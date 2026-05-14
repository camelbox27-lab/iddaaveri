"""
patch_missing_scores.py - Eksik IY/MS Skoru olan maclari yeniden ceker

Kullanim:
    python scripts/patch_missing_scores.py --input output/iddaagecmismaclar_clean.xlsx
    python scripts/patch_missing_scores.py --input output/iddaagecmismaclar_clean.xlsx --resume
    python scripts/patch_missing_scores.py --input output/iddaagecmismaclar_clean.xlsx --max-hours 5.5

Mantik:
1. Excel'i oku, MS Skoru veya IY Skoru bos olan satirlari bul
2. Mackolik API'den o gun icin match listesi cek -> MS kodu ile eslesir
3. Skoru guncelle, yoksa iddaa sayfasini scrape et
4. Guncellenmis Excel'i kaydet
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import io
import time
import re
import unicodedata
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).resolve().parent.parent
BASE_URL = "https://www.mackolik.com"
API_URL  = "https://www.mackolik.com/perform/p0/ajax/components/competition/livescores/json?"

_SESSION = requests.Session()
_SESSION.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'tr,en-US;q=0.7,en;q=0.3',
    'Connection': 'keep-alive',
})

_throttle_delay = 0.05
_throttle_lock  = threading.Lock()

def _throttle_hit():
    global _throttle_delay
    with _throttle_lock:
        _throttle_delay = min(_throttle_delay * 1.5, 1.0)

def _throttle_ok():
    global _throttle_delay
    with _throttle_lock:
        _throttle_delay = max(_throttle_delay * 0.8, 0.02)


def _norm(v):
    if not v: return ""
    return " ".join(str(v).replace("\xa0", " ").split()).strip()

def _fold(s: str) -> str:
    if not s: return ""
    s = str(s).replace("I", "i").replace("İ", "i").lower()
    for old, new in [("ş","s"),("ğ","g"),("ü","u"),("ö","o"),
                     ("ç","c"),("ı","i")]:
        s = s.replace(old, new)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return " ".join(s.split())


def is_score_empty(val) -> bool:
    if val is None or (isinstance(val, float) and str(val) == 'nan'):
        return True
    s = str(val).strip()
    return s in ('', '-', 'nan', 'None', 'none')


def parse_date(val) -> dt.date | None:
    if not val:
        return None
    s = str(val).strip()
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y'):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


# ── API ile gun verisi ───────────────────────────────────────────────────────
def fetch_api_day(target_date: dt.date, max_retries: int = 4) -> dict[str, dict]:
    """API'den o gune ait maci cek. MS kodu -> {iy_skor, ms_skor, mac_url}"""
    params = {'sports[]': 'Soccer', 'matchDate': target_date.strftime('%Y-%m-%d')}
    data = {}
    for attempt in range(max_retries):
        try:
            resp = _SESSION.get(API_URL, params=params, timeout=15)
            resp.raise_for_status()
            jr = resp.json()
            data = jr.get('data', jr)
            break
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep((attempt + 1) * 2)
            else:
                print(f"  [API] {target_date} basarisiz: {e}", flush=True)
                return {}

    matches_raw  = data.get('matches', {})
    competitions = data.get('competitions', {})

    if isinstance(matches_raw, list):
        matches_raw = {str(i): m for i, m in enumerate(matches_raw)}
    if isinstance(competitions, list):
        competitions = {str(c.get('id', i)): c for i, c in enumerate(competitions)}

    result: dict[str, dict] = {}
    for mid, m in (matches_raw if isinstance(matches_raw, dict) else {}).items():
        code = str(m.get('iddaaCode', '') or '')
        if not code or code == 'None':
            continue

        score = m.get('score', {}) or {}
        ft_h = score.get('home', '')
        ft_a = score.get('away', '')
        ms_skor = f"{ft_h}-{ft_a}" if ft_h != '' and ft_a != '' else ''

        ht = score.get('ht', {}) or {}
        ht_h = str(ht.get('home', '')).strip() if isinstance(ht, dict) else ''
        ht_a = str(ht.get('away', '')).strip() if isinstance(ht, dict) else ''
        iy_skor = f"{ht_h}-{ht_a}" if (ht_h or ht_a) else ''

        # Mac URL
        comp_id  = str(m.get('competitionId', ''))
        comp     = (competitions.get(comp_id) or {})
        match_slug   = m.get('matchSlug', '') or m.get('slug', '')
        comp_slug    = comp.get('competitionSlug', '')
        country_slug = comp.get('countrySlug', '')
        season_slug  = comp.get('seasonSlug', '')

        if match_slug and comp_slug and country_slug:
            mac_url = f"{BASE_URL}/futbol/{country_slug}/{comp_slug}/{season_slug}/mac/{match_slug}/iddaa"
        else:
            mac_url = ''

        result[code] = {
            'iy_skor': iy_skor,
            'ms_skor': ms_skor,
            'mac_url': mac_url,
            'home': (m.get('homeTeam', {}) or {}).get('name', ''),
            'away': (m.get('awayTeam', {}) or {}).get('name', ''),
        }
    return result


# ── Iddaa sayfasi scrape ─────────────────────────────────────────────────────
def scrape_iddaa_page(url: str, max_retries: int = 3) -> dict:
    """Iddaa sayfasindan IY/MS skor + oranlari cek."""
    result = {}
    for attempt in range(max_retries):
        try:
            time.sleep(_throttle_delay)
            resp = _SESSION.get(url, timeout=12, allow_redirects=True)
            if resp.status_code == 404:
                break
            if resp.status_code in (429, 500, 502, 503):
                _throttle_hit()
                if attempt < max_retries - 1:
                    time.sleep((attempt + 1) * 3)
                    continue
                break
            resp.raise_for_status()
            _throttle_ok()
            html = resp.text
            soup = BeautifulSoup(html, 'html.parser')

            # IY Skor — mackolik score widget'indan
            iy_el = soup.select_one('.match-row__half-time-score, [class*="half-time"], .p0c-soccer-match-details-score__half-time')
            if iy_el:
                iy_raw = _norm(iy_el.get_text()).replace('IY', '').replace('iy', '').strip()
                if re.match(r'\d+-\d+', iy_raw):
                    result['iy_skor'] = iy_raw

            # MS Skor
            ms_el = soup.select_one('.p0c-soccer-match-details-score__score, [class*="score__score"], .match-row__score')
            if ms_el:
                ms_raw = _norm(ms_el.get_text()).strip()
                if re.match(r'\d+-\d+', ms_raw):
                    result['ms_skor'] = ms_raw

            # Header'dan alternatif skor
            hdr = soup.select_one('[class*="p0c-soccer-match-details-header__score"]')
            if hdr and not result.get('ms_skor'):
                txt = _norm(hdr.get_text())
                m2 = re.search(r'(\d+)\s*-\s*(\d+)', txt)
                if m2:
                    result['ms_skor'] = f"{m2.group(1)}-{m2.group(2)}"

            return result
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep((attempt + 1) * 2)
    return result


# ── Progress ─────────────────────────────────────────────────────────────────
def load_progress(progress_file: Path) -> dict:
    if progress_file.exists():
        try:
            return json.loads(progress_file.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {}


def save_progress(progress_file: Path, data: dict):
    progress_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


# ── Excel kaydet ─────────────────────────────────────────────────────────────
HEADERS_ORDER = [
    "Ev Sahibi", "Deplasman", "Tarih", "Saat", "Lig",
    "MS Kodu", "IY Skor", "MS Skor",
    "MS1", "MS0", "MS2",
    "CS 1X", "CS 12", "CS X2",
    "IY1", "IY0", "IY2",
    "AU 0.5 Alt", "AU 0.5 Ust",
    "AU 1.5 Alt", "AU 1.5 Ust",
    "AU 2.5 Alt", "AU 2.5 Ust",
    "AU 3.5 Alt", "AU 3.5 Ust",
    "AU 4.5 Alt", "AU 4.5 Ust",
    "KG Var", "KG Yok",
    "HND1", "HNDX", "HND2",
    "HND2-1", "HND2-X", "HND2-2",
    "IY AU 0.5 Alt", "IY AU 0.5 Ust",
    "IY AU 1.5 Alt", "IY AU 1.5 Ust",
    "IY/MS 1/1", "IY/MS 1/X", "IY/MS 1/2",
    "IY/MS X/1", "IY/MS X/X", "IY/MS X/2",
    "IY/MS 2/1", "IY/MS 2/X", "IY/MS 2/2",
    "TG 0-1", "TG 2-3", "TG 4-5", "TG 6+",
    "T1 1.5 Ust", "T1 2.5 Ust",
    "T2 1.5 Ust", "T2 2.5 Ust",
]


def save_df_to_excel(df: pd.DataFrame, path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "MS Oranlari"
    hfill = PatternFill("solid", fgColor="2E7D32")
    hfont = Font(color="FFFFFF", bold=True)

    cols = [c for c in HEADERS_ORDER if c in df.columns]
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center")

    for _, row in df[cols].iterrows():
        ws.append([None if pd.isna(v) else v for v in row])

    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(w + 2, 40)

    wb.save(path)


# ── Ana fonksiyon ────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='Eksik IY/MS skorlari patch et')
    parser.add_argument('--input',     default='output/iddaagecmismaclar_clean.xlsx')
    parser.add_argument('--output',    default=None, help='Cikti dosyasi (yoksa input uzerine yazar)')
    parser.add_argument('--progress',  default='progress_patch.json')
    parser.add_argument('--max-hours', type=float, default=5.5)
    parser.add_argument('--workers',   type=int, default=6)
    parser.add_argument('--resume',    action='store_true')
    args = parser.parse_args()

    input_path   = BASE_DIR / args.input
    output_path  = BASE_DIR / (args.output or args.input)
    progress_file = BASE_DIR / args.progress
    deadline = dt.datetime.now() + dt.timedelta(hours=args.max_hours)

    print("=" * 60)
    print("Patch: Eksik IY/MS Skor Scripti")
    print(f"Girdi : {input_path}")
    print(f"Cikti : {output_path}")
    print(f"Max   : {args.max_hours} saat")
    print("=" * 60)

    # Excel oku
    print(f"\n[1] Excel okunuyor...")
    df = pd.read_excel(input_path, dtype=str)
    print(f"    {len(df):,} satir, {len(df.columns)} sutun")

    # Standart kolon isimlerini normalize et
    col_map = {}
    for c in df.columns:
        cf = _fold(str(c))
        if cf in ('ev sahibi',): col_map[c] = 'Ev Sahibi'
        elif cf in ('deplasman', 'konuk ekip'): col_map[c] = 'Deplasman'
        elif cf in ('tarih',): col_map[c] = 'Tarih'
        elif cf in ('saat',): col_map[c] = 'Saat'
        elif cf in ('lig',): col_map[c] = 'Lig'
        elif cf in ('ms kodu',): col_map[c] = 'MS Kodu'
        elif cf in ('iy skor',): col_map[c] = 'IY Skor'
        elif cf in ('ms skor',): col_map[c] = 'MS Skor'
    df = df.rename(columns=col_map)

    # Gerekli kolonlar yoksa ekle
    for col in ('IY Skor', 'MS Skor', 'MS Kodu', 'Tarih'):
        if col not in df.columns:
            df[col] = ''

    # Eksik skorlu satirlari bul
    ms_bos = df['MS Skor'].apply(is_score_empty)
    iy_bos = df['IY Skor'].apply(is_score_empty)
    eksik_mask = ms_bos | iy_bos
    eksik_idxs = df[eksik_mask].index.tolist()

    print(f"    MS Skor bos: {ms_bos.sum():,}")
    print(f"    IY Skor bos: {iy_bos.sum():,}")
    print(f"    Toplam eksik: {len(eksik_idxs):,} satir")

    # Progress yukle
    progress = load_progress(progress_file) if args.resume else {}
    done_codes = set(progress.get('done_codes', []))
    print(f"    Onceden tamamlanan: {len(done_codes):,}")

    # Tarihe gore grupla
    from collections import defaultdict
    date_groups: dict[str, list[int]] = defaultdict(list)
    for idx in eksik_idxs:
        tarih = str(df.at[idx, 'Tarih']).strip()
        date_groups[tarih].append(idx)

    print(f"\n[2] {len(date_groups)} farkli tarih isleniyor...\n")

    updated_count = 0
    api_hit_count = 0
    batch_save_interval = 500  # Her 500 guncelleme kaydet
    last_save = 0

    for tarih_str, idxs in sorted(date_groups.items()):
        if dt.datetime.now() >= deadline:
            print(f"\n[!] Zaman doldu. {updated_count} guncelleme yapildi.", flush=True)
            break

        target_date = parse_date(tarih_str)
        if not target_date:
            continue

        # API'den gun verisi cek
        api_data = fetch_api_day(target_date)
        api_hit_count += 1

        for idx in idxs:
            ms_kodu = str(df.at[idx, 'MS Kodu']).strip()
            if ms_kodu in done_codes:
                continue

            updated = False

            # API ile dene
            if ms_kodu and ms_kodu in api_data:
                api_match = api_data[ms_kodu]
                if is_score_empty(df.at[idx, 'MS Skor']) and api_match.get('ms_skor'):
                    df.at[idx, 'MS Skor'] = api_match['ms_skor']
                    updated = True
                if is_score_empty(df.at[idx, 'IY Skor']) and api_match.get('iy_skor'):
                    df.at[idx, 'IY Skor'] = api_match['iy_skor']
                    updated = True

                # Hala bos ise iddaa sayfasini dene
                if (is_score_empty(df.at[idx, 'MS Skor']) or is_score_empty(df.at[idx, 'IY Skor'])) and api_match.get('mac_url'):
                    page_data = scrape_iddaa_page(api_match['mac_url'])
                    if is_score_empty(df.at[idx, 'MS Skor']) and page_data.get('ms_skor'):
                        df.at[idx, 'MS Skor'] = page_data['ms_skor']
                        updated = True
                    if is_score_empty(df.at[idx, 'IY Skor']) and page_data.get('iy_skor'):
                        df.at[idx, 'IY Skor'] = page_data['iy_skor']
                        updated = True

            if updated:
                updated_count += 1

            done_codes.add(ms_kodu)

        time.sleep(0.1)

        # Ara kayit
        if updated_count - last_save >= batch_save_interval:
            print(f"  [KAYIT] {updated_count} guncelleme -> {output_path}", flush=True)
            save_df_to_excel(df, output_path)
            progress['done_codes'] = list(done_codes)
            save_progress(progress_file, progress)
            last_save = updated_count

    # Son kayit
    print(f"\n[3] Final kayit...")
    save_df_to_excel(df, output_path)
    progress['done_codes'] = list(done_codes)
    save_progress(progress_file, progress)

    print(f"\n{'='*60}")
    print(f"OZET:")
    print(f"  Eksik satirlar : {len(eksik_idxs):,}")
    print(f"  API sorgusu    : {api_hit_count:,}")
    print(f"  Guncellenen    : {updated_count:,}")
    print(f"  Cikti          : {output_path}")
    print("="*60)


if __name__ == '__main__':
    main()
