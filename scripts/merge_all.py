"""
merge_all.py - Tum Excel dosyalarini birlestirip final temiz Excel olusturur

Kaynaklar:
  1. output/iddaagecmismaclar_patched.xlsx  (ana veri + patch edilmis skorlar)
  2. output/eksik_2017_08_2018_03.xlsx       (2017-08 / 2018-03)
  3. output/eksik_2018_04_2018_11.xlsx       (2018-04 / 2018-11)
  4. output/eksik_2018_12_2019_07.xlsx       (2018-12 / 2019-07)
  5. output/guncel_2026_03_2026_05.xlsx      (2026-03 / 2026-05)
  6. output/iddaa_uefa_*.xlsx                (UEFA maclari - varsa)

Cikti: output/iddaagecmismaclar_FINAL.xlsx
"""
from __future__ import annotations

import sys
import io
import datetime as dt
from pathlib import Path

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR    = Path(__file__).resolve().parent.parent
OUTPUT_DIR  = BASE_DIR / "output"
FINAL_PATH  = OUTPUT_DIR / "iddaagecmismaclar_FINAL.xlsx"

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

# Alternatif kolon ismi eslesmesi (eski dosyalar farkli isim kullanabilir)
COL_ALIASES = {
    'Konuk Ekip': 'Deplasman',
    'konuk ekip': 'Deplasman',
    'Ev Sahibi ': 'Ev Sahibi',
}

DUP_KEYS = ['Ev Sahibi', 'Deplasman', 'Tarih', 'Saat']


def is_score_valid(val) -> bool:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return False
    s = str(val).strip()
    return s not in ('', '-', 'nan', 'None', 'none')


def count_filled(row) -> int:
    count = 0
    for v in row:
        if v is not None and not (isinstance(v, float) and np.isnan(v)):
            s = str(v).strip()
            if s not in ('', '-', 'nan', 'None'):
                count += 1
    return count


def pick_best_row(group_df: pd.DataFrame) -> int:
    """Duplikat grubundan en iyi satiri sec."""
    if len(group_df) == 1:
        return group_df.index[0]

    # MS Skor dolu olanı tercih et
    if 'MS Skor' in group_df.columns:
        ms_valid = group_df['MS Skor'].apply(is_score_valid)
        if ms_valid.sum() == 1:
            return group_df[ms_valid].index[0]
        elif ms_valid.sum() > 1:
            group_df = group_df[ms_valid]

    # IY Skor dolu olanı tercih et
    if 'IY Skor' in group_df.columns:
        iy_valid = group_df['IY Skor'].apply(is_score_valid)
        if iy_valid.sum() == 1:
            return group_df[iy_valid].index[0]
        elif iy_valid.sum() > 1:
            group_df = group_df[iy_valid]

    # En fazla dolu sutun
    filled = group_df.apply(lambda r: count_filled(r.tolist()), axis=1)
    return filled.idxmax()


def parse_date(val) -> dt.date | None:
    if not val or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def normalize_df(df: pd.DataFrame, source_label: str) -> pd.DataFrame:
    """DataFrame kolonlarini standart HEADERS_ORDER'a gore normalize et."""
    # Alias'lari uygula
    df = df.rename(columns=COL_ALIASES)

    # Eksik kolonlari bos olarak ekle
    for col in HEADERS_ORDER:
        if col not in df.columns:
            df[col] = None

    # Sadece istenen kolonları tut, siralama duzelt
    df = df[HEADERS_ORDER].copy()

    # Tarih formatini standartlastir
    def fix_date(v):
        d = parse_date(v)
        return d.strftime('%d.%m.%Y') if d else v
    df['Tarih'] = df['Tarih'].apply(fix_date)

    # String kolonlari temizle
    for col in ('Ev Sahibi', 'Deplasman', 'Saat', 'Lig', 'MS Kodu', 'IY Skor', 'MS Skor'):
        df[col] = df[col].apply(lambda v:
            None if (v is None or (isinstance(v, float) and np.isnan(v)))
            else str(v).strip()
        )

    return df


def load_excel(path: Path, source: str) -> pd.DataFrame | None:
    if not path.exists():
        print(f"  [ATLA] {path.name} bulunamadi", flush=True)
        return None
    try:
        df = pd.read_excel(path, dtype=str)
        print(f"  [OK]   {path.name}: {len(df):,} satir", flush=True)
        return normalize_df(df, source)
    except Exception as e:
        print(f"  [HATA] {path.name}: {e}", flush=True)
        return None


def save_final_excel(df: pd.DataFrame, path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "MS Oranlari"
    hfill = PatternFill("solid", fgColor="2E7D32")
    hfont = Font(color="FFFFFF", bold=True)

    ws.append(HEADERS_ORDER)
    for cell in ws[1]:
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        ws.append([None if pd.isna(v) else v for v in row])

    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(w + 2, 40)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"  Kaydedildi: {path}", flush=True)


def main():
    print("=" * 65)
    print("FINAL MERGE: Tum Veriler Birlestiriliyor")
    print(f"Hedef: 2017-08-01 - 2026-05-14")
    print("=" * 65)

    # ── Dosya listesi (oncelik sirasi: daha gec / daha tam olanlar once) ──
    sources = [
        # Ana veri (patch edilmis)
        (OUTPUT_DIR / "iddaagecmismaclar_patched.xlsx", "patched"),
        # Patch yapilmamissa clean versiyonu dene
        (OUTPUT_DIR / "iddaagecmismaclar_clean.xlsx",   "clean"),
        # Eksik donemler
        (OUTPUT_DIR / "eksik_2017_08_2018_03.xlsx",     "eksik-1"),
        (OUTPUT_DIR / "eksik_2018_04_2018_11.xlsx",     "eksik-2"),
        (OUTPUT_DIR / "eksik_2018_12_2019_07.xlsx",     "eksik-3"),
        # Guncel donem
        (OUTPUT_DIR / "guncel_2026_03_2026_05.xlsx",    "guncel"),
    ]

    # UEFA dosyalari varsa ekle
    for uefa_path in sorted(OUTPUT_DIR.glob("iddaa_uefa_*.xlsx")):
        sources.append((uefa_path, f"uefa-{uefa_path.stem}"))

    # Eski rescrape dosyalari varsa ekle (iddaa_YYYY_YYYY.xlsx)
    for extra in sorted(OUTPUT_DIR.glob("iddaa_2*.xlsx")):
        if extra.stem not in ('iddaagecmismaclar_FINAL', 'iddaagecmismaclar'):
            sources.append((extra, f"extra-{extra.stem}"))

    print(f"\n[1] Dosyalar yukleniyor...")
    dfs = []
    for path, label in sources:
        df = load_excel(path, label)
        if df is not None and len(df) > 0:
            dfs.append(df)

    if not dfs:
        print("HATA: Hicbir kaynak dosya yuklenemedi!", flush=True)
        sys.exit(1)

    print(f"\n[2] Birlestiriliyor...")
    df_all = pd.concat(dfs, ignore_index=True)
    print(f"    Ham toplam: {len(df_all):,} satir", flush=True)

    # ── Tarih filtresi: 2017-08-01 ile 2026-05-14 ──
    print(f"\n[3] Tarih filtresi uygulanıyor (2017-08-01 → 2026-05-14)...")
    start_filter = dt.date(2017, 8, 1)
    end_filter   = dt.date(2026, 5, 14)

    def in_range(val) -> bool:
        d = parse_date(val)
        if not d:
            return False
        return start_filter <= d <= end_filter

    mask_date = df_all['Tarih'].apply(in_range)
    outside   = (~mask_date).sum()
    df_all    = df_all[mask_date].reset_index(drop=True)
    print(f"    Tarih araligi disinda silinen: {outside:,}", flush=True)
    print(f"    Kalan: {len(df_all):,}", flush=True)

    # ── Bos Ev Sahibi satirlari sil ──
    df_all = df_all[df_all['Ev Sahibi'].notna() & (df_all['Ev Sahibi'].str.strip() != '')].reset_index(drop=True)
    print(f"    Bos EV SAHİBİ silindi, kalan: {len(df_all):,}", flush=True)

    # ── Duplikat temizle (akilli) ──
    print(f"\n[4] Duplikatlar temizleniyor...")
    dup_mask   = df_all.duplicated(subset=DUP_KEYS, keep=False)
    dup_count  = dup_mask.sum()
    print(f"    Duplikat satir toplami: {dup_count:,}", flush=True)

    if dup_count > 0:
        dup_df     = df_all[dup_mask]
        dup_groups = dup_df.groupby(DUP_KEYS, dropna=False)
        keep_idxs  = set()
        for _, grp in dup_groups:
            keep_idxs.add(pick_best_row(grp))

        non_dup_idxs = set(df_all[~dup_mask].index.tolist())
        final_idxs   = sorted(non_dup_idxs | keep_idxs)
        df_all       = df_all.loc[final_idxs].reset_index(drop=True)

        removed_dups = dup_count - len(keep_idxs)
        print(f"    Silinen duplikat: {removed_dups:,}", flush=True)

    print(f"    Temiz toplam: {len(df_all):,}", flush=True)

    # ── Tarihe gore sirala ──
    print(f"\n[5] Tarihe gore siralanıyor...")
    df_all['_sort'] = df_all['Tarih'].apply(parse_date)
    df_all = df_all.sort_values('_sort', na_position='last').drop(columns=['_sort'])
    df_all = df_all.reset_index(drop=True)

    # ── Istatistik raporu ──
    print(f"\n[6] Istatistik Raporu:")
    ms_bos = df_all['MS Skor'].apply(lambda v: not is_score_valid(v)).sum()
    iy_bos = df_all['IY Skor'].apply(lambda v: not is_score_valid(v)).sum()
    print(f"    Toplam mac        : {len(df_all):,}")
    print(f"    MS Skor bos       : {ms_bos:,}")
    print(f"    IY Skor bos       : {iy_bos:,}")

    # Aylik dagilim ozeti
    from collections import Counter
    ay_sayac: Counter = Counter()
    for v in df_all['Tarih']:
        d = parse_date(v)
        if d:
            ay_sayac[f"{d.year}-{d.month:02d}"] += 1

    print(f"\n    Aylik dagilim (ilk / son 5):")
    months = sorted(ay_sayac.items())
    for ay, say in months[:5]:
        print(f"      {ay}: {say:5d} mac")
    print("      ...")
    for ay, say in months[-5:]:
        print(f"      {ay}: {say:5d} mac")

    eksik_aylar = [ay for ay, say in months if say == 0]
    if eksik_aylar:
        print(f"\n    UYARI - Hic mac olmayan aylar: {eksik_aylar}")
    else:
        print(f"\n    Tum aylar kapli (bos ay yok)")

    # ── Kaydet ──
    print(f"\n[7] Final Excel kaydediliyor...")
    save_final_excel(df_all, FINAL_PATH)

    print(f"\n{'='*65}")
    print(f"TAMAMLANDI: {FINAL_PATH}")
    print(f"Toplam satir: {len(df_all):,}")
    print("="*65)


if __name__ == '__main__':
    main()
