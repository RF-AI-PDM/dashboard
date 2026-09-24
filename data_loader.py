"""Data loading and formatting utilities for Dashboard Monitoring Pekerjaan.

Reads and processes sheets from Excel file data/Data.xlsx.
All source code must be ASCII-only.
"""

from __future__ import annotations

from datetime import date, datetime
import math
import os
import re
from typing import Any, Dict, List, Optional
import numpy as np
import pandas as pd


def fmt_rp(x: Any) -> str:
    """Format number as Indonesian Rupiah string (e.g. 'Rp 16.618.722.076').

    Returns '-' if x is None, NaN, or cannot be converted to a valid number.
    """
    if x is None or pd.isna(x):
        return "-"
    try:
        val = float(x)
        if math.isnan(val) or math.isinf(val):
            return "-"
        val_int = int(round(val))
        formatted = f"{val_int:,}".replace(",", ".")
        return f"Rp {formatted}"
    except (ValueError, TypeError):
        return "-"


def fmt_pct(x: Any) -> str:
    """Format decimal/ratio as percentage string with comma (e.g. 0.6924 -> '69,2%').

    Returns '-' if x is None, NaN, or invalid.
    """
    if x is None or pd.isna(x):
        return "-"
    try:
        val = float(x)
        if math.isnan(val) or math.isinf(val):
            return "-"
        pct_val = val * 100.0
        formatted = f"{pct_val:.1f}%".replace(".", ",")
        return formatted
    except (ValueError, TypeError):
        return "-"


def _clean_str(val: Any, default: str = "") -> str:
    """Helper to convert value to cleanly stripped string."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    return default if s.lower() == "nan" else s


# Status ringkas yang dipertahankan apa adanya; selain ini dianggap catatan progres.
_CANONICAL_STATUS = {"SELESAI", "CANCEL", "BELUM ADA STATUS"}
_CANONICAL_PAYMENT = {"PAYMENT", "PROSES", "CANCEL", "BELUM ADA STATUS"}


def _normalize_status(df: pd.DataFrame) -> None:
    """Rapikan kolom STATUS / STATUS PEMBAYARAN agar berisi kategori ringkas.

    Sel STATUS di Excel sering diisi catatan panjang (mis. '1. Menunggu BA dari user ...').
    Catatan tersebut dipindahkan ke kolom Kendala (jika belum ada) dan STATUS
    diringkas menjadi 'PROSES'. Teks asli tetap tersimpan di kolom 'STATUS (Asli)'.
    """
    if "STATUS" not in df.columns:
        return

    df["STATUS (Asli)"] = df["STATUS"]
    kendala_col = "Kendala" if "Kendala" in df.columns else None

    for idx, raw in df["STATUS"].items():
        text = str(raw).strip()
        upper = text.upper()
        if upper in _CANONICAL_STATUS:
            df.at[idx, "STATUS"] = upper
            continue
        # Bukan status baku -> anggap sedang berjalan, simpan catatannya di Kendala.
        df.at[idx, "STATUS"] = "PROSES"
        if kendala_col:
            existing = str(df.at[idx, kendala_col]).strip()
            if not existing:
                df.at[idx, kendala_col] = text
            elif text not in existing:
                df.at[idx, kendala_col] = f"{existing}\n{text}"

    if "STATUS PEMBAYARAN" in df.columns:
        def _pay(v: object) -> str:
            u = str(v).strip().upper()
            if u in _CANONICAL_PAYMENT:
                return u
            if "PAYMENT" in u:
                return "PROSES PAYMENT"
            if "MENUNGGU" in u:
                return "MENUNGGU"
            return "PROSES"

        df["STATUS PEMBAYARAN (Asli)"] = df["STATUS PEMBAYARAN"]
        df["STATUS PEMBAYARAN"] = df["STATUS PEMBAYARAN"].apply(_pay)


def _parse_monitoring(raw: pd.DataFrame) -> pd.DataFrame:
    """Parse sheet 'Monitoring' (shape 38x39).

    Row 0: Header.
    Rows 3..36: Data rows. Row 37: TOTAL (skip).
    Filter: Keep only rows where col 2 (Judul Pekerjaan) is non-empty.
    """
    col_mapping = {
        1: "NO",
        2: "Judul Pekerjaan",
        3: "Projeck",
        4: "User",
        5: "Proses Kontrak",
        6: "Eksekutor",
        7: "DMR/TOR/Spesifikasi",
        8: "LOI",
        9: "Nomor Informasi Harga Mitra",
        10: "IH Sebelum PPN",
        11: "IH Setelah PPN",
        12: "Undangan PL",
        13: "Kontrak IPS",
        14: "Nilai Kontrak Sebelum PPN (IPS)",
        15: "Nilai Kontrak Setelah PPN (IPS)",
        16: "Mulai Pekerjaan",
        17: "Berakhir Pekerjaan",
        18: "Addendum Kontrak",
        19: "Ket Addendum",
        20: "BAST/TUG",
        21: "Invoice",
        22: "Nilai Invoice Sebelum PPN",
        23: "Nilai Invoice Setelah PPN",
        24: "DENDA",
        25: "PEMBAYARAN PIPS",
        26: "Nama Vendor",
        27: "Nomor Kontrak Vendor",
        28: "Nilai Kontrak Mitra Sebelum PPN",
        29: "Nilai Kontrak Mitra Setelah PPN",
        30: "BA Vendor",
        31: "Nilai BA Vendor Setelah PPN",
        32: "DENDA2",
        33: "PEMBAYARAN MITRA",
        34: "MARGIN",
        35: "Kendala",
        36: "Periode Penagihan",
        37: "STATUS",
        38: "STATUS PEMBAYARAN",
    }

    numeric_col_indices = [10, 11, 14, 15, 22, 23, 24, 28, 29, 31, 32, 34]
    numeric_names = [col_mapping[i] for i in numeric_col_indices if i in col_mapping]

    # Find where TOTAL row starts
    end_row = len(raw)
    for r in range(3, len(raw)):
        cell_val = str(raw.iloc[r, 1]).strip().upper() if pd.notna(raw.iloc[r, 1]) else ""
        if cell_val == "TOTAL":
            end_row = r
            break

    # Slice rows 3 to end_row
    data_slice = raw.iloc[3:end_row].copy()

    # Filter rows where col 2 (Judul Pekerjaan) is not empty
    mask = (
        data_slice[2].notna()
        & (data_slice[2].astype(str).str.strip() != "")
        & (data_slice[2].astype(str).str.strip().str.lower() != "nan")
    )
    data_slice = data_slice[mask]

    # Build clean DataFrame
    clean_df = pd.DataFrame()
    for col_idx, col_name in col_mapping.items():
        if col_idx < data_slice.shape[1]:
            clean_df[col_name] = data_slice[col_idx]
        else:
            clean_df[col_name] = np.nan

    # Coerce numeric columns
    for num_col in numeric_names:
        if num_col in clean_df.columns:
            clean_df[num_col] = pd.to_numeric(clean_df[num_col], errors="coerce")

    # Clean text columns and fill empty status
    for col in clean_df.columns:
        if col not in numeric_names:
            if col in ["STATUS", "STATUS PEMBAYARAN"]:
                clean_df[col] = clean_df[col].apply(
                    lambda v: "BELUM ADA STATUS"
                    if pd.isna(v) or not str(v).strip() or str(v).strip().lower() == "nan"
                    else str(v).strip()
                )
            elif col == "NO":
                clean_df[col] = clean_df[col].apply(
                    lambda v: int(v)
                    if pd.notna(v) and str(v).replace(".0", "").isdigit()
                    else (_clean_str(v) or "-")
                )
            else:
                clean_df[col] = clean_df[col].apply(lambda v: _clean_str(v, default=""))

    _normalize_status(clean_df)

    clean_df.reset_index(drop=True, inplace=True)
    return clean_df


def _get_float(series: Any, idx: int, default: float = 0.0) -> float:
    """Safe float getter from pandas Series or list."""
    if len(series) > idx and pd.notna(series[idx]):
        try:
            return float(series[idx])
        except (ValueError, TypeError):
            return default
    return default


def _get_str(series: Any, idx: int, default: str) -> str:
    """Safe string getter from pandas Series or list."""
    if len(series) > idx and pd.notna(series[idx]):
        s = str(series[idx]).strip()
        if s and s.lower() != "nan":
            return s
    return default


def _parse_pencapaian(raw: pd.DataFrame, realisasi_sum: float) -> Dict[str, Any]:
    """Parse sheet 'PENCAPAIAN' without relying on formula cells.

    Constant targets are read from C6 (idx 2), E6 (idx 4), G6 (idx 6) in row 5 (0-based).
    Realisasi is computed from monitoring sum.
    """
    labels_row = raw.iloc[4] if len(raw) > 4 else []
    values_row = raw.iloc[5] if len(raw) > 5 else []

    target_tahun = _get_float(values_row, 2, 0.0)
    target_s1 = _get_float(values_row, 4, 0.0)
    target_s2 = _get_float(values_row, 6, 0.0)

    realisasi_tahun = realisasi_sum
    capaian_s1 = realisasi_sum
    capaian_s2 = realisasi_sum

    ratio_tahun = (realisasi_tahun / target_tahun) if target_tahun != 0 else 0.0
    ratio_s1 = (capaian_s1 / target_s1) if target_s1 != 0 else 0.0
    ratio_s2 = (capaian_s2 / target_s2) if target_s2 != 0 else 0.0

    deviasi_tahun = realisasi_tahun - target_tahun
    deviasi_s1 = capaian_s1 - target_s1
    deviasi_s2 = capaian_s2 - target_s2

    return {
        "labels": {
            "target_tahun": _get_str(labels_row, 2, "TARGET TAHUN 2026"),
            "realisasi_tahun": _get_str(labels_row, 3, "REALISASI TAHUN 2026"),
            "target_s1": _get_str(labels_row, 4, "TARGET SMESTER 1"),
            "capaian_s1": _get_str(labels_row, 5, "Pencapaian Semester 1"),
            "target_s2": _get_str(labels_row, 6, "TARGET SMESTER 2"),
            "capaian_s2": _get_str(labels_row, 7, "Pencapaian Semester 2"),
        },
        "target_tahun": target_tahun,
        "realisasi_tahun": realisasi_tahun,
        "target_s1": target_s1,
        "capaian_s1": capaian_s1,
        "target_s2": target_s2,
        "capaian_s2": capaian_s2,
        "ratio_tahun": ratio_tahun,
        "ratio_s1": ratio_s1,
        "ratio_s2": ratio_s2,
        "deviasi_tahun": deviasi_tahun,
        "deviasi_s1": deviasi_s1,
        "deviasi_s2": deviasi_s2,
    }


def _parse_akumulatif_kpi(raw: pd.DataFrame, realisasi_sum: float) -> Dict[str, Any]:
    """Parse sheet 'Grafik akumulatif' without relying on formula cells.

    Targets from C6 (idx 2), E6 (idx 4), G6 (idx 6).
    capaian_s1 from F6 (idx 5) if numeric else realisasi_sum.
    realisasi_tahun = capaian_s2 = realisasi_sum.
    """
    labels_row = raw.iloc[4] if len(raw) > 4 else []
    values_row = raw.iloc[5] if len(raw) > 5 else []

    target_tahun = _get_float(values_row, 2, 0.0)
    target_s1 = _get_float(values_row, 4, 0.0)
    f6_val = _get_float(values_row, 5, 0.0)
    capaian_s1 = f6_val if f6_val > 0 else realisasi_sum
    target_s2 = _get_float(values_row, 6, 0.0)

    realisasi_tahun = realisasi_sum
    capaian_s2 = realisasi_sum

    ratio_tahun = (realisasi_tahun / target_tahun) if target_tahun != 0 else 0.0
    ratio_s1 = (capaian_s1 / target_s1) if target_s1 != 0 else 0.0
    ratio_s2 = (capaian_s2 / target_s2) if target_s2 != 0 else 0.0

    deviasi_tahun = realisasi_tahun - target_tahun
    deviasi_s1 = capaian_s1 - target_s1
    deviasi_s2 = capaian_s2 - target_s2

    return {
        "labels": {
            "target_tahun": _get_str(labels_row, 2, "TARGET TAHUN 2026"),
            "realisasi_tahun": _get_str(labels_row, 3, "REALISASI TAHUN 2026"),
            "target_s1": _get_str(labels_row, 4, "TARGET SMESTER 1"),
            "capaian_s1": _get_str(labels_row, 5, "Pencapaian Semester 1"),
            "target_s2": _get_str(labels_row, 6, "TARGET SMESTER 2"),
            "capaian_s2": _get_str(labels_row, 7, "Pencapaian Semester 2"),
        },
        "target_tahun": target_tahun,
        "realisasi_tahun": realisasi_tahun,
        "target_s1": target_s1,
        "capaian_s1": capaian_s1,
        "target_s2": target_s2,
        "capaian_s2": capaian_s2,
        "ratio_tahun": ratio_tahun,
        "ratio_s1": ratio_s1,
        "ratio_s2": ratio_s2,
        "deviasi_tahun": deviasi_tahun,
        "deviasi_s1": deviasi_s1,
        "deviasi_s2": deviasi_s2,
    }


def _parse_monitoring_akumulatif(raw: pd.DataFrame) -> pd.DataFrame:
    """Parse sheet 'Monitoring Akumulatif '."""
    col_mapping = {
        1: "NO",
        2: "Judul Pekerjaan",
        3: "User",
        4: "Kontrak",
        5: "Nilai Kontrak Setelah PPN",
        6: "Nama Vendor",
        7: "Kontrak Vendor",
        8: "Nilai Kontrak Mitra Setelah PPN",
        9: "STATUS",
    }
    numeric_indices = [5, 8]

    # Find total row
    end_row = len(raw)
    for r in range(1, len(raw)):
        cell_val = str(raw.iloc[r, 1]).strip().upper() if pd.notna(raw.iloc[r, 1]) else ""
        if cell_val == "TOTAL":
            end_row = r
            break

    data_slice = raw.iloc[1:end_row].copy()
    mask = (
        data_slice[2].notna()
        & (data_slice[2].astype(str).str.strip() != "")
        & (data_slice[2].astype(str).str.strip().str.lower() != "nan")
    )
    data_slice = data_slice[mask]

    clean_df = pd.DataFrame()
    for col_idx, col_name in col_mapping.items():
        if col_idx < data_slice.shape[1]:
            clean_df[col_name] = data_slice[col_idx]
        else:
            clean_df[col_name] = np.nan

    for idx in numeric_indices:
        col_name = col_mapping[idx]
        if col_name in clean_df.columns:
            clean_df[col_name] = pd.to_numeric(clean_df[col_name], errors="coerce")

    for col in clean_df.columns:
        if col not in ["Nilai Kontrak Setelah PPN", "Nilai Kontrak Mitra Setelah PPN"]:
            if col == "STATUS":
                clean_df[col] = clean_df[col].apply(
                    lambda v: "BELUM ADA STATUS"
                    if pd.isna(v) or not str(v).strip() or str(v).strip().lower() == "nan"
                    else str(v).strip()
                )
            elif col == "NO":
                clean_df[col] = clean_df[col].apply(
                    lambda v: int(v)
                    if pd.notna(v) and str(v).replace(".0", "").isdigit()
                    else (_clean_str(v) or "-")
                )
            else:
                clean_df[col] = clean_df[col].apply(lambda v: _clean_str(v, default=""))

    clean_df.reset_index(drop=True, inplace=True)
    return clean_df


def _parse_potensi(raw: pd.DataFrame) -> pd.DataFrame:
    """Parse sheet 'POTENSI 2025 DAN 2026'."""
    header_row = 3
    for r in range(min(10, len(raw))):
        c2 = str(raw.iloc[r, 2]).strip().upper() if pd.notna(raw.iloc[r, 2]) else ""
        if "JUDUL" in c2:
            header_row = r
            break

    end_row = len(raw)
    for r in range(header_row + 1, len(raw)):
        c2 = str(raw.iloc[r, 2]).strip().upper() if pd.notna(raw.iloc[r, 2]) else ""
        if c2 == "TOTAL":
            end_row = r
            break

    data_slice = raw.iloc[header_row + 1:end_row].copy()
    mask = data_slice[2].notna() & (data_slice[2].astype(str).str.strip() != "")
    data_slice = data_slice[mask]

    clean_df = pd.DataFrame()
    clean_df["JUDUL"] = data_slice[2].apply(lambda v: _clean_str(v, default=""))
    clean_df["POTENSI NILAI"] = pd.to_numeric(data_slice[3], errors="coerce").fillna(0.0)

    def _clean_periode(v: Any) -> str:
        if pd.isna(v):
            return "-"
        s = str(v).replace(".0", "").strip()
        return s if s and s.lower() != "nan" else "-"

    clean_df["PERIODE"] = data_slice[4].apply(_clean_periode) if 4 in data_slice.columns else "-"
    clean_df.reset_index(drop=True, inplace=True)
    return clean_df


HISTORI_COLUMNS = ["Timestamp", "Judul Pekerjaan", "Field", "Nilai Lama", "Nilai Baru"]


def _parse_histori(raw: pd.DataFrame) -> pd.DataFrame:
    """Parse sheet 'Histori' (change log written by excel_writer.append_histori).

    Row 0: header. Rows 1..: append-only log rows, no TOTAL marker.
    """
    if raw.empty or len(raw) < 2:
        return pd.DataFrame(columns=HISTORI_COLUMNS)

    data_slice = raw.iloc[1:].copy()
    mask = (
        data_slice[1].notna()
        & (data_slice[1].astype(str).str.strip() != "")
        & (data_slice[1].astype(str).str.strip().str.lower() != "nan")
    )
    data_slice = data_slice[mask]

    clean_df = pd.DataFrame()
    for col_idx, col_name in enumerate(HISTORI_COLUMNS):
        if col_idx < data_slice.shape[1]:
            clean_df[col_name] = data_slice[col_idx].apply(lambda v: _clean_str(v, default=""))
        else:
            clean_df[col_name] = ""

    clean_df.reset_index(drop=True, inplace=True)
    return clean_df


def _parse_penagihan(raw: pd.DataFrame) -> List[Dict[str, Any]]:
    """Parse sheet 'PENAGIHAN' (checklist items)."""
    sections: List[Dict[str, Any]] = []
    current_section: Optional[Dict[str, Any]] = None

    for r in range(len(raw)):
        c0 = raw.iloc[r, 0] if raw.shape[1] > 0 else np.nan
        c1 = raw.iloc[r, 1] if raw.shape[1] > 1 else np.nan

        c0_str = _clean_str(c0)
        c1_str = _clean_str(c1)

        if c0_str:
            current_section = {
                "letter": c0_str,
                "title": c1_str,
                "items": []
            }
            sections.append(current_section)
        elif c1_str and current_section is not None:
            current_section["items"].append(c1_str)

    return sections


# Indonesian month dictionary for parsing
ID_MONTHS: Dict[str, int] = {
    "januari": 1, "jan": 1,
    "februari": 2, "feb": 2,
    "maret": 3, "mar": 3,
    "april": 4, "apr": 4,
    "mei": 5,
    "juni": 6, "jun": 6,
    "juli": 7, "jul": 7,
    "agustus": 8, "agu": 8, "ags": 8,
    "september": 9, "sep": 9,
    "oktober": 10, "okt": 10,
    "november": 11, "nov": 11,
    "desember": 12, "des": 12,
}


def parse_id_date(text: Any) -> Optional[date]:
    """Extract the first date-like pattern from text.

    Supports:
      - 'd Month yyyy' with Indonesian month names (full or 3-letter)
      - dd/mm/yyyy or dd-mm-yyyy
      - yyyy-mm-dd
      - pandas/python datetime objects

    Returns datetime.date or None.
    """
    if text is None or pd.isna(text):
        return None

    if isinstance(text, (datetime, pd.Timestamp)):
        return text.date()
    if isinstance(text, date):
        return text

    s = str(text).strip()
    if not s or s.lower() == "nan":
        return None

    # Try 'd Month yyyy' e.g. '04 September 2026' or '4 sep 2026'
    m_id = re.search(r"\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b", s)
    if m_id:
        d_str, mon_str, y_str = m_id.groups()
        mon_lower = mon_str.lower()
        if mon_lower in ID_MONTHS:
            try:
                return date(int(y_str), ID_MONTHS[mon_lower], int(d_str))
            except ValueError:
                pass

    # Try dd/mm/yyyy or dd-mm-yyyy
    m_dmy = re.search(r"\b(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\b", s)
    if m_dmy:
        g1, g2, g3 = m_dmy.groups()
        try:
            return date(int(g3), int(g2), int(g1))
        except ValueError:
            pass

    # Try yyyy-mm-dd
    m_ymd = re.search(r"\b(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})\b", s)
    if m_ymd:
        g1, g2, g3 = m_ymd.groups()
        try:
            return date(int(g1), int(g2), int(g3))
        except ValueError:
            pass

    return None


def parse_duration_days(text: Any) -> Optional[int]:
    """Parse duration in days from text patterns like '14 hari kalender', '30 Hari', '20 (Dua Puluh) hari'.

    Returns int or None.
    """
    if text is None or pd.isna(text):
        return None
    s = str(text).strip()
    if not s or s.lower() == "nan":
        return None

    # Matches: '14 hari', '30 Hari kalender', '20 (Dua Puluh) hari kalender'
    m = re.search(r"(\d+)\s*(?:\([^)]*\))?\s*hari\b", s, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None


def load_all(path: str) -> Dict[str, Any]:
    """Load and parse all sheets from Excel file.

    Returns dict with keys:
      - monitoring (pd.DataFrame)
      - pencapaian (dict)
      - akumulatif_kpi (dict)
      - monitoring_akumulatif (pd.DataFrame)
      - potensi (pd.DataFrame)
      - penagihan (list)
      - file_mtime (float)
    """
    with pd.ExcelFile(path) as excel:
        sheet_names = excel.sheet_names
        sheet_map = {name.strip(): name for name in sheet_names}

        def _get_sheet_df(clean_name: str) -> pd.DataFrame:
            actual_name = sheet_map.get(clean_name)
            if actual_name:
                return excel.parse(actual_name, header=None)
            return pd.DataFrame()

        # Sheet 'Monitoring'
        raw_mon = _get_sheet_df("Monitoring")
        monitoring = _parse_monitoring(raw_mon) if not raw_mon.empty else pd.DataFrame()

        # Sheet 'Monitoring Akumulatif '
        raw_mon_ak = _get_sheet_df("Monitoring Akumulatif")
        monitoring_akumulatif = _parse_monitoring_akumulatif(raw_mon_ak) if not raw_mon_ak.empty else pd.DataFrame()

        # Compute realisasi sums
        mon_realisasi_sum = (
            float(monitoring["Nilai Kontrak Setelah PPN (IPS)"].sum())
            if "Nilai Kontrak Setelah PPN (IPS)" in monitoring.columns
            else 0.0
        )
        ak_realisasi_sum = (
            float(monitoring_akumulatif["Nilai Kontrak Mitra Setelah PPN"].sum())
            if "Nilai Kontrak Mitra Setelah PPN" in monitoring_akumulatif.columns
            else 0.0
        )

        # Sheet 'PENCAPAIAN'
        raw_pencapaian = _get_sheet_df("PENCAPAIAN")
        pencapaian = (
            _parse_pencapaian(raw_pencapaian, mon_realisasi_sum)
            if not raw_pencapaian.empty
            else {}
        )

        # Sheet 'Grafik akumulatif'
        raw_grafik = _get_sheet_df("Grafik akumulatif")
        akumulatif_kpi = (
            _parse_akumulatif_kpi(raw_grafik, ak_realisasi_sum)
            if not raw_grafik.empty
            else {}
        )

        # Sheet 'POTENSI 2025 DAN 2026'
        raw_potensi = _get_sheet_df("POTENSI 2025 DAN 2026")
        potensi = _parse_potensi(raw_potensi) if not raw_potensi.empty else pd.DataFrame()

        # Sheet 'PENAGIHAN'
        raw_penagihan = _get_sheet_df("PENAGIHAN")
        penagihan = _parse_penagihan(raw_penagihan) if not raw_penagihan.empty else []

        # Sheet 'Histori' (change log, may not exist yet on older workbooks)
        raw_histori = _get_sheet_df("Histori")
        histori = _parse_histori(raw_histori) if not raw_histori.empty else pd.DataFrame(columns=HISTORI_COLUMNS)

    file_mtime = os.path.getmtime(path) if os.path.exists(path) else 0.0

    return {
        "monitoring": monitoring,
        "pencapaian": pencapaian,
        "akumulatif_kpi": akumulatif_kpi,
        "monitoring_akumulatif": monitoring_akumulatif,
        "potensi": potensi,
        "penagihan": penagihan,
        "histori": histori,
        "file_mtime": file_mtime,
    }
