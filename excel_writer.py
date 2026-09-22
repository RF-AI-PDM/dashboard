"""Excel writer module for Dashboard Monitoring Pekerjaan.

Writes back to data/Data.xlsx using openpyxl while preserving formulas and styles.
All source code must be ASCII-only.
"""

from __future__ import annotations

from copy import copy
from datetime import datetime
import math
import os
import re
import shutil
from typing import Any, Dict, List, Optional
import openpyxl
import pandas as pd


class ExcelLocked(Exception):
    """Raised when the Excel file is locked or cannot be written."""
    pass


def backup_excel(path: str, backup_dir: Optional[str] = None) -> str:
    """Create a backup of the Excel file, keeping only the 20 newest backups."""
    if backup_dir is None:
        backup_dir = os.path.join(os.path.dirname(path), "backup")
    os.makedirs(backup_dir, exist_ok=True)

    stem = os.path.splitext(os.path.basename(path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"{stem}_{timestamp}.xlsx"
    backup_path = os.path.join(backup_dir, backup_filename)

    counter = 1
    while os.path.exists(backup_path):
        backup_path = os.path.join(backup_dir, f"{stem}_{timestamp}_{counter}.xlsx")
        counter += 1

    shutil.copy2(path, backup_path)

    # Prune old backups for this stem, keeping only newest 20
    matching_files = [
        os.path.join(backup_dir, f)
        for f in os.listdir(backup_dir)
        if f.startswith(f"{stem}_") and f.endswith(".xlsx") and os.path.isfile(os.path.join(backup_dir, f))
    ]
    matching_files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    if len(matching_files) > 20:
        for old_file in matching_files[20:]:
            try:
                os.remove(old_file)
            except OSError:
                pass

    return backup_path


def _find_total_row(ws: Any, col_letter: str, start_row: int) -> int:
    """Find row number where cell value in col_letter equals 'TOTAL'."""
    for r in range(start_row, ws.max_row + 1):
        val = ws[f"{col_letter}{r}"].value
        if val is not None and str(val).strip().upper() == "TOTAL":
            return r
    raise ValueError(f"Baris TOTAL tidak ditemukan di kolom {col_letter} mulai baris {start_row}")


def _ensure_slots(
    ws: Any,
    first_data_row: int,
    total_row: int,
    needed: int,
    merged_total_range: Optional[str] = None
) -> int:
    """Ensure sufficient row slots before the TOTAL row.

    If needed > (total_row - first_data_row), inserts rows before TOTAL.
    Unmerges and re-merges the total range, and copies styles from the last data row.
    Returns the new total_row.
    """
    slots = total_row - first_data_row
    if needed <= slots:
        return total_row

    amount = needed - slots
    actual_merged: Optional[tuple[int, int]] = None

    # Check and unmerge total range if present
    if merged_total_range:
        for rng in list(ws.merged_cells.ranges):
            if str(rng) == merged_total_range or (rng.min_row == total_row and rng.max_row == total_row):
                actual_merged = (rng.min_col, rng.max_col)
                ws.unmerge_cells(str(rng))
                break

    last_data_row = total_row - 1
    ws.insert_rows(total_row, amount)
    new_total_row = total_row + amount

    # Re-merge at the new total row
    if actual_merged:
        min_col, max_col = actual_merged
        min_letter = openpyxl.utils.get_column_letter(min_col)
        max_letter = openpyxl.utils.get_column_letter(max_col)
        ws.merge_cells(f"{min_letter}{new_total_row}:{max_letter}{new_total_row}")

    # Copy styles from last_data_row to each newly inserted row
    for r in range(total_row, new_total_row):
        for c in range(1, ws.max_column + 1):
            src_cell = ws.cell(row=last_data_row, column=c)
            tgt_cell = ws.cell(row=r, column=c)
            if src_cell.has_style:
                tgt_cell.font = copy(src_cell.font)
                tgt_cell.border = copy(src_cell.border)
                tgt_cell.fill = copy(src_cell.fill)
                tgt_cell.number_format = copy(src_cell.number_format)
                tgt_cell.alignment = copy(src_cell.alignment)

    return new_total_row


def _rewrite_total_formulas(
    ws: Any,
    total_row: int,
    first_data_row: int,
    numeric_col_letters: List[str]
) -> None:
    """Rewrite SUM formulas on the TOTAL row for specified numeric columns."""
    for col in numeric_col_letters:
        ws[f"{col}{total_row}"] = f"=SUM({col}{first_data_row}:{col}{total_row - 1})"


def _clean_num_val(v: Any) -> Optional[float]:
    """Convert value to float if valid numeric, else None."""
    if v is None or pd.isna(v):
        return None
    try:
        val = float(v)
        if math.isnan(val) or math.isinf(val):
            return None
        return val
    except (ValueError, TypeError):
        return None


def _clean_str_val(v: Any) -> Optional[str]:
    """Convert value to stripped string if non-empty, else None."""
    if v is None or pd.isna(v):
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    return s


# Monitoring column definition: (1-based Excel column index, column letter, column name, is_numeric)
MONITORING_COLS = [
    (2, "B", "NO", False),
    (3, "C", "Judul Pekerjaan", False),
    (4, "D", "Projeck", False),
    (5, "E", "User", False),
    (6, "F", "Proses Kontrak", False),
    (7, "G", "Eksekutor", False),
    (8, "H", "DMR/TOR/Spesifikasi", False),
    (9, "I", "LOI", False),
    (10, "J", "Nomor Informasi Harga Mitra", False),
    (11, "K", "IH Sebelum PPN", True),
    (12, "L", "IH Setelah PPN", True),
    (13, "M", "Undangan PL", False),
    (14, "N", "Kontrak IPS", False),
    (15, "O", "Nilai Kontrak Sebelum PPN (IPS)", True),
    (16, "P", "Nilai Kontrak Setelah PPN (IPS)", True),
    (17, "Q", "Mulai Pekerjaan", False),
    (18, "R", "Berakhir Pekerjaan", False),
    (19, "S", "Addendum Kontrak", False),
    (20, "T", "Ket Addendum", False),
    (21, "U", "BAST/TUG", False),
    (22, "V", "Invoice", False),
    (23, "W", "Nilai Invoice Sebelum PPN", True),
    (24, "X", "Nilai Invoice Setelah PPN", True),
    (25, "Y", "DENDA", True),
    (26, "Z", "PEMBAYARAN PIPS", True),
    (27, "AA", "Nama Vendor", False),
    (28, "AB", "Nomor Kontrak Vendor", False),
    (29, "AC", "Nilai Kontrak Mitra Sebelum PPN", True),
    (30, "AD", "Nilai Kontrak Mitra Setelah PPN", True),
    (31, "AE", "BA Vendor", False),
    (32, "AF", "Nilai BA Vendor Setelah PPN", True),
    (33, "AG", "DENDA2", True),
    (34, "AH", "PEMBAYARAN MITRA", True),
    (35, "AI", "MARGIN", True),
    (36, "AJ", "Kendala", False),
    (37, "AK", "Periode Penagihan", False),
    (38, "AL", "STATUS", False),
    (39, "AM", "STATUS PEMBAYARAN", False),
]

MONITORING_NUMERIC_TOTAL_COLS = ["K", "L", "O", "P", "W", "X", "Y", "Z", "AC", "AD", "AF"]


def write_monitoring(path: str, df: pd.DataFrame, backup_dir: Optional[str] = None) -> str:
    """Write DataFrame back to sheet 'Monitoring' in Excel file.

    Performs automatic backup, expands row slots if needed, clears previous data,
    writes updated rows with renumbered NO, updates TOTAL row SUM formulas, and
    updates formula references to Monitoring!P<row> across other sheets.
    """
    backup_path = backup_excel(path, backup_dir=backup_dir)

    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e

    ws = wb["Monitoring"]
    first_data_row = 5
    old_total_row = _find_total_row(ws, "B", first_data_row)

    total_row = _ensure_slots(
        ws,
        first_data_row,
        old_total_row,
        len(df),
        merged_total_range=f"B{old_total_row}:H{old_total_row}"
    )

    # Ensure row 4 is template row NO=0 with empty title so it is kept as template
    ws.cell(row=4, column=2).value = 0
    ws.cell(row=4, column=3).value = None

    # Clear columns B..AM from row 5 to total_row - 1
    for r in range(first_data_row, total_row):
        for col_idx, col_letter, _, _ in MONITORING_COLS:
            ws.cell(row=r, column=col_idx).value = None

    # Write rows in df order
    for row_i, (_, row_data) in enumerate(df.iterrows()):
        curr_row = first_data_row + row_i
        for col_idx, col_letter, col_name, is_numeric in MONITORING_COLS:
            if col_name == "NO":
                ws.cell(row=curr_row, column=col_idx).value = row_i + 1
            elif is_numeric:
                val = _clean_num_val(row_data.get(col_name))
                ws.cell(row=curr_row, column=col_idx).value = val
            else:
                raw_val = row_data.get(col_name)
                s_val = _clean_str_val(raw_val)
                if col_name in ("STATUS", "STATUS PEMBAYARAN"):
                    if s_val == "BELUM ADA STATUS":
                        s_val = None
                ws.cell(row=curr_row, column=col_idx).value = s_val

    # Rewrite total formulas
    _rewrite_total_formulas(ws, total_row, first_data_row, MONITORING_NUMERIC_TOTAL_COLS)

    # If total_row moved or changed, update formula references in all sheets
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for r in range(1, sheet.max_row + 1):
            for c in range(1, sheet.max_column + 1):
                cell_val = sheet.cell(row=r, column=c).value
                if isinstance(cell_val, str) and cell_val.startswith("="):
                    if "Monitoring!P" in cell_val:
                        new_val = re.sub(r"Monitoring!P\d+", f"Monitoring!P{total_row}", cell_val)
                        if new_val != cell_val:
                            sheet.cell(row=r, column=c).value = new_val

    try:
        wb.save(path)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e
    finally:
        wb.close()

    return backup_path


AKUMULATIF_COLS = [
    (2, "B", "NO", False),
    (3, "C", "Judul Pekerjaan", False),
    (4, "D", "User", False),
    (5, "E", "Kontrak", False),
    (6, "F", "Nilai Kontrak Setelah PPN", True),
    (7, "G", "Nama Vendor", False),
    (8, "H", "Kontrak Vendor", False),
    (9, "I", "Nilai Kontrak Mitra Setelah PPN", True),
    (10, "J", "STATUS", False),
]

AKUMULATIF_TOTAL_COLS = ["F", "G", "H", "I"]


def write_monitoring_akumulatif(path: str, df: pd.DataFrame, backup_dir: Optional[str] = None) -> str:
    """Write DataFrame back to sheet 'Monitoring Akumulatif '."""
    backup_path = backup_excel(path, backup_dir=backup_dir)

    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e

    sheet_name = "Monitoring Akumulatif " if "Monitoring Akumulatif " in wb.sheetnames else "Monitoring Akumulatif"
    ws = wb[sheet_name]
    first_data_row = 2
    old_total_row = _find_total_row(ws, "B", first_data_row)

    total_row = _ensure_slots(
        ws,
        first_data_row,
        old_total_row,
        len(df),
        merged_total_range=f"B{old_total_row}:D{old_total_row}"
    )

    # Clear columns B..J from row 2 to total_row - 1
    for r in range(first_data_row, total_row):
        for col_idx, _, _, _ in AKUMULATIF_COLS:
            ws.cell(row=r, column=col_idx).value = None

    # Write rows
    for row_i, (_, row_data) in enumerate(df.iterrows()):
        curr_row = first_data_row + row_i
        for col_idx, _, col_name, is_numeric in AKUMULATIF_COLS:
            if col_name == "NO":
                ws.cell(row=curr_row, column=col_idx).value = row_i + 1
            elif is_numeric:
                ws.cell(row=curr_row, column=col_idx).value = _clean_num_val(row_data.get(col_name))
            else:
                s_val = _clean_str_val(row_data.get(col_name))
                if col_name == "STATUS" and s_val == "BELUM ADA STATUS":
                    s_val = None
                ws.cell(row=curr_row, column=col_idx).value = s_val

    # Rewrite totals for F, G, H, I
    _rewrite_total_formulas(ws, total_row, first_data_row, AKUMULATIF_TOTAL_COLS)

    # Update references in formulas e.g. 'Monitoring Akumulatif '!I8
    for s_name in wb.sheetnames:
        sheet = wb[s_name]
        for r in range(1, sheet.max_row + 1):
            for c in range(1, sheet.max_column + 1):
                cell_val = sheet.cell(row=r, column=c).value
                if isinstance(cell_val, str) and cell_val.startswith("="):
                    if "Monitoring Akumulatif" in cell_val and "!I" in cell_val:
                        new_val = re.sub(r"('?Monitoring Akumulatif\s*'?\s*!\s*I)\d+", rf"\g<1>{total_row}", cell_val)
                        if new_val != cell_val:
                            sheet.cell(row=r, column=c).value = new_val

    try:
        wb.save(path)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e
    finally:
        wb.close()

    return backup_path


def write_potensi(path: str, df: pd.DataFrame, backup_dir: Optional[str] = None) -> str:
    """Write DataFrame back to sheet 'POTENSI 2025 DAN 2026'."""
    backup_path = backup_excel(path, backup_dir=backup_dir)

    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e

    sheet_name = "POTENSI 2025 DAN 2026"
    ws = wb[sheet_name]
    first_data_row = 5
    old_total_row = _find_total_row(ws, "C", first_data_row)

    total_row = _ensure_slots(ws, first_data_row, old_total_row, len(df), merged_total_range=None)

    # Clear columns C..E from row 5 to total_row - 1
    for r in range(first_data_row, total_row):
        ws.cell(row=r, column=3).value = None
        ws.cell(row=r, column=4).value = None
        ws.cell(row=r, column=5).value = None

    # Write data
    for row_i, (_, row_data) in enumerate(df.iterrows()):
        curr_row = first_data_row + row_i
        ws.cell(row=curr_row, column=3).value = _clean_str_val(row_data.get("JUDUL"))
        ws.cell(row=curr_row, column=4).value = _clean_num_val(row_data.get("POTENSI NILAI"))
        ws.cell(row=curr_row, column=5).value = _clean_str_val(row_data.get("PERIODE"))

    # Rewrite total formula for D
    ws[f"D{total_row}"] = f"=SUM(D{first_data_row}:D{total_row - 1})"

    try:
        wb.save(path)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e
    finally:
        wb.close()

    return backup_path


def write_targets(
    path: str,
    pencapaian: Optional[Dict[str, Any]] = None,
    akumulatif: Optional[Dict[str, Any]] = None,
    backup_dir: Optional[str] = None
) -> str:
    """Update target constants in sheet 'PENCAPAIAN' and 'Grafik akumulatif'."""
    backup_path = backup_excel(path, backup_dir=backup_dir)

    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e

    # Update PENCAPAIAN
    if pencapaian and "PENCAPAIAN" in wb.sheetnames:
        ws_p = wb["PENCAPAIAN"]
        if "target_tahun" in pencapaian:
            ws_p["C6"] = _clean_num_val(pencapaian["target_tahun"])
        if "target_s1" in pencapaian:
            ws_p["E6"] = _clean_num_val(pencapaian["target_s1"])
        if "target_s2" in pencapaian:
            ws_p["G6"] = _clean_num_val(pencapaian["target_s2"])

    # Update Grafik akumulatif
    if akumulatif and "Grafik akumulatif" in wb.sheetnames:
        ws_g = wb["Grafik akumulatif"]
        if "target_tahun" in akumulatif:
            ws_g["C6"] = _clean_num_val(akumulatif["target_tahun"])
        if "target_s1" in akumulatif:
            ws_g["E6"] = _clean_num_val(akumulatif["target_s1"])
        if "capaian_s1" in akumulatif:
            ws_g["F6"] = _clean_num_val(akumulatif["capaian_s1"])
        if "target_s2" in akumulatif:
            ws_g["G6"] = _clean_num_val(akumulatif["target_s2"])

    try:
        wb.save(path)
    except PermissionError as e:
        raise ExcelLocked("Tutup file Excel terlebih dahulu") from e
    finally:
        wb.close()

    return backup_path
