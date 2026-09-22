# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-user Streamlit dashboard (Bahasa Indonesia UI) for monitoring procurement/contract work at PLN Indonesia Power Services. The only data source is `data/Data.xlsx`; the app both reads it and writes edits back to it. There is no git repo, no test suite, and no web backend — everything is local Python.

## Commands

All commands assume the project venv at `.\venv` (Python 3.12). On Windows use `venv\Scripts\python.exe`; there is no system-wide `python` guaranteed on PATH.

```powershell
# Run the app (or just double-click run.bat — it creates the venv and installs deps if missing)
venv\Scripts\python.exe -m streamlit run app.py --browser.gatherUsageStats false

# Install / update dependencies
venv\Scripts\python.exe -m pip install -r requirements.txt

# Syntax check every module
venv\Scripts\python.exe -m py_compile app.py data_loader.py excel_writer.py ppt_export.py theme.py

# Headless smoke test of the whole app (no browser). Must print EXC: []
venv\Scripts\python.exe -c "from streamlit.testing.v1 import AppTest; at=AppTest.from_file(r'D:\Artefact\DASHBOARD\app.py', default_timeout=180); at.run(); print('EXC:', [str(e.value) for e in at.exception])"

# Verify loader numbers against the known Excel totals
venv\Scripts\python.exe -c "import data_loader as dl; d=dl.load_all('data/Data.xlsx'); assert abs(d['pencapaian']['realisasi_tahun']-16618722076)<1; assert abs(d['akumulatif_kpi']['realisasi_tahun']-17222968236)<1; print('ok')"

# Build the PPT without the UI
venv\Scripts\python.exe -c "import data_loader as dl, ppt_export; d=dl.load_all('data/Data.xlsx'); open('_t.pptx','wb').write(ppt_export.build_pptx(d, d['monitoring']))"
```

`AppTest.from_file` needs an **absolute** path; it resolves relative paths against the calling script, not the CWD.

## Hard rules

- **Never write to `data/Data.xlsx` during tests.** Copy it to `data/_test_copy.xlsx`, pass `backup_dir="data/_test_backup"` to the writer functions, and delete both afterward. The writer functions all accept `backup_dir`. To exercise the UI save flows (Simpan buttons, Tambah Pekerjaan form) through `AppTest`, set `DASHBOARD_EXCEL=<abs path to the copy>` in the environment before importing the app; `app.py` reads that env var for `EXCEL_PATH`.
- **All `.py` source must be ASCII-only** (no emoji, no unicode symbols in strings or comments). The files were previously corrupted by a PowerShell `Get-Content`/`Set-Content` round-trip that re-encoded UTF-8; ASCII-only makes that impossible. Verify with a byte scan, not `grep -P` (Git Bash grep lacks UTF-8 locale support here).
- Do not use PowerShell `Get-Content ... | Set-Content` to bulk-edit source files. Use the Edit tool or a Python script.
- `use_container_width=True` is deprecated in the installed Streamlit; use `width='stretch'`.
- A `~$Data.xlsx` lock file in `data/` means Excel has the workbook open; any save from the app raises `excel_writer.ExcelLocked`. The UI already surfaces this; do not "fix" it by deleting the lock file.

## Architecture

Five modules, one direction of data flow, Excel as the single source of truth:

```text
data/Data.xlsx  --read-->  data_loader.load_all()  -->  app.py (Streamlit)  -->  ppt_export.build_pptx()
      ^                                                       |
      +------------------ excel_writer.write_*() <------------+   (Kelola Data tab)
```

### `data_loader.py` — the only reader

`load_all(path)` returns one dict: `monitoring` (DataFrame), `pencapaian` (dict), `akumulatif_kpi` (dict), `monitoring_akumulatif`, `potensi`, `penagihan` (list of sections), `file_mtime`. Everything else in the app consumes this dict.

Key behaviours that are easy to break:

- **Sheets are read with `header=None`** and parsed by fixed 0-based row/column indices. Excel row `r` is pandas index `r-1`; Excel column `B` is pandas column `1`. The column-name mapping for the Monitoring sheet lives in `_parse_monitoring`; the same mapping must stay consistent with the column-letter list in `excel_writer.write_monitoring`.
- **Data rows are delimited by the `TOTAL` marker**, not by hard-coded counts, so the loader survives row insertions. Template/marker rows (Monitoring rows 2–4 in Excel) are skipped by the "Judul Pekerjaan non-empty" filter.
- **KPI realisasi/ratio/deviasi are computed in Python, never read from formula cells.** After `openpyxl` saves a workbook, formula cells have no cached value and pandas reads them as `NaN`. Only the constant target cells (`PENCAPAIAN!C6/E6/G6`, `Grafik akumulatif!C6/E6/F6/G6`) are read from Excel. `deviasi = realisasi - target` (negative when under target).
- **`_normalize_status`** collapses the free-text `STATUS` column to `SELESAI / PROSES / CANCEL / BELUM ADA STATUS` and moves any non-canonical text into `Kendala`. Originals are preserved in `STATUS (Asli)` / `STATUS PEMBAYARAN (Asli)`. These two helper columns exist only in memory; the writer drops them.
- `parse_id_date` / `parse_duration_days` handle Indonesian date text (`"04 September 2026"`, `"14 hari kalender"`) for the Gantt chart. Most rows in the current data have no parseable dates, so the Gantt is expected to be sparse.

### `excel_writer.py` — the only writer

Loads with `openpyxl.load_workbook(path)` (**not** `data_only=True`) so formulas and styles survive. Every `write_*` call: backs up to `data/backup/` (keeps 20 newest), locates the `TOTAL` row by scanning, inserts rows above it if the DataFrame outgrew the slots (unmerging and re-merging the TOTAL row's merged range, because openpyxl does not move merges on `insert_rows`), clears and rewrites the data block, rewrites the `TOTAL` row's `=SUM(...)` formulas, and — if the TOTAL row moved — rewrites every cross-sheet formula string in the workbook that referenced the old row (e.g. `Monitoring!P38` → `Monitoring!P69`). `PermissionError` on save is re-raised as `ExcelLocked`.

Writer functions: `write_monitoring`, `write_monitoring_akumulatif`, `write_potensi`, `write_targets`. Each returns the backup path. The sheet name `"Monitoring Akumulatif "` has a **trailing space** in the workbook; match it with `.strip()` or the exact literal.

### `app.py` — UI

One file, seven top-level tabs (`Ringkasan`, `Monitoring Pekerjaan`, `Akumulatif`, `Potensi Pendapatan`, `Alur Penagihan`, `Kelola Data`, `Export PPT`); `Kelola Data` has six sub-tabs (Monitoring editor, Tambah Pekerjaan form, Akumulatif editor, Potensi editor, Target, Cadangan/restore).

- Data is loaded once via `@st.cache_data` keyed on `(path, mtime)`, so editing the Excel file externally and pressing "Muat Ulang Data" (which calls `st.cache_data.clear()`) picks up changes.
- **Sidebar filter widgets are keyed** (`f_search`, `f_user`, `f_proses`, `f_eksekutor`, `f_status`, `f_bayar`, `f_vendor`). Streamlit forbids setting a widget's `session_state` after that widget has been instantiated in the same run, so two indirection flags exist: `_reset_filters` (set by the Reset button) and `_pending_filter` (set by chart click handlers). Both are consumed at the top of the sidebar block **before** the multiselects are created. Any new "set a filter programmatically" feature must go through `_pending_filter`, not assign to `f_*` directly.
- Chart drill-down uses `st.plotly_chart(..., on_select="rerun", selection_mode="points")`; the selected category name is pushed into `_pending_filter`.
- Before every save in `Kelola Data`, the app compares `data["file_mtime"]` with the current `os.path.getmtime` and refuses to write if the file changed underneath it (stale-edit guard).
- Every save ends with `st.cache_data.clear(); st.rerun()`, which wipes anything rendered in that run. Success feedback therefore goes through `st.session_state["_flash"]`, displayed once just above the tab bar on the next run. Do not call `st.success` directly before a rerun.
- The Monitoring editor overlays edited rows onto the full DataFrame **by index label**, not by position: `st.data_editor` with `num_rows="dynamic"` keeps original labels for surviving rows and assigns fresh labels to added rows, so positional matching corrupts hidden columns after a deletion.
- The Monitoring editor shows a column subset; hidden columns are merged back onto the full DataFrame by row position before writing, so adding a visible column to the editor must not reorder rows.

### `ppt_export.py`

`build_pptx(data, monitoring_df, gantt_png=None) -> bytes`. Uses python-pptx **native** charts (`CategoryChartData`) and tables, not images, so the output is editable in PowerPoint. `monitoring_df` is the *filtered* frame from the sidebar, so the PPT reflects the current filter. The optional Gantt slide only renders if `kaleido` is importable; it is deliberately not in `requirements.txt`.

### `theme.py`

Single place for brand colors (PLN/IPS navy `#0A3D8F` + yellow `#FDB913`), `STATUS_COLORS` / `PAYMENT_COLORS` maps keyed by the canonical status strings from `data_loader`, `ORG_NAME`, and `LOGO_PATH` (`data/logo.png`, optional — sidebar and PPT fall back to a text badge). `app.py` and `ppt_export.py` both import from here; `.streamlit/config.toml` duplicates the primary color for Streamlit's own widgets and must be kept in sync by hand.
