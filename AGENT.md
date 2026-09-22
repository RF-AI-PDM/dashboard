# AGENT.md

Guidance for AI coding agents working in this repo. (See also `CLAUDE.md`, same content, Claude-specific framing.)

## What this is

Single-user Streamlit dashboard (Bahasa Indonesia UI) for monitoring procurement/contract work at PLN Indonesia Power Services. Only data source: `data/Data.xlsx`; app reads it and writes edits back. No git repo, no test suite, no web backend — everything local Python.

## Commands

Assume project venv at `.\venv` (Python 3.12). Windows: use `venv\Scripts\python.exe`; no guaranteed system-wide `python` on PATH.

```powershell
# Run the app (or double-click run.bat — creates venv and installs deps if missing)
venv\Scripts\python.exe -m streamlit run app.py --browser.gatherUsageStats false

# Install / update dependencies
venv\Scripts\python.exe -m pip install -r requirements.txt

# Syntax check every module
venv\Scripts\python.exe -m py_compile app.py data_loader.py excel_writer.py ppt_export.py theme.py

# Headless smoke test of whole app (no browser). Must print EXC: []
venv\Scripts\python.exe -c "from streamlit.testing.v1 import AppTest; at=AppTest.from_file(r'D:\Artefact\DASHBOARD\app.py', default_timeout=180); at.run(); print('EXC:', [str(e.value) for e in at.exception])"

# Verify loader numbers against known Excel totals
venv\Scripts\python.exe -c "import data_loader as dl; d=dl.load_all('data/Data.xlsx'); assert abs(d['pencapaian']['realisasi_tahun']-16618722076)<1; assert abs(d['akumulatif_kpi']['realisasi_tahun']-17222968236)<1; print('ok')"

# Build PPT without UI
venv\Scripts\python.exe -c "import data_loader as dl, ppt_export; d=dl.load_all('data/Data.xlsx'); open('_t.pptx','wb').write(ppt_export.build_pptx(d, d['monitoring']))"
```

`AppTest.from_file` needs **absolute** path; resolves relative paths against calling script, not CWD.

## Hard rules

- **Never write to `data/Data.xlsx` during tests.** Copy to `data/_test_copy.xlsx`, pass `backup_dir="data/_test_backup"` to writer functions, delete both after. Writer functions all accept `backup_dir`.
- **All `.py` source must be ASCII-only** (no emoji, no unicode symbols in strings or comments). Files previously corrupted by PowerShell `Get-Content`/`Set-Content` round-trip re-encoding UTF-8; ASCII-only makes that impossible. Verify with byte scan, not `grep -P` (Git Bash grep lacks UTF-8 locale support here).
- Don't use PowerShell `Get-Content ... | Set-Content` to bulk-edit source files. Use proper edit tool or Python script.
- `use_container_width=True` deprecated in installed Streamlit; use `width='stretch'`.
- A `~$Data.xlsx` lock file in `data/` means Excel has workbook open; any save from app raises `excel_writer.ExcelLocked`. UI already surfaces this; don't "fix" by deleting lock file.

## Architecture

Five modules, one direction of data flow, Excel as single source of truth:

```
data/Data.xlsx  --read-->  data_loader.load_all()  -->  app.py (Streamlit)  -->  ppt_export.build_pptx()
      ^                                                       |
      +------------------ excel_writer.write_*() <------------+   (Kelola Data tab)
```

### `data_loader.py` — only reader

`load_all(path)` returns one dict: `monitoring` (DataFrame), `pencapaian` (dict), `akumulatif_kpi` (dict), `monitoring_akumulatif`, `potensi`, `penagihan` (list of sections), `file_mtime`. Everything else in app consumes this dict.

Key behaviours easy to break:

- **Sheets read with `header=None`**, parsed by fixed 0-based row/column indices. Excel row `r` is pandas index `r-1`; Excel column `B` is pandas column `1`. Column-name mapping for Monitoring sheet lives in `_parse_monitoring`; must stay consistent with column-letter list in `excel_writer.write_monitoring`.
- **Data rows delimited by `TOTAL` marker**, not hard-coded counts, so loader survives row insertions. Template/marker rows (Monitoring rows 2-4 in Excel) skipped by "Judul Pekerjaan non-empty" filter.
- **KPI realisasi/ratio/deviasi computed in Python, never read from formula cells.** After `openpyxl` saves workbook, formula cells have no cached value and pandas reads them as `NaN`. Only constant target cells (`PENCAPAIAN!C6/E6/G6`, `Grafik akumulatif!C6/E6/F6/G6`) read from Excel. `deviasi = realisasi - target` (negative when under target).
- **`_normalize_status`** collapses free-text `STATUS` column to `SELESAI / PROSES / CANCEL / BELUM ADA STATUS`, moves any non-canonical text into `Kendala`. Originals preserved in `STATUS (Asli)` / `STATUS PEMBAYARAN (Asli)`. These two helper columns exist only in memory; writer drops them.
- `parse_id_date` / `parse_duration_days` handle Indonesian date text (`"04 September 2026"`, `"14 hari kalender"`) for Gantt chart. Most rows in current data have no parseable dates, so Gantt expected sparse.

### `excel_writer.py` — only writer

Loads with `openpyxl.load_workbook(path)` (**not** `data_only=True`) so formulas and styles survive. Every `write_*` call: backs up to `data/backup/` (keeps 20 newest), locates `TOTAL` row by scanning, inserts rows above it if DataFrame outgrew slots (unmerging and re-merging TOTAL row's merged range, since openpyxl doesn't move merges on `insert_rows`), clears and rewrites data block, rewrites `TOTAL` row's `=SUM(...)` formulas, and — if TOTAL row moved — rewrites every cross-sheet formula string in workbook that referenced old row (e.g. `Monitoring!P38` -> `Monitoring!P69`). `PermissionError` on save re-raised as `ExcelLocked`.

Writer functions: `write_monitoring`, `write_monitoring_akumulatif`, `write_potensi`, `write_targets`. Each returns backup path. Sheet name `"Monitoring Akumulatif "` has **trailing space** in workbook; match with `.strip()` or exact literal.

### `app.py` — UI

One file, seven top-level tabs (`Ringkasan`, `Monitoring Pekerjaan`, `Akumulatif`, `Potensi Pendapatan`, `Alur Penagihan`, `Kelola Data`, `Export PPT`); `Kelola Data` has six sub-tabs (Monitoring editor, Tambah Pekerjaan form, Akumulatif editor, Potensi editor, Target, Cadangan/restore).

- Data loaded once via `@st.cache_data` keyed on `(path, mtime)`, so editing Excel file externally and pressing "Muat Ulang Data" (calls `st.cache_data.clear()`) picks up changes.
- **Sidebar filter widgets keyed** (`f_search`, `f_user`, `f_proses`, `f_eksekutor`, `f_status`, `f_bayar`, `f_vendor`). Streamlit forbids setting widget's `session_state` after that widget instantiated in same run, so two indirection flags exist: `_reset_filters` (set by Reset button) and `_pending_filter` (set by chart click handlers). Both consumed at top of sidebar block **before** multiselects created. Any new "set filter programmatically" feature must go through `_pending_filter`, not assign to `f_*` directly.
- Chart drill-down uses `st.plotly_chart(..., on_select="rerun", selection_mode="points")`; selected category name pushed into `_pending_filter`.
- Before every save in `Kelola Data`, app compares `data["file_mtime"]` with current `os.path.getmtime` and refuses write if file changed underneath it (stale-edit guard).
- Monitoring editor shows column subset; hidden columns merged back onto full DataFrame by row position before writing, so adding visible column to editor must not reorder rows.

### `ppt_export.py`

`build_pptx(data, monitoring_df, gantt_png=None) -> bytes`. Uses python-pptx **native** charts (`CategoryChartData`) and tables, not images, so output editable in PowerPoint. `monitoring_df` is the *filtered* frame from sidebar, so PPT reflects current filter. Optional Gantt slide only renders if `kaleido` importable; deliberately not in `requirements.txt`.

### `theme.py`

Single place for brand colors (PLN/IPS navy `#0A3D8F` + yellow `#FDB913`), `STATUS_COLORS` / `PAYMENT_COLORS` maps keyed by canonical status strings from `data_loader`, `ORG_NAME`, and `LOGO_PATH` (`data/logo.png`, optional — sidebar and PPT fall back to text badge). `app.py` and `ppt_export.py` both import from here; `.streamlit/config.toml` duplicates primary color for Streamlit's own widgets, must be kept in sync by hand.
