# Dashboard Monitoring Proyek — PLN Indonesia Power Services

Dashboard Streamlit (UI Bahasa Indonesia) untuk memantau pekerjaan procurement/kontrak. Sumber data satu-satunya: `data/Data.xlsx` — aplikasi membaca sekaligus menulis balik perubahan ke file ini. Single-user, jalan lokal, tanpa server/backend terpisah.

## Fitur

- **Ringkasan** — KPI pencapaian & akumulatif tahun berjalan
- **Monitoring Pekerjaan** — daftar & filter pekerjaan (pencarian, user, proses, eksekutor, status, status bayar, vendor)
- **Akumulatif** — grafik akumulatif KPI
- **Potensi Pendapatan**
- **Alur Penagihan**
- **Kelola Data** — edit langsung dari dashboard (Monitoring, Tambah Pekerjaan, Akumulatif, Potensi, Target, Cadangan/restore)
- **Export PPT** — export slide PowerPoint (chart & tabel native, bisa diedit lagi di PowerPoint), mengikuti filter yang aktif

## Instalasi

Laptop belum ada Python/tools sama sekali? Ikuti **[INSTALL.md](INSTALL.md)** — tutorial step by step dari nol.

Sudah ada Python 3.12 + venv terpasang, cukup:

```powershell
run.bat
```

atau manual:

```powershell
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe -m streamlit run app.py --browser.gatherUsageStats false
```

Buka `http://localhost:8501` di browser.

## Struktur project

```
data/Data.xlsx        sumber data satu-satunya (dibaca & ditulis)
data_loader.py         baca Excel -> dict data
excel_writer.py        tulis balik perubahan ke Excel
app.py                  UI Streamlit, 7 tab
ppt_export.py           export ke .pptx
theme.py                warna brand, konstanta tampilan
run.bat                 setup venv + install paket + jalankan app
```

## Dokumentasi lain

- **[INSTALL.md](INSTALL.md)** — tutorial instalasi laptop baru
- **[CLAUDE.md](CLAUDE.md)** / **[AGENT.md](AGENT.md)** — panduan arsitektur & aturan kerja untuk AI coding agent (perilaku loader/writer, hal-hal yang gampang rusak kalau diubah sembarangan)

## Catatan penting

- File `data/Data.xlsx` adalah satu-satunya sumber kebenaran — jangan edit struktur sheet/kolom tanpa menyesuaikan `data_loader.py` & `excel_writer.py`.
- Kalau muncul file `~$Data.xlsx` di folder `data`, artinya Excel sedang membuka file itu di tempat lain — tutup Excel dulu sebelum menyimpan dari dashboard.
- Tidak ada git repo / test suite bawaan di project ini.
