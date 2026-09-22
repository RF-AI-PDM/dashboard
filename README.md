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

## Deploy ke Streamlit Community Cloud

Aplikasi ini butuh proses Python yang tetap hidup (WebSocket) dan bisa menulis file lokal (`excel_writer.py` menyimpan balik ke `data/Data.xlsx`), jadi **tidak bisa** dideploy sebagai fungsi serverless (Vercel, Netlify, dsb). Platform yang cocok: [Streamlit Community Cloud](https://share.streamlit.io) (gratis, resmi dari Streamlit).

1. Push repo ini ke GitHub (sudah: `RF-AI-PDM/dashboard`, branch `master`).
2. Buka [share.streamlit.io](https://share.streamlit.io), sign in pakai akun GitHub yang sama.
3. **New app** → pilih repo `RF-AI-PDM/dashboard`, branch `master`, main file path `app.py`.
4. Klik **Deploy**. Streamlit Cloud akan `pip install -r requirements.txt` lalu `streamlit run app.py` di container mereka sendiri (bukan `venv` lokal kamu — `venv/` sengaja tidak ikut di-commit, lihat `.gitignore`).
5. URL publik siap dalam 1-2 menit (format `https://<nama-app>.streamlit.app`).

**Batasan penting yang perlu dipahami** (bukan bug, memang cara kerja platform gratis ini):

- **Penyimpanan bersifat sementara (ephemeral).** Edit lewat tab "Kelola Data" tetap tersimpan ke `data/Data.xlsx` *selama container itu hidup*, tapi hilang lagi ke versi terakhir yang ada di GitHub begitu container di-restart — ini terjadi otomatis saat: push commit baru ke repo, app di-reboot manual dari dashboard Streamlit Cloud, atau setelah container "tidur" karena tidak ada yang akses lama (free tier). Kalau perubahan data perlu permanen, commit & push `data/Data.xlsx` yang sudah diupdate kembali ke GitHub.
- **App tidur setelah idle lama** (free tier) — akses pertama setelah tidur perlu ~30-60 detik untuk bangun (cold start).
- Backup otomatis (`data/backup/`) tetap dibuat di container saat ada yang klik Simpan, tapi juga ikut hilang saat container restart (memang sengaja masuk `.gitignore`, tidak boleh menumpuk di git).
- Kalau butuh data yang benar-benar persisten multi-user (bukan cuma "gudang laporan"), platform ini bukan jawaban jangka panjang — perlu pindah ke database (lihat catatan di `CLAUDE.md`).

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
- Tidak ada test suite otomatis (lihat perintah verifikasi manual di `CLAUDE.md`).
