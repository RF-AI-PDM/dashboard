# Tutorial Instalasi di Laptop Baru

Panduan ini untuk memasang dan menjalankan dashboard di laptop Windows yang **belum ada software apapun** (belum ada Python, belum ada apa-apa).

## Yang dibutuhkan

- Laptop Windows 10/11
- Koneksi internet (untuk download Python dan paket-paket sekali di awal)
- Folder project ini (`DASHBOARD`) beserta file `data/Data.xlsx` di dalamnya

---

## Langkah 1 — Install Python

1. Buka browser, kunjungi https://www.python.org/downloads/
2. Klik tombol download versi terbaru (disarankan Python 3.12.x, sesuai yang dipakai project ini).
3. Jalankan installer yang sudah didownload.
4. **PENTING:** di layar pertama installer, centang kotak **"Add python.exe to PATH"** di bagian bawah, baru klik **Install Now**.

   Kalau langkah ini terlewat, nanti perintah `python` tidak akan dikenali di Command Prompt / PowerShell.

5. Setelah selesai, cek instalasi berhasil. Buka **PowerShell** (klik Start, ketik "PowerShell", Enter), lalu ketik:

   ```powershell
   python --version
   ```

   Harus muncul versi Python (misal `Python 3.12.x`). Kalau muncul error "not recognized", restart laptop dulu (supaya PATH ter-refresh), lalu coba lagi.

---

## Langkah 2 — Salin folder project ke laptop baru

1. Salin seluruh folder `DASHBOARD` (via flashdisk, share network, cloud drive, dll) ke lokasi mana saja di laptop baru, misalnya `D:\Artefact\DASHBOARD`.
2. Pastikan folder `data` ikut tersalin, dan di dalamnya ada file `Data.xlsx`. Tanpa file ini aplikasi tidak bisa jalan.
3. **Jangan** salin folder `venv` lama (kalau ada) dari laptop sebelumnya — venv harus dibuat baru di laptop ini karena environment Python beda mesin bisa rusak. Kalau folder `venv` ikut tersalin, hapus saja folder tersebut sebelum lanjut ke langkah berikut.

---

## Langkah 3 — Jalankan aplikasi

1. Masuk ke folder project di File Explorer.
2. Double-click **`run.bat`**.
3. Jendela hitam (Command Prompt) akan terbuka dan otomatis:
   - membuat virtual environment (folder `venv`) — proses ini sekali saja, agak lama (bisa 1-3 menit)
   - menginstall semua paket yang dibutuhkan (`streamlit`, `pandas`, `openpyxl`, `plotly`, `python-pptx`) — juga sekali saja, tergantung kecepatan internet, biasanya 2-5 menit
   - menjalankan dashboard dan membuka browser otomatis ke `http://localhost:8501`
4. Kalau Windows Defender / firewall menampilkan pop-up minta izin akses jaringan, klik **Allow access**.
5. Kalau berhasil, dashboard akan tampil di browser. Biarkan jendela Command Prompt tetap terbuka selama memakai dashboard — menutup jendela itu akan mematikan aplikasinya.

Untuk menjalankan lagi di lain waktu, cukup double-click `run.bat` lagi — tidak akan install ulang karena `venv` dan paket sudah ada, jadi langsung cepat jalan.

---

## Menghentikan aplikasi

Klik di jendela Command Prompt yang terbuka, tekan `Ctrl+C`, lalu tutup jendelanya. Atau tutup langsung jendelanya.

---

## Troubleshooting

**"python is not recognized as an internal or external command"**
Python belum ke-install dengan benar atau lupa centang "Add to PATH" saat instalasi. Install ulang Python dan pastikan centang opsi itu, lalu restart laptop.

**`run.bat` langsung tertutup / error saat install paket**
Biasanya karena koneksi internet bermasalah saat `pip install`. Buka PowerShell, masuk ke folder project, jalankan manual supaya errornya kebaca:

```powershell
cd "D:\Artefact\DASHBOARD"
venv\Scripts\python.exe -m pip install -r requirements.txt
```

**File `Data.xlsx` tidak ketemu / error saat load data**
Pastikan file `data/Data.xlsx` benar-benar ada di folder `data` (nama file dan lokasi harus persis sama).

**Ada file `~$Data.xlsx` di folder `data`**
Artinya `Data.xlsx` sedang terbuka di Microsoft Excel di komputer lain/sesi lain. Tutup Excel-nya dulu sebelum menyimpan perubahan dari dashboard. Jangan hapus file `~$` ini manual.

**Browser tidak terbuka otomatis**
Buka browser manual, ketik alamat: `http://localhost:8501`

**Mau install ulang dari nol (venv rusak)**
Hapus folder `venv` di dalam project, lalu double-click `run.bat` lagi — akan dibuat ulang otomatis.
