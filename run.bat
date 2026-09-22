@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
  echo Membuat virtual environment...
  python -m venv venv
)
venv\Scripts\python.exe -m pip --version >nul 2>&1 || venv\Scripts\python.exe -m ensurepip
venv\Scripts\python.exe -c "import streamlit, plotly, pptx, openpyxl, pandas" 2>nul || venv\Scripts\python.exe -m pip install -r requirements.txt
echo Menjalankan dashboard di http://localhost:8501 (browser akan terbuka otomatis) ...
echo Tekan Ctrl+C di jendela ini untuk menghentikan dashboard.
venv\Scripts\python.exe -m streamlit run app.py --browser.gatherUsageStats false
pause
