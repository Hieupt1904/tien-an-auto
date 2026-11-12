@echo off
rem Tao file TienAn.exe tu source (chay 1 lan tren may co Python + mang)
rem Ket qua: dist\TienAn\TienAn.exe  (chep ca thu muc dist\TienAn di la chay duoc)
chcp 65001 >nul
cd /d "%~dp0.."
python -m pip install --upgrade pyinstaller || goto :err
python -m PyInstaller --noconfirm --clean ^
  --name TienAn ^
  --add-data "src\template.json;." ^
  --paths src --paths src\app ^
  --hidden-import _api --hidden-import _pull --hidden-import _report ^
  --hidden-import _xlsio --hidden-import _xls --hidden-import _ole ^
  --hidden-import _http --hidden-import _ocr --hidden-import _session --hidden-import _log ^
  src\app\app_run.py || goto :err
echo.
echo Xong. Thu muc chay duoc: dist\TienAn
goto :eof
:err
echo Co loi khi dong goi. Xem thong bao phia tren.
pause
