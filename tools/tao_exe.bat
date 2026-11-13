@echo off
rem ===================================================================
rem  Dong goi TienAn.exe (MOT file, khong can Python tren may dich)
rem  Chay file nay MOT LAN tren may Windows co Python 3 + mang.
rem  Ket qua: dist\TienAn.exe  -> chep di dau cung chay duoc.
rem
rem  LUU Y: .exe phai build TREN WINDOWS. Build tren Linux/macOS se ra
rem  binary cua he do, khong chay duoc tren Windows.
rem ===================================================================
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

echo [1/3] Kiem tra Python...
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% --version || goto :nopython

echo [2/3] Cai PyInstaller...
%PY% -m pip install --upgrade --quiet pyinstaller || goto :err

echo [3/3] Dong goi...
%PY% -m PyInstaller --noconfirm --clean --onefile ^
  --name TienAn ^
  --add-data "src\template.json;." ^
  --paths src --paths src\app ^
  --hidden-import _api --hidden-import _pull --hidden-import _report ^
  --hidden-import _xlsio --hidden-import _xls --hidden-import _ole ^
  --hidden-import _http --hidden-import _ocr --hidden-import _session ^
  --hidden-import _log ^
  --exclude-module tkinter.test --exclude-module test ^
  src\app\app_run.py || goto :err

if not exist "dist\TienAn.exe" goto :err

echo.
echo ====================================================
echo  XONG: dist\TienAn.exe
echo.
echo  Cach dung tren may truong (khong can cai gi):
echo    1. Chep TienAn.exe ra mot thu muc rieng
echo       (vd Desktop\Tien an)
echo    2. Bam dup TienAn.exe
echo    3. Lan dau nhap tai khoan + mat khau roi bam Luu
echo    4. Moi ngay: mo len, bam "CAP NHAT DU LIEU"
echo.
echo  File tao ra nam trong thu muc "Bao cao" canh TienAn.exe
echo ====================================================
goto :eof

:nopython
echo.
echo KHONG TIM THAY PYTHON 3.
echo Tai tai: https://www.python.org/downloads/windows/
echo Khi cai nho tich "Add python.exe to PATH".
pause
goto :eof

:err
echo.
echo Co loi khi dong goi. Xem thong bao phia tren.
pause
