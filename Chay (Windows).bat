@echo off
chcp 65001 >nul
title Tien an - tu dong cap nhat
cd /d "%~dp0"
where py >nul 2>nul && (py -3 "src\app\app_run.py" & goto :end)
where python >nul 2>nul && (python "src\app\app_run.py" & goto :end)
echo Khong tim thay Python. Hay cai Python 3 tu python.org (nho tich "Add to PATH").
pause
:end
