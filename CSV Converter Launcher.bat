@echo off
title CSV to Excel Converter - Launcher
color 0A
echo.
echo =========================================
echo   CSV to Excel Converter - Launcher
echo =========================================
echo.

REM Python kontrolü
echo [1/4] Python kontrol ediliyor...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ HATA: Python bulunamadi!
    echo.
    echo Python'u indirip yukleyin: https://python.org
    echo.
    pause
    exit /b 1
) else (
    echo ✅ Python bulundu
)

REM Gerekli kütüphaneleri kontrol et ve yükle
echo.
echo [2/4] Gerekli kutuphaneler kontrol ediliyor...

REM pandas kontrolü
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  pandas bulunamadi, yukleniyor...
    pip install pandas
    if errorlevel 1 (
        echo ❌ pandas yuklenemedi!
        goto :error
    )
) else (
    echo ✅ pandas mevcut
)

REM openpyxl kontrolü  
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  openpyxl bulunamadi, yukleniyor...
    pip install openpyxl
    if errorlevel 1 (
        echo ❌ openpyxl yuklenemedi!
        goto :error
    )
) else (
    echo ✅ openpyxl mevcut
)

REM tkinter kontrolü (genellikle Python ile gelir)
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo ❌ tkinter bulunamadi! Python kurulumunuzu kontrol edin.
    goto :error
) else (
    echo ✅ tkinter mevcut
)

echo.
echo [3/4] Tum kutuphaneler hazir!
echo.
echo [4/4] CSV Converter baslatiliyor...
echo.

REM Python dosyasını çalıştır
if exist "csv_converter.py" (
    python csv_converter.py
) else if exist "csv_to_xlsx_converter.py" (
    python csv_to_xlsx_converter.py
) else (
    echo ❌ HATA: Python dosyasi bulunamadi!
    echo.
    echo Asagidaki dosyalardan birini bu klasore koyun:
    echo - csv_converter.py
    echo - csv_to_xlsx_converter.py
    echo.
    goto :error
)

REM Program bittikten sonra
echo.
echo Program kapatildi.
goto :end

:error
echo.
echo ❌ Bir hata olustu! Lutfen kontrol edin.
echo.
pause
exit /b 1

:end
echo.
echo Program basariyla tamamlandi.
timeout /t 3 >nul
exit /b 0