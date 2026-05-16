@echo off
:: ============================================================
:: Build script — Word Tools para Windows (.exe)
:: Requisitos: Python 3.10+, pip install pyinstaller pywin32 PyQt6 python-docx
:: ============================================================

echo.
echo  Construindo Word Tools para Windows...
echo.

:: Limpa builds anteriores
if exist dist\WordTools rmdir /s /q dist\WordTools
if exist build rmdir /s /q build

:: Gera o executavel
pyinstaller ^
    --noconfirm ^
    --windowed ^
    --onedir ^
    --name "WordTools" ^
    --icon "assets\icon.ico" ^
    --add-data "app;app" ^
    main.py

echo.
if exist dist\WordTools\WordTools.exe (
    echo  [OK] Build concluido: dist\WordTools\WordTools.exe
) else (
    echo  [ERRO] Build falhou. Veja o log acima.
    exit /b 1
)

echo.
echo  Para distribuir: compacte a pasta dist\WordTools\ em um .zip
echo  ou use o Inno Setup para gerar um instalador .exe
echo.
pause
