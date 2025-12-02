@echo off
echo ========================================
echo  CRIAÇÃO DO EXECUTÁVEL SORTEADOR INRAD
echo ========================================
echo.

echo 1. Instalando dependências...
pip install -r requirements.txt

echo.
echo 2. Criando executável com PyInstaller...
echo.

pyinstaller --onefile ^
            --windowed ^
            --name "Sorteador_INRAD" ^
            --icon=icon.ico ^
            --add-data "audio/*.mp3;audio" ^
            --hidden-import=pandas ^
            --hidden-import=openpyxl ^
            --hidden-import=pygame ^
            Sorteador-INRAD-2025.py

echo.
echo ========================================
echo  EXECUTÁVEL CRIADO COM SUCESSO!
echo  O arquivo estará em: dist\Sorteador_INRAD.exe
echo ========================================

pause