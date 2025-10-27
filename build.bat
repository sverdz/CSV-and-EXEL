@echo off
REM Build script for Data Processor GUI
REM Створення exe файлу з GUI інтерфейсом

echo ========================================
echo   Data Processor - Build Script
echo ========================================
echo.

REM Перевірка наявності PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [ERROR] PyInstaller не встановлено!
    echo Встановіть: pip install pyinstaller
    pause
    exit /b 1
)

echo [1/5] Очищення попередніх збірок...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

echo [2/5] Перевірка залежностей...
pip install -r requirements.txt

echo [3/5] Створення spec файлу...
pyi-makespec ^
    --name="DataProcessor" ^
    --onefile ^
    --windowed ^
    --icon=NONE ^
    --add-data="data_processor.py;." ^
    data_processor_gui.py

echo [4/5] Збірка EXE файлу...
pyinstaller --clean --noconfirm DataProcessor.spec

echo [5/5] Перевірка результату...
if exist "dist\DataProcessor.exe" (
    echo.
    echo ========================================
    echo   УСПІХ!
    echo ========================================
    echo.
    echo EXE файл створено: dist\DataProcessor.exe
    echo Розмір:
    dir dist\DataProcessor.exe | find "DataProcessor.exe"
    echo.
    echo Можете запустити: dist\DataProcessor.exe
    echo.
) else (
    echo.
    echo [ERROR] Не вдалося створити EXE файл!
    echo Перевірте помилки вище.
    echo.
)

echo Натисніть будь-яку клавішу для виходу...
pause >nul
