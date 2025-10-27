#!/bin/bash
# Build script for Data Processor GUI (Linux/Mac)
# Створення exe/binary файлу з GUI інтерфейсом

echo "========================================"
echo "  Data Processor - Build Script"
echo "========================================"
echo ""

# Перевірка наявності PyInstaller
if ! python3 -c "import PyInstaller" 2>/dev/null; then
    echo "[ERROR] PyInstaller не встановлено!"
    echo "Встановіть: pip install pyinstaller"
    exit 1
fi

echo "[1/5] Очищення попередніх збірок..."
rm -rf build dist *.spec

echo "[2/5] Перевірка залежностей..."
pip install -r requirements.txt

echo "[3/5] Створення spec файлу..."
pyi-makespec \
    --name="DataProcessor" \
    --onefile \
    --windowed \
    --add-data="data_processor.py:." \
    data_processor_gui.py

echo "[4/5] Збірка binary файлу..."
pyinstaller --clean --noconfirm DataProcessor.spec

echo "[5/5] Перевірка результату..."
if [ -f "dist/DataProcessor" ]; then
    echo ""
    echo "========================================"
    echo "  УСПІХ!"
    echo "========================================"
    echo ""
    echo "Binary файл створено: dist/DataProcessor"
    ls -lh dist/DataProcessor
    echo ""
    echo "Можете запустити: ./dist/DataProcessor"
    echo ""
else
    echo ""
    echo "[ERROR] Не вдалося створити binary файл!"
    echo "Перевірте помилки вище."
    echo ""
fi
