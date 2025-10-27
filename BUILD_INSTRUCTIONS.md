# Інструкція зі збірки EXE файлу

## Швидкий старт (Windows)

1. Відкрийте командний рядок або PowerShell
2. Перейдіть в папку проекту:
   ```bash
   cd CSV-and-EXEL
   ```
3. Запустіть:
   ```bash
   build.bat
   ```
4. Готовий EXE буде в папці `dist/DataProcessor.exe`

## Покрокова інструкція

### Крок 1: Встановіть Python

Якщо Python ще не встановлено:
- Завантажте з https://www.python.org/downloads/
- Версія: Python 3.8 або новіша
- Під час встановлення обов'язково відмітьте "Add Python to PATH"

### Крок 2: Встановіть залежності

```bash
pip install -r requirements.txt
```

Це встановить:
- pandas
- openpyxl
- xlsxwriter
- tqdm
- pyinstaller

### Крок 3: Збірка EXE

#### Варіант А: Автоматичний (рекомендовано)

```bash
build.bat
```

#### Варіант Б: Ручний

```bash
# Очистити попередні збірки
rmdir /s /q build dist
del /q *.spec

# Створити EXE
pyinstaller --clean --noconfirm DataProcessor.spec
```

### Крок 4: Тестування

```bash
dist\DataProcessor.exe
```

Повинно відкритися вікно програми з GUI інтерфейсом.

## Налаштування збірки

### Зміна імені файлу

Відредагуйте `DataProcessor.spec`:
```python
name='DataProcessor',  # Змініть на потрібне ім'я
```

### Додання іконки

1. Створіть або завантажте .ico файл (наприклад, `app.ico`)
2. Покладіть його в папку проекту
3. Відредагуйте `DataProcessor.spec`:
   ```python
   icon='app.ico',
   ```

### Зменшення розміру EXE

#### Опція 1: UPX компресія (вже включена)
```python
upx=True,
```

#### Опція 2: Виключення непотрібних модулів
У `DataProcessor.spec` додайте до `excludes`:
```python
excludes=[
    'matplotlib',
    'scipy',
    'IPython',
    'notebook',
    'pytest',
    'PIL',  # Якщо не працюєте з зображеннями
],
```

### Консольна версія (з вікном консолі)

Змініть в `DataProcessor.spec`:
```python
console=True,  # Було False
```

## Linux / Mac

### Встановлення

```bash
pip install -r requirements.txt
```

### Збірка

```bash
chmod +x build.sh
./build.sh
```

Готовий binary буде в `dist/DataProcessor`

### Запуск

```bash
./dist/DataProcessor
```

## Типові проблеми

### Помилка: "PyInstaller не знайдено"

```bash
pip install pyinstaller
```

### Помилка: "Модуль pandas не знайдено"

```bash
pip install -r requirements.txt
```

### EXE не запускається

1. Перевірте антивірус (може блокувати)
2. Запустіть з консолі для перегляду помилок:
   ```bash
   dist\DataProcessor.exe
   ```
3. Перебудуйте з консоллю для діагностики:
   ```python
   console=True  # в DataProcessor.spec
   ```

### EXE дуже повільно запускається

Це нормально для першого запуску. PyInstaller розпаковує файли при старті.

### Антивірус блокує EXE

Додайте в виключення або підпишіть EXE цифровим підписом.

## Додаткові опції PyInstaller

### Збірка в один файл (замість папки)

Вже налаштовано в spec файлі:
```python
exe = EXE(
    ...
    [],  # Порожній список = один файл
)
```

### Збірка в папку (з DLL окремо)

Змініть на:
```python
exe = EXE(
    pyz,
    a.scripts,
    [],  # Порожній
    exclude_binaries=True,  # Додати
    ...
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='DataProcessor'
)
```

## Розповсюдження

### Що включити в реліз:

1. `dist/DataProcessor.exe` - головний файл
2. `README.md` - інструкція користувача
3. Приклади файлів (опціонально)

### Що НЕ треба включати:

- Папку `build/`
- `.spec` файли
- Вихідний код Python (якщо не потрібен)
- `__pycache__/`

## Версіонування

Додайте версію в код (`data_processor_gui.py`):

```python
VERSION = "1.0.0"
self.root.title(f"Обробник CSV та Excel файлів v{VERSION}")
```

## Автоматизація релізів

### GitHub Actions приклад

Створіть `.github/workflows/build.yml`:

```yaml
name: Build EXE

on:
  push:
    tags:
      - 'v*'

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
        with:
          python-version: '3.10'
      - run: pip install -r requirements.txt
      - run: pyinstaller DataProcessor.spec
      - uses: actions/upload-artifact@v2
        with:
          name: DataProcessor
          path: dist/DataProcessor.exe
```

## Підтримка

При виникненні проблем:
1. Перевірте Python версію: `python --version`
2. Перевірте PyInstaller: `pyinstaller --version`
3. Перегляньте логи збірки
4. Створіть issue в GitHub репозиторії

---

**Успішної збірки!** 🚀
