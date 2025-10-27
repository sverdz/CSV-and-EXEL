@echo off
REM Скрипт для вирішення конфлікту в README.md

echo ========================================
echo   Вирішення конфлікту в README.md
echo ========================================
echo.

echo [1] Перевірка статусу...
git status

echo.
echo Оберіть варіант вирішення:
echo.
echo 1 - Взяти версію з сервера (рекомендовано)
echo 2 - Залишити свою локальну версію
echo 3 - Повністю перезавантажити з сервера
echo 4 - Відкрити README для ручного редагування
echo 5 - Скасувати
echo.

set /p choice="Ваш вибір (1-5): "

if "%choice%"=="1" goto SERVER_VERSION
if "%choice%"=="2" goto LOCAL_VERSION
if "%choice%"=="3" goto FULL_RESET
if "%choice%"=="4" goto MANUAL_EDIT
if "%choice%"=="5" goto CANCEL

:SERVER_VERSION
echo.
echo [2] Беремо версію з сервера...
git checkout --theirs README.md
git add README.md
echo [3] Створюємо коміт...
git commit -m "Resolve conflict: accept remote README version"
echo [4] Відправляємо на сервер...
git push
echo.
echo ✓ Готово! Використано версію з сервера.
goto END

:LOCAL_VERSION
echo.
echo [2] Залишаємо локальну версію...
git checkout --ours README.md
git add README.md
echo [3] Створюємо коміт...
git commit -m "Resolve conflict: keep local README version"
echo [4] Відправляємо на сервер...
git push
echo.
echo ✓ Готово! Використано локальну версію.
goto END

:FULL_RESET
echo.
echo ⚠ УВАГА: Це видалить всі локальні незбережені зміни!
set /p confirm="Продовжити? (yes/no): "
if not "%confirm%"=="yes" goto CANCEL

echo.
echo [2] Отримуємо дані з сервера...
git fetch origin
echo [3] Перезавантажуємо репозиторій...
git reset --hard origin/claude/merge-scripts-optimize-011CUWuVpzv5oHPmeHsFCF6e
echo.
echo ✓ Готово! Репозиторій повністю синхронізовано з сервером.
goto END

:MANUAL_EDIT
echo.
echo [2] Відкриваємо README.md...
start notepad README.md
echo.
echo Відредагуйте файл вручну:
echo   - Знайдіть маркери ^<^<^<^<^<^<^< HEAD
echo   - Видаліть маркери та залиште потрібний код
echo   - Збережіть та закрийте файл
echo.
pause
echo.
echo [3] Додаємо зміни...
git add README.md
echo [4] Створюємо коміт...
git commit -m "Resolve README conflict manually"
echo [5] Відправляємо на сервер...
git push
echo.
echo ✓ Готово!
goto END

:CANCEL
echo.
echo Операцію скасовано.
goto END

:END
echo.
echo ========================================
echo   Поточний статус:
echo ========================================
git status
echo.
pause
