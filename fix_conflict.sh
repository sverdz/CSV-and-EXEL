#!/bin/bash
# Скрипт для вирішення конфлікту в README.md

echo "========================================"
echo "  Вирішення конфлікту в README.md"
echo "========================================"
echo ""

echo "[1] Перевірка статусу..."
git status

echo ""
echo "Оберіть варіант вирішення:"
echo ""
echo "1 - Взяти версію з сервера (рекомендовано)"
echo "2 - Залишити свою локальну версію"
echo "3 - Повністю перезавантажити з сервера"
echo "4 - Відкрити README для ручного редагування"
echo "5 - Скасувати"
echo ""

read -p "Ваш вибір (1-5): " choice

case $choice in
    1)
        echo ""
        echo "[2] Беремо версію з сервера..."
        git checkout --theirs README.md
        git add README.md
        echo "[3] Створюємо коміт..."
        git commit -m "Resolve conflict: accept remote README version"
        echo "[4] Відправляємо на сервер..."
        git push
        echo ""
        echo "✓ Готово! Використано версію з сервера."
        ;;
    2)
        echo ""
        echo "[2] Залишаємо локальну версію..."
        git checkout --ours README.md
        git add README.md
        echo "[3] Створюємо коміт..."
        git commit -m "Resolve conflict: keep local README version"
        echo "[4] Відправляємо на сервер..."
        git push
        echo ""
        echo "✓ Готово! Використано локальну версію."
        ;;
    3)
        echo ""
        echo "⚠ УВАГА: Це видалить всі локальні незбережені зміни!"
        read -p "Продовжити? (yes/no): " confirm
        if [ "$confirm" = "yes" ]; then
            echo ""
            echo "[2] Отримуємо дані з сервера..."
            git fetch origin
            echo "[3] Перезавантажуємо репозиторій..."
            git reset --hard origin/claude/merge-scripts-optimize-011CUWuVpzv5oHPmeHsFCF6e
            echo ""
            echo "✓ Готово! Репозиторій повністю синхронізовано з сервером."
        else
            echo "Операцію скасовано."
        fi
        ;;
    4)
        echo ""
        echo "[2] Відкриваємо README.md..."

        # Спроба відкрити в різних редакторах
        if command -v nano &> /dev/null; then
            nano README.md
        elif command -v vim &> /dev/null; then
            vim README.md
        elif command -v vi &> /dev/null; then
            vi README.md
        else
            echo "Редактор не знайдено. Відкрийте файл вручну: README.md"
            read -p "Натисніть Enter після редагування..."
        fi

        echo ""
        echo "[3] Додаємо зміни..."
        git add README.md
        echo "[4] Створюємо коміт..."
        git commit -m "Resolve README conflict manually"
        echo "[5] Відправляємо на сервер..."
        git push
        echo ""
        echo "✓ Готово!"
        ;;
    5)
        echo ""
        echo "Операцію скасовано."
        ;;
    *)
        echo ""
        echo "Невірний вибір!"
        ;;
esac

echo ""
echo "========================================"
echo "  Поточний статус:"
echo "========================================"
git status
echo ""
