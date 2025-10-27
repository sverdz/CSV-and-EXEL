# Як вирішити конфлікт у README.md

## Автоматичне вирішення (рекомендовано)

### Варіант 1: Взяти версію з сервера

```bash
# Скасувати локальні зміни та взяти версію з GitHub
git checkout --theirs README.md
git add README.md
git commit -m "Resolve conflict: accept remote version"
git push
```

### Варіант 2: Залишити свою версію

```bash
# Залишити свою локальну версію
git checkout --ours README.md
git add README.md
git commit -m "Resolve conflict: keep local version"
git push
```

### Варіант 3: Взяти найновішу версію з гілки

```bash
# Повністю перезавантажити з сервера
git fetch origin
git reset --hard origin/claude/merge-scripts-optimize-011CUWuVpzv5oHPmeHsFCF6e
```

⚠️ **УВАГА**: Варіант 3 видалить всі ваші локальні незбережені зміни!

## Ручне вирішення

Якщо хочете об'єднати зміни вручну:

### 1. Відкрийте README.md в редакторі

Знайдіть маркери конфлікту:

```
<<<<<<< HEAD
Ваша версія
=======
Версія з сервера
>>>>>>> branch-name
```

### 2. Виберіть потрібні зміни

Видаліть маркери (`<<<<<<<`, `=======`, `>>>>>>>`) та залиште потрібний код:

```markdown
# Правильна об'єднана версія
```

### 3. Збережіть та закомітьте

```bash
git add README.md
git commit -m "Resolve merge conflict in README.md"
git push
```

## Якщо нічого не допомагає

### Повне перестворення з сервера:

```bash
# 1. Створіть резервну копію своїх змін (якщо потрібно)
cp README.md README.md.backup

# 2. Видаліть локальний файл
rm README.md

# 3. Відновіть з сервера
git checkout origin/claude/merge-scripts-optimize-011CUWuVpzv5oHPmeHsFCF6e -- README.md

# 4. Додайте та закомітьте
git add README.md
git commit -m "Restore README from remote"
git push
```

## Перевірка стану

Після вирішення конфлікту перевірте:

```bash
# Статус git
git status

# Має бути: "working tree clean"

# Перегляд останніх комітів
git log --oneline -5

# Перевірка що все синхронізовано
git pull
```

## Запобігання конфліктам у майбутньому

1. **Завжди робіть pull перед push:**
   ```bash
   git pull
   git add .
   git commit -m "message"
   git push
   ```

2. **Працюйте в окремих гілках:**
   ```bash
   git checkout -b my-feature
   # робіть зміни
   git push -u origin my-feature
   ```

3. **Використовуйте rebase замість merge:**
   ```bash
   git pull --rebase
   ```

## Допомога

Якщо жодне з рішень не працює, надішліть вивід команди:

```bash
git status > git_status.txt
```

І прикріпіть файл `git_status.txt` до issue.

---

**Швидке рішення (якщо нема важливих локальних змін):**

```bash
git fetch origin
git reset --hard origin/claude/merge-scripts-optimize-011CUWuVpzv5oHPmeHsFCF6e
```

Це поверне ваш репозиторій до стану сервера.
