Правила с clasp (синхронизация с гугл скриптом):

✅ Пошагово:
🔁 Шаг 1: Активируй нужный аккаунт вручную
Если хочешь поработать, например, с аккаунтом Sloboda:
bash
Копировать
Редактировать
cp ~/.clasprc_sloboda.json ~/.clasprc.json
Или с аккаунтом Personal:
bash
Копировать
Редактировать
cp ~/.clasprc_personal.json ~/.clasprc.json
📌 Это переключает текущий активный аккаунт, который использует clasp.

📁 Шаг 2: Перейди в нужный проект
bash
Копировать
Редактировать
cd ~/путь*к*папке_проекта
Пример:

bash
Копировать
Редактировать
cd ~/projects/sloboda-webapp/
🚀 Шаг 3: Работай с проектом как обычно
bash
Копировать
Редактировать
clasp pull # подтянуть код с Google
clasp push # отправить код на сервер
clasp open # открыть проект в браузере
🧠 Запомни:
Перед каждым переходом между проектами важно выполнить cp ~/.clasprc\_\*.json ~/.clasprc.json — иначе clasp будет работать с "не тем" аккаунтом.

Сами проекты не конфликтуют, если ты правильно подменяешь .clasprc.json.

❓ Пример: хочу поработать с аккаунтом Sloboda
bash
Копировать
Редактировать
cp ~/.clasprc_sloboda.json ~/.clasprc.json
cd ~/projects/sloboda-webapp/
clasp pull
code .

Влить бранч на гитхабе:

Если смерджила бранч в гитхабе и сюда нужно подтянуть:
git pull origin main

Файл с правами доступов к приложению (имейлы): https://docs.google.com/spreadsheets/d/1xBB9PdNKdu8hBCAAe-iEW5cTW4_oQWOZjDzwFn8bH_U/edit?gid=0#gid=0