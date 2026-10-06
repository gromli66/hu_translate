# Переводчик HU→RU

Внутренний веб-сервис: переводит венгерскую эксплуатационную документацию АЭС (DOCX, PDF, сканы) на русский
через портал go.ai-rosatom.ru и возвращает перевод в исходном формате. Версия сырая — для старта и проверки.

## Развернуть

Нужен Docker с Compose. Серверу нужен доступ к GitHub, Docker Hub, PyPI (для сборки) и к go.ai-rosatom.ru.

**1. Код и запуск** — сервис поднимется на порту 8010 (другой порт: до запуска `echo HUT_PORT=9000 > .env`):

```bash
git clone https://github.com/gromli66/hu_translate.git && cd hu_translate
```

```bash
docker compose up -d --build
```

**2. Глоссарии.** Комплект «Пакш» уже в репозитории — [projects/paks/](projects/paks/), делать ничего не нужно.
Это данные заказчика: репозиторий приватный, наружу не выкладывать.

Новый комплект со своим глоссарием (docx, xlsx или md с таблицей «венгерский | русский | комментарий»):

```bash
cp глоссарий.docx projects/ && docker compose exec hut python manage.py project add <имя> projects/глоссарий.docx
```

**3. Пользователи** — пароль команда спросит. Роли: `user` (по умолчанию), `expert`, `admin`.

```bash
docker compose exec hut python manage.py user add admin --role admin
```

```bash
docker compose exec hut python manage.py user add ivanov
```

Пользователю передать адрес `http://<сервер>:8010`, логин и пароль.

Обновить: `git pull && docker compose up -d --build`. Всё состояние — в папке `data/` (сервис сам раз в сутки
кладёт копию базы в `data/backups/`). Если что-то не так: `docker compose logs --tail 50`.

## Пользоваться

1. Открыть адрес, войти.
2. «Профиль» → вставить свой токен портала go.ai-rosatom.ru (где взять — в «Справке»).
3. «Переводы» → выбрать проект, загрузить файлы → дождаться → «Вычитка» (по желанию) → «Скачать».

Подробно — меню «Справка» в сервисе. Эксперту (пополнение глоссария) — «Эксперт» → «Справка».

## Что где

```
server/      веб-сервис (FastAPI + HTMX)
src/         движок перевода
projects/    глоссарии проектов (paks — «Пакш»)
data/        база и переводы (создаётся при запуске)
docs/        DEV.md — устройство, качество, ограничения; PLAN_service.md
tests/       pytest
manage.py    пользователи и проекты
Dockerfile, docker-compose.yml
```
