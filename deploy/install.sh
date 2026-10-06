#!/usr/bin/env bash
# Установка и обновление переводчика HU→RU. Запускать из папки, куда распакован архив:  ./install.sh
# Повторный запуск (после распаковки нового архива поверх старого) обновляет сервис: данные, пользователи и
# настройки сохраняются.
# Без вопросов (для автоматизации):  HUT_PORT=8010 ADMIN_LOGIN=admin ADMIN_PASSWORD='…' ./install.sh
set -euo pipefail
cd "$(dirname "$0")"
say() { printf '\n== %s\n' "$*"; }

command -v docker >/dev/null || { echo "Нужен Docker: https://docs.docker.com/engine/install/"; exit 1; }
docker compose version >/dev/null 2>&1 || { echo "Нужен Docker Compose v2 (команда «docker compose»)."; exit 1; }

say "Загружаю образ"
docker load < hut-image.tar.gz | tail -1

if [ ! -f .env ]; then
  port="${HUT_PORT:-}"
  [ -n "$port" ] || read -rp "Порт, на котором пользователи будут открывать сервис [8010]: " port
  port="${port:-8010}"
  key=$(docker run --rm hut:latest python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
  cat > .env <<EOF
# Настройки переводчика, создано install.sh $(date +%F).
# Ключ не терять и не менять: без него сохранённые токены пользователей не прочитать.
HUT_SECRET_KEY=$key
HUT_PORT=$port
HUT_MAX_JOBS=3
HUT_MAX_UPLOAD_MB=200
HUT_BACKUP_KEEP=14
HUT_KEEP_DAYS=0
HUT_CACHE_DAYS=0
HUT_COOKIE_SECURE=0
ROSATOM_AI_BASE=https://go.ai-rosatom.ru
ROSATOM_MODEL=privateLLM
EOF
  chmod 600 .env
fi
port=$(grep -E '^HUT_PORT=' .env | cut -d= -f2)
mkdir -p data projects

say "Запускаю сервис"
docker compose up -d
for _ in $(seq 1 60); do
  curl -fsS -m 3 "http://127.0.0.1:${port}/health" >/dev/null 2>&1 && break
  sleep 2
done

if ! docker compose exec -T hut python manage.py user list | grep -q " admin "; then
  say "Администратор сервиса"
  login="${ADMIN_LOGIN:-}"
  [ -n "$login" ] || read -rp "Логин администратора: " login
  if [ -n "${ADMIN_PASSWORD:-}" ]; then
    docker compose exec -T hut python manage.py user add "$login" --role admin --password "$ADMIN_PASSWORD"
  else
    docker compose exec hut python manage.py user add "$login" --role admin
  fi
fi

say "Проверка"
./hut check || true
addr=$(hostname -I 2>/dev/null | awk '{print $1}') || true
say "Готово. Адрес для пользователей: http://${addr:-<адрес сервера>}:${port}"
echo "Новый пользователь: ./hut user add <логин>   (подробнее — ADMIN.md)"
