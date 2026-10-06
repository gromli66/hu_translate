#!/usr/bin/env bash
# Приёмка на сервере после «docker compose up -d» (запускать из папки docker/):  bash deploy_check.sh
# Ничего не меняет, только проверяет. Итог — число OK / FAIL.
cd "$(dirname "$0")"
ok=0; fail=0
check() {   # check "что проверяем" команда…
  local name="$1"; shift
  if out=$("$@" 2>&1); then echo "OK    $name"; ok=$((ok + 1)); else echo "FAIL  $name: ${out:0:200}"; fail=$((fail + 1)); fi
}
port="$(grep -E '^HUT_PORT=' .env 2>/dev/null | cut -d= -f2)"; port="${port:-8010}"

check "контейнер hut запущен"            sh -c 'docker ps --filter name=^hut$ --filter status=running -q | grep -q .'
check "ответ /health"                    curl -fsS "http://127.0.0.1:${port}/health"
check "страница входа открывается"       sh -c "curl -fsS http://127.0.0.1:${port}/login | grep -q 'Вход'"
check "ключ шифрования задан в .env"     sh -c "grep -qE '^HUT_SECRET_KEY=.+' .env"
check "Tesseract с венгерским"           sh -c "docker exec hut tesseract --list-langs | grep -qx hun"
check "из контейнера виден портал"       docker exec hut python -c "import urllib.request,urllib.error,os
try: urllib.request.urlopen(os.environ.get('ROSATOM_AI_BASE','https://go.ai-rosatom.ru')+'/api/models', timeout=15)
except urllib.error.HTTPError as e: assert e.code in (401, 403), e   # без токена портал отвечает отказом — значит, доступен"
check "есть хотя бы один проект"         sh -c "ls ../projects/*/project.json >/dev/null"
check "есть админ"                       sh -c "docker exec hut python manage.py user list | grep -q admin"
check "на диске больше 5 ГБ свободно"    sh -c "[ \$(df -Pk .. | awk 'NR==2 {print \$4}') -gt 5000000 ]"
check "ротация логов контейнера"         sh -c "docker inspect hut --format '{{.HostConfig.LogConfig.Config}}' | grep -q max-size"

echo "итог: OK $ok, FAIL $fail"
[ "$fail" -eq 0 ]
