#!/usr/bin/env bash
# Сборка образа и архив для сервера (там нет интернета для apt/pip).
# Запускать на машине с интернетом из корня репозитория:  bash docker/build_release.sh
# Результат: dist/hut_<дата>_<коммит>.tar.gz и его sha256 — переносится на сервер (см. README, «Выкат на сервер»).
set -euo pipefail
cd "$(dirname "$0")/.."
[ -f docker/.env ] || cp docker/.env.template docker/.env      # compose требует файл; ключ на сервере свой
tag="$(date +%Y-%m-%d)_$(git rev-parse --short HEAD)"
docker compose -f docker/docker-compose.yml build
mkdir -p dist
out="dist/hut_${tag}.tar.gz"
docker save hut:latest | gzip > "$out"
sha256sum "$out" | tee "$out.sha256"
ls -lh "$out"
