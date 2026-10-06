#!/usr/bin/env bash
# Архив для сервера: образ + всё для установки (install.sh, hut, docker-compose.yml, ADMIN.md, проекты с глоссариями).
# Запускать на машине с интернетом из корня репозитория:  bash docker/build_release.sh
# Результат: dist/hut_<дата>_<коммит>.tar.gz — на сервере: tar xzf … && cd hut && ./install.sh
# Архив содержит глоссарии заказчика (projects/*/data) — не выкладывать за пределы контура.
set -euo pipefail
cd "$(dirname "$0")/.."
tag="$(date +%Y-%m-%d)_$(git rev-parse --short HEAD)"
docker build -f docker/Dockerfile -t hut:latest .
rm -rf dist/hut && mkdir -p dist/hut
cp deploy/install.sh deploy/hut deploy/docker-compose.yml deploy/ADMIN.md dist/hut/
chmod +x dist/hut/install.sh dist/hut/hut
cp -r projects dist/hut/projects
docker save hut:latest | gzip > dist/hut/hut-image.tar.gz
out="dist/hut_${tag}.tar.gz"
tar czf "$out" -C dist hut
rm -rf dist/hut
sha256sum "$out" | tee "$out.sha256"
ls -lh "$out"
