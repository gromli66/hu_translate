# -*- coding: utf-8 -*-
"""Подставной исполнитель задачи для тестов сервиса: без портала, повторяет контракт server/runner.py
(progress.json, work/out). Файл с «fail» в имени — задача падает. FAKE_SLEEP — сколько «работать»."""
import os
import sys
import json
import time
from pathlib import Path

jd = Path(sys.argv[1])
mode = sys.argv[2] if len(sys.argv) > 2 else "translate"
tok = os.environ.get("ROSATOM_AI_TOKEN", "")
if mode == "edits":                       # применение правок: забрать отложенные правки и «пересобрать»
    (jd / "edits_pending.json").unlink(missing_ok=True)
    (jd / "progress.json").write_text(json.dumps({"stage": "готово", "pct": 100,
                                                  "summary": {"segments": 3, "issues": 0, "suspicious": 0}}), encoding="utf-8")
    sys.exit(0)
(jd / "token_seen.json").write_text(json.dumps({"len": len(tok)}), encoding="utf-8")   # сам токен не пишем
print("работаю", flush=True)
(jd / "progress.json").write_text(json.dumps({"stage": "перевод", "pct": 30, "eta_s": 120}), encoding="utf-8")
time.sleep(float(os.environ.get("FAKE_SLEEP", "0.3")))
files = sorted((jd / "in").iterdir())
if any("fail" in f.name for f in files):
    print("RuntimeError: тестовый сбой", flush=True)
    sys.exit(1)
out = jd / "work" / "out"
out.mkdir(parents=True, exist_ok=True)
for f in files:
    (out / f"{f.stem}_RU{f.suffix}").write_bytes(b"translated")
(jd / "progress.json").write_text(json.dumps({"stage": "готово", "pct": 100,
                                              "summary": {"segments": 10, "issues": 1, "suspicious": 2}}), encoding="utf-8")
