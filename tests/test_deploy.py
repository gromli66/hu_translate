# -*- coding: utf-8 -*-
"""Команды администратора: новый проект с глоссарием заказчика (./hut project add → manage.py project add)."""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MD = ("# Г\n\n## Правила применения\n\n- коды не переводятся\n\n## Системы\n\n"
      "| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n| szivattyú | насос | строгий |\n")


def run(tmp_path, *args):
    env = {"HUT_DATA": str(tmp_path / "data"), "HUT_PROJECTS": str(tmp_path / "projects"), "PYTHONIOENCODING": "utf-8",
           "SYSTEMROOT": __import__("os").environ.get("SYSTEMROOT", ""), "PATH": __import__("os").environ.get("PATH", "")}
    return subprocess.run([sys.executable, str(ROOT / "manage.py"), *args], capture_output=True, text=True,
                          encoding="utf-8", env=env, cwd=ROOT)


def test_project_add_creates_engine_glossary(tmp_path):
    src = tmp_path / "customer.md"
    src.write_text(MD, encoding="utf-8")
    r = run(tmp_path, "project", "add", "uj-komplekt", str(src))
    assert r.returncode == 0, r.stderr
    assert "терминов 2" in r.stdout
    pj = tmp_path / "projects" / "uj-komplekt" / "project.json"
    assert json.loads(pj.read_text(encoding="utf-8"))["glossary"] == "data/glossary.md"
    sys.path.insert(0, str(ROOT / "src"))
    import glossary as G
    rules, entries = G.load(pj.parent / "data" / "glossary.md")
    assert rules == "- коды не переводятся" and [(e.hu, e.ru) for e in entries] == [("szelep", "клапан"), ("szivattyú", "насос")]
    assert "строгий" in entries[1].note
    again = run(tmp_path, "project", "add", "uj-komplekt", str(src))           # второй раз — отказ с подсказкой
    assert again.returncode != 0 and "Импорт глоссария заказчика" in again.stderr
    assert "uj-komplekt" in run(tmp_path, "project", "list").stdout
