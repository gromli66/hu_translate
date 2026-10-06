# -*- coding: utf-8 -*-
"""Перевод комплекта эксплуатационных документов HU→RU на портале go.ai-rosatom.ru.

  python translate_komplekt.py run <документы или папка> --project projects/paks/project.json --work work/paks
      извлечение → перевод разделами → смысловая проверка с исправлением → DOCX/PDF + таблицы вычитки в <work>/out
  python translate_komplekt.py edits --project … --work …
      правки редактора (колонка G таблиц вычитки) → перевод, память переводов проекта, похожие фразы по образцу
  python translate_komplekt.py terms <документы> --project … --work … --out terms.md
      кандидаты в термины нового комплекта, переведённые порталом, — на проверку специалисту
Шаги по отдельности: extract / translate / review / assemble. Ответы портала кэшируются: перезапуск после обрыва дешёвый."""
import sys, argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
sys.stdout.reconfigure(encoding="utf-8")
import pipeline as P


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("run", "extract", "translate", "review", "assemble", "edits", "terms"):
        sp = sub.add_parser(name)
        if name in ("run", "extract", "terms"):
            sp.add_argument("inputs", nargs="+", help="файлы DOCX/PDF или папки с ними")
        sp.add_argument("--project", required=name != "extract", help="project.json проекта")
        sp.add_argument("--work", required=True, help="рабочая папка комплекта (json, ocr, out)")
        sp.add_argument("--workers", type=int, default=4, help="одновременных запросов перевода (по умолчанию 4)")
        sp.add_argument("--review-workers", type=int, default=6, help="одновременных запросов проверки (по умолчанию 6)")
        sp.add_argument("--no-review", action="store_true", help="без смысловой проверки (быстрее вдвое)")
        sp.add_argument("--redo", action="store_true", help="перевести заново уже переведённые документы")
        sp.add_argument("--sim", type=float, default=0.8, help="edits: порог похожести фраз для перевода по образцу")
        sp.add_argument("--out", help="terms: куда записать md")
        sp.add_argument("--top", type=int, default=300, help="terms: сколько кандидатов")
    a = ap.parse_args()
    prj = P.load_project(a.project) if a.project else None
    if a.cmd in ("run", "extract", "terms"):
        P.extract(a.inputs, a.work)
    if a.cmd in ("run", "translate"):
        P.translate(a.work, prj, a.workers, redo=a.redo)
    if a.cmd == "review" or (a.cmd == "run" and not a.no_review):
        P.review(a.work, prj, a.review_workers)
    if a.cmd in ("run", "assemble"):
        P.assemble(a.work)
    if a.cmd == "edits":
        P.edits(a.work, prj, a.sim, a.workers)
    if a.cmd == "terms":
        P.terms(a.work, prj, a.out or str(Path(a.work) / "terms_candidates.md"), a.top)


if __name__ == "__main__":
    main()
