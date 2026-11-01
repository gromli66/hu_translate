# -*- coding: utf-8 -*-
"""Внутри контейнера `hut` (запускает tools/gate_baseline.py, руками не нужен):
    python gate_asm_all.py <src_dir> <out_dir> <corpus.txt>
собирает кодом <src_dir> каждый JSON из <corpus.txt> (по пути на строку, по одному процессу Python — как ночной
`asm_all.py`) в <out_dir> и печатает на каждый документ строку `GATE<TAB>{json}`: имя, секунды, листов PDF, откуда
взят PDF (`info["pdf"]["from"]`, у сканов и титулов `info` без `pdf` — PDF «на месте») и ошибка, если сборка упала."""
import json
import sys
import time
import traceback
from pathlib import Path

SRC, OUT, CORPUS = sys.argv[1], Path(sys.argv[2]), Path(sys.argv[3])
sys.path.insert(0, SRC)
from assemble import assemble_all  # noqa: E402
import pymupdf  # noqa: E402


def record(stem, sec, pdf_pages, info, err=None):
    pdf = (info or {}).get("pdf")
    return dict(stem=stem, sec=round(sec, 1), pdf_pages=pdf_pages, err=err,
                pdf_from=pdf.get("from") if isinstance(pdf, dict) else None,
                pdf_reason=pdf.get("reason") if isinstance(pdf, dict) else None)


OUT.mkdir(parents=True, exist_ok=True)
files = [ln.strip() for ln in CORPUS.read_text(encoding="utf-8").splitlines() if ln.strip()]
print("всего JSON:", len(files), flush=True)
for i, jf in enumerate(files, 1):
    stem = Path(jf).stem
    t = time.time()
    try:
        d = json.loads(Path(jf).read_text(encoding="utf-8"))
        info = assemble_all(jf, d, OUT)
        pdf = OUT / f"{stem}_RU.pdf"
        n = len(pymupdf.open(str(pdf))) if pdf.exists() else 0
        rec = record(stem, time.time() - t, n, info)
    except Exception as e:                              # упавший документ — строка в отчёте, а не конец прогона
        traceback.print_exc()
        rec = record(stem, time.time() - t, 0, None, "%s: %s" % (type(e).__name__, str(e)[:200]))
    print("[%d/%d] %s" % (i, len(files), stem[:50]), flush=True)
    print("GATE\t" + json.dumps(rec, ensure_ascii=False), flush=True)
print("DONE", flush=True)
