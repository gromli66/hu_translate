# -*- coding: utf-8 -*-
"""Стенд точности вёрстки переведённых PDF (сканы форм и страницы с текстовым слоем): извлечение → перевод (без смысловой проверки) → сборка → метрики.

Метрики по каждой странице переведённого PDF:
  наложения  — пары строк перевода, перекрытых больше чем на четверть меньшей (текст на тексте);
  на линейке — строки перевода, через которые проходит линейка таблицы (текст вылез из ячейки);
  мелкий     — строки кеглем меньше 4 pt (не читается);
  DOCX       — слов PDF, потерянных в DOCX.
usage: python tools/layout_bench.py <папка с PDF> <рабочая папка> [--project projects/paks/project.json]"""
import re
import sys
import json
import argparse
import collections
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.stdout.reconfigure(encoding="utf-8")

import fitz
import pipeline as P
import segments as S


def page_lines(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            txt = "".join(s["text"] for s in ln["spans"]).strip()
            if txt:
                out.append({"bbox": fitz.Rect(ln["bbox"]), "size": max(s["size"] for s in ln["spans"]), "t": txt})
    return out


def overlaps(lines):
    n = 0
    for i, a in enumerate(lines):
        for b in lines[i + 1:]:
            inter = a["bbox"] & b["bbox"]
            if inter.is_valid and not inter.is_empty:
                small = min(a["bbox"].get_area(), b["bbox"].get_area()) or 1
                if inter.get_area() / small > 0.25:
                    n += 1
    return n


def crossings(lines, hs, vs):
    n = 0
    for ln in lines:
        r = ln["bbox"]
        if any(r.y0 + 1.5 < y < r.y1 - 1.5 and min(r.x1, x1) - max(r.x0, x0) > 5 for x0, x1, y in hs) or \
                any(r.x0 + 1.5 < x < r.x1 - 1.5 and min(r.y1, y1) - max(r.y0, y0) > 0.5 * r.height for y0, y1, x in vs):
            n += 1
    return n


def docx_lost(pdf, docx):
    from docx import Document
    from docx.oxml.ns import qn
    W = lambda t: collections.Counter(re.findall(r"[А-Яа-яЁё]{3,}|[A-Z0-9]{4,}", t))
    d = Document(docx)
    parts = []
    for p in d.element.body.iter(qn("w:p")):
        if any(q is not p for q in p.iter(qn("w:p"))):
            continue
        parts.append("".join((el.text or "") if el.tag == qn("w:t") else " " for el in p.iter()
                             if el.tag in (qn("w:t"), qn("w:br"), qn("w:tab"))))
    pw, dw = W(" ".join(p.get_text() for p in fitz.open(pdf))), W(" ".join(parts))
    return sum((pw - dw).values()), sum(pw.values())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("work")
    ap.add_argument("--project", default=str(ROOT / "projects" / "paks" / "project.json"))
    a = ap.parse_args()
    prj = P.load_project(a.project)
    work = Path(a.work)
    P.extract([a.inputs], work)
    P.translate(work, prj)
    P.assemble(work)
    tot = collections.Counter()
    for jf in P.doc_files(work):
        d = json.loads(jf.read_text(encoding="utf-8"))
        ocr = d["info"].get("ocr_pages", [])
        ru = work / "out" / f"{jf.stem}_RU.pdf"
        src = fitz.open(d["file"])
        doc = fitz.open(ru)
        done = sorted({s["page"] for s in d["segs"]})
        for k, pno in enumerate(done):
            lines = page_lines(doc[k])
            hs, vs = S.page_rules(src[pno]) if hasattr(S, "page_rules") else ([], [])
            # «на линейке» — только для сканов: на страницах с текстовым слоем метрика шумит на рамках полей
            # (у исходного венгерского feladat — 437)
            m = {"строк": len(lines), "наложения": overlaps(lines), "на линейке": crossings(lines, hs, vs) if pno in ocr else 0,
                 "мелкий": sum(ln["size"] < 4 for ln in lines),
                 "сегментов": sum(1 for s in d["segs"] if s["page"] == pno)}
            tot.update(m)
            print(f"{jf.stem} стр.{pno + 1} ({'скан' if pno in ocr else 'текст'}): " + ", ".join(f"{k_} {v}" for k_, v in m.items()))
        docx = work / "out" / f"{jf.stem}_RU.docx"
        if docx.exists():
            lost, allw = docx_lost(ru, docx)
            tot.update({"DOCX потеряно": lost, "DOCX слов": allw})
            print(f"{jf.stem}: DOCX потеряно {lost} из {allw}")
    print("ИТОГО: " + ", ".join(f"{k} {v}" for k, v in tot.items()))


if __name__ == "__main__":
    main()
