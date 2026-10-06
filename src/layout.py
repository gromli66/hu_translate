# -*- coding: utf-8 -*-
"""Раскладка страницы PDF: фрагменты строк → абзацы (единицы перевода) с рамками.

Один алгоритм для текстового слоя и для слов Tesseract:
- строка режется на фрагменты по горизонтальному разрыву > GAP (столбцы таблиц без вертикальных линий);
- фрагмент присоединяется к открытому абзацу, если он ниже вплотную, перекрывается по x,
  между ними нет горизонтальной линейки, начертание то же и фрагмент не начинается с маркера пункта;
- для вставки перевода рамка абзаца расширяется вправо/вниз до соседей, линеек и полей."""
import re
import fitz

GAP = 30.0          # выключка по ширине в узких ячейках даёт пробелы до ~25 pt
MARK = re.compile(r"^\s*(\(?\d+(\.\d+)*\.?\)?|[a-zA-Z]\.?\)|[A-Z]\.\d*\.?|[-–•▪]|\d+/\d+\.)\s")


def _union(bbs):
    return [min(b[0] for b in bbs), min(b[1] for b in bbs), max(b[2] for b in bbs), max(b[3] for b in bbs)]


def _vrule_in(vs, xa, xb, y0, y1):
    return any(xa - 1 <= x <= xb + 1 and min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0) for ry0, ry1, x in vs)


def frags_from_text(page, vs=()):
    """Фрагменты из текстового слоя: [{t, bbox, size, bold}]. vs — вертикальные линейки."""
    d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_MEDIABOX_CLIP)
    out = []
    for b in d["blocks"]:
        if b.get("type") != 0:
            continue
        for ln in b["lines"]:
            if abs(ln["dir"][1]) > 0.1:          # повёрнутый текст не трогаем
                continue
            cur = None
            for s in ln["spans"]:
                txt = s["text"]
                if not txt.strip():
                    if cur is not None:
                        cur["t"] += txt
                    continue
                bb = list(s["bbox"])
                bold = bool(s["flags"] & 16) or "Bold" in s["font"]
                if cur is not None and bb[0] - cur["bbox"][2] <= GAP and                         not _vrule_in(vs, cur["bbox"][2], bb[0], bb[1], bb[3]):
                    cur["t"] += txt
                    cur["bbox"] = _union([cur["bbox"], bb])
                    cur["chars"].append((len(txt.strip()), s["size"], bold))
                else:
                    if cur is not None:
                        out.append(cur)
                    cur = {"t": txt, "bbox": bb, "chars": [(len(txt.strip()), s["size"], bold)]}
            if cur is not None:
                out.append(cur)
    # выключенная строка: PyMuPDF отдаёт каждое слово отдельной «строкой» — склеить по базовой линии
    out.sort(key=lambda f: (round((f["bbox"][1] + f["bbox"][3]) / 4), f["bbox"][0]))
    merged = []
    for f in out:
        m = merged[-1] if merged else None
        if m is not None and abs(m["bbox"][1] - f["bbox"][1]) < 2.5 and abs(m["bbox"][3] - f["bbox"][3]) < 2.5                 and 0 <= f["bbox"][0] - m["bbox"][2] <= GAP and not _vrule_in(vs, m["bbox"][2], f["bbox"][0], f["bbox"][1], f["bbox"][3]):
            m["t"] = m["t"].rstrip() + " " + f["t"].lstrip()
            m["bbox"] = _union([m["bbox"], f["bbox"]]); m["chars"] += f["chars"]
        else:
            merged.append(f)
    out = merged
    for f in out:
        n = sum(c[0] for c in f["chars"]) or 1
        f["size"] = max(f["chars"], key=lambda c: c[0])[1]
        f["bold"] = sum(c[0] for c in f["chars"] if c[2]) > n / 2
        f["t"] = f["t"].strip()
        del f["chars"]
    return [f for f in out if f["t"]]


def frags_from_words(words):
    """Фрагменты из слов Tesseract: words = [{t, line, box, h}] в порядке чтения."""
    out, cur = [], None
    for w in words:
        if not w["t"]:
            continue
        if cur is not None and cur["line"] == w["line"] and w["box"][0] - cur["bbox"][2] <= GAP:
            cur["t"] += " " + w["t"]; cur["bbox"] = _union([cur["bbox"], w["box"]]); cur["hs"].append(w["h"])
        else:
            if cur is not None:
                out.append(cur)
            cur = {"t": w["t"], "bbox": list(w["box"]), "line": w["line"], "hs": [w["h"]]}
    if cur is not None:
        out.append(cur)
    for f in out:
        hs = sorted(f.pop("hs")); f.pop("line")
        f["size"] = round(hs[len(hs) // 2] * 1.25, 1)   # высота заглавной/строчной с выносными → кегль, грубо
        f["bold"] = False
    return out


def rules(page):
    """Горизонтальные и вертикальные линейки (тонкие залитые прямоугольники и отрезки)."""
    hs, vs = [], []
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.height <= 2.5 and r.width >= 15:
            hs.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
        elif r.width <= 2.5 and r.height >= 8:
            vs.append((r.y0, r.y1, (r.x0 + r.x1) / 2))
        else:
            for it in dr["items"]:
                if it[0] == "l":
                    p, q = it[1], it[2]
                    if abs(p.y - q.y) < 1 and abs(p.x - q.x) >= 15:
                        hs.append((min(p.x, q.x), max(p.x, q.x), p.y))
                    elif abs(p.x - q.x) < 1 and abs(p.y - q.y) >= 8:
                        vs.append((min(p.y, q.y), max(p.y, q.y), p.x))
                elif it[0] == "re":
                    rr = it[1]
                    if rr.width >= 15 and rr.height > 2.5:     # рамка ячейки: четыре стороны
                        hs += [(rr.x0, rr.x1, rr.y0), (rr.x0, rr.x1, rr.y1)]
                        vs += [(rr.y0, rr.y1, rr.x0), (rr.y0, rr.y1, rr.x1)]
    return hs, vs


def _rule_between(hs, x0, x1, ya, yb):
    for rx0, rx1, y in hs:
        if ya - 0.5 <= y <= yb + 0.5 and min(x1, rx1) - max(x0, rx0) > 5:
            return True
    return False


def _in_cell(vs, bb):
    """Фрагмент зажат вертикальными линейками слева и справа (ячейка таблицы)."""
    x0, y0, x1, y1 = bb
    left = any(x <= x0 + 1 and x0 - x < 250 and ry0 <= y1 and ry1 >= y0 for ry0, ry1, x in vs)
    right = any(x >= x1 - 1 and x - x1 < 250 and ry0 <= y1 and ry1 >= y0 for ry0, ry1, x in vs)
    return left and right


def _right_limit(frags, vs, bb, page_right):
    """Где кончается место справа от строки: вертикальная линейка или соседний фрагмент на той же высоте."""
    x0, y0, x1, y1 = bb
    lim = page_right
    for ry0, ry1, x in vs:
        if x >= x1 - 1 and min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0):
            lim = min(lim, x)
    for f in frags:
        fx0, fy0, fx1, fy1 = f["bbox"]
        if fx0 >= x1 + 1 and min(y1, fy1) - max(y0, fy0) > 0.5 * (y1 - y0):
            lim = min(lim, fx0)
    return lim


def group(frags, hs, vs=()):
    page_right = max((f["bbox"][2] for f in frags), default=0)
    frags = sorted(frags, key=lambda f: (round(f["bbox"][1] / 3), f["bbox"][0]))
    pars = []
    for f in frags:
        x0, y0, x1, y1 = f["bbox"]
        best = None
        if not MARK.match(f["t"]):
            for p in reversed(pars[-12:]):
                px0, py0, px1, py1 = p["bbox"]
                last = p["frags"][-1]["bbox"]
                gap = y0 - last[3]
                size = max(p["size"], f["size"])
                if not (-0.35 * size <= gap <= 0.6 * size):
                    continue
                if min(x1, px1) - max(x0, px0) <= 0:
                    continue
                if p["bold"] != f["bold"] or abs(p["size"] - f["size"]) > 1.5:
                    continue
                if _rule_between(hs, x0, x1, last[3], y0):
                    continue
                # короткая строка (до преграды справа далеко) закрывает абзац
                lim = _right_limit(frags, vs, last, page_right)
                width = lim - px0
                if width > 0 and lim - last[2] > 0.45 * width:
                    continue
                best = p
                break
        if best is None:
            pars.append({"frags": [f], "bbox": list(f["bbox"]), "size": f["size"], "bold": f["bold"]})
        else:
            best["frags"].append(f); best["bbox"] = _union([best["bbox"], f["bbox"]])
    for p in pars:
        p["text"] = join_lines([f["t"] for f in p["frags"]])
    return pars


def join_lines(lines):
    out = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not out:
            out = ln
        elif out.endswith("-") and ln[:1].islower():
            keep = re.search(r"[A-ZÁÉÍÓÖŐÚÜŰ0-9]-$", out)
            out = (out if keep else out[:-1]) + ln
        else:
            out += " " + ln
    return out


def avail_rects(pars, hs, vs, page_rect):
    """Рамка под перевод: вправо до соседа/вертикальной линейки/правого поля, вниз до соседа/линейки."""
    if not pars:
        return
    wide = sorted(p["bbox"][2] for p in pars if p["bbox"][2] - p["bbox"][0] > 0.5 * page_rect.width)
    right_margin = wide[len(wide) // 2] if wide else max(p["bbox"][2] for p in pars)
    for p in pars:
        x0, y0, x1, y1 = p["bbox"]
        lim_x = max(right_margin, x1)
        for q in pars:
            if q is p:
                continue
            qx0, qy0, qx1, qy1 = q["bbox"]
            if qx0 >= x1 - 1 and min(y1, qy1) - max(y0, qy0) > 0:
                lim_x = min(lim_x, qx0 - 3)
        for ry0, ry1, x in vs:
            if x >= x1 - 0.5 and min(y1, ry1) - max(y0, ry0) > 0:
                lim_x = min(lim_x, x - 2)
        lim_y = page_rect.y1 - 20
        for q in pars:
            if q is p:
                continue
            qx0, qy0, qx1, qy1 = q["bbox"]
            if qy0 >= y1 - 1 and min(lim_x, qx1) - max(x0, qx0) > 0:
                lim_y = min(lim_y, qy0 - 1)
        for rx0, rx1, y in hs:
            if y >= y1 - 0.5 and min(lim_x, rx1) - max(x0, rx0) > 5:
                lim_y = min(lim_y, y - 1)
        p["avail"] = [x0, y0, max(x1, lim_x), max(y1, lim_y)]
