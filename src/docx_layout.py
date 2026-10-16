# -*- coding: utf-8 -*-
"""DOCX распознанного PDF (текст листа — контурами, перевод — по сегментам OCR): та же модель блоков, что у размеченного
PDF (docx_struct), но по самим сегментам листа — дальше тот же профиль оформления, сборщик и печать PDF из DOCX.

Раньше такой PDF («Печать в PDF», 3SZ19, 3PR42) получал DOCX из переведённого PDF: строки резались, налезали,
кегль у каждого сегмента свой (оценка по картинке) — Максим, 16.10: «плавающий шрифт». Здесь:
- абзац = сегмент (перевод готов, портал не нужен), кегль тела — один на документ (медиана), шрифт — по семейству;
- колонтитулы — тексты, повторяющиеся в верхней/нижней части листов (цифры не в счёт), номер листа — полем;
- заголовок — короткий полужирный абзац по центру, уровень — по глубине номера («1.1.» — 2);
- пункт с тире — по маркеру «-» в текстовом слое листа (он там есть, текст нарисован контурами), отступ — уровнем.
Таблицы не восстанавливаются: их ячейки идут абзацами в порядке чтения."""
import re
from collections import Counter
from statistics import median

FAMILY = {"monospace": "Courier New", "serif": "Times New Roman", "sans-serif": "Arial"}
NUM = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?\s+\S")


def _zone(s, H):
    y0, y1 = s["loc"]["bbox"][1], s["loc"]["bbox"][3]
    return "head" if y1 < 0.14 * H else "foot" if y0 > 0.86 * H else None


def _sig(text):
    return re.sub(r"\d+", "#", text.strip())


def _running(segs, H, pages):
    """Колонтитулы: (зона, текст без цифр) или (зона, высота строки), что повторяется хотя бы на половине листов (и не
    меньше чем на двух). По высоте — потому что OCR режет шапку по-разному: «Megnevezés … UTASÍTÁSA 72» на одном листе
    и «Verziószám: 0 Oldalszám:» отдельно — на другом (3PR42, 16.10)."""
    cnt, pos = Counter(), {}
    for s in segs:
        z = _zone(s, H)
        if z:
            cnt[(z, _sig(s["text"]))] += 1
            pos.setdefault((z, round(s["loc"]["bbox"][1] / 4)), Counter())[_sig(s["text"])] += 1
    need = max(2, 0.5 * pages)
    # на высоте — и текст в основном тот же (≥ половины листов): название раздела в рамке на той же высоте меняется
    # от раздела к разделу — это текст, а не колонтитул
    return {k for k, n in cnt.items() if n >= need} | \
        {k for k, c in pos.items() if sum(c.values()) >= need and c.most_common(1)[0][1] >= 0.5 * pages}


def _is_running(s, keys, H):
    z = _zone(s, H)
    return bool(z) and ((z, _sig(s["text"])) in keys or (z, round(s["loc"]["bbox"][1] / 4)) in keys)


def _dashes(doc):
    """Маркеры пунктов «-»/«–» текстового слоя: {лист: [(x1, y_середина)]}."""
    out = {}
    for pno, page in enumerate(doc):
        for w in page.get_text("words"):
            if w[4] in ("-", "–", "—", "•"):
                out.setdefault(pno, []).append((w[2], (w[1] + w[3]) / 2, w[3] - w[1]))
    return out


def load(src, segs):
    """Блоки распознанного PDF по сегментам. Возвращает (блоки, сведения), как docx_struct.load."""
    import fitz
    doc = fitz.open(str(src))
    W, H = doc[0].rect.width, doc[0].rect.height
    body = [s for s in segs if "shot" not in s["loc"] and s["loc"].get("bbox") and s["text"].strip()]
    run_keys = _running(body, H, len(doc))
    content = [s for s in body if not _is_running(s, run_keys, H)]
    # бегущее название раздела в рамке сверху листа (3PR42: «1.15. Átfolyás …» на каждом листе раздела) — то же, что
    # вверху предыдущего листа, — не текст (на всех листах сразу его не найти: у каждого раздела своё)
    top = {}
    for s in content:
        if s["loc"]["bbox"][3] < 0.2 * H:
            top.setdefault(s["page"], set()).add(_sig(s["text"]))
    content = [s for s in content if not (s["loc"]["bbox"][3] < 0.2 * H and _sig(s["text"]) in top.get(s["page"] - 1, ()))]
    if not content:
        return [], {"width": W, "height": H, "pages": len(doc), "artifacts": {"chars": [], "imgs": []}}
    size = round(median(s["loc"].get("size") or 10 for s in content) * 2) / 2
    font = FAMILY.get(Counter(s["loc"].get("family") for s in content).most_common(1)[0][0], "Arial")
    left = sorted(s["loc"]["bbox"][0] for s in content)[len(content) // 10]
    dashes = _dashes(doc)
    blocks, cur = [], None                                  # cur — открытый список из пунктов с тире
    last_head = {}                                          # уровень → текст последнего заголовка
    for group in _lines(sorted(content, key=lambda s: s["id"])):
        s = group[0]
        loc = s["loc"]
        x0, y0, x1, y1 = loc["bbox"]
        first = (loc.get("frags") or [loc["bbox"]])[0]
        lh = max(first[3] - first[1], 6)
        geom = {"page": s["page"] + 1, "top": y0, "bottom": max(g["loc"]["bbox"][3] for g in group), "x0": x0,
                "x1": max(g["loc"]["bbox"][2] for g in group), "first_x0": first[0], "size": size,
                "lines": max(len(g["loc"].get("frags") or [1]) for g in group)}
        mono = font == "Courier New"                        # моноширинный основной шрифт — не код (docx_build.is_code)
        runs = []
        for i, g in enumerate(group):
            # сегменты одной строки (оглавление: «1.4.» ⇥ «Tengelyelmozdulás …» ⇥ «13») — через табуляцию по месту
            # оригинала: память переводит части по отдельности, склеенный текст в ней не находился (16.10)
            t = g["text"] + ("\t" if i + 1 < len(group) else "")
            r = {"t": t, "font": font, "size": size, "b": bool(g["loc"].get("bold")), "i": False, "color": "000000",
                 "body_mono": mono}
            if i + 1 < len(group):
                if i + 2 == len(group) and re.fullmatch(r"\d{1,4}", group[-1]["text"].strip()):
                    r["rtab"] = True                        # номер листа в конце строки оглавления — к правому полю
                else:
                    r["tabs"] = [group[i + 1]["loc"]["bbox"][0]]
            runs.append(r)
        run = runs[0]
        text = "".join(r["t"] for r in runs).strip()
        m = NUM.match(text)
        if len(group) == 1 and loc.get("align") == "center" and loc.get("bold") and len(text) < 120:
            level = min(len(m.group(1).split(".")) if m else 1, 5)
            norm = re.sub(r"\s+", " ", text)
            if last_head.get(level) == norm:
                continue        # бегущее название раздела в начале каждого листа (3SZ19: 41 и 42 раза) — не новый раздел
            last_head[level] = norm
            for lv in range(level + 1, 6):
                last_head.pop(lv, None)
            blocks.append({"k": "h", "level": level, "runs": [run], "geom": geom})
            cur = None
            continue
        mark = next((d for d in dashes.get(s["page"], [])
                     if abs(d[1] - (first[1] + first[3]) / 2) < 0.6 * lh and 0 < first[0] - d[0] < 30), None)
        para = {"k": "p", "runs": runs, "geom": geom, "seg": True}   # seg — абзац собран, merge_inline не клеит
        if mark:
            if cur is None or abs(cur["x0"] - x0) > 6:
                cur = {"k": "list", "kind": "Dash", "items": [], "x0": x0}
                blocks.append(cur)
            cur["items"].append([para])
            continue
        if cur is not None and cur["items"] and abs(first[0] - cur["x0"]) < 8 and loc.get("align") != "center":
            cur["items"][-1].append(para)                   # продолжение пункта (OCR разбил пункт на сегменты)
            continue
        cur = None
        indent = round(max(0.0, first[0] - left) / 9) * 9
        if indent >= 9 and loc.get("align") != "center":
            para["indent"] = min(indent, 90)                # отступ по оригиналу — уровнем стиля (сборщик)
        blocks.append(para)
    for b in blocks:
        b.pop("x0", None)
    return blocks, {"width": W, "height": H, "pages": len(doc), "artifacts": {"chars": [], "imgs": []},
                    "font": font, "size": size}


def _lines(segs):
    """Подряд идущие сегменты одной строки листа (тот же лист, середина первой строки — в пределах половины её
    высоты, следующий правее предыдущего) — группой."""
    out = []
    for s in segs:
        f = (s["loc"].get("frags") or [s["loc"]["bbox"]])[0]
        if out:
            p = out[-1][-1]
            pf = (p["loc"].get("frags") or [p["loc"]["bbox"]])[0]
            if p["page"] == s["page"] and abs((pf[1] + pf[3]) / 2 - (f[1] + f[3]) / 2) < 0.5 * max(pf[3] - pf[1], 6) \
                    and f[0] > pf[2] - 2:
                out[-1].append(s)
                continue
        out.append([s])
    return out


def _grid(zsegs):
    """Сетка колонтитула по сегментам одного листа: строки — по высоте, столбцы — по левым краям (±12 pt)."""
    rows = {}
    for s in sorted(zsegs, key=lambda s: (s["loc"]["bbox"][1], s["loc"]["bbox"][0])):
        key = next((k for k in rows if abs(k - s["loc"]["bbox"][1]) < 5), s["loc"]["bbox"][1])
        rows.setdefault(key, []).append(s)
    xs = []
    for x in sorted(s["loc"]["bbox"][0] for s in zsegs):
        if not xs or x - xs[-1] > 12:
            xs.append(x)
    right = max(s["loc"]["bbox"][2] for s in zsegs) + 6
    cols = [x - 4 for x in xs] + [right]
    out = {"rows": [], "widths": [cols[i + 1] - cols[i] for i in range(len(xs))], "aligns": ["left"] * len(xs),
           "inner": False, "box": True, "font": None, "size": None}
    for _, ss in sorted(rows.items()):
        cells = [""] * len(xs)
        for s in ss:
            ci = max(i for i, x in enumerate(xs) if s["loc"]["bbox"][0] >= x - 12)
            cells[ci] = (cells[ci] + " " + s["text"].strip()).strip()
        out["rows"].append({"cells": cells, "bold": all(s["loc"].get("bold") for s in ss)})
    # столбец из одних чисел («Oldalszám: | 41») — к предыдущему: узкий столбец разносил номер листа по строкам
    for ci in range(len(xs) - 1, 0, -1):
        if all(re.fullmatch(r"\d{1,4}", r["cells"][ci]) or not r["cells"][ci] for r in out["rows"]) \
                and any(r["cells"][ci] for r in out["rows"]):
            for r in out["rows"]:
                r["cells"][ci - 1] = (r["cells"][ci - 1] + " " + r["cells"][ci]).strip()
                del r["cells"][ci]
            out["widths"][ci - 1] += out["widths"][ci]
            del out["widths"][ci], out["aligns"][ci]
    return out


def layout(src, blocks, segs, info):
    """Раскладка листа (как docx_struct.layout): поля — по тексту тела, колонтитулы — сеткой по повторяющимся
    сегментам типичного листа, номер листа — число, растущее с листом."""
    import fitz
    doc = fitz.open(str(src))
    W, H = doc[0].rect.width, doc[0].rect.height
    body = [s for s in segs if "shot" not in s["loc"] and s["loc"].get("bbox") and s["text"].strip()]
    run_keys = _running(body, H, len(doc))
    runs = [s for s in body if _is_running(s, run_keys, H)]
    geoms = [b["geom"] for b in blocks if b.get("geom")] + \
        [p["geom"] for b in blocks if b["k"] == "list" for it in b["items"] for p in it if p.get("geom")]
    left = sorted(g["x0"] for g in geoms)[len(geoms) // 10] if geoms else 62
    right = W - sorted(g["x1"] for g in geoms)[-max(1, len(geoms) // 10)] if geoms else 59
    top = min((g["top"] for g in geoms), default=72)
    bottom = max((g["bottom"] for g in geoms), default=H - 72)
    by_page = {}
    for s in runs:
        by_page.setdefault(s["page"], []).append(s)
    pages = sorted(by_page)
    # образец — лист с самым частым числом сегментов колонтитулов (полный набор, а не случайно слитый OCR)
    common = Counter(len(by_page[p]) for p in pages).most_common(1)[0][0] if pages else 0
    same = [p for p in pages if len(by_page[p]) == common]
    rep = same[len(same) // 2] if same else None
    page_tok = None
    if rep is not None:
        for s in by_page[rep]:                               # число, растущее с листом: «Oldalszám: 7»
            m = re.search(r"\d+", s["text"])
            if not m:
                continue
            same = [t for p in pages for t in by_page[p] if _sig(t["text"]) == _sig(s["text"])
                    and abs(t["loc"]["bbox"][1] - s["loc"]["bbox"][1]) < 4]       # то же место на листе
            vals = Counter(int(re.search(r"\d+", t["text"]).group()) - t["page"] for t in same if re.search(r"\d+", t["text"]))
            # смещение номера от листа — одно у большинства (3PR42: нумерация сдвигается посреди документа)
            if len(same) >= 3 and vals and vals.most_common(1)[0][1] >= 0.6 * len(same):
                page_tok = m.group()
                break
    head = _grid([s for s in by_page.get(rep, []) if _zone(s, H) == "head"]) if rep is not None and \
        any(_zone(s, H) == "head" for s in by_page.get(rep, [])) else None
    footz = [s for s in by_page.get(rep, []) if _zone(s, H) == "foot"] if rep is not None else []
    foot = _grid(footz) if footz else {"rows": [], "widths": [], "aligns": [], "inner": False, "box": False}
    head_bottom = max((s["loc"]["bbox"][3] for s in by_page.get(rep, []) if _zone(s, H) == "head"), default=36)
    foot_top = min((s["loc"]["bbox"][1] for s in footz), default=H - 36)
    return {"W": W, "H": H, "left": left, "right": max(right, 20), "top": max(top - 4, head_bottom + 6),
            "bottom": max(H - bottom - 4, H - foot_top + 6), "header": 24, "footer": 18, "title_left": left,
            "head_imgs": [], "head_lines": [], "head_grid": head,
            "stamp": foot["rows"], "stamp_widths": foot["widths"], "stamp_aligns": foot["aligns"],
            "stamp_font": info.get("font") or "Times New Roman", "stamp_size": info.get("size"),
            "stamp_box": foot["box"], "stamp_inner": False, "page_tok": page_tok,
            "first_is_body": any(s["page"] == 0 for s in runs)}    # колонтитулы и на 1-м листе — титула нет
