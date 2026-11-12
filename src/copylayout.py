# -*- coding: utf-8 -*-
"""Раскладка копии скана по строкам Surya (решение Максима 12.11 «встраивай»; проверка —
pid_pipeline_deploy/_scratch/scan_layout_2026-11/REPORT.md, «Что брать в переводчик», пп. 3–5).

Текст и деление на абзацы — у портала (копия, rucopy), положение строк — со скана: строки Surya (их многоугольники,
tools/surya_read.py) сопоставляются с кусками копии по тексту, и по краям строк ставится:
- выравнивание и отступ коротких кусков (1–2 строки на скане) — пункты списка, «По месту требования», заголовки;
- отступ первой строки длинных абзацев (1,5–8 высот строки);
- куски на одной строке скана («(подпись)» ⇥ «Сливкин Д.Н.», «Директор» ⇥ «Е.А. Киреев») — одним абзацем через
  табуляцию с позицией второго куска, как на скане (портал ставит их друг под другом по центру).
Кусок — абзац копии или его строка между переносами (<br> портала: пункты списка в одном <p>); абзац с такими
кусками делится на абзацы по переносам, только если хоть одному куску что-то ставится.
Не сопоставилось уверенно — остаётся как у портала. Текст скан не меняет никогда. Таблицы (ячейки) не трогаются:
их куски только забирают свои строки, чтобы те не ушли в абзацы.
Слова Surya 0.17 непригодны (одна рамка на всю строку) — только строки. «По ширине» со скана не берётся: концы строк
режут печати, Surya рвёт строку по широким пробелам."""
import math
import re
import statistics as st
from copy import deepcopy

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
LAT = str.maketrans("abcehkmoptxyABCEHKMOPTXY", "авсенкмортхуАВСЕНКМОРТХУ")
LINE_MIN = 0.5          # строка Surya — куску, в котором есть не меньше половины её трёхбуквенных кусков
COVER_MIN = 0.6         # кусок принимается, если его строки на скане покрывают не меньше 60 % его текста
HEAD_MIN = 0.5          # и первая строка содержит его начало («Регистрационный номер…» под печатью — Surya прочла
                        # правую половину, и выходил отступ 354 px, Декларация о соответствии)
SHORT_MIN = 4           # строка Surya короче (букв и цифр) — не сопоставляется: «от», обрывки печатей
LONG_MIN = 10           # строка не короче — сдвигает порядок чтения


def norm(t):
    return re.sub(r"[^0-9a-zа-яё]+", "", t.translate(LAT).lower())


def grams(t, n=3):
    return {t[i:i + n] for i in range(max(len(t) - n + 1, 1))} if t else set()


def is_break(el):
    """Перенос строки (не разрыв листа или колонки)."""
    return el.tag == W + "br" and el.get(W + "type") in (None, "textWrapping")


def pieces(p):
    """Текст абзаца по кускам между переносами строки."""
    out = [""]
    for r in p.iter(W + "r"):
        for c in r:
            if c.tag == W + "t":
                out[-1] += c.text or ""
            elif c.tag == W + "tab":
                out[-1] += "\t"
            elif is_break(c):
                out.append("")
    return out


def has_drawing(p):
    return p.find(f".//{W}drawing") is not None


def ends_page(p):
    return p.find(f".//{W}br[@{W}type='page']") is not None or p.find(f"{W}pPr/{W}sectPr") is not None


def units(doc):
    """Куски копии по листам: {лист: [{"p": абзац, "k": номер куска в абзаце, "t": текст, "cell": в таблице}]}.
    Лист кончается разрывом листа или раздела (как copycheck.paragraphs)."""
    page, out = 0, {}
    for el in doc.element.body.iterchildren():
        if el.tag == W + "p":
            for k, t in enumerate(pieces(el)):
                out.setdefault(page, []).append({"p": el, "k": k, "t": t, "cell": False})
            if ends_page(el):
                page += 1
        elif el.tag == W + "tbl":
            for p in el.iter(W + "p"):
                out.setdefault(page, []).append({"p": p, "k": 0, "t": "".join(pieces(p)), "cell": True})
    return out


# ------------------------------------------------------------------ геометрия строк
def skew(lines):
    """Перекос листа — медиана наклона верхней и нижней кромок многоугольников строк шире четверти листа."""
    wide = max((ln["box"][2] for ln in lines), default=0) * 0.25
    angs = []
    for ln in lines:
        p = ln.get("poly") or []
        if len(p) != 4 or ln["box"][2] - ln["box"][0] < wide:
            continue
        for (ax, ay), (bx, by) in ((p[0], p[1]), (p[3], p[2])):
            if bx - ax > 0:
                angs.append(math.atan2(by - ay, bx - ax))
    return st.median(angs) if len(angs) >= 6 else 0.0


def deskew(lines):
    """Строки с рамкой rb — по многоугольнику, повёрнутому обратно на перекос листа (поворот вокруг начала: важны
    только положения строк друг относительно друга)."""
    a = skew(lines)
    c, s = math.cos(-a), math.sin(-a)
    out = []
    for ln in lines:
        pts = ln.get("poly") or [[ln["box"][0], ln["box"][1]], [ln["box"][2], ln["box"][3]]]
        q = [(x * c - y * s, x * s + y * c) for x, y in pts]
        out.append(dict(ln, rb=[min(p[0] for p in q), min(p[1] for p in q), max(p[0] for p in q), max(p[1] for p in q)]))
    return out, a


def rows_of(lines):
    """Ряды — строки, перекрытые по высоте больше чем на половину меньшей; ряды сверху вниз, в ряду — слева направо."""
    rows = []
    for ln in sorted(lines, key=lambda ln: (ln["rb"][1] + ln["rb"][3]) / 2):
        b = ln["rb"]
        if rows:
            cur = rows[-1]
            r0, r1 = st.median(x["rb"][1] for x in cur), st.median(x["rb"][3] for x in cur)
            if min(b[3], r1) - max(b[1], r0) > 0.5 * min(b[3] - b[1], r1 - r0):
                cur.append(ln)
                continue
        rows.append([ln])
    return [sorted(r, key=lambda ln: ln["rb"][0]) for r in rows]


def column(lines, hmed):
    """Колонка текста — «частые края»: самый левый левый край, общий у 3+ строк, и самый правый правый край, общий у
    3+ строк (медиана длинных строк давала поле на уровне отступа пунктов списка); нет — 5-й и 95-й процентили."""
    x0s, x1s = sorted(ln["rb"][0] for ln in lines), sorted(ln["rb"][2] for ln in lines)
    t0 = 0.8 * hmed
    fl = [v for v in x0s if sum(abs(v - u) <= t0 for u in x0s) >= 3]
    fr = [v for v in x1s if sum(abs(v - u) <= t0 for u in x1s) >= 3]
    if fl and fr and max(fr) - min(fl) > 0.3 * x1s[-1]:
        return min(fl), max(fr), x0s
    return x0s[len(x0s) // 20], x1s[len(x1s) * 19 // 20], x0s


# ------------------------------------------------------------------ решения по листу
def plan_page(us, lines, dpi):
    """Куски листа + строки Surya → ({номер куска: решение}, [связки кусков одной строки]).
    Решение: {"align": 0|1|2|None, "left": pt|None, "first": pt|None}; связка: [(номер куска, ("left"|"right", pt))]
    — первый кусок с позицией None."""
    ls = [ln for ln in lines if len(norm(ln.get("t") or "")) >= 2]
    if len(ls) < 5:
        return {}, []
    ls, _ = deskew(ls)
    hmed = st.median(ln["rb"][3] - ln["rb"][1] for ln in ls)
    left, right, x0s = column(ls, hmed)
    width = right - left
    if width <= 0:
        return {}, []
    tol = max(0.8 * hmed, 0.012 * width)
    pt = 72.0 / dpi
    rows = rows_of(ls)
    order = [ln for r in rows for ln in r]
    for ri, r in enumerate(rows):
        for ln in r:
            ln["row"] = ri
    ug = [grams(norm(u["t"])) if len(norm(u["t"])) >= 3 else set() for u in us]
    # строка → кусок: больше всего общих трёхбуквенных кусков; при равных — ближний вперёд от куска предыдущей строки
    # (портал пишет в порядке чтения; повтор «ООО «Авангард-НСК»» в шапке и в абзаце иначе путался)
    # короткая строка (обрывок печати «PEHO.», «Директор») порядок не сдвигает: «Директор» есть и в абзаце выше
    # («назначен новый директор»), и обрывок печати уводил порядок назад — Информационное письмо
    prev = 0
    for ln in order:
        t = norm(ln["t"])
        ln["u"] = None
        if len(t) < SHORT_MIN:
            continue
        g = grams(t)
        sc = [len(g & u) / len(g) for u in ug]
        top = max(sc, default=0)
        if top < LINE_MIN:
            continue
        cand = [k for k, v in enumerate(sc) if v >= top - 0.05]
        ln["u"] = min(cand, key=lambda k: (k - prev) if k >= prev else 3 * (prev - k))
        if len(t) >= LONG_MIN:
            prev = ln["u"]
    # строки куска — подряд в порядке чтения; короткие строки вдали от основного блока куска (обрывки печатей) — не
    # его; длинные вдали — кусок спутан, не берётся
    bad = set()
    for k in range(len(us)):
        at = [i for i, ln in enumerate(order) if ln["u"] == k]
        if not at:
            continue
        runs = [[at[0]]]
        for i in at[1:]:
            if any(order[j]["u"] not in (None, k) for j in range(runs[-1][-1] + 1, i)):
                runs.append([i])
            else:
                runs[-1].append(i)
        size = lambda run: sum(len(norm(order[i]["t"])) for i in run)
        main = max(runs, key=size)
        for run in runs:
            if run is main:
                continue
            if size(run) >= LONG_MIN:
                bad.add(k)
            else:
                for i in run:
                    order[i]["u"] = None
    ok = {}
    for k, u in enumerate(us):
        mem = [ln for ln in order if ln["u"] == k]
        if not mem or not ug[k] or k in bad:
            continue
        got = set().union(*(grams(norm(ln["t"])) for ln in mem))
        if len(ug[k] & got) / len(ug[k]) < COVER_MIN:
            continue
        rws = {}
        for ln in mem:
            rws.setdefault(ln["row"], []).append(ln)
        rr = [rws[i] for i in sorted(rws)]
        head = grams(norm(u["t"])[:12])
        first = set().union(*(grams(norm(ln["t"])) for ln in rr[0]))
        if len(head & first) / len(head) < HEAD_MIN:
            continue
        ok[k] = [[min(ln["rb"][0] for ln in r), min(ln["rb"][1] for ln in r), max(ln["rb"][2] for ln in r),
                  max(ln["rb"][3] for ln in r)] for r in rr]
    stop = lambda x: sum(abs(x - u) <= 0.8 * hmed for u in x0s) >= 4   # «табулятор»: край, общий у 4+ строк листа
    sides = lambda lg, rg: lg > 2 * tol and rg > 2 * tol and abs(lg - rg) <= max(tol, 0.03 * width)
    dec = {}
    for k, bx in ok.items():
        if us[k]["cell"]:
            continue
        x0, x1 = [q[0] for q in bx], [q[2] for q in bx]
        lg, rg = x0[0] - left, right - max(x1)
        d = None
        if len(bx) == 1:
            if lg <= 2 * tol and rg <= 2 * tol:
                # строка почти во всю колонку: по центру она или влево, в допуск не различить («ФЕДЕРАЛЬНОЕ
                # АГЕНТСТВО…» по центру листа выходил влево с отступом, Крепеж сертификат) — как у портала
                continue
            if sides(lg, rg) and not stop(x0[0]):
                d = {"align": 1, "left": 0.0}
            elif rg <= tol and lg > 0.25 * width:
                d = {"align": 2, "left": 0.0}
            elif lg > tol and rg <= 2 * tol:
                # строка с отступом до правого края — красная строка абзаца («В связи с этим, …»): в копии он может
                # перенестись, и продолжение — от поля, а не под отступом
                d = {"align": 0, "left": 0.0, "first": round(lg * pt, 1)}
            else:
                d = {"align": 0, "left": round(lg * pt, 1) if lg > tol else 0.0}
            d.setdefault("first", 0.0)
        elif len(bx) == 2:
            c = [(p + q) / 2 for p, q in zip(x0, x1)]
            if abs(c[0] - c[1]) <= tol and sides(min(x0) - left, right - max(x1)):
                d = {"align": 1, "left": 0.0, "first": 0.0}
            elif abs(x0[0] - x0[1]) <= tol:
                m = min(x0) - left
                if m > tol and all(q[2] - q[0] >= 0.8 * width for q in bx):
                    continue        # обе строки почти во всю колонку и не от поля: по центру или с отступом — не
                    # различить (мелкий текст под рамкой сертификата выходил влево с отступом, Крепеж сертификат)
                d = {"align": 0, "left": round(m * pt, 1) if m > tol else 0.0, "first": 0.0}
            elif 1.5 * hmed <= x0[0] - x0[1] <= 8 * hmed:
                m = x0[1] - left
                d = {"align": 0, "left": round(m * pt, 1) if m > tol else 0.0, "first": round((x0[0] - x0[1]) * pt, 1)}
            elif abs(x1[0] - x1[1]) <= tol and rg <= tol and lg > 0.25 * width:
                d = {"align": 2, "left": 0.0, "first": 0.0}
        else:
            rest = x0[1:]
            r0 = st.median(rest)
            if max(rest) - min(rest) <= tol and abs(r0 - left) <= tol and 1.5 * hmed <= x0[0] - r0 <= 8 * hmed:
                d = {"align": None, "left": None, "first": round((x0[0] - r0) * pt, 1)}
        if d:
            dec[k] = d
    # куски одной строки: ряд из кусков подряд идущих абзацев, каждый — целиком в этом ряду
    links = []
    for r in rows:
        seq = []
        for ln in r:
            if seq and seq[-1][0] == ln["u"]:
                seq[-1][1].append(ln)
            else:
                seq.append((ln["u"], [ln]))
        seq = [(k, g) for k, g in seq if k is not None]
        if len(seq) < 2 or any(k not in ok or us[k]["cell"] or len(ok[k]) != 1 for k, _ in seq):
            continue
        ks = [k for k, _ in seq]
        if any(not between_empty(us, a, b) for a, b in zip(ks, ks[1:])):
            continue
        gaps = [g2[0]["rb"][0] - g1[-1]["rb"][2] for (_, g1), (_, g2) in zip(seq, seq[1:])]
        if min(gaps) <= 1.5 * hmed or max(gaps) <= 3 * hmed:
            continue
        link = [(ks[0], None)]
        for j, (k, g) in enumerate(seq[1:], 1):
            x0, x1 = g[0]["rb"][0], g[-1]["rb"][2]
            if j == len(seq) - 1 and right - x1 <= tol:
                link.append((k, ("right", None)))
            else:
                link.append((k, ("left", round((x0 - left) * pt, 1))))
        lg = ok[ks[0]][0][0] - left
        dec[ks[0]] = {"align": 0, "left": round(lg * pt, 1) if lg > tol else 0.0, "first": 0.0}
        for k in ks[1:]:
            dec.pop(k, None)
        links.append(link)
    return dec, links


WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def loose(p):
    """Абзац без текста, в котором разве что картинки поверх текста от края листа (печать, подпись — rucopy.place):
    его можно слить с соседним без сдвига картинок."""
    if p.find(f".//{WP}inline") is not None:
        return False
    return all(a.find(f"{WP}positionH").get("relativeFrom") == "page" == a.find(f"{WP}positionV").get("relativeFrom")
               for a in p.iter(WP + "anchor"))


def between_empty(us, a, b):
    """Куски a < b идут подряд в копии: между ними только абзацы без текста — пустые или с подписью/печатью поверх
    текста («Руководитель органа по сертификации» ⇥ «В.В. Ерофеев», между ними — абзац подписи, Крепеж сертификат)."""
    if b <= a:
        return False
    for u in us[a + 1:b]:
        if u["cell"] or u["t"].strip() or not loose(u["p"]):
            return False
    return True


# ------------------------------------------------------------------ правка Word
def split(p):
    """Абзац по переносам строки → [абзацы] (первый — сам абзац, остальные — сразу за ним, с тем же оформлением
    абзаца; отбивка после — только у последнего, чтобы лист не удлинялся; разрыв раздела — у последнего)."""
    from docx.oxml import OxmlElement
    ppr = p.find(W + "pPr")
    kids = [c for c in p if c is not ppr]
    groups = [[]]
    for c in kids:
        p.remove(c)
        if c.tag == W + "r" and any(is_break(x) for x in c):
            rpr = c.find(W + "rPr")

            def fresh():
                r = OxmlElement("w:r")
                if rpr is not None:
                    r.append(deepcopy(rpr))
                return r
            part = fresh()
            for x in [x for x in c if x is not rpr]:
                if is_break(x):
                    if len(part) > (rpr is not None):
                        groups[-1].append(part)
                    groups.append([])
                    part = fresh()
                else:
                    part.append(x)
            if len(part) > (rpr is not None):
                groups[-1].append(part)
        else:
            groups[-1].append(c)
    sect = ppr.find(W + "sectPr") if ppr is not None else None
    if sect is not None:
        ppr.remove(sect)
    out, prev = [], p
    for i, g in enumerate(groups):
        if i == 0:
            el = p
        else:
            el = OxmlElement("w:p")
            if ppr is not None:
                el.append(deepcopy(ppr))
            prev.addnext(el)
            prev = el
        for c in g:
            el.append(c)
        out.append(el)
    if sect is not None:
        last = out[-1]
        lp = last.find(W + "pPr")
        if lp is None:
            lp = OxmlElement("w:pPr")
            last.insert(0, lp)
        lp.append(sect)
    if len(out) > 1:
        from docx.shared import Pt
        from docx.text.paragraph import Paragraph
        for el in out[:-1]:
            Paragraph(el, None).paragraph_format.space_after = Pt(0)
        # пробел у переноса (портал пишет «…,<br> и с новой строки «    - разделение…») — не текст: в начале абзаца
        # он сдвигал строку
        for el in out:
            ts = [t for t in el.iter(W + "t")]
            for t in ts:
                t.text = (t.text or "").lstrip()
                if t.text:
                    break
            for t in reversed(ts):
                t.text = (t.text or "").rstrip()
                if t.text:
                    break
    return out


def text_width(doc, p):
    """Ширина колонки (pt) раздела, к которому относится абзац (свойства раздела — в его последнем абзаце)."""
    body = doc.element.body
    seen = False
    for el in body.iterchildren():
        seen = seen or el is p
        s = el.find(f"{W}pPr/{W}sectPr") if el.tag == W + "p" else (el if el.tag == W + "sectPr" else None)
        if seen and s is not None:
            sz, mar = s.find(W + "pgSz"), s.find(W + "pgMar")
            if sz is not None and mar is not None:
                tw = lambda e, a: int(e.get(W + a) or 0) / 20
                return tw(sz, "w") - tw(mar, "left") - tw(mar, "right")
    return None


def apply(doc, surya_pages, dpi):
    """Раскладка копии по строкам Surya на месте. surya_pages — {лист: [{"t", "box", "poly"}]} (строки листа в px при
    dpi). → (кусков с раскладкой со скана, связок кусков одной строки)."""
    from docx.enum.text import WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    n_dec = n_link = 0
    for page, us in units(doc).items():
        dec, links = plan_page(us, surya_pages.get(page) or [], dpi)
        if not dec and not links:
            continue
        touched = set(dec) | {k for link in links for k, _ in link}
        # абзацы с несколькими кусками, где хоть одному что-то ставится, — делятся по переносам
        el_of, split_done = {}, {}
        for k, u in enumerate(us):
            p = u["p"]
            multi = sum(1 for v in us if v["p"] is p) > 1
            if multi and any(j in touched for j, v in enumerate(us) if v["p"] is p):
                if id(p) not in split_done:
                    split_done[id(p)] = split(p)
                el_of[k] = split_done[id(p)][u["k"]]
            else:
                el_of[k] = p
        for k, d in dec.items():
            pf = Paragraph(el_of[k], None).paragraph_format
            if d["align"] is not None:
                pf.alignment = d["align"]
            if d["left"] is not None:
                pf.left_indent = Pt(d["left"])
            if d["first"] is not None:
                pf.first_line_indent = Pt(d["first"])
            n_dec += 1
        for link in links:
            els = [el_of[k] for k, _ in link]
            ks = [k for k, _ in link]
            # абзацы без текста между кусками (пустые, с подписью поверх текста) — в первый: картинки от края листа
            # не сдвигаются, а пустая строка между кусками одной строки скана не нужна
            mid = [el_of[j] for a, b in zip(ks, ks[1:]) for j in range(a + 1, b)]
            if any(ends_page(el) for el in els[1:] + mid):         # разрыв листа или раздела — не сливать
                continue
            first = els[0]
            for el in dict.fromkeys(mid):
                for c in [c for c in el if c.tag != W + "pPr"]:
                    first.append(c)
                el.getparent().remove(el)
            tw = text_width(doc, first)
            pf = Paragraph(first, None).paragraph_format
            for el, (_, (kind, pos)) in zip(els[1:], link[1:]):
                r = OxmlElement("w:r")
                r.append(OxmlElement("w:tab"))
                first.append(r)
                for c in [c for c in el if c.tag != W + "pPr"]:
                    first.append(c)
                el.getparent().remove(el)
                if kind == "right" and tw:
                    pf.tab_stops.add_tab_stop(Pt(tw), WD_TAB_ALIGNMENT.RIGHT)
                else:
                    pf.tab_stops.add_tab_stop(Pt(min(pos, (tw or pos + 20) - 20)), WD_TAB_ALIGNMENT.LEFT)
            n_link += 1
    return n_dec, n_link
