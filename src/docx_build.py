# -*- coding: utf-8 -*-
"""Сборка DOCX из блоков docx_struct.load: стили по оригиналу, настоящие списки и нумерация заголовков,
подписи «Рисунок N. …» полем SEQ, оглавление и список рисунков полями TOC, колонтитулы со штампом.
Прямого кегля и отступов на отдельных абзацах нет — всё стилями («как будто документ писали по-русски»)."""
import io
import re
from collections import Counter

from docx_struct import text_of, insert_imgs

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
BULLETS = {"Disc": ("", "Symbol"), "Circle": ("o", "Courier New"), "Square": ("", "Wingdings"),
           "Dash": ("–", None)}                     # тире из текста пункта (ÜFK «– ha az …») — шрифтом абзаца
NUMFMT = {"Decimal": "decimal", "LowerAlpha": "lowerLetter", "UpperAlpha": "upperLetter",
          "LowerRoman": "lowerRoman", "UpperRoman": "upperRoman"}
LEAD = re.compile(r"^\s*[•▪◦·o\-–](?:\s+|$)")
CAPTION = re.compile(r"^\s*(\d+)\.\s*(ábra|táblázat)\s+(.*)$", re.S | re.I)
CAP_RU = re.compile(r"^\s*(?:(?:Рисунок|Рис\.?|рис\.?|Таблица|Табл\.?|табл\.?)\s*\d+\s*[.:]?"
                    r"|\d+\s*\.?\s*(?:рис\.?|табл\.?|рисунок|таблица)\.?)\s*", re.I)


def _el(tag, **attrs):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k.replace("__", ":")), str(v))
    return e


def _field(par, instr, result=""):
    """Поле Word (TOC, SEQ, PAGE, NUMPAGES): пересчитывается при открытии (settings: updateFields)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    r = par.add_run(); r._r.append(_el("w:fldChar", w__fldCharType="begin"))
    r = par.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "
    r._r.append(it)
    r = par.add_run(); r._r.append(_el("w:fldChar", w__fldCharType="separate"))
    res = par.add_run(result)
    r = par.add_run(); r._r.append(_el("w:fldChar", w__fldCharType="end"))
    return res


def _figure(b):
    """Абзац-рисунок (не значок в строке): одни картинки, хотя бы одна выше 24 pt."""
    imgs = [r["img"] for r in b["runs"] if "img" in r]
    return bool(imgs) and len(imgs) == len(b["runs"]) and max(float(i["bottom"]) - float(i["top"]) for i in imgs) > 24


def is_code(runs):
    """Абзац кода: больше половины символов — моноширинным (по символам, а не по прогонам: абзац «текст + имя файла
    Consolas» иначе уходил в стиль кода целиком)."""
    c = Counter()
    for r in runs:
        if "t" in r and r["t"].strip():
            # body_mono — основной шрифт документа моноширинный (3SZ19 набран Courier): это не код (docx_layout)
            c[(r.get("font") or "").lower().startswith(("consolas", "courier")) and not r.get("body_mono")] += len(r["t"].strip())
    return c[True] > c[False]


def merge_inline(blocks):
    """Соседние абзацы одной строки (перекрёстная ссылка Word — отдельный элемент: «(», «3», «. fejezet)») —
    в один абзац; пустые абзацы-отбивки — прочь (интервалы дают стили)."""
    out = []
    for b in blocks:
        if b["k"] == "list":
            b = dict(b, items=[merge_inline(it) for it in b["items"]])
        elif b["k"] == "table":
            b = dict(b, rows=[[dict(c, blocks=merge_inline(c["blocks"])) for c in r] for r in b["rows"]])
        if b["k"] == "p" and not b["runs"]:
            continue
        prev = out[-1] if out else None
        g, pg = b.get("geom") or {}, (prev or {}).get("geom") or {}
        # ссылка рядом с абзацем (Link/Span) — часть абзаца; хвост после неё на той же строке — тоже
        joins = prev is not None and prev["k"] == "p" and b["k"] == "p" and not _figure(prev) and not _figure(b) \
            and not b.get("seg") and not prev.get("seg") and (      # абзацы по сегментам листа уже собраны (docx_layout)
            b.get("inline")
            or (prev.get("tail_inline") and g and g.get("page") == pg.get("page")
                and abs(g["top"] - (prev.get("last_top") or 0)) < 6)
            or (g and pg and g["page"] == pg["page"] and abs(g["top"] - prev.get("last_top", pg["top"])) < 6
                and g["first_x0"] > pg["x0"]))
        if joins:
            runs = [dict(r) for r in b["runs"]]
            # значок в строке, нарисованный отдельно от текста: в тексте на его месте пустые скобки «()» — туда же
            # и в переводе (иначе значок уезжал в конец абзаца, а в тексте оставалось «( )»)
            free = len(re.findall(r"\(\s*\)", text_of(prev["runs"]))) - sum(1 for r in prev["runs"] if r.get("paren"))
            for r in runs:
                if "img" in r and free > 0:
                    r["paren"] = True
                    free -= 1
            if all("img" in r and not r.get("paren") for r in runs):
                merged = insert_imgs(prev["runs"], runs)        # значок без скобок — на своё место в строке
            else:
                pt, nt = text_of(prev["runs"]), text_of(runs)
                if pt and nt and "t" in runs[0] and not pt.endswith((" ", "\t", "(", "„", "«")) \
                        and not nt.startswith((" ", ".", ",", ";", ":", ")", "”", "»")):
                    runs[0]["t"] = " " + runs[0]["t"]          # «(lásd:» + «5.6.4.7» + «fejezet)» — с пробелами
                merged = prev["runs"] + runs
            out[-1] = dict(prev, runs=merged, last_top=g.get("top", prev.get("last_top")),
                           tail_inline=bool(b.get("inline")))
            continue
        out.append(b)
    return out


_HELV = None


def _text_w(text, size):
    """Ширина строки, pt: латиница — по метрике Helvetica, кириллица и прочее — средней шириной Tahoma (у Helvetica
    PDF кириллицы нет, шрифтов Word на сервере нет)."""
    global _HELV
    if _HELV is None:
        import fitz
        _HELV = fitz.Font("helv")
    return sum(_HELV.text_length(ch, fontsize=size) if ord(ch) < 256 else 0.6 * size for ch in text)


def _span_starts(row, xs, ncol):
    """Ячейки неполной строки — по столбцам рамки (xs): разбиение столбцов на len(row) подряд идущих групп, текст
    ячейки внутри своей группы и ближе всего к её середине или левому краю. Шапка ÜFK «Eltérő állapot» стоит над
    двумя столбцами «A. | текст» — раньше объединялся только хвост строки, и шапка съезжала (16.10). Без рамки —
    хвост строки, как раньше. Возвращает столбец начала каждой ячейки."""
    k = len(row)
    tail = list(range(k))
    if len(xs) != ncol + 1 or k == 0:
        return tail
    geo = [c.get("geom") for c in row]
    if all(g is None for g in geo):
        return tail
    from functools import lru_cache

    def cost(i, a, z):
        g = geo[i]
        if g is None:
            return 5.0 * (z - a)                             # пустая ячейка — по столбцу, если ничто не мешает
        gl, gr = xs[a], xs[z + 1]
        inside = 0 if gl - 3 <= g["x0"] and g["x1"] <= gr + 3 else 1000
        return inside + min(abs((gl + gr) / 2 - (g["x0"] + g["x1"]) / 2), abs(g["x0"] - gl - 5))

    @lru_cache(None)
    def best(i, a):
        if i == k:
            return (0.0, ()) if a == ncol else (float("inf"), ())
        out = (float("inf"), ())
        for z in range(a, ncol - (k - i - 1)):
            c, rest = best(i + 1, z + 1)
            c += cost(i, a, z)
            if c < out[0]:
                out = (c, (a,) + rest)
        return out

    c, starts = best(0, 0)
    return list(starts) if c < 1000 else tail


def _autonum_ok(blocks):
    """Автонумерация заголовков воспроизводит номера оригинала (UTS: 157 из 157) — иначе номера остаются текстом:
    у ÜFK «DF 1.2.1.», «BK 3.1.1.» — автонумерация давала «1.2.1 DF 1.2.1. …» (16.10); заголовки без номеров
    автонумерация нумеровала сама."""
    cnt, ok, n = [0] * 6, 0, 0
    for b in blocks:
        if b["k"] != "h":
            continue
        lv = min(b["level"], 5)
        cnt[lv - 1] += 1
        for i in range(lv, 6):
            cnt[i] = 0
        m = re.match(r"\s*(\d+(?:\.\d+)*)\.?\s", text_of(b["runs"]))
        n += 1
        ok += bool(m and m[1] == ".".join(str(c) for c in cnt[:lv]))
    return n > 0 and ok >= 0.9 * n


def _keep_tab(src, ru):
    """«подпись ⇥ значение»: перевод потерял табуляцию («Номер технического изменения 10776» — значение не в столбце,
    титул UTS, 16.10) — вернуть её перед значением, если оно стоит в конце перевода дословно."""
    if not ru or src.count("\t") != 1 or "\t" in ru:
        return ru
    val, r = src.split("\t")[1].strip(), ru.rstrip()
    if val and r.endswith(val) and len(r) > len(val):
        return r[:-len(val)].rstrip() + "\t" + val
    return ru


def _col_align(rows, xs, page):
    """Выравнивание столбцов таблицы по оригиналу (рамка xs на первом листе): {("head" | "body", столбец): "center" |
    "right"}; по однострочным ячейкам полных строк, большинством. Шапка UTS — по центру, номера версий — вправо, а
    в переводе всё стояло влево (16.10)."""
    if len(xs) < 2:
        return {}
    ncol = len(xs) - 1
    votes = {}
    head = True
    for r in rows:
        head = head and all(c["th"] for c in r)
        if len(r) != ncol:
            continue
        for ci, c in enumerate(r):
            g = c.get("geom")
            if not g or g.get("page") != page or (g.get("lines") or 1) > 1:
                continue
            left, right = g["x0"] - xs[ci], xs[ci + 1] - g["x1"]
            v = "center" if abs(left - right) < 4 and left > 8 else "right" if right < 10 < left - 8 else "left"
            votes.setdefault(("head" if head else "body", ci), []).append(v)
    out = {}
    for k, v in votes.items():
        best = Counter(v).most_common(1)[0][0]
        if best != "left":
            out[k] = best
    return out


def _base_group(text, groups, parts):
    told = [i for i, s in enumerate(parts) if s.strip() and s.strip() not in text] or list(range(len(groups)))
    return max(told, key=lambda i: len(parts[i].strip()))


def _verbatim(text, groups, parts):
    """Все куски не основного начертания стоят в переводе дословно (путь курсивом, команда моноширинным) —
    начертание расставит _restyle, перевод с метками у портала не нужен."""
    base = _base_group(text, groups, parts)
    return all(parts[i].strip() in text for i, g in enumerate(groups)
               if g[0] != groups[base][0] and parts[i].strip())


def _restyle(text, groups, parts):
    """Перевод без меток начертания (из памяти сегментов страниц) — основным начертанием абзаца (самая длинная из
    переведённых групп — тех, чего дословно в переводе нет), а куски, перешедшие в перевод дословно (путь, имя файла,
    команда курсивом или моноширинным), — своим: раньше весь абзац брал начертание первой группы, и курсив путей
    пропадал (UTS, 16.10)."""
    base = _base_group(text, groups, parts)
    spans = []
    for i, (g, p) in enumerate(zip(groups, parts)):
        s = p.strip()
        if i == base or len(s) < 3 or g[0] == groups[base][0]:
            continue
        j = text.find(s)
        if j >= 0 and not any(a < j + len(s) and j < b for a, b, _ in spans):
            spans.append((j, j + len(s), g[0]))
    out, pos = [], 0
    for a, b, key in sorted(spans):
        if a > pos:
            out.append((groups[base][0], text[pos:a]))
        out.append((key, text[a:b]))
        pos = b
    if pos < len(text) or not out:
        out.append((groups[base][0], text[pos:]))
    return out


TAGS = re.compile(r"</?f\d+>")
KEY = lambda t: re.sub(r"\s+", "", TAGS.sub("", t or ""))
HU_LETTERS = re.compile(r"[áéíóöőúüűÁÉÍÓÖŐÚÜŰ]")


def code_part(src, ru):
    """Моноширинная часть абзаца с метками. Без венгерских букв (команда, путь, вывод программы) — как в оригинале:
    портал переводил «sync» → «синхронизировать», «gatewayconfig ipaddr» → «настройка IP-адреса шлюза» (UTS, 15.10).
    Заполнитель «<beállítani kívánt IP cím>» переводится, но в тех же угловых скобках (портал их терял)."""
    if not HU_LETTERS.search(src):
        return src
    s = src.strip()
    if s.startswith("<") and s.endswith(">") and not ru.strip().startswith("<"):
        return src[:len(src) - len(src.lstrip())] + "<" + ru.strip() + ">" + src[len(src.rstrip()):]
    return ru
# метка пункта в начале сегмента страницы: «a.», «1)», составная «E.1.», «4.1.2.», «a.)», маркер-тире «–»; дерево Word
# отдаёт её отдельно (Lbl), и сборка ищет перевод текста без неё (ÜFK I: 680 абзацев не находились, 16.10)
_LBL = r"(?:(?:\d{1,3}|LL|[ivxIVX]{2,5})(?:\.(?:\d{1,3}|LL))*(?:\.\)|[.)])|[–\-•▪◦·])"
LABEL_T = re.compile(r"^(\s*(?:<f\d+>)?)\s*" + _LBL.replace("LL", "[A-Za-z]") + r"\s+(?=\S)")
LABEL_RU_T = re.compile(r"^(\s*(?:<f\d+>)?)\s*" + _LBL.replace("LL", "[A-Za-zА-Яа-яЁё]") + r"\s+(?=\S)")


class Memory:
    """Переводы для сборки: готовые (проверенные) сегменты по тексту без пробелов; «подпись ⇥ значение» — по частям;
    заголовок — и с номером («1 Előzmények» в памяти PDF). Чего нет — в missing (переводится порталом и
    добавляется через add)."""

    def __init__(self, pairs):
        from segments import needs_translation
        self.need = needs_translation
        self.m, self.tagged, self.missing = {}, {}, []
        for hu, ru in pairs:
            self.add(hu, ru)

    def add(self, hu, ru):
        if not ru:
            return
        if TAGS.search(hu):
            self.tagged.setdefault(hu.strip(), ru)
        else:
            self.m.setdefault(KEY(hu), TAGS.sub("", ru))
        # пункт списка: номер «a.» в памяти — частью текста, в сборке номер даёт автонумерация (docx_struct.relist);
        # в тексте с метками начертания номер — внутри первой метки: «<f1>a. az UTS …</f1><f2>ping x_pb</f2>»
        mh, mr = LABEL_T.match(hu), LABEL_RU_T.match(ru)
        if mh and mr:
            self.add(mh.group(1) + hu[mh.end():], mr.group(1) + ru[mr.end():])
        elif mh:                                    # перевод без метки (тире перевод опустил) — как есть
            self.add(mh.group(1) + hu[mh.end():], ru)

    def get(self, text):
        if not self.need(TAGS.sub("", text)):
            return text
        if TAGS.search(text):
            return self.tagged.get(text.strip())
        r = self.m.get(KEY(text))
        if r is None and "\t" in text:
            parts = [p for p in text.split("\t")]
            got = [self.get(p) if p.strip() else p for p in parts]
            if all(g is not None for g in got):
                r = "\t".join(g.strip() if p.strip() else g for p, g in zip(parts, got))
        return r

    def __call__(self, text):
        r = self.get(text)
        if r is None and text.strip():
            self.missing.append(text)
        return r


class Builder:
    """tr(text) → перевод или None (оставить как есть); image(img) → PNG-байты картинки (переведённый
    скриншот или оригинал) или None."""

    def __init__(self, tr, image, profile=None):
        """profile — профиль оформления оригинала (docx_profile.extract); без него — значения UTS по умолчанию."""
        from docx import Document
        self.doc, self.tr, self.image = Document(), tr, image
        self.P = profile or {}
        self.nums, self._abstract = {}, {}
        self.code_groups = {}                  # текст с метками → номера моноширинных групп (code_part)
        self.left = 62
        self.L = None
        self._toc_n = 0
        self._styles()
        self._heading_numbering()

    def _role(self, role, font="Tahoma", size=10, bold=False):
        r = (self.P.get("roles") or {}).get(role) or {}
        return r.get("font", font), r.get("size", size), r.get("b", bold)

    def _space(self, key, default):
        return (self.P.get("space") or {}).get(key, default)

    # ---------------------------------------------------------------- стили
    def _font(self, s, name="Tahoma", size=10, bold=False, italic=False):
        from docx.shared import Pt, RGBColor
        s.font.name, s.font.size, s.font.bold, s.font.italic = name, Pt(size), bold, italic
        s.font.color.rgb = RGBColor(0, 0, 0)
        rf = s.element.get_or_add_rPr().find(W + "rFonts")
        if rf is None:
            rf = _el("w:rFonts"); s.element.get_or_add_rPr().insert(0, rf)
        for k in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(W + k, name)
        for k in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
            rf.attrib.pop(W + k, None)

    def _para(self, s, before=0, after=6, line=1.08, keep=False, align=None, indent=None):
        from docx.shared import Pt
        pf = s.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), line
        pf.keep_with_next = keep
        if align is not None:
            pf.alignment = align
        if indent is not None:
            pf.left_indent = Pt(indent)

    def _styles(self):
        """Стили по профилю оригинала (Ф2 шрифт ролей, Ф3 H1 с новой страницы, Ф5 выключка, Ф6 межстрочный,
        интервалы); без профиля — значения UTS."""
        from docx.shared import Pt
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL, WD_LINE_SPACING
        st = self.doc.styles
        P = self.P
        align = P.get("align") or {}
        f, size, bold = self._role("body")
        self._font(st["Normal"], f, size, bold)
        self._para(st["Normal"], after=self._space("body", 8), line=1.15,
                   align=AL.JUSTIFY if align.get("body") == "both" else None)
        if P.get("line"):                                    # шаг строк оригинала; «не меньше» — значки в строке не режет
            st["Normal"].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            st["Normal"].paragraph_format.line_spacing = Pt(P["line"])
        heads = P.get("heads") or {}
        for lv, (size, bold) in {1: (14, True), 2: (14, False), 3: (12, False), 4: (11, False), 5: (11, True)}.items():
            s = st[f"Heading {lv}"]
            f, size, bold = self._role(f"H{lv}", "Tahoma", size, bold)
            self._font(s, f, size, bold)
            self._para(s, before=self._space(f"H{lv}_before", 14 if lv <= 2 else 10),
                       after=self._space(f"H{lv}_after", 6), line=1.0, keep=True, align=AL.LEFT)
            s.paragraph_format.page_break_before = bool((heads.get(str(lv)) or {}).get("page_break"))

        def new(name, base="Normal"):
            s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = st[base]
            return s
        cap = P.get("caption") or {}
        f, size, bold = self._role("caption", "Tahoma", self._role("body")[1], True)
        s = new("Caption UTS"); self._font(s, f, size, bold)
        self._para(s, before=4, after=12, align=AL.CENTER if cap.get("center", True) else AL.LEFT)
        s = new("Figure UTS")
        self._para(s, before=6, after=2, keep=True, align=AL.CENTER if (P.get("figure") or {}).get("center", True) else AL.LEFT)
        f, size, bold = self._role("td", *self._role("body"))
        s = new("Table Text"); self._font(s, f, size, False); self._para(s, after=0, line=1.0, align=AL.LEFT)
        s.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        s = new("Table Head UTS", "Table Text"); s.paragraph_format.keep_with_next = True   # шапка не отрывается от тела
        s = new("Table Text Small", "Table Text"); s.font.size = Pt(9)
        s = new("Table Head Small", "Table Text Small"); s.paragraph_format.keep_with_next = True
        f, size, bold = self._role("code", "Consolas", 10)
        if not f.lower().startswith(("consolas", "courier")):
            f, size, bold = "Consolas", self._role("body")[1], False
        s = new("Code UTS"); self._font(s, f, size, bold); self._para(s, after=0, line=1.0, align=AL.LEFT)
        s.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        s = new("Code Center UTS", "Code UTS"); s.paragraph_format.alignment = AL.CENTER
        s = new("List UTS")                                       # межстрочный — от «Normal»
        s.paragraph_format.space_after = Pt(self._space("list", 3))
        s.paragraph_format.alignment = AL.JUSTIFY if align.get("list") == "both" else AL.LEFT
        s = new("Center UTS"); s.paragraph_format.alignment = AL.CENTER   # Ф13: абзац по центру в оригинале
        s = new("Title UTS"); self._font(s, "Verdana", 18, bold=True); self._para(s, before=24, after=24, align=AL.CENTER)
        s = new("Front UTS"); self._font(s, size=14); self._para(s, after=4)
        s = new("Small UTS"); self._font(s, size=8); self._para(s, after=0, line=1.0)
        s = new("TOC Title UTS"); self._font(s, size=14, bold=True); self._para(s, before=6, after=12)
        s = new("TOF Title UTS", "TOC Title UTS")
        s = new("Header UTS"); self._font(s, size=10); self._para(s, after=0, line=1.0, align=AL.CENTER)
        s = new("Stamp UTS"); self._font(s, "Times New Roman", 10); self._para(s, after=0, line=1.0, align=AL.LEFT)

    def _heading_numbering(self):
        """Многоуровневая нумерация заголовков (1, 1.1, 1.1.1 …), привязана к стилям Heading 1–5: в оригинале все 157
        номеров совпадают с автонумерацией по уровням — Word ведёт её сам и при правке перенумерует."""
        num = self.doc.part.numbering_part.element
        an = _el("w:abstractNum", w__abstractNumId=900)
        an.append(_el("w:multiLevelType", w__val="multilevel"))
        heads = self.P.get("heads") or {}
        for lv in range(5):
            l = _el("w:lvl", w__ilvl=lv)
            l.append(_el("w:start", w__val=1)); l.append(_el("w:numFmt", w__val="decimal"))
            l.append(_el("w:pStyle", w__val=f"Heading{lv + 1}"))
            l.append(_el("w:lvlText", w__val=".".join(f"%{i + 1}" for i in range(lv + 1))))
            l.append(_el("w:lvlJc", w__val="left"))
            # Ф4: текст заголовка — на расстоянии от номера, как в оригинале (UTS: 21,6 / 28,8 / 36 / 43,2 / 50,4 pt;
            # профиль не даёт номеру подойти к тексту ближе 4 pt)
            tab = (heads.get(str(lv + 1)) or {}).get("tab")
            tw = round(tab * 20) if tab else (851 if lv < 2 else 1134)
            ppr = _el("w:pPr"); ppr.append(_el("w:tabs")); ppr[-1].append(_el("w:tab", w__val="left", w__pos=tw))
            ppr.append(_el("w:ind", w__left=tw, w__hanging=tw)); l.append(ppr)
            an.append(l)
        first_num = num.find(W + "num")
        (first_num.addprevious if first_num is not None else num.append)(an)
        n = _el("w:num", w__numId=900); n.append(_el("w:abstractNumId", w__val=900)); num.append(n)
        for lv in range(1, 6):
            ppr = self.doc.styles[f"Heading {lv}"].element.get_or_add_pPr()
            numpr = _el("w:numPr"); numpr.append(_el("w:ilvl", w__val=lv - 1)); numpr.append(_el("w:numId", w__val=900))
            ppr.insert(0, numpr)

    def _list_ind(self, lv):
        """Ф7: отступы пункта уровня lv (маркер, текст) в pt — по профилю; уровни глубже замеренных — с тем же шагом."""
        lists = self.P.get("lists") or {}
        if not lists:
            return 18 * (2 * lv + 1), 36 * (lv + 1)
        known = sorted(int(k) for k in lists)
        top = known[-1]
        if lv <= top and str(lv) in lists:
            v = lists[str(lv)]
            return v["marker"], v["text"]
        v = lists[str(top)]
        stepx = (v["text"] - lists[str(known[0])]["text"]) / top if top else 36
        return v["marker"] + stepx * (lv - top), v["text"] + stepx * (lv - top)

    def _list_num(self, kind, start=1, level=0, suffix="."):
        """Новый экземпляр списка вида kind: нумерация — с номера первого пункта оригинала (start, на уровне level;
        список после рисунка продолжает «7.»), вид метки — как в оригинале («a.» / «a.)»), маркеры уровней • o ▪."""
        num = self.doc.part.numbering_part.element
        key = (kind, suffix or ".")
        if key not in self._abstract:
            aid = 901 + len(self._abstract)
            an = _el("w:abstractNum", w__abstractNumId=aid)
            an.append(_el("w:multiLevelType", w__val="hybridMultilevel"))
            for lv in range(4):
                l = _el("w:lvl", w__ilvl=lv)
                l.append(_el("w:start", w__val=1))
                fontn = None
                if kind in NUMFMT:
                    fmt = NUMFMT[kind] if lv == 0 else ["lowerLetter", "lowerRoman", "decimal"][(lv - 1) % 3]
                    l.append(_el("w:numFmt", w__val=fmt))
                    l.append(_el("w:lvlText", w__val=f"%{lv + 1}" + ((suffix or ".") if lv == 0 else ".")))
                else:
                    ch, fontn = BULLETS.get(kind if lv == 0 else ["Circle", "Square", "Disc"][(lv - 1) % 3],
                                            BULLETS["Disc"])
                    l.append(_el("w:numFmt", w__val="bullet")); l.append(_el("w:lvlText", w__val=ch))
                l.append(_el("w:lvlJc", w__val="left"))
                mark, text = self._list_ind(lv)
                ppr = _el("w:pPr"); ppr.append(_el("w:ind", w__left=round(text * 20), w__hanging=round((text - mark) * 20)))
                l.append(ppr)
                if fontn:
                    rpr = _el("w:rPr"); rpr.append(_el("w:rFonts", w__ascii=fontn, w__hAnsi=fontn, w__hint="default"))
                    l.append(rpr)
                an.append(l)
            first_num = num.find(W + "num")
            (first_num.addprevious if first_num is not None else num.append)(an)
            self._abstract[key] = aid
        nid = 1000 + len(self.nums)
        n = _el("w:num", w__numId=nid); n.append(_el("w:abstractNumId", w__val=self._abstract[key]))
        for lv in range(4):
            o = _el("w:lvlOverride", w__ilvl=lv)
            o.append(_el("w:startOverride", w__val=start if lv == min(level, 3) else 1)); n.append(o)
        num.append(n)
        self.nums[nid] = kind
        return nid

    # ---------------------------------------------------------------- прогоны
    def _img_run(self, par, img):
        from docx.shared import Pt
        data = self.image(img)
        if data:
            par.add_run().add_picture(io.BytesIO(data), width=Pt(float(img["x1"]) - float(img["x0"])))

    def _put_runs(self, par, runs, style_bold=False):
        """Текст с переводом. Одно начертание — перевод целиком; несколько (жирное «FIGYELEM!» + текст) — перевод
        с метками <fN>, каждой группе своё начертание. Картинки в начале — перед текстом, иначе — после."""
        from docx.shared import RGBColor
        texts = [r for r in runs if "t" in r]
        imgs = [r for r in runs if "img" in r and not r.get("paren")]
        pimgs = [r for r in runs if "img" in r and r.get("paren")]      # значки в «()» — на то же место в переводе
        # значок в начале — первым после пустых прогонов (табуляция после снятого маркера пункта)
        head = next((r for r in runs if "img" in r or r["t"].strip()), None)
        lead_img = head is not None and "img" in head and not head.get("paren")
        groups = []
        for r in texts:
            key = (bool(r.get("b")), bool(r.get("i")), (r.get("font") or "").lower().startswith(("consolas", "courier")),
                   r.get("color") if r.get("color") not in ("000000", None) else None)
            if groups and (groups[-1][0] == key or not r["t"].strip()):
                groups[-1][1].append(r)
            else:
                groups.append([key, [r]])
        # группа из одних пробелов (пробел перед текстом после снятого маркера) — в следующую: иначе пустая метка
        # <f1></f1> и лишний запрос к порталу
        merged = []
        for g in groups:
            if merged and not "".join(x["t"] for x in merged[-1][1]).strip():
                merged[-1] = [g[0], merged[-1][1] + g[1]]
            else:
                merged.append(g)
        groups = merged
        parts = ["".join(x["t"] for x in g[1]) for g in groups]
        out = []
        if len(groups) <= 1:
            src = "".join(parts)
            ru = _keep_tab(src, self.tr(src) if src.strip() else src)
            out = [(groups[0][0] if groups else None, ru if ru is not None else src)]
        else:
            tagged = "".join(f"<f{i + 1}>{t}</f{i + 1}>" for i, t in enumerate(parts))
            code = {i + 1 for i, g in enumerate(groups) if g[0][2]}          # моноширинные группы — code_part
            if code:
                self.code_groups[tagged] = code
            get = getattr(self.tr, "get", None)
            plain0 = get("".join(parts)) if get else None
            if get and plain0 is not None and get(tagged) is None and _verbatim(plain0, groups, parts):
                ru = None       # есть перевод без меток, прочие начертания — у дословных кусков: портал не нужен
            else:
                ru = self.tr(tagged)
            got = re.findall(r"<f(\d+)>(.*?)</f\1>", ru or "", re.S)
            if ru and sorted(int(k) for k, _ in got) == list(range(1, len(groups) + 1)):
                d = {int(k): t for k, t in got}
                d.update({k: code_part(parts[k - 1], d[k]) for k in code})
                out = []
                for i, g in enumerate(groups):
                    t = d[i + 1]
                    if out and out[-1][1] and not out[-1][1].endswith((" ", "\t")) and t and not t.startswith((" ", ".", ",", ":", ";", ")")):
                        t = " " + t
                    out.append((g[0], t))
            else:
                plain = self.tr("".join(parts))
                out = _restyle(plain if plain is not None else "".join(parts), groups, parts)
        if lead_img:
            for im in imgs:
                self._img_run(par, im["img"])
            if out and out[0][1]:
                out[0] = (out[0][0], " " + out[0][1].lstrip())
        def text_run(t, key):
            run = par.add_run(t)
            if key:
                b, i, code, color = key
                if b and not style_bold:
                    run.bold = True
                if i:
                    run.italic = True
                if code and par.style.name != "Code UTS":
                    run.font.name = "Consolas"
                if color and color != "ffffff":
                    run.font.color.rgb = RGBColor.from_string(color.upper())
        for key, t in out:
            while t and pimgs and re.search(r"\(\s*\)", t):
                m = re.search(r"\(\s*\)", t)
                text_run(t[:m.start() + 1], key)
                self._img_run(par, pimgs.pop(0)["img"])
                t = ")" + t[m.end():]
            if t:
                text_run(t, key)
        for im in ([] if lead_img else imgs) + pimgs:
            self._img_run(par, im["img"])
        # пункт оглавления оригинала без номера раздела (шапка таблицы) — уровень структуры, чтобы его видело поле TOC
        lv = getattr(self, "toc_extra", {}).get(KEY(text_of(runs)))
        if lv:
            ppr = par._p.get_or_add_pPr()
            after = next((c for c in ppr if c.tag in {W + n for n in ("divId", "cnfStyle", "rPr", "sectPr", "pPrChange")}), None)
            el = _el("w:outlineLvl", w__val=lv - 1)            # порядок элементов pPr — по схеме, иначе Word ругается
            after.addprevious(el) if after is not None else ppr.append(el)
        # табуляция — позиции по оригиналу (пары «подпись ⇥ значение» на титуле, колонки в абзаце)
        from docx.shared import Pt
        for x in sorted({round(x) for r in runs for x in r.get("tabs", [])}):
            if x - self.left > 6:
                par.paragraph_format.tab_stops.add_tab_stop(Pt(x - self.left))
        if self.L and any(r.get("rtab") for r in runs):
            # номер листа в конце строки оглавления — к правому полю: русское название длиннее, и номер
            # по позиции оригинала перескакивал на следующую строку (3SZ19, 16.10)
            from docx.enum.text import WD_TAB_ALIGNMENT
            par.paragraph_format.tab_stops.add_tab_stop(Pt(self.L["W"] - self.L["right"] - self.left), WD_TAB_ALIGNMENT.RIGHT)

    # ---------------------------------------------------------------- блоки
    def paragraph(self, b, style="Normal", tgt=None):
        runs = b["runs"]
        tgt = tgt if tgt is not None else self.doc
        text = text_of(runs).strip()
        m = CAPTION.match(text)
        if m and style == "Normal":
            return self.caption(m, tgt)
        if all("img" in r for r in runs):
            par = tgt.add_paragraph(style="Figure UTS" if style == "Normal" else style)
            for r in runs:
                self._img_run(par, r["img"])
            return par
        if style == "Normal" and is_code(runs):
            style = "Code UTS"
        size = (b.get("geom") or {}).get("size")
        if style == "Normal" and size:
            if size >= 17:
                style = "Title UTS"
            elif size >= 13:
                style = "Front UTS"
            elif size <= 8:
                style = "Small UTS"
        if style in ("Normal", "Code UTS") and tgt is self.doc and self.L and self._centered(b):
            style = "Center UTS" if style == "Normal" else "Code Center UTS"
        if style == "Normal" and b.get("indent"):
            style = self._indented(b["indent"])              # отступ оригинала — уровнем стиля (docx_layout)
        par = tgt.add_paragraph(style=style)
        self._put_runs(par, runs, style_bold=style in ("Title UTS",))
        return par

    def _indented(self, ind):
        """Стиль абзаца с отступом слева ind pt (шаг 9 pt): отступ — стилем, а не прямым оформлением абзаца."""
        from docx.enum.style import WD_STYLE_TYPE
        from docx.shared import Pt
        name = f"Indent {int(ind)} UTS"
        if name not in [s.name for s in self.doc.styles]:
            s = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = self.doc.styles["Normal"]
            s.paragraph_format.left_indent = Pt(ind)
        return name

    def caption(self, m, tgt):
        """Подпись в одном формате: «Рисунок N. Название» / «Таблица N. Название», номер — полем SEQ
        (список рисунков строится по нему). Переводится только название."""
        n, kind, title = int(m.group(1)), m.group(2).lower(), m.group(3).strip()
        label = "Рисунок" if kind == "ábra" else "Таблица"
        par = tgt.add_paragraph(style="Caption UTS")
        par.add_run(label + " ")
        _field(par, f"SEQ {label} \\* ARABIC", str(n))
        ru = self.tr(m.group(0).strip())
        cut = CAP_RU.sub("", ru, count=1).strip() if ru else ""
        if not cut or cut == (ru or "").strip():
            cut = self.tr(title) or title
        cut = cut.lstrip("–—-:. ")                           # «9. táblázat – …» → «Таблица 9. Название», без тире
        par.add_run(". " + cut.rstrip("."))
        return par

    def _centered(self, b):
        from docx_profile import centered
        g = b.get("geom") or {}
        if (g.get("page") or 0) == 1 and g.get("x0") is not None:
            # титул: центр — по листу (у титула свои поля); «Érvényes … visszavonásig» ÜFK стоял влево (16.10)
            return abs((g["x0"] + g["x1"]) / 2 - self.L["W"] / 2) < 15 and g["x1"] - g["x0"] < 0.7 * self.L["W"]
        return (g.get("page") or 0) >= 2 and centered(g, self.L)

    def heading(self, b):
        full = text_of(b["runs"]).strip()
        get = getattr(self.tr, "get", None)
        ru = get(full) if get else None                      # в памяти PDF заголовок — с номером
        auto = getattr(self, "autonum", True)
        if ru:
            par = self.doc.add_paragraph(style=f"Heading {min(b['level'], 5)}")
            par.add_run(re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", ru) if auto else ru)
            return
        runs = [dict(r) for r in b["runs"] if "t" in r]
        if not auto:                                         # номер оригинала — текстом, его не даёт нумерация
            for r in runs:
                r["b"] = False
            par = self.doc.add_paragraph(style=f"Heading {min(b['level'], 5)}")
            self._put_runs(par, runs)
            return
        while runs and re.fullmatch(r"\s*\d+(?:\.\d+)*\s*", runs[0]["t"]):
            runs = runs[1:]                                  # номер отдельным прогоном — его даёт нумерация стиля
        if runs:
            runs[0]["t"] = re.sub(r"^\s*\d+(?:\.\d+)*\s+", "", runs[0]["t"]).lstrip()
        for r in runs:
            r["b"] = False                                   # начертание — от стиля
        par = self.doc.add_paragraph(style=f"Heading {min(b['level'], 5)}")
        self._put_runs(par, runs)

    def list_(self, b, level=0, tgt=None):
        from docx.shared import Pt
        nid = self._list_num(b["kind"], b.get("start", 1), level, b.get("suffix") or ".")
        tgt = tgt if tgt is not None else self.doc
        for item in b["items"]:
            first = True
            for x in item:
                if x["k"] == "list":
                    self.list_(x, level + 1, tgt); continue
                if x["k"] == "table":
                    self.table(x); continue
                if x["k"] != "p" or not x["runs"]:
                    continue
                runs = [dict(r) for r in x["runs"]]
                if runs and "t" in runs[0] and LEAD.match(runs[0]["t"]):
                    runs[0]["t"] = LEAD.sub("", runs[0]["t"])
                    if not runs[0]["t"].strip():          # маркер отдельным прогоном (Symbol) — прочь вместе с ним
                        runs = runs[1:]
                        if runs and "t" in runs[0]:
                            runs[0]["t"] = runs[0]["t"].lstrip()
                if all("img" in r for r in runs) and not first:
                    self.paragraph(dict(x, runs=runs), tgt=tgt); continue
                par = tgt.add_paragraph(style="List UTS")
                if first:
                    numpr = _el("w:numPr"); numpr.append(_el("w:ilvl", w__val=min(level, 3)))
                    numpr.append(_el("w:numId", w__val=nid))
                    par._p.get_or_add_pPr().insert(0, numpr)
                else:
                    par.paragraph_format.left_indent = Pt(self._list_ind(level)[1])
                self._put_runs(par, runs)
                first = False

    def _borders(self, t, color):
        """Линии таблицы цветом оригинала (служебные таблицы UTS — серые 808080, остальные — чёрные)."""
        b = _el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(_el(f"w:{side}", w__val="single", w__sz=4, w__space=0, w__color=color))
        t._tbl.tblPr.append(b)

    def _cell_margins(self, t, fr):
        """Ф11: поля ячеек по оригиналу: сверху и снизу — по таблице (до текста за вычетом межстрочного зазора шрифта),
        слева и справа — обычные для документа (у таблицы с числами вправо левый зазор шире поля, даты переносились)."""
        tp = self.P.get("table") or {}
        m = _el("w:tblCellMar")
        for side, key, cut in (("top", "pad_top", 2.0), ("left", "pad_left", 0), ("bottom", "pad_bottom", 2.0),
                               ("right", "pad_left", 0)):
            own = fr.get(key) if side in ("top", "bottom") else None
            v = own if own is not None and own < 12 else tp.get(key)
            if v is None:
                continue
            m.append(_el(f"w:{side}", w__w=max(0, round((v - cut) * 20)), w__type="dxa"))
        t._tbl.tblPr.append(m)

    def table(self, b):
        from docx.shared import Pt
        rows = [r for r in b["rows"] if r]
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        t = self.doc.add_table(rows=len(rows), cols=ncol)
        fr = b.get("frame") or {}
        if b.get("borders", True):
            t.style = self.doc.styles["Table Grid"]
            if fr.get("line") and fr["line"] != "000000":
                self._borders(t, fr["line"])
            self._cell_margins(t, fr)
        else:
            self._no_borders(t)
        sizes = Counter(round(p["geom"]["size"]) for r in rows for c in r for p in c["blocks"] if p.get("geom"))
        small = bool(sizes) and sizes.most_common(1)[0][0] <= 9
        cell_style = "Table Text Small" if small else "Table Text"
        head_style = "Table Head Small" if small else "Table Head UTS"
        # Ф10: ширины столбцов — по рамке оригинала; без рамки — по левым краям ячеек полных строк
        xs = fr.get("xs") or []
        widths = None
        bounds = xs if len(xs) == ncol + 1 else []           # границы столбцов — для объединённых ячеек
        if len(xs) == ncol + 1:
            widths = [xs[i + 1] - xs[i] for i in range(ncol)]
        else:
            full = [r for r in rows if len(r) == ncol]
            lefts = [min([r[c]["geom"]["x0"] for r in full if r[c].get("geom")] or [None], key=lambda v: v or 0)
                     for c in range(ncol)]
            if all(v is not None for v in lefts) and len(set(round(v) for v in lefts)) == ncol:
                # рамка неполная (ÜFK: линии только между группами шапки) — границы по левым краям текста полных строк;
                # правый край — по самому правому тексту таблицы, шапки тоже (у столбца «8 óra» шапка
                # «Ellenőrzési ciklusidő» шире — по полным строкам столбец выходил в 27 pt, шапка рассыпалась)
                right = max(c["geom"]["x1"] for r in rows for c in r if c.get("geom")) + 6
                bounds = lefts + [right]
                rights = lefts[1:] + [right]
                if not b.get("borders", True) and self.L:
                    # раскладка без рамок (значок ⇥ подпись): последний столбец — до правого поля, а не по длине
                    # венгерского текста (русский длиннее и переносился в узком столбце, UTS стр. 93); отступ — как в оригинале
                    rights[-1] = max(rights[-1], self.L["W"] - self.L["right"])
                    ind = lefts[0] - self.left - 5
                    if ind > 3:
                        tp = t._tbl.tblPr
                        nxt = next((c for c in tp if c.tag in {W + n for n in (
                            "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption")}), None)
                        el = _el("w:tblInd", w__w=round(ind * 20), w__type="dxa")   # порядок tblPr — по схеме
                        nxt.addprevious(el) if nxt is not None else tp.append(el)
                widths = [max(rights[c] - lefts[c], 20) for c in range(ncol)]
        if widths:
            self._fit_words(rows, widths, 9 if small else 10)
            t.autofit = False
            for c in range(ncol):
                for cell in t.columns[c].cells:
                    cell.width = Pt(widths[c])
            grid = t._tbl.tblGrid
            for gc, w in zip(grid.findall(W + "gridCol"), widths):
                gc.set(W + "w", str(round(w * 20)))
        aligns = _col_align(rows, xs, (b.get("geom") or {}).get("page"))
        in_head = True
        for ri, row in enumerate(rows):
            starts = _span_starts(row, bounds, ncol) if len(row) < ncol else list(range(len(row)))
            for ci, a in enumerate(starts):                  # объединённые ячейки строки — по столбцам рамки
                z = starts[ci + 1] - 1 if ci + 1 < len(starts) else ncol - 1
                if z > a:
                    t.cell(ri, a).merge(t.cell(ri, z))
            in_head = in_head and all(c["th"] for c in row)
            head = in_head and ri < 3 and len(rows) > 1
            if head and len(rows) > 3:
                t.rows[ri]._tr.get_or_add_trPr().append(_el("w:tblHeader"))
            # короткая строка не рвётся между листами (в оригинале строки целые; длинная — пусть рвётся, иначе не влезет)
            size = sum(len(text_of(x["runs"])) for c in row for x in c["blocks"] if x["k"] == "p")
            if size < 600 and not any("img" in r for c in row for x in c["blocks"] if x["k"] == "p" for r in x["runs"]):
                t.rows[ri]._tr.get_or_add_trPr().append(_el("w:cantSplit"))
            for ci, cell in enumerate(row):
                dc = t.cell(ri, starts[ci])
                first = dc.paragraphs[0]
                used = False
                style = head_style if head else cell_style
                al = aligns.get(("head" if head else "body", ci)) if len(row) == ncol else None
                if al:
                    style = self._aligned(style, al)
                for x in cell["blocks"]:
                    if x["k"] == "list":
                        self.list_(x, 0, dc); continue
                    if x["k"] != "p" or not x["runs"]:
                        continue
                    if not used:
                        first.style = self.doc.styles[style]
                        if all("img" in r for r in x["runs"]):
                            for r in x["runs"]:
                                self._img_run(first, r["img"])
                        else:
                            self._put_runs(first, x["runs"])
                        used = True
                    else:
                        self.paragraph(x, style, tgt=dc)
                if not used and len(dc.paragraphs) > 1:      # ячейка начинается со списка — пустой первый абзац прочь
                    first._p.getparent().remove(first._p)
                elif not used:                               # пустая ячейка — стилем таблицы: «Обычный» (8 pt после,
                    first.style = self.doc.styles[style]     # 1,15) раздувал строку (история изменений UTS, 16.10)
                if head and fr.get("fill") and b.get("borders", True):
                    dc._tc.get_or_add_tcPr().append(_el("w:shd", w__val="clear", w__color="auto", w__fill=fr["fill"]))
        self.doc.add_paragraph(style="Table Text")        # Word: после таблицы — абзац

    def _fit_words(self, rows, widths, size):
        """Столбец не уже самого длинного слова перевода с полями ячейки: «Редакция» в узком столбце «Törés / verzió»
        рвалась «Редакци / я» (UTS), метка «A.1.» — «A. / 1.» (ÜFK, 16.10). Недостающее — у столбца с самым большим
        запасом; ширина таблицы та же."""
        get = getattr(self.tr, "get", None)
        if get is None:
            return
        pad = 2 * (self.P.get("table") or {}).get("pad_left", 5.4) + 5     # поля ячейки слева и справа (_cell_margins)
        ncol = len(widths)
        need = [0.0] * ncol
        for row in rows:
            if len(row) != ncol:
                continue
            for ci, cell in enumerate(row):
                for x in cell["blocks"]:
                    if x["k"] != "p" or not x["runs"]:
                        continue
                    src = text_of(x["runs"]).strip()
                    ru = TAGS.sub("", (get(src) if src else None) or src)
                    bold = any(r.get("b") for r in x["runs"] if r.get("t", "").strip())
                    w = max((_text_w(wd, size) for wd in ru.split()), default=0) * (1.1 if bold else 1.0)
                    need[ci] = max(need[ci], w + pad)
        for c in range(ncol):
            short = need[c] - widths[c]
            if short <= 0:
                continue
            d = max((k for k in range(ncol) if k != c), key=lambda k: widths[k] - need[k], default=None)
            give = min(short, widths[d] - need[d]) if d is not None else 0
            if give > 0:
                widths[c] += give
                widths[d] -= give

    def _title_styles(self, blocks):
        """Стили крупного, «переднего» и мелкого текста (титул) — шрифтом, кеглем и начертанием оригинала: было
        зашито Verdana 18 полужирный под UTS, а титул ÜFK набран Times New Roman (16.10)."""
        pick = {"Title UTS": lambda s: s >= 17, "Front UTS": lambda s: 13 <= s < 17, "Small UTS": lambda s: s <= 8}
        for name, ok in pick.items():
            runs = [r for b in _paras(blocks) if b["k"] == "p" and (b.get("geom") or {}).get("size")
                    and ok(b["geom"]["size"]) for r in b["runs"] if r.get("t", "").strip()]
            if runs:
                font, size, bold = Counter((r.get("font"), r.get("size"), bool(r.get("b"))) for r in runs).most_common(1)[0][0]
                self._font(self.doc.styles[name], font or "Tahoma", size or 10, bold)

    def _aligned(self, style, al):
        """Стиль ячейки с выравниванием столбца оригинала (стилем, а не прямым оформлением абзаца)."""
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        name = f"{style} {al}"
        if name not in [s.name for s in self.doc.styles]:
            s = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = self.doc.styles[style]
            s.paragraph_format.alignment = AL.CENTER if al == "center" else AL.RIGHT
        return name

    def _toc_styles(self, blocks):
        """Оглавления оригинала (docx_struct.toc_info): заголовки — свои стили по оригиналу (кегль, цвет, центр),
        строки уровней — стили Word «toc N» (полужирные — как в оригинале), пункты без номера раздела — абзацы, которым
        сборка даст уровень структуры (иначе поле TOC их не видит)."""
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        from docx.oxml.ns import qn
        from docx.shared import RGBColor
        self.toc_extra = {}
        for b in blocks:
            info = (b.get("info") or {}) if b["k"] == "toc" else None
            if info is None:
                continue
            for text, lv in info.get("extra", []):
                self.toc_extra[KEY(text)] = lv
            t = info.get("title")
            name = "TOF Title UTS" if info.get("figures") else "TOC Title UTS"
            if t:
                s = self.doc.styles[name]
                self._font(s, t.get("font") or "Tahoma", t.get("size") or 14, bool(t.get("b")))
                if (t.get("color") or "000000") != "000000":
                    s.font.color.rgb = RGBColor.from_string(t["color"].upper())
                s.paragraph_format.alignment = AL.CENTER if t.get("center") else AL.LEFT
            if info.get("figures"):
                continue
            from docx.shared import Pt
            for i, bold in enumerate(info.get("bold", [])):
                sid = f"TOC{i + 1}"
                s = next((x for x in self.doc.styles if x.style_id == sid), None)
                if s is None:
                    s = self.doc.styles.add_style(f"toc {i + 1}", WD_STYLE_TYPE.PARAGRAPH)
                    s.element.set(qn("w:styleId"), sid)
                    # встроенный стиль Word «toc N», а не пользовательский: иначе Word собирает оглавление своими
                    # «Оглавление N» и наши полужирный, отступы и шаг строк не видит
                    s.element.attrib.pop(qn("w:customStyle"), None)
                    s.base_style = self.doc.styles["Normal"]
                pf = s.paragraph_format
                pf.alignment = AL.LEFT                  # от «Обычного» по ширине — строка разреживалась по буквам
                pf.space_before = pf.space_after = Pt(0)
                ind = info.get("indent") or []
                pf.left_indent = Pt(ind[i] if i < len(ind) else 0)
                if info.get("pitch"):
                    pf.line_spacing = Pt(info["pitch"])     # шаг строк — как в оглавлении оригинала
                s.font.bold = bold

    def toc(self, b=None):
        info = (b or {}).get("info") or {}
        figs = info.get("figures", self._toc_n > 0)
        title = (info.get("title") or {}).get("t")
        text = (self.tr(title) if title else None) or ("Список рисунков" if figs else "Содержание")
        self.doc.add_paragraph(text, style="TOF Title UTS" if figs else "TOC Title UTS")
        par = self.doc.add_paragraph()
        if figs:
            _field(par, 'TOC \\h \\z \\c "Рисунок"', "[Список рисунков: в Word — правой кнопкой → «Обновить поле»]")
        else:                                       # уровней — сколько в оглавлении оригинала
            _field(par, f'TOC \\o "1-{info.get("levels") or 3}" \\h \\z \\u',
                   "[Оглавление: в Word — правой кнопкой → «Обновить поле»]")
        self._toc_n += 1

    def block(self, b):
        k = b["k"]
        if k == "h":
            self.heading(b)
        elif k == "p":
            self.paragraph(b)
        elif k == "list":
            self.list_(b)
        elif k == "table":
            self.table(b)
        elif k == "toc":
            self.toc(b)

    # ---------------------------------------------------------------- лист, колонтитулы, титул
    def _page(self, sec, L, left, right, top, bottom):
        from docx.shared import Pt
        sec.page_width, sec.page_height = Pt(L["W"]), Pt(L["H"])
        sec.left_margin, sec.right_margin = Pt(left), Pt(right)
        sec.top_margin, sec.bottom_margin = Pt(top), Pt(bottom)
        sec.header_distance, sec.footer_distance = Pt(L["header"]), Pt(L["footer"])

    def _no_borders(self, table):
        tblpr = table._tbl.tblPr
        b = _el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(_el(f"w:{side}", w__val="nil"))
        tblpr.append(b)

    def _zone_table(self, part, grid, style, page_tok=None):
        """Таблица колонтитула по сетке оригинала (docx_struct._zone_grid): ширины и выравнивание столбцов, рамка —
        все линии (штамп UTS) / только внешняя (строка-рамка ÜFK) / без рамки; номер листа — полем PAGE."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        rows, widths, aligns = grid["rows"], grid.get("widths") or [], grid.get("aligns") or []
        ncol = max(len(r["cells"]) for r in rows)
        t = part.add_table(rows=len(rows), cols=ncol, width=Pt(sum(widths) or 400))
        if grid.get("inner", True):
            t.style = self.doc.styles["Table Grid"]
        elif grid.get("box"):
            b = _el("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                b.append(_el(f"w:{side}", w__val="single" if side in ("top", "left", "bottom", "right") else "nil",
                             w__sz=4, w__space=0, w__color="000000"))
            t._tbl.tblPr.append(b)
        else:
            self._no_borders(t)
        for c, w in zip(t.columns, widths):
            for cell in c.cells:
                cell.width = Pt(w)
        if widths:                             # и сетка таблицы: LibreOffice берёт ширины из неё, а не из ячеек
            t.autofit = False
            for gc, w in zip(t._tbl.tblGrid.findall(W + "gridCol"), widths):
                gc.set(W + "w", str(round(w * 20)))
        al_of = {"right": AL.RIGHT, "center": AL.CENTER, "left": AL.LEFT}   # «влево» — явно: стиль шапки по центру
        for ri, row in enumerate(rows):
            for ci, txt in enumerate(row["cells"]):
                p = t.cell(ri, ci).paragraphs[0]
                p.style = self.doc.styles[style]
                if ci < len(aligns) and aligns[ci] in al_of:
                    p.alignment = al_of[aligns[ci]]
                self._zone_text(p, txt, row["bold"], page_tok)
        return t

    def _zone_text(self, p, txt, bold, page_tok):
        """Текст ячейки колонтитула: «N/NN» → PAGE/NUMPAGES (штамп UTS), число листа из колонтитула оригинала
        («85. oldal», docx_struct._typical_page) → PAGE в переводе на том же месте; прочее — перевод."""
        m = re.match(r"^(.*?)(\d+)\s*/\s*(\d+)\s*$", txt)
        if m:
            head = (m.group(1).strip() + "    ") if m.group(1).strip() else ""
            r = p.add_run(head); r.bold = bold
            f = _field(p, "PAGE", m.group(2)); f.bold = bold
            r = p.add_run("/"); r.bold = bold
            f = _field(p, "NUMPAGES", m.group(3)); f.bold = bold
            p.alignment = 2
            return
        tok = (page_tok or "").rstrip(".")
        num = re.compile(rf"(?<!\d){re.escape(tok)}(?!\d)") if tok else None
        ru = (self.tr(txt) or txt) if txt else ""
        mm = num.search(ru) if num and num.search(txt) else None
        if mm:
            r = p.add_run(ru[:mm.start()]); r.bold = bold
            f = _field(p, "PAGE", tok); f.bold = bold
            r = p.add_run(ru[mm.end():]); r.bold = bold
        else:
            r = p.add_run(ru); r.bold = bold

    def _header(self, sec, L, logos_only=False):
        """Шапка: логотипы по краям, строки по центру (название и номер документа), линия снизу; шапка без логотипов
        (строка-таблица ÜFK) — таблицей по сетке оригинала."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        hdr = sec.header
        hdr.is_linked_to_previous = True                     # новый раздел копирует ссылку на шапку титула —
        hdr.is_linked_to_previous = False                    # своя часть, иначе обе шапки пишутся в одну
        grid = L.get("head_grid")
        if grid and grid.get("rows"):
            if not logos_only:
                t = self._zone_table(hdr, grid, "Header UTS", L.get("page_tok"))
                hdr.paragraphs[0]._p.addprevious(t._tbl)
                hdr.paragraphs[0].paragraph_format.space_after = Pt(0)
            return
        width = L["W"] - L["left"] - L["right"]
        t = hdr.add_table(rows=1, cols=3, width=Pt(width))
        hdr.paragraphs[0]._p.addprevious(t._tbl)             # таблица — первой: пустой абзац под ней несёт линию
        self._no_borders(t)
        mid_x = L["W"] / 2
        lefts = [i for i in L["head_imgs"] if (i["x0"] + i["x1"]) / 2 < mid_x]
        rights = [i for i in L["head_imgs"] if (i["x0"] + i["x1"]) / 2 >= mid_x]
        wl = max([i["x1"] - i["x0"] for i in lefts] + [60]) + 6
        wr = max([i["x1"] - i["x0"] for i in rights] + [60]) + 6
        for c, w in zip(t.columns, (wl, width - wl - wr, wr)):
            for cell in c.cells:
                cell.width = Pt(w)
        for cell, imgs, al in ((t.cell(0, 0), lefts, AL.LEFT), (t.cell(0, 2), rights, AL.RIGHT)):
            p = cell.paragraphs[0]; p.alignment = al
            for im in imgs:
                data = self.image({"stream": type("S", (), {"objid": im["xref"]})(), **im})
                if data:
                    p.add_run().add_picture(io.BytesIO(data), width=Pt(im["x1"] - im["x0"]))
        mid = t.cell(0, 1)
        if not logos_only:
            for i, ln in enumerate(L["head_lines"]):
                p = mid.paragraphs[0] if i == 0 else mid.add_paragraph()
                p.style = self.doc.styles["Header UTS"]
                if i:
                    p.paragraph_format.space_before = Pt(8)
                p.add_run(self.tr(ln["t"]) or ln["t"])
        for p in ([] if logos_only else list(hdr.paragraphs)):   # пустой абзац колонтитула — под линию
            ppr = p._p.get_or_add_pPr()
            bdr = _el("w:pBdr"); bdr.append(_el("w:bottom", w__val="single", w__sz=6, w__space=1, w__color="000000"))
            ppr.append(bdr)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 0.3

    def _footer(self, sec, L):
        """Штамп листа: сетка, выравнивание и рамка по оригиналу (_zone_table), номер листа — полем."""
        from docx.shared import Pt
        ftr = sec.footer
        ftr.is_linked_to_previous = True
        ftr.is_linked_to_previous = False
        if not L["stamp"]:
            return
        self._zone_table(ftr, {"rows": L["stamp"], "widths": L["stamp_widths"], "aligns": L.get("stamp_aligns"),
                               "inner": L.get("stamp_inner", True), "box": L.get("stamp_box", True)},
                         "Stamp UTS", L.get("page_tok"))
        for p in list(ftr.paragraphs):
            p.paragraph_format.space_after = Pt(0)

    def build(self, blocks, out, L):
        """L — раскладка листа (docx_struct.layout). Титул (стр. 1) — свой раздел без колонтитулов,
        с полями титула; тело — поля, шапка и штамп оригинала."""
        from docx.enum.section import WD_SECTION
        from docx.shared import Pt
        if (self.P.get("page") or {}).get("right"):
            L = dict(L, right=self.P["page"]["right"])      # правое поле — по краю выключки текста, не по таблицам
        self.L = L
        self._toc_styles(blocks)
        self._title_styles(blocks)
        self.autonum = _autonum_ok(blocks)
        if not self.autonum:                    # номера заголовков — текстом, как в оригинале: стили без нумерации
            for lv in range(1, 6):
                ppr = self.doc.styles[f"Heading {lv}"].element.pPr
                for np_ in ([] if ppr is None else ppr.findall(W + "numPr")):
                    ppr.remove(np_)
                if ppr is not None:
                    for ind in ppr.findall(W + "ind"):
                        ppr.remove(ind)
        # шрифт штампа и шапки-таблицы — как в колонтитулах оригинала (было зашито под UTS: Times New Roman 10, Tahoma 10)
        if L.get("stamp_font"):
            self._font(self.doc.styles["Stamp UTS"], L["stamp_font"], L.get("stamp_size") or 10)
        hg = L.get("head_grid") or {}
        if hg.get("font"):
            self._font(self.doc.styles["Header UTS"], hg["font"], hg.get("size") or 10)
        first = self.doc.sections[0]
        prev_bottom = 40                        # от верхнего поля титула: название ÜFK стоит с середины листа
        # первый лист с колонтитулами тела (3SZ19 начинается оглавлением с шапкой) — не титул (docx_layout)
        in_title = had_title = not L.get("first_is_body")
        if in_title:
            self._page(first, L, L["title_left"], 20, 40, 30)
            self.left = L["title_left"]
        else:
            self._page(first, L, L["left"], L["right"], L["top"], L["bottom"])
            self.left = L["left"]
            self._header(first, L)
            self._footer(first, L)
        title_foot = []
        for b in merge_inline(blocks):
            g = b.get("geom") or {}
            page = g.get("page")
            if in_title and page is not None and page >= 2:
                in_title = False
                sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
                self._page(sec, L, L["left"], L["right"], L["top"], L["bottom"])
                self.left = L["left"]
                self._header(sec, L)
                self._footer(sec, L)
            if in_title and b["k"] == "p" and _figure(b) is False and all("img" in r for r in b["runs"]):
                prev_bottom = max(prev_bottom, g.get("bottom") or 0)   # отступ текста титула — от низа логотипов
                continue                                     # логотипы титула — в шапку титула
            if in_title and b["k"] == "p" and all("img" in r for r in b["runs"]):
                if g and g.get("top", 999) < 110:
                    prev_bottom = max(prev_bottom, g.get("bottom") or 0)
                    continue
            if in_title and b["k"] == "p" and g.get("top", 0) > 0.8 * L["H"]:
                title_foot.append(b)                         # юридический текст внизу титула — в подвал титула
                continue
            self.block(b)
            if in_title and b["k"] == "p" and g and self.doc.paragraphs:
                par = self.doc.paragraphs[-1]
                if par.alignment == 1 or (par.style.paragraph_format.alignment or 0) == 1:
                    # по центру — центр там же, где в оригинале (поля титула несимметричны)
                    a = 2 * (g["x0"] + g["x1"]) / 2 - L["title_left"] - (L["W"] - 20)
                    if a > 3:
                        par.paragraph_format.left_indent = Pt(a)
                    elif a < -3:
                        par.paragraph_format.right_indent = Pt(-a)
                else:
                    dx = g["x0"] - L["title_left"]
                    if abs(dx) > 3:
                        par.paragraph_format.left_indent = Pt(dx)
                # большие пустоты титула (название посреди листа, юридический текст внизу) — как в оригинале
                if prev_bottom is not None and g.get("top") and g["top"] - prev_bottom > 24:
                    par.paragraph_format.space_before = Pt(0.85 * (g["top"] - prev_bottom))
                prev_bottom = g.get("bottom") or prev_bottom
        # после add_section объект первого раздела указывает на последний (sectPr тела документа) — шапку титула
        # ставить разделу [0], иначе шапки менялись местами
        if had_title:
            self._header(self.doc.sections[0], L, logos_only=True)
        if title_foot:
            foot = self.doc.sections[0].footer
            foot.is_linked_to_previous = False
            first_p = foot.paragraphs[0] if foot.paragraphs and not foot.paragraphs[0].text else None
            for b in title_foot:
                self.paragraph(b, tgt=foot)
            if first_p is not None:
                first_p._p.getparent().remove(first_p._p)
        self._tab_groups()
        self.doc.settings.element.append(_el("w:updateFields", w__val="true"))
        self.doc.save(out)

    def _tab_groups(self):
        """Пары «подпись ⇥ значение» подряд с одной табуляцией (титул: «Műszaki változás száma⇥10776»): позиция —
        общая и не ближе самой длинной русской подписи, иначе значение длинной подписи уезжало на следующую
        табуляцию («Номер технического изменения 10776» не в столбце, UTS 16.10). Ширина — латиница по метрике
        Helvetica, кириллица и прочее — средней шириной Tahoma (у Helvetica PDF нет кириллицы, шрифтов Word на
        сервере нет), с запасом."""
        from docx.shared import Pt
        normal = (self.doc.styles["Normal"].font.size or Pt(10)).pt
        pars = self.doc.paragraphs

        def stop(p):
            ts = p.paragraph_format.tab_stops
            return ts[0].position if len(ts) == 1 and "\t" in p.text else None

        i = 0
        while i < len(pars):
            pos = stop(pars[i])
            if pos is None:
                i += 1
                continue
            j = i
            while j < len(pars) and stop(pars[j]) == pos:
                j += 1
            need = 0.0
            for p in pars[i:j]:
                size = (p.style.font.size.pt if p.style.font.size else normal)
                ind = p.paragraph_format.left_indent.pt if p.paragraph_format.left_indent else 0.0
                need = max(need, ind + 1.1 * _text_w(p.text.split("\t")[0], size) + 8)
            if need > pos.pt:
                for p in pars[i:j]:
                    p.paragraph_format.tab_stops.clear_all()
                    p.paragraph_format.tab_stops.add_tab_stop(Pt(round(need)))
            i = j


# ---------------------------------------------------------------- в конвейере (pipeline.extract, struct_fill; assemble)
STRUCT_COVER = 0.9


def _paras(blocks):
    """Абзацы и заголовки в порядке чтения, со вложенными списками и ячейками таблиц."""
    for b in blocks:
        if b["k"] == "list":
            for it in b["items"]:
                yield from _paras(it)
        elif b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    yield from _paras(c["blocks"])
        elif b.get("runs"):
            yield b


def job_memory(segs, res):
    """Память сборки из сегментов задачи: текст страниц, абзацы разметки (loc.struct_only) — переводы после проверки,
    шлифовки и правок редактора; надписи скриншотов — не здесь (их несут картинки)."""
    return Memory([(s["text"], (res.get(str(s["id"])) or {}).get("ru")) for s in segs if "shot" not in s["loc"]])


def coverage(pdf, segs):
    """Доля уникальных текстов страниц PDF, которые есть в тексте дерева разметки (0 — разметки нет). Оглавление
    собирается полем, номер пункта — автонумерацией, их не считаем. Неполная разметка теряла бы текст в DOCX:
    такой PDF идёт прежним путём (UTS: 99,2 %, не нашлись только штамп и шапка листа — артефакты)."""
    import fitz
    import docx_struct
    from assemble import TOC
    from segments import needs_translation
    doc = fitz.open(pdf)
    if doc.xref_get_key(doc.pdf_catalog(), "StructTreeRoot")[0] == "null":
        return 0.0
    # дерево — по соглашениям экспорта Word (Lbl/LBody, TOC, ячейки TD); разметку Acrobat/PDF24 ставят автоматически:
    # абзацы прямо в строке таблицы мимо ячеек — DOCX разваливался (31RDGT: столбцы нулевой ширины, 102 стр., 16.10)
    meta = doc.metadata or {}
    if not re.search(r"Microsoft.{0,4}Word", (meta.get("producer") or "") + " " + (meta.get("creator") or ""), re.I):
        return 0.0
    blocks, _ = docx_struct.load(str(pdf))
    blob = "".join(KEY(text_of(b["runs"])) for b in _paras(blocks))
    keys = set()
    for s in segs:
        t = TAGS.sub("", s["text"])
        if "frags" not in s["loc"] or not needs_translation(t) or TOC.match(t.strip()):
            continue
        m = LABEL_T.match(t)
        keys.add(KEY(t[m.end():] if m else t))
    return sum(k in blob for k in keys) / len(keys) if keys else 0.0


def keep_code(hu, ru, groups):
    """Перевод с метками, где моноширинные группы groups (номера, Builder.code_groups) — по code_part, как их
    поставит сборка: таблица вычитки показывает то же, что DOCX."""
    src = {int(k): t for k, t in re.findall(r"<f(\d+)>(.*?)</f\1>", hu, re.S)}
    return re.sub(r"<f(\d+)>(.*?)</f\1>", lambda m: f"<f{m[1]}>{code_part(src[int(m[1])], m[2])}</f{m[1]}>"
                  if int(m[1]) in groups and int(m[1]) in src else m[0], ru, flags=re.S)


def missing(pdf, segs, res):
    """Тексты, которые сборка DOCX по разметке спросит у памяти задачи и не найдёт: [(текст, стр. от 0 или None,
    номера моноширинных групп)] в порядке документа. Абзац с метками начертания и он же без меток (запасной ход
    сборки) — один раз, первым."""
    import tempfile
    from pathlib import Path
    import docx_struct
    blocks, _ = docx_struct.load(str(pdf))
    L = docx_struct.layout(str(pdf), blocks)
    mem = job_memory(segs, res)
    probe = Builder(mem, lambda img: None)
    with tempfile.TemporaryDirectory() as tmp:
        probe.build(blocks, str(Path(tmp) / "probe.docx"), L)
    pages = [(KEY(text_of(p["runs"])), (p.get("geom") or {}).get("page")) for p in _paras(merge_inline(blocks))]
    out, seen = [], set()
    for t in mem.missing:
        k = KEY(t)
        if not k or k in seen:
            continue
        seen.add(k)
        # тот же абзац; иначе — абзац, частью которого текст является (или наоборот), сравнимой длины: короткий
        # абзац «UTS» содержится почти в любом тексте и уводил страницу в начало документа
        pg = next((p for pk, p in pages if pk == k), None) or next(
            (p for pk, p in pages if pk and (pk in k or k in pk) and 2 * min(len(pk), len(k)) >= max(len(pk), len(k))),
            None)
        out.append((t, pg - 1 if pg else None, sorted(probe.code_groups.get(t, ()))))
    return out
