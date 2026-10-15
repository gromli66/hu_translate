# -*- coding: utf-8 -*-
"""Сборка DOCX из блоков docx_struct.load: стили по оригиналу, настоящие списки и нумерация заголовков,
подписи «Рисунок N. …» полем SEQ, оглавление и список рисунков полями TOC, колонтитулы со штампом.
Прямого кегля и отступов на отдельных абзацах нет — всё стилями («как будто документ писали по-русски»)."""
import io
import re
from collections import Counter

from docx_struct import text_of

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
BULLETS = {"Disc": ("", "Symbol"), "Circle": ("o", "Courier New"), "Square": ("", "Wingdings")}
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
            c[(r.get("font") or "").lower().startswith(("consolas", "courier"))] += len(r["t"].strip())
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
        joins = prev is not None and prev["k"] == "p" and b["k"] == "p" and not _figure(prev) and not _figure(b) and (
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
            pt, nt = text_of(prev["runs"]), text_of(runs)
            if pt and nt and "t" in runs[0] and not pt.endswith((" ", "\t", "(", "„", "«")) \
                    and not nt.startswith((" ", ".", ",", ";", ":", ")", "”", "»")):
                runs[0]["t"] = " " + runs[0]["t"]              # «(lásd:» + «5.6.4.7» + «fejezet)» — с пробелами
            out[-1] = dict(prev, runs=prev["runs"] + runs, last_top=g.get("top", prev.get("last_top")),
                           tail_inline=bool(b.get("inline")))
            continue
        out.append(b)
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
LABEL_T = re.compile(r"^(\s*(?:<f\d+>)?)\s*(?:\d{1,3}|[A-Za-z]|[ivxIVX]{2,5})[.)]\s+(?=\S)")
LABEL_RU_T = re.compile(r"^(\s*(?:<f\d+>)?)\s*(?:\d{1,3}|[A-Za-zА-Яа-яЁё]|[ivxIVX]{2,5})[.)]\s+(?=\S)")


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
        s = new("Header UTS"); self._font(s, size=10); self._para(s, after=0, line=1.0, align=AL.CENTER)
        s = new("Stamp UTS"); self._font(s, "Times New Roman", 10); self._para(s, after=0, line=1.0)

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

    def _list_num(self, kind, start=1, level=0):
        """Новый экземпляр списка вида kind: нумерация — с номера первого пункта оригинала (start, на уровне level;
        список после рисунка продолжает «7.»), маркеры уровней • o ▪."""
        num = self.doc.part.numbering_part.element
        if kind not in self._abstract:
            aid = 901 + len(self._abstract)
            an = _el("w:abstractNum", w__abstractNumId=aid)
            an.append(_el("w:multiLevelType", w__val="hybridMultilevel"))
            for lv in range(4):
                l = _el("w:lvl", w__ilvl=lv)
                l.append(_el("w:start", w__val=1))
                fontn = None
                if kind in NUMFMT:
                    fmt = NUMFMT[kind] if lv == 0 else ["lowerLetter", "lowerRoman", "decimal"][(lv - 1) % 3]
                    l.append(_el("w:numFmt", w__val=fmt)); l.append(_el("w:lvlText", w__val=f"%{lv + 1}."))
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
            self._abstract[kind] = aid
        nid = 1000 + len(self.nums)
        n = _el("w:num", w__numId=nid); n.append(_el("w:abstractNumId", w__val=self._abstract[kind]))
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
        lead_img = bool(runs) and "img" in runs[0] and not runs[0].get("paren")
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
            ru = self.tr(src) if src.strip() else src
            out = [(groups[0][0] if groups else None, ru if ru is not None else src)]
        else:
            tagged = "".join(f"<f{i + 1}>{t}</f{i + 1}>" for i, t in enumerate(parts))
            code = {i + 1 for i, g in enumerate(groups) if g[0][2]}          # моноширинные группы — code_part
            if code:
                self.code_groups[tagged] = code
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
                out = [(groups[0][0], plain if plain is not None else "".join(parts))]
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
        # табуляция — позиции по оригиналу (пары «подпись ⇥ значение» на титуле, колонки в абзаце)
        from docx.shared import Pt
        for x in sorted({round(x) for r in runs for x in r.get("tabs", [])}):
            if x - self.left > 6:
                par.paragraph_format.tab_stops.add_tab_stop(Pt(x - self.left))

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
        par = tgt.add_paragraph(style=style)
        self._put_runs(par, runs, style_bold=style in ("Title UTS",))
        return par

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
        return (g.get("page") or 0) >= 2 and centered(g, self.L)

    def heading(self, b):
        full = text_of(b["runs"]).strip()
        get = getattr(self.tr, "get", None)
        ru = get(full) if get else None                      # в памяти PDF заголовок — с номером
        if ru:
            par = self.doc.add_paragraph(style=f"Heading {min(b['level'], 5)}")
            par.add_run(re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", ru))
            return
        runs = [dict(r) for r in b["runs"] if "t" in r]
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
        nid = self._list_num(b["kind"], b.get("start", 1), level)
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
        if len(xs) == ncol + 1:
            widths = [xs[i + 1] - xs[i] for i in range(ncol)]
        else:
            full = [r for r in rows if len(r) == ncol]
            lefts = [min([r[c]["geom"]["x0"] for r in full if r[c].get("geom")] or [None], key=lambda v: v or 0)
                     for c in range(ncol)]
            if all(v is not None for v in lefts) and len(set(round(v) for v in lefts)) == ncol:
                rights = lefts[1:] + [max(r[-1]["geom"]["x1"] for r in full if r[-1].get("geom")) + 6]
                widths = [max(rights[c] - lefts[c], 20) for c in range(ncol)]
        if widths:
            t.autofit = False
            for c in range(ncol):
                for cell in t.columns[c].cells:
                    cell.width = Pt(widths[c])
            grid = t._tbl.tblGrid
            for gc, w in zip(grid.findall(W + "gridCol"), widths):
                gc.set(W + "w", str(round(w * 20)))
        in_head = True
        for ri, row in enumerate(rows):
            if len(row) < ncol:                              # объединённая ячейка: хвост строки
                t.cell(ri, len(row) - 1).merge(t.cell(ri, ncol - 1))
            in_head = in_head and all(c["th"] for c in row)
            head = in_head and ri < 3 and len(rows) > 1
            if head and len(rows) > 3:
                t.rows[ri]._tr.get_or_add_trPr().append(_el("w:tblHeader"))
            # короткая строка не рвётся между листами (в оригинале строки целые; длинная — пусть рвётся, иначе не влезет)
            size = sum(len(text_of(x["runs"])) for c in row for x in c["blocks"] if x["k"] == "p")
            if size < 600 and not any("img" in r for c in row for x in c["blocks"] if x["k"] == "p" for r in x["runs"]):
                t.rows[ri]._tr.get_or_add_trPr().append(_el("w:cantSplit"))
            for ci, cell in enumerate(row):
                dc = t.cell(ri, ci)
                first = dc.paragraphs[0]
                used = False
                style = head_style if head else cell_style
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
                if head and fr.get("fill") and b.get("borders", True):
                    dc._tc.get_or_add_tcPr().append(_el("w:shd", w__val="clear", w__color="auto", w__fill=fr["fill"]))
        self.doc.add_paragraph(style="Table Text")        # Word: после таблицы — абзац

    def toc(self):
        if self._toc_n == 0:
            self.doc.add_paragraph(self.tr("Tartalom") or "Содержание", style="TOC Title UTS")
            par = self.doc.add_paragraph()
            _field(par, 'TOC \\o "1-4" \\h \\z \\u', "[Оглавление: в Word — правой кнопкой → «Обновить поле»]")
        else:
            self.doc.add_paragraph("Список рисунков", style="TOC Title UTS")
            par = self.doc.add_paragraph()
            _field(par, 'TOC \\h \\z \\c "Рисунок"', "[Список рисунков: в Word — правой кнопкой → «Обновить поле»]")
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
            self.toc()

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

    def _header(self, sec, L, logos_only=False):
        """Шапка: логотипы по краям, строки по центру (название и номер документа), линия снизу."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        hdr = sec.header
        hdr.is_linked_to_previous = True                     # новый раздел копирует ссылку на шапку титула —
        hdr.is_linked_to_previous = False                    # своя часть, иначе обе шапки пишутся в одну
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
        """Штамп листа: сетка по рамкам ячеек оригинала, номер листа «N/NN» — полями PAGE и NUMPAGES."""
        from docx.shared import Pt
        ftr = sec.footer
        ftr.is_linked_to_previous = True
        ftr.is_linked_to_previous = False
        grid = L["stamp"]
        if not grid:
            return
        ncol = max(len(r["cells"]) for r in grid)
        t = ftr.add_table(rows=len(grid), cols=ncol, width=Pt(sum(L["stamp_widths"]) or 400))
        t.style = self.doc.styles["Table Grid"]
        for c, w in zip(t.columns, L["stamp_widths"]):
            for cell in c.cells:
                cell.width = Pt(w)
        for ri, row in enumerate(grid):
            for ci, txt in enumerate(row["cells"]):
                p = t.cell(ri, ci).paragraphs[0]
                p.style = self.doc.styles["Stamp UTS"]
                m = re.match(r"^(.*?)(\d+)\s*/\s*(\d+)\s*$", txt)
                if m:
                    head = (m.group(1).strip() + "    ") if m.group(1).strip() else ""
                    r = p.add_run(head); r.bold = row["bold"]
                    f = _field(p, "PAGE", m.group(2)); f.bold = row["bold"]
                    r = p.add_run("/"); r.bold = row["bold"]
                    f = _field(p, "NUMPAGES", m.group(3)); f.bold = row["bold"]
                    p.alignment = 2
                else:
                    r = p.add_run(self.tr(txt) or txt if txt else ""); r.bold = row["bold"]
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
        first = self.doc.sections[0]
        prev_bottom = None
        self._page(first, L, L["title_left"], 20, 40, 30)
        self.left = L["title_left"]
        in_title = True
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
                continue                                     # логотипы титула — в шапку титула
            if in_title and b["k"] == "p" and all("img" in r for r in b["runs"]):
                if g and g.get("top", 999) < 110:
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
        self._header(self.doc.sections[0], L, logos_only=True)
        if title_foot:
            foot = self.doc.sections[0].footer
            foot.is_linked_to_previous = False
            first_p = foot.paragraphs[0] if foot.paragraphs and not foot.paragraphs[0].text else None
            for b in title_foot:
                self.paragraph(b, tgt=foot)
            if first_p is not None:
                first_p._p.getparent().remove(first_p._p)
        self.doc.settings.element.append(_el("w:updateFields", w__val="true"))
        self.doc.save(out)


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
