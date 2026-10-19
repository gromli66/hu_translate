# -*- coding: utf-8 -*-
"""Сборка DOCX из блоков docx_struct.load: стили по оригиналу, настоящие списки и нумерация заголовков,
подписи «Рисунок N. …» полем SEQ, оглавление и список рисунков полями TOC, колонтитулы со штампом.
Прямого кегля и отступов на отдельных абзацах нет — всё стилями («как будто документ писали по-русски»)."""
import io
import re
from collections import Counter

from docx_struct import text_of, insert_imgs, PAGE_MARK

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
BULLETS = {"Disc": ("\uf0b7", "Symbol"), "Circle": ("o", "Courier New"), "Square": ("\uf0a7", "Wingdings"),
           "Dash": ("–", None)}                     # тире из текста пункта (ÜFK «– ha az …») — шрифтом абзаца
NUMFMT = {"Decimal": "decimal", "LowerAlpha": "lowerLetter", "UpperAlpha": "upperLetter",
          "LowerRoman": "lowerRoman", "UpperRoman": "upperRoman"}
LEAD = re.compile(r"^\s*[•▪◦·o\-–\uf020-\uf0ff](?:\s+|$)")    # и маркер Symbol/Wingdings (частная область)
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


def is_label(b):
    """Строка-подпись внутри раздела («Beavatkozások:», «Korlátozó feltétel:»): короткая, полужирная, с двоеточием в
    конце — свой стиль с интервалами оригинала. Иначе её больший отступ сверху (19 pt против 6 у ÜFK) уходил в медиану
    интервала всех абзацев, и текст разъезжался (157 листов собирались в 169, 16.10)."""
    runs = [r for r in b.get("runs", []) if "t" in r and r["t"].strip()]
    t = text_of(b.get("runs", [])).strip()
    return bool(runs) and len(t) <= 80 and t.endswith(":") and all(r.get("b") for r in runs)


def is_code(runs):
    """Абзац кода: больше половины символов — моноширинным (по символам, а не по прогонам: абзац «текст + имя файла
    Consolas» иначе уходил в стиль кода целиком)."""
    c = Counter()
    for r in runs:
        if "t" in r and r["t"].strip():
            # body_mono — основной шрифт документа моноширинный (3SZ19 набран Courier): это не код (docx_layout)
            c[(r.get("font") or "").lower().startswith(("consolas", "courier")) and not r.get("body_mono")] += len(r["t"].strip())
    return c[True] > c[False]


FONT_GUESS = [(r"mono|courier|consolas|menlo|lucida console", "modern"),
              (r"sans|arial|helvetica|calibri|verdana|tahoma|segoe|trebuchet|franklin|futura|gill|myriad|roboto|"
               r"lato|ubuntu|dejavu sans|liberation sans|pt sans|century gothic", "swiss"),
              (r"times|georgia|cambria|garamond|book antiqua|palatino|serif|minion|bookman|century|baskerville", "roman")]


def font_table(path, classes):
    """Таблица шрифтов DOCX — класс каждого шрифта документа (рубленый / с засечками / моноширинный): Word и
    LibreOffice подбирают по нему замену неустановленному шрифту. Класс — из PDF (docx_struct.font_classes), иначе
    по имени."""
    import zipfile
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    if "word/fontTable.xml" not in files:
        return
    used = set()
    for n, b in files.items():
        if re.match(r"word/(document|styles|header\d*|footer\d*|numbering)\.xml$", n):
            used |= set(re.findall(r'w:(?:ascii|hAnsi|cs)="([^"]+)"', b.decode("utf8", "ignore")))
    ft = files["word/fontTable.xml"].decode("utf8")
    have = set(re.findall(r'<w:font w:name="([^"]+)"', ft))
    add = []
    for f in sorted(used - have):
        k = classes.get(f) or next((c for rx, c in FONT_GUESS if re.search(rx, f, re.I)), None)
        if k:
            add.append(f'<w:font w:name="{f}"><w:charset w:val="00"/><w:family w:val="{k}"/>'
                       f'<w:pitch w:val="{"fixed" if k == "modern" else "variable"}"/></w:font>')
    if not add:
        return
    files["word/fontTable.xml"] = ft.replace("</w:fonts>", "".join(add) + "</w:fonts>").encode("utf8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)


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
        # строчный элемент — только у конца абзаца (тот же лист, до следующей строки): подписи рисунков JPDS тоже
        # строчные, и «Рисунок 14 … Рисунок 15 … Рисунок 20» слипались в один абзац с разных листов (17.10)
        size = pg.get("size") or 10
        low = max(pg.get("bottom") or pg.get("top") or 0, (prev or {}).get("last_bottom") or 0)
        near = not g.get("top") or not pg.get("top") or g.get("page") == pg.get("page") \
            and pg["top"] - 3 <= g["top"] <= low + 0.6 * size
        cap = bool(__import__("docx_struct").CAPTION_ANY.match(text_of(b.get("runs") or []))) and \
            abs((g.get("top") or 0) - (prev or {}).get("last_top", pg.get("top") or 0)) >= 6
        joins = prev is not None and prev["k"] == "p" and b["k"] == "p" and not _figure(prev) and not _figure(b) \
            and not b.get("seg") and not prev.get("seg") and (      # абзацы по сегментам листа уже собраны (docx_layout)
            b.get("inline") and near and not cap
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
                # пробел — только если на листе между кусками промежуток: Word кладёт «i» и «5» в разные элементы
                # («Core i 5», «по явятся» в ТЭ357, 17.10)
                a, z = prev["runs"][-1], runs[0]
                touch = a.get("x1") is not None and z.get("x") is not None and a.get("page") == z.get("page") \
                    and abs((a.get("top") or 0) - (z.get("top") or 0)) < 0.5 * (a.get("size") or 10) \
                    and -1 <= z["x"] - a["x1"] < 0.12 * (a.get("size") or 10)
                if pt and nt and "t" in runs[0] and not touch and not pt.endswith((" ", "\t", "(", "„", "«")) \
                        and not nt.startswith((" ", ".", ",", ";", ":", ")", "”", "»")):
                    runs[0]["t"] = " " + runs[0]["t"]          # «(lásd:» + «5.6.4.7» + «fejezet)» — с пробелами
                merged = prev["runs"] + runs
            out[-1] = dict(prev, runs=merged, last_top=g.get("top", prev.get("last_top")),
                           last_bottom=max(g.get("bottom") or 0, prev.get("last_bottom") or 0),
                           tail_inline=bool(b.get("inline")), **__import__("docx_struct").join_notes(prev, b))
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


from docx.exceptions import InvalidSpanError


_TRPR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader",
         "tblCellSpacing", "jc", "hidden"]


def _trpr_put(tr, el):
    """Свойство строки таблицы — на место по схеме (Word строг к порядку элементов trPr)."""
    pr = tr.get_or_add_trPr()
    k = _TRPR.index(el.tag.split("}")[1])
    nxt = next((c for c in pr if c.tag.split("}")[1] in _TRPR and _TRPR.index(c.tag.split("}")[1]) > k), None)
    nxt.addprevious(el) if nxt is not None else pr.append(el)


def _span_starts(row, xs, ncol):
    """Ячейки неполной строки — по столбцам рамки (xs): каждой ячейке — группа соседних столбцов, текст ячейки внутри
    своей группы и ближе всего к её середине или левому краю. Шапка ÜFK «Eltérő állapot» стоит над двумя столбцами
    «A. | текст» — раньше объединялся только хвост строки, и шапка съезжала (16.10). Столбец может остаться ничьим —
    его занимает ячейка строкой выше («Описание» на две строки шапки над «КРТ №1 … №4», ИН-1172: «КРТ №1» растягивалась
    на столбец «Описание», 17.10). Без рамки — хвост строки, как раньше. Возвращает [(первый, последний столбец)]."""
    k = len(row)
    tail = [(i, i) for i in range(k - 1)] + [(k - 1, ncol - 1)] if k else []
    if len(xs) != ncol + 1 or k == 0:
        return tail
    geo = [c.get("geom") for c in row]
    if all(g is None for g in geo):
        return tail
    from functools import lru_cache
    skip = 25.0                                             # ничей столбец — дороже точного места, дешевле чужого

    def cost(i, a, z):
        g = geo[i]
        if g is None:
            return 5.0 * (z - a)                             # пустая ячейка — по столбцу, если ничто не мешает
        gl, gr = xs[a], xs[z + 1]
        inside = 0 if gl - 3 <= g["x0"] and g["x1"] <= gr + 3 else 1000
        # столбец, до которого текст ячейки не доходит, — дороже ничьего: ничей закроет ячейка сверху (объединение по
        # вертикали, OAH 1.1 «MEGJEGYZÉS»: описание растягивалось на два столбца, 19.10)
        wide = sum(30 for c in range(a + 1, z + 1) if g["x1"] < xs[c] + 3)
        return inside + wide + min(abs((gl + gr) / 2 - (g["x0"] + g["x1"]) / 2), abs(g["x0"] - gl - 5))

    @lru_cache(None)
    def best(i, a):
        if i == k:
            return skip * (ncol - a), ()
        out = (float("inf"), ())
        for s0 in range(a, ncol - (k - i - 1)):
            for z in range(s0, ncol - (k - i - 1)):
                c, rest = best(i + 1, z + 1)
                c += cost(i, s0, z) + skip * (s0 - a)
                if c < out[0]:
                    out = (c, ((s0, z),) + rest)
        return out

    c, spans = best(0, 0)
    return list(spans) if c < 1000 else tail


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


def _head_dots(blocks):
    """Уровни заголовков, у которых в оригинале точка после номера («1. Introduction», а «2.1 Scope» — без: Eskom):
    больше половины номеров уровня — с точкой."""
    cnt = {}
    for b in blocks:
        if b["k"] != "h":
            continue
        m = re.match(r"\s*\d+(?:\.\d+)*(\.?)\s", text_of(b["runs"]))
        if m:
            c = cnt.setdefault(min(b["level"], 5), [0, 0])
            c[bool(m[1])] += 1
    return {lv for lv, (no, yes) in cnt.items() if yes > no}


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
    for ri, r in enumerate(rows):
        head = head and all(c["th"] for c in r)
        # первая строка без пометки шапки — своим голосом: шапка по центру решала за всё тело, и текст ячеек ÜFK
        # (скан) вставал по центру, 19.10
        key = "head" if head else "first" if ri == 0 else "body"
        if len(r) != ncol:
            continue
        for ci, c in enumerate(r):
            g = c.get("geom")
            if not g or g.get("page") != page or (g.get("lines") or 1) > 1:
                continue
            left, right = g["x0"] - xs[ci], xs[ci + 1] - g["x1"]
            v = "center" if abs(left - right) < 4 and left > 8 else "right" if right < 10 < left - 8 else "left"
            votes.setdefault((key, ci), []).append(v)
    out = {}
    for k, v in votes.items():
        best = Counter(v).most_common(1)[0][0]
        if best != "left" or k[0] == "first":
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
PIPE = re.compile(r"^\s*\|\s*|\s*\|\s*$")
KEY = lambda t: re.sub(r"\s+", "", TAGS.sub("", t or ""))
HU_LETTERS = re.compile(r"[áéíóöőúüűÁÉÍÓÖŐÚÜŰ]")


def code_part(src, ru):
    """Моноширинная часть абзаца с метками. Без венгерских букв (команда, путь, вывод программы) — как в оригинале:
    портал переводил «sync» → «синхронизировать», «gatewayconfig ipaddr» → «настройка IP-адреса шлюза» (UTS, 15.10).
    Заполнитель «<beállítani kívánt IP cím>» переводится, но в тех же угловых скобках (портал их терял)."""
    if not HU_LETTERS.search(src):
        return src
    # листинг из нескольких строк: строка без венгерских букв (ключ=значение, путь, команда) — как в оригинале,
    # с венгерскими (комментарий) — переводом той же строки (UTS: «LogFile=…» → «Файл журнала=…», 19.10)
    sl, rl = src.split("\n"), (ru or "").split("\n")
    if len(sl) > 1 and len(sl) == len(rl):
        return "\n".join(r if HU_LETTERS.search(a) else a for a, r in zip(sl, rl))
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
        # «Alfanumerika |» — вертикаль таблицы, прочитанная OCR; сборка (docx_layout) спрашивает текст без неё
        c = PIPE.sub("", hu)
        if c != hu and c.strip():
            self.add(c, PIPE.sub("", ru))
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
        s = new("List Cell UTS", "Table Text")                   # пункт в ячейке: шрифт ячейки, влево — по ширине
        s.paragraph_format.space_after = Pt(self._space("list", 3))   # в узкой ячейке «есть   готовность   от» (ТЭ357)
        s = new("Center UTS"); s.paragraph_format.alignment = AL.CENTER   # Ф13: абзац по центру в оригинале
        s = new("Label UTS")                                      # строка-подпись раздела (is_label)
        body_after = self._space("body", 8)
        self._para(s, before=max(0.0, self._space("label_before", body_after) - body_after),
                   after=self._space("label_after", body_after), keep=True)
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

    def _list_ind(self, lv, kind=None):
        """Ф7: отступы пункта уровня lv (маркер, текст) в pt — по профилю, своего вида списка, если замерен («a.)» у
        ÜFK — от поля, «–» — с отступом); уровни глубже замеренных — с тем же шагом."""
        own = ((self.P.get("lists_kind") or {}).get(kind) or {}).get(str(lv))
        if own:
            return own["marker"], own["text"]
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

    def _lvl_ind(self, lv, kind):
        """Отступы уровня списка (маркер, текст), pt, — как в нумерации: под метку и зазор после неё."""
        mark, text = self._list_ind(lv, kind)
        if kind in ("LowerRoman", "UpperRoman") and lv == 0:
            text = max(text, mark + 24)                     # «iii.», «iv.» шире висячего отступа — текст уезжал (ÜFK)
        elif kind not in NUMFMT and kind != "Label":
            # маркер шире висячего отступа (у оригинала «- » — дефис и пробел, 7 pt): тире «–» вставало вплотную к
            # тексту (НГТУ, 19.10) — место под знак и зазор
            text = max(text, mark + 9)
        return mark, text

    def _list_num(self, kind, start=1, level=0, suffix=".", own=False):
        """Новый экземпляр списка вида kind: нумерация — с номера первого пункта оригинала (start, на уровне level;
        список после рисунка продолжает «7.»), вид метки — как в оригинале («a.» / «a.)»), маркеры уровней • o ▪."""
        num = self.doc.part.numbering_part.element
        key = (kind, suffix or ".", own)
        if key not in self._abstract:
            aid = 901 + len(self._abstract)
            an = _el("w:abstractNum", w__abstractNumId=aid)
            an.append(_el("w:multiLevelType", w__val="hybridMultilevel"))
            for lv in range(4):
                l = _el("w:lvl", w__ilvl=lv)
                l.append(_el("w:start", w__val=1))
                fontn = None
                if kind == "Label":                         # метка — в тексте пункта («ba)», «(a)»): без маркера
                    l.append(_el("w:numFmt", w__val="none")); l.append(_el("w:lvlText", w__val=""))
                elif kind in NUMFMT:
                    fmt = NUMFMT[kind] if lv == 0 else ["lowerLetter", "lowerRoman", "decimal"][(lv - 1) % 3]
                    l.append(_el("w:numFmt", w__val=fmt))
                    l.append(_el("w:lvlText", w__val=f"%{lv + 1}" + ((suffix or ".") if lv == 0 else ".")))
                else:
                    # тире — на всех уровнях (у 3SZ19, 3PR42 вложенные пункты — тоже «-», а не «o»)
                    # own — знак с листа (текстовый слой): тот же на любом уровне
                    ch, fontn = BULLETS.get(kind if lv == 0 or kind == "Dash" or own else ["Circle", "Square", "Disc"][(lv - 1) % 3],
                                            BULLETS["Disc"])
                    l.append(_el("w:numFmt", w__val="bullet")); l.append(_el("w:lvlText", w__val=ch))
                l.append(_el("w:lvlJc", w__val="left"))
                mark, text = self._lvl_ind(lv, kind)
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
        data = img.get("png") or self.image(img)          # png — снимок рисунка с листа (docx_layout)
        if data:
            w = float(img["x1"]) - float(img["x0"])
            if getattr(self, "L", None):                    # не шире поля набора
                w0 = w
                w = min(w, self.L["W"] - self.L["left"] - self.L["right"])
                # и не выше поля набора без места под подпись: снимок во весь лист уводил «Рисунок 17» на следующий
                # лист (ИН-1065, 18.10)
                h = (float(img.get("bottom") or 0) - float(img.get("top") or 0)) * w / w0 if w0 > 0 else 0
                room = self.L["H"] - self.L["top"] - self.L["bottom"] - 48
                if h > room > 100:
                    w *= room / h
            par.add_run().add_picture(io.BytesIO(data), width=Pt(w))

    def _own_size(self, runs, style):
        """Кегль абзаца, если он не кегль его стиля (на 1 pt и больше): титул 14 pt и выходные данные 10 pt при теле
        12 pt печатались 12 pt (OAH 4.22, 18.10). Только при точном кегле — из знаков PDF (L["sizes"]), не из оценки
        распознавания."""
        if not (getattr(self, "L", None) or {}).get("sizes"):
            return None
        from docx.shared import Pt
        c = Counter()
        for r in runs:
            if "t" in r and r.get("size"):
                c[float(r["size"])] += len(r["t"].strip())
        if not c or not max(c.values()):
            return None
        size = c.most_common(1)[0][0]
        st = self.doc.styles[style]
        while st is not None and st.font.size is None:
            st = st.base_style
        have = st.font.size.pt if st is not None else (self.doc.styles["Normal"].font.size or Pt(10)).pt
        return size if abs(size - have) >= 1 else None

    def _put_runs(self, par, runs, style_bold=False, size=None):
        """Текст с переводом. Одно начертание — перевод целиком; несколько (жирное «FIGYELEM!» + текст) — перевод
        с метками <fN>, каждой группе своё начертание. Картинки в начале — перед текстом, иначе — после. size — свой
        кегль абзаца (_own_size)."""
        from docx.shared import RGBColor, Pt
        texts = [r for r in runs if "t" in r]
        imgs = [r for r in runs if "img" in r and not r.get("paren") and not r.get("at_tab")]
        timgs = [r for r in runs if r.get("at_tab")]          # значки на месте табуляции (docx_struct.inline_pictures)
        pimgs = [r for r in runs if "img" in r and r.get("paren")]      # значки в «()» — на то же место в переводе
        # значок в начале — первым после пустых прогонов (табуляция после снятого маркера пункта)
        head = next((r for r in runs if "img" in r or r["t"].strip()), None)
        lead_img = head is not None and "img" in head and not head.get("paren")
        groups = []
        for r in texts:
            # моноширинный — «код», если это не основной шрифт документа (3SZ19 набран Courier: иначе каждая группа
            # абзаца с полужирной командой шла «кодом», и code_part возвращал венгерское «NYITNI» вместо перевода)
            key = (bool(r.get("b")), bool(r.get("i")),
                   (r.get("font") or "").lower().startswith(("consolas", "courier")) and not r.get("body_mono"),
                   r.get("color") if r.get("color") not in ("000000", None) else None, bool(r.get("u")))
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
            if groups and groups[0][0][2] and ru is not None:      # весь абзац — код: то же правило, что у кусков кода
                ru = code_part(src, ru)
            out = [(groups[0][0] if groups else None, ru if ru is not None else src)]
        elif all(p.endswith("	") for p in parts[:-1]):
            # части через табуляцию, у каждой своё начертание (строка оглавления «1.3.» ⇥ «Generátor …» ⇥ «16»,
            # docx_layout) — перевод по частям: память их знает, метки начертания портала не нужны и табуляции целы
            for g, p in zip(groups, parts):
                body = p[:-1] if p.endswith("	") else p
                t = self.tr(body) if body.strip() else body
                out.append((g[0], (t if t is not None else body) + ("	" if p.endswith("	") else "")))
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
            # индексы распознавания портала («N<sub>R</sub>», «m<sup>3</sup>») — индексами, а не тегами в тексте (3PR42)
            for piece in re.split(r"(<su[bp]>.*?</su[bp]>)", t):
                m = re.fullmatch(r"<(su[bp])>(.*?)</su[bp]>", piece, re.S)
                if not piece or m and not m.group(2):
                    continue
                run = par.add_run(m.group(2) if m else piece)
                if size:
                    run.font.size = Pt(size)
                if m:
                    if m.group(1) == "sub":
                        run.font.subscript = True
                    else:
                        run.font.superscript = True
                if key:
                    b, i, code, color, u = key
                    if b and not style_bold:
                        run.bold = True
                    if i:
                        run.italic = True
                    if u:
                        run.underline = True
                    if code and par.style.name != "Code UTS":
                        run.font.name = "Consolas"
                    if color and color != "ffffff":
                        run.font.color.rgb = RGBColor.from_string(color.upper())
        for key, t in out:
            while t and timgs and "	" in t:
                k = t.index("	")
                text_run(t[:k].rstrip() + " ", key)
                self._img_run(par, timgs.pop(0)["img"])
                t = " " + t[k + 1:].lstrip()
            while t and pimgs and re.search(r"\(\s*\)", t):
                m = re.search(r"\(\s*\)", t)
                text_run(t[:m.start() + 1], key)
                self._img_run(par, pimgs.pop(0)["img"])
                t = ")" + t[m.end():]
            if t:
                text_run(t, key)
        for im in ([] if lead_img else imgs) + pimgs + timgs:
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
        # в ячейке — от левого края текста ячейки, не от поля листа: табулятор за краем узкой ячейки уводил текст после
        # него («A.1. ⇥ Helyre kell állítani…» ÜFK II — пропадало, 18.10)
        base = getattr(self, "_tab_base", None)
        base = self.left if base is None else base
        for x in sorted({round(x) for r in runs for x in r.get("tabs", [])}):
            if x - base > 6:
                par.paragraph_format.tab_stops.add_tab_stop(Pt(x - base))
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
        elif style == "Normal" and is_label(b):
            style = "Label UTS"
        size = (b.get("geom") or {}).get("size")
        if style == "Normal" and size:
            # пороги — от кегля тела, а не числом: у инструкций по ГОСТ тело 14 pt, и «от 13 pt — вводный лист»
            # уводило все 632 абзаца ИН-ЦАСУТП-1065 в стиль без отступа первой строки (16.10)
            bs = self._role("body")[1] or 10
            if size >= max(17, 1.5 * bs):
                style = "Title UTS"
            elif size >= max(13, bs + 2.5):
                style = "Front UTS"
            elif size <= min(8, bs - 2):
                style = "Small UTS"
        if style in ("Normal", "Code UTS") and tgt is self.doc and self.L and self._centered(b):
            style = "Center UTS" if style == "Normal" else "Code Center UTS"
        if style == "Normal" and b.get("indent"):
            style = self._indented(b["indent"])              # отступ оригинала — уровнем стиля (docx_layout)
        elif style == "Normal" and tgt is self.doc and self.L and b.get("geom") and getattr(self, "left", None):
            # абзац тела правее поля (OAH: тело сдвинуто от заголовков на 37 pt) — стилем отступа, шаг 3 pt;
            # висячий (первая строка левее остальных: «номер ⇥ название» перечня стандартов) — стилем с висячим
            g = b["geom"]
            rest, fx = g.get("rest_x0"), g.get("first_x0")
            if rest is not None and fx is not None and fx < rest - 5:
                d, h = round((rest - self.left) / 3) * 3, round((rest - fx) / 3) * 3
                if 0 < d < 0.6 * (self.L["W"] - self.L["left"] - self.L["right"]):
                    style = self._indented(d, h)
            else:
                # сдвиг абзаца — по строкам после первой; у однострочного — за вычетом отступа первой строки стиля
                # (МТ-ЦАСУТП-14: первая строка с 36 pt давала однострочным ещё 36 pt слева, 16.10)
                x = rest if rest is not None else (fx if fx is not None else g.get("x0") or 0) - \
                    ((self.P or {}).get("first_line") or 0)
                d = round((x - self.left) / 3) * 3
                if 6 <= d < 0.5 * (self.L["W"] - self.L["left"] - self.L["right"]):
                    style = self._indented(d)
        # абзац у поля без красной строки в документе с красной строкой («Пояснительная записка должна содержать…»
        # НГТУ, 85 pt при теле с 99) — стилем тела без отступа первой строки (19.10)
        g = b.get("geom") or {}
        if (self.P or {}).get("first_line") and style == "Normal" and tgt is self.doc \
                and getattr(self, "left", None) is not None and g.get("first_x0") is not None \
                and abs(g["first_x0"] - self.left) <= 3 \
                and (g.get("rest_x0") is None or abs(g["rest_x0"] - g["first_x0"]) <= 3):
            style = self._flush()
        par = tgt.add_paragraph(style=style)
        # в ячейке кегль — стилем таблицы (ужатым под ширину, «Table Text Small»)
        self._put_runs(par, runs, style_bold=style in ("Title UTS",),
                       size=self._own_size(runs, style) if tgt is self.doc else None)
        # строка-заголовок не из разметки (вся полужирная, короткая, без знака конца фразы: «FŐIGAZGATÓI ELŐSZÓ» OAH,
        # «4. ТЕХНИЧЕСКИЕ ДАННЫЕ …» ИН-1065) — не последней строкой листа, со следующим абзацем (оставалась внизу
        # листа, текст — на следующем, 18.10)
        ink = [r for r in runs if "t" in r and r["t"].strip()]
        if tgt is self.doc and ink and all(r.get("b") for r in ink) and len(text) <= 160 \
                and not text.endswith((".", ",", ";", ":")) and re.search(r"[^\W\d_]{3}", text):
            par.paragraph_format.keep_with_next = True
        # вводная фраза сразу под заголовком («… az alábbi hibaüzenet jelenik meg:» + рисунок) — с тем, что за ней:
        # заголовок с одной фразой оставался внизу листа, рисунок уходил на следующий (UTS 10.6, 18.10). Только под
        # заголовком и короткая: у каждой фразы с «:» перед снимком JPDS вырос на 3 листа пустот (18.10)
        if tgt is self.doc and text.endswith(":") and len(text) <= 250 and getattr(self, "_last_k", None) == "h":
            par.paragraph_format.keep_with_next = True
        for n in b.get("notes") or []:
            self._footnote(par, n)
        return par

    def _footnote(self, par, note):
        """Сноска Word: знак ссылки в абзаце (на место оригинала — долей длины текста: в конце или после слова) и
        текст сноски с переводом в footnotes.xml (часть создаётся при первой сноске)."""
        import copy
        from docx.enum.style import WD_STYLE_TYPE
        from docx.shared import Pt
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.part import XmlPart
        from docx.opc.packuri import PackURI
        from docx.oxml import parse_xml
        st = self.doc.styles
        if not getattr(self, "notes_part", None):
            ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
            sep = ('<w:footnote w:type="{t}" w:id="{i}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                   'w:lineRule="auto"/></w:pPr><w:r><w:{t}/></w:r></w:p></w:footnote>')
            xml = (f"<w:footnotes {ns}>" + sep.format(t="separator", i=-1) + sep.format(t="continuationSeparator", i=0)
                   + "</w:footnotes>")
            self.notes_part = XmlPart(PackURI("/word/footnotes.xml"),
                                      "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                                      parse_xml(xml), self.doc.part.package)
            self.doc.part.relate_to(self.notes_part, RT.FOOTNOTES)
            self.note_id = 0
            names = [x.name for x in st]
            if "Footnote Text" not in names:
                s_ = st.add_style("Footnote Text", WD_STYLE_TYPE.PARAGRAPH)
                s_.base_style = st["Normal"]
                sizes = [r.get("size") for r in note["runs"] if r.get("size")]
                s_.font.size = Pt(sizes[0] if sizes else 9)
                pf = s_.paragraph_format
                pf.space_before, pf.space_after, pf.first_line_indent, pf.line_spacing = Pt(0), Pt(0), Pt(0), 1.0
            if "Footnote Reference" not in names:
                c_ = st.add_style("Footnote Reference", WD_STYLE_TYPE.CHARACTER)
                c_.font.superscript = True
        self.note_id += 1
        nid = str(self.note_id)
        tmp = self.doc.add_paragraph(style="Footnote Text")    # перевод — тем же путём, что у абзаца
        self._put_runs(tmp, [dict(r) for r in note["runs"]])
        p_el = tmp._p
        p_el.getparent().remove(p_el)
        ref = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr>'
                        '<w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>')
        sp = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                       '<w:t xml:space="preserve"> </w:t></w:r>')
        ppr = p_el.find(W + "pPr")
        at = list(p_el).index(ppr) + 1 if ppr is not None else 0
        p_el.insert(at, sp)
        p_el.insert(at, ref)
        fn = parse_xml(f'<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="{nid}"/>')
        fn.append(p_el)
        self.notes_part.element.append(fn)
        mark = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr>'
                         f'<w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{nid}"/></w:r>')
        runs = par._p.findall(W + "r")
        texts = [("".join(t.text or "" for t in r.findall(W + "t")), r) for r in runs]
        total = sum(len(t) for t, _ in texts)
        if not texts or total == 0:
            par._p.append(mark)
            return
        if note.get("pos", 1.0) >= 0.9:
            texts[-1][1].addnext(mark)
            return
        goal, acc = note["pos"] * total, 0
        for t, r in texts:                                  # после слова, ближайшего к месту оригинала
            if acc + len(t) >= goal:
                k = int(goal - acc)
                m = re.search(r"\s", t[k:])
                cut = k + m.start() if m else len(t)
                if cut < len(t):
                    r2 = copy.deepcopy(r)
                    for el in r.findall(W + "t"):
                        r.remove(el)
                    for el in r2.findall(W + "t"):
                        r2.remove(el)
                    a = _el("w:t"); a.text = t[:cut]; a.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    b2 = _el("w:t"); b2.text = t[cut:]; b2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    r.append(a); r2.append(b2)
                    r.addnext(r2)
                r.addnext(mark)
                return
            acc += len(t)
        texts[-1][1].addnext(mark)

    def _flush(self):
        """Стиль тела без отступа первой строки — абзацу у поля в документе с красной строкой."""
        from docx.enum.style import WD_STYLE_TYPE
        from docx.shared import Pt
        name = "Body Flush UTS"
        if name not in [s.name for s in self.doc.styles]:
            s = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = self.doc.styles["Normal"]
            s.paragraph_format.first_line_indent = Pt(0)
        return name

    def _indented(self, ind, hang=0):
        """Стиль абзаца с отступом слева ind pt (и висячим hang pt): отступ — стилем, а не прямым оформлением абзаца."""
        from docx.enum.style import WD_STYLE_TYPE
        from docx.shared import Pt
        name = f"Indent {int(ind)} UTS" if not hang else f"Hanging {int(ind)} {int(hang)} UTS"
        if name not in [s.name for s in self.doc.styles]:
            s = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = self.doc.styles["Normal"]
            s.paragraph_format.left_indent = Pt(ind)
            if hang:
                s.paragraph_format.first_line_indent = Pt(-hang)
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
        txt = text_of(b.get("runs") or [])
        if (g.get("page") or 0) >= 2 and len(txt) > 40 and LABEL_T.match(txt):
            return False                # пункт с номером по центру не бывает: «1. Az érvényét …» с отступом (ÜFK II);
            # короткие «A. melléklet», «II. FEJEZET» и титул — по месту
        if (g.get("page") or 0) == 1 and g.get("x0") is not None:
            # титул: центр — по листу (у титула свои поля); «Érvényes … visszavonásig» ÜFK стоял влево (16.10)
            return abs((g["x0"] + g["x1"]) / 2 - self.L["W"] / 2) < 15 and g["x1"] - g["x0"] < 0.7 * self.L["W"]
        return (g.get("page") or 0) >= 2 and centered(g, self.L)

    def heading(self, b):
        fresh = getattr(self, "fresh", False)
        par = self._heading(b)
        if b.get("center") and par is not None:              # заголовок по центру в оригинале (НГТУ, docx_layout)
            par.alignment = 1
        if b.get("break") and par is not None and not fresh:     # разрыв листа оригинала (docx_struct.load)
            par.paragraph_format.page_break_before = True

    def _heading(self, b):
        full = text_of(b["runs"]).strip()
        get = getattr(self.tr, "get", None)
        ru = get(full) if get else None                      # в памяти PDF заголовок — с номером
        auto = getattr(self, "autonum", True)
        if ru:
            par = self.doc.add_paragraph(style=self._head_style(b))
            run = par.add_run(re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", ru) if auto else ru)
            # цвет заголовка — общий цвет его текста в оригинале (синие «Приложение 1 …» ИН-1172 выходили чёрными)
            cols = {(r.get("color") or "000000").lower() for r in b["runs"] if r.get("t", "").strip()}
            if len(cols) == 1 and cols - {"000000", "ffffff"}:
                from docx.shared import RGBColor
                run.font.color.rgb = RGBColor.from_string(cols.pop().upper())
            return par
        runs = [dict(r) for r in b["runs"] if "t" in r]
        if not auto:                                         # номер оригинала — текстом, его не даёт нумерация
            for r in runs:
                r["b"] = False
            par = self.doc.add_paragraph(style=self._head_style(b))
            self._put_runs(par, runs)
            return par
        while runs and re.fullmatch(r"\s*\d+(?:\.\d+)*\.?\s*", runs[0]["t"]):
            runs = runs[1:]                                  # номер отдельным прогоном — его даёт нумерация стиля
        if runs:                                             # и с точкой: «1. Introduction» печаталось «1 1. Introduction»
            runs[0]["t"] = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s+", "", runs[0]["t"]).lstrip()
        for r in runs:
            r["b"] = False                                   # начертание — от стиля
        par = self.doc.add_paragraph(style=self._head_style(b))
        self._put_runs(par, runs)
        if not text_of(runs).strip():                        # заголовок картинками без пункта оглавления — картинки
            for r in b["runs"]:
                if "img" in r:
                    self._img_run(par, r["img"])
        return par

    def _head_style(self, b):
        """Стиль заголовка; заголовка нет в оглавлении оригинала — тот же вид, но без уровня структуры (поле TOC его
        не берёт)."""
        from docx.enum.style import WD_STYLE_TYPE
        name = f"Heading {min(b['level'], 5)}"
        keys = getattr(self, "toc_keys", None)
        t = text_of(b["runs"])
        if not keys or not KEY(t) or _in_toc(t, keys):       # в оглавлении ÜFK «1.1. BEVEZETÉS», в тексте «Bevezetés»
            return name
        own = name + " NoTOC"
        if own not in [s.name for s in self.doc.styles]:
            st = self.doc.styles.add_style(own, WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = self.doc.styles[name]
            st.next_paragraph_style = self.doc.styles["Normal"]
            ppr = st.element.get_or_add_pPr()
            for el in ppr.findall(W + "outlineLvl"):
                ppr.remove(el)
            ppr.append(_el("w:outlineLvl", w__val=9))
        return own

    def list_(self, b, level=0, tgt=None):
        from docx.shared import Pt
        nid = self._list_num(b["kind"], b.get("start", 1), level, b.get("suffix") or ".", bool(b.get("own")))
        tgt = tgt if tgt is not None else self.doc
        # свой список стоит иначе, чем вид списка по документу (больше 3 pt), — отступы абзацу, по оригиналу
        own = None
        geo = b.get("geo")
        if geo and tgt is self.doc and getattr(self, "left", None) is not None:
            m, t = geo["marker"] - self.left, geo["text"] - self.left
            am, at = self._list_ind(level, b["kind"])
            # тире может выступать за поле набора (скан 3SZ19: поле — по началу текста, тире на 4,75 pt левее) — у поля;
            # иначе пункт брал отступы профиля, и текст уезжал на 64 pt (19.10)
            if (abs(m - am) > 3 or abs(t - at) > 3) and m >= -15 and t > max(m, 0.0):
                own = (max(m, 0.0), t)
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
                # картинка и подпись «Рисунок N» внутри пункта — своими абзацами (по центру, как в оригинале), не
                # продолжением пункта с отступом (JPDS, 17.10)
                if not first and (all("img" in r for r in runs)
                                  or __import__("docx_struct").CAPTION_ANY.match(text_of(runs))):
                    self.paragraph(dict(x, runs=runs), tgt=tgt); continue
                par = tgt.add_paragraph(style="List UTS" if tgt is self.doc else "List Cell UTS")
                if first:
                    numpr = _el("w:numPr"); numpr.append(_el("w:ilvl", w__val=min(level, 3)))
                    numpr.append(_el("w:numId", w__val=nid))
                    par._p.get_or_add_pPr().insert(0, numpr)
                    # отступы — и самому абзацу: стиль пункта (отступ первой строки 0 при красной строке тела,
                    # build) перебивает висячий отступ нумерации — маркер вставал у поля вплотную к тексту (НГТУ, 19.10)
                    m, t = own or self._lvl_ind(min(level, 3), b["kind"])
                    par.paragraph_format.left_indent = Pt(t)
                    par.paragraph_format.first_line_indent = Pt(m - t)
                else:
                    par.paragraph_format.left_indent = Pt(own[1] if own else self._lvl_ind(min(level, 3), b["kind"])[1])
                self._put_runs(par, runs)
                for n in x.get("notes") or []:
                    self._footnote(par, n)
                first = False

    def _borders(self, t, color):
        """Линии таблицы цветом оригинала (служебные таблицы UTS — серые 808080, остальные — чёрные)."""
        b = _el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(_el(f"w:{side}", w__val="single", w__sz=4, w__space=0, w__color=color))
        _tblpr_put(t, b)

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
        _tblpr_put(t, m)

    def table(self, b):
        from docx.shared import Pt
        rows = [r for r in b["rows"] if r]
        if not rows:
            return
        # ячейки по сетке линий (docx_layout, «span»): столбцов — по сетке; ни в одной строке может не быть всех
        ncol = max(sum(c["span"] for c in r) for r in rows) if all("span" in c for r in rows for c in r) \
            else max(len(r) for r in rows)
        t = self.doc.add_table(rows=len(rows), cols=ncol)
        fr = b.get("frame") or {}
        # линии — у каждой ячейки свои, как в оригинале (docx_struct.cell_edges), если сетка там неполная
        edges = [c.get("edges") for r in rows for c in r]
        own_edges = b.get("borders", True) and edges and all(edges) and not all(all(e.values()) for e in edges)
        if b.get("borders", True):
            t.style = self.doc.styles["Table Grid"]
            if own_edges:
                self._no_borders(t)
            elif fr.get("line") and fr["line"] != "000000":
                self._borders(t, fr["line"])
            self._cell_margins(t, fr)
        else:
            self._no_borders(t)
        sizes = Counter(round(p["geom"]["size"]) for r in rows for c in r for p in c["blocks"] if p.get("geom"))
        small = bool(sizes) and sizes.most_common(1)[0][0] <= 9
        cell_style = "Table Text Small" if small else "Table Text"
        head_style = "Table Head Small" if small else "Table Head UTS"
        fit_size = 9 if small else 10
        tiny = sizes.most_common(1)[0][0] if sizes else None
        if tiny is not None and tiny <= 8:
            # таблица мельче 9 pt (параметры защит МТ-14 на 12 столбцов — 6–7 pt): кегль оригинала стилем, иначе в
            # узких столбцах слова рвались посередине («отключе|н», 18.10)
            from docx.shared import Pt
            from docx.enum.style import WD_STYLE_TYPE
            names = []
            for base in ("Table Text Small", "Table Head Small"):
                name = f"{base} {tiny:g}"
                if name not in [x.name for x in self.doc.styles]:
                    st = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                    st.base_style = self.doc.styles[base]
                    st.font.size = Pt(max(tiny, 5))
                names.append(name)
            cell_style, head_style = names
            fit_size = max(tiny, 5)
        if b.get("grid"):
            # сетка без линеек (docx_layout._grids): кегль — свой стиль, чтобы влезла (прямой кегль — нарушение
            # оформления); поля ячеек — по 1 pt, как в оригинале
            from docx.shared import Pt
            from docx.enum.style import WD_STYLE_TYPE
            if b.get("font_size"):
                name = f"Table Text {b['font_size']:g}"
                if name not in [x.name for x in self.doc.styles]:
                    st = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
                    st.base_style = self.doc.styles["Table Text"]
                    st.font.size = Pt(b["font_size"])
                cell_style = head_style = name
            mar = _el("w:tblCellMar")
            for side in ("left", "right"):
                mar.append(_el(f"w:{side}", w__w=20, w__type="dxa"))
            look = t._tbl.tblPr.find(W + "tblLook")          # порядок tblPr — по схеме: tblCellMar перед tblLook
            look.addprevious(mar) if look is not None else t._tbl.tblPr.append(mar)
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
            # картинка в ячейке — не шире своего столбца: иначе её срезает край ячейки (значки окна JPDS «□», «×» —
            # видна левая полоска, 17.10); недостающее — у самого просторного столбца
            need = [0.0] * ncol
            for r in rows:
                if len(r) != ncol:
                    continue
                for c, cell in enumerate(r):
                    for x in _paras(cell["blocks"]):
                        for run in x.get("runs") or []:
                            im = run.get("img")
                            if im and im.get("x1") is not None:
                                need[c] = max(need[c], float(im["x1"]) - float(im["x0"]) + 12)
            for c in range(ncol):
                if widths[c] < need[c]:
                    k = max(range(ncol), key=lambda i: widths[i] - need[i])
                    d = need[c] - widths[c]
                    if k != c and widths[k] - d >= max(need[k], 20):
                        widths[k] -= d
                        widths[c] = need[c]
            self._fit_words(rows, widths, fit_size)
            t.autofit = False
            for c in range(ncol):
                for cell in t.columns[c].cells:
                    cell.width = Pt(widths[c])
            grid = t._tbl.tblGrid
            for gc, w in zip(grid.findall(W + "gridCol"), widths):
                gc.set(W + "w", str(round(w * 20)))
            # поле слева и справа — не больше, чем оставляет место тексту в самом узком столбце (ÜFK II л. 49: 20
            # столбцов по 23 pt при полях 13,7 pt — числа не помещались и не печатались, 17.10)
            mar = t._tbl.tblPr.find(W + "tblCellMar")
            lim = max(1.0, (min(widths) - 12) / 2)
            for el in ([] if mar is None else [mar.find(W + "left"), mar.find(W + "right")]):
                if el is not None and int(el.get(W + "w")) / 20 > lim:
                    el.set(W + "w", str(round(lim * 20)))
        aligns = _col_align(rows, xs, (b.get("geom") or {}).get("page"))
        in_head = True
        for ri, row in enumerate(rows):
            if all("span" in c for c in row):               # ячейки по сетке линий (docx_layout) — места известны
                starts = [sum(c["span"] for c in row[:i]) for i in range(len(row))]
                spans = [(a, (starts[i + 1] - 1 if i + 1 < len(starts) else ncol - 1)) for i, a in enumerate(starts)]
            else:
                spans = _span_starts(row, bounds, ncol) if len(row) < ncol else [(i, i) for i in range(len(row))]
                starts = [a for a, _ in spans]
            for a, z in spans:                               # объединённые ячейки строки — по столбцам рамки
                if z > a:
                    t.cell(ri, a).merge(t.cell(ri, z))
            covered = {c for a, z in spans for c in range(a, z + 1)}
            for c in range(ncol):                            # ничей столбец — ячейка сверху тянется вниз
                if c not in covered and ri > 0:
                    try:
                        t.cell(ri - 1, c).merge(t.cell(ri, c))
                    except InvalidSpanError:                 # сверху объединение шире столбца — оставить пустой
                        pass
            in_head = in_head and all(c["th"] for c in row)
            head = in_head and ri < 3 and len(rows) > 1
            # повтор шапки на новом листе — как в оригинале: таблица, разрезанная листом без повтора шапки, — без него
            # (JPDS «Перечень сокращений»: «БД База данных» печаталась второй раз вверху следующего листа, 18.10)
            if head and len(rows) > 3 and (b.get("head_repeat") or not b.get("split")):
                _trpr_put(t.rows[ri]._tr, _el("w:tblHeader"))
            # короткая строка не рвётся между листами (в оригинале строки целые; длинная — пусть рвётся, иначе не влезет)
            size = sum(len(text_of(x["runs"])) for c in row for x in c["blocks"] if x["k"] == "p")
            if size < 600 and not any("img" in r for c in row for x in c["blocks"] if x["k"] == "p" for r in x["runs"]):
                _trpr_put(t.rows[ri]._tr, _el("w:cantSplit"))
            # пустая строка (графы для подписей листов ознакомления ИН-1065) — не ниже, чем в оригинале: без текста
            # она схлопывалась в 6 pt, лист ознакомления выходил вдвое короче (19.10); высота — по линиям таблицы
            # строки таблицы по линиям листа (docx_layout: строка — промежуток между линиями) — и с текстом: таблицы НГТУ
            # с просторными строками (31 pt при строке текста) сжимались вдвое
            ys = fr.get("ys") or []
            if (not size or fr.get("row_lines")) and len(ys) == len(rows) + 1 and ys[ri + 1] - ys[ri] > 6:
                # строка без текста и картинок — ровно по оригиналу: пустой абзац кеглем тела (14 pt, 1,5 интервала)
                # раздувал графы листа изменений ИН-1065 с 18 до 27 pt, хвост листа уходил на новый (19.10)
                blank = not size and not any("img" in r for c in row for x in c["blocks"] if x["k"] == "p"
                                             for r in x["runs"])
                _trpr_put(t.rows[ri]._tr, _el("w:trHeight", w__val=round((ys[ri + 1] - ys[ri]) * 20),
                                              w__hRule="exact" if blank else "atLeast"))
            for ci, cell in enumerate(row):
                dc = t.cell(ri, starts[ci])
                self._tab_base = min((float(x["geom"]["x0"]) for x in cell["blocks"]
                                      if (x.get("geom") or {}).get("x0") is not None),
                                     default=(cell.get("geom") or {}).get("x0"))
                if own_edges:
                    self._cell_borders(dc, cell["edges"], fr.get("line") or "000000")
                if cell.get("vcont"):                       # продолжение ячейки сверху (docx_struct.sub_rows)
                    continue
                first = dc.paragraphs[0]
                used = False
                style = head_style if head else cell_style
                al = aligns.get(("head" if head else "body", ci)) if len(row) == ncol else None
                if not head and ri == 0 and len(row) == ncol:      # первая строка — по себе, без своих голосов — как тело
                    al = aligns.get(("first", ci), al)
                    al = None if al == "left" else al
                g = cell.get("geom")
                if al is None and len(row) < ncol and len(bounds) == ncol + 1 and g and g.get("page") ==                         (b.get("geom") or {}).get("page"):
                    # объединённая ячейка (шапка ÜFK «Eltérő állapot» над двумя столбцами) — по своим границам
                    a0, z1 = bounds[starts[ci]], bounds[starts[ci + 1] if ci + 1 < len(starts) else ncol]
                    if abs((g["x0"] - a0) - (z1 - g["x1"])) < 6 and g["x0"] - a0 > 8:
                        al = "center"
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
                fill = fr.get("fill") if head and b.get("borders", True) else None
                fill = fill or cell.get("fill")                     # заливка ячейки по оригиналу (docx_struct)
                if fill:
                    dc._tc.get_or_add_tcPr().append(_el("w:shd", w__val="clear", w__color="auto", w__fill=fill))
            for ci, cell in enumerate(row):
                if cell.get("vcont") and ri > 0:
                    try:
                        t.cell(ri - 1, starts[ci]).merge(t.cell(ri, starts[ci]))
                    except InvalidSpanError:
                        pass
            # шапка (или первая строка) — со следующей строкой: «Контроль | Периодичность» оставалась одна внизу листа,
            # строки — на следующем (ÜFK I, перевод, л. 86 и 99; 19.10)
            if (head or ri == 0) and ri + 1 < len(rows):
                for tc in t.rows[ri]._tr.tc_lst:
                    for p in tc.iter(W + "p"):
                        p.get_or_add_pPr().get_or_add_keepNext()
        self._tab_base = None
        self.doc.add_paragraph(style="Table Text")        # Word: после таблицы — абзац

    def _cell_borders(self, dc, edges, color):
        tcpr = dc._tc.get_or_add_tcPr()
        el = _el("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            el.append(_el(f"w:{side}", w__val="single", w__sz=4, w__space=0, w__color=color) if edges.get(side)
                      else _el(f"w:{side}", w__val="nil"))
        nxt = next((c for c in tcpr if c.tag in {W + n for n in (
            "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark")}), None)
        nxt.addprevious(el) if nxt is not None else tcpr.append(el)

    def _fit_words(self, rows, widths, size):
        """Столбец не уже самого длинного слова перевода с полями ячейки: «Редакция» в узком столбце «Törés / verzió»
        рвалась «Редакци / я» (UTS), метка «A.1.» — «A. / 1.» (ÜFK, 16.10). Недостающее — у столбца с самым большим
        запасом; ширина таблицы та же."""
        get = getattr(self.tr, "get", None)
        if get is None:
            return
        # поля ячейки слева и справа (_cell_margins); у профиля без таблиц с текстом в рамке их нет (3PR42) — обычные
        pad = 2 * ((self.P.get("table") or {}).get("pad_left") or 5.4) + 5
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
                # позиция сразу за номером пункта («13.1.1» шире ближайшей табуляции — название уезжало к номеру
                # листа, а номер листа — на новую строку, UTS с переводом, 19.10); запас — под номер вида «13.1.1»
                # шрифтом оглавления: 7 pt на знак и 6 pt зазора. Висячий отступ на ту же позицию: вторая строка
                # длинного пункта — под текстом, не под номером (OAH 1.1)
                base, num = (ind[i] if i < len(ind) else 0), 7 * (2 * i + 2) + 6
                pf.left_indent, pf.first_line_indent = Pt(base + num), Pt(-num)
                pf.tab_stops.add_tab_stop(Pt(base + num))
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

    def _shrink_tail(self):
        """Перед разрывом листа: пустой абзац Word после таблицы — в 1 pt. Таблица кончалась у низа листа, этот абзац
        уходил на следующий лист один, а за ним текст «с новой страницы» — пустой лист (ÜFK I л. 75, 18.10)."""
        from docx.shared import Pt
        body = self.doc.element.body
        last = next((e for e in reversed(body) if e.tag != W + "sectPr"), None)
        prev = last.getprevious() if last is not None else None
        if last is not None and last.tag == W + "p" and not "".join(last.itertext()).strip()                 and prev is not None and prev.tag == W + "tbl":
            par = self.doc.paragraphs[-1]
            if par._p is last:
                par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.line_spacing = Pt(1)

    def block(self, b):
        k = b["k"]
        try:
            self._block(b)
        finally:
            self._last_k = k                                # вид прошлого блока (вводная фраза под заголовком)

    def _block(self, b):
        k = b["k"]
        fresh, self.fresh = getattr(self, "fresh", False), False
        if b.get("break") and not fresh:
            self._shrink_tail()
        if k == "h":
            self.heading(b)
        elif k == "p":
            par = self.paragraph(b)
            # разрыв листа оригинала (docx_struct.load); первый абзац нового раздела и так с нового листа
            if b.get("break") and par is not None and not fresh:
                par.paragraph_format.page_break_before = True
        elif k == "list":
            self.list_(b, b.get("level", 0))                # уровень по отступу оригинала (docx_layout)
        elif k == "table":
            if b.get("break") and not fresh:
                # разрыв перед таблицей — пустым абзацем в 1 pt «с новой страницы»: у таблицы своего разрыва нет
                from docx.shared import Pt
                par = self.doc.add_paragraph(style="Table Text")
                par.paragraph_format.page_break_before = True
                par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.line_spacing = Pt(1)
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

    def _num_start(self, sec, value):
        """Нумерация листов раздела — с числа оригинала (pgNumType; место в sectPr — по схеме, иначе Word ругается)."""
        if value is None or value < 0:
            return
        sp = sec._sectPr
        for el in sp.findall(W + "pgNumType"):
            sp.remove(el)
        el = _el("w:pgNumType", w__start=int(value))
        nxt = next((c for c in sp if c.tag in {W + n for n in (
            "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
            "printerSettings", "sectPrChange")}), None)
        nxt.addprevious(el) if nxt is not None else sp.append(el)

    def _no_borders(self, table):
        b = _el("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b.append(_el(f"w:{side}", w__val="nil"))
        _tblpr_put(table, b)

    def _zone_table(self, part, grid, style, page_tok=None, page_marked=False):
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
            _tblpr_put(t, b)
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
        merge = grid.get("merge")
        for ri, row in enumerate(rows):
            for ci, txt in enumerate(row["cells"]):
                cell = t.cell(ri, ci)
                p = cell.paragraphs[0]
                p.style = self.doc.styles[style]
                if isinstance(txt, dict):                   # ячейка по линиям (docx_struct._rule_grid): строки, картинки
                    for im in (grid.get("imgs") or {}).get((ri, ci), []):
                        data = self.image({"stream": type("S", (), {"objid": im["xref"]})(), **im})
                        if data:                            # не шире ячейки с полями — иначе логотип обрезан
                            w = im["x1"] - im["x0"]
                            if ci < len(widths):
                                w = min(w, widths[ci] - 12)
                            p.add_run().add_picture(io.BytesIO(data), width=Pt(w))
                    for li, (line, bold) in enumerate(txt["lines"]):
                        q = p if li == 0 and not p.runs else cell.add_paragraph(style=self.doc.styles[style])
                        q.alignment = al_of.get(txt.get("align"), AL.LEFT)
                        self._zone_text(q, line, bold, page_tok, page_marked)
                    if not txt["lines"]:
                        p.alignment = al_of.get(txt.get("align"), AL.LEFT)
                    continue
                for li, line in enumerate(txt.split("\n")):     # перенос в ячейке (docx_struct._zone_grid)
                    q = p if li == 0 else cell.add_paragraph(style=self.doc.styles[style])
                    if ci < len(aligns) and aligns[ci] in al_of:
                        q.alignment = al_of[aligns[ci]]
                    self._zone_text(q, line, row["bold"], page_tok, page_marked)
        if merge:                                           # ячейки без линии между строками — одна
            for ci in range(ncol):
                ri = 0
                while ri < len(rows):
                    z = ri
                    while z + 1 < len(rows) and merge[z + 1][ci]:
                        z += 1
                    if z > ri:
                        t.cell(ri, ci).merge(t.cell(z, ci))
                    ri = z + 1
        return t

    def _zone_text(self, p, txt, bold, page_tok, page_marked=False):
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
        if PAGE_MARK in txt:                    # номер помечен на месте (docx_struct.layout): поле — он, k-й такой же
            k = len(num.findall(txt[:txt.index(PAGE_MARK)])) if num else 0
            txt = re.sub(PAGE_MARK + "+", tok, txt)
            ru = self.tr(txt) or txt
            ms = list(num.finditer(ru)) if num else []
            mm = (ms[k] if len(ms) == len(num.findall(txt)) else ms[-1]) if ms else None
        else:
            ru = (self.tr(txt) or txt) if txt else ""
            mm = num.search(ru) if num and num.search(txt) and not page_marked else None
        if mm:
            r = p.add_run(ru[:mm.start()]); r.bold = bold
            f = _field(p, "PAGE", tok); f.bold = bold
            rest = ru[mm.end():]
            m2 = re.match(r"(\s*(?:/|of|из|von|de|di|z)\s*)(\d+)", rest, re.I)   # «20/57», «4 of 13» — и число листов полем
            if m2:
                r = p.add_run(m2.group(1)); r.bold = bold
                f = _field(p, "NUMPAGES", m2.group(2)); f.bold = bold
                rest = rest[m2.end():]
            r = p.add_run(rest); r.bold = bold
        else:
            r = p.add_run(ru); r.bold = bold

    def _zone_paras(self, part, zp, style, page_tok, page_marked):
        """Колонтитул без рамки (docx_struct._zone_paras) — абзацами по строкам оригинала: одна часть — выравниванием,
        несколько — табуляциями по местам (по центру / у правого поля / слева); линии — границами абзаца;
        картинка (логотип) — своим абзацем по месту."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL, WD_TAB_ALIGNMENT as TA
        al = {"center": AL.CENTER, "right": AL.RIGHT}
        for i, row in enumerate(zp["rows"]):
            p = part.paragraphs[0] if i == 0 else part.add_paragraph()
            p.style = self.doc.styles[style]
            pf = p.paragraph_format
            pf.left_indent = pf.first_line_indent = Pt(0)
            pf.space_after = Pt(0)
            if row.get("img"):
                im = row["img"]
                p.alignment = al.get(row["kind"], AL.LEFT)
                data = self.image({"stream": type("S", (), {"objid": im["xref"]})(), **im})
                if data:
                    p.add_run().add_picture(io.BytesIO(data), width=Pt(im["x1"] - im["x0"]))
            else:
                parts = row["parts"]
                if len(parts) == 1 and not (parts[0][1] == "left" and parts[0][2] > 6):
                    p.alignment = al.get(parts[0][1], AL.LEFT)
                    self._zone_text(p, parts[0][0], row.get("bold"), page_tok, page_marked)
                else:
                    p.alignment = AL.LEFT
                    for j, (txt, kind, pos) in enumerate(parts):
                        if j or kind != "left" or pos > 6:
                            pf.tab_stops.add_tab_stop(Pt(pos), {"center": TA.CENTER, "right": TA.RIGHT}.get(kind, TA.LEFT))
                            p.add_run("\t")
                        self._zone_text(p, txt, row.get("bold"), page_tok, page_marked)
                if (row.get("color") or "000000").lower() not in ("000000", "ffffff"):   # цвет строки — как в оригинале
                    from docx.shared import RGBColor
                    for r in p.runs:
                        r.font.color.rgb = RGBColor.from_string(row["color"].upper())
            for key, side in (("rule_above", "top"), ("rule_below", "bottom")):
                if row.get(key):
                    ppr = p._p.get_or_add_pPr()
                    bdr = ppr.find(W + "pBdr")
                    if bdr is None:
                        bdr = _el("w:pBdr")
                        ppr.insert_element_before(bdr, *PPR_AFTER_PBDR)
                    bdr.append(_el(f"w:{side}", w__val="single", w__sz=4, w__space=1, w__color="000000"))

    def _header(self, sec, L, logos_only=False):
        """Шапка: логотипы по краям, строки по центру (название и номер документа), линия снизу; шапка без логотипов
        (строка-таблица ÜFK) — таблицей по сетке оригинала."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
        hdr = sec.header
        hdr.is_linked_to_previous = True                     # новый раздел копирует ссылку на шапку титула —
        hdr.is_linked_to_previous = False                    # своя часть, иначе обе шапки пишутся в одну
        grid = L.get("head_grid")
        if grid and grid.get("paras"):
            if not logos_only:
                self._zone_paras(hdr, grid["paras"], "Header UTS", L.get("page_tok"), L.get("page_marked"))
            return
        if grid and grid.get("rows"):
            if not logos_only:
                t = self._zone_table(hdr, grid, "Header UTS", L.get("page_tok"), L.get("page_marked"))
                hdr.paragraphs[0]._p.addprevious(t._tbl)
                hdr.paragraphs[0].paragraph_format.space_after = Pt(0)
            return
        if not L.get("head_imgs") and not L.get("head_lines"):
            return                                  # шапки нет — и таблицы-пустышки нет
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
            c.width = Pt(w)                                  # и сетка таблицы: LibreOffice берёт ширину из неё
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
            hl = L["head_lines"]
            # строки начинаются с одного края и разной длины — выключка влево (НГТУ: три строки справа от герба
            # печатались по центру столбиком); иначе — по центру, как у UTS
            flush = len(hl) > 1 and all("x0" in ln for ln in hl) and max(ln["x0"] for ln in hl) - min(
                ln["x0"] for ln in hl) < 3 and max(ln["x1"] for ln in hl) - min(ln["x1"] for ln in hl) > 6
            for i, ln in enumerate(hl):
                p = mid.paragraphs[0] if i == 0 else mid.add_paragraph()
                p.style = self.doc.styles["Header UTS"]
                if flush:
                    p.alignment = AL.LEFT
                    p.paragraph_format.left_indent = Pt(max(0.0, ln["x0"] - L["left"] - wl))
                    if i:                                    # промежуток между строками — как в оригинале
                        p.paragraph_format.space_before = Pt(max(0.0, ln["top"] - hl[i - 1]["bottom"] - 0.2 * ln["size"]))
                elif i:
                    p.paragraph_format.space_before = Pt(8)
                p.add_run(self.tr(ln["t"]) or ln["t"])
        for p in ([] if logos_only or not L.get("head_rule", True) else list(hdr.paragraphs)):   # пустой абзац — под линию
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
        if L.get("stamp_paras"):
            self._zone_paras(ftr, L["stamp_paras"], "Stamp UTS", L.get("page_tok"), L.get("page_marked"))
            return
        if not L["stamp"]:
            return
        self._zone_table(ftr, {"rows": L["stamp"], "widths": L["stamp_widths"], "aligns": L.get("stamp_aligns"),
                               "inner": L.get("stamp_inner", True), "box": L.get("stamp_box", True),
                               "merge": L.get("stamp_merge")},
                         "Stamp UTS", L.get("page_tok"), L.get("page_marked"))
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
        if self.P.get("first_line"):            # отступ первой строки — только у основного текста (docx_profile)
            from docx.enum.style import WD_STYLE_TYPE
            for st in self.doc.styles:
                if st.type == WD_STYLE_TYPE.PARAGRAPH and st.name != "Normal" and \
                        st.paragraph_format.first_line_indent is None:
                    st.paragraph_format.first_line_indent = Pt(0)
            self.doc.styles["Normal"].paragraph_format.first_line_indent = Pt(self.P["first_line"])
        tocs = [b for b in blocks if b["k"] == "toc" and not ((b.get("info") or {}).get("figures"))]
        self.toc_keys = [k for t in tocs for k in (t.get("info") or {}).get("keys") or []] or None
        if self.toc_keys:
            # сверке верим, если пункты оглавления почти все нашлись среди заголовков; иначе (оглавление прочитано не
            # так) заголовки все в оглавлении, как раньше — поле TOC не пустеет. По доле совпавших заголовков мерить
            # нельзя: у инструкции ИН-АСУТП-1172 обычные абзацы оформлены стилем заголовка 2 (в её оглавлении их нет),
            # заголовков втрое больше пунктов, и сверка выключалась — абзацы шли в оглавление (16.10)
            hs = [text_of(b["runs"]) for b in blocks if b["k"] == "h" and KEY(text_of(b["runs"]))]
            found = sum(any(_in_toc(h, [k]) for h in hs) for k in self.toc_keys if k)
            if found < 0.6 * len([k for k in self.toc_keys if k]):
                self.toc_keys = None
        self.autonum = _autonum_ok(blocks)
        for l in self.doc.part.numbering_part.element.iter(W + "lvl"):    # точка после номера — как в оригинале
            if l.getparent().get(W + "abstractNumId") == "900" and int(l.get(W + "ilvl")) + 1 in _head_dots(blocks):
                t = l.find(W + "lvlText")
                if not t.get(W + "val").endswith("."):
                    t.set(W + "val", t.get(W + "val") + ".")
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
        prev_bottom = L["top"]                  # от верхнего поля: название ÜFK стоит с середины листа
        # почти пустой лист вне титула (выходные данные OAH на обороте титула: «Kiadta» на четверти листа, «A kiadvány
        # beszerezhető» — на двух третях) — большие пустоты тоже по оригиналу; было — всё сжато наверх (17.10)
        # только абзацы и заголовки: лист с таблицей, списком или рисунком не пустой (их текст здесь не считан —
        # отступы в 450 pt перед заголовками под таблицами, N3a 56 → 61 лист)
        chars, busy = Counter(), set()
        for b in blocks:
            g0 = b.get("geom") or (b["items"][0][0].get("geom") if b["k"] == "list" and b.get("items") and b["items"][0]
                                   else None) or {}
            if not g0.get("page"):
                continue
            if b["k"] in ("p", "h") and not all("img" in r for r in b.get("runs") or [{"img": 1}]):
                chars[g0["page"]] += len(text_of(b.get("runs") or []).strip())
            else:
                busy.add(g0["page"])
        sparse = {pg for pg, n in chars.items() if n < 400 and pg >= 2 and pg not in busy} & set(L.get("sparse") or ())
        sparse_bottom = {}
        # первый лист с колонтитулами тела (3SZ19 начинается оглавлением с шапкой) — не титул (docx_layout)
        in_title = had_title = not L.get("first_is_body")
        # поля титула — как у тела, места строк — отступами и центром от них: поле по самому частому краю строк
        # титула (центрованных) сжимало название OAH в столбец в полстраницы (16.10)
        self._page(first, L, L["left"], L["right"], L["top"], L["bottom"])
        self.left = L["left"]
        if not in_title:
            self._header(first, L)
            self._footer(first, L)
            if L.get("first"):                  # особый первый лист (ONS: без подвала) — «особый колонтитул» Word
                first.different_first_page_header_footer = True
                fp = type("First", (), {"header": first.first_page_header, "footer": first.first_page_footer})()
                self._header(fp, dict(L, **L["first"]))
                self._footer(fp, dict(L, **L["first"]))
        # логотипы титула — те же, что в шапке тела (UTS), — шапкой титула; иные картинки титула — на листе (OAH)
        head_x = {i["xref"] for i in L.get("head_imgs") or []}
        logos_head = False
        blocks = merge_inline(blocks)
        if in_title:                            # титул — сверху вниз по листу (логотип OAH в дереве — после названия)
            n1 = next((i for i, b in enumerate(blocks) if ((b.get("geom") or {}).get("page") or 1) >= 2), len(blocks))
            blocks = sorted(blocks[:n1], key=lambda b: (b.get("geom") or {}).get("top") or 0) + blocks[n1:]
        # вводная часть (оглавление, лист изменений) — свой подвал без номера; тело — нумерация с числа оригинала
        front, ns = L.get("front"), L.get("num_start")
        # номер листа во вводной части — только если её шапка та же, что у тела (синтетика: «2» в шапке оглавления
        # пропадал); своя шапка или подвал без номера (ÜFK «Tartalomjegyzék») — без номера, как было
        Lf = dict(L, **{k: v for k, v in (front or {}).items() if k != "to"},
                  page_tok=None if front and "head_grid" in front else L.get("page_tok")) if front else None
        in_front = False
        if not in_title and ns:
            self._num_start(first, ns["value"])
        title_foot = []
        for b in blocks:
            g = b.get("geom") or {}
            page = g.get("page")
            if in_title and page is not None and page >= 2:
                in_title = False
                sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
                self.fresh = True
                self._page(sec, L, L["left"], L["right"], L["top"], L["bottom"])
                self.left = L["left"]
                in_front = bool(front) and page <= front["to"]
                self._header(sec, Lf if in_front else L)
                self._footer(sec, Lf if in_front else L)
                if ns and not in_front:
                    self._num_start(sec, ns["value"])
            if in_front and page is not None and page > front["to"]:
                in_front = False
                sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
                self.fresh = True
                self._page(sec, L, L["left"], L["right"], L["top"], L["bottom"])
                self._header(sec, L)
                self._footer(sec, L)
                if ns:
                    self._num_start(sec, ns["value"])
            # логотипы обложки — те же, что в шапке тела, или стоящие в полосе шапки (выше поля тела; «atom erőmű» UTS
            # уезжал за правый край), — шапкой титула из логотипов тела
            if in_title and b["k"] == "p" and b["runs"] and all("img" in r for r in b["runs"]) and head_x and (all(
                    getattr(r["img"].get("stream"), "objid", None) in head_x for r in b["runs"])
                    or (g.get("bottom") or 999) <= L["top"] + 20):
                prev_bottom = max(prev_bottom, g.get("bottom") or 0)   # отступ текста титула — от низа логотипов
                logos_head = True
                continue
            if in_title and b["k"] == "p" and b["runs"] and all("img" in r for r in b["runs"]) and g.get("x1") \
                    and (g["x1"] - g["x0"]) * ((g.get("bottom") or g["top"]) - g["top"]) > 0.4 * L["W"] * L["H"]:
                continue                                     # фон обложки (UTS) — не картинкой в потоке
            # юридический текст внизу титула — в подвал титула; только ниже поля тела (UTS: 780 pt при поле 728),
            # дата титула выше поля — в потоке (OAH 4.22: «2015. szeptember» на 681 pt уходила в подвал к логотипу)
            if in_title and b["k"] == "p" and g.get("top", 0) > max(0.8 * L["H"], L["H"] - L["bottom"] - 2):
                title_foot.append(b)
                continue
            self.block(b)
            pg = g.get("page")
            if not in_title and pg in sparse and b["k"] in ("p", "h") and g.get("top") and self.doc.paragraphs:
                gap = g["top"] - sparse_bottom.get(pg, L["top"])
                if gap > 36:
                    self.doc.paragraphs[-1].paragraph_format.space_before = Pt(0.85 * gap)
                sparse_bottom[pg] = g.get("bottom") or g["top"]
            # повёрнутый текст (вдоль края титула JPDS) — абзацем без места: центр по его середине у края листа давал
            # правый отступ 450 pt и столбик по букве (16.10)
            if in_title and b["k"] == "p" and g and self.doc.paragraphs and not g.get("rotated"):
                par = self.doc.paragraphs[-1]
                if par.alignment == 1 or (par.style.paragraph_format.alignment or 0) == 1:
                    # по центру — центр там же, где в оригинале
                    a = 2 * (g["x0"] + g["x1"]) / 2 - L["left"] - (L["W"] - L["right"])
                    if a > 3:
                        par.paragraph_format.left_indent = Pt(a)
                    elif a < -3:
                        par.paragraph_format.right_indent = Pt(-a)
                elif all("img" in r for r in b["runs"]):            # картинка титула — по месту
                    if abs((g["x0"] + g["x1"]) / 2 - L["W"] / 2) < 15 or abs(
                            2 * (g["x0"] + g["x1"]) / 2 - L["left"] - (L["W"] - L["right"])) < 15:
                        par.alignment = 1
                    else:
                        par.paragraph_format.left_indent = Pt(max(0, g["x0"] - L["left"]))
                else:
                    dx = g["x0"] - L["left"]
                    if abs(dx) > 3:
                        par.paragraph_format.left_indent = Pt(dx)
                # большие пустоты титула (название посреди листа, юридический текст внизу) — как в оригинале
                if prev_bottom is not None and g.get("top") and g["top"] - prev_bottom > 24:
                    par.paragraph_format.space_before = Pt(0.85 * (g["top"] - prev_bottom))
                prev_bottom = g.get("bottom") or prev_bottom
            elif in_title and b["k"] == "table" and g.get("bottom"):
                prev_bottom = g["bottom"]            # пустота под таблицей титула — от её низа (Eskom: титул на 2 листа)
        # после add_section объект первого раздела указывает на последний (sectPr тела документа) — шапку титула
        # ставить разделу [0], иначе шапки менялись местами
        # колонтитулы обложки — как у тела, но спрятаны под фоном (UTS): шапка титула — логотипы тела
        if had_title and (logos_head or not L.get("first") and L.get("head_imgs")):
            self._header(self.doc.sections[0], L, logos_only=True)
        elif had_title and L.get("first"):      # свои колонтитулы титула (NSU: шапка с названием вуза)
            self._header(self.doc.sections[0], dict(L, **L["first"]))
            self._footer(self.doc.sections[0], dict(L, **L["first"]))
        if title_foot:
            foot = self.doc.sections[0].footer
            foot.is_linked_to_previous = False
            # пустой первый абзац подвала — прочь, но не абзац с логотипом (подвал титула OAH, 16.10)
            first_p = foot.paragraphs[0] if foot.paragraphs and not foot.paragraphs[0].text \
                and not foot.paragraphs[0]._p.findall(".//" + W + "drawing") else None
            for b in title_foot:
                par = self.paragraph(b, tgt=foot)
                g = b.get("geom") or {}
                if par is not None and g.get("x1") and abs(2 * (g["x0"] + g["x1"]) / 2 - L["left"]
                                                           - (L["W"] - L["right"])) < 15:
                    par.alignment = 1                        # по центру — как в оригинале («2015. szeptember»)
            if first_p is not None:
                first_p._p.getparent().remove(first_p._p)
        self._tab_groups()
        self.doc.settings.element.append(_el("w:updateFields", w__val="true"))
        self.doc.save(out)
        font_table(out, (L or {}).get("font_class") or {})

    def _tab_groups(self):
        """Пары «подпись ⇥ значение» подряд с одной табуляцией (титул: «Műszaki változás száma⇥10776»): позиция —
        общая и не ближе самой длинной русской подписи, иначе значение длинной подписи уезжало на следующую
        табуляцию («Номер технического изменения 10776» не в столбце, UTS 16.10). Ширина — латиница по метрике
        Helvetica, кириллица и прочее — средней шириной Tahoma (у Helvetica PDF нет кириллицы, шрифтов Word на
        сервере нет), с запасом."""
        from docx.shared import Pt
        normal = (self.doc.styles["Normal"].font.size or Pt(10)).pt
        pars = self.doc.paragraphs

        from docx.enum.text import WD_TAB_ALIGNMENT

        def stop(p):
            # правая табуляция (номер листа в строке оглавления, docx_layout) — не «подпись ⇥ значение»: её уводило
            # за край листа по длине самого длинного названия (3SZ19, 716 pt)
            ts = p.paragraph_format.tab_stops
            return ts[0].position if len(ts) == 1 and "\t" in p.text and ts[0].alignment != WD_TAB_ALIGNMENT.RIGHT \
                else None

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
            # «подпись» длиннее строки — не пара «подпись ⇥ значение», а промежуток в тексте (место значка ESC в абзаце
            # ИН-1065: табуляция уезжала на 827 pt, хвост абзаца — за край листа, 17.10)
            width = (self.L["W"] - self.L["left"] - self.L["right"]) if getattr(self, "L", None) else 450
            if pos.pt < need <= 0.75 * width:
                for p in pars[i:j]:
                    p.paragraph_format.tab_stops.clear_all()
                    p.paragraph_format.tab_stops.add_tab_stop(Pt(round(need)))
            i = j


# ---------------------------------------------------------------- в конвейере (pipeline.extract, struct_fill; assemble)
STRUCT_COVER = 0.9


PPR_AFTER_PBDR = tuple("w:" + t for t in (
    "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
    "rPr", "sectPr", "pPrChange"))                        # что идёт в pPr после pBdr (порядок схемы)
TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
         "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
         "tblDescription"]


def _tblpr_put(table, el):
    """Свойство таблицы — на своё место в tblPr (порядок по схеме; строгие проверки ругались на tblBorders после
    tblLook); прежнее такое же — прочь."""
    tp = table._tbl.tblPr
    for old in tp.findall(el.tag):
        tp.remove(old)
    name = el.tag.split("}")[1]
    later = {W + n for n in TBLPR[TBLPR.index(name) + 1:]}
    nxt = next((c for c in tp if c.tag in later), None)
    nxt.addprevious(el) if nxt is not None else tp.append(el)


_NUMHEAD = re.compile(r"^[A-Za-zÁÉÍÓÖŐÚÜŰ]{0,4}\d+(?:\.\d+)*\.?")


def _in_toc(text, keys):
    """Заголовок есть в оглавлении оригинала: сравнение без пробелов и регистра, с номером раздела и без него
    («1.1. BEVEZETÉS» в оглавлении — «Bevezetés» в тексте, когда номер даёт автонумерация)."""
    k = KEY(text).casefold()
    variants = {k[:30], _NUMHEAD.sub("", k)[:30]} - {""}
    for e in keys:
        e = e.casefold()
        for ev in {e[:30], _NUMHEAD.sub("", e)[:30]} - {""}:
            if any(ev.startswith(v) or v.startswith(ev) for v in variants):
                return True
    return False


def _paras(blocks):
    """Абзацы и заголовки в порядке чтения, со вложенными списками и ячейками таблиц, и тексты их сносок."""
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
            for n in b.get("notes") or []:
                yield {"k": "p", "runs": n["runs"]}


def blocks_layout(pdf, segs, mode):
    """Блоки и раскладка листа по виду PDF: "struct" — по дереву разметки (docx_struct); "layout" — распознанный
    (контуры): по сегментам OCR и раскладка по ним (docx_layout); "text" — текстовый слой без разметки: блоки по
    сегментам, начертание по шрифтам слоя, колонтитулы — правилами размеченного по знакам вне текста тела."""
    import docx_struct
    if mode in ("layout", "text"):
        import docx_layout
        blocks, info = docx_layout.load(str(pdf), segs, text=mode == "text")
        if mode == "text":
            return blocks, dict(docx_struct.layout(str(pdf), blocks, content=info.get("content_boxes")), sizes=True)
        return blocks, docx_layout.layout(str(pdf), blocks, segs, info)
    blocks, _ = docx_struct.load(str(pdf))
    blocks = docx_struct.compose_figures(blocks, str(pdf))   # не в load: load зовёт и coverage при извлечении
    docx_struct.list_geo(blocks)
    return blocks, dict(docx_struct.layout(str(pdf), blocks), sizes=True)    # кегль — из знаков PDF, точный


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
    # и Word 2007 — «Microsoft® Office Word 2007»: прежнее «Microsoft.{0,4}Word» его не узнавало, и руководства OAH
    # шли переводом поверх PDF (живая проверка, 16.10)
    if not re.search(r"Microsoft\W{0,3}(?:Office\s+)?Word", (meta.get("producer") or "") + " " + (meta.get("creator") or ""),
                     re.I):
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


def missing(pdf, segs, res, mode="struct"):
    """Тексты, которые сборка DOCX по разметке спросит у памяти задачи и не найдёт: [(текст, стр. от 0 или None,
    номера моноширинных групп)] в порядке документа. Абзац с метками начертания и он же без меток (запасной ход
    сборки) — один раз, первым. mode="layout" — распознанный PDF (docx_layout): абзацы с полужирными словами."""
    import tempfile
    from pathlib import Path
    blocks, L = blocks_layout(pdf, segs, mode)
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
