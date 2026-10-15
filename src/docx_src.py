# -*- coding: utf-8 -*-
"""DOCX для PDF с текстовым слоем — из оригинала, а не из переведённого PDF.

Прежний путь (pdf2docx по переведённому PDF) давал в Word чужие шрифты (Nimbus Sans, Charis SIL — их встраивает
insert_htmlbox), ужатый кегль из тесных ячеек PDF и разваленные сканы-страницы (UTS, 15.10: 107 стр. вместо 99,
всё засечками). Теперь:
- извлечение (plan): pdf2docx по оригиналу → абзацы DOCX сопоставляются с сегментами PDF той же страницы
  (точно; склейкой подряд идущих сегментов; частями одного сегмента — перевод целиком в первую часть).
  Несопоставленные абзацы становятся своими сегментами (loc.docx_only) и переводятся как все;
- сборка (build): pdf2docx по оригиналу с переведёнными скриншотами → переводы в прогоны (assemble.build_docx) →
  шрифты, которых Word не знает (Symbol для «•», TimesNewRomanPSMT), заменяются.
В Word таблица настоящая: строка растёт под русский текст, кегль не ужимается."""
import re, zipfile, tempfile
from pathlib import Path

import segments as SG

TAGS = re.compile(r"</?f\d+>")
GROUP = re.compile(r"<f(\d+)>(.*?)</f\1>", re.S)
KEY = lambda t: re.sub(r"\s+", "", TAGS.sub("", t or ""))
FONTS = {"Symbol": "Arial", "TimesNewRomanPS": "Times New Roman", "TimesNewRomanPSMT": "Times New Roman",
         "TimesNewRomanPS-BoldMT": "Times New Roman", "ArialMT": "Arial", "Arial-BoldMT": "Arial",
         "CourierNewPSMT": "Courier New"}


PPR_AFTER_SPACING = ("ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
                     "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")
BULLET_T = re.compile(r"^[•◦▪o\-–]$")


def normalize(path):
    """Правка DOCX от pdf2docx до сопоставления (одинаково при извлечении и при сборке):
    - абзац из нескольких пунктов списка («•⇥текст↵•⇥текст…») или строк оглавления — на абзацы по пункту: иначе перевод уходит в первый
      прогон, а маркеры остаются в конце абзаца (UTS стр. 26, 15.10); пункту — висячий отступ по табуляции;
    - точная высота строк таблицы — «не меньше»: русский текст длиннее, Word обрезал строки таблицы."""
    import copy
    from lxml import etree
    W = SG.W
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    is_bullet = lambda r: r is not None and r.tag == W + "r" and BULLET_T.match("".join(t.text or "" for t in r.iter(W + "t")))
    has_br = lambda r: r.tag == W + "r" and r.find(W + "br") is not None and not "".join(t.text or "" for t in r.iter(W + "t"))
    for p in list(root.iter(W + "p")):
        kids = list(p)
        # оглавление: pdf2docx собирает весь блок страницы в один абзац с переносами — по строке на абзац, тогда
        # строка сопоставляется с сегментом PDF и получает тот же перевод («Рисунок 1.», а не «1. рис.»)
        toc = len(re.findall(r"\.{4,}\s*\d+", "".join(t.text or "" for t in p.iter(W + "t")))) >= 2
        cuts = [i for i in range(len(kids) - 1) if has_br(kids[i]) and (toc or is_bullet(kids[i + 1]))]
        if not cuts:
            continue
        ppr = p.find(W + "pPr")
        sect = ppr.find(W + "sectPr") if ppr is not None else None
        if sect is not None:
            ppr.remove(sect)
        prev, bounds = p, cuts + [len(kids)]
        for a, b in zip(bounds, bounds[1:]):
            q = etree.Element(W + "p")
            if ppr is not None:
                q.append(copy.deepcopy(ppr))
                # внутри исходного абзаца отступа «перед» между строками не было — у кусков его нет
                # (оглавление UTS с отступом у каждой строки растягивалось на лишние листы)
                sp = q.find(f"{W}pPr/{W}spacing")
                if sp is not None:
                    sp.set(W + "before", "0")
            for el in kids[a + 1:b]:                 # перенос строки перед маркером уходит
                q.append(el)
            p.remove(kids[a])
            prev.addnext(q); prev = q
        # и отступа «после» между строками не было — он остаётся только у последнего куска
        piece = p
        while piece is not prev:
            sp = piece.find(f"{W}pPr/{W}spacing")
            if sp is not None:
                sp.set(W + "after", "0")
            piece = piece.getnext()
        if sect is not None:
            last_ppr = prev.find(W + "pPr")
            if last_ppr is None:
                last_ppr = etree.SubElement(prev, W + "pPr"); prev.insert(0, last_ppr)
            last_ppr.append(sect)
    for p in root.iter(W + "p"):                      # висячий отступ у пункта «•⇥текст»
        runs = [r for r in p if r.tag == W + "r"]
        ppr = p.find(W + "pPr")
        if len(runs) < 2 or ppr is None or not is_bullet(runs[0]) or runs[1].find(W + "tab") is None:
            continue
        tab, ind = ppr.find(f"{W}tabs/{W}tab"), ppr.find(W + "ind")
        if tab is None or ind is None:
            continue
        pos, left = int(tab.get(W + "pos", "0")), int(ind.get(W + "left", "0"))
        if pos > left:
            ind.set(W + "left", str(pos)); ind.set(W + "hanging", str(pos - left))
            ind.attrib.pop(W + "firstLine", None)
    for h in root.iter(W + "trHeight"):
        if h.get(W + "hRule") == "exact":
            h.set(W + "hRule", "atLeast")
    # пустой абзац с разрывом раздела (конец каждой страницы pdf2docx) — высотой в 1 pt: на заполненной странице
    # он выпадал на новый лист, и оставалась пустая страница с одним колонтитулом
    for p in root.iter(W + "p"):
        ppr = p.find(W + "pPr")
        if ppr is None or ppr.find(W + "sectPr") is None or "".join(t.text or "" for t in p.iter(W + "t")).strip():
            continue
        for tag in ("spacing", "rPr"):
            el = ppr.find(W + tag)
            if el is not None:
                ppr.remove(el)
        sp = etree.Element(W + "spacing"); sp.set(W + "before", "0"); sp.set(W + "after", "0")
        sp.set(W + "line", "20"); sp.set(W + "lineRule", "exact")
        rpr = etree.Element(W + "rPr"); etree.SubElement(rpr, W + "sz").set(W + "val", "2")
        # порядок детей pPr по схеме: spacing — до ind/jc/…, rPr — перед sectPr
        after = [x for x in ppr if x.tag in {W + t for t in PPR_AFTER_SPACING}]
        (after[0].addprevious if after else ppr.append)(sp)
        ppr.find(W + "sectPr").addprevious(rpr)
    with zipfile.ZipFile(path) as z:
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        types = etree.fromstring(z.read("[Content_Types].xml"))
    # вертикальные отступы pdf2docx повторяют позиции строк PDF; русский текст длиннее, и хвост страницы
    # переливался на следующий лист — отступ перед абзацем ×0,6, межстрочный «авто» ×0,92 (не меньше одинарного)
    for sp in root.iter(W + "spacing"):
        if sp.get(W + "before"):
            sp.set(W + "before", str(int(float(sp.get(W + "before")) * 0.6)))
        if sp.get(W + "lineRule") == "auto" and sp.get(W + "line"):
            sp.set(W + "line", str(max(240, int(float(sp.get(W + "line")) * 0.92))))
    parts = _footers(root, rels, types)
    # пустые абзацы в конце страницы (pdf2docx ставит их вокруг таблиц; после переноса штампа их по 2–4) вместе
    # с абзацем-разрывом выталкивали его на новый лист — оставалась пустая страница с одним колонтитулом
    empty = lambda e: (e.tag == W + "p" and not "".join(t.text or "" for t in e.iter(W + "t")).strip()
                       and e.find(f".//{W}drawing") is None and e.find(f".//{W}pict") is None
                       and e.find(f".//{W}object") is None and e.find(f"{W}pPr/{W}sectPr") is None)
    for sp in [x.getparent().getparent() for x in root.iter(W + "sectPr") if x.getparent().tag == W + "pPr"]:
        prev = sp.getprevious()
        while prev is not None and empty(prev):
            nxt = prev.getprevious(); prev.getparent().remove(prev); prev = nxt
    tmp = Path(path).with_suffix(".norm.docx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "word/_rels/document.xml.rels":
                data = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "[Content_Types].xml":
                data = etree.tostring(types, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
        for name, data in parts.items():
            zout.writestr(name, data)
    tmp.replace(path)


def _footers(root, rels, types):
    """Таблица-штамп внизу листа (повторяется в конце многих страниц: «Létesítmény: … Verzió, Lapszám») —
    в настоящий нижний колонтитул раздела. pdf2docx кладёт её в тело страницы: русский текст длиннее, строки
    растут — и штамп выпадал на отдельную пустую страницу (UTS: 8 листов → 14, 15.10).
    Возвращает новые части пакета {имя: байты}; rels и types дополняются на месте."""
    import copy
    from collections import Counter
    from lxml import etree
    W = SG.W
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    body = root.find(W + "body")
    sections, cur = [], []
    for el in body:
        if el.tag == W + "sectPr":
            continue
        cur.append(el)
        if el.tag == W + "p" and el.find(f"{W}pPr/{W}sectPr") is not None:
            sections.append((cur, el.find(f"{W}pPr/{W}sectPr"))); cur = []
    if cur and body.find(W + "sectPr") is not None:
        sections.append((cur, body.find(W + "sectPr")))
    text = lambda el: "".join(t.text or "" for t in el.iter(W + "t"))
    # подпись штампа — весь его текст без цифр и пробелов: штамп на всех листах один, меняются номер листа и дата;
    # по первым 12 знакам штампом считалась таблица шагов TU-документов (шапка повторяется на каждом листе,
    # тело — разное) и уходила в колонтитул целиком (30RDIT_TU, красная команда 15.10)
    sig = lambda tbl: re.sub(r"[\s\d]+", "", text(tbl))
    rows_h = lambda tbl: sum(int(h.get(W + "val", "0")) for h in (tr.find(f"{W}trPr/{W}trHeight")
                                                                   for tr in tbl.findall(W + "tr")) if h is not None)
    tails = []
    for els, sect in sections:
        tail = None
        for el in reversed(els):
            if el.tag == W + "tbl":
                tail = el; break
            if el.tag == W + "p" and text(el).strip():
                break
        tails.append(tail)
    page_h = max([int(x.get(W + "h", "0")) for x in root.iter(W + "pgSz")] + [16838])
    # штамп невысокий (не больше пятой части листа) и повторяется хотя бы на трети листов
    counts = Counter(sig(t) for t in tails if t is not None and sig(t) and rows_h(t) <= 0.2 * page_h)
    common = {k for k, n in counts.items() if n >= max(2, 0.3 * len(sections))}
    parts = {}
    for els, sect in sections:
        # последний штамп раздела, даже если после него есть текст: на титуле UTS штамп спрятан под фоном
        # во весь лист, а ниже идёт юридический текст
        tbl = next((el for el in reversed(els) if el.tag == W + "tbl" and sig(el) in common
                    and rows_h(el) <= 0.2 * page_h), None)
        if tbl is None:
            continue
        n = len(parts) + 1
        name, rid = f"word/footer{100 + n}.xml", f"rIdStamp{n}"
        ftr = etree.Element(W + "ftr", nsmap=root.nsmap)
        ftr.append(copy.deepcopy(tbl))
        etree.SubElement(ftr, W + "p")
        height = rows_h(tbl)          # по строкам самой таблицы: pdf2docx вкладывает штамп таблицей в таблицу
        tbl.getparent().remove(tbl)
        parts[name] = etree.tostring(ftr, xml_declaration=True, encoding="UTF-8", standalone=True)
        etree.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                         Id=rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                         Target=name.split("/", 1)[1])
        etree.SubElement(types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                         PartName="/" + name,
                         ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        ref = etree.Element(W + "footerReference"); ref.set(W + "type", "default"); ref.set(R + "id", rid)
        sect.insert(0, ref)
        mar = sect.find(W + "pgMar")
        if mar is not None:
            bottom = int(mar.get(W + "bottom", "500"))
            mar.set(W + "footer", str(bottom))
            mar.set(W + "bottom", str(bottom + height + 120))
    return parts


def paged_segments(docx):
    """Сегменты DOCX (segments.docx_segments) с номером страницы: склейка постраничных DOCX кладёт sectPr
    в последний абзац каждой страницы (assemble._merge_docx)."""
    from lxml import etree
    ds = SG.docx_segments(docx)
    root = etree.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    W = SG.W
    # граница страницы — закладка «_hutPage» (assemble._merge_docx); свои разрывы разделов pdf2docx ставит и внутри
    # страницы, поэтому по ним — только в DOCX без закладок
    mark = lambda el: any((b.get(W + "name") or "").startswith("_hutPage") for b in el.iter(W + "bookmarkStart"))
    marks = mark(root)
    is_end = mark if marks else (lambda p: p.find(W + "pPr/" + W + "sectPr") is not None)
    page_of, pg = {}, 0
    for pi, p in enumerate(root.iter(W + "p")):
        page_of[pi] = pg
        if is_end(p):
            pg += 1
    for d in ds:
        d["page"] = page_of.get(d["loc"]["p"]) if d["loc"]["part"] == "word/document.xml" else None
    return ds


def plan(pdf_path, segs, docx_out, start_id):
    """DOCX по оригиналу и сопоставление его абзацев с сегментами PDF. Возвращает (новые сегменты, план) или
    (None, ошибка): без DOCX сборка идёт прежним путём."""
    from assemble import docx_from_pdf
    Path(docx_out).parent.mkdir(parents=True, exist_ok=True)
    info = docx_from_pdf(pdf_path, docx_out)
    if info.get("error"):
        return None, info
    normalize(docx_out)
    ds = paged_segments(docx_out)
    by_page = {}
    for s in segs:
        if "frags" in s["loc"]:                      # только текст страницы: не скриншоты и не абзацы DOCX
            by_page.setdefault(s["page"], []).append(s)
    keys = {s["id"]: KEY(s["text"]) for s in segs}
    steps, cursor, parts = [], {}, {}
    for j, d in enumerate(ds):
        if not SG.needs_translation(TAGS.sub("", d["text"])):
            steps.append(None); continue
        k, cands = KEY(d["text"]), by_page.get(d["page"], [])
        # несколько непустых начертаний («FIGYELEM!» жирным + текст), а перевод сегмента PDF без меток — ушёл бы
        # в первую группу, весь абзац стал бы жирным (красная команда 15.10): такой абзац переводится сам, с метками.
        # Только для целого абзаца: часть сегмента (pdf2docx режет абзац на строки и слова) идёт цепочкой частей,
        # иначе обрывки «jogosultsággal», «rendelkező» переводились бы по слову (UTS стр. 55)
        multi = len([g for g in GROUP.findall(d["text"]) if g[1].strip()]) > 1
        hit = next((s for s in (cands if d["page"] is not None else segs) if "frags" in s["loc"] and keys[s["id"]] == k),
                   None)
        if hit:
            steps.append({"k": k, "own": True} if multi else {"k": k, "ids": [hit["id"]]}); continue
        run = None
        for i in range(len(cands)):                  # абзац DOCX = несколько сегментов PDF подряд
            acc, ids = "", []
            for s in cands[i:i + 10]:
                acc += keys[s["id"]]; ids.append(s["id"])
                if acc == k:
                    run = ids
                if not k.startswith(acc) or run:
                    break
            if run:
                break
        if run:
            steps.append({"k": k, "own": True} if multi else {"k": k, "ids": run}); continue
        # абзац DOCX — очередная часть сегмента PDF (pdf2docx режет абзац по строкам или ячейкам)
        part = next((s for s in cands if len(k) >= 2 and keys[s["id"]][cursor.get(s["id"], 0):].startswith(k)), None)
        if part:
            cursor[part["id"]] = cursor.get(part["id"], 0) + len(k)
            parts.setdefault(part["id"], []).append(j)
            steps.append({"k": k, "part": part["id"]}); continue
        steps.append({"k": k, "own": True})
    for sid, js in parts.items():
        if cursor[sid] == len(keys[sid]):            # сегмент собран из частей целиком: перевод — в первую
            steps[js[0]] = {"k": steps[js[0]]["k"], "ids": [sid]}
            for j in js[1:]:
                steps[j] = {"k": steps[j]["k"], "empty": True, "of": sid}
        else:
            for j in js:
                steps[j] = {"k": steps[j]["k"], "own": True}
    extra = []
    for j, st in enumerate(steps):
        if st and st.get("own"):
            d = ds[j]
            st["own"] = start_id + len(extra)
            extra.append({"id": st["own"], "text": d["text"], "src": d["src"], "page": d["page"],
                          "loc": {"docx_only": True}})
    return extra, {"steps": steps, "pages": info.get("pages")}


def _toc_tabs(path):
    """Строка оглавления «текст ….. 19» — «текст⇥19» с правой табуляцией и точками-заполнителями на ширину текста
    раздела: pdf2docx набирает точки текстом, перевод длиннее — строки переносились и разъезжались (UTS, 15.10)."""
    import copy
    from lxml import etree
    from assemble import TOC
    W = SG.W
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    paras = list(body.iter(W + "p"))
    sects = [(i, p.find(f"{W}pPr/{W}sectPr")) for i, p in enumerate(paras) if p.find(f"{W}pPr/{W}sectPr") is not None]
    sects.append((len(paras), body.find(W + "sectPr")))
    done = 0
    for i, p in enumerate(paras):
        runs = [r for r in p.iter(W + "r") if r.find(W + "t") is not None]
        text = "".join(t.text or "" for t in p.iter(W + "t"))
        m = TOC.match(text.strip()) if runs else None
        if not m or len(re.findall(r"\.{3,}", text)) != 1:
            continue
        sect = next(s for k, s in sects if k >= i)
        if sect is None or sect.find(W + "pgSz") is None or sect.find(W + "pgMar") is None:
            continue
        mar, sz = sect.find(W + "pgMar"), sect.find(W + "pgSz")
        width = int(sz.get(W + "w")) - int(float(mar.get(W + "left", "0"))) - int(float(mar.get(W + "right", "0")))
        first = runs[0]
        for r in runs[1:]:
            for ch in list(r):
                if ch.tag in (W + "t", W + "tab", W + "br"):
                    r.remove(ch)
        for ch in list(first):
            if ch.tag in (W + "t", W + "tab", W + "br"):
                first.remove(ch)
        t1 = etree.SubElement(first, W + "t"); t1.text = m.group("t").strip(" .…")
        t1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        etree.SubElement(first, W + "tab")
        etree.SubElement(first, W + "t").text = m.group("n")
        ppr = p.find(W + "pPr")
        if ppr is None:
            ppr = etree.Element(W + "pPr"); p.insert(0, ppr)
        for tag in ("tabs", "jc"):
            el = ppr.find(W + tag)
            if el is not None:
                ppr.remove(el)
        tabs = etree.Element(W + "tabs")
        tab = etree.SubElement(tabs, W + "tab"); tab.set(W + "val", "right"); tab.set(W + "leader", "dot")
        tab.set(W + "pos", str(max(width - 20, 1000)))
        # порядок детей pPr: tabs — до spacing/ind; jc — после ind
        anchor = next((x for x in ppr if x.tag in {W + t for t in ("suppressAutoHyphens", "kinsoku", "wordWrap",
                       "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
                       "snapToGrid", "spacing", "ind") + PPR_AFTER_SPACING}), None)
        (anchor.addprevious if anchor is not None else ppr.append)(tabs)
        done += 1
    tmp = Path(path).with_suffix(".toc.docx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    tmp.replace(path)
    return done


def _fix_fonts(path):
    """Шрифты, которых Word не знает или которыми рисует не то: «•» в Symbol Word показывает квадратом."""
    pat = re.compile(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")(%s)(")' % "|".join(map(re.escape, FONTS)))
    tmp = Path(path).with_suffix(".tmp.docx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                data = pat.sub(lambda m: m.group(1) + FONTS[m.group(2)] + m.group(3), data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def build(pdf_for_docx, out, segs, res, dplan):
    """DOCX перевода по плану извлечения. pdf_for_docx — оригинал с переведёнными скриншотами."""
    from assemble import docx_from_pdf, build_docx
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "src.docx"
        info = docx_from_pdf(pdf_for_docx, src)
        if info.get("error"):
            return info
        normalize(src)
        ds = paged_segments(src)
        steps = dplan["steps"]
        by_key = {}
        for st in steps:                  # «пустая часть» не перекрывает настоящий перевод того же текста («RC»)
            if st and (st["k"] not in by_key or by_key[st["k"]].get("empty")):
                by_key[st["k"]] = st
        if len(ds) != len(steps) or any(st and KEY(d["text"]) != st["k"] for d, st in zip(ds, steps)):
            # pdf2docx разобрал страницы иначе, чем при извлечении — сопоставление по тексту абзаца
            steps = [by_key.get(KEY(d["text"])) if SG.needs_translation(TAGS.sub("", d["text"])) else None for d in ds]
        ru = lambda i: TAGS.sub("", (res.get(str(i)) or {}).get("ru") or "").strip()
        pseudo, missed = {}, 0
        # части сегмента, у которого нет перевода, не стираются: иначе венгерский текст обрезался бы
        lacking = {i for st in steps if st and st.get("ids") for i in st["ids"] if not ru(i)}
        for d, st in zip(ds, steps):
            if not st:
                continue
            if st.get("empty"):
                if st.get("of") in lacking:
                    continue
                t = " "                                  # часть абзаца, перевод которого ушёл в первую часть
            elif "own" in st:
                t = (res.get(str(st["own"])) or {}).get("ru") or ""
            else:
                t = " ".join(x for x in (ru(i) for i in st["ids"]) if x)
                groups = GROUP.findall(d["text"])
                if t and len(groups) > 1:
                    # группы из одних пробелов и табуляций («⇥текст» пункта списка) остаются, перевод — в текстовую
                    first = next(i for i, g in enumerate(groups) if g[1].strip())
                    t = "".join(f"<f{n}>{t if i == first else (g if not g.strip() else '')}</f{n}>"
                                for i, (n, g) in enumerate(groups))
            if not t:
                missed += 1; continue
            pseudo[str(d["id"])] = {"ru": t}
        build_docx(src, out, ds, pseudo)
    _toc_tabs(out)
    _fix_fonts(out)
    return {"converter": "pdf2docx(оригинал)", "pages": info.get("pages"), "paragraphs": len(pseudo), "missed": missed}
