"""Подгонка перевода DOCX «на месте» под листы оригинала (20.10). Документ, свёрстанный «раздел = лист» (3PR04: 514
разделов на 514 листах, номер листа в шапке — текстом), после перевода расползался: русский текст длиннее, у 137
разделов две-три строки уходили на новый лист — 514 → 651 листов, почти пустые листы с тем же номером в шапке.

Раздел перевода, занявший больше листов, чем в оригинале, — мельче кеглем (по разделу, до 80 %), пока не уложится.
Сколько листов у раздела — печатью LibreOffice с метками начала разделов (невидимый текст в 1 pt в первом абзаце
раздела): у оригинала и у перевода. Без LibreOffice — без подгонки."""
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MARK = "QZS{}QZ"
FLOOR = 0.8
# rPr по схеме: что идёт после кегля (новый w:sz — перед ними, иначе Word считает файл повреждённым)
AFTER_SZ = {"highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
            "eastAsianLayout", "specVanish", "oMath"}


def _q(tag):
    return W + tag


def _sections(body):
    """Части тела по разделам: [[элементы раздела]] — раздел кончается абзацем с w:sectPr (последний — sectPr тела)."""
    out, cur = [], []
    for el in body:
        if el.tag == _q("sectPr"):
            continue
        cur.append(el)
        if el.tag == _q("p") and el.find(f"{_q('pPr')}/{_q('sectPr')}") is not None:
            out.append(cur)
            cur = []
    if cur:
        out.append(cur)
    return out


def _first_p(els):
    for el in els:
        p = next(el.iter(_q("p")), None)
        if p is not None:
            return p
    return None


def _marked(src, dst):
    """Копия DOCX с меткой в начале каждого раздела (белый текст 1 pt) → число разделов."""
    n = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                root = etree.fromstring(data)
                secs = _sections(root.find(_q("body")))
                n = len(secs)
                for i, sec in enumerate(secs):
                    p = _first_p(sec)
                    if p is None:
                        continue
                    r = etree.Element(_q("r"))
                    rpr = etree.SubElement(r, _q("rPr"))
                    etree.SubElement(rpr, _q("color")).set(_q("val"), "FFFFFF")
                    for t in ("sz", "szCs"):
                        etree.SubElement(rpr, _q(t)).set(_q("val"), "2")
                    etree.SubElement(r, _q("t")).text = MARK.format(i)
                    ppr = p.find(_q("pPr"))
                    if ppr is not None:
                        ppr.addnext(r)
                    else:
                        p.insert(0, r)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    return n


def _layout(pdf):
    """По напечатанному PDF с метками: {раздел: (лист начала, листов до начала следующего, заполнение последнего листа
    0..1)}. Мера одна у оригинала и перевода — сравниваются числа листов."""
    import fitz
    with fitz.open(str(pdf)) as doc:
        start = {}
        H = doc[0].rect.height if len(doc) else 842
        body = []                                   # (верх, низ) текста тела листа без колонтитулов
        for pno, pg in enumerate(doc):
            for m in re.finditer(r"QZS(\d+)QZ", pg.get_text()):
                start.setdefault(int(m.group(1)), pno)
            bs = [b for b in pg.get_text("blocks") if b[1] > 0.1 * H and b[3] < 0.93 * H and b[4].strip()]
            body.append((min(b[1] for b in bs), max(b[3] for b in bs)) if bs else None)
        total = len(doc)
    known = sorted((p, i) for i, p in start.items())
    out = {}
    for k, (p, i) in enumerate(known):
        nxt = known[k + 1][0] if k + 1 < len(known) else total
        span = max(1, nxt - p)
        full = [b for b in body[p:p + span - 1] if b] or [b for b in body[p:p + 1] if b]
        top = min((b[0] for b in full), default=0.1 * H)
        bottom = max((b[1] for b in full), default=0.9 * H)
        tail = body[p + span - 1] if p + span - 1 < len(body) else None
        fill = 1.0 if not tail or bottom <= top else min(1.0, max(0.05, (tail[1] - top) / (bottom - top)))
        out[i] = (p, span, fill)
    return out


def _style_sizes(styles_xml):
    """({id стиля: кегль в полупунктах с учётом basedOn или None}, кегль умолчаний документа)."""
    root = etree.fromstring(styles_xml)
    raw, base = {}, {}
    for st in root.iter(_q("style")):
        sid = st.get(_q("styleId"))
        sz = st.find(f"{_q('rPr')}/{_q('sz')}")
        raw[sid] = int(sz.get(_q("val"))) if sz is not None and (sz.get(_q("val")) or "").isdigit() else None
        b = st.find(_q("basedOn"))
        base[sid] = b.get(_q("val")) if b is not None else None
    dd = root.find(f"{_q('docDefaults')}/{_q('rPrDefault')}/{_q('rPr')}/{_q('sz')}")
    default = int(dd.get(_q("val"))) if dd is not None else 20

    def size(sid, seen=()):
        if sid is None or sid in seen or sid not in raw:
            return None
        return raw[sid] if raw[sid] is not None else size(base[sid], seen + (sid,))
    return {sid: size(sid) for sid in raw}, default


def _put_sz(rpr, tag, val):
    el = rpr.find(_q(tag))
    if el is None:
        el = etree.Element(_q(tag))
        nxt = next((c for c in rpr if etree.QName(c).localname in AFTER_SZ
                    or tag == "sz" and etree.QName(c).localname == "szCs"), None)
        if nxt is not None:
            nxt.addprevious(el)
        else:
            rpr.append(el)
    el.set(_q("val"), val)


def _shrink_p(p, k, sizes, default):
    """Абзац мельче в k раз: кегль прогонов (свой, стиля знака, стиля абзаца, умолчаний) и точные интервалы."""
    ppr = p.find(_q("pPr"))
    ps = ppr.find(_q("pStyle")) if ppr is not None else None
    psz = sizes.get(ps.get(_q("val"))) if ps is not None else sizes.get("Normal")
    sp = ppr.find(_q("spacing")) if ppr is not None else None
    if sp is not None:
        for a in ("before", "after"):
            v = sp.get(_q(a))
            if v and v.isdigit():
                sp.set(_q(a), str(round(int(v) * k)))
        line = sp.get(_q("line")) or ""
        if sp.get(_q("lineRule")) in ("exact", "atLeast") and line.isdigit():
            sp.set(_q("line"), str(round(int(line) * k)))
    for r in p.iter(_q("r")):
        if next(r.iterancestors(_q("p")), None) is not p:
            continue                                # прогон вложенного абзаца (надпись) — у своего абзаца
        rpr = r.find(_q("rPr"))
        if rpr is None:
            rpr = etree.Element(_q("rPr"))
            r.insert(0, rpr)
        rs = rpr.find(_q("rStyle"))
        sz = rpr.find(_q("sz"))
        own = int(sz.get(_q("val"))) if sz is not None and (sz.get(_q("val")) or "").isdigit() else None
        cur = own or (sizes.get(rs.get(_q("val"))) if rs is not None else None) or psz or default
        new = str(max(2, round(cur * k)))
        _put_sz(rpr, "sz", new)
        _put_sz(rpr, "szCs", new)


def _shrink(src, dst, ks):
    """Копия DOCX, где у разделов ks {номер: доля} абзацы мельче (_shrink_p)."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        sizes, default = _style_sizes(zin.read("word/styles.xml")) if "word/styles.xml" in names else ({}, 20)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml" and ks:
                root = etree.fromstring(data)
                for i, sec in enumerate(_sections(root.find(_q("body")))):
                    if ks.get(i):
                        for el in sec:
                            for p in el.iter(_q("p")):
                                _shrink_p(p, ks[i], sizes, default)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)


def fit(orig, ru, printer, work=None):
    """Подогнать ru (DOCX перевода «на месте», меняется на месте) под листы orig. printer(docx, pdf) — печать
    LibreOffice. → {"sections", "over" — разделов длиннее оригинала, "fixed", "left"}."""
    tmp = Path(work or tempfile.mkdtemp(prefix="docx_fit_"))
    tmp.mkdir(parents=True, exist_ok=True)
    try:
        return _fit(orig, ru, printer, tmp)
    finally:
        if work is None:
            shutil.rmtree(tmp, ignore_errors=True)


def _fit(orig, ru, printer, tmp):
    n = _marked(orig, tmp / "orig_m.docx")
    printer(tmp / "orig_m.docx", tmp / "orig_m.pdf")
    want = {i: s for i, (_, s, _) in _layout(tmp / "orig_m.pdf").items()}
    ks, info = {}, {"sections": n, "over": 0, "fixed": 0, "left": 0}
    cur = Path(ru)
    for step in range(3):
        _marked(cur, tmp / f"ru_m{step}.docx")
        printer(tmp / f"ru_m{step}.docx", tmp / f"ru_m{step}.pdf")
        got = _layout(tmp / f"ru_m{step}.pdf")
        over = {i: v for i, v in got.items() if i in want and v[1] > want[i]}
        if step == 0:
            info["over"] = len(over)
        if not over or step == 2:
            info["left"] = len(over)
            break
        for i, (_, span, fill) in over.items():
            need = (span - 1) + fill                # листов текста раздела
            k0 = ks.get(i, 1.0)
            # высота текста ~ кегль² (строка ниже, и в строку входит больше); запас 2 %; вторая попытка — ещё −6 %
            k = k0 * min(0.98, (want[i] / need) ** 0.5 * 0.98) if step == 0 else k0 * 0.94
            ks[i] = max(FLOOR, k)
        _shrink(ru, tmp / f"ru_fit{step}.docx", ks)
        cur = tmp / f"ru_fit{step}.docx"
    if ks:
        shutil.copyfile(cur, ru)
    info["fixed"] = info["over"] - info["left"]
    return info
