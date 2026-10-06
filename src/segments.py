# -*- coding: utf-8 -*-
"""Извлечение сегментов (единиц перевода) из трёх видов источников.

Сегмент = абзац / ячейка / блок PDF. Поля: id, text (для модели; при смешанном
форматировании — с метками <f1>…</f1>), src (чистый текст), loc (как вернуть
перевод на место), page.

- DOCX: все w:p во всех частях (тело, таблицы, врезки, колонтитулы); врезки в DOCX
  лежат дважды (DrawingML + VML-запасной), одинаковый текст переводится один раз.
- PDF с текстовым слоем: блоки PyMuPDF с bbox, переносы слов склеиваются.
- PDF без текстового слоя (3PR42, текст кривыми): Tesseract даёт слова с рамками,
  портал — второе чтение; коды/прописные/спецзнаки берутся у портала, строчные
  слова — у Tesseract (замер 01.10: ошибки чтецов не пересекаются)."""
import re, json, difflib, subprocess, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OCR_DIR = ROOT / "work" / "ocr"                  # переопределяется конвейером: <рабочая папка>/ocr
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HU_WORD = re.compile(r"[a-záéíóöőúüű]{2,}")


def needs_translation(text):
    """Есть ли что переводить: строчное слово или аббревиатура из букв (не чистые коды/числа)."""
    t = re.sub(r"<\/?f\d+>", "", text)
    return bool(HU_WORD.search(t)) or bool(re.search(r"\b[A-ZÁÉÍÓÖŐÚÜŰ]{2,}\b", t))


# ---------------------------------------------------------------- DOCX
def _own_runs(p):
    """w:r этого абзаца без w:r вложенных абзацев (врезка внутри рисунка)."""
    out = []
    for r in p.iter(W + "r"):
        a = r.getparent()
        while a is not None and a.tag != W + "p":
            a = a.getparent()
        if a is p:
            out.append(r)
    return out


def run_text(r):
    s = []
    for ch in r:
        tag = ch.tag
        if tag == W + "t":
            s.append(ch.text or "")
        elif tag == W + "tab":
            s.append("\t")
        elif tag in (W + "br", W + "cr"):
            s.append("\n")
        elif tag == W + "noBreakHyphen":
            s.append("‑")
    return "".join(s)


def run_sig(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        return ()
    sig = []
    for name in ("b", "i", "u", "caps", "strike"):
        el = rpr.find(W + name)
        if el is not None and el.get(W + "val") not in ("0", "false", "none"):
            sig.append(name)
    va = rpr.find(W + "vertAlign")
    if va is not None and va.get(W + "val") in ("superscript", "subscript"):
        sig.append(va.get(W + "val"))
    return tuple(sig)


LEAD = re.compile(r"\s*[-☐☒□■]+\s*")


def _symbol_run(r, t):
    if re.fullmatch(r"[-☐☒□■▪●○✓✔\s]+", t) and t.strip():
        return True
    rpr = r.find(W + "rPr")
    fonts = rpr.find(W + "rFonts") if rpr is not None else None
    if fonts is not None and re.search(r"Wingdings|Symbol|Webdings", " ".join(fonts.attrib.values())):
        return True
    return False


def docx_parts(path):
    """Имена XML-частей с текстом: тело, колонтитулы, сноски."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    keep = ["word/document.xml"] + sorted(n for n in names if re.match(r"word/(header|footer|footnotes|endnotes)\d*\.xml$", n))
    return keep


def docx_segments(path):
    from lxml import etree
    segs = []
    with zipfile.ZipFile(path) as z:
        for part in docx_parts(path):
            root = etree.fromstring(z.read(part))
            for pi, p in enumerate(root.iter(W + "p")):
                runs = _own_runs(p)
                texts = [run_text(r) for r in runs]
                if not "".join(texts).strip():
                    continue
                # группы подряд идущих прогонов с одинаковым начертанием
                groups = []
                for ri, (r, t) in enumerate(zip(runs, texts)):
                    if not t:
                        continue
                    # значки шрифтов Wingdings/Symbol (флажки «□») — не текст: модель их выбрасывает,
                    # прогон остаётся на месте нетронутым
                    if _symbol_run(r, t):
                        continue
                    sg = run_sig(r)
                    if groups and (groups[-1]["sig"] == sg or not t.strip()):
                        groups[-1]["runs"].append(ri); groups[-1]["text"] += t
                    else:
                        groups.append({"sig": sg, "runs": [ri], "text": t})
                # ведущий значок в одном прогоне с текстом («  E2\t…», 3925 абзацев 3PR04) —
                # отрезать от текста для модели, вернуть при сборке
                lead = ""
                if groups:
                    m = LEAD.match(groups[0]["text"])
                    if m and m.group().strip():
                        lead = m.group()
                        groups[0]["text"] = groups[0]["text"][len(lead):]
                        if not groups[0]["text"]:
                            groups.pop(0)           # прогон из одного значка — не трогать
                            lead = ""
                if not groups:
                    continue
                src = "".join(g["text"] for g in groups)
                if len(groups) > 1:
                    text = "".join(f"<f{k+1}>{g['text']}</f{k+1}>" for k, g in enumerate(groups))
                else:
                    text = src
                segs.append({"id": len(segs), "text": text, "src": src, "page": None,
                             "loc": {"part": part, "p": pi, "groups": [g["runs"] for g in groups], "lead": lead}})
    return segs


# ---------------------------------------------------------------- PDF
OCR_PROMPT = ("Перепиши весь текст страницы дословно, на языке оригинала (венгерский), сохраняя порядок, "
              "нумерацию пунктов, переносы абзацев и все коды/обозначения посимвольно (буквы с диакритикой: "
              "á é í ó ö ő ú ü ű) и знаки ≤ ≥ < > ± ÷ Δ ° %. Ничего не переводи, не пропускай и не добавляй. Выведи только текст.")
_NORM = re.compile(r"[^\w%Δ÷±≤≥<>]+")


def _norm(tok):
    return _NORM.sub("", tok.lower())


def _codeish(tok):
    # ≤ ≥ < > — тоже «код»: иначе «ti ≤ 70» портала проигрывало «tiz 70» Tesseract (→ «при 10 70 °C», 01.10)
    core = re.sub(r"[^\w%Δ÷±≤≥<>„”\"]", "", tok)
    if not core:
        return False
    if re.search(r"[0-9%Δ÷±≤≥<>„”\"]", core):
        return True
    letters = [c for c in core if c.isalpha()]
    return len(letters) >= 2 and all(c.isupper() for c in letters)


def _choose(t, p):
    """t — слово Tesseract, p — слово портала: коды/прописные/спецзнаки/пунктуация — портал, строчные — Tesseract."""
    if not re.search(r"\w", t) or not re.search(r"\w", p):
        return p
    return p if (_codeish(t) or _codeish(p)) else t


def merge_readers(words, portal_text):
    """words: слова Tesseract [{t, box, line, h}] в порядке чтения; правит w['t'] по второму чтению.
    Возвращает новый список (слова, которых нет у Tesseract, приклеиваются к соседу слева)."""
    ptoks = portal_text.split()
    a = [_norm(w["t"]) for w in words]
    b = [_norm(t) for t in ptoks]
    out = [dict(w) for w in words]
    keep = [True] * len(words)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op in ("equal", "replace") and i2 - i1 == j2 - j1:
            for k in range(i2 - i1):
                out[i1 + k]["t"] = _choose(words[i1 + k]["t"], ptoks[j1 + k])
        elif op == "replace":
            ts = [w["t"] for w in words[i1:i2]]
            if any(_codeish(t) for t in ts + ptoks[j1:j2]):
                out[i1]["t"] = " ".join(ptoks[j1:j2])
                for k in range(i1 + 1, i2):
                    keep[k] = False
        elif op == "delete":
            for k in range(i1, i2):
                if not _norm(words[k]["t"]):        # мусор пунктуации («- .» на месте тире списка)
                    keep[k] = False
        elif op == "insert" and i1 > 0:
            out[i1 - 1]["t"] += " " + " ".join(ptoks[j1:j2])
    return [w for w, k in zip(out, keep) if k]


def ocr_paragraphs(words, portal_text):
    """Абзацы OCR-страницы: границы — по строкам чтения портала, текст — сверка чтецов,
    рамка — объединение рамок совпавших слов Tesseract. Возвращает (абзацы, слов только у Tesseract)."""
    plines = [ln.strip() for ln in portal_text.splitlines() if ln.strip()]
    ptoks, pline = [], []
    for li, ln in enumerate(plines):
        for t in ln.split():
            ptoks.append(t); pline.append(li)
    a = [_norm(w["t"]) for w in words]
    b = [_norm(t) for t in ptoks]
    text = list(ptoks)
    box = [None] * len(ptoks)
    lost = []
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op in ("equal", "replace") and i2 - i1 == j2 - j1:
            for k in range(i2 - i1):
                text[j1 + k] = _choose(words[i1 + k]["t"], ptoks[j1 + k]); box[j1 + k] = words[i1 + k]["box"]
        elif op == "replace":
            # разное число слов: текст портала, рамки — все слова Tesseract этого куска
            ub = [w["box"] for w in words[i1:i2]]
            for k in range(j1, j2):
                box[k] = [min(b_[0] for b_ in ub), min(b_[1] for b_ in ub), max(b_[2] for b_ in ub), max(b_[3] for b_ in ub)]
            if not any(_codeish(w["t"]) for w in words[i1:i2]) and not any(_codeish(t) for t in ptoks[j1:j2])                     and (i2 - i1) > (j2 - j1):
                lost += [w["t"] for w in words[i1:i2] if re.search(r"[a-záéíóöőúüű]{3,}", w["t"])][(j2 - j1):]
        elif op == "delete":
            lost += [w["t"] for w in words[i1:i2] if re.search(r"[a-záéíóöőúüű]{3,}", w["t"])]
    pars = []
    for li, ln in enumerate(plines):
        idx = [j for j in range(len(ptoks)) if pline[j] == li]
        bxs = [box[j] for j in idx if box[j] is not None]
        if not bxs:
            continue
        pars.append({"t": " ".join(text[j] for j in idx),
                     "bbox": [min(b_[0] for b_ in bxs), min(b_[1] for b_ in bxs), max(b_[2] for b_ in bxs), max(b_[3] for b_ in bxs)],
                     "bold": False})
    return pars, lost


def _tess_words(png_path, dpi):
    r = subprocess.run(["tesseract", str(png_path), "stdout", "-l", "hun", "--psm", "4", "tsv"],
                       capture_output=True, text=True, encoding="utf-8")
    words = []
    for ln in r.stdout.splitlines()[1:]:
        c = ln.split("\t")
        if len(c) < 12 or c[0] != "5" or not c[11].strip():
            continue
        x, y, w, h = map(int, c[6:10])
        words.append({"t": c[11], "key": (int(c[2]), int(c[3])), "line": (int(c[2]), int(c[3]), int(c[4])),
                      "box": [x * 72 / dpi, y * 72 / dpi, (x + w) * 72 / dpi, (y + h) * 72 / dpi],
                      "h": h * 72 / dpi})
    return words


def _dehyphen(words):
    """Склейка переносов на концах строк Tesseract (внутри абзаца Tesseract)."""
    out = []
    for w in words:
        if out and out[-1]["t"].endswith("-") and out[-1]["line"] != w["line"] and out[-1]["key"] == w["key"] \
                and w["t"][:1].islower():
            prev = out[-1]
            keep = re.search(r"[A-ZÁÉÍÓÖŐÚÜŰ0-9]-$", prev["t"])
            out[-1] = dict(prev, t=(prev["t"] if keep else prev["t"][:-1]) + w["t"])
        else:
            out.append(dict(w))
    return out


def pdf_segments(path, pages=None, ocr="auto", dpi_tess=300, dpi_portal=200, ratio=0.85):
    """ocr: 'auto' — страница идёт через OCR, если в текстовом слое < ratio слов Tesseract."""
    import fitz
    from concurrent.futures import ThreadPoolExecutor
    from layout import frags_from_text, frags_from_words, rules, group, avail_rects
    doc = fitz.open(path)
    work = Path(OCR_DIR) / Path(path).stem; work.mkdir(parents=True, exist_ok=True)
    rng = list(pages if pages is not None else range(len(doc)))

    def tess(pno):
        f = work / f"p{pno:04d}.json"
        if f.exists():
            return json.loads(f.read_text(encoding="utf-8"))
        png = work / f"p{pno:04d}.png"
        fitz.open(path)[pno].get_pixmap(dpi=dpi_tess, colorspace=fitz.csGRAY).save(str(png))
        ws = _tess_words(png, dpi_tess); png.unlink()
        f.write_text(json.dumps(ws, ensure_ascii=False), encoding="utf-8")
        return ws

    with ThreadPoolExecutor(4) as ex:
        tw = dict(zip(rng, ex.map(tess, rng)))

    def portal(pno):
        from llm import chat, img_b64
        png = fitz.open(path)[pno].get_pixmap(dpi=dpi_portal, colorspace=fitz.csGRAY).tobytes("png")
        return chat([{"role": "user", "content": OCR_PROMPT}], images=[img_b64(png)], max_tokens=8000, tag=f"ocr{pno}")

    plan = {}
    for pno in rng:
        n_text = len(re.findall(r"\w{2,}", doc[pno].get_text()))
        n_tess = len([w for w in tw[pno] if re.search(r"\w{2,}", w["t"])])
        plan[pno] = (ocr is True) or (ocr == "auto" and n_tess >= 15 and n_text < ratio * n_tess)
    ocr_pages = [p for p in rng if plan[p]]
    with ThreadPoolExecutor(4) as ex:
        pr = dict(zip(ocr_pages, ex.map(portal, ocr_pages)))

    segs, errors, lost = [], [], {}
    for pno in rng:
        page = doc[pno]
        hs, vs = rules(page)
        if plan[pno]:
            r = pr[pno]
            ws = _dehyphen(tw[pno])
            if r.get("text"):
                lines, lost[pno] = ocr_paragraphs(ws, r["text"])
                # кегль страницы: шаг строк Tesseract / 1.2 (кегль из высоты слова врёт на выносных)
                tops = {}
                for w in ws:
                    tops.setdefault(tuple(w["line"]), w["box"][1])
                ys = sorted(tops.values())
                steps = sorted(b_ - a_ for a_, b_ in zip(ys, ys[1:]) if 6 < b_ - a_ < 30)
                size = round(steps[len(steps) // 2] / 1.2, 1) if steps else 10.0
                # остатки текстового слоя на странице знают настоящий кегль
                tl = sorted(sp["size"] for b_ in page.get_text("dict")["blocks"] for l_ in b_.get("lines", [])
                            for sp in l_["spans"] if len(sp["text"].strip()) > 2)
                if len(tl) >= 5:
                    size = tl[len(tl) // 2]
                for ln in lines:
                    ln["size"] = size
                pars = group(lines, hs, vs)
            else:
                errors.append((pno, r.get("error")))
                pars = group(frags_from_words(ws), hs, vs)
        else:
            pars = group(frags_from_text(page, vs), hs, vs)
        avail_rects(pars, hs, vs, page.rect)
        for p in pars:
            segs.append({"id": len(segs), "text": p["text"], "src": p["text"], "page": pno,
                         "loc": {"bbox": [round(v, 2) for v in p["bbox"]], "avail": [round(v, 2) for v in p["avail"]],
                                 "frags": [[round(v, 2) for v in f["bbox"]] for f in p["frags"]],
                                 "size": p["size"], "bold": p["bold"], "ocr": plan[pno]}})
    return segs, {"ocr_pages": ocr_pages, "errors": errors, "lost_words": {k: v for k, v in lost.items() if v}}
