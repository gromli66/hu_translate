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
    return bool(HU_WORD.search(t)) or bool(re.search(r"\b[A-ZÁÉÍÓÖŐÚÜŰ]{2,}\b", t)) \
        or bool(re.search(r"[А-Яа-яЁё]", t))          # русский источник (проект "pair": "ru-en"): и «№ п/п»


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


# поле номера страницы: его результат («24» в «Стр.24») — не текст. Прежде он шёл в сегмент, перевод «Page 24»
# ложился в первый прогон, а результат поля пустел — LibreOffice/Word пересчитывали поле, и выходило «Page 2402»
# (ТЗ тренажёра RU→EN, 20.10)
PAGE_FIELD = re.compile(r"^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b")


def _field_result_runs(runs):
    """Номера прогонов — результатов полей номера страницы (сложное поле fldChar begin…separate…end или fldSimple)."""
    out, stack = set(), []                 # stack: [инструкция, идёт результат] на каждое открытое поле
    for i, r in enumerate(runs):
        par = r.getparent()
        if par is not None and par.tag == W + "fldSimple" and PAGE_FIELD.match(par.get(W + "instr") or ""):
            out.add(i)
            continue
        fc = r.find(W + "fldChar")
        if fc is not None:
            kind = fc.get(W + "fldCharType")
            if kind == "begin":
                stack.append(["", False])
            elif kind == "separate" and stack:
                stack[-1][1] = True
            elif kind == "end" and stack:
                stack.pop()
            continue
        it = r.find(W + "instrText")
        if it is not None and stack:
            stack[-1][0] += it.text or ""
        elif stack and stack[-1][1] and PAGE_FIELD.match(stack[-1][0]):
            out.add(i)
    return out


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
                page_no = _field_result_runs(runs)
                for ri, (r, t) in enumerate(zip(runs, texts)):
                    if not t or ri in page_no:     # номер страницы (поле) — не текст, прогон не трогается
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
# знак градуса — в сравнении слов: без него «°C.» портала становилось «c» и совпадало с «c», которым Tesseract прочёл
# «<» («t < 98 °C.» → «t < 98 c 98 9C.», 3SZ19, 19.10)
_NORM = re.compile(r"[^\w%Δ÷±≤≥<>°]+")


def _norm(tok):
    return _NORM.sub("", tok.lower())


# полужирный Markdown в чтении портала: «**» в начале слова … «**» в конце слова той же строки, без «**» между ними.
# Звёздочки кода — не пара: «(**GB02)» 3SZ19, «a **RD31P702B1» 30RDIT — за ними буква, перед ними скобка или пробел
_BOLD_MD = re.compile(r"(?<!\S)\*\*(?=[^\s*])((?:(?!\*\*)[^\n])+?)(?<=[^\s*])\*\*(?![\w*])")


def _unmark(text):
    """Текст чтения портала без полужирной разметки Markdown: «**3PR42**», «**1.19/2. … szintnövekedése**» — звёздочки
    у «кодовых» слов шли в текст (слово портала), у обычных — нет: заголовок «**1.19/2. …» не совпадал с названием в
    рамке и печатался дважды (3PR42 л. 72, 20.10)."""
    return _BOLD_MD.sub(r"\1", text)


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


# пунктир под подпись («Aláírás: ............») OCR читает буквами: «sszsseseeeeeeeseeee…» — такого венгерского слова
# не бывает (30RDGT nyilv, 06.10: мусор уходил в перевод и раздувал строки формы); заменяем на «…»
# шум начинается с точек после подписи поля («Aláírás:……sss») или стоит отдельным словом — иначе он съедал
# последнюю букву подписи («s» из «Aláírás»)
_LEADER = re.compile(r"[.,…:;]+[szeégí.,:;…itn0-9]{8,}|(?<![\wÀ-ɏ])[szeégí][szeégí.,:;…itn]{9,}")
_TRIPLE = re.compile(r"([a-zé])\1\1")


def _denoise(tok):
    def rep(m):
        s = m.group()
        # в мусоре пунктира всегда есть буква трижды подряд («eee», «sss») или ряд точек в начале; в словах — нет
        # («szükségessége:» опознавался как пунктир и переводился «Обучение треб…», 06.10)
        noisy = _TRIPLE.search(s) or re.match(r"[.,…:;]{3,}", s)
        return "…" if noisy and sum(c in "szeégí.,:;…0123456789" for c in s) / len(s) >= 0.8 else s
    return _LEADER.sub(rep, tok)


def _sim(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio() if a and b else 0.0


INK_HAND = 0.08     # плотность краски ниже — штрих ручки (0,03–0,05), а не печать (0,11–0,25 даже у кода в скобках)
# подпись поля «Aláírás:», «Dátum:» — всегда печать: пунктир рядом снижает плотность краски, а рукопись двоеточием
# не кончается
LABEL = re.compile(r"^[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]{3,}:")
SYMBOL = re.compile(r"[\w<>≤≥=±÷+%°Δ×→]")
# флажок ☐ Tesseract читает как «L]», «L1]», «[]»
CHECKBOX = re.compile(r"^\[?[LlI1|]{0,2}\]$|^\[$")
# подпись поля, слипшаяся с рукописью или линейкой слева: «oílNév:», «INév:»
GLUED = re.compile(r"^(.+?)([A-ZÁÉÍÓÖŐÚÜŰ][a-záéíóöőúüű]{2,}:)$")


def _boxlike(w, line_h):
    """Флажок ☑/☐/☒, прочитанный буквами («VI», «[5», «£]»): до трёх знаков в почти квадратной рамке, неуверенно,
    и рамка выше букв строки или есть скобка — печатное «A» перед кодом тоже бывает квадратным и с уверенностью 0.
    Такое слово — картинка флажка: не переводится и не закрывается плашкой (32RDGT, 07.10)."""
    x0, y0, x1, y1 = w["box"]
    return len(w["raw"]) <= 3 and w.get("c", 100) < 60 and 0.75 <= (x1 - x0) / max(y1 - y0, 0.1) <= 1.3 \
        and (w["h"] >= 1.15 * line_h or bool(re.search(r"[\[\]£|]", w["raw"])))


def _split_glued(words):
    """«oílNév:» — одно слово Tesseract из рукописной даты и подписи поля: целиком оно сходило за подпись (двоеточие)
    и печаталось вместе с рукописью кеглем 15 (32RDGT, лист регистрации, 07.10). Подпись — отдельным словом,
    рамка делится по числу букв; обрывок линейки («I», «|») отбрасывается."""
    out = []
    for w in words:
        m = GLUED.match(w["t"])
        if not m or re.search(r"[A-ZÁÉÍÓÖŐÚÜŰ]{2}", m.group(1)):
            out.append(w)
            continue
        x0, y0, x1, y1 = w["box"]
        cut = x0 + (x1 - x0) * len(m.group(1)) / len(w["t"])
        if not re.fullmatch(r"[|Il1\[\]]", m.group(1)):
            out.append(dict(w, t=m.group(1), box=[x0, y0, cut, y1]))
        out.append(dict(w, t=m.group(2), box=[cut, y0, x1, y1]))
    return out


def ocr_words(words, portal_text):
    """Печатные слова OCR-страницы с текстом, уточнённым чтением портала. Рукописное (подписи, даты, фамилии
    от руки) и шум пунктира отбрасываются — их не переводят и не закрывают плашкой.
    Рукописное слово: уверенность Tesseract < 60 и мало краски (INK_HAND). Подтверждение порталом рукопись
    не отличает: портал переписывает и рукописные фамилии с датами (титул 30RDGT, 06.10).
    Зона рукописи — рамки рукописных слов; неуверенное (< 90) бледное слово внутри зоны тоже рукописное.
    Возвращает (слова, слов портала без места на странице)."""
    # обрывки пунктира («sss», «see» высотой 0,5–1 pt) — не слова
    ws = [dict(w, t=_denoise(w["t"]), raw=w["t"]) for w in _split_glued(words) if w["h"] >= 2.0]
    # рукопись по геометрии строки: неуверенное слово вдвое выше соседей — рукописная дата/подпись поверх формы
    # (нижняя медиана: из двух слов «медиана» иначе — большее, и высокое слово не считалось высоким)
    by_line = {}
    for w in ws:
        by_line.setdefault(w["line"], []).append(w["h"])
    for w in ws:
        hs_ = sorted(by_line[w["line"]])
        tall = len(hs_) >= 2 and w["h"] > 1.6 * hs_[(len(hs_) - 1) // 2]
        # подпись поля — по исходному тексту: очистка пунктира съедает двоеточие («Aláírás:…»)
        cf, ink = w.get("c", 100), w.get("ink", 1)
        # рукописные цифры Tesseract читает и с уверенностью 67–75 («242.» вместо «2017.»), но краски в них
        # как у рукописи (0,04–0,075): очень бледное — рукопись всегда, бледное — при уверенности ниже 80
        w["hand"] = (ink < 0.05 or (cf < 80 and ink < INK_HAND) or (cf < 60 and tall)) and not LABEL.match(w["raw"])
        w["box_like"] = _boxlike(w, hs_[(len(hs_) - 1) // 2])
        # буллет • читается как «e»: одиночный знак ниже строки и почти сплошная краска (0,66 против 0,2–0,45
        # у букв) — картинка, как флажок; иначе «e» переводилось в «е», а сам буллет закрывался плашкой (32RDGT)
        w["box_like"] |= len(w["raw"]) == 1 and w["h"] < 0.7 * hs_[(len(hs_) - 1) // 2] and ink > 0.55
    # флажки и буллеты — картинки: в сверку с порталом не идут ни у Tesseract, ни у портала — иначе кусок
    # «• RD 3.1.;» портала доставался буллету «e» и пропадал вместе с ним (32RDGT, 07.10)
    ws = [w for w in ws if not w["box_like"]]
    ptoks, pline = [], []
    for li, ln in enumerate(_unmark(portal_text).splitlines()):
        for t in ln.split():
            if not re.fullmatch(r"[•·▪●○◦■□☐☑☒✓✔]+", t):
                ptoks.append(t); pline.append(li)
    a, b = [_norm(w["t"]) for w in ws], [_norm(t) for t in ptoks]
    wline = [None] * len(ws)                         # строка чтения портала, к которой прижато слово
    conf = [False] * len(ws)
    lost = []
    if ptoks:
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        # знак сравнения или единицы у портала на месте неуверенного слова Tesseract («<» прочитан как «c») — подтверждён
        sign = lambda p: not re.search(r"\w", p) and re.search(r"[<>≤≥=±÷%°Δ×→]", p)
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            # в куске — хвост соседней строки чтения портала («-» пункта ниже): «t c 459C.» против «t < 45°C. -» —
            # длины кусков расходились, «c» доставалось «45°C.», а «<» пропадал (3SZ19, 19.10); слова портала — своей
            # строки (той, что у слова перед куском), 1:1
            if op == "replace" and i1 > 0 and wline[i1 - 1] is not None and i2 - i1 != j2 - j1:
                js = [j for j in range(j1, j2) if pline[j] == wline[i1 - 1]]
                if len(js) == i2 - i1 and len({w["line"] for w in ws[i1:i2]}) == 1:
                    for k, j in zip(range(i1, i2), js):
                        w, p = ws[k], ptoks[j]
                        w["t"] = p if w.get("c", 100) < 60 else _choose(w["t"], p)
                        conf[k] = a[k] == b[j] or _sim(a[k], b[j]) >= 0.5 or bool(sign(p))
                        wline[k] = pline[j]
                    lost += [ptoks[j] for j in range(j1, j2) if j not in js]
                    continue
            if op in ("equal", "replace") and i2 - i1 == j2 - j1:
                for k in range(i2 - i1):
                    w, p = ws[i1 + k], ptoks[j1 + k]
                    # неуверенно прочитанное слово портал знает лучше («Dánum:» → «Dátum:»)
                    w["t"] = p if w.get("c", 100) < 60 else _choose(w["t"], p)
                    conf[i1 + k] = op == "equal" or _sim(a[i1 + k], b[j1 + k]) >= 0.5 or bool(sign(p))
                    wline[i1 + k] = pline[j1 + k]
            elif op == "replace" and i2 - i1 <= 3 and j2 - j1 <= 3 and not any(w["hand"] for w in ws[i1:i2]) \
                    and len({w["line"] for w in ws[i1:i2]}) == 1 and len(set(pline[j1:j2])) == 1 \
                    and (any(_codeish(w["t"]) for w in ws[i1:i2]) or any(_codeish(t) for t in ptoks[j1:j2])):
                # код разбит по-разному («5090-nál» у Tesseract, «50%-nál» у портала): текст портала целиком —
                # в первое слово, рамка — на все слова куска
                from layout import _union
                ws[i1]["t"] = " ".join(ptoks[j1:j2])
                ws[i1]["box"] = _union([w["box"] for w in ws[i1:i2]])
                ws[i1]["c"] = max(w.get("c", 100) for w in ws[i1:i2])
                conf[i1] = True
                for k in range(i1 + 1, i2):
                    ws[k]["t"] = ""
            elif op == "replace":
                for k in range(i1, i2):
                    best = max(range(j1, j2), key=lambda j: _sim(a[k], b[j]))
                    if _sim(a[k], b[best]) >= 0.5:
                        conf[k] = True
                        if ws[k].get("c", 100) < 60:
                            ws[k]["t"] = ptoks[best]
            elif op == "insert":
                # знаки и коды, которых Tesseract не увидел («%», «≤»), — к слову слева; рукописные фамилии,
                # прочитанные порталом, не подклеиваем: рукопись остаётся как есть
                codes = [t for t in ptoks[j1:j2] if _codeish(t) or not re.search(r"\w", t)]
                same_line = i1 > 0 and wline[i1 - 1] is not None and all(pline[j] == wline[i1 - 1] for j in range(j1, j2))
                if codes and same_line and len(codes) == j2 - j1:
                    ws[i1 - 1]["t"] += " " + " ".join(codes)
                else:
                    lost += ptoks[j1:j2]
    # рукопись: Tesseract не уверен и краски мало (штрих ручки тонкий: 0,03–0,05 против 0,18–0,25 у печати,
    # включая плохо прочитанные печатные коды «NRc19», «260€C» — замер 06.10)
    # очень неуверенное (< 30) слово, которое портал не подтвердил, — рукопись или шум: рукописная дата
    # «2017.04.21.» у Tesseract — «LJOMT-DÁL», «242. 74. 09» (формы 30RDGT, 06.10); печатные коды с низкой
    # уверенностью («NRc19» → «NR<1%») портал подтверждает
    hand = [w["hand"] or (w.get("c", 100) < 30 and not c and not LABEL.match(w["raw"])) for w, c in zip(ws, conf)]
    zones = [w["box"] for w, h in zip(ws, hand) if h and len(_norm(w["t"])) >= 2]
    near = lambda bx: any(bx[0] < z[2] + 3 and bx[2] > z[0] - 3 and bx[1] < z[3] + 3 and bx[3] > z[1] - 3 for z in zones)
    out = []
    for w, h in zip(ws, hand):
        # знаки сравнения и единиц — смысл («v > 1,0 мм»: Tesseract читает «>» как «5», портал правит на «>»);
        # одиночные точки и запятые пунктира — нет
        if not SYMBOL.search(w["t"]) or h or CHECKBOX.match(w["t"]) or w["box_like"]:
            continue
        if w.get("c", 100) < 90 and w.get("ink", 1) < INK_HAND and near(w["box"]) and not LABEL.match(w["raw"]):
            continue                                   # рядом с рукописью и сам похож на неё («TÓ» в фамилии от руки)
        out.append(w)
    return out, lost


DESC = re.compile(r"[gjpqyÁÉÍÓÖŐÚÜŰ,;]")


def _split_at_rules(words, vs):
    """Слово Tesseract поперёк вертикали таблицы («blokkvezénylőiJA.1.», «szerintijA.2.2.»: линия прочитана буквой,
    текст двух столбцов склеен — ÜFK-скан, 19.10) — двумя словами по линии; буква-линия у места разреза (±2 знака
    от доли ширины) отбрасывается."""
    out = []
    for w in words:
        x0, y0, x1, y1 = w["box"]
        t = w["t"]
        xs = [x for ry0, ry1, x in vs if x0 + 2 < x < x1 - 2 and min(y1, ry1) - max(y0, ry0) > 0.5 * (y1 - y0)]
        if not xs or len(t) < 3:
            out.append(w)
            continue
        x = xs[0]
        s, n = w.get("raw") or t, len(t.strip())
        if s != t and s.startswith(t.strip()) and len(s) > n + 1:
            # чтение портала отбросило хвост за линией («EltérőlA.2.1.» → «Eltérő»), рамка — прежняя: хвост из сырого
            a, b = t, s[n + 1:] if s[n] in "|Jj[]!lIi1" else s[n:]
        else:
            k = round(len(t) * (x - x0) / (x1 - x0))
            near = sorted((i for i in range(max(0, k - 2), min(len(t), k + 3)) if t[i] in "|Jj[]!lI1"),
                          key=lambda i: abs(i - k))
            a, b = (t[:near[0]], t[near[0] + 1:]) if near else (t[:k], t[k:])
        if a:
            out.append(dict(w, t=a, box=[x0, y0, x - 1, y1]))
        if b:
            out.append(dict(w, t=b, box=[x + 1, y0, x1, y1]))
    return out


def ocr_frags(words, vs=()):
    """Фрагменты OCR-страницы: слова одной строки Tesseract подряд, без большого разрыва и без линейки между ними.
    Кегль — по высоте слов без выносных вниз и без прописных с диакритикой (высота такого слова ≈ 0,72 кегля),
    жирность — по плотности краски относительно слов той же высоты."""
    from layout import GAP, CODE_GAP, _vrule_in, _union, code_line
    words = _split_at_rules(words, vs)
    inks = sorted(w["ink"] for w in words if "ink" in w)
    # рамка слова, задетого рукописной подписью, раздута по высоте («Aláírás:» 9,4 pt вместо 3,8): по вертикали —
    # как у обычных слов той же строки, иначе и кегль перевода, и белая плашка вырастают и затирают подпись
    by_line = {}
    for w in words:
        by_line.setdefault(w["line"], []).append(w)
    words = [dict(w, box=list(w["box"])) for w in words]
    for w in words:
        mates = sorted(by_line[w["line"]], key=lambda m: m["h"])
        med = mates[len(mates) // 2]["h"]
        normal = [m for m in mates if m["h"] <= 1.5 * med]
        if w["h"] > 1.6 * med and len(normal) >= 1:
            w["box"][1] = sorted(m["box"][1] for m in normal)[len(normal) // 2]
            w["box"][3] = sorted(m["box"][3] for m in normal)[len(normal) // 2]
            w["h"] = w["box"][3] - w["box"][1]
    out, cur = [], None
    for w in words:
        if cur is not None and cur["line"] == w["line"] and w["box"][0] - cur["bbox"][2] <= GAP \
                and not _vrule_in(vs, cur["bbox"][2], w["box"][0], w["box"][1], w["box"][3]) \
                and not (w["box"][0] - cur["bbox"][2] >= CODE_GAP and code_line(cur["t"]) and code_line(w["t"])):
            cur["t"] += " " + w["t"]; cur["bbox"] = _union([cur["bbox"], w["box"]]); cur["ws"].append(w)
        else:
            if cur is not None:
                out.append(cur)
            cur = {"t": w["t"], "bbox": list(w["box"]), "line": w["line"], "ws": [w]}
    if cur is not None:
        out.append(cur)
    for f in out:
        ws = f.pop("ws"); f.pop("line")
        # кегль — по уверенно прочитанным словам: рукописная дата, исправленная порталом («2025.ot,» → «2025.04.07.»),
        # раздувала кегль строки «Dátum: …» до 15 pt (32RDGT, 07.10)
        sure = [w for w in ws if w.get("c", 100) >= 80] or ws
        plain = sorted(w["h"] for w in sure if not DESC.search(w["t"]) and re.search(r"[a-zá-ű0-9]", w["t"], re.I))
        hs = plain or sorted(w["h"] * 0.8 for w in sure)
        size = hs[len(hs) // 2] / 0.72
        # проверка по ширине: средняя буква ≈ 0,5 кегля — раздутая по высоте рамка (подпись от руки, диакритика
        # над буквами) не даёт кегль 13 у подписи поля и у «Verziószám:»
        chars = sum(len(w["t"]) for w in ws) + len(ws) - 1
        by_width = (f["bbox"][2] - f["bbox"][0]) / (0.5 * max(chars, 1))
        if chars >= 4:
            size = min(size, 1.25 * by_width)
        f["size"] = round(min(max(size, 4.0), 16.0), 1)
        f["key"] = ws[0]["key"]
        same = [w2["ink"] for w2 in words if "ink" in w2 and abs(w2["h"] - ws[0]["h"]) < 0.25 * ws[0]["h"]] or inks
        med = sorted(same)[len(same) // 2] if same else 0
        n = sum(len(w["t"]) for w in ws) or 1
        f["bold"] = med > 0 and sum(len(w["t"]) for w in ws if w.get("ink", 0) > 1.3 * med) > n / 2
        # по одной-двум буквам («A» перед кодом, «0» в таблице) штрих не измерить: жирность от краски, иначе
        # одиночный знак отрывался от абзаца или приклеивался к жирной шапке
        st = sorted(w["stroke"] for w in sure if w.get("stroke"))
        letters = sum(len(re.findall(r"\w", w["t"])) for w in sure)
        f["_st"] = st[len(st) // 2] / f["size"] if st and letters >= 3 else None
    # жирность — по толщине штриха относительно кегля: жирное в 1,3–1,7 раза толще медианы страницы (замер 32RDGT:
    # обычный текст 0,72–0,96 pt, заголовки 1,44–2,64 pt); медиана — по числу букв, кэш без штриха — по краске
    rel = sorted((f["_st"], len(f["t"])) for f in out if f["_st"])
    if rel:
        acc, half = 0, sum(n for _, n in rel) / 2
        for ref, n in rel:
            acc += n
            if acc >= half:
                break
        for f in out:
            if f["_st"]:
                f["bold"] = f["_st"] > 1.3 * ref
    for f in out:
        f.pop("_st")
    # моноширинный шрифт (Courier): буква ≈ 0,6 кегля против 0,42–0,44 у Times (замер 06.10: 3SZ19 — 0,65,
    # 3PR42 и формы 30RDGT — 0,42–0,44); решение по странице, длинные строки — каждая сама
    ratio = lambda f: (f["bbox"][2] - f["bbox"][0]) / (len(f["t"]) * f["size"])
    long_r = sorted(ratio(f) for f in out if len(f["t"]) >= 8)
    mono_page = bool(long_r) and long_r[len(long_r) // 2] > 0.55
    for f in out:
        # одним решением на страницу: по отдельной строке оценка сбивается (цифры шире букв — «TO 140387»
        # выходил моноширинным среди Times)
        f["family"] = "monospace" if mono_page else "serif"
    # кегль внутри абзаца Tesseract сглаживается медианой: разброс оценки по строкам рвал абзацы на строки
    by_par = {}
    for f in out:
        by_par.setdefault(f["key"], []).append(f["size"])
    for f in out:
        sizes = sorted(by_par[f.pop("key")])
        med = sizes[len(sizes) // 2]
        if abs(f["size"] - med) <= 0.25 * med:
            f["size"] = med
    return out




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
                      "h": h * 72 / dpi, "c": float(c[10])})
    return words


def _ink(gray, box, dpi):
    """Доля тёмных пикселей в рамке слова: жирное начертание плотнее при той же высоте."""
    k = dpi / 72
    x0, y0, x1, y1 = (int(v * k) for v in box)
    cut = gray[max(y0, 0):max(y1, y0 + 1), max(x0, 0):max(x1, x0 + 1)]
    return float((cut < 128).mean()) if cut.size else 0.0


def _stroke(gray, box, dpi):
    """Толщина штриха слова, pt: медиана длин горизонтальных отрезков краски. Плотность краски жирное не отличает —
    у прописных она выше сама по себе: заголовки титула 32RDGT не выходили жирными (07.10)."""
    import numpy as np
    k = dpi / 72
    x0, y0, x1, y1 = (int(v * k) for v in box)
    cut = gray[max(y0, 0):max(y1, y0 + 1), max(x0, 0):max(x1, x0 + 1)] < 128
    if not cut.any():
        return 0.0
    d = np.diff(np.pad(cut.astype(np.int8), ((0, 0), (1, 1))), axis=1)
    runs = np.nonzero(d == -1)[1] - np.nonzero(d == 1)[1]
    return round(float(np.median(runs)) / k, 2)


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


def page_missing_share(page, tess_words):
    """Доля настоящих слов Tesseract (от 4 букв, с тремя строчными подряд), которых в текстовом слое нет ни по месту,
    ни по написанию (похожее написание засчитывается: коды Tesseract дробит и искажает — «31RD315401»; кружки отметок
    читает как «oO», рамки — «LL»; 30RDGT feladat, 06.10).
    - Надписи внутри картинок (скриншоты, логотипы) — не потеря слоя, если слой на странице есть: страницы руководства
      со скриншотами и логотипом в колонтитуле уходили в OCR целиком и перерисовывались как сканы (UTS, 17 из 99, 15.10).
    - Слово, место которого занято словом слоя, не потеряно: точки-заполнители оглавления Tesseract читает как
      «eeeteenáááánázsái», «KEKE» — оглавление считалось сканом.
    Замер по кэшу OCR 15 документов: родные страницы — 0,00; сканы и страницы с текстом кривыми — от 0,17;
    3SZ19 стр. 80 (была текстовой) — 0,13.
    - Невидимый текст (слой OCR поверх скана, текст под фоном) слово не «накрывает»: скан с невидимым слоем
      иначе не уходил в OCR, а его текст отбрасывался как скрытый — страница пропадала (красная команда 15.10).
    - Картинка от половины листа — скан (тело страницы, вклеенное в шаблон с текстовым колонтитулом), её слова
      считаются.
    Возвращает (доля, слова Tesseract вне картинок-скриншотов)."""
    import fitz
    from layout import hidden_text, _covered
    rot = page.rotation_matrix
    hidden = hidden_text(page)
    visible = [w for w in page.get_text("words") if not (hidden and _covered(hidden, w[:4]))]
    text = " ".join(w[4] for w in visible)
    words = tess_words
    if len(re.findall(r"\w{2,}", text)) >= 20:
        area = page.rect.get_area()
        imgs = [r for r in (fitz.Rect(i["bbox"]) * rot for i in page.get_image_info()) if r.get_area() < 0.5 * area]
        centre = lambda w: fitz.Point((w["box"][0] + w["box"][2]) / 2, (w["box"][1] + w["box"][3]) / 2)
        words = [w for w in words if not any(r.contains(centre(w)) for r in imgs)]
    boxes = [fitz.Rect(w[:4]) * rot for w in visible]

    def covered(w):
        r = fitz.Rect(w["box"]); a = r.get_area() or 1
        return any((r & b).get_area() > 0.3 * a for b in boxes if b.intersects(r))

    cand = [w for w in words if re.fullmatch(r"\w{4,}", w["t"]) and re.search(r"[a-záéíóöőúüű]{3}", w["t"].lower())]
    if not cand:
        return 0.0, words
    tl = {x.lower() for x in re.findall(r"\w{4,}", text)}
    miss = sum(1 for w in cand if not covered(w) and w["t"].lower() not in tl
               and not difflib.get_close_matches(w["t"].lower(), tl, 1, 0.8))
    return miss / len(cand), words


def page_rules(page, dpi=300):
    """Линейки страницы: векторные (layout.rules) плюс найденные на картинке (сканы форм)."""
    import numpy as np
    import fitz
    from layout import rules, rules_from_image
    hs, vs = rules(page)
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    ih, iv = rules_from_image(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width), dpi)
    return hs + ih, vs + iv


def pdf_segments(path, pages=None, ocr="auto", dpi_tess=300, dpi_portal=200, ratio=0.85, progress=None, shots=True,
                 roles=False):
    """ocr: 'auto' — страница идёт через OCR, если в текстовом слое < ratio слов Tesseract и заметной доли
    настоящих слов Tesseract в слое нет (page_missing_share > 15 %).
    progress(доля 0..1) — ход по страницам: первая половина — Tesseract, вторая — чтение сканов порталом
    (на 146 сканах это 10 минут, а процент стоял на нуле — «зависло», 07.10).
    shots — надписи скриншотов на страницах с текстовым слоем (shots.shot_segments): сегменты идут сразу после
    текста своей страницы, номера сегментов — по порядку.
    roles — листы с текстовым слоем: границы абзацев и роли строк (заголовок, пункт, подпись) — метками портала
    (roles.py; PDF без разметки, режим "text"); сбой портала — лист по правилам."""
    import fitz
    from concurrent.futures import ThreadPoolExecutor
    from layout import frags_from_text, rules, rules_from_image, group, avail_rects, aligns
    doc = fitz.open(path)
    work = Path(OCR_DIR) / Path(path).stem; work.mkdir(parents=True, exist_ok=True)
    rng = list(pages if pages is not None else range(len(doc)))

    def tess(pno):
        """Слова Tesseract (с уверенностью и плотностью краски) и линейки, найденные на картинке; кэш на диске."""
        import numpy as np
        f = work / f"p{pno:04d}.json"
        if f.exists():
            cached = json.loads(f.read_text(encoding="utf-8"))
            if isinstance(cached, dict):            # старый кэш (список слов без уверенности) — пересчитать
                for w in cached["words"]:           # из JSON номера строк приходят списками — ключи словаря
                    w["line"], w["key"] = tuple(w["line"]), tuple(w["key"])
                if all("stroke" in w for w in cached["words"]):
                    return cached
                # кэш без толщины штриха: досчитать по картинке, Tesseract не перезапускать
                pix = fitz.open(path)[pno].get_pixmap(dpi=dpi_tess, colorspace=fitz.csGRAY)
                gray = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
                for w in cached["words"]:
                    w["stroke"] = _stroke(gray, w["box"], dpi_tess)
                f.write_text(json.dumps(cached, ensure_ascii=False), encoding="utf-8")
                return cached
        png = work / f"p{pno:04d}.png"
        pix = fitz.open(path)[pno].get_pixmap(dpi=dpi_tess, colorspace=fitz.csGRAY)
        pix.save(str(png))
        ws = _tess_words(png, dpi_tess); png.unlink()
        gray = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
        for w in ws:
            w["ink"] = round(_ink(gray, w["box"], dpi_tess), 3)
            w["stroke"] = _stroke(gray, w["box"], dpi_tess)
        ih, iv = rules_from_image(gray, dpi_tess)
        out = {"words": ws, "hs": ih, "vs": iv}
        f.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        return out

    import threading
    lock, cnt = threading.Lock(), {"t": 0, "p": 0}

    def counted(fn, key, base, total):
        def run(pno):
            r = fn(pno)
            with lock:                                  # и счётчик, и запись progress.json — по одному потоку
                cnt[key] += 1
                if progress:
                    progress(base + 0.5 * cnt[key] / total)
            return r
        return run

    with ThreadPoolExecutor(4) as ex:
        tpage = dict(zip(rng, ex.map(counted(tess, "t", 0.0, len(rng) or 1), rng)))
    tw = {p: tpage[p]["words"] for p in rng}

    def portal(pno):
        from llm import chat, img_b64
        png = fitz.open(path)[pno].get_pixmap(dpi=dpi_portal, colorspace=fitz.csGRAY).tobytes("png")
        return chat([{"role": "user", "content": OCR_PROMPT}], images=[img_b64(png)], max_tokens=8000, tag=f"ocr{pno}")

    from layout import hidden_text, _covered
    plan = {}
    for pno in rng:
        hid = hidden_text(doc[pno])           # невидимый слой OCR поверх скана — не текст страницы
        n_text = len([w for w in doc[pno].get_text("words") if re.search(r"\w{2,}", w[4])
                      and not (hid and _covered(hid, w[:4]))])
        share, outside = page_missing_share(doc[pno], tw[pno]) if ocr == "auto" else (0.0, tw[pno])
        n_tess = len([w for w in outside if re.search(r"\w{2,}", w["t"])])
        plan[pno] = (ocr is True) or (ocr == "auto" and n_tess >= 15 and n_text < ratio * n_tess and share > 0.15)
    ocr_pages = [p for p in rng if plan[p]]
    with ThreadPoolExecutor(4) as ex:
        pr = dict(zip(ocr_pages, ex.map(counted(portal, "p", 0.5, len(ocr_pages) or 1), ocr_pages)))

    labels, pre = {}, {}
    if roles:
        import roles as R
        text_pg = [p for p in rng if not plan[p]]
        labels = R.doc_labels(path, text_pg)
        # распознанные листы — по кускам OCR (строки собираются из кусков: текстового слоя у скана нет)
        for pno in [p for p in rng if plan[p]]:
            hs_, vs_ = rules(doc[pno])
            hs_, vs_ = hs_ + [tuple(x) for x in tpage[pno]["hs"]], vs_ + [tuple(x) for x in tpage[pno]["vs"]]
            words_, lost_ = ocr_words(_dehyphen(tw[pno]), pr[pno].get("text") or "")
            pre[pno] = (hs_, vs_, lost_, ocr_frags(words_, vs_))
        if pre:
            labels.update(R.frag_labels({p: v[3] for p, v in pre.items()},
                                        {p: (doc[p].rect.width, doc[p].rect.height) for p in pre}))
    segs, errors, lost = [], [], {}
    for pno in rng:
        page = doc[pno]
        hs, vs = rules(page)
        if plan[pno]:
            r = pr[pno]
            # у скана нет векторных линий — линейки таблиц найдены на картинке
            hs, vs = hs + [tuple(x) for x in tpage[pno]["hs"]], vs + [tuple(x) for x in tpage[pno]["vs"]]
            if not r.get("text"):
                errors.append((pno, r.get("error")))
            if pno in pre:
                hs, vs, lost[pno], frags = pre[pno]
            else:
                words, lost[pno] = ocr_words(_dehyphen(tw[pno]), r.get("text") or "")
                frags = ocr_frags(words, vs)
            pars = group(frags, hs, vs, form=True)
            if labels.get(pno) and labels[pno][1]:
                from layout import join_lines
                pars = R.apply(pars, labels[pno][0], labels[pno][1], join_lines)
        else:
            pars = group(frags_from_text(page, vs), hs, vs)
            if labels.get(pno) and labels[pno][1]:
                from layout import join_lines
                pars = R.apply(pars, labels[pno][0], labels[pno][1], join_lines)
        avail_rects(pars, hs, vs, page.rect)
        aligns(pars, vs)
        for p in pars:
            segs.append({"id": len(segs), "text": p["text"], "src": p["text"], "page": pno,
                         "loc": {"bbox": [round(v, 2) for v in p["bbox"]], "avail": [round(v, 2) for v in p["avail"]],
                                 "align": p["align"], "cell": p["cell"],
                                 "frags": [[round(v, 2) for v in f["bbox"]] for f in p["frags"]],
                                 "size": p["size"], "bold": p["bold"], "family": p.get("family", "serif"),
                                 "ocr": plan[pno]}})
            if labels.get(pno) and labels[pno][1]:
                segs[-1]["loc"]["roles"] = True             # абзацы листа — по меткам портала
                if p.get("role"):
                    segs[-1]["loc"]["role"] = p["role"]
    n_shots = 0
    text_pages = {p for p in rng if not plan[p]}
    if shots and text_pages:
        import shots as SH
        extra = SH.shot_segments(doc, work / "shots", len(segs), pages=text_pages)
        n_shots = len(extra)
        order = {p: i for i, p in enumerate(rng)}
        segs = sorted(segs + extra, key=lambda x: (order[x["page"]], "shot" in x["loc"]))
        for i, x in enumerate(segs):
            x["id"] = i
    return segs, {"ocr_pages": ocr_pages, "errors": errors, "lost_words": {k: v for k, v in lost.items() if v},
                  "shot_labels": n_shots}
