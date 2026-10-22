# -*- coding: utf-8 -*-
"""Надписи на скриншотах: растровые картинки на страницах с текстовым слоем (окна программы в руководстве UTS).

Извлечение (shot_segments):
- картинка — скриншот, если она не украшение (стоит на >30 % страниц: логотипы колонтитула), не значок (<3000 pt²)
  и не фон/скан во весь лист;
- Tesseract (hun) по картинке, увеличенной до ~3×, прямо и негативом; рамки строк от детектора текста (PP-OCRv4 det
  через OpenCV DNN), не накрытые словами Tesseract, дочитываются Tesseract-ом по вырезкам — так находятся кнопки и
  вкладки светлым по тёмному, заголовки окон с тенью; слова → надписи (одна строка, разрыв меньше высоты буквы);
- портал видит картинку и список надписей: исправляет чтение по картинке и называет язык каждой надписи. В перевод —
  только венгерские (lang=hu): английские окна (VirtualBox, вход SUSE, веб-страница Softing) остаются как есть;
  коды, имена, числа, даты, сокращения (CAN, PROFIBUS) портал относит к «без языка» — тоже.
Надписи — обычные сегменты (перевод, проверка, таблица вычитки); loc.shot = {xref, box (пиксели картинки), h}.

Сборка (render): заливка рамки надписи цветом фона (медиана по кромке), перевод цветом надписи, кегль по высоте
строки, место вправо — до первого не-фонового столбца; не влезает — ужимается. Картинка заменяется в PDF.
Прежняя «белая плашка поверх надписи» (режим скана на всю страницу) давала грязь — UTS, 15.10.

Язык источника — параметр lang (пара проекта: hu-ru → hu, ru-en → ru): словарь Tesseract и подсказки порталу (LANGS).

Картинки DOCX (add_docx_shots, docx_media; 22.10): растр тела документа (word/media: PNG, JPEG, BMP, TIFF) — так же, но
строки одной подписи (схема, чертёж) склеиваются в одну надпись (_blocks), а перевод переносится по её ширине; сегменты —
сразу после абзаца с картинкой, loc.shot = {media, box, h, t, size[, lines, ts]}. Сборка пишет перерисованную картинку
в ту же запись того же формата и размера в пикселях — место и размер на листе не меняются. Векторные (EMF, WMF, SVG)
и картинки только колонтитулов не трогаются."""
import json, re, difflib, sys, threading, zipfile
from pathlib import Path

DECOR_SHARE, MIN_AREA, FULL = 0.3, 3000, 0.5   # FULL: от половины листа — скан или фон, не скриншот
UPSCALE_TO = 3000          # Tesseract: ширина увеличенной картинки, px (мелкий шрифт интерфейса ~8–10 px)
# детектор строк текста: PP-OCRv4 det (PaddleOCR, Apache-2.0; файл из rapidocr_onnxruntime 1.4.4), 4,7 МБ.
# Увеличение 1,5× (2× — та же полнота на UTS, вдвое дольше), площадь — до 2,5 Мп: память сети ~0,7 ГБ
DET_MODEL = Path(__file__).parent / "models" / "ch_PP-OCRv4_det_infer.onnx"
DET_SCALE, DET_MAX_PX = 1.5, 2.5e6
_det = {"net": None, "error": None, "lock": threading.Lock()}   # сеть OpenCV DNN не потокобезопасна

# Что переводить, решает язык надписи: в перевод — только язык источника (lang=hu). Прежний вопрос «переводить ли» (промпт
# «скриншот на венгерском», «английские слова не переводить») портал решал через раз в обе стороны: английские окна
# VirtualBox, Softing, вход SUSE (UTS стр. 19, 88–89, 95) то переводил, то нет, а роли «Operátor»/«Adminisztrátor» то
# оставлял; язык надписи по картинке он называет устойчиво (замер 15.10 — work/uts_2026-10-15/shots_bench/README.md,
# «Английские окна»)
# Язык источника — из пары проекта ("hu-ru" → hu, "ru-en" → ru): словарь Tesseract и подсказки порталу. Промпт hu —
# прежний слово в слово (ключ кэша портала: уже прочитанные скриншоты не перечитываются)
LANGS = {
    "hu": {"tess": "hun", "name": "венгерский", "by": "венгерским",
           "letters": "венгерские буквы á é í ó ö ő ú ü ű только там, где они есть на картинке",
           "with": "и вместе с кодом или числом: «UTS üzenet», «1 perc»", "other": ("en", "de")},
    # сокращения кириллицей (АРМ, МФУ) — русские: в английском переводе им не место. Словарь — с английским: латиница
    # в русских подписях схем (VESA, 3D) одним rus читалась кириллицей («уеза»), а похожие буквы портал выправляет (_FOLD)
    "ru": {"tess": "rus+eng", "name": "русский", "by": "русским",
           "letters": "латиницу в кодах и словах — VESA, 3D, RS-485 — не заменяй кириллицей и наоборот",
           "with": "и сокращения кириллицей: АРМ, МФУ; вместе с кодом или числом: «АРМ 3», «5 мин»",
           "other": ("en", "de"), "fold": True},
}
LANG_NAMES = {"hu": "венгерский", "ru": "русский", "en": "английский", "de": "немецкий"}
SHOT = "Это скриншот интерфейса программы."
PICTURE = "Это рисунок из документа: схема, чертёж, фото или скриншот."      # картинки DOCX


def check_prompt(lang="hu", what=SHOT):
    """Промпт проверки надписей порталом (шаблон с {labels})."""
    L = LANGS[lang]
    codes = "|".join(f'"{c}"' for c in (lang, *L["other"], "-"))
    others = ", ".join(f"{c} — {LANG_NAMES[c]}" for c in L["other"])
    return "".join([
        what, " Ниже — надписи с него, прочитанные OCR с ", L["by"], " словарём, по одной на строке "
        "в виде «номер: текст».\n{labels}\n\n"
        "Для каждой надписи верни объект JSON {{\"n\": номер, \"t\": \"текст надписи точно как на картинке\", "
        "\"lang\": ", codes, "}}.\n"
        "- t: исправь ошибки OCR по картинке (", L["letters"], "); "
        "если надписи на картинке нет или это мусор OCR — t пустая строка.\n"
        "- lang: язык надписи на картинке: ", lang, " — ", L["name"], " (", L["with"], "), ",
        others, ", \"-\" — без языка: коды и идентификаторы (X0YZ11K500_ZD11, SKRVR_X_TEST), имена "
        "файлов и путей, числа, даты, время, имена людей, сокращения из заглавных латинских букв. OCR читал ", L["by"],
        " словарём и мог исказить английские слова — решай по картинке.\n"
        "Не добавляй надписей, которых нет в списке. Выведи только массив JSON."])


def find_shots(doc):
    """{xref: [(номер страницы, рамка на странице как видна)]} — картинки-скриншоты документа."""
    import fitz
    uses, places = {}, {}
    for page in doc:
        rot = page.rotation_matrix
        area = page.rect.get_area()
        for ii in page.get_image_info(xrefs=True):
            x = ii.get("xref") or 0
            r = fitz.Rect(ii["bbox"]) * rot
            if not x or r.get_area() < MIN_AREA or r.get_area() >= FULL * area:
                continue
            uses.setdefault(x, set()).add(page.number)
            places.setdefault(x, []).append((page.number, r))
    n = max(len(doc), 1)
    return {x: places[x] for x in places if len(uses[x]) <= max(1, DECOR_SHARE * n) and not _masked(doc, x)}


def _masked(doc, xref):
    """Картинка с маской прозрачности — логотип, а не скриншот: логотип «atom erőmű» на титуле UTS стоит отдельной
    картинкой (не повторяется по листам) и уходил в перевод (15.10)."""
    # трафарет (/ImageMask, 1 бит без цветового пространства — подпись, печать JBIG2) — тоже не скриншот:
    # на нём _rgb падал и ронял извлечение всей задачи (красная команда 15.10)
    masks = any(doc.xref_get_key(xref, k)[0] != "null" for k in ("SMask", "Mask"))
    return masks or doc.xref_get_key(xref, "ImageMask")[1] == "true"


def _rgb(doc, xref):
    import fitz
    pix = fitz.Pixmap(doc, xref)
    if pix.alpha or pix.colorspace is None or pix.colorspace.n != 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
        if pix.alpha:
            pix = fitz.Pixmap(pix, 0)
    return pix


def _np(pix):
    import numpy as np
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width * pix.n]\
        .reshape(pix.height, pix.width, pix.n)


def _tsv(png, lang="hu", psm=11):
    """Слова Tesseract (словарь языка источника; psm 11 — разреженный текст) с листа png: [(текст, x, y, w, h,
    уверенность)] — с буквой или цифрой, от 30."""
    import subprocess
    r = subprocess.run(["tesseract", str(png), "stdout", "-l", LANGS[lang]["tess"], "--psm", str(psm), "tsv"],
                       capture_output=True, text=True, encoding="utf-8")
    out = []
    for ln in r.stdout.splitlines()[1:]:
        c = ln.split("	")
        if len(c) < 12 or c[0] != "5" or not c[11].strip() or float(c[10]) < 30 or not re.search(r"\w", c[11]):
            continue
        out.append((c[11], *(int(v) for v in c[6:10]), float(c[10])))
    return out


def _hit(box, words):
    return [o for o in words if min(box[2], o["box"][2]) > max(box[0], o["box"][0])
            and min(box[3], o["box"][3]) > max(box[1], o["box"][1])]


def det_boxes(img):
    """Рамки строк текста [x0, y0, x1, y1] (пиксели img, RGB) — детектор DB (PP-OCRv4 det) через OpenCV DNN.
    Разбор карты как у DB: порог 0.3, связные области со средней уверенностью от 0.5, расширение на площадь·1.6/периметр."""
    import cv2
    import numpy as np
    h, w = img.shape[:2]
    k = min(DET_SCALE, (DET_MAX_PX / max(h * w, 1)) ** 0.5)
    H, W = max(32, round(h * k / 32) * 32), max(32, round(w * k / 32) * 32)
    x = cv2.resize(img[:, :, ::-1], (W, H), interpolation=cv2.INTER_CUBIC).astype(np.float32) / 127.5 - 1.0   # BGR
    with _det["lock"]:
        if _det["net"] is None:
            _det["net"] = cv2.dnn.readNetFromONNX(str(DET_MODEL))
        _det["net"].setInput(x.transpose(2, 0, 1)[None].copy())
        pred = _det["net"].forward()[0, 0].copy()
    n, lab, st, _ = cv2.connectedComponentsWithStats(cv2.dilate((pred > 0.3).astype(np.uint8), np.ones((2, 2), np.uint8)),
                                                     connectivity=4)
    out = []
    for i in range(1, n):
        x0, y0, bw, bh = (int(v) for v in st[i][:4])
        if min(bw, bh) < 3 or pred[y0:y0 + bh, x0:x0 + bw][lab[y0:y0 + bh, x0:x0 + bw] == i].mean() < 0.5:
            continue
        d = bw * bh * 1.6 / (2 * (bw + bh))
        out.append([max(0.0, (x0 - d) * w / W), max(0.0, (y0 - d) * h / H),
                    min(float(w), (x0 + bw + d) * w / W), min(float(h), (y0 + bh + d) * h / H)])
    return out


def _covered(box, words):
    """Доля ширины рамки детектора под словами Tesseract (неуверенные — мусор во всю кнопку — не в счёт)."""
    iv = sorted((max(o["box"][0], box[0]), min(o["box"][2], box[2])) for o in _hit(box, words)
                if o["c"] >= 50 and min(o["box"][3], box[3]) - max(o["box"][1], box[1]) > 0.3 * (box[3] - box[1]))
    tot, end = 0.0, box[0]
    for a, b in iv:
        tot += max(0.0, b - max(a, end)); end = max(end, b)
    return tot / max(box[2] - box[0], 1)


def det_words(img, boxes, work_png, lang="hu"):
    """Слова в рамках детектора. Вырезка (бинаризация Otsu: белый заголовок с тенью на градиенте в сером Tesseract не
    читает; высота ~40 px) — прямо и негативом, все столбиком на одном листе, один вызов Tesseract; у рамки берётся
    вариант с большей средней уверенностью. Возвращает слова [{t, box (пиксели img), h, c}]."""
    import cv2
    import numpy as np
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    H, W = gray.shape
    crops, y = [], 20
    for i, b in enumerate(boxes):
        x0, y0, x1, y1 = max(int(b[0]) - 2, 0), max(int(b[1]) - 2, 0), min(int(b[2]) + 3, W), min(int(b[3]) + 3, H)
        if x1 <= x0 or y1 <= y0:
            continue
        s = min(5.0, max(1.5, 40.0 / (y1 - y0)))
        c = cv2.resize(gray[y0:y1, x0:x1], None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)
        c = cv2.threshold(c, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        for v, cc in enumerate((c, 255 - c)):
            # почти сплошь тёмные строки и столбцы у краёв вырезки — рамка кнопки или вкладки: в рамке Tesseract не
            # читает надпись («Eredmények» на выделенной вкладке, UTS стр. 27). Стираются линии, а не связные
            # области: акцент «Á» касается линии шапки таблицы, и область уносила букву («FORRÁS» → «FORR»).
            # Только у варианта со светлым фоном: у тёмного «рамка» — сам фон, остаётся полоса, и лист не читался
            dark = cc == 0
            if dark.mean() < 0.5:
                hh, ww = cc.shape
                side = min(0.35 * ww, 0.6 * hh)
                rows = np.nonzero(dark.mean(axis=1) >= 0.7)[0]
                rows = rows[(rows < 0.35 * hh) | (rows >= 0.65 * hh)]
                dark = dark.copy(); dark[rows, :] = False   # столбцы — уже без строк рамки: с ними ножка «P» шла за край
                cols = np.nonzero(dark.mean(axis=0) >= 0.7)[0]
                cols = cols[(cols < side) | (cols >= ww - side)]
                if len(rows) or len(cols):
                    cc = cc.copy(); cc[rows, :] = 255; cc[:, cols] = 255
            crops.append((cc, x0, y0, s, y, i, v))
            y += cc.shape[0] + 24
    if not crops:
        return []
    sheet = np.full((y + 20, max(c[0].shape[1] for c in crops) + 40), 255, np.uint8)
    for c, _, _, _, yy, _, _ in crops:
        # кромка вырезки — цветом её фона: у тёмной кнопки на белом листе Tesseract видел рамку
        sheet[yy - 6:yy + c.shape[0] + 6, 14:26 + c.shape[1]] = int(np.median(c))
        sheet[yy:yy + c.shape[0], 20:20 + c.shape[1]] = c
    cv2.imwrite(str(work_png), sheet)
    got = {}
    for t, x, yw, w, h, conf in _tsv(work_png, lang):
        n = len(re.sub(r"\W", "", t))
        if n < 2 or n <= 3 and conf < 75:          # значки, разделители вкладок и крестик окна читаются коротко и неуверенно
            continue
        for c, x0, y0, s, yy, i, v in crops:
            if yy <= yw + h / 2 <= yy + c.shape[0]:
                got.setdefault((i, v), []).append({"t": t, "box": [x0 + (x - 20) / s, y0 + (yw - yy) / s,
                                                                   x0 + (x - 20 + w) / s, y0 + (yw - yy + h) / s],
                                                   "h": h / s, "c": conf})
                break
    words = []
    for i in {k[0] for k in got}:
        a, b = got.get((i, 0), []), got.get((i, 1), [])
        words += a if (sum(w["c"] for w in a) / len(a) if a else -1) >= (sum(w["c"] for w in b) / len(b) if b else -1) else b
    return words


def tess_labels(img, work_png, lang="hu"):
    """Надписи картинки (массив RGB) по Tesseract и детектору: [{t, box (пиксели исходной картинки), h, c}]."""
    import cv2
    k = max(1.0, min(4.0, UPSCALE_TO / max(img.shape[1], 1)))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    big = cv2.resize(gray, None, fx=k, fy=k, interpolation=cv2.INTER_CUBIC)
    words = []
    # второй проход по негативу: светлый текст на тёмных кнопках («LECSATLAKOZÁS») Tesseract не видит;
    # слово негатива берётся, если на его месте прямой проход ничего не нашёл
    for img_k in (big, 255 - big):
        cv2.imwrite(str(work_png), img_k)
        first = list(words)
        for t, x, y, w, h, c in _tsv(work_png, lang):
            box = [x / k, y / k, (x + w) / k, (y + h) / k]
            if not _hit(box, first):
                words.append({"t": t, "box": box, "h": h / k, "c": c})
    # третий источник — рамки детектора, которые слова Tesseract не накрыли (UTS: кнопки «Mentés», «Mégsem», вкладка
    # «Tesztelés», заголовок «Teszt utasítás lezárása»; полнота на 14 скриншотах 84 → 98 %, 15.10). Уверенное слово
    # из вырезки вытесняет неуверенный мусор Tesseract на своём месте («ém]» во всю кнопку «Mégsem»), но не похожее на
    # него неуверенное слово: верное «intézkedései» (34) менялось на «Intézkedésel» (86), и портал уже не мог исправить
    if _det["error"] is None:
        try:
            boxes = [b for b in det_boxes(img) if _covered(b, words) < 0.8]
            for w in det_words(img, boxes, work_png, lang) if boxes else []:
                hit = _hit(w["box"], words)
                if hit and w["c"] >= 80 and all(o["c"] < 50 for o in hit) and difflib.SequenceMatcher(
                        None, w["t"].lower(), " ".join(o["t"] for o in hit).lower()).ratio() < 0.6:
                    words = [o for o in words if all(o is not x for x in hit)]
                    hit = []
                if not hit:
                    words.append(w)
        except cv2.error as e:              # модель не читается этой сборкой OpenCV — дальше только Tesseract
            _det["error"] = f"{type(e).__name__}: {e}"[:300]
            print(f"shots: детектор надписей отключён — {_det['error']}", file=sys.stderr, flush=True)
    return _labels(words)


def _labels(words):
    """Слова → надписи: одна линия, разрыв меньше высоты строки. Линии — по близости середин по высоте: слова одной
    строки из разных проходов расходятся на 1–2 px, и округление по 4 px клало «Ideiglenesen» и «lezár» в разные строки."""
    lines = []
    for w in sorted(words, key=lambda w: (w["box"][1] + w["box"][3]) / 2):
        yc = (w["box"][1] + w["box"][3]) / 2
        if lines and abs(yc - lines[-1][0]) < 0.35 * min(w["h"], lines[-1][1]):
            lines[-1][2].append(w)
        else:
            lines.append([yc, w["h"], [w]])
    words = [w for ln in lines for w in sorted(ln[2], key=lambda w: w["box"][0])]
    labels = []
    for w in words:
        best = None
        for lb in labels[-8:]:
            b = lb["box"]
            ov = min(b[3], w["box"][3]) - max(b[1], w["box"][1])
            h = min(b[3] - b[1], w["h"])
            if ov > 0.5 * h and 0 <= w["box"][0] - b[2] <= 1.0 * max(lb["h"], w["h"]):
                best = lb
        if best:
            best["t"] += " " + w["t"]
            b = best["box"]
            best["box"] = [min(b[0], w["box"][0]), min(b[1], w["box"][1]), max(b[2], w["box"][2]), max(b[3], w["box"][3])]
            best["c"] = min(best["c"], w["c"])
        else:
            labels.append({"t": w["t"], "box": list(w["box"]), "h": w["h"], "c": w["c"]})
    for lb in labels:
        lb["h"] = lb["box"][3] - lb["box"][1]
        lb["box"] = [round(v, 1) for v in lb["box"]]
    return labels


def portal_png(pix):
    """Картинка для портала: мелкий шрифт интерфейса — увеличение до ~1600 px по ширине."""
    import fitz
    k = max(1.0, min(2.0, 1600 / max(pix.width, 1)))
    return (fitz.Pixmap(pix, int(pix.width * k), int(pix.height * k)) if k > 1 else pix).tobytes("png")


_OBJ_START = re.compile(r'\{\s*"')
# объект, который не разбирается JSON-ом, — по ключам: t — всё между "t": " и ", "lang" (жадно, в пределах объекта)
_OBJ_LOOSE = re.compile(r'\{\s*"n"\s*:\s*"?(\d+)"?\s*,\s*"t"\s*:\s*"(.*)"\s*,\s*"lang"\s*:\s*"([^"]*)"\s*\}', re.S)
FOLLOWUPS = 3              # дозапросов по недостающим надписям после оборванного ответа
_FOLD = str.maketrans("aeopcyxAEOPCYXBKMHT", "аеорсухАЕОРСУХВКМНТ")     # латиница, похожая на кириллицу


def _answer_objects(text):
    """{n: объект} из ответа портала. По объекту, не массивом: оборванный ответ отдаёт всё, что успел (разбор массива
    целиком терял всю картинку). Объект — от «{"» до следующего: регулярка «{без скобок}» теряла надпись с фигурными
    скобками (GUID, UTS p67), а json.loads — с прямыми кавычками внутри строки («„W aktiv", „Y aktiv" lassú…», p43)."""
    dec = json.JSONDecoder()
    starts = [m.start() for m in _OBJ_START.finditer(text)] + [len(text)]
    got = {}
    for a, b in zip(starts, starts[1:]):
        try:
            o = dec.raw_decode(text, a)[0]
        except ValueError:
            m = _OBJ_LOOSE.match(text, a, b)
            if not m:
                continue                     # оборванный хвост ответа
            o = {"n": int(m[1]), "t": m[2], "lang": m[3]}
        try:
            got[int(o["n"])] = o
        except (ValueError, TypeError, KeyError):
            continue
    return got


def _script(w):
    return "c" if re.search(r"[А-Яа-яЁё]", w) else "l" if re.search(r"[A-Za-z]", w) else ""


def _merge(ocr, new):
    """Уверенное чтение Tesseract, поправленное порталом пословно: слово портала берётся, если отличается от слова OCR
    только знаками (± вместо подчёркнутого «+», кавычки «„для») или алфавитом (rus+eng: «уеза» вместо «vesa», «Moy» вместо
    «МФУ»); прочие правки — нет (опечатка оригинала «направяющей» остаётся: рисунок так написан)."""
    a, b = ocr.split(), new.split()
    if len(a) != len(b):
        return ocr
    core = lambda w: re.sub(r"[\W_]", "", w.translate(_FOLD))
    # смена алфавита — у слова той же длины: мусор «ao» иначе получал чужое «проводов.» из сдвинутого ответа
    swap = lambda x, y: _script(x) and _script(y) and _script(x) != _script(y) and abs(len(core(x)) - len(core(y))) <= 1
    return " ".join(y if core(x) == core(y) or swap(x, y) else x for x, y in zip(a, b))


def _realign(labels, got):
    """Ответ портала по рисунку иногда идёт своим порядком чтения, а не по номерам списка (договор, 22.10: со 2-й надписи
    всё сдвинуто на одну — «2. В конструкцию…» получил «2400-2500» и выпал как число). Ответ, непохожий на надпись
    своего номера, переносится на надпись, на которую похож (от 0,8) и которой не досталось похожего ответа; пустые
    ответы (мусор) — на свободные номера как есть."""
    sim = lambda a, b: difflib.SequenceMatcher(None, a.translate(_FOLD).lower(), b.translate(_FOLD).lower()).ratio()
    out = {n: o for n, o in got.items() if 1 <= n <= len(labels) and sim(str(o.get("t", "")), labels[n - 1]["t"]) >= 0.6}
    for n, o in got.items():
        t = str(o.get("t", "")).strip()
        if t and out.get(n) is not o:
            s, m = max(((sim(t, lb["t"]), i + 1) for i, lb in enumerate(labels) if i + 1 not in out), default=(0, 0))
            if s >= 0.8:
                out[m] = dict(o, n=m)
    for n, o in got.items():
        if n not in out and not str(o.get("t", "")).strip():
            out[n] = o
    return out


def check_labels(png, labels, tag, lang="hu", what=SHOT, merge=False):
    """Портал по картинке (PNG, portal_png) исправляет чтение и решает, что переводить (надписи на языке источника lang).
    Сбой портала — надписи остаются непереведёнными (на скриншоте лучше венгерская надпись, чем неверно прочитанная).
    merge — рисунки DOCX: уверенное чтение портал правит пословно (_merge); скриншоты PDF — как замерено 15.10."""
    from llm import chat, img_b64
    prompt = check_prompt(lang, what)

    def ask(part, tg):
        text = "\n".join(f"{i + 1}: {lb['t']}" for i, lb in enumerate(part))
        r = chat([{"role": "user", "content": prompt.format(labels=text)}], images=[img_b64(png)],
                 max_tokens=6000, tag=tg)
        if not r.get("text"):
            return None, None
        return _answer_objects(r["text"]), r.get("finish")

    if not labels:
        return []
    got, finish = ask(labels, tag)
    if got is None:
        return None                          # сбой портала: не кэшировать «нечего переводить»
    if merge:
        got = _realign(labels, got)
    miss = [i for i in range(len(labels)) if i + 1 not in got]
    # модель сама уходит в рассуждения и обрывает ответ на 6000 токенах (UTS: 2 из 71 скриншота, 15.10) — недостающие
    # надписи дозапросом; обрывается и он (p26: 51 надпись, 10 + 19 за два ответа) — снова, пока прибавляет
    for k in range(FOLLOWUPS):
        if not miss or finish != "length":
            break
        more, finish = ask([labels[i] for i in miss], tag + "+" * (k + 1))
        for j, i in enumerate(miss):
            if more and j + 1 in more:
                got[i + 1] = dict(more[j + 1], n=i + 1)
        left = [i for i in miss if i + 1 not in got]
        if len(left) == len(miss):
            break                            # не прибавил: повтор того же запроса — тот же ответ из кэша
        miss = left
    out = []
    for i, lb in enumerate(labels):
        o = got.get(i + 1)
        if not o or o.get("lang") != lang or not str(o.get("t", "")).strip():
            continue
        # уверенное чтение Tesseract портал не переписывает: он «исправлял» надпись по соседней кнопке
        # (FELCSATLAKOZÁS → LECSATLAKOZÁS — обратный смысл, UTS стр. 26, 15.10); неуверенное — только похожим.
        # Замена латиницы на похожую кириллицу («APM» → «АРМ») — всегда: Tesseract rus+eng читает русское
        # сокращение латиницей уверенно (для венгерского ничего не меняет: латиница в латиницу переходит взаимно однозначно)
        t = str(o["t"]).strip()
        ft, fo = t.translate(_FOLD), lb["t"].translate(_FOLD)
        if ft != fo and (lb["c"] >= 85 or difflib.SequenceMatcher(None, ft.lower(), fo.lower()).ratio() < 0.6):
            t = _merge(lb["t"], t) if merge else lb["t"]
            if merge and len(re.findall(r"[^\W\d_]", t)) < 2:
                continue                     # осталось уверенное чтение-мусор («Е]» на месте «Инстр + 3D»)
        if merge and LANGS[lang].get("fold"):
            # слово целиком из латинских двойников кириллицы в русской надписи — кириллица: «APM» (так прочли и Tesseract,
            # и портал) иначе шло в перевод латинским сокращением и оставалось как есть
            t = " ".join(w.translate(_FOLD) if re.fullmatch(r"[aeopcyxAEOPCYXBKMHT]+", re.sub(r"\W", "", w) or "-") else w
                         for w in t.split())
        out.append(dict(lb, t=t, ocr=lb["t"]))
    return out


def shot_segments(doc, work, start_id, pages=None, progress=None, lang="hu"):
    """Сегменты надписей всех скриншотов документа (кэш чтения — work/shot_<xref>.json; не hu — shot_<xref>.<lang>.json)."""
    from concurrent.futures import ThreadPoolExecutor
    work = Path(work); work.mkdir(parents=True, exist_ok=True)
    shots = {x: pl for x, pl in find_shots(doc).items() if pages is None or pl[0][0] in pages}
    sfx = "" if lang == "hu" else f".{lang}"

    xs = sorted(shots, key=lambda x: shots[x][0][0])
    # всё, что трогает MuPDF (картинка, PNG для портала), — в этом потоке: документ PyMuPDF не потокобезопасен;
    # в пуле — только Tesseract (отдельный процесс) и запрос к порталу
    res, todo = {}, []
    for x in xs:
        f = work / f"shot_{x}{sfx}.json"
        if f.exists():
            res[x] = json.loads(f.read_text(encoding="utf-8"))
        else:
            try:
                pix = _rgb(doc, x)
                todo.append((x, _np(pix).copy(), portal_png(pix), [pix.width, pix.height]))
            except (ValueError, RuntimeError) as e:     # необычная картинка — остаётся как есть
                res[x] = {"size": [0, 0], "labels": [], "error": f"{type(e).__name__}: {e}"[:200]}

    def one(job):
        x, img, png, size = job
        try:
            labels = tess_labels(img, work / f"shot_{x}.png", lang)
            (work / f"shot_{x}.png").unlink(missing_ok=True)
            good = check_labels(png, labels, f"shot{x}", lang)
        except Exception as e:              # noqa: BLE001 — одна картинка не валит извлечение задачи
            return {"size": size, "labels": [], "error": f"{type(e).__name__}: {e}"[:200]}
        if good is None:                    # портал не ответил — в кэш не пишем, следующий прогон спросит снова
            return {"size": size, "labels": [], "error": "портал не ответил"}
        out = {"size": size, "labels": good, "ocr": len(labels)}
        (work / f"shot_{x}{sfx}.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        return out

    with ThreadPoolExecutor(4) as ex:
        for i, (job, r) in enumerate(zip(todo, ex.map(one, todo))):
            res[job[0]] = r
            if progress:
                progress((i + 1) / max(len(todo), 1))
    segs = []
    for x in xs:
        pno, rect = shots[x][0]
        for lb in res[x]["labels"]:
            segs.append({"id": start_id + len(segs), "text": lb["t"], "src": lb["t"], "page": pno,
                         "loc": {"shot": {"xref": x, "box": lb["box"], "h": lb["h"], "t": lb["t"], "size": res[x]["size"]},
                                 "bbox": [round(v, 2) for v in rect], "ocr": False}})
    return segs


# ------------------------------------------------------------------ сборка
def _bg_color(img, b, pad=1, exact=False):
    """Цвет фона надписи — пик сглаженной гистограммы яркости рамки с кромкой (буквы — меньшинство). Медиана по кромке
    и мода по корзинам цвета ошибались: серый фон шапки с шумом сглаживания делился на две корзины, и побеждала
    белая линия-разделитель — шапка (белое на сером) перерисовывалась тёмным по белому (UTS, 15.10).
    exact (рисунки DOCX) — самый частый цвет у пика по полосам в 8 px слева и справа от рамки на её строках: среднее
    с полутонами букв (254,255,253) на белом листе оставляло на месте стёртого слова пятно, а в рамке, ужатой по краске
    (_tight), белых букв шапки таблицы больше, чем серого фона (и кольцо вокруг задевало белую строку под шапкой), —
    шапка перерисовывалась белой плашкой."""
    import numpy as np
    H, W = img.shape[:2]
    y0, y1 = max(int(b[1]), 0), min(int(b[3]) + 1, H)
    sides = [img[y0:y1, max(int(b[0]) - 8, 0):max(int(b[0]), 0)], img[y0:y1, min(int(b[2]) + 1, W):min(int(b[2]) + 9, W)]]
    if exact and sum(s.size for s in sides):
        px = np.concatenate([s.reshape(-1, 3) for s in sides]).astype(float)
    else:
        x0, y0, x1, y1 = max(int(b[0]) - pad, 0), max(int(b[1]) - pad, 0), min(int(b[2]) + pad + 1, W), min(int(b[3]) + pad + 1, H)
        px = img[y0:y1, x0:x1].reshape(-1, 3).astype(float)
    lum = px @ np.array([0.299, 0.587, 0.114])
    hist = np.bincount(np.clip(lum, 0, 255).astype(int), minlength=256).astype(float)
    peak = int(np.convolve(hist, np.ones(17), mode="same").argmax())
    if exact:
        vals, cnt = np.unique(px[np.abs(lum - peak) <= 12], axis=0, return_counts=True)
        return vals[cnt.argmax()]
    return px[np.abs(lum - peak) <= 12].mean(axis=0)


def _ink(img, b, bg, q=0.6):
    """Пиксели букв в рамке: (маска, цвет букв — медиана пикселей не ближе q·max к фону). Мелкий тонкий шрифт схем
    (рисунки DOCX) — q=0.85: при 0.6 в медиану шли полутона сглаживания, и чёрная подпись перерисовывалась серой."""
    import numpy as np
    cut = img[int(b[1]):int(b[3]) + 1, int(b[0]):int(b[2]) + 1].astype(float)
    d = np.abs(cut - bg).sum(axis=2)
    if not d.size or d.max() <= 30:
        return np.zeros(d.shape, bool), np.array([0.0, 0.0, 0.0])
    m = d >= 0.5 * d.max()
    return m, np.median(cut[d >= q * d.max()], axis=0)


def _stroke(mask):
    """Толщина штриха — медиана горизонтальных отрезков букв, px."""
    import numpy as np
    if not mask.any():
        return 0.0
    d = np.diff(np.pad(mask.astype(np.int8), ((0, 0), (1, 1))), axis=1)
    return float(np.median(np.nonzero(d == -1)[1] - np.nonzero(d == 1)[1]))


def _bold(mask, h):
    """Жирное — толщина штриха от 0,16 высоты строки."""
    return h > 0 and _stroke(mask) >= 0.16 * h


DESC = r"[gjpqy,;руфдцщ]"          # буквы с выносным элементом вниз (кириллица — русский источник)


def _em(text):
    """Высота рамки Tesseract в долях кегля: от верха (прописные с акцентом выше) до низа (выносные)."""
    top = 0.9 if re.search(r"[ÁÉÍÓÖŐÚÜŰЁЙ]", text) else 0.73 if re.search(r"[A-ZÁ-Űbdfhklt0-9áéíóöőúüűА-ЯЁбёйф]", text) \
        else 0.53
    return top + (0.21 if re.search(DESC, text) else 0.0)


def _free_right(img, b, bg, limit, ink=1):
    """Сколько места вправо от рамки, пока столбец пикселей в полосе надписи — фон (не фона в нём меньше ink пикселей:
    на скане чертежа одна соринка останавливала поиск, и подпись «Тяговое усилие:» ужималась втрое)."""
    import numpy as np
    H, W = img.shape[:2]
    y0, y1 = max(int(b[1]), 0), min(int(b[3]) + 1, H)
    x = int(b[2]) + 1
    while x < min(W, int(limit)):
        col = img[y0:y1, x].astype(float)
        if (np.abs(col - bg).sum(axis=1) > 60).sum() >= ink:
            break
        x += 1
    return x - 2


def _free_left(img, b, bg, limit, ink=1):
    """Сколько места влево от рамки, пока столбец пикселей в полосе надписи — фон (см. _free_right)."""
    import numpy as np
    H, W = img.shape[:2]
    y0, y1 = max(int(b[1]), 0), min(int(b[3]) + 1, H)
    x = int(b[0]) - 1
    while x > max(0, int(limit)):
        col = img[y0:y1, x].astype(float)
        if (np.abs(col - bg).sum(axis=1) > 60).sum() >= ink:
            break
        x -= 1
    return x + 2


def _speck(b):
    """Соринка скана, а не препятствие: меньше пятой части высоты строки (не меньше 3 px) в столбце."""
    return max(3, round(0.2 * (b[3] - b[1])))


def _draw(text, size, bold, fg, bg, width, center=False, height=None, pitch=None):
    """Строка перевода на фоне bg: (массив пикселей, строка базовой линии, итоговый кегль). height — подпись в несколько
    строк (_blocks): перевод переносится по ширине с шагом строк pitch и ужимается, если не влез в height; базовая
    линия — первой строки."""
    import fitz
    col = "rgb(%d,%d,%d)" % tuple(int(v) for v in fg)
    wrap = height is not None
    ph = int(height if wrap else size * 1.5) + 4
    css = (f"* {{font-family: sans-serif; font-size: {size:.2f}px; line-height: {f'{pitch / size:.3f}' if wrap else '1.2'};"
           f" color: {col};{'' if wrap else ' white-space: nowrap;'}"
           f"{' font-weight: bold;' if bold else ''}{' text-align: center;' if center else ''}}}")
    doc = fitz.open(); pg = doc.new_page(width=max(width, 4), height=ph)
    pg.draw_rect(pg.rect, color=None, fill=tuple(float(v) / 255 for v in bg))
    spare, scale = pg.insert_htmlbox(fitz.Rect(0, 2, max(width, 4), ph), text, css=css, scale_low=0.35)
    if spare < 0:
        return None, 0, 0                   # не влезает даже на 35 % — надпись остаётся оригинальной
    arr = _np(pg.get_pixmap(dpi=72, alpha=False))
    # базовая линия — по низу «Н» того же кегля, нарисованной чёрным (цветом надписи белая «Н» на белом не видна)
    ref = fitz.open(); rp = ref.new_page(width=size * 2 + 4, height=ph)
    rp.insert_htmlbox(fitz.Rect(0, 2, size * 2 + 4, ph), "Н",
                      css=css.replace(f"{size:.2f}px", f"{size * scale:.2f}px").replace(col, "rgb(0,0,0)"))
    import numpy as np
    r = _np(rp.get_pixmap(dpi=72, alpha=False))
    rows = np.nonzero((r.astype(int).sum(axis=2) < 600).any(axis=1))[0]
    base = int(rows.max()) + 1 if len(rows) else int(size)
    return arr, base, size * scale


def render(doc, xref, items):
    """Картинка xref с переведёнными надписями. items: [(loc.shot, перевод)]. Возвращает Pixmap RGB."""
    import fitz
    import numpy as np
    img = render_img(_np(_rgb(doc, xref)).copy(), items)
    return fitz.Pixmap(fitz.csRGB, img.shape[1], img.shape[0], np.ascontiguousarray(img).tobytes(), False)


def _plan_block(img, sh, ru, boxes):
    """Подпись в несколько строк (картинки DOCX, _blocks): стираются все строки, перевод переносится по ширине подписи
    с тем же шагом строк; вправо — до чужой надписи или не-фона в полосе любой строки."""
    import numpy as np
    W = img.shape[1]
    lines, ts, b = sh["lines"], sh["ts"], sh["box"]
    bg = _bg_color(img, lines[0], exact=True)
    mask, fg = _ink(img, lines[0], bg, 0.85)
    size = float(np.median([(ln[3] - ln[1]) / _em(t) for ln, t in zip(lines, ts)]))
    bases = [ln[3] - (0.21 * size if re.search(DESC, t) else 0.0) for ln, t in zip(lines, ts)]
    pitch = max(float(np.median(np.diff(bases))), 0.9 * size)
    lim = []
    for ln in lines:
        nxt = min([o[0] for o in boxes if o is not b and o[0] > ln[2] and min(o[3], ln[3]) - max(o[1], ln[1]) > 0] + [W])
        lim.append(_free_right(img, ln, bg, min(nxt - 2, b[2] + (b[2] - b[0]) * 0.5 + 20), _speck(ln)))
    x0, x1 = b[0], max(min(lim), b[2])
    # строки по середине (заголовок в две строки) — перевод по центру той же середины
    mids = [(ln[0] + ln[2]) / 2 for ln in lines]
    center = max(mids) - min(mids) < 0.6 * size and max(ln[0] for ln in lines) - b[0] > size
    if center:
        left = min(max(_free_left(img, ln, bg, b[0] - (b[2] - b[0]) * 0.5 - 20, _speck(ln)) for ln in lines), b[0])
        mid = (b[0] + b[2]) / 2
        half = min(mid - left, x1 - mid)
        x0, x1 = mid - half, mid + half
    return {"b": b, "lines": lines, "bg": bg, "fg": fg, "stroke": _stroke(mask), "bold": False, "size": size,
            "x0": x0, "x1": x1, "center": center, "base": bases[0], "ru": ru, "pitch": pitch,
            "height": len(lines) * pitch}


def render_img(img, items, pic=False):
    """Массив RGB картинки с переведёнными надписями (на месте). items: [(loc.shot, перевод)].
    pic — рисунок DOCX (схема, чертёж): тонкие линии (подчёркивание выносок) не стираются, цвет букв — по самым тёмным
    пикселям, сглаживание стирается целиком (на белом от старого слова оставалась серая тень; договор, 22.10)."""
    import cv2
    import numpy as np
    H, W = img.shape[:2]
    boxes = [sh["box"] for sh, _ in items]
    if pic:
        # линии, которые не стираются вместе с буквами, — по связным кускам маски: кусок внутри рамки надписи — штрих
        # крупной буквы (трафаретный заголовок чертежа: ножки «M», «T» длиннее LINE_PX оставались), выходящий за
        # рамку — подчёркивание или выноска
        _, comp, st, _ = cv2.connectedComponentsWithStats(_line_mask(img).astype(np.uint8), connectivity=8)
    plan = []
    for sh, ru in items:
        if sh.get("lines"):
            plan.append(_plan_block(img, sh, ru, boxes))
            continue
        b = sh["box"]
        bg = _bg_color(img, b, exact=pic)
        mask, fg = _ink(img, b, bg, 0.85 if pic else 0.6)
        src = sh.get("t", "")
        size = (b[3] - b[1]) / _em(src) if src else (b[3] - b[1]) / 0.78
        # на рисунке место вправо — до 3 ширин надписи: короткая подпись таблицы чертежа («Vonóerő:» → «Тяговое
        # усилие:») при 1,5 ужималась вдвое при свободном поле справа
        ext = 3.0 if pic else 1.5
        nxt = min([o[0] for o in boxes if o is not b and o[0] > b[2] and min(o[3], b[3]) - max(o[1], b[1]) > 0] + [W])
        right = max(_free_right(img, b, bg, min(nxt - 2, b[2] + (b[2] - b[0]) * ext + 20), _speck(b) if pic else 1), b[2])
        prv = max([o[2] for o in boxes if o is not b and o[2] < b[0] and min(o[3], b[3]) - max(o[1], b[1]) > 0] + [0])
        left = min(_free_left(img, b, bg, max(prv + 2, b[0] - (b[2] - b[0]) * ext - 20), _speck(b) if pic else 1), b[0])
        # надпись посередине свободного места (заголовок столбца, кнопка) — перевод по центру, место в обе стороны;
        # на рисунке — только зажатая с обеих сторон (ячейка, кнопка): подписи таблицы чертежа с пустым полем вокруг
        # иначе разъезжались от общего левого края
        fl, fr = b[0] - left, right - b[2]
        center = fl > 3 and fr > 3 and abs(fl - fr) < 0.35 * max(fl, fr)
        if pic:
            center = center and right + 2 < min(nxt - 2, b[2] + (b[2] - b[0]) * ext + 20) and \
                left - 2 > max(prv + 2, b[0] - (b[2] - b[0]) * ext - 20)
        base = b[3] - (0.21 * size if re.search(DESC, src) else 0.0)
        plan.append({"b": b, "bg": bg, "fg": fg, "bold": _bold(mask, b[3] - b[1]), "size": size,
                     "x0": left if center else b[0], "x1": right, "center": center, "base": base, "ru": ru,
                     "stroke": _stroke(mask)})
    if pic:
        # жирное на рисунке — штрих заметно толще, чем у остальных надписей картинки: у мелкого шрифта схемы штрих
        # 1–2 px, и мерка «0,16 высоты рамки» делала жирной обычную подпись («Окно для вывода», договор, 22.10)
        med = float(np.median([p["stroke"] for p in plan])) if plan else 0.0
        for p in plan:
            p["bold"] = len(plan) >= 3 and p["stroke"] >= 1.4 * med
    # один кегль у надписей одного размера: оценка по рамке гуляет на 10–20 % («Результаты» крупнее соседей); на рисунке —
    # до 30 %: у подписи схемы выносная буква, слитая с подчёркиванием, уходит из рамки («32" монитор» — на треть мельче)
    sizes = [p["size"] for p in plan]
    tol = 0.3 if pic else 0.15
    for p in plan:
        near = sorted(s for s in sizes if abs(s - p["size"]) <= tol * (max(s, p["size"]) if pic else p["size"]))
        if len(near) >= 3:
            p["size"] = near[len(near) // 2]
    # перевод рисуется заранее: не влезший (кнопка «Ki/Be» → «Выключено полностью») не стирает оригинал —
    # иначе на месте надписи оставалась пустота (красная команда 15.10)
    for p in plan:
        p["draw"] = _draw(p["ru"], p["size"], p["bold"], p["fg"], p["bg"], int(p["x1"] - p["x0"] + 2), p["center"],
                          p.get("height"), p.get("pitch"))
    plan = [p for p in plan if p["draw"][0] is not None]
    for p in plan:                               # сначала стереть все надписи, потом писать — соседи не затираются
        # стираются только буквы (с каймой сглаживания), не вся рамка: рамка задевала белую линию под шапкой таблицы
        for b in p.get("lines") or [p["b"]]:
            pad = 3 if pic else 1
            y0, x0 = max(int(b[1]) - pad, 0), max(int(b[0]) - pad, 0)
            cut = img[y0:min(int(b[3]) + pad + 1, H), x0:min(int(b[2]) + pad + 1, W)]
            d = np.abs(cut.astype(int) - p["bg"].astype(int)).sum(axis=2)
            raw = d > max((0.08 if pic else 0.2) * d.max(), 24)
            # рисунок: кайма в 2 px — у букв скана (JPEG) светлый ореол, при 1 px от стёртого заголовка оставалась тень
            m = cv2.dilate(raw.astype(np.uint8), np.ones((5, 5) if pic else (3, 3), np.uint8)) > 0
            if pic:
                cc = comp[y0:y0 + cut.shape[0], x0:x0 + cut.shape[1]]
                own = [i for i in np.unique(cc[cc > 0]) if st[i, 0] >= b[0] - 3 and st[i, 1] >= b[1] - 3
                       and st[i, 0] + st[i, 2] <= b[2] + 3 and st[i, 1] + st[i, 3] <= b[3] + 3]
                m &= ~((cc > 0) & ~np.isin(cc, own))
                # строка, почти сплошь отличная от фона (в запасе над и под рамкой — больше чем наполовину), — край
                # плашки или строка таблицы, не буквы: стёртая, она давала серую «ступеньку» под переведённым заголовком
                rows, full = np.arange(cut.shape[0]) + y0, raw.mean(axis=1)
                m[(full >= 0.9) | ((rows < int(b[1])) | (rows > int(b[3]))) & (full >= 0.5)] = False
            cut[m] = p["bg"].astype("uint8")
    for p in plan:
        b = p["b"]
        arr, base, _ = p["draw"]
        x0, y0 = max(int(p["x0"]), 0), int(round(p["base"] - base))
        ys = slice(max(y0, 0), min(y0 + arr.shape[0], H)); xs = slice(x0, min(x0 + arr.shape[1], W))
        a = arr[ys.start - y0:ys.stop - y0, :xs.stop - x0]
        # переносятся только пиксели букв: полоса заплатки выше рамки и не должна затирать соседние строки и линии
        m = np.abs(a.astype(int) - p["bg"].astype(int)).sum(axis=2) > 12
        region = img[ys, xs]
        region[m] = a[m]
    return img


def apply(doc, segs, res):
    """Заменяет в doc картинки-скриншоты переведёнными. Возвращает (заменено, ошибки)."""
    tags = re.compile(r"</?f\d+>")
    by, page_of = {}, {}
    for s in segs:
        sh = s["loc"].get("shot")
        if not sh or _masked(doc, sh["xref"]):
            continue
        r = res.get(str(s["id"])) or {}
        ru = tags.sub("", r.get("ru") or "").strip()
        page_of.setdefault(sh["xref"], s["page"])
        if ru and not r.get("copied") and ru != s["text"]:
            by.setdefault(sh["xref"], []).append((sh, ru))
    done, errors = 0, []
    for x, items in by.items():
        try:
            doc[page_of[x]].replace_image(x, pixmap=render(doc, x, items))
            done += 1
        except Exception as e:                  # noqa: BLE001 — одна картинка не валит сборку: остаётся оригинал
            errors.append(f"картинка {x} (стр. {page_of[x] + 1}): {type(e).__name__}: {e}"[:200])
    return done, errors


# ------------------------------------------------------------------ картинки DOCX
DOC = "word/document.xml"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_VML_IMG = "{urn:schemas-microsoft-com:vml}imagedata"
_SVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"
RASTER = {".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG", ".bmp": "BMP", ".tif": "TIFF", ".tiff": "TIFF"}
MIN_PX = 4000              # меньше (в пикселях) — значок или маркер, не рисунок с надписями
PORTAL_MAX = 2400          # фото в 12 Мп портал не примет: для него — уменьшенная копия (надписи сверяются по тексту)


def docx_pictures(path):
    """Растровые картинки тела DOCX: ([(запись word/media/…, номер абзаца тела с первой ссылкой)], {запись: почему
    пропущена}). Векторные (EMF, WMF, SVG) — пропуск: их надписи не растр, перерисовка заменила бы чертёж картинкой;
    PNG-заместитель при SVG Word не показывает. Картинки только колонтитулов (логотип, бланк) не смотрятся."""
    import posixpath
    from lxml import etree
    from segments import W
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        rels = {}
        if "word/_rels/document.xml.rels" in names:
            for r in etree.fromstring(z.read("word/_rels/document.xml.rels")):
                t = r.get("Target") or ""
                if r.get("TargetMode") != "External":
                    rels[r.get("Id")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("word", t))
        root = etree.fromstring(z.read(DOC))
    pics, skipped, seen = [], {}, set()
    for pi, p in enumerate(root.iter(W + "p")):
        for el in p.iter(_BLIP, _VML_IMG):
            name = rels.get(el.get(_R + "embed") if el.tag == _BLIP else el.get(_R + "id"))
            if not name or name not in names or name in seen:
                continue
            seen.add(name)
            ext = Path(name).suffix.lower()
            if el.tag == _BLIP and next(el.iter(_SVG_BLIP), None) is not None:
                skipped[name] = "векторная (SVG; PNG — запасная, Word её не показывает)"
            elif ext not in RASTER:
                skipped[name] = f"векторная ({ext[1:].upper()})" if ext in (".emf", ".wmf", ".emz", ".wmz", ".svg") \
                    else f"формат {ext or '?'} не читается"
            else:
                pics.append((name, pi))
    return pics, skipped


def _open_raster(data):
    """Картинка записи word/media: (RGB uint8 — прозрачное наложено на белый лист, исходная картинка PIL)."""
    import io
    import numpy as np
    from PIL import Image                   # Pillow — зависимость pdfplumber
    im = Image.open(io.BytesIO(data))
    im.load()
    a = np.asarray(im.convert("RGBA")).astype(np.float32)
    al = a[:, :, 3:] / 255.0
    return (a[:, :, :3] * al + 255.0 * (1 - al)).round().astype(np.uint8), im


def _save_raster(img, im, fmt, changed):
    """Перерисованная картинка в формате исходной: тот же размер в пикселях, DPI, ICC и EXIF (поворот снимка), у JPEG —
    те же таблицы квантования (качество как у исходной); прозрачный фон остаётся прозрачным, кроме перерисованного."""
    import io
    import numpy as np
    from PIL import Image, JpegImagePlugin
    kw = {k: im.info[k] for k in ("dpi", "icc_profile", "exif") if im.info.get(k)}
    if fmt != "JPEG" and ("A" in im.getbands() or "transparency" in im.info):
        alpha = np.asarray(im.convert("RGBA"))[:, :, 3].copy()
        alpha[changed] = 255
        out = Image.fromarray(np.dstack([img, alpha]), "RGBA")
    else:
        out = Image.fromarray(img, "RGB")
        if im.mode in ("L", "1"):
            out = out.convert("L")
    if fmt == "JPEG" and getattr(im, "quantization", None):
        kw["qtables"] = im.quantization
        if JpegImagePlugin.get_sampling(im) >= 0:
            kw["subsampling"] = JpegImagePlugin.get_sampling(im)
    buf = io.BytesIO()
    out.save(buf, fmt, **kw)
    return buf.getvalue()


def _portal_img(img):
    import fitz
    import numpy as np
    pix = fitz.Pixmap(fitz.csRGB, img.shape[1], img.shape[0], np.ascontiguousarray(img).tobytes(), False)
    k = PORTAL_MAX / max(pix.width, pix.height)
    if k < 1:
        pix = fitz.Pixmap(pix, int(pix.width * k), int(pix.height * k))
    return portal_png(pix)


LINE_PX = 30               # линия чертежа: тёмный отрезок длиннее (px) и тоньше 9 px


def _line_mask(img):
    """Пиксели тонких длинных линий (подчёркивание выносок, рамки, размерные линии) — _lines вместе."""
    hor, ver = _lines(img)
    return hor | ver


def _lines(img):
    """(горизонтальные, вертикальные) тонкие линии: темнее окрестности (чёрная шляпа) и тоньше 9 px поперёк, длиннее
    LINE_PX вдоль. Не по абсолютной яркости: заливки (зелёные плашки «АРМ» на плане зала) стирались вместе с буквами,
    а полоски серой шапки таблицы между светлыми буквами и краем шапки шли за линии — белый текст шапки размывался
    (стенд UTS: «Azonosító», «Kategória»…). Светлые линии (белая шляпа) не берутся: на стенде UTS минус одна надпись."""
    import cv2
    import numpy as np
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    k = lambda w, h: cv2.getStructuringElement(cv2.MORPH_RECT, (w, h))
    thin = lambda w, h: (cv2.morphologyEx(gray, cv2.MORPH_BLACKHAT, k(w, h)) > 40).astype(np.uint8)
    return (cv2.morphologyEx(thin(1, 9), cv2.MORPH_OPEN, k(LINE_PX, 1)) > 0,
            cv2.morphologyEx(thin(9, 1), cv2.MORPH_OPEN, k(1, LINE_PX)) > 0)


def _unline(img):
    """Картинка без тонких длинных линий: они касаются букв, и Tesseract резал слова («Съемный кож|ух», «по направяю|щей
    vesa» — договор, 22.10). Пиксели линии — медианой окрестности."""
    import cv2
    m = _line_mask(img)
    out = img.copy()
    out[m] = cv2.medianBlur(img, 9)[m]
    return out


def line_words(img, boxes, work_png, lang="hu"):
    """Слова строк детектора, строка целиком (psm 6, все строки столбиком на одном листе — один вызов Tesseract): на
    схемах выноски в несколько строк стоят тесно, и разреженный режим (det_words) терял короткие слова («с», «в»,
    «АРМ») и резал слова у линий. Рамки детектора на тесных строках перекрываются — в вырезке строки соседние строки
    забелены по середине между ними («регулирования» не читалось вовсе). Прямо и негативом (светлое на плашке), у строки —
    вариант с большей средней уверенностью. [{t, box (пиксели img), h, c, line (номер рамки)}]."""
    import cv2
    import numpy as np
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    H, W = gray.shape
    crops, y, band = [], 20, {}
    for i, b in enumerate(boxes):
        x0, y0, x1, y1 = max(int(b[0]) - 2, 0), max(int(b[1]) - 2, 0), min(int(b[2]) + 3, W), min(int(b[3]) + 3, H)
        if x1 <= x0 or y1 <= y0:
            continue
        g = gray[y0:y1, x0:x1].copy()
        cy = (b[1] + b[3]) / 2
        band[i] = [y0, y1]                  # полоса строки: рамки слов — не выше и не ниже середины до соседних строк
        for j, o in enumerate(boxes):
            oy = (o[1] + o[3]) / 2
            if j == i or min(o[2], x1) <= max(o[0], x0) or min(o[3], y1) <= max(o[1], y0) or abs(oy - cy) < 0.3 * (b[3] - b[1]):
                continue                    # не задевает или та же строка (соседнее слово)
            xa, xb, mid = max(int(o[0]) - x0, 0), min(int(o[2]) + 1 - x0, x1 - x0), int((cy + oy) / 2) - y0
            if oy < cy:
                g[:max(mid, 0), xa:xb] = 255
                band[i][0] = max(band[i][0], y0 + mid)
            else:
                g[max(mid + 1, 0):, xa:xb] = 255
                band[i][1] = min(band[i][1], y0 + mid + 1)
        s = min(5.0, max(1.0, 40.0 / (y1 - y0)))
        c = cv2.resize(g, None, fx=s, fy=s, interpolation=cv2.INTER_CUBIC)
        # и бинаризованная: серую вырезку (тёмное «АРМ» на зелёной плашке) Tesseract на листе принимал за картинку, а
        # бинаризация портит тонкий шрифт на белом («vesa» → «уеза») — у строки выигрывает вариант увереннее
        b2 = cv2.threshold(c, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        for v, cc in enumerate((c, b2, 255 - b2)):
            crops.append((cc, x0, y0, s, y, i, v))
            y += cc.shape[0] + 30
    if not crops:
        return []
    sheet = np.full((y + 20, max(c[0].shape[1] for c in crops) + 40), 255, np.uint8)
    for c, _, _, _, yy, _, _ in crops:
        sheet[yy - 8:yy + c.shape[0] + 8, 12:28 + c.shape[1]] = int(np.median(c))    # кромка — цветом фона вырезки
        sheet[yy:yy + c.shape[0], 20:20 + c.shape[1]] = c
    cv2.imwrite(str(work_png), sheet)
    got = {}
    for t, x, yw, w, h, conf in _tsv(work_png, lang, psm=6):
        for c, x0, y0, s, yy, i, v in crops:
            if yy <= yw + h / 2 <= yy + c.shape[0]:
                got.setdefault((i, v), []).append({"t": t, "box": [x0 + (x - 20) / s, y0 + (yw - yy) / s,
                                                                   x0 + (x - 20 + w) / s, y0 + (yw - yy + h) / s],
                                                   "h": h / s, "c": conf, "line": i})
                break
    words = []
    for i in {k[0] for k in got}:
        v, ws = max(((v, got.get((i, v), [])) for v in range(3)),
                    key=lambda vw: sum(w["c"] for w in vw[1]) / len(vw[1]) if vw[1] else -1)
        ws.sort(key=lambda w: w["box"][0])
        for w, nx in zip(ws, ws[1:]):
            # рамка, накрывшая следующее слово целиком, — до него: Tesseract дал «"Azonosító» рамку во всю строку шапки,
            # и разрыв до соседнего столбца пропадал (просто перекрытые рамки — «TU» и «azonosító» — не трогаются)
            if nx["box"][0] > w["box"][0] and w["box"][2] >= nx["box"][2]:
                w["box"][2] = nx["box"][0] - 1
        for w in ws:
            bx = w["box"]
            bx = [bx[0], max(bx[1], band[i][0]), bx[2], min(bx[3], band[i][1])] if bx[3] > band[i][0] and \
                bx[1] < band[i][1] else bx
            w["box"] = _tight(gray, bx, light=v == 2)
            w["h"] = w["box"][3] - w["box"][1]
        words += ws
    return words


def _tight(gray, b, light=False):
    """Рамка слова, ужатая по краске: на клетчатом фоне плана рамка Tesseract выходила вдвое выше строки — строки
    подписи не склеивались, кегль перевода раздувался (договор, 22.10). Порог — посередине между самым тёмным пикселем
    (light — самым светлым) и медианой рамки: Оцу на краю зелёной плашки делил белое поле и плашку, а не буквы и фон."""
    import numpy as np
    H, W = gray.shape
    x0, y0, x1, y1 = max(int(b[0]), 0), max(int(b[1]), 0), min(int(np.ceil(b[2])), W), min(int(np.ceil(b[3])), H)
    cut = gray[y0:y1, x0:x1].astype(int)
    if cut.size < 4:
        return b
    med, ext = float(np.median(cut)), float(cut.max() if light else cut.min())
    if abs(ext - med) < 30:
        return b
    ink = cut > (ext + med) / 2 if light else cut < (ext + med) / 2
    ink[:, ink.mean(axis=0) >= 0.9] = False         # контур плашки в рамке — линия во всю высоту, не буква
    ink[ink.mean(axis=1) >= 0.9, :] = False
    rows, cols = np.nonzero(ink.any(axis=1))[0], np.nonzero(ink.any(axis=0))[0]
    if not len(rows):
        return b
    return [float(max(b[0], x0 + cols[0])), float(max(b[1], y0 + rows[0])),
            float(min(b[2], x0 + cols[-1] + 1)), float(min(b[3], y0 + rows[-1] + 1))]


def _wall(gray, a, b):
    """Между рамками a и b (слева направо) — столбец не цвета фона во всю высоту строки с запасом: разделитель столбцов
    таблицы (на серой шапке — светлый, короче линии _lines), буквы такого не дают."""
    import numpy as np
    x0, x1 = int(a[2]) + 1, int(b[0])
    if x1 <= x0:
        return False
    cut = gray[max(int(min(a[1], b[1])) - 2, 0):int(max(a[3], b[3])) + 3, x0:x1].astype(int)
    return bool(cut.size) and bool(((np.abs(cut - np.median(cut)) > 30).mean(axis=0) >= 0.9).any())


def _line_labels(words, gray=None):
    """Слова строк детектора → надписи: строка — по рамке детектора (не по высоте слов: у «„для» с кавычкой рамка выше
    соседей, и разбор по высоте уводил слово в отдельную строку), внутри строки — разрыв больше высоты строки или
    разделитель между словами (_wall по серой картинке gray) делит надписи: заголовки соседних столбцов таблицы — разные
    надписи («Azonosító» и «OS felhasználó azonosító» сливались, и портал отбрасывал первое)."""
    labels = []
    for i in sorted({w["line"] for w in words}):
        ws = sorted((w for w in words if w["line"] == i), key=lambda w: w["box"][0])
        # одиночный знак заметно выше строки — разделитель столбцов, прочитанный «I»: растягивал рамку надписи до строки
        # таблицы, и заплатка цвета шапки залезала в белую строку
        med = sorted(w["box"][3] - w["box"][1] for w in ws)[len(ws) // 2]
        ws = [w for w in ws if len(w["t"]) > 1 or w["box"][3] - w["box"][1] <= 1.3 * med] or ws
        h = max(w["box"][3] - w["box"][1] for w in ws)
        cur = None
        for w in ws:
            if cur and w["box"][0] - cur["box"][2] <= h and not (gray is not None and _wall(gray, cur["box"], w["box"])):
                b = cur["box"]
                cur.update(t=cur["t"] + " " + w["t"], c=min(cur["c"], w["c"]),
                           box=[b[0], min(b[1], w["box"][1]), max(b[2], w["box"][2]), max(b[3], w["box"][3])])
            else:
                cur = {"t": w["t"], "box": list(w["box"]), "c": w["c"]}
                labels.append(cur)
    for lb in labels:
        lb["h"] = lb["box"][3] - lb["box"][1]
        lb["box"] = [round(v, 1) for v in lb["box"]]
    return labels


def pic_labels(img, work_png, lang="hu"):
    """Надписи рисунка DOCX (схема, чертёж, фото, скриншот): строки детектора читаются целиком (line_words) на картинке
    без тонких линий (_unline); слова Tesseract по всей картинке — только вне прочитанных строк (детектор пропускал
    надписи на цветных плашках) и отдельными надписями: склеенное с подписью мусорное слово у стрелки выноски
    раздувало её рамку и кегль. [{t, box, h, c}], как у tess_labels."""
    import cv2
    clean = _unline(img)
    words = []
    if _det["error"] is None:
        try:
            boxes = det_boxes(clean)
            words = line_words(clean, boxes, work_png, lang) if boxes else []
        except cv2.error as e:              # модель не читается этой сборкой OpenCV — дальше только Tesseract
            _det["error"] = f"{type(e).__name__}: {e}"[:300]
            print(f"shots: детектор надписей отключён — {_det['error']}", file=sys.stderr, flush=True)
    k = max(1.0, min(4.0, UPSCALE_TO / max(img.shape[1], 1)))
    gray = cv2.cvtColor(clean, cv2.COLOR_RGB2GRAY)
    big = cv2.resize(gray, None, fx=k, fy=k, interpolation=cv2.INTER_CUBIC)
    extra = []
    for neg, img_k in enumerate((big, 255 - big)):
        cv2.imwrite(str(work_png), img_k)
        first = words + extra
        for t, x, y, w, h, c in _tsv(work_png, lang):
            box = _tight(gray, [x / k, y / k, (x + w) / k, (y + h) / k], light=bool(neg))
            if not _hit(box, first):
                extra.append({"t": t, "box": box, "h": box[3] - box[1], "c": c})
    return _line_labels(words, cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)) + _labels(extra)


def _blocks(labels):
    """Строки одной фразы (подпись схемы в несколько строк) — одна надпись: перевод по строкам ломал фразу («Ось 32"
    монитора, | возможно смещение | по направляющей vesa | крепления ±20 мм.», договор, 22.10). Строка продолжает
    надпись выше, если стоит прямо под ней (зазор меньше высоты строки, тот же кегль, левый край или середина рядом) и
    начинается строчной буквой — или цифрой, когда строка выше не кончается точкой; перенос по дефису — тоже. Пункты меню
    и кнопки столбиком начинаются прописной и остаются отдельными. Только картинки DOCX: скриншоты PDF — по строкам."""
    order = sorted(range(len(labels)), key=lambda i: (labels[i]["box"][1], labels[i]["box"][0]))
    blocks = []                             # [номер первой строки, [строки]]
    for i in order:
        lb, t = labels[i], labels[i]["t"].strip()
        bx, h1 = lb["box"], lb["box"][3] - lb["box"][1]
        first = re.match(r"\W*(\w)", t)
        best = None
        for bl in blocks:
            p = bl[1][-1]
            pb, pt, h0 = p["box"], p["t"].strip(), p["box"][3] - p["box"][1]
            h, gap = min(h0, h1), bx[1] - pb[3]
            if not -0.3 * h <= gap <= 0.9 * h or max(h0, h1) > 1.4 * h:
                continue
            if abs(bx[0] - pb[0]) > 1.5 * h and abs(bx[0] + bx[2] - pb[0] - pb[2]) / 2 > 1.5 * h:
                continue
            cont = pt.endswith("-") or first and (first[1].islower() or
                                                  first[1].isdigit() and not re.search(r"[.!?:;]$", pt))
            if cont and (best is None or gap < best[0]):
                best = (gap, bl)
        if best:
            best[1][1].append(lb)
        else:
            blocks.append([i, [lb]])
    out = []
    for _, lines in sorted(blocks, key=lambda bl: bl[0]):
        if len(lines) == 1:
            out.append(lines[0]); continue
        text = lines[0]["t"].strip()
        for ln in lines[1:]:
            t = ln["t"].strip()
            # перенос: «закреп-» + «ления» — без дефиса; «500-» + «600» — дефис остаётся
            text = text[:-1] + t if text.endswith("-") and t[:1].islower() else text + ("" if text.endswith("-") else " ") + t
        boxes = [ln["box"] for ln in lines]
        out.append({"t": text, "box": [min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes),
                                       max(b[3] for b in boxes)],
                    "h": lines[0]["h"], "c": min(ln["c"] for ln in lines), "lines": boxes,
                    "ts": [ln["t"].strip() for ln in lines]})
    return out


def docx_shot_segments(path, work, lang="hu", progress=None):
    """Надписи картинок тела DOCX: (сегменты без номеров — номера даёт add_docx_shots, {запись: почему пропущена}).
    Кэш чтения — work/<имя картинки>[.<lang>].json; сбой портала не кэшируется."""
    from concurrent.futures import ThreadPoolExecutor
    from PIL import Image
    pics, skipped = docx_pictures(path)
    work, sfx = Path(work), "" if lang == "hu" else f".{lang}"
    res, todo = {}, []
    with zipfile.ZipFile(path) as z:
        for name, _ in pics:
            f = work / f"{Path(name).name}{sfx}.json"
            if f.exists():
                res[name] = json.loads(f.read_text(encoding="utf-8")); continue
            try:
                img = _open_raster(z.read(name))[0]
            except (OSError, ValueError, Image.DecompressionBombError) as e:      # битая или неизвестная картинка
                skipped[name] = f"не открылась: {type(e).__name__}: {e}"[:200]; continue
            if img.shape[0] * img.shape[1] < MIN_PX:
                skipped[name] = "значок"; continue
            todo.append((name, img, f))
    if todo:
        work.mkdir(parents=True, exist_ok=True)

    def one(job):
        name, img, f = job
        size = [img.shape[1], img.shape[0]]
        tmp = work / f"{Path(name).name}.tess.png"
        try:
            # одиночный знак (цифра размера, «Е» штриховки) не переводится, а длинный список мусора сбивал портал с
            # нумерации ответа
            labels = [lb for lb in pic_labels(img, tmp, lang) if len(re.findall(r"[^\W_]", lb["t"])) >= 2]
            tmp.unlink(missing_ok=True)
            good = check_labels(_portal_img(img), labels, f"pic:{name}", lang, PICTURE, merge=True)
        except Exception as e:              # noqa: BLE001 — одна картинка не валит извлечение задачи
            return {"size": size, "labels": [], "error": f"{type(e).__name__}: {e}"[:200]}
        if good is None:                    # портал не ответил — в кэш не пишем, следующий прогон спросит снова
            return {"size": size, "labels": [], "error": "портал не ответил"}
        out = {"size": size, "labels": _blocks(good), "ocr": len(labels)}
        f.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        return out

    with ThreadPoolExecutor(4) as ex:
        for i, (job, r) in enumerate(zip(todo, ex.map(one, todo))):
            res[job[0]] = r
            if progress:
                progress((i + 1) / len(todo))
    segs = []
    for name, pi in pics:
        r = res.get(name)
        if not r:
            continue
        if r.get("error"):
            skipped[name] = r["error"]
        for lb in r["labels"]:
            sh = {"media": name, "p": pi, "box": lb["box"], "h": lb["h"], "t": lb["t"], "size": r["size"]}
            if lb.get("lines"):
                sh.update(lines=lb["lines"], ts=lb["ts"])
            segs.append({"id": None, "text": lb["t"], "src": lb["t"], "page": None, "loc": {"shot": sh}})
    return segs, skipped


def add_docx_shots(segs, path, work, lang="hu"):
    """Сегменты DOCX (segments.docx_segments) + надписи его картинок — сразу после абзаца тела, где стоит картинка
    (перевод разделами видит текст вокруг рисунка); номера сегментов — по порядку. Возвращает (сегменты, сведения):
    без картинок — те же сегменты и пустые сведения."""
    extra, skipped = docx_shot_segments(path, work, lang)
    if not extra and not skipped:
        return segs, {}
    pend, out = list(extra), []
    for s in segs:
        while pend and (s["loc"].get("part") != DOC or pend[0]["loc"]["shot"]["p"] < s["loc"]["p"]):
            out.append(pend.pop(0))
        out.append(s)
    out += pend
    for i, s in enumerate(out):
        s["id"] = i
    info = {"shot_labels": len(extra)}
    if skipped:
        info["shot_skipped"] = skipped
    return out, info


def docx_media(src, segs, res):
    """Перерисованные картинки DOCX: ({запись word/media/…: байты}, ошибки). Картинка остаётся той же записью того же
    формата и размера в пикселях — разметка документа (место, размер на листе, обрезка) не меняется."""
    tags = re.compile(r"</?f\d+>")
    by = {}
    for s in segs:
        sh = s["loc"].get("shot")
        if not sh or "media" not in sh:
            continue
        r = res.get(str(s["id"])) or {}
        ru = tags.sub("", r.get("ru") or "").strip()
        if ru and not r.get("copied") and ru != s["text"]:
            by.setdefault(sh["media"], []).append((sh, ru))
    out, errors = {}, []
    if not by:
        return out, errors
    with zipfile.ZipFile(src) as z:
        for name, items in by.items():
            try:
                img, im = _open_raster(z.read(name))
                new = render_img(img.copy(), items, pic=True)
                out[name] = _save_raster(new, im, RASTER[Path(name).suffix.lower()], (new != img).any(axis=2))
            except Exception as e:              # noqa: BLE001 — одна картинка не валит сборку: остаётся оригинал
                errors.append(f"картинка {name}: {type(e).__name__}: {e}"[:200])
    return out, errors
