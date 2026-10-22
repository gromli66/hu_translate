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
Прежняя «белая плашка поверх надписи» (режим скана на всю страницу) давала грязь — UTS, 15.10."""
import json, re, difflib, sys, threading
from pathlib import Path

DECOR_SHARE, MIN_AREA, FULL = 0.3, 3000, 0.5   # FULL: от половины листа — скан или фон, не скриншот
UPSCALE_TO = 3000          # Tesseract: ширина увеличенной картинки, px (мелкий шрифт интерфейса ~8–10 px)
# детектор строк текста: PP-OCRv4 det (PaddleOCR, Apache-2.0; файл из rapidocr_onnxruntime 1.4.4), 4,7 МБ.
# Увеличение 1,5× (2× — та же полнота на UTS, вдвое дольше), площадь — до 2,5 Мп: память сети ~0,7 ГБ
DET_MODEL = Path(__file__).parent / "models" / "ch_PP-OCRv4_det_infer.onnx"
DET_SCALE, DET_MAX_PX = 1.5, 2.5e6
_det = {"net": None, "error": None, "lock": threading.Lock()}   # сеть OpenCV DNN не потокобезопасна

# Что переводить, решает язык надписи: в перевод — только lang=hu. Прежний вопрос «переводить ли» (промпт «скриншот на
# венгерском», «английские слова не переводить») портал решал через раз в обе стороны: английские окна VirtualBox,
# Softing, вход SUSE (UTS стр. 19, 88–89, 95) то переводил, то нет, а роли «Operátor»/«Adminisztrátor» то оставлял;
# язык надписи по картинке он называет устойчиво (замер 15.10 — work/uts_2026-10-15/shots_bench/README.md,
# «Английские окна»)
CHECK_PROMPT = (
    "Это скриншот интерфейса программы. Ниже — надписи с него, прочитанные OCR с венгерским словарём, по одной на строке "
    "в виде «номер: текст».\n{labels}\n\n"
    "Для каждой надписи верни объект JSON {{\"n\": номер, \"t\": \"текст надписи точно как на картинке\", "
    "\"lang\": \"hu\"|\"en\"|\"de\"|\"-\"}}.\n"
    "- t: исправь ошибки OCR по картинке (венгерские буквы á é í ó ö ő ú ü ű только там, где они есть на картинке); "
    "если надписи на картинке нет или это мусор OCR — t пустая строка.\n"
    "- lang: язык надписи на картинке: hu — венгерский (и вместе с кодом или числом: «UTS üzenet», «1 perc»), "
    "en — английский, de — немецкий, \"-\" — без языка: коды и идентификаторы (X0YZ11K500_ZD11, SKRVR_X_TEST), имена "
    "файлов и путей, числа, даты, время, имена людей, сокращения из заглавных латинских букв. OCR читал венгерским "
    "словарём и мог исказить английские слова — решай по картинке.\n"
    "Не добавляй надписей, которых нет в списке. Выведи только массив JSON.")


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
            # логотип в зоне шапки (небольшая картинка у верха листа, уже трети его ширины) — не скриншот: название
            # организации на нём «переводилось» вперемешку алфавитов («Юрszágos Atomenergia Управление», E42, 22.10)
            if r.y1 <= 0.22 * page.rect.height and r.width < 0.3 * page.rect.width and r.get_area() < 0.03 * area:
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


def _tsv(png):
    """Слова Tesseract (hun, psm 11) с листа png: [(текст, x, y, w, h, уверенность)] — с буквой или цифрой, от 30."""
    import subprocess
    r = subprocess.run(["tesseract", str(png), "stdout", "-l", "hun", "--psm", "11", "tsv"],
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


def det_words(img, boxes, work_png):
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
    for t, x, yw, w, h, conf in _tsv(work_png):
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


def tess_labels(img, work_png):
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
        for t, x, y, w, h, c in _tsv(work_png):
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
            for w in det_words(img, boxes, work_png) if boxes else []:
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


def check_labels(png, labels, tag):
    """Портал по картинке (PNG, portal_png) исправляет чтение и решает, что переводить. Сбой портала — надписи
    остаются непереведёнными (на скриншоте лучше венгерская надпись, чем неверно прочитанная)."""
    from llm import chat, img_b64

    def ask(part, tg):
        text = "\n".join(f"{i + 1}: {lb['t']}" for i, lb in enumerate(part))
        r = chat([{"role": "user", "content": CHECK_PROMPT.format(labels=text)}], images=[img_b64(png)],
                 max_tokens=6000, tag=tg)
        if not r.get("text"):
            return None, None
        return _answer_objects(r["text"]), r.get("finish")

    if not labels:
        return []
    got, finish = ask(labels, tag)
    if got is None:
        return None                          # сбой портала: не кэшировать «нечего переводить»
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
        if not o or o.get("lang") != "hu" or not str(o.get("t", "")).strip():
            continue
        # уверенное чтение Tesseract портал не переписывает: он «исправлял» надпись по соседней кнопке
        # (FELCSATLAKOZÁS → LECSATLAKOZÁS — обратный смысл, UTS стр. 26, 15.10); неуверенное — только похожим
        t = str(o["t"]).strip()
        if lb["c"] >= 85 or difflib.SequenceMatcher(None, t.lower(), lb["t"].lower()).ratio() < 0.6:
            t = lb["t"]
        out.append(dict(lb, t=t, ocr=lb["t"]))
    return out


def shot_segments(doc, work, start_id, pages=None, progress=None):
    """Сегменты надписей всех скриншотов документа (кэш чтения — work/shot_<xref>.json)."""
    from concurrent.futures import ThreadPoolExecutor
    work = Path(work); work.mkdir(parents=True, exist_ok=True)
    shots = {x: pl for x, pl in find_shots(doc).items() if pages is None or pl[0][0] in pages}

    xs = sorted(shots, key=lambda x: shots[x][0][0])
    # всё, что трогает MuPDF (картинка, PNG для портала), — в этом потоке: документ PyMuPDF не потокобезопасен;
    # в пуле — только Tesseract (отдельный процесс) и запрос к порталу
    res, todo = {}, []
    for x in xs:
        f = work / f"shot_{x}.json"
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
            labels = tess_labels(img, work / f"shot_{x}.png")
            (work / f"shot_{x}.png").unlink(missing_ok=True)
            good = check_labels(png, labels, f"shot{x}")
        except Exception as e:              # noqa: BLE001 — одна картинка не валит извлечение задачи
            return {"size": size, "labels": [], "error": f"{type(e).__name__}: {e}"[:200]}
        if good is None:                    # портал не ответил — в кэш не пишем, следующий прогон спросит снова
            return {"size": size, "labels": [], "error": "портал не ответил"}
        out = {"size": size, "labels": good, "ocr": len(labels)}
        (work / f"shot_{x}.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
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
def _bg_color(img, b, pad=1):
    """Цвет фона надписи — пик сглаженной гистограммы яркости рамки с кромкой (буквы — меньшинство). Медиана по кромке
    и мода по корзинам цвета ошибались: серый фон шапки с шумом сглаживания делился на две корзины, и побеждала
    белая линия-разделитель — шапка (белое на сером) перерисовывалась тёмным по белому (UTS, 15.10)."""
    import numpy as np
    H, W = img.shape[:2]
    x0, y0, x1, y1 = max(int(b[0]) - pad, 0), max(int(b[1]) - pad, 0), min(int(b[2]) + pad + 1, W), min(int(b[3]) + pad + 1, H)
    px = img[y0:y1, x0:x1].reshape(-1, 3).astype(float)
    lum = px @ np.array([0.299, 0.587, 0.114])
    hist = np.bincount(np.clip(lum, 0, 255).astype(int), minlength=256).astype(float)
    peak = int(np.convolve(hist, np.ones(17), mode="same").argmax())
    return px[np.abs(lum - peak) <= 12].mean(axis=0)


def _ink(img, b, bg):
    """Пиксели букв в рамке: (маска, цвет букв)."""
    import numpy as np
    cut = img[int(b[1]):int(b[3]) + 1, int(b[0]):int(b[2]) + 1].astype(float)
    d = np.abs(cut - bg).sum(axis=2)
    if not d.size or d.max() <= 30:
        return np.zeros(d.shape, bool), np.array([0.0, 0.0, 0.0])
    m = d >= 0.5 * d.max()
    return m, np.median(cut[d >= 0.6 * d.max()], axis=0)


def _bold(mask, h):
    """Жирное — толщина штриха (медиана горизонтальных отрезков букв) от 0,16 высоты строки."""
    import numpy as np
    if not mask.any() or h <= 0:
        return False
    d = np.diff(np.pad(mask.astype(np.int8), ((0, 0), (1, 1))), axis=1)
    runs = np.nonzero(d == -1)[1] - np.nonzero(d == 1)[1]
    return float(np.median(runs)) >= 0.16 * h


def _em(text):
    """Высота рамки Tesseract в долях кегля: от верха (прописные с акцентом выше) до низа (выносные)."""
    top = 0.9 if re.search(r"[ÁÉÍÓÖŐÚÜŰ]", text) else 0.73 if re.search(r"[A-ZÁ-Űbdfhklt0-9áéíóöőúüű]", text) else 0.53
    return top + (0.21 if re.search(r"[gjpqy,;]", text) else 0.0)


def _free_right(img, b, bg, limit):
    """Сколько места вправо от рамки, пока столбец пикселей в полосе надписи — фон."""
    import numpy as np
    H, W = img.shape[:2]
    y0, y1 = max(int(b[1]), 0), min(int(b[3]) + 1, H)
    x = int(b[2]) + 1
    while x < min(W, int(limit)):
        col = img[y0:y1, x].astype(float)
        if np.abs(col - bg).sum(axis=1).max() > 60:
            break
        x += 1
    return x - 2


def _free_left(img, b, bg, limit):
    """Сколько места влево от рамки, пока столбец пикселей в полосе надписи — фон."""
    import numpy as np
    H, W = img.shape[:2]
    y0, y1 = max(int(b[1]), 0), min(int(b[3]) + 1, H)
    x = int(b[0]) - 1
    while x > max(0, int(limit)):
        col = img[y0:y1, x].astype(float)
        if np.abs(col - bg).sum(axis=1).max() > 60:
            break
        x -= 1
    return x + 2


def _draw(text, size, bold, fg, bg, width, center=False):
    """Строка перевода на фоне bg: (массив пикселей, строка базовой линии, итоговый кегль)."""
    import fitz
    col = "rgb(%d,%d,%d)" % tuple(int(v) for v in fg)
    ph = int(size * 1.5) + 4
    css = (f"* {{font-family: sans-serif; font-size: {size:.2f}px; line-height: 1.2; color: {col}; white-space: nowrap;"
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
    pix = _rgb(doc, xref)
    img = _np(pix).copy()
    H, W = img.shape[:2]
    boxes = [sh["box"] for sh, _ in items]
    plan = []
    for sh, ru in items:
        b = sh["box"]
        bg = _bg_color(img, b)
        mask, fg = _ink(img, b, bg)
        src = sh.get("t", "")
        size = (b[3] - b[1]) / _em(src) if src else (b[3] - b[1]) / 0.78
        nxt = min([o[0] for o in boxes if o is not b and o[0] > b[2] and min(o[3], b[3]) - max(o[1], b[1]) > 0] + [W])
        right = max(_free_right(img, b, bg, min(nxt - 2, b[2] + (b[2] - b[0]) * 1.5 + 20)), b[2])
        prv = max([o[2] for o in boxes if o is not b and o[2] < b[0] and min(o[3], b[3]) - max(o[1], b[1]) > 0] + [0])
        left = min(_free_left(img, b, bg, max(prv + 2, b[0] - (b[2] - b[0]) * 1.5 - 20)), b[0])
        # надпись посередине свободного места (заголовок столбца, кнопка) — перевод по центру, место в обе стороны
        fl, fr = b[0] - left, right - b[2]
        center = fl > 3 and fr > 3 and abs(fl - fr) < 0.35 * max(fl, fr)
        base = b[3] - (0.21 * size if re.search(r"[gjpqy,;]", src) else 0.0)
        plan.append({"b": b, "bg": bg, "fg": fg, "bold": _bold(mask, b[3] - b[1]), "size": size,
                     "x0": left if center else b[0], "x1": right, "center": center, "base": base, "ru": ru})
    # один кегль у надписей одного размера: оценка по рамке гуляет на 10–20 % («Результаты» крупнее соседей)
    sizes = [p["size"] for p in plan]
    for p in plan:
        near = sorted(s for s in sizes if abs(s - p["size"]) <= 0.15 * p["size"])
        if len(near) >= 3:
            p["size"] = near[len(near) // 2]
    import cv2
    # перевод рисуется заранее: не влезший (кнопка «Ki/Be» → «Выключено полностью») не стирает оригинал —
    # иначе на месте надписи оставалась пустота (красная команда 15.10)
    for p in plan:
        p["draw"] = _draw(p["ru"], p["size"], p["bold"], p["fg"], p["bg"], int(p["x1"] - p["x0"] + 2), p["center"])
    plan = [p for p in plan if p["draw"][0] is not None]
    for p in plan:                               # сначала стереть все надписи, потом писать — соседи не затираются
        # стираются только буквы (с каймой сглаживания), не вся рамка: рамка задевала белую линию под шапкой таблицы
        b = p["b"]
        y0, x0 = max(int(b[1]) - 1, 0), max(int(b[0]) - 1, 0)
        cut = img[y0:min(int(b[3]) + 2, H), x0:min(int(b[2]) + 2, W)]
        d = np.abs(cut.astype(int) - p["bg"].astype(int)).sum(axis=2)
        m = cv2.dilate((d > max(0.2 * d.max(), 24)).astype(np.uint8), np.ones((3, 3), np.uint8)) > 0
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
    return fitz.Pixmap(fitz.csRGB, W, H, np.ascontiguousarray(img).tobytes(), False)


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
