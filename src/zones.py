# -*- coding: utf-8 -*-
"""Колонтитул листа — один решатель для текстового слоя (T) и распознанных листов (O): E9b-Z, E9_DESIGN.md §6.

Решение «эта строка — колонтитул или бегущее название раздела, не тело» жило россыпью в docx_layout.load (повтор по
тексту и по высоте, метка портала, меняющееся название в рамке, повтор с прошлого листа, название над заголовком) и
частью — в пути А (docx_struct.layout, только текстовый слой). Здесь — одна функция decide: вход — сегменты тела по
листам, линии листов (таблицы листа — table_heads, рамка шапки), метки портала; выход — тело и причина по каждому снятому
сегменту. Шаг 0 — перенос без смены выхода; правила Z1, Z7, Z4, Z2, Z5, Z8, Z11 — по одному поверх (E9B_Z_REPORT.md)."""
import re
from collections import Counter

import geom as GM
import patterns as PT

# поле шапки листа: подпись из букв (до 40 знаков, без цифр), двоеточие, значение — «Azonosító: 3PR10/INFO»,
# «Oldalszám:1», «Megnevezés: …»; «Érintett fejezetek:» без значения и «1.16/3. A fővízkör…» — не поле
FIELD = re.compile(r"\s*[^\W\d_][^:\d]{0,40}:\s*\S")


def zone(s, H):
    y0, y1 = s["loc"]["bbox"][1], s["loc"]["bbox"][3]
    return "head" if y1 < GM.HEAD_EDGE * H else "foot" if y0 > GM.FOOT_EDGE * H else None


def sig(text):
    return re.sub(r"\d+", "#", text.strip())


def running_keys(segs, H, pages):
    """Колонтитулы: (зона, текст без цифр) или (зона, высота строки), что повторяется хотя бы на половине листов (и не
    меньше чем на двух). По высоте — потому что OCR режет шапку по-разному: «Megnevezés … UTASÍTÁSA 72» на одном листе
    и «Verziószám: 0 Oldalszám:» отдельно — на другом (3PR42, 16.10). Считаются листы, а не куски: строка таблицы из
    трёх кусков у верха одностраничного документа — не «три повтора» (30RDGT modj: шапка таблицы и первая строка
    уходили в шапку DOCX, в таблице на их месте — пустые строки, 20.10)."""
    cnt, pos = {}, {}
    for s in segs:
        z = zone(s, H)
        if z:
            cnt.setdefault((z, sig(s["text"])), set()).add(s["page"])
            pos.setdefault((z, round(s["loc"]["bbox"][1] / 4)), {}).setdefault(sig(s["text"]), set()).add(s["page"])
    need = max(2, 0.5 * pages)
    # на высоте — и текст в основном тот же (≥ половины листов): название раздела в рамке на той же высоте меняется
    # от раздела к разделу — это текст, а не колонтитул
    return {k for k, ps in cnt.items() if len(ps) >= need} | \
        {k for k, c in pos.items() if len(set().union(*c.values())) >= need and max(map(len, c.values())) >= 0.5 * pages}


def is_running(s, keys, H):
    z = zone(s, H)
    return bool(z) and ((z, sig(s["text"])) in keys or (z, round(s["loc"]["bbox"][1] / 4)) in keys)


def stable_running(s, pool, keys, H):
    """Колонтитул — строка, что стоит так же почти на всех листах: повтор по тексту (_is_running) или строка с меткой
    портала, если на 80 % листов с колонтитулами (pool — повторы и метки портала) на той же высоте похожий текст (OCR
    читает шапку по-разному). Название раздела в рамке (3SZ19: «1. A turbina üzemzavarai» / «2. …») портал метит
    колонтитулом, но оно меняется от раздела к разделу — это текст: в шапке всех листов стояло название второго раздела
    (шапка ломалась на три столбца), а в теле заголовок раздела пропадал в начале раздела и оставался посреди (19.10)."""
    if is_running(s, keys, H):
        return True
    from difflib import SequenceMatcher
    z, y, k = zone(s, H), s["loc"]["bbox"][1], sig(s["text"])
    if not z:
        return False
    have = {t["page"] for t in pool if zone(t, H) == z}
    hit = {t["page"] for t in pool if zone(t, H) == z and abs(t["loc"]["bbox"][1] - y) < 4
           and SequenceMatcher(None, k, sig(t["text"])).ratio() >= 0.8}
    return len(hit) >= 0.8 * len(have)


def table_heads(rules, body, keys, H):
    """Содержимое таблиц, в которых под первой строкой сетки от двух столбцов есть текст тела (не повтор колонтитула), —
    тело, не колонтитул и не бегущее название. Шапка такой таблицы повторяется на каждом листе (Word печатает её на
    каждом листе) — и снималась как колонтитул или бегущее название: у 30RDGT шапка таблицы уходила в шапку DOCX
    узкой сеткой, 15 листов → 68; у 31RDGT заголовки столбцов пропадали со второго листа (20.10). Не только первая
    строка: подписи строк («Berendezés alfanumerika», «Beépítés helye» вверху каждого листа 3PR10, «Alfanumerika»,
    «Állapot» в 3SZ17, 3SZ23, 3FI18) стоят в верхней части листа и повторяются на следующем — их снимало правило
    «название раздела вверху листа повторяет прошлый лист»: в переводе оставались на одном листе из девяти (29.10).
    Свидетельство решателя старшего ранга (E9b-Z, Z7): что внутри таблицы с телом — не колонтитул ни одним правилом
    ниже (повтор, метка портала, название раздела, повтор с прошлого листа). Таблицы листа — docx_layout._tables по
    линиям rules[1]. → {id сегмента}."""
    import fitz
    from docx_layout import _center, _tables
    per = rules[1]
    by_page = {}
    for s in body:
        by_page.setdefault(s["page"], []).append(s)
    sure, maybe = set(), []
    inside = set()
    for pno, (hs, vs) in enumerate(per):
        segs = by_page.get(pno, [])
        for xs, ys, _, _ in _tables(hs, vs):
            if len(xs) < 3 or len(ys) < 3:
                continue
            whole = fitz.Rect(xs[0], ys[0], xs[-1], ys[-1])
            head = fitz.Rect(xs[0], ys[0], xs[-1], ys[1])
            rest = fitz.Rect(xs[0], ys[1], xs[-1], ys[-1])
            body_ = [s for s in segs if rest.contains(fitz.Point(*_center(s))) and not is_running(s, keys, H)]
            if not body_:
                continue                                    # штамп: вся таблица повторяется на листах — колонтитул
            ids = [s for s in segs if head.contains(fitz.Point(*_center(s)))]
            # тело таблицы — и вне зон колонтитулов; таблица целиком в зоне шапки или подвала — штамп листа (ниже первой
            # строки меняется текст: номер, раздел, мусор слоя «���» разной длины), если её первая строка стоит на том
            # же месте и на других листах; тогда шапка таблицы — только если та же строка — шапка таблицы с телом на
            # другом листе. Иначе штамп шёл в тело один раз, у остальных листов шапки не было (ред-тим 20.10);
            # продолжение таблицы в одну-две строки вверху листа — по-прежнему таблица
            if any(not zone(s, H) for s in body_):
                sure |= {s["id"] for s in ids}
                inside |= {s["id"] for s in segs if whole.contains(fitz.Point(*_center(s)))}
            else:
                maybe.append(ids)
    place = lambda s: (sig(s["text"]), round(s["loc"]["bbox"][0] / 6), round(s["loc"]["bbox"][1] / 6))
    seen = Counter(place(s) for p in by_page.values() for s in p)
    sigs = {sig(s["text"]) for p in by_page.values() for s in p if s["id"] in sure}
    return inside | {s["id"] for ids in maybe if ids and (all(sig(s["text"]) in sigs for s in ids)
                                                         or not any(seen[place(s)] > 1 for s in ids)) for s in ids}


def _num_keys(text):
    """Числа строки и их ключи — до 10 знаков перед числом, числа раньше — «#» (как docx_struct.layout.num_keys)."""
    return [(re.sub(r"\d+", "#", text[:m.start()])[-10:], m) for m in re.finditer(r"\d+", text)]


def rows_of(segs):
    """Строки колонтитула листа: сегменты по высоте (±5 pt), в строке — слева направо (как docx_layout._grid)."""
    rows = {}
    for s in sorted(segs, key=lambda s: (s["loc"]["bbox"][1], s["loc"]["bbox"][0])):
        k = next((k for k in rows if abs(k - s["loc"]["bbox"][1]) < 5), s["loc"]["bbox"][1])
        rows.setdefault(k, []).append(s)
    return [sorted(r, key=lambda s: s["loc"]["bbox"][0]) for _, r in sorted(rows.items())]


def _row_spans(segs):
    """Текст колонтитула листа по строкам и место каждого сегмента в нём: (текст, [(начало, сегмент)])."""
    out, at = "", []
    for r in rows_of(segs):
        if out:
            out += " | "
        for j, s in enumerate(r):
            if j:
                out += " "
            at.append((len(out), s))
            out += s["text"].strip()
    return out, at


def sheet_numbers(zsegs):
    """Номер листа оригинала по листам — счётчик в колонтитуле (путь А для распознанных листов, E9b-Z, Z8; как
    docx_struct.layout для текстового слоя): число, у которого по тексту перед ним (10 знаков, по строкам колонтитула —
    OCR рвёт «Oldalszám:» и «41» на куски и ставит число то в первую строку шапки, то во вторую) на листах
    ≥ max(3, ½) разных значений; номер листа — первое такое число на листе; верен, если ≥ 60 % пар соседних листов
    растут на 1, иначе номеров нет. zsegs — {лист от 1: сегменты колонтитула}. → ({лист: номер}, {ключи счётчика})."""
    texts = {p: _row_spans(ss)[0] for p, ss in zsegs.items()}
    vals = {}
    for t in texts.values():
        for k, m in _num_keys(t):
            vals.setdefault(k, []).append(m.group())
    counter = {k for k, v in vals.items() if len(set(v)) >= max(3, 0.5 * len(v))}
    no = {}
    for p, t in texts.items():
        n = next((int(m.group()) for k, m in _num_keys(t) if k in counter), None)
        if n is not None:
            no[p] = n
    pairs = [(no[p], no[p + 1]) for p in no if p + 1 in no]
    if not pairs or sum(b == a + 1 for a, b in pairs) < 0.6 * len(pairs):
        return {}, counter
    # номер — только прочитанный и в монотонном ряду: число не больше соседа слева или не меньше соседа справа (OCR
    # прочёл не ту цифру, повтор) — листа без номера, без угадывания; лист без номера — прежним порядком (продолжает
    # раздел прошлого листа, номер — счётом страниц); сбоев много (номер есть меньше чем у ¾ листов) или ряд всё равно
    # не растёт — номеров нет у всего документа (решение оркестратора по Z8, 09.11)
    bad = {p for p in no if p - 1 in no and no[p] <= no[p - 1] or p + 1 in no and no[p] >= no[p + 1]}
    no = {p: n for p, n in sorted(no.items()) if p not in bad}
    if len(no) < 0.75 * len(texts) or any(b <= a for a, b in zip(list(no.values()), list(no.values())[1:])):
        return {}, counter
    return no, counter


def mark_counter(segs, counter, mark):
    """Копии сегментов колонтитула листа, где цифры счётчика листа (sheet_numbers) заменены знаком mark — там сборка
    ставит поле PAGE (docx_build._zone_text, как метка пути А docx_struct.layout). → (сегменты, число счётчика или None)."""
    text, at = _row_spans(segs)
    hit = next((m for k, m in _num_keys(text) if k in counter), None)
    if hit is None:
        return segs, None
    out = []
    for a, s in at:
        t = s["text"].strip()
        if a <= hit.start() and hit.end() <= a + len(t):
            i = hit.start() - a
            s = dict(s, text=t[:i] + mark * len(hit.group()) + t[i + len(hit.group()):])
        out.append(s)
    return out, hit.group()


class Why(dict):
    """{id сегмента: причина снятия — правило решателя} + heads — id содержимого таблиц с телом (свидетельство «тело»
    старшего ранга, table_heads)."""
    heads = frozenset()


def decide(body, H, pages, rules):
    """Тело листа и снятое (колонтитул, бегущее название раздела). body — сегменты листа с текстом и рамкой; pages —
    листов в документе; rules — docx_layout._page_rules (rules[1] — линии листов: таблицы, свидетельство «тело»
    старшего ранга, table_heads; rules[4] — низ рамки шапки по листам). → (content, Why)."""
    from docx_layout import _role
    run_keys = running_keys(body, H, pages)
    # шапка таблицы (строка заголовков столбцов) повторяется на каждом листе, но это не колонтитул и не бегущее название
    heads = table_heads(rules, body, run_keys, H)
    # распознанная строка с меткой портала «колонтитул» (roles.py) — свидетельство колонтитула второго ранга (_role, п. 1):
    # OCR читает шапку по-разному на разных листах («T0.7. sz. útmutató» / «3.7. sz. útmutató»), и повтор по тексту её
    # не узнаёт (замер на сканах 18.10: колонтитулы правилами 0,27); колонтитул — если стоит так же почти на всех листах
    hdr = {s["id"] for s in body if s["loc"].get("ocr") and s["loc"].get("role") == "header"}
    pool = [s for s in body if s["id"] not in heads and (is_running(s, run_keys, H) or s["id"] in hdr)]
    content = [s for s in body if _role({
        "table_head": s["id"] in heads, "repeat": is_running(s, run_keys, H), "ocr": bool(s["loc"].get("ocr")),
        "label": s["loc"].get("role"), "stable": s["id"] in hdr and stable_running(s, pool, run_keys, H)})
        != "running"]
    why = Why()
    kept = {s["id"] for s in content}
    for s in body:
        if s["id"] not in kept:
            why[s["id"]] = "повтор" if (zone(s, H), sig(s["text"])) in run_keys else \
                "повтор по высоте" if is_running(s, run_keys, H) else "метка портала"
    # меняющееся название раздела в рамке (метка портала, но не на всех листах одинаково) — один раз, первым появлением:
    # OCR читает его на каждом листе чуть по-разному, и точное сравнение с предыдущим листом ниже оставляло копии
    # посреди раздела (3SZ19, 19.10); похожие строки верхней зоны — одна группа
    from difflib import SequenceMatcher
    titles = {sig(s["text"]) for s in pool if s["id"] in hdr and not stable_running(s, pool, run_keys, H)}
    titles = [k for k in titles if len(k) >= 6]
    like = lambda k, q: SequenceMatcher(None, k, q).ratio() >= 0.8
    seen, drop = [], set()
    for s in sorted(content, key=lambda s: (s["page"], s["loc"]["bbox"][1])):
        k = sig(s["text"])
        if zone(s, H) != "head" or s["id"] in heads or not any(like(k, t) for t in titles):
            continue
        if any(like(k, q) for q in seen):
            drop.add(id(s))
        else:
            seen.append(k)
    why.update({s["id"]: "название раздела" for s in content if id(s) in drop})
    content = [s for s in content if id(s) not in drop]
    # бегущее название раздела в рамке сверху листа (3PR42: «1.15. Átfolyás …» на каждом листе раздела) и шапка серии
    # листов (3PR10 «Azonosító: 3PR10/INFO», «verziószám: 0» на листах 20–34) — то же, что в зоне шапки соседнего листа:
    # повтор на подряд идущих листах (на всех листах сразу его не найти: у каждого раздела и серии своё)
    # название — от 6 знаков с буквами: короткое значение («3», «N/A» титульных сведений Eskom) повторяется само по себе
    # и пропадало (17.10)
    # и не строка, что начинается ниже рамки шапки листа: это тело — «1. Ismérvek:» 3SZ19 стоит под рамкой на одном
    # месте почти на каждом листе (раздел — с нового листа) и снималась как повтор с прошлого листа на 39 листах из 68
    # (20.10); лист без рамки — как прежде
    frames = rules[4]
    # «под рамкой» — по середине строки: по верху с допуском 1 pt заголовок сразу под рамкой (верх 91,4 при рамке 92,6,
    # 30RR2126GT стр. 10–12) считался «в рамке», и «A 30RR25S002 ellenőrzése» снималась повтором прошлого листа (цифры —
    # маской) — два заголовка из трёх пропадали (E9_DESIGN §5, F-1)
    under = lambda s: frames[s["page"]] is not None and \
        (s["loc"]["bbox"][1] + s["loc"]["bbox"][3]) / 2 >= frames[s["page"]]
    titled = lambda s: s["id"] not in heads and not under(s) and len(re.sub(r"\W", "", s["text"])) >= 6 \
        and re.search(r"[^\W\d_]{3}", s["text"])
    band = lambda s: s["loc"]["bbox"][3] < 0.2 * H
    # пробелы — без счёта: OCR читает название то с двойным пробелом после номера, то с одним («1.1.  Turbina …»
    # / «1.1. Turbina …», 3SZ19), и повтор с прошлого листа оставался заголовком посреди раздела (19.10)
    key = lambda t: re.sub(r"\s+", " ", sig(t)).strip()
    # повтор — со строкой соседнего листа вне таблиц с телом (свидетельство старшего ранга, table_heads): строка тела
    # листа регистрации «Érintett fejezetek:» 30RL2126GT сверялась с той же подписью в ячейке таблицы изменений на
    # прошлом листе и пропадала (E9b-Z, Z4)
    top = {}
    for s in content:
        if band(s) and s["id"] not in heads:
            top.setdefault(s["page"], set()).add(key(s["text"]))
    # строка-поле шапки — «подпись: значение» («Azonosító: 3PR10/INFO», «Oldalszám:1», «Megnevezés: 30YA12…»): на
    # первом листе серии она — шапка этого листа, как и на следующих, а не текст и не заголовок раздела: оставалась в
    # теле первого листа серии (3PR10 л. 20, 3FI18 л. 33, 30GF*51GT л. 15), у 1SZ01 «Oldalszám:1» полужирной — заголовком;
    # путь А (docx_struct.layout) берёт её в шапку серии из знаков вне тела
    field = lambda s: FIELD.match(s["text"])
    run_from = lambda s: band(s) and titled(s) and key(s["text"]) in top.get(s["page"] + 1, ()) \
        and key(s["text"]) not in top.get(s["page"] - 1, ())
    # первое появление бегущего названия (то же — вверху следующего листа) — заголовок раздела, даже без метки портала:
    # «1.8. Turbina csapágyak rezgése …» 3SZ19 шло обычным абзацем, а со следующего листа снималось как повтор (19.10)
    # (название — полужирное или по центру: строка тела вверху листа, повторённая на следующем, — не заголовок)
    first = {(s["page"], key(s["text"])) for s in content if run_from(s) and not field(s)
             and (s["loc"].get("bold") or s["loc"].get("align") == "center")}
    lead = lambda s: run_from(s) and field(s)
    prev = lambda s: band(s) and titled(s) and key(s["text"]) in top.get(s["page"] - 1, ())
    why.update({s["id"]: "шапка серии" for s in content if lead(s)})
    why.update({s["id"]: "повтор с прошлого листа" for s in content if prev(s)})
    content = [s for s in content if not prev(s) and not lead(s)]
    # (факт листа, старше метки портала — _role: метка «текст» у такой строки заголовок не отменяет)
    content = [dict(s, loc=dict(s["loc"], title=True)) if (s["page"], key(s["text"])) in first else s
               for s in content]
    # и на первом листе раздела: название в рамке сверху, а ниже на том же листе — тот же текст заголовком
    # («1.3/1. SZBV kazetta leesése» дважды подряд, 3PR42)
    norm = lambda t: re.sub(r"\s+", " ", sig(t))
    lower = {}                                  # лист → [(верх, текст)]; заголовок бывает и в верхней пятой части листа
    # только заголовок ниже (полужирный, по центру, метка портала «заголовок»): при любом повторе пропадали «Sziasztok!»
    # письма 31/32SDTÖMGT, «Schnepp Péter» 30RYBSZGT, «BLOKKINDÍTÁSKOR» титула 30TKBSZ36BIGT (E9_DESIGN §5, F-2)
    for s in content:
        # заголовок — как у _role: метка портала «заголовок» или полужирный по центру / с номером; полужирная строка
        # слева без номера — тело («Sziasztok!» письма, «BLOKKINDÍTÁSKOR» в названии титула)
        if s["loc"].get("role") == "heading" or s["loc"].get("bold") and (
                s["loc"].get("align") == "center" or re.match(rf"\s*{PT.ITEM_LABEL}\s", s["text"])):
            lower.setdefault(s["page"], []).append((s["loc"]["bbox"][1], norm(s["text"])))
    over = lambda s: s["loc"]["bbox"][3] < 0.2 * H and titled(s) and any(
        y > s["loc"]["bbox"][3] and t == norm(s["text"]) for y, t in lower.get(s["page"], ()))
    why.update({s["id"]: "над заголовком" for s in content if over(s)})
    content = [s for s in content if not over(s)]
    why.heads = frozenset(heads)
    return content, why
