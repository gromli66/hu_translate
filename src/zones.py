# -*- coding: utf-8 -*-
"""Колонтитул листа — один решатель для текстового слоя (T) и распознанных листов (O): E9b-Z, E9_DESIGN.md §6.

Решение «эта строка — колонтитул или бегущее название раздела, не тело» жило россыпью в docx_layout.load (повтор по
тексту и по высоте, метка портала, меняющееся название в рамке, повтор с прошлого листа, название над заголовком) и
частью — в пути А (docx_struct.layout, только текстовый слой). Здесь — одна функция decide: вход — сегменты тела по
листам, линии листов (рамка шапки), шапки таблиц (_head_rows), метки портала; выход — тело и причина по каждому снятому
сегменту. Шаг 0 — перенос без смены выхода; правила Z1, Z7, Z4, Z2, Z5, Z8, Z11 — по одному поверх (E9B_Z_REPORT.md)."""
import re

import geom as GM
import patterns as PT


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


class Why(dict):
    """{id сегмента: причина снятия — правило решателя} + heads — id шапок таблиц (свидетельство «тело» высшего ранга, _head_rows)."""
    heads = frozenset()


def decide(body, H, pages, rules, head_rows):
    """Тело листа и снятое (колонтитул, бегущее название раздела). body — сегменты листа с текстом и рамкой; pages —
    листов в документе; rules — _page_rules (rules[4] — низ рамки шапки по листам); head_rows(keys) — id шапок таблиц
    (docx_layout._head_rows: тело, даже если повторяется). → (content, Why)."""
    from docx_layout import _role
    run_keys = running_keys(body, H, pages)
    # шапка таблицы (строка заголовков столбцов) повторяется на каждом листе, но это не колонтитул и не бегущее название
    heads = head_rows(run_keys)
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
    # бегущее название раздела в рамке сверху листа (3PR42: «1.15. Átfolyás …» на каждом листе раздела) — то же, что
    # вверху предыдущего листа, — не текст (на всех листах сразу его не найти: у каждого раздела своё)
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
    # пробелы — без счёта: OCR читает название то с двойным пробелом после номера, то с одним («1.1.  Turbina …»
    # / «1.1. Turbina …», 3SZ19), и повтор с прошлого листа оставался заголовком посреди раздела (19.10)
    key = lambda t: re.sub(r"\s+", " ", sig(t)).strip()
    top = {}
    for s in content:
        if s["loc"]["bbox"][3] < 0.2 * H:
            top.setdefault(s["page"], set()).add(key(s["text"]))
    # первое появление бегущего названия (то же — вверху следующего листа) — заголовок раздела, даже без метки портала:
    # «1.8. Turbina csapágyak rezgése …» 3SZ19 шло обычным абзацем, а со следующего листа снималось как повтор (19.10)
    # (название — полужирное или по центру: строка тела вверху листа, повторённая на следующем, — не заголовок)
    first = {(s["page"], key(s["text"])) for s in content if s["loc"]["bbox"][3] < 0.2 * H and titled(s)
             and (s["loc"].get("bold") or s["loc"].get("align") == "center")
             and key(s["text"]) in top.get(s["page"] + 1, ()) and key(s["text"]) not in top.get(s["page"] - 1, ())}
    prev = lambda s: s["loc"]["bbox"][3] < 0.2 * H and titled(s) and key(s["text"]) in top.get(s["page"] - 1, ())
    why.update({s["id"]: "повтор с прошлого листа" for s in content if prev(s)})
    content = [s for s in content if not prev(s)]
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
