# -*- coding: utf-8 -*-
"""Единство терминов внутри документа (аудит 29.10, причина №7): одно венгерское слово переведено по-разному в разных
сегментах («VÍZKÖDDEL OLTÓ» → «водяным туманом» ×31, «паром» ×1, «облаком» ×1). Без портала, по готовому res:
1) detect — корни венгерских слов → русские слова, которые с ними сопутствуют; больше одного — вариант разошёлся
   (основа слова из русского — glossary._stem, венгерского — termx.stem);
2) same_text — один и тот же венгерский сегмент целиком переведён по-разному («Név» → «Ф. И. О.» / «Имя»);
3) choose — вариант по глоссарию, иначе по частоте; выбор порталом — следующий этап (поле "chosen" можно заполнить извне);
4) unify — подстановка выбранного варианта во все сегменты документа; в конвейере выключена (HU_UNIFY_TERMS=1 или
   "unify_terms": true в проекте), отчёт — лист «Термины» таблицы вычитки (sheet) — всегда.
Подстановка меняет только слово с тем же окончанием, что есть у выбранного в документе: падеж не угадывается."""
import re
from collections import Counter, defaultdict

import glossary as G
from termx import STOP, FULLWORD, TAG, stem as hu_stem

CYR = re.compile(r"[А-Яа-яЁё]{4,}")
RU_STOP = frozenset("""этот этой этих того если когда после перед также только всех всей быть будет может должен должны
который которая которые которых которой данн каждый любой другой иной тому этого между через более менее при для без
над под при или как так что его их она они оно все вся весь всё нет нельзя""".split())


def _ru_words(text):
    """Слова русского текста: [(поверхностная форма, основа)] в порядке следования."""
    return [(w, G._stem(w)) for w in CYR.findall(TAG.sub(" ", text)) if G._stem(w) not in RU_STOP]


def _root(lw):
    """Корень венгерского слова: основа без суффикса падежа, первые 7 букв (приставки и словоизменение склеивают формы)."""
    return hu_stem(lw)[:7]


def _key(text):
    return re.sub(r"[\W_]+", " ", TAG.sub(" ", text).lower()).strip()


def items_of(segs, res):
    """[(id, венгерский, русский)] — переведённые порталом сегменты (не копии, не из памяти переводов, не с метками)."""
    out = []
    for s in segs:
        r = res.get(str(s["id"])) or {}
        ru = r.get("ru")
        if ru and not r.get("copied") and not r.get("tm") and not r.get("edited") and "shot" not in s["loc"] \
                and not TAG.search(s["text"]) and not TAG.search(ru):
            out.append((s["id"], s["text"], ru))
    return out


def detect(items, min_n=3, min_share=0.25, min_lift=3.0, min_prec=0.6):
    """Слова оригинала, перевод которых разошёлся → [{root, n, hu_forms, variants: [{stem, forms, ids}]}].
    Слово перевода связано с корнем, если встречается минимум в двух сегментах с корнем и доле min_share их, чаще с ним,
    чем без него (min_prec), и в min_lift раз чаще, чем в среднем по документу. Если главный вариант почти всегда стоит
    после/перед одним и тем же словом («водяным туманом»), редкие варианты берутся из того же места сегментов, где
    главного нет («водяным паром», «водяным облаком»)."""
    N = len(items)
    roots, forms_hu = defaultdict(list), defaultdict(Counter)
    ru_w, df = {}, Counter()
    for i, hu, ru in items:
        ws = _ru_words(ru)
        ru_w[i] = ws
        df.update({st for _, st in ws})
        for w in FULLWORD.findall(TAG.sub(" ", hu)):
            lw = w.lower()
            if len(lw) >= 6 and lw not in STOP:
                roots[_root(lw)].append(i)
                forms_hu[_root(lw)][lw] += 1
    out = []
    for r, ids in roots.items():
        ids = sorted(set(ids))
        n = len(ids)
        if n < min_n:
            continue
        c, surf = Counter(), defaultdict(Counter)
        for i in ids:
            for w, st in ru_w[i]:
                surf[st][w.lower()] += 1
            c.update({st for _, st in ru_w[i]})
        main = [st for st, k in c.items() if k >= 2 and k / n >= min_share and k / df[st] >= min_prec
                and (k / n) / (df[st] / N) >= min_lift]
        if not main:
            continue
        seqs = {i: [st for _, st in ru_w[i]] for i in ids}
        variants = {st: [i for i in ids if st in seqs[i]] for st in main}
        # «водяным туманом / паром / облаком»: два главных слова подряд (a → b) — a держит место, b — то, что на нём стоит;
        # в сегментах, где b нет, на месте b стоит другое слово после a — вариант
        slot = {}
        for a in main:
            after = Counter(q[j + 1] for q in seqs.values() for j in range(len(q) - 1) if q[j] == a)
            for b_, k in after.items():
                if b_ in main and k >= 2 and k / c[a] >= 0.6:
                    alt = defaultdict(list)
                    for i in ids:
                        q = seqs[i]
                        if b_ in q:
                            continue
                        for j in range(len(q) - 1):
                            if q[j] == a and q[j + 1] not in main and q[j + 1] not in RU_STOP:
                                alt[q[j + 1]].append(i)
                    if alt:
                        slot[b_] = alt
        for b_, alt in slot.items():
            for x, xs in alt.items():
                variants.setdefault(x, []).extend(xs)
        # вариант — слово, которое не стоит в сегментах с главным (туман/пар/облако взаимоисключающи, туман/водян — нет)
        dom_ids = set(variants[max(main, key=lambda st: c[st])])
        if slot:
            top_b = max(slot, key=lambda b_: len(variants[b_]))
            # главное слово места покрывает большинство сегментов корня, а каждый вариант — редкость (туман 38 из 48;
            # паром, облаком — по одному); иначе это разные вещи («азотных баллонов» и «в баке»), а не вариант перевода
            top = [top_b] + [x for x in slot[top_b] if len(set(slot[top_b][x])) <= max(1, 0.15 * len(set(variants[top_b])))]
            if len(variants[top_b]) < 0.6 * n:
                top = []
        else:
            top = [st for st in main if st in variants and
                   (not set(variants[st]) & dom_ids or len(set(variants[st]) & dom_ids) <= 0.1 * len(set(variants[st])))]
            if max(main, key=lambda st: c[st]) not in top:
                top.append(max(main, key=lambda st: c[st]))
        if len(top) < 2:
            continue
        out.append({"root": r, "n": n, "kind": "slot" if slot else "excl", "hu_forms": [w for w, _ in forms_hu[r].most_common(3)],
                    "variants": [{"stem": st, "forms": [w for w, _ in surf[st].most_common(3)] or [st],
                                  "ids": sorted(set(variants[st]))}
                                 for st in sorted(top, key=lambda s: -len(set(variants[s])))]})
    # те же сегменты вариантов под разными корнями (vízköddel / rendszer / gépház — одна фраза): остаётся самый
    # специфичный корень (меньше сегментов), прочие — в hu_forms
    best = {}
    for g in out:
        k = (g["kind"], tuple(tuple(v["ids"]) for v in g["variants"][1:]))      # редкие варианты те же сегменты
        if k not in best or g["n"] < best[k]["n"]:
            if k in best:
                g["hu_forms"] = g["hu_forms"] + [f for f in best[k]["hu_forms"] if f not in g["hu_forms"]][:3]
            best[k] = g
        else:
            best[k]["hu_forms"] = best[k]["hu_forms"] + [f for f in g["hu_forms"] if f not in best[k]["hu_forms"]][:3]
    out = sorted(best.values(), key=lambda g: -g["n"])
    return out


def same_text(items):
    """Сегмент целиком (без различий в регистре, знаках и числах), переведённый по-разному: [{hu, variants: [{ru, ids}]}]."""
    by = defaultdict(lambda: defaultdict(list))
    for i, hu, ru in items:
        k = _key(hu)
        if len(k) >= 3:
            by[k][ru.strip()].append(i)
    out = []
    for k, d in by.items():
        if len({_key(v) for v in d}) > 1:               # различие только в знаках/регистре («Лист.» и «лист») — не вариант
            out.append({"hu": k, "variants": [{"ru": v, "ids": ids} for v, ids in sorted(d.items(), key=lambda x: -len(x[1]))]})
    out.sort(key=lambda g: -sum(len(v["ids"]) for v in g["variants"]))
    return out


def choose(group, entries=()):
    """Выбранный вариант: основа, которая есть в русской части глоссария у статьи с этим корнем; иначе самый частый.
    Пишет group["chosen"] и group["by"] ("глоссарий" | "частота"); уже заполненное (порталом) не трогает."""
    if group.get("chosen") is not None:
        return group["chosen"]
    vs = group["variants"]
    if "root" in group:
        for e in entries:
            if group["root"] in {_root(w.lower()) for w in FULLWORD.findall(e.hu) if len(w) >= 6}:
                ru_st = {G._stem(w) for w in CYR.findall(e.ru)}
                hit = [v for v in vs if v["stem"] in ru_st]
                if hit:
                    group["chosen"], group["by"] = hit[0]["stem"], "глоссарий"
                    return group["chosen"]
        group["chosen"], group["by"] = vs[0]["stem"], "частота"
    else:
        group["chosen"], group["by"] = vs[0]["ru"], "частота"
    return group["chosen"]


def _same_case(old, new):
    return new.upper() if len(old) > 1 and old.isupper() else new.capitalize() if old[:1].isupper() else new


def unify(items, groups, texts, entries=()):
    """Подстановка выбранного варианта во все сегменты. items — как у detect; groups — detect + same_text; texts —
    {id: русский}, меняется на месте (чего нет — берётся из items). Возвращает правки [{id, was, now, why}] и пропуски
    [{id, why}] (у выбранного слова в документе нет формы с тем же окончанием)."""
    edits, skipped = [], []
    ru_of = {i: ru for i, _, ru in items}
    for g in groups:
        if g.get("kind") == "excl":                     # совпадение слов без общего места во фразе — шумно, только в отчёт
            continue
        ch = choose(g, entries)
        if "root" in g:
            keep = next(v for v in g["variants"] if v["stem"] == ch)
            pool = {}                                   # окончание → форма выбранного слова, виденная в документе
            for i in keep["ids"]:
                for w, st in _ru_words(texts.get(i, ru_of[i])):
                    if st == ch:
                        pool.setdefault(w.lower()[-2:], w.lower())
            for v in g["variants"]:
                if v["stem"] == ch:
                    continue
                for i in v["ids"]:
                    t = texts.get(i, ru_of[i])
                    new = t
                    for w, st in _ru_words(t):
                        if st != v["stem"]:
                            continue
                        f = pool.get(w.lower()[-2:])
                        if f is None:
                            skipped.append({"id": i, "why": f"нет формы «{keep['forms'][0]}» с окончанием «{w[-2:]}» ({w})"})
                            continue
                        new = re.sub(rf"(?<![А-Яа-яЁё]){re.escape(w)}(?![А-Яа-яЁё])", _same_case(w, f), new, count=1)
                    if new != t:
                        edits.append({"id": i, "was": t, "now": new, "why": f"{g['root']}: {v['forms'][0]} → {keep['forms'][0]}"})
                        texts[i] = new
        else:
            for v in g["variants"]:
                if v["ru"] == ch:
                    continue
                for i in v["ids"]:
                    now = ch.upper() if v["ru"].isupper() and not ch.isupper() else ch      # шапка прописными остаётся прописными
                    edits.append({"id": i, "was": texts.get(i, ru_of[i]), "now": now, "why": f"«{g['hu']}»: целиком"})
                    texts[i] = now
    return edits, skipped


def unify_doc(d, entries=()):
    """Документ задания (segs, res): подставляет выбранные варианты в res[...]['ru'], прежний текст — в r['unified'].
    Возвращает (правки, пропуски). Вызывается только под флагом."""
    items = items_of(d["segs"], d["res"])
    edits, skipped = unify(items, same_text(items) + detect(items), {}, entries)
    for e in edits:
        r = d["res"][str(e["id"])]
        r["unified"] = r.get("unified") or e["was"]
        r["ru"] = e["now"]
    return edits, skipped


def sheet(wb, d, entries=()):
    """Лист «Термины» в книгу wb: что разошлось в документе, какой вариант выбран и чем (глоссарий / частота)."""
    from openpyxl.styles import Alignment, Font
    items = items_of(d["segs"], d["res"])
    groups = same_text(items) + detect(items)
    if "Термины" in wb.sheetnames:
        del wb["Термины"]
    if not groups:                                      # расхождений нет — таблица вычитки прежняя, без пустого листа
        return groups
    ws = wb.create_sheet("Термины")
    ws.append(["слово оригинала", "вариант перевода", "сегментов", "выбран", "чем выбран", "тип", "№ сегментов (первые)"])
    for g in groups:
        ch = choose(g, entries)
        what = "/".join(g["hu_forms"]) if "root" in g else f"«{g['hu']}» (сегмент целиком)"
        for v in g["variants"]:
            mine = v.get("stem", v.get("ru")) == ch
            ws.append([what, "/".join(v["forms"]) if "forms" in v else v["ru"], len(v["ids"]), "да" if mine else "",
                       g.get("by", "") if mine else "",
                       {"slot": "то же место фразы", "excl": "сопутствующие слова (проверить)"}.get(g.get("kind"), "сегмент целиком"),
                       ", ".join(str(i) for i in v["ids"][:8])])
    for col, w in zip("ABCDEFG", (28, 50, 10, 8, 12, 24, 40)):
        ws.column_dimensions[col].width = w
    for c in ws[1]:
        c.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    return groups
