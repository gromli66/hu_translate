# -*- coding: utf-8 -*-
"""Венгерские отрицающие морфемы: разбор слова для подсказки модели и проверка отрицания в переводе.

Системная слабость Qwen 32b (замер 01.10): «tömörtelenség» (негерметичность) верно лишь в 49 из 86 фраз —
«разрыв», «закупорка», «герметичность подтверждена»; «inhermetikusság» → «герметичность».
Модель не разбирает слово на морфемы и теряет отрицание. Лечение без словарей:
1) подсказка-разбор «tömör + -telen (без, не-) + -ség» для каждого такого слова в куске;
2) проверка: в оригинале слов с отрицанием N — в переводе маркеров отрицания не меньше N."""
import re

from glossary import _is_suffix_chain

# суффиксы лишения/отрицания: -talan/-telen, -tlan/-tlen, -atlan/-etlen, -hatatlan/-hetetlen
NEG = re.compile(r"(hatatlan|hetetlen|talan|telen|atlan|etlen|otlan|tlan|tlen)")
HU_WORD = re.compile(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]{6,}")
# слова, где -telen/-talan не несёт отрицания по смыслу перевода или уже «вшито» (hirtelen — внезапно)
NOT_NEG = ("hirtelen", "kedvezőtlen",
           # отрицание «растворяется» в переводе: haladéktalanul → немедленно, maradéktalanul → полностью,
           # közvetlenül → напрямую, függetlenül → вне зависимости (ложные тревоги на прогоне 01.10)
           "haladéktalan", "maradéktalan", "közvetlen", "független")
NEG_WORDS = {"nem", "sem", "nincs", "nincsenek", "ne", "tilos"}   # «nélkül» — нет: «késlekedés nélkül» → «немедленно»
TAIL_EXTRA = ("ság", "ség", "ul", "ül", "ná", "né", "ok", "ek", "ság", "ségé", "ságá", "sága", "sége")

# маркеры отрицания/лишения в русском: не-, без-/бес-, обез-/обес-, де-(аэр/газ/мин), ни-, нет, отсутств, утечк/течь
RU_NEG = re.compile(r"(?<![а-яё])(не|без|бес|обез|обес|ни|де(?=[агмзсф])|отсутств|нет|утеч|теч[ьи]|разгерм|запрещ|наруш)",
                    re.I)
# начинаются на «не», но не отрицание (иначе «необходимо» прятало ошибку «герметичность подтверждена»)
RU_NOT_NEG = re.compile(r"(?<![а-яё])(необходим|немедлен|нескольк|некотор|нередк|неделя|недел|немн|неон)", re.I)


def neg_words(text):
    """Слова оригинала с отрицающей морфемой → [(слово, корень, суффикс, хвост)]."""
    out = []
    for w in HU_WORD.findall(text):
        lw = w.lower()
        if lw.startswith(NOT_NEG):
            continue
        if lw.startswith("inhermetikus") or lw.startswith("instabil"):
            out.append((w, w[2:], "in-", ""))
            continue
        for m in NEG.finditer(lw):
            stem, tail = lw[:m.start()], lw[m.end():]
            if len(stem) >= 3 and (not tail or _is_suffix_chain(tail) or tail in TAIL_EXTRA
                                   or any(tail.startswith(t) and _is_suffix_chain(tail[len(t):]) for t in TAIL_EXTRA)):
                out.append((w, stem, m.group(), tail))
                break
    return out


NEM_PAIR = re.compile(r"(?<![\wÀ-ɏ])(nem|sem)\s+([a-záéíóöőúüű]{4,})", re.I)
# после «nem» — сказуемое, тут модель не ошибается; опасно отрицание признака: «nem sérült» (неповреждённый)
NEM_SKIP = {"kell", "lehet", "szabad", "szükséges", "teljesül", "teljesült", "teljesülnek", "megengedett", "lehetséges",
            "működik", "működött", "indul", "indult", "zár", "zárt", "nyit", "nyitott", "került", "történt", "változik",
            "megfelelő", "lett", "volt", "lesz", "végezhető", "tilos", "elvárt", "szabadon"}


def nem_pairs(text):
    """«nem sérült» и т. п.: отрицание признака отдельным словом. Модель переводила «a nem sérült transzformátor»
    как «неисправной» — смысл наоборот, а проверка не видела потери («не» есть в «неисправной», вычитка 01.10)."""
    return [(n, w) for n, w in NEM_PAIR.findall(text) if w.lower() not in NEM_SKIP]


def hint(text):
    """Подсказка для промпта: разбор слов с отрицанием, или пустая строка."""
    ws = neg_words(text)
    pairs = nem_pairs(text)
    if not ws and not pairs:
        return ""
    if not ws:
        return ("Отрицание отдельным словом относится к следующему слову — сохрани его смысл, не заменяй антонимом:\n" +
                "\n".join(f"- {n} {w} = НЕ + {w}" for n, w in dict.fromkeys(pairs)))
    seen, lines = set(), []
    for w, stem, suf, tail in ws:
        if w.lower() in seen:
            continue
        seen.add(w.lower())
        if suf == "in-":
            lines.append(f"- {w} = in- (отрицание «не-») + {stem}")
        else:
            lines.append(f"- {w} = {stem} + -{suf} (лишение/отрицание: «без», «не-»)" + (f" + -{tail}" if tail else ""))
    lines += [f"- {n} {w} = НЕ + {w} (отрицание отдельным словом — сохрани, не заменяй антонимом)"
              for n, w in dict.fromkeys(pairs)]
    return ("Слова с отрицающей морфемой — отрицание обязано сохраниться в переводе "
            "(tömörtelenség = tömör «плотный, герметичный» + -telen «без» → «негерметичность», а НЕ «герметичность», "
            "«разрыв» или «закупорка»):\n" + "\n".join(lines))


def neg_count_hu(text):
    toks = re.findall(r"[a-záéíóöőúüű]+", text.lower())
    return len(neg_words(text)) + sum(1 for t in toks if t in NEG_WORDS)


def neg_count_ru(text):
    clean = RU_NOT_NEG.sub(" ", text)
    return len(RU_NEG.findall(clean))


def neg_lost(hu, ru):
    """Отрицаний в переводе меньше, чем в оригинале → вероятная потеря/инверсия."""
    n_hu = neg_count_hu(hu)
    return n_hu > 0 and neg_count_ru(ru) < n_hu


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    for t, ru in [("- Amennyiben a felső blokk tömörtelensége egyértelműen megállapítást nyert, intézkedni kell a blokk leállítására.",
                   "- Если герметичность верхнего блока однозначно подтверждена, необходимо принять меры по остановке блока."),
                  ("ÜFK 4.8.3.2. Üzemanyag inhermetikusság", "ПУБЭ 4.8.3.2. Негерметичность топлива"),
                  ("A korlátozó feltétel NEM TELJESÜL, a szivattyú üzemképtelen.", "Ограничивающее условие НЕ ВЫПОЛНЯЕТСЯ, насос неработоспособен."),
                  ("Haladéktalanul el kell kezdeni a sótalanvíz pótlását.", "Незамедлительно начать подпитку обессоленной водой.")]:
        print(neg_words(t), "| потеря:", neg_lost(t, ru))
        print(hint(t))
