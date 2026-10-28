# -*- coding: utf-8 -*-
"""Перевод сегментов порталом: глоссарий + проверки + повтор упавших. Пара языков — Translator(pair=…), см. langs.py:
hu-ru — проверки и постобработка ниже (исходные); ru-en — check_en; прочие пары — check_generic, system_generic.

Translator.run_batch переводит пачку сегментов одним запросом (конвейер подаёт раздел документа ≤6000 знаков
с контекстом предыдущего раздела, см. pipeline.py): в запросе только найденные в тексте термины глоссария
(основной заказчика → строгие → подсказки), контекст, разбор слов с отрицающими морфемами.
Проверки сегмента (жёсткие → повтор поштучно с рассуждением и перечнем замечаний, до 2 раз):
пропуск, метки <fN>, [TAB]/[BR], коды и числа оригинала, остатки венгерского, иероглифы, нет кириллицы,
длина, перевёрнутые действия (открыть/закрыть, вкл/выкл…), потерянные отрицания, обязательные термины комплекта.
Мягкая (только в отчёт): термины основного глоссария. Постобработка: fix_units, fix_caps, first_use."""
import re, json
from pathlib import Path

import glossary as G
import morph
from llm import chat, STAT


SYSTEM_DOMAIN = 'Ты — профессиональный переводчик эксплуатационной документации АЭС с венгерского языка на русский. Документы — АЭС «Пакш» (ВВЭР-440): инструкции по эксплуатации (KU), пределы и условия безопасной эксплуатации (ÜFK), инструкции по ликвидации нарушений и аварий.'
SYSTEM = """Ты — профессиональный переводчик эксплуатационной документации АЭС с венгерского языка на русский. Документы — АЭС «Пакш» (ВВЭР-440): инструкции по эксплуатации (KU), пределы и условия безопасной эксплуатации (ÜFK), инструкции по ликвидации нарушений и аварий.

Правила глоссария (обязательны):
{rules}

Дополнительные правила:
- Переводи полностью и точно, без пропусков, сокращений, пояснений и комментариев от себя. Стиль — официально-деловой, как в российских эксплуатационных инструкциях АЭС.
- Термины из списка «Термины глоссария» используй именно в указанном виде (склоняя по-русски). В записи «венгерский → русский [примечание]» скобки внутри русского варианта — это расшифровка или допустимый вариант, а не часть термина: в текст ставь один основной вариант без этих скобок (SZBV → «СУЗ», а не «СУЗ (ОР СУЗ)»). Примечание в квадратных скобках — подсказка для тебя, в перевод его не переносить.
- Аббревиатуры из глоссария пиши только русским вариантом, без венгерского оригинала в скобках (оригинал при первом употреблении расставит программа). Исключение — аббревиатуры, которые глоссарий велит оставлять латиницей.
- Латинские буквенные обозначения, которых нет в глоссарии (AMV, AVV, DHJ, FHJ, RTK, NNY), оставляй латиницей как в оригинале, не заменяй похожими русскими буквами.
- Коды оборудования и сигналов (30TL02D001, TK52S402, PKHU.GEN.DT.ATL), номера пунктов и ссылок (2.1., ÜFK 4.3.1.), обозначения состояний (M1…M7), числа и знаки (±, ÷, Δ, <, >) переноси посимвольно. Десятичная запятая остается запятой.
- Слова, написанные в оригинале ПРОПИСНЫМИ (NEM LEHET NAGYOBB, MINT; AZONNAL; HA … AKKOR; VAGY; ÉS), переводятся и в переводе тоже пишутся ПРОПИСНЫМИ (НЕ ДОЛЖНА ПРЕВЫШАТЬ; НЕМЕДЛЕННО; ЕСЛИ … ТО; ИЛИ; И).
- Служебные метки: [TAB] — табуляция, [TAB×3] — три табуляции подряд, [BR] — перенос строки, [BR×2] — два переноса; переноси каждую метку в перевод в точности как есть и на то же место относительно текста. Метки форматирования <f1>…</f1>, <f2>…</f2> сохраняй все, переводи текст внутри них.
- Формат ответа: на каждый входной <s id="N">…</s> ровно один <s id="N">перевод</s>, в том же порядке. Никакого другого текста."""

# Обратное направление RU→EN (проект "pair": "ru-en"; ТЗ тренажёра, 20.10): свой промпт и свои проверки.
# Венгерские проверки (остатки венгерского = любая латиница, «нет кириллицы») английский перевод валили бы целиком.
SYSTEM_RU_EN_DOMAIN = "You are a professional translator of technical documentation from Russian into English."
SYSTEM_RU_EN = SYSTEM_RU_EN_DOMAIN + """

Glossary rules (mandatory):
{rules}

Additional rules:
- Translate completely and accurately, without omissions, abridgement, explanations or comments of your own. Style: formal technical English of a technical specification, consistent American spelling.
- Use the terms from "Glossary terms for this fragment" exactly as given. In an entry "Russian → English [note]" the note is a hint for you; do not copy it into the translation.
- No Cyrillic in the translation. A Russian abbreviation missing from the glossary is translated by its meaning (an established English abbreviation or the full English term), never transliterated.
- Codes, identifiers, Latin names and abbreviations (Thai Binh 2 Power Plant, Babcock & Wilcox, KKS, SCADA, TCP/IP), section and item numbers (1.1.1.3, 7.2.) are copied character by character. Numbers stay as in the original; the decimal comma becomes a decimal point (0,4 кВ → 0.4 kV). Units are written in English (кВ → kV, МВт → MW, т/ч → t/h, сек → s).
- Words written in CAPITALS in the original stay in CAPITALS in the translation.
- Russian quotation marks «…» become "…".
- Service markers: [TAB] is a tab, [TAB×3] three tabs in a row, [BR] a line break, [BR×2] two line breaks; copy each marker into the translation exactly as is, in the same position relative to the text (between the corresponding words). Keep all formatting markers <f1>…</f1>, <f2>…</f2>, translating the text inside them.
- Answer format: for each input <s id="N">…</s> exactly one <s id="N">translation</s>, in the same order. No other text."""
CYR_LEFT = re.compile(r"[А-Яа-яЁё]+")
LAT_NAME = re.compile(r"(?<![A-Za-z])[A-Za-z][A-Za-z&]+(?![A-Za-z])")

# Прочие пары (ru-hu, en-ru, hu-en, en-hu; 22.10): общий промпт по языкам пары. Не проверены на комплекте так, как
# hu-ru и ru-en: числа, единицы и кавычки — по правилам языка цели; латиница/кириллица — по письменности источника.
DECIMAL = {"ru": "comma", "hu": "comma", "en": "point"}
UNITS_OF = {"ru": "кВ, МВт, т/ч, с", "hu": "kV, MW, t/h, s", "en": "kV, MW, t/h, s"}
QUOTES = {"ru": "«…»", "hu": "„…”", "en": "\"…\""}


def system_generic(src, tgt):
    import langs as L
    S, T = L.NAME_EN[src], L.NAME_EN[tgt]
    if DECIMAL[src] == DECIMAL[tgt]:
        dec = f"the decimal {DECIMAL[tgt]} stays as in the original"
    else:
        dec = f"the decimal {DECIMAL[src]} becomes a decimal {DECIMAL[tgt]} (0{',' if DECIMAL[src] == 'comma' else '.'}4 → " \
              f"0{',' if DECIMAL[tgt] == 'comma' else '.'}4)"
    script = (f"No Cyrillic in the translation. A Russian abbreviation missing from the glossary is translated by its "
              f"meaning (an established {T} abbreviation or the full {T} term), never transliterated." if src == "ru" else
              f"Latin letter designations, equipment and signal codes missing from the glossary stay in Latin exactly as "
              f"in the original; do not replace them with similar Cyrillic letters." if tgt == "ru" else
              f"Equipment and signal codes, Latin letter designations and proper names are copied as in the original.")
    return (f"You are a professional translator of technical documentation from {S} into {T}." + """

Glossary rules (mandatory):
{rules}

Additional rules:
""" + f"""- Translate completely and accurately, without omissions, abridgement, explanations or comments of your own. Style: formal technical {T}.
- Use the terms from "Glossary terms for this fragment" exactly as given (inflected as {T} grammar requires). In an entry "{S} → {T} [note]" the note is a hint for you; do not copy it into the translation.
- {script}
- Codes, identifiers, names, section and item numbers (2.1., 7.2.) are copied character by character. Numbers stay as in the original; {dec}. Units are written in {T} ({UNITS_OF[tgt]}).
- Words written in CAPITALS in the original stay in CAPITALS in the translation.
- Quotation marks: {QUOTES[tgt]}.
""" + """- Service markers: [TAB] is a tab, [TAB×3] three tabs in a row, [BR] a line break, [BR×2] two line breaks; copy each marker into the translation exactly as is, in the same position relative to the text (between the corresponding words). Keep all formatting markers <f1>…</f1>, <f2>…</f2>, translating the text inside them.
- Answer format: for each input <s id="N">…</s> exactly one <s id="N">translation</s>, in the same order. No other text.""")

CODE = re.compile(r"[\wÀ-ɏ.\-/÷()+±%]*\d[\wÀ-ɏ.\-/÷()+±%]*")
HU_SUFFIX = re.compile(r"-[a-záéíóöőúüű]+$")
LATIN_WORD = re.compile(r"(?<![\w])[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű][a-záéíóöőúüű]{2,}(?![\w])")
CJK = re.compile(r"[　-鿿가-힯฀-๿]")
LETTERS = re.compile(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüűА-Яа-яЁё]+")


def mixed_script(src, out):
    """Слова перевода из букв двух алфавитов («СОORУЖЕНИЕ» — латинские O и R внутри русского слова: сбой портала,
    корпус HU 22.10), кроме записанных так в самом оригинале."""
    have = set(LETTERS.findall(TAG.sub(" ", src)))
    return [w for w in LETTERS.findall(TAG.sub(" ", out))
            if CYR.search(w) and re.search(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]", w) and w not in have]
CYR = re.compile(r"[а-яА-Я]")
TAG = re.compile(r"</?f\d+>")


NUM_GLUED = re.compile(r"(?<=\d\.)(?=(?!(?:docx?|xlsx?|pdf|txt|html?|json|xml|png|jpe?g|zip)\b)[a-záéíóöőúüű]{4,}(?![\w.]))")


def to_model(text):
    """Серия табуляций — одна метка [TAB×4]: подряд идущие [TAB][TAB][TAB][TAB] модель пересчитывает
    неверно (42 из 50 вызовов DOCX-замера были повторами из-за этого, 01.10). Слово, приставшее к номеру точкой
    («ugrás a(z) 4.1.2.29.fejezetre»), отделяется пробелом: слитным модель оставляла его по-венгерски (28 строк, 28.10)."""
    t = re.sub(r"\t+", lambda m: "[TAB]" if len(m.group()) == 1 else f"[TAB×{len(m.group())}]", NUM_GLUED.sub(" ", text))
    return re.sub(r"\n+", lambda m: "[BR]" if len(m.group()) == 1 else f"[BR×{len(m.group())}]", t)


def from_model(text):
    t = re.sub(r"\s*\[TAB(?:[×x](\d+))?\]\s*", lambda m: "\t" * int(m.group(1) or 1), text)
    t = re.sub(r"\s*\[BR(?:[×x](\d+))?\]\s*", lambda m: "\n" * int(m.group(1) or 1), t)
    return t.replace("ё", "е").replace("Ё", "Е")


FILLER = re.compile(r"[.…_]{3,}")       # линия-заполнитель бланка: «20.......év.........hó.....nap»


def codes_of(src):
    out = set()
    for m in CODE.finditer(TAG.sub(" ", FILLER.sub(" ", src))):
        c = HU_SUFFIX.sub("", m.group().strip(".,;:()-"))
        c = c.strip(".,;:()")
        if not c or not re.search(r"\d", c):
            continue
        # «70 %-ig» → «70»; «3-án» → «3»; «6.csoport» → «6»; «110%os» → «110%»; «128bar» → «128»; «max.2» → «2»
        c = re.sub(r"-[a-záéíóöőúüű]+$", "", c)
        c = re.sub(r"(?<=\d)\.[a-záéíóöőúüű]{2,}$", "", c)
        c = re.sub(r"(?<=[\d%])[a-záéíóöőúüű]+$", "", c)
        c = re.sub(r"^[a-záéíóöőúüű]+\.", "", c)
        c = re.sub(r"(?<=^\d)\.[A-ZÁÉÍÓÖŐÚÜŰ][a-záéíóöőúüű]{2,}$|(?<=^\d\d)\.[A-ZÁÉÍÓÖŐÚÜŰ][a-záéíóöőúüű]{2,}$", "", c)  # «5.Teszt» → «5»
        c = re.sub(r"^[a-záéíóöőúüű]{3,}(?:-[a-záéíóöőúüű]+)?\((?=\w*\d)", "", c)     # «üzemállapotban(M2», «bar-os(123» → «M2», «123»
        c = re.sub(r"^[a-záéíóöőúüű]{3,}-(?=\d)", "", c)                               # «sebességét-5» → «5»
        c = re.sub(r"^a(?=\d{2}[A-Z])", "", c)                                         # «a30TH60S201» (артикль «a» + код) → «30TH60S201»
        c = re.sub(r"(?<=\d)\.[A-ZÁÉÍÓÖŐÚÜŰ][a-záéíóöőúüű]{2,}$", "", c)             # «2.1.Személyi» → «2.1»
        c = c.strip(".,;:()")
        um = re.fullmatch(r"(\d[\d.,]*)?(?:mbar|bar|kV|kW|MW|mA|nA|Hz)/(\d+)", c)     # «2bar/10 perc» → «2» и «10»: единица переводится
        if um:
            out.update(g for g in um.groups() if g)
            continue
        c = re.sub(r"(?<=\d)(kV|kW|MW|mA|nA|Hz|mbar|bar)$", "", c)   # «0,4kV» → «0,4»: единица переводится (fix_units)
        if c:
            out.add(c)
    return out


def code_present(c, nru, gloss_abbr):
    """Код есть в переводе. Код с переводимой аббревиатурой глоссария (ÜV-1 → АЗ-1) — достаточно номера.
    Единица («g/dm3») или число с единицей («400KV») — засчитывается перевод единицы (fix_units): верное «400 кВ»
    давало замечание, а неверное «400KV» проходило чисто, и повтор перевода выбирал неверное (редтим 05.10)."""
    if _norm_code(c) in nru:
        return True
    low = c.lower()
    umap = {u.lower(): r for u, r in UNITS}
    if low in umap:
        flat = lambda t: t.replace("³", "3").replace("²", "2")        # «дм3/ч» и «дм³/ч» — одна единица
        return flat(umap[low]) in flat(nru) or low in nru.lower()
    mu = re.match(r"^(\d[\d.,]*)\s?([a-z/0-9]+)$", low)
    if mu and mu.group(2) in umap:
        return _norm_code(mu.group(1)) in nru
    m = re.match(r"^([A-ZÁÉÍÓÖŐÚÜŰ]{2,})-?(\d[\w.\-/]*)$", c)
    if m and m.group(1) in gloss_abbr and _norm_code(m.group(2)) in nru:
        return True
    m = re.match(r"^(\d+)\.\s?([A-ZÁÉÍÓÖŐÚÜŰ]{2,})$", c)        # «1.VT» → «СВО-1»: номер впереди, аббревиатура переводится
    return bool(m and m.group(2) in gloss_abbr and re.search(rf"(?<!\d){m.group(1)}(?!\d)", nru))


def _norm_code(s):
    return s.replace(" ", "").replace("‑", "-").replace("–", "-").replace("—", "-")


HU_CAPS_WORDS = frozenset("""nem lehet nagyobb kisebb mint vagy és ha akkor azonnal tilos csak kell egyenlő legalább
legfeljebb több kevesebb igen nincs van nyitva zárva fel le be ki nyit zár indul áll üzemel teljesül
megengedett szabad sem amíg majd után előtt""".split())

LAT_ABBR = re.compile(r"(?<![\wÀ-ɏ])[A-ZÁÉÍÓÖŐÚÜŰ]{2,}(?![\wÀ-ɏ])")


NUM = re.compile(r"\d+")
CAPS_PHRASE = re.compile(r"(?<![\wÀ-ɏ])[A-ZÁÉÍÓÖŐÚÜŰ]{2,}(?:[ ,]+[A-ZÁÉÍÓÖŐÚÜŰ]{2,})+(?![\wÀ-ɏ])")


# Обратное действие: в оригинале только «открыть», в переводе только «закрыть» (и наоборот). Смысловая проверка порталом
# пропустила «be kell nyitni» → «закрыть» (свежая выборка, 02.10); код ловит это без модели.
# (название, HU «да», HU «нет», RU «да», RU «нет»)
ACTIONS = [
    ("открыть/закрыть",
     r"(?<![a-záéíóöőúüű])(ki|meg|vissza|be)?nyit(?!ott\s+primer)",
     r"(?<![a-záéíóöőúüű])(le|be|vissza|el)?zár(?!lat|ójel|óvíz|ó\s+víz)(?![a-záéíóöőúüű]*kör)",
     r"открыт|открыв|откро|открой|открыл", r"закрыт|закрыв|закро|закрой|закрыл"),
    ("включить/отключить",
     r"(?<![-a-záéíóöőúüű])bekapcsol|(?<![-\w])be\s+kell\s+kapcsolni",
     r"(?<![a-záéíóöőúüű])kikapcsol|(?<![-\w])ki\s+kell\s+kapcsolni",
     r"(?<![а-я])включ", r"отключ|выключ(?!ател)|обесточ"),
    ("пуск/останов",
     r"(?<![a-záéíóöőúüű])(el|újra|vissza)?indít|(?<![a-záéíóöőúüű])indul(?!tak\s+ki)",
     r"leállít|leáll(?!ás)",
     r"пуск|запуст|запуска|пущен|пусти|пустят", r"останов|остановл|остана"),
    ("нагрузить/разгрузить",                          # «leterhelni» → «Нагрузить» (3PR04, слепая вычитка 05.10)
     r"felterhel", r"leterhel|(?<![a-záéíóöőúüű])le\s+és\s+felterhel",      # «le és felterhelés» = разгрузка и нагружение
     r"(?<!раз)нагру[зж]|набор мощност|набрать мощност|набора мощност", r"разгру[зж]|снижени[ея] нагрузк|снизить нагрузк"),
]


def action_flip(src, ru):
    h, r = src.lower(), ru.lower()
    out = []
    for name, hy, hn, ry, rn in ACTIONS:
        a, b = bool(re.search(hy, h)), bool(re.search(hn, h))
        c, d = bool(re.search(ry, r)), bool(re.search(rn, r))
        yes, no = name.split("/")
        if a and not b and d and not c:
            out.append(f"обратное действие: в оригинале «{yes}», в переводе «{no}»")
        if b and not a and c and not d:
            out.append(f"обратное действие: в оригинале «{no}», в переводе «{yes}»")
    return out


# торговые марки и стандарты, которые остаются латиницей, хотя стоят среди прописных слов фразы («STORZ KAPOCS», «ASME OM KÓD»)
LATIN_BRANDS = frozenset({"STORZ", "ASME"})
NAME_FIELD = re.compile(r"(?i)\bNév\s*:\s*([^|\n]*?)(?=\s+(?:Szervezet|Aláírás|Dátum|Beosztás)\b|\||$)")
MAIL = re.compile(r"([^;<:\n]*?)\s*<([^<>@\s]+@[^<>\s]+)>")
HYPHEN_CAPS = re.compile(r"(?<=[A-ZÁÉÍÓÖŐÚÜŰ])- (?=[A-ZÁÉÍÓÖŐÚÜŰ]{2})")      # «MENNYI- SÉG» — перенос слова в узкой ячейке


def name_words(src):
    """Слова-имена оригинала: поле «Név:», «Фамилия Имя <адрес>» и сам адрес. По глоссарию остаются латиницей —
    замечанием «венгерский остался» на них шумело больше половины всех замечаний партии (28.10)."""
    plain = TAG.sub(" ", src)
    out = set()
    for m in NAME_FIELD.finditer(plain):
        out.update(LETTERS.findall(m.group(1)))
    for m in MAIL.finditer(plain):
        out.update(LETTERS.findall(m.group(1)))
        out.update(LETTERS.findall(m.group(2)))
    return out


def check(src, ru, allowed_latin=frozenset(), gloss_abbr=frozenset(), vocab=frozenset(), abbr_keep=frozenset()):
    """Жёсткие замечания к переводу сегмента (пустой список — годен).
    gloss_abbr — аббревиатуры глоссария (их переводят); vocab — строчные слова документа:
    прописное слово, которое в документе встречается и строчными (VAGY ~ vagy), — это текст, его переводят;
    прочие латинские прописные (AMV, DHJ) — обозначения, обязаны остаться латиницей."""
    issues = []
    if not ru.strip():
        return ["пусто"]
    if sorted(TAG.findall(src)) != sorted(TAG.findall(ru)):
        issues.append("метки форматирования")
    if src.count("\t") != ru.count("\t") or src.count("\n") != ru.count("\n"):
        issues.append("[TAB]/[BR]")
    nru = _norm_code(ru)
    miss = [c for c in codes_of(src) if not code_present(c, nru, gloss_abbr)]
    if miss:
        issues.append("нет кодов/чисел: " + ", ".join(sorted(miss)[:6]))
    # обратная проверка: число в переводе, которого нет в оригинале (OCR «ti ≤ 70» → «tiz 70» → «при 10 70 °C»)
    src_nums = {n.replace(",", ".") for n in re.findall(r"\d+(?:[.,]\d+)?", TAG.sub(" ", src))}
    # «30RS61S004,30RS61S005» (коды через запятую без пробела): «004,30» — не дробь; в переводе «004, 30RS61S005»
    src_nums |= {p for m in re.finditer(r"(?<=[A-Za-z])\d+[.,]\d+", TAG.sub(" ", src)) for p in re.split(r"[.,]", m.group())}
    extra = sorted({n for n in re.findall(r"\d+(?:[.,]\d+)?", TAG.sub(" ", ru))
                    if len(re.sub(r"\D", "", n)) >= 2 and n.replace(",", ".") not in src_nums})
    if extra:
        issues.append("лишнее число: " + ", ".join(extra[:4]))
    plain = HYPHEN_CAPS.sub("", TAG.sub(" ", src))
    phrase_words = set()                  # слова прописных фраз оригинала — текст, их переводят
    for m in CAPS_PHRASE.finditer(plain):
        words = [w for w in LAT_ABBR.findall(m.group()) if w not in abbr_keep]
        if any(len(w) >= 4 and len(re.findall(r"[AEIOUÁÉÍÓÖŐÚÜŰ]", w)) >= 2 for w in words):
            phrase_words.update(words)
    if not plain.isupper():           # строка целиком прописными — это текст (НЕ ДОПУСКАЕТСЯ), а не обозначения
        lost = sorted({a for a in LAT_ABBR.findall(plain) if a not in gloss_abbr and len(a) <= 6
                       and a not in phrase_words
                       and a.lower() not in vocab and not re.search(r"[ÁÉÍÓÖŐÚÜŰ]", a)} - set(LAT_ABBR.findall(ru)))
        if lost:
            issues.append("латиница потеряна: " + ", ".join(lost[:6]))
    left = [w for w in LATIN_WORD.findall(TAG.sub(" ", ru)) if w not in allowed_latin]
    caps_ru = set(LAT_ABBR.findall(TAG.sub(" ", ru)))
    # прописное слово документа (VAGY ~ vagy) — текст; аббревиатуры раздела глоссария (BRO, ÁNEREM) остаются латиницей законно
    left += [a for a in caps_ru if a.lower() in vocab and a not in abbr_keep and len(a) >= 3]   # «NR» — индекс, не слово
    # прописная фраза оригинала из 2+ слов (NEM LETT ELVÉGEZVE, KISEBB VAGY EGYENLŐ) — текст, а не обозначения
    left += [w for w in phrase_words if w in caps_ru and len(w) >= 3 and w not in left]
    names = name_words(src) | {w for d in DESIGNATOR.findall(plain) for w in LETTERS.findall(d)}     # обозначения — код
    # прописное слово внутри обозначения с цифрами («3SZ17/INFO») — часть кода, а не слово текста
    names |= {w for tok in re.findall(r"\S*\d\S*", plain) for w in LETTERS.findall(tok) if w.isupper()}
    left = [w for w in left if w not in names and w not in LATIN_BRANDS]
    if left:
        issues.append("венгерский остался: " + ", ".join(left[:6]))
    if CJK.search(ru):
        issues.append("иероглифы")
    mixed = mixed_script(src, ru)
    if mixed:
        issues.append("слово из двух алфавитов: " + ", ".join(mixed[:4]))
    if morph.neg_lost(TAG.sub(" ", src), TAG.sub(" ", ru)):
        issues.append("отрицание потеряно: " + ", ".join(w for w, *_ in morph.neg_words(TAG.sub(" ", src))[:4]))
    issues += action_flip(TAG.sub(" ", src), TAG.sub(" ", ru))
    plain_src = TAG.sub("", src)
    if re.search(r"[a-záéíóöőúüű]{3,}", plain_src) and not CYR.search(ru):
        issues.append("не переведено")
    if len(plain_src) >= 25:
        k = len(TAG.sub("", ru)) / len(plain_src)
        if not 0.5 <= k <= 2.6:
            issues.append(f"длина ×{k:.2f}")
    return issues


def check_en(src, en):
    """Жёсткие замечания к английскому переводу русского сегмента (RU→EN): пусто, метки, [TAB]/[BR], числа (десятичная
    запятая → точка), лишние числа, латиница оригинала (названия, коды), кириллица в переводе, иероглифы, длина."""
    if not en.strip():
        return ["пусто"]
    issues = []
    if sorted(TAG.findall(src)) != sorted(TAG.findall(en)):
        issues.append("метки форматирования")
    if src.count("\t") != en.count("\t") or src.count("\n") != en.count("\n"):
        issues.append("[TAB]/[BR]")
    # разряды: «18 636 856,37» (пробел, неразрывный) = «18,636,856.37» (запятая — только в английском: в русском
    # «0,125» — дробь); договор: каждая сумма уходила на повтор
    group = lambda t, seps: re.sub(rf"(?<=\d)[{seps}](?=\d{{3}}(?!\d))", "", TAG.sub(" ", t))
    nums = lambda t: {n.replace(",", ".") for n in re.findall(r"\d+(?:[.,]\d+)?", t)}
    s_n, e_n = nums(group(src, "   ")), nums(group(en, "   ,"))
    miss = sorted(s_n - e_n)
    if miss:
        issues.append("нет кодов/чисел: " + ", ".join(miss[:6]))
    extra = sorted(n for n in e_n - s_n if len(re.sub(r"\D", "", n)) >= 2)
    if extra:
        issues.append("лишнее число: " + ", ".join(extra[:4]))
    lost = sorted(set(LAT_NAME.findall(TAG.sub(" ", src))) - set(LAT_NAME.findall(TAG.sub(" ", en))))
    if lost:
        issues.append("латиница потеряна: " + ", ".join(lost[:6]))
    cyr = CYR_LEFT.findall(TAG.sub(" ", en))
    if len(cyr) > 0.5 * len(re.findall(r"\w+", TAG.sub(" ", en))):
        issues.append("не переведено")                 # badness: хуже любого перевода с замечаниями
    elif cyr:
        issues.append("кириллица осталась: " + ", ".join(dict.fromkeys(cyr[:6])))
    if CJK.search(en):
        issues.append("иероглифы")
    plain_src = TAG.sub("", src)
    if len(plain_src) >= 25:
        k = len(TAG.sub("", en)) / len(plain_src)
        if not 0.45 <= k <= 2.0:
            issues.append(f"длина ×{k:.2f}")
    return issues


THOUSANDS = {"ru": "   ", "hu": "   ", "en": ",  "}
HU_LETTERS = re.compile(r"[A-Za-z]*[áéíóöőúüűÁÉÍÓÖŐÚÜŰ][A-Za-záéíóöőúüűÁÉÍÓÖŐÚÜŰ]*")


def check_generic(src, out, s_lang, t_lang):
    """Жёсткие замечания к переводу для пар без своей проверки (ru-hu, en-ru, hu-en, en-hu): пусто, метки,
    [TAB]/[BR], числа (разряды и десятичный знак — по языкам источника и цели), лишние числа, латиница русского
    источника, остатки языка источника (кириллица; венгерские буквы; английские слова в русском переводе),
    иероглифы, длина."""
    if not out.strip():
        return ["пусто"]
    issues = []
    if sorted(TAG.findall(src)) != sorted(TAG.findall(out)):
        issues.append("метки форматирования")
    if src.count("\t") != out.count("\t") or src.count("\n") != out.count("\n"):
        issues.append("[TAB]/[BR]")
    group = lambda t, lang: re.sub(rf"(?<=\d)[{THOUSANDS[lang]}](?=\d{{3}}(?!\d))", "", TAG.sub(" ", t))
    nums = lambda t: {n.replace(",", ".") for n in re.findall(r"\d+(?:[.,]\d+)?", t)}
    s_n, o_n = nums(group(src, s_lang)), nums(group(out, t_lang))
    miss = sorted(s_n - o_n)
    if miss:
        issues.append("нет кодов/чисел: " + ", ".join(miss[:6]))
    extra = sorted(n for n in o_n - s_n if len(re.sub(r"\D", "", n)) >= 2)
    if extra:
        issues.append("лишнее число: " + ", ".join(extra[:4]))
    plain_s, plain_o = TAG.sub(" ", src), TAG.sub(" ", out)
    if s_lang == "ru":
        lost = sorted(set(LAT_NAME.findall(plain_s)) - set(LAT_NAME.findall(plain_o)))
        if lost:
            issues.append("латиница потеряна: " + ", ".join(lost[:6]))
    left = []
    if s_lang == "ru":
        left = CYR_LEFT.findall(plain_o)
    elif s_lang == "hu":
        left = [w for w in HU_LETTERS.findall(plain_o) if w in plain_s]
    elif t_lang == "ru":                                  # en → ru: английские слова источника в переводе
        src_words = set(re.findall(r"(?<![A-Za-z])[a-z]{4,}(?![A-Za-z])", plain_s))
        left = [w for w in re.findall(r"(?<![A-Za-z])[a-z]{4,}(?![A-Za-z])", plain_o) if w in src_words]
    words = re.findall(r"\w+", plain_o)
    if t_lang == "ru" and re.search(r"[A-Za-z]{3,}", plain_s) and not CYR.search(plain_o) \
            or left and len(left) > 0.5 * len(words) \
            or len(plain_s.strip()) >= 15 and re.sub(r"\s+", " ", plain_o).strip() == re.sub(r"\s+", " ", plain_s).strip():
        issues.append("не переведено")
    elif left:
        issues.append({"ru": "кириллица осталась: ", "hu": "венгерский остался: "}.get(s_lang, "английский остался: ")
                      + ", ".join(dict.fromkeys(left[:6])))
    if CJK.search(out):
        issues.append("иероглифы")
    mixed = mixed_script(src, out) if t_lang == "ru" else []
    if mixed:
        issues.append("слово из двух алфавитов: " + ", ".join(mixed[:4]))
    plain_src = TAG.sub("", src)
    if len(plain_src) >= 25:
        k = len(TAG.sub("", out)) / len(plain_src)
        if not 0.45 <= k <= 2.2:
            issues.append(f"длина ×{k:.2f}")
    return issues


UNITS = [("mg/dm3", "мг/дм³"), ("g/dm3", "г/дм³"), ("mg/kg", "мг/кг"), ("g/kg", "г/кг"), ("kg/cm2", "кгс/см²"),
         ("m3/h", "м³/ч"), ("m3/óra", "м³/ч"), ("dm3/óra", "дм³/ч"), ("t/óra", "т/ч"), ("l/óra", "л/ч"), ("kg/óra", "кг/ч"), ("mm/s", "мм/с"), ("t/h", "т/ч"), ("l/h", "л/ч"), ("mbar", "мбар"), ("bar", "бар"),
         ("kV", "кВ"), ("KV", "кВ"), ("kW", "кВт"), ("MW", "МВт"), ("Hz", "Гц"), ("mA", "мА"), ("nA", "нА"), ("1/min", "1/мин")]
UNIT_RE = re.compile(r"(?<=\d)\s?(" + "|".join(re.escape(u) for u, _ in UNITS) + r")(?!\w|/(?!\d))")   # «1,5 bar/10 мин» — единица перед «/10»
PER_MIN = re.compile(r"(?<![\w/])1/min(?!\w)")             # «Fordulatszám (1/min)» — без числа впереди


# несуществующие слова, которые модель собирает сама от глагола глоссария («вывесить плакат» → «ВЫВЕСИВАНИЕ»: 107 строк
# партии «Пакш», 28.10): основа — слово, замена — настоящее; регистр сохраняется
NONWORDS = [(re.compile(r"(?i)вывесиван"), "вывешиван")]


def _same_case(old, new):
    return new.upper() if old.isupper() else new.capitalize() if old[:1].isupper() else new


def fix_nonwords(ru):
    for rx, new in NONWORDS:
        ru = rx.sub(lambda m: _same_case(m.group(), new), ru)
    return ru


def fix_units(ru):
    """Единицы после числа — по разделу «Единицы» глоссария: «10 g/dm3» → «10 г/дм³», «0,4kV» → «0,4 кВ»
    (модель оставляла латиницу, вычитка v2 01.10)."""
    m = dict(UNITS)
    # служебная пометка глоссария, скопированная моделью в текст («Собрать электрическую схему [уточнить] …»):
    # 35 сегментов v3, 21 — v31 (слепое сравнение 04.10); убирается здесь, т. к. fix_units стоит на всех путях вывода
    ru = re.sub(r"\s*\[(уточнить|проверено Claude)\]", "", ru)
    return fix_nonwords(PER_MIN.sub("1/мин", UNIT_RE.sub(lambda x: " " + m[x.group(1)], ru)))


CYR_CAPS = re.compile(r"(?<![А-ЯЁа-яё])[А-ЯЁ]{5,}(?![А-ЯЁа-яё])")


def fix_caps(src, ru, keep=frozenset()):
    """Лишние прописные: в оригинале нет ни одного прописного слова (кроме обозначений), а модель пишет
    «НЕМЕДЛЕННО отключить», «ОСТАНОВИТЬ турбину» (правило про AZONNAL → НЕМЕДЛЕННО применено не к месту,
    вычитка v2 01.10). Слова от 5 букв понижаются; первое слово сегмента — с заглавной. keep — прописные слова,
    которые пишет сам глоссарий (ДАПЭН, АПЭН, САОЗ): «KÜTSZ» без слова с гласными в оригинале давало «дапэн», 25 строк 28.10."""
    plain = TAG.sub(" ", src)
    # «НЕ затронутой»: подсказка «nem X = НЕ + X» давала прописное «НЕ», когда в оригинале «nem» строчное (02.10)
    if not re.search(r"(?<![\wÀ-ɏ])(NEM|NE|SEM)(?![\wÀ-ɏ])", plain):
        ru = re.sub(r"(?<![А-ЯЁа-яё])НЕ(?=\s+[а-яё])",
                    lambda m: "Не" if re.fullmatch(r"[\s\-–•\d.)(]*", ru[:m.start()]) else "не", ru)
    if re.search(r"(?<![\wÀ-ɏ])[A-ZÁÉÍÓÖŐÚÜŰ]{2,}[ÁÉÍÓÖŐÚÜŰ]*[A-ZÁÉÍÓÖŐÚÜŰ]*(?![\wÀ-ɏ])", plain):
        words = re.findall(r"(?<![\wÀ-ɏ])[A-ZÁÉÍÓÖŐÚÜŰ]{3,}(?![\wÀ-ɏ])", plain)
        # в оригинале есть прописное СЛОВО (с гласными, не аббревиатура) — прописные в переводе законны
        if any(len(re.findall(r"[AEIOUÁÉÍÓÖŐÚÜŰ]", w)) >= 2 for w in words):
            return ru
    def low(m):
        w = m.group()
        if w in keep:
            return w
        start = m.start() == 0 or re.fullmatch(r"[\s\-–•\d.)(]*", ru[:m.start()]) is not None
        return w.capitalize() if start else w.lower()
    return CYR_CAPS.sub(low, ru)


LIST_ISSUES = ("нет кодов/чисел", "лишнее число", "латиница потеряна", "венгерский остался", "кириллица осталась",
               "английский остался")


def badness(issues):
    """Пустой/непереведённый хуже любого перевода с замечаниями (иначе при равном числе замечаний
    оставался пустой — 6 пустых на ÜFK I, замер 01.10). Списочные замечания весят по числу позиций:
    «венгерский остался: Digirec» лучше, чем «… szint, szint, Blokkolás, Digirec» (раньше — поровну, редтим 05.10)."""
    n = 0
    for i in issues:
        if i in ("пусто", "не переведено"):
            n += 100
        elif i.startswith(LIST_ISSUES) and ":" in i:
            n += len([x for x in i.split(":", 1)[1].split(",") if x.strip()]) or 1
        else:
            n += 1
    return n


def gloss_misses(src, ru, entries, outer=None):
    """Мягкая проверка: русские термины найденных записей глоссария есть в переводе (по основам).
    Запись, все места которой лежат внутри более длинного термина (outer — весь глоссарий), не проверяется:
    «főgőz → острый пар» внутри «főgőz kollektor → главный паровой коллектор» давало ложные тревоги (v3, 02.10)."""
    low = ru.lower()
    miss = []
    found = G.find(src, entries)
    wider = G.find(src, outer, limit=100000) if outer is not None else found
    sp = {}
    span_of = lambda x: sp.setdefault(id(x), G.spans(src, x))
    def nested(e):
        mine = span_of(e)
        return bool(mine) and all(any(f is not e and j2 - i2 > j - i and i2 <= i and j <= j2
                                      for f in wider for i2, j2 in span_of(f)) for i, j in mine)
    for e in found:
        if nested(e):
            continue
        variants = [v.strip() for v in re.split(r"\s+/\s+|;", re.sub(r"\([^)]*\)", "", e.ru)) if v.strip()]
        ok = False
        for v in variants:
            ws = [w for w in re.findall(r"[А-Яа-яA-Za-z0-9]+", v) if len(w) >= 4 or w.isupper()]
            if not ws:
                ok = True; break
            if all(w.lower()[:max(3, min(5, len(w) - 2))] in low for w in ws):
                ok = True; break
        if not ok:
            miss.append(e.hu)
    return miss


DESIGNATOR = re.compile(r"(?<![\w.])\d+\.[A-ZÁÉÍÓÖŐÚÜŰ]{1,6}(?:\.[A-ZÁÉÍÓÖŐÚÜŰ0-9]{1,6}){2,}(?![\w])")


def mask_designators(text):
    """Точечные обозначения оборудования (1.GF.H.ÁG.SZEK.FEDÉL) → метки «X1», «X2»: {метка: обозначение}."""
    back = {}
    for n, c in enumerate(dict.fromkeys(DESIGNATOR.findall(text)), 1):
        back[f"X{n}"] = c
        text = re.sub(rf"(?<![\w.]){re.escape(c)}(?![\w])", f"X{n}", text)
    return text, back


def unmask_designators(ru, back):
    for k, c in back.items():
        ru = re.sub(rf"(?<![\w.]){k}(?![\w])", lambda _: c, ru)
    return ru


class Translator:
    def __init__(self, glossary, think=False, retry_think=True, extra=None, domain=None, tm=None, pair=None):
        import langs as L
        self.src, self.tgt = L.parse_pair(pair)
        self.hu_ru, self.ru_en = (self.src, self.tgt) == ("hu", "ru"), (self.src, self.tgt) == ("ru", "en")
        self.tm = json.loads(Path(tm).read_text(encoding="utf-8")) if tm and Path(tm).exists() else {}
        # глоссарий, записанный в обратную сторону (шапка «| Венгерский | Русский |» у пары ru-hu), — развёрнутым
        self.rules, self.entries = G.load(glossary, None if self.hu_ru else (self.src, self.tgt))
        self.main_ids = {id(e) for e in self.entries}             # записи основного глоссария — приоритет в промпте
        self.extra_n, self.extra_entries = 0, []
        if extra and Path(extra).exists():
            # глоссарий документа/комплекта (termx.py, проверен): основной приоритетнее — дубли ключей отбрасываются;
            # его термины обязательны — сегмент без них уходит на повтор (единообразие по всему комплекту)
            # дубль — только полное совпадение термина: проверка «G.find внутри ключа» выкидывала «fel kell bórozni»
            # (нашлось «kell») и «biztonsági hűtővíz» — до модели доходило 75 из 91 записи (01.10)
            main = {tuple(w.lower() for w in seq) for e in self.entries for seq in e.alts}
            add = [e for e in G.load(extra, None if self.hu_ru else (self.src, self.tgt))[1]
                   if not any(tuple(w.lower() for w in seq) in main for seq in e.alts)]
            self.entries += add
            self.extra_n, self.extra_entries = len(add), add
        # строгие термины комплекта — один и тот же список на все проверки (кэш индекса узнаёт список по объекту)
        self.strict_entries = [e for e in self.extra_entries if "строгий" in (e.note or "")] or self.extra_entries
        self.abbrs = G.abbreviations(self.entries)
        self.allowed_latin = self._allowed_latin()
        self.ru_caps = frozenset(w for e in self.entries for w in re.findall(r"(?<![А-ЯЁа-яё])[А-ЯЁ]{5,}(?![А-ЯЁа-яё])", e.ru))
        # всё, что глоссарий переводит: аббревиатуры раздела + прописные венгерские ключи остальных записей
        self.gloss_abbr = frozenset(set(self.abbrs) | {w for e in self.entries for seq in e.alts for w in seq if G._is_abbr(w)})
        self.think, self.retry_think = think, retry_think
        self.vocab = frozenset()
        base = SYSTEM if self.hu_ru else SYSTEM_RU_EN if self.ru_en else system_generic(self.src, self.tgt)
        self.system = base.format(rules=self.rules)
        if domain:                                  # другая предметная область/комплект — первая фраза промпта
            first = SYSTEM_DOMAIN if self.hu_ru else base.split("\n", 1)[0]
            self.system = self.system.replace(first, domain)

    def post(self, src, out):
        """Постобработка перевода по паре: hu-ru — прописные и единицы (fix_caps, fix_units); перевод на русский —
        единицы; прочие пары — без изменений."""
        if self.hu_ru:
            return fix_units(fix_caps(src, out, self.ru_caps))
        return fix_units(out) if self.tgt == "ru" else out

    def term_issues(self, src, ru):
        """Обязательные термины глоссария комплекта, которых нет в переводе (по основам).
        Если в файле есть пометки «строгий» (плотный глоссарий агентов, тысячи записей) — обязательны только они,
        иначе синонимы и падежи гоняли бы лишние повторы; без пометок обязательны все (прежний режим)."""
        strict = self.strict_entries
        if not strict:
            return []
        miss = gloss_misses(TAG.sub("", src), ru, strict, outer=self.entries)
        ru_of = {e.hu: e.ru for e in strict}
        return [f"термин документа: {m} → {ru_of.get(m, '?')}" for m in miss[:3]]

    def full_check(self, src, ru):
        if self.ru_en:
            return check_en(src, ru) + (self.term_issues(src, ru) if ru else [])
        if not self.hu_ru:
            return check_generic(src, ru, self.src, self.tgt) + (self.term_issues(src, ru) if ru else [])
        return check(src, ru, self.allowed_latin, self.gloss_abbr, self.vocab, frozenset(self.abbrs)) + \
            (self.term_issues(src, ru) if ru else [])

    @staticmethod
    def term_line(e):
        """«SZBV → СУЗ [в глоссарии также: ОР СУЗ; примечание]» — скобки русского варианта уходят в примечание,
        иначе модель вставляет «СУЗ (ОР СУЗ)» в каждое упоминание (замер 01.10)."""
        ru = re.sub(r"\s*\([^)]*\)", "", e.ru).strip()
        paren = re.findall(r"\(([^)]*)\)", e.ru)
        if not ru:
            ru, paren = e.ru, []
        own = re.sub(r"^(строгий|риск высокий)(;\s*)?", "", (e.note or "").strip())
        own = re.sub(r"^(строгий|риск высокий)(;\s*)?", "", own)
        own = re.sub(r"\[уточнить\]|проверено Claude|(?<![\w])строгий(?![\w])", "", own)   # служебное — не модели
        own = re.sub(r"(;\s*){2,}", "; ", own).strip("; ")[:110]
        note = "; ".join(([f"в глоссарии также: {', '.join(paren)}"] if paren else []) + ([own] if own else []))
        return f"- {e.hu} → {ru}" + (f"  [{note}]" if note else "")

    def ru_main(self, abbr):
        """Основной русский вариант аббревиатуры: без скобок, первый из «X / Y»."""
        ru = re.sub(r"\([^)]*\)", "", self.abbrs[abbr].ru)
        return re.split(r"\s+/\s+|;", ru)[0].strip() or self.abbrs[abbr].ru

    def _allowed_latin(self):
        """Строчная латиница, разрешённая в переводе: то, что сам глоссарий оставляет латиницей."""
        ok = set()
        for e in self.entries:
            ok.update(w for w in LATIN_WORD.findall(e.ru))
        return frozenset(ok)

    def _prompt(self, units, first_abbr, context, note=""):
        text = "\n".join(u["text"] for u in units)
        # плотный глоссарий: в пачке бывает >150 терминов. Срез по длине выкидывал однословные сокращения
        # основного глоссария (KR → БРУ-К, AR → БРУ-А) — модель оставляла латиницу (эталон g3/g4, 02.10).
        # Порядок: основной глоссарий заказчика → строгие → остальные подсказки
        found = G.find(text, self.entries, limit=100000)
        main_ids = self.main_ids
        strict = lambda e: "строгий" in (e.note or "")
        hits = ([e for e in found if id(e) in main_ids] + [e for e in found if id(e) not in main_ids and strict(e)] +
                [e for e in found if id(e) not in main_ids and not strict(e)])[:200]
        hits += G.paren_abbr_hits(text, self.entries, found)     # второе значение сокращения: «… (TG) → турбогенератор»
        parts = []
        if hits:
            parts.append(("Glossary terms for this fragment:\n" if not self.hu_ru else "Термины глоссария для этого фрагмента:\n")
                         + "\n".join(self.term_line(e) for e in hits))
        # правило первого употребления «СУЗ (SZBV)» модель внутри пачки не держит (ставит скобки
        # в каждое упоминание, замер 01.10) — его расставляет first_use() по всему документу
        if context:
            parts.append(("Preceding text of the document (for context, do NOT translate):\n" if not self.hu_ru else
                          "Предыдущий текст документа (для контекста, НЕ переводить):\n") + to_model(context))
        if note:
            parts.append(note)
        # разбор слов с отрицающими морфемами (tömör+telen+ség): без него «негерметичность» становилась
        # «герметичностью»/«разрывом»/«закупоркой» (49 из 86 верно); с ним 6 из 6 исправлено (опыт 01.10)
        h = morph.hint(text) if self.src == "hu" else ""      # венгерские морфемы отрицания
        if h:
            parts.append(h)
        segs = "\n".join(f'<s id="{i+1}">{to_model(u["text"])}</s>' for i, u in enumerate(units))
        parts.append(("Translate the segments:\n" if not self.hu_ru else "Переведи сегменты:\n") + segs)
        return "\n\n".join(parts)

    def _call(self, units, first_abbr, context, think, note=""):
        user = self._prompt(units, first_abbr, context, note)
        src_len = sum(len(u["text"]) for u in units)
        r = chat([{"role": "system", "content": self.system}, {"role": "user", "content": user}],
                 max_tokens=int(800 + src_len * 1.6) + (6000 if think else 0), think=think,
                 temperature=getattr(self, "temperature", 0), sampling=getattr(self, "sampling", None))
        got = {}
        text = r.get("text", "")
        # незакрытая метка (обрыв, забытый </s>) — сегмент до следующей метки или конца ответа
        for m in re.finditer(r'<s id="?(\d+)"?>(.*?)(?:</s>|(?=<s id=)|\Z)', text, re.S):
            i = int(m.group(1)) - 1
            if 0 <= i < len(units) and i not in got and m.group(2).strip():
                got[i] = from_model(m.group(2).strip())
        if len(units) == 1 and not got and text.strip() and "<s" not in text:
            got[0] = from_model(text.strip())          # один сегмент, ответ без меток
        return got, r

    def run_batch(self, plan):
        units = plan["units"]
        got, r = self._call(units, plan["first_abbr"], plan["context"], self.think, plan.get("note", ""))
        res = {}
        for i, u in enumerate(units):
            ru = got.get(i, "")
            res[u["key"]] = {"ru": ru, "issues": self.full_check(u["text"], ru), "tries": 1}
        # повтор упавших поштучно, с рассуждением и перечнем замечаний
        for i, u in enumerate(units):
            cur = res[u["key"]]
            for attempt in range(2):
                if not cur["issues"]:
                    break
                note = ("В прошлый раз перевод этого сегмента не прошел проверку: " + "; ".join(cur["issues"]) +
                        ". Исправь. Если в оригинале есть код или число — оно должно быть в переводе без изменений." +
                        (" Проверь, что каждое отрицание оригинала сохранено (см. разбор слов выше)."
                         if any(i.startswith("отрицание") for i in cur["issues"]) else ""))
                if self.ru_en:
                    note = ("The previous translation of this segment failed the check (issues in Russian): "
                            + "; ".join(cur["issues"]) + ". Fix it and answer in English. Every number and code of the "
                            "original must be in the translation unchanged (decimal comma → decimal point).")
                elif not self.hu_ru:
                    import langs as L
                    note = ("The previous translation of this segment failed the check (issues in Russian): "
                            + "; ".join(cur["issues"]) + f". Fix it and answer in {L.NAME_EN[self.tgt]}. Every number "
                            "and code of the original must be in the translation (with the decimal separator of the "
                            "target language).")
                # обозначение «1.GF.H.ÁG.SZEK.FEDÉL» модель каждый раз «расшифровывает» по-своему (48 строк 30YCGT, 28.10):
                # в повторе оно уходит под меткой и возвращается дословно
                lost = any(i.startswith("нет кодов") for i in cur["issues"])
                masked, back = mask_designators(u["text"]) if lost and self.hu_ru else (u["text"], {})
                g2, _ = self._call([{**u, "text": masked}], [a for a in plan["first_abbr"] if a in u["text"]],
                                   plan["context"], self.retry_think, note + ("" if attempt == 0 else " (повторная попытка)"))
                ru2 = unmask_designators(g2.get(0, ""), back)
                iss2 = self.full_check(u["text"], ru2)
                if ru2 and badness(iss2) < badness(cur["issues"]):
                    cur = {"ru": ru2, "issues": iss2, "tries": cur["tries"] + 1}
                else:
                    cur["tries"] += 1
            res[u["key"]] = cur
        return res

    def first_use(self, segs, out):
        """«СУЗ (SZBV)» при первом употреблении в документе, дальше «СУЗ». Только для коротких русских
        аббревиатур (СУЗ, ИЭ, АРМ, ПУБЭ): длинные описательные варианты склоняются, их не трогаем."""
        short = {a: self.ru_main(a) for a in self.abbrs
                 if re.fullmatch(r"[А-ЯЁ][А-ЯЁ0-9\-]{1,7}", self.ru_main(a))}
        seen = set()
        for s in segs:
            r = out.get(s["id"])
            if not r or r.get("copied") or not r["ru"]:
                continue
            ru = r["ru"]
            for a in G.abbr_in(TAG.sub(" ", s["text"]), short):
                rus = short[a]
                pat_dup = re.compile(rf"(?<![А-ЯЁA-Z]){re.escape(rus)}\s*\({re.escape(a)}\)")
                ru = pat_dup.sub(rus, ru)                       # скобки, поставленные моделью, снять
                if a not in seen:
                    m = re.search(rf"(?<![А-ЯЁA-Z]){re.escape(rus)}(?![А-ЯЁA-Z])", ru)
                    if m:
                        ru = ru[:m.end()] + f" ({a})" + ru[m.end():]
                        seen.add(a)
            r["ru"] = ru
