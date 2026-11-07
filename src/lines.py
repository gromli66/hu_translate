# -*- coding: utf-8 -*-
"""Слово неделимо (UNIFIED_DESIGN.md, П3, этап E4): перенос слова в конце строки решается здесь — одной функцией и одним
списком исключений для строк текстового слоя (layout.join_lines, метки портала roles.apply) и слов Tesseract
(segments._dehyphen). Прежде было три реализации с разными исключениями: у слов Tesseract «fűtő-» + «és» склеивалось в
«fűtőés», «KI-» + «KAPCSOLVA» не склеивалось вовсе. Сборка (docx_layout._join) зовёт glue с E9.7г."""
import re

UPPER = "A-ZÁÉÍÓÖŐÚÜŰА-ЯЁ"
HYPHEN = re.compile(r"[^\W\d_]-$")                                  # строка кончается буквой и дефисом (не «–»)
# Исключения — дефис не перенос или часть слова:
MINUS = re.compile(r"(?:^|[^\d\s])\s+-$|^-$")                         # «felé -»: отдельный дефис после слова (не числа)
SUSPENDED = re.compile(r"(?:és|vagy|illetve|ill\.|и|или|либо|and|or)(?:\s|$)", re.I)   # «fűtő- és hűtőrendszer»
KEEP = re.compile(rf"[{UPPER}0-9]-$|\w-\w+-$")      # после прописной или цифры («ÜFK-» + «s»), составное «loss-of-» + «coolant»
CAPS_WRAP = re.compile(rf"(?:^|[^\w-])([{UPPER}]{{2,}})-$")         # слово из двух и более прописных, дефис в конце
CAPS_HEAD = re.compile(rf"([{UPPER}]{{2,}})(?![\w-])")              # продолжение начинается словом из прописных


def hyphen_break(prev, nxt):
    """Слово разорвано переносом строки: строка кончается буквой и дефисом (не «–»), следующая начинается буквой.
    Обе строки — одна фраза, что бы ни стояло между ними («oltó-» / «berendezés» в блоке схемы, «KI-» / «KAPCSOLVA»,
    «pa-» / «lacknyomás»): иначе слово переводилось двумя кусками."""
    return bool(HYPHEN.search(prev.rstrip())) and nxt.lstrip()[:1].isalpha()


# служебное слово в конце строки — фраза на нём не кончается: артикль, союз (языки пар проекта: венгерский, русский,
# английский; строчными — прописное «A» в конце строки бывает и обозначением)
OPEN_TAIL = re.compile(r"(?:^|\s)(?:a|az|egy|és|vagy|illetve|ill\.|the|an|and|or|of|и|или|а|в|на|с|по|к|для)$")


def mid_phrase(prev, nxt):
    """Строка оборвалась на середине фразы: в ней ≥ 3 слов, она кончается буквой или запятой, следующая начинается
    строчной буквой. Одиночные значения столбика («rendben» ×4) и пункты с точкой не подходят.
    Продолжение с числа или кода — когда строка кончается запятой или служебным словом: «…háromutú szelepet a» /
    «30UJ97B001 felé NYITNI.» (3SZ17 л. 9), «…7 bar-os gőzrendszer (RX),» / «5 bar-os gőzrendszer …» (1SZ01 л. 8) шли
    порознь. Строка, кончающаяся словом или точкой, перед кодом — нет: там код обычно новая строка перечня."""
    prev, nxt = prev.rstrip(), nxt.lstrip()
    if len(prev.split()) < 3:
        return False
    if re.search(r"[^\W\d_]$|,$", prev) and nxt[:1].isalpha() and nxt[:1].islower():
        return True
    first = nxt.split()[0] if nxt else ""
    return bool(re.search(r"\d", first)) and (prev.endswith(",") or bool(OPEN_TAIL.search(prev)))


def caps_wrap(out, ln):
    """«KI-» + «KAPCSOLVA», «REND-» + «BEN.»: слово из прописных, разорванное переносом. Обе части от двух букв и одна из
    них не короче четырёх; пара коротких аббревиатур («RO-ROP», «ÜBO-ÜBO», «RP-UL») — составная, дефис остаётся."""
    a, b = CAPS_WRAP.search(out), CAPS_HEAD.match(ln)
    return bool(a and b) and (len(a.group(1)) >= 4 or len(b.group(1)) >= 4)


def glue(prev, nxt, ranges=True):
    """Перенос слова: prev кончается «-», nxt — продолжение того же слова → склеенный текст без пробела (дефис уходит,
    если он не часть слова); не перенос — None. Продолжение со строчной буквы — перенос, кроме союза после дефиса;
    с прописной — только слово из прописных (caps_wrap); с цифры после «цифра-» — диапазон («1000-» + «1200» →
    «1000-1200», дефис на месте; прежде «1000- 1200», а копия правил в сборке, docx_layout._join, давала «1000-1200», E9.7г).
    Отдельный «-» после слова (MINUS) + число — знак минуса: «felé -» + «1,2 mm» → «felé -1,2 mm» (3SZ19; прежде «- 1,2»,
    тире) — тире бывает только между словами; «20 -» + «25 bar» — диапазон через пробелы, не минус.
    ranges=False — без диапазонов и минуса: у слов Tesseract «118-» в конце строки и номер штампа «10135772» в начале следующей —
    разные числа; склеенное «118-10135772» не сверялось с чтением портала и терялось (титулы-сканы, 20 JSON)."""
    if not prev.endswith("-") or SUSPENDED.match(nxt):
        return None
    if ranges and nxt[:1].isdigit() and (prev[-2:-1].isdigit() or MINUS.search(prev)):
        return prev + nxt
    if nxt[:1].islower():
        return (prev if KEEP.search(prev) else prev[:-1]) + nxt
    if caps_wrap(prev, nxt):
        return prev[:-1] + nxt
    return None


def join_lines(lines):
    """Строки абзаца → текст: через пробел, перенос слова — без пробела (glue)."""
    out = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        out = (glue(out, ln) or out + " " + ln) if out else ln
    return out
