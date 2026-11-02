# -*- coding: utf-8 -*-
"""Слово неделимо (UNIFIED_DESIGN.md, П3, этап E4): перенос слова в конце строки решается здесь — одной функцией и одним
списком исключений для строк текстового слоя (layout.join_lines, метки портала roles.apply) и слов Tesseract
(segments._dehyphen). Прежде было три реализации с разными исключениями: у слов Tesseract «fűtő-» + «és» склеивалось в
«fűtőés», «KI-» + «KAPCSOLVA» не склеивалось вовсе. Сборка (docx_layout._join) пока держит свою копию тех же правил."""
import re

UPPER = "A-ZÁÉÍÓÖŐÚÜŰА-ЯЁ"
HYPHEN = re.compile(r"[^\W\d_]-$")                                  # строка кончается буквой и дефисом (не «–»)
# Исключения — дефис не перенос или часть слова:
SUSPENDED = re.compile(r"(?:és|vagy|illetve|ill\.|и|или|либо|and|or)(?:\s|$)", re.I)   # «fűtő- és hűtőrendszer»
KEEP = re.compile(rf"[{UPPER}0-9]-$|\w-\w+-$")      # после прописной или цифры («ÜFK-» + «s»), составное «loss-of-» + «coolant»
CAPS_WRAP = re.compile(rf"(?:^|[^\w-])([{UPPER}]{{2,}})-$")         # слово из двух и более прописных, дефис в конце
CAPS_HEAD = re.compile(rf"([{UPPER}]{{2,}})(?![\w-])")              # продолжение начинается словом из прописных


def hyphen_break(prev, nxt):
    """Слово разорвано переносом строки: строка кончается буквой и дефисом (не «–»), следующая начинается буквой.
    Обе строки — одна фраза, что бы ни стояло между ними («oltó-» / «berendezés» в блоке схемы, «KI-» / «KAPCSOLVA»,
    «pa-» / «lacknyomás»): иначе слово переводилось двумя кусками."""
    return bool(HYPHEN.search(prev.rstrip())) and nxt.lstrip()[:1].isalpha()


def mid_phrase(prev, nxt):
    """Строка оборвалась на середине фразы: в ней ≥ 3 слов, она кончается буквой или запятой, следующая начинается
    строчной буквой. Одиночные значения столбика («rendben» ×4) и пункты с точкой не подходят."""
    prev, nxt = prev.rstrip(), nxt.lstrip()
    return len(prev.split()) >= 3 and bool(re.search(r"[^\W\d_]$|,$", prev)) and nxt[:1].isalpha() and nxt[:1].islower()


def caps_wrap(out, ln):
    """«KI-» + «KAPCSOLVA», «REND-» + «BEN.»: слово из прописных, разорванное переносом. Обе части от двух букв и одна из
    них не короче четырёх; пара коротких аббревиатур («RO-ROP», «ÜBO-ÜBO», «RP-UL») — составная, дефис остаётся."""
    a, b = CAPS_WRAP.search(out), CAPS_HEAD.match(ln)
    return bool(a and b) and (len(a.group(1)) >= 4 or len(b.group(1)) >= 4)


def glue(prev, nxt):
    """Перенос слова: prev кончается «-», nxt — продолжение того же слова → склеенный текст без пробела (дефис уходит,
    если он не часть слова); не перенос — None. Продолжение со строчной буквы — перенос, кроме союза после дефиса;
    с прописной — только слово из прописных (caps_wrap)."""
    if not prev.endswith("-") or SUSPENDED.match(nxt):
        return None
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
