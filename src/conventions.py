# -*- coding: utf-8 -*-
"""Языковые соглашения комплекта (projects/<имя>/data/conventions.json) и словарь документа (<документ>_terms.json):
проверка пары «венгерский → русский» и правка кодом всего детерминированного. Утверждено Максимом 15.10: ссылки на
разделы — «(см. раздел N)», на рисунки — «рисунок N» в падеже, даты — как в оригинале.

Правила (номера — как в отчёте МЕТРИКА_struct_docx.md):
  Я0 HTML-сущности (&lt; &gt;) в тексте        Я6 коды, пути, имена файлов — дословно
  Я2 «рис. N» → «рисунок N» в падеже            Я7 сокращения комплекта (TU → ИИ …)
  Я3 «(см. раздел N)», «раздел N» вместо главы   Я8 одна форма предупреждения «ВНИМАНИЕ!»
  Я4 кавычки «ёлочки»                            Я9 числовые даты — как в оригинале
  Я5 названия — латиницей (UTS, Ethernet …)      Я10 словарь документа
                                                 Я11 венгерские слова в переводе
check() → [(правило, что не так, можно ли кодом)], fix_code() → (новый перевод, [правила])."""
import html
import json
import re
import unicodedata
from pathlib import Path

TAGS = re.compile(r"</?f\d+>")
CYR = "А-Яа-яЁё"
CAPTION_HU = re.compile(r"^\s*\d+\.\s*(ábra|táblázat)\b", re.I)
# код: имя файла, путь, команда, идентификатор — с «_», цифрой, слэшем в начале или расширением
CODE = re.compile(r"(?<![\w/\\])(?:[/\\][A-Za-z0-9_.\-/\\]+|[A-Za-z0-9][A-Za-z0-9_\-]*(?:[._/\\][A-Za-z0-9_\-]+)+)(?![\w])")
DATE = re.compile(r"\b(\d{4})\.\s?(\d{2})\.\s?(\d{2})\.?")
NAMES = {"Я0": "HTML-сущность в тексте", "Я2": "ссылка на рисунок/таблицу", "Я3": "ссылка на раздел",
         "Я4": "кавычки", "Я5": "название латиницей", "Я6": "код дословно", "Я7": "сокращение комплекта",
         "Я8": "предупреждение", "Я9": "дата как в оригинале", "Я10": "термин документа",
         "Я11": "венгерский в переводе"}


def load(path, terms=None):
    """Соглашения комплекта + словарь документа (необяз.)."""
    rules = json.loads(Path(path).read_text(encoding="utf-8")) if path else {}
    rules["terms"] = json.loads(Path(terms).read_text(encoding="utf-8"))["terms"] if terms and Path(terms).exists() else []
    return rules


def _codes(hu):
    out = []
    for c in set(CODE.findall(hu)):
        c = c.rstrip(".")
        if not re.search(r"[A-Za-z]", c):
            continue
        # «5.sz» (5. szám), «10.3Nem», «a.BER» — номер, слипшийся с текстом, а не код
        if re.fullmatch(r"\d+\.[a-záéíóöőúüű]{1,4}", c) or re.match(r"\d+(?:\.\d+)+[A-Z][a-z]", c) \
                or re.match(r"(?:\d{1,2}|[A-Za-z])\.[A-Za-z]", c):
            continue
        if "_" in c or re.search(r"\d", c) or c[0] in "/\\" or re.fullmatch(r"[\w\-]+\.[A-Za-z]{2,4}", c):
            out.append(c)
    return out


def _cap(s, upper):
    return (s[:1].upper() + s[1:]) if upper else s


# ---------------------------------------------------------------- ссылки (Я2, Я3)
CASE_GLAVA = {"глава": "раздел", "главы": "раздела", "главе": "разделе", "главу": "раздел", "главой": "разделом",
              "главах": "разделах", "главам": "разделам", "главами": "разделами"}
FIG_PREP = {"на": "рисунке", "в": "рисунке", "к": "рисунку", "по": "рисунку", "с": "рисунком", "под": "рисунком",
            "из": "рисунка", "до": "рисунка"}
NUM = r"\d+(?:\.\d+)*"


def _refs_fix(ru):
    t = ru
    # Я3: «(см.: 5.5)», «см. 7.2 и 7.3» → «(см. раздел 5.5)», «см. разделы 7.2 и 7.3»
    t = re.sub(rf"([Сс]м)\.:?\s*({NUM})(\s*(?:и|,)\s*)({NUM})(?!\s*\])",
               lambda m: f"{m[1]}. разделы {m[2]}{m[3]}{m[4]}", t)
    t = re.sub(rf"([Сс]м)\.:?\s*({NUM})(?![\d\]])", lambda m: f"{m[1]}. раздел {m[2]}", t)
    t = re.sub(rf"([Сс]м)\.:\s*(раздел|рисун|\[)", r"\1. \2", t)
    # «(глава 4)», «(раздел 4)» → «(см. раздел 4)»
    t = re.sub(rf"\((?:[Гг]лава|[Рр]аздел)\s+({NUM})\)", r"(см. раздел \1)", t)
    # «глава/главе … N» → «раздел/разделе … N»
    t = re.sub(rf"(?<![{CYR}])([Гг])(лава|лавы|лаве|лаву|лавой|лавах|лавам|лавами)(\s+{NUM})",
               lambda m: _cap(CASE_GLAVA["г" + m[2]], m[1] == "Г") + m[3], t)
    # «см. п. 5.2» → «см. раздел 5.2»
    t = re.sub(rf"([Сс]м)\.\s*п\.\s*({NUM}\.\d+)", r"\1. раздел \2", t)
    # Я2: «рис. N» → «рисунок N» в падеже по предлогу; «см. рис. N» → «см. рисунок N»
    t = re.sub(rf"(?<![{CYR}])(на|в|к|по|с|под|из|до|На|В|К|По|С|Под|Из|До)\s+рис\.\s*(\d+)",
               lambda m: f"{m[1]} {FIG_PREP[m[1].lower()]} {m[2]}", t)
    t = re.sub(r"([Сс]м\.:?\s*|\()рис\.\s*(\d+)", lambda m: m[1].replace(":", "") + "рисунок " + m[2], t)
    # «Рисунок 33» посреди фразы — со строчной
    t = re.sub(rf"(?<=[{CYR},;:)»\w]\s)(Рисун(?:ок|ке|ка|ку|ком))(\s+\d+)", lambda m: m[1].lower() + m[2], t)
    return t


def _refs_left(ru):
    """Что кодом не правится: «рис. N» без предлога, «N. рис.» (порядок оригинала), «таблица» через сокращение."""
    return re.findall(rf"(?<![{CYR}])рис\.\s*\d+|\b\d+\.\s*(?:рис|табл)\.|(?<![{CYR}])табл\.\s*\d+", ru)


# ---------------------------------------------------------------- кавычки (Я4)
def _quotes_fix(hu, ru):
    """Прямые "…" (если в источнике их нет — значит, не код) и венгерские „…” вне «ёлочек» → «…»."""
    t = ru
    out, pos = [], 0
    for m in re.finditer(r'(?<![=\w])„([^„“”"«»\n]+)[“”"]', t):       # „x001" — открывающая венгерская
        depth = t[:m.start()].count("«") - t[:m.start()].count("»")
        out.append(t[pos:m.start()] + (m[0] if depth > 0 else "«" + m[1] + "»"))
        pos = m.end()
    t = "".join(out) + t[pos:]
    if '"' in t and '"' not in hu:                           # прямые; после «=» — значение в настройке, не трогать
        t = re.sub(r'(?<![=\w])"([^"\n]+)"', r"«\1»", t)
    return t


def check_text(ru, rules):
    """Проверка готового текста DOCX без источника: что видно по самому русскому тексту."""
    t = TAGS.sub("", ru)
    out = []
    if re.search(r"&(?:lt|gt|amp|quot|#\d+);", t):
        out.append(("Я0", "HTML-сущность"))
    fixed = _refs_fix(t)
    if fixed != t or _refs_left(fixed) or re.search(rf"(?<![{CYR}])[Гг]лав\w*\s+{NUM}", t):
        out.append(("Я2/Я3", "формат ссылки"))
    if re.search(r'"[^"\n]+"', t) or re.search(r"„[^“”«»„\n]+[“”]", t) and t.count("«") == 0:
        out.append(("Я4", "кавычки"))
    for k in rules.get("cyrillic_to_latin", {}):
        if re.search(rf"(?<![{CYR}]){re.escape(k)}", t):
            out.append(("Я5", f"«{k}» — латиницей"))
    if re.match(r"\s*(ВНИМАНИЕ\s*[:.]|Внимание\s*[!:.])", t):
        out.append(("Я8", "«ВНИМАНИЕ!»"))
    if _hungarian(t):
        out.append(("Я11", "венгерский остался: " + ", ".join(_hungarian(t)[:5])))
    return out


# ---------------------------------------------------------------- проверка и правка
def check(hu, ru, rules, kind="text"):
    """Нарушения в переводе: [(правило, пояснение, кодом_ли)]. kind: text | caption | heading."""
    hu_t, ru_t = TAGS.sub("", hu), TAGS.sub("", ru)
    out = []
    if re.search(r"&(?:lt|gt|amp|quot|#\d+);", ru_t):
        out.append(("Я0", "HTML-сущность в тексте", True))
    if kind == "text" and not CAPTION_HU.match(hu_t):
        if _refs_fix(ru_t) != ru_t:
            out.append(("Я2/Я3", "формат ссылки", True))
        left = _refs_left(_refs_fix(ru_t))
        if left:
            out.append(("Я2", "ссылка " + ", ".join(left[:3]) + " — «рисунок N» в падеже", False))
        if re.search(rf"(?<![{CYR}])[Гг]лав\w*\s+{NUM}", _refs_fix(ru_t)):
            out.append(("Я3", "«глава N» — «раздел N»", False))
    if _quotes_fix(hu_t, ru_t) != ru_t:
        out.append(("Я4", "кавычки «ёлочки»", True))
    for name in rules.get("latin", []):
        n_hu = _name_count(name, hu_t)
        if n_hu and _name_count(name, ru_t) < n_hu:
            cyr = [k for k, v in rules.get("cyrillic_to_latin", {}).items() if v.lower().startswith(name.lower())
                   and re.search(rf"(?<![{CYR}]){re.escape(k)}", ru_t)]
            out.append(("Я5", f"«{name}» — латиницей" + (f", в оригинале {n_hu} раз" if n_hu > 1 else ""), bool(cyr)))
    for c in _codes(hu_t):
        if c not in ru_t:
            out.append(("Я6", f"«{c}» — дословно", _code_guess(c, ru_t) is not None))
    for lat, cyr in rules.get("abbr", {}).items():
        if re.search(rf"(?<![\w]){re.escape(lat)}(?![\w])", hu_t) and re.search(rf"(?<![\w]){re.escape(lat)}(?![\w])", ru_t) \
                and cyr not in ru_t:
            out.append(("Я7", f"«{lat}» — «{cyr}»", True))
    w = rules.get("warning") or {}
    if w and re.match(rf"\s*{w['hu']}", hu_t, re.I) and not ru_t.lstrip().startswith(w["ru"]):
        out.append(("Я8", f"«{w['ru']}»", bool(re.match(r"\s*(ВНИМАНИЕ|Внимание)\s*[:!.]?", ru_t))))
    for m in DATE.finditer(hu_t):
        if m[0].rstrip(".") not in ru_t:
            out.append(("Я9", f"дата «{m[0]}» — как в оригинале", f"{m[3]}.{m[2]}.{m[1]}" in ru_t))
    for e in rules.get("terms", []):
        m = re.search(e["bad"], ru_t)
        if m and re.search(e["hu"], hu_t) and not (e.get("unless_hu") and re.search(e["unless_hu"], hu_t)):
            code = bool(e.get("swap")) or bool(e.get("drop") and re.search(e["drop"], ru_t))
            out.append(("Я10", f"«{m[0]}» → {e['ru']}", code))
    # «НСБ (BLÜGY)» — правило первого употребления сокращения (translate.first_use), не остаток венгерского
    left = [w for w in _hungarian(ru_t) if not (w.isupper() and f"({w})" in ru_t and w in hu_t)]
    if left:
        out.append(("Я11", "венгерский остался: " + ", ".join(left[:5]), False))
    return out


# путь, имя файла — по символам пути (шаблон «<UTS környezet>/<év>» — текст в скобках, не путь)
CODE_LIKE = re.compile(r"[\w.\-]*[/\\_][\w./\\\-]*|[\w\-]+\.[a-z]{2,4}(?!\w)")


def _name_count(name, t):
    """Сколько раз название стоит в тексте отдельным словом (UTS-t — тоже); пути и имена файлов не в счёт. Название
    латиницей хоть раз — мало: «…UTS… при запуске СОИ другой среды» и «uts_kmod.conf» проходили проверку (UTS, 16.10)."""
    return len(re.findall(rf"(?<![\w]){re.escape(name)}(?![\w])", CODE_LIKE.sub(" ", t), re.I))


def _hungarian(ru):
    """Венгерские слова в переводе: с венгерскими буквами, строчные или ПРОПИСНЫЕ (имя с прописной — название фирмы,
    «Paksi Atomerőmű Zrt.», — латиницей законно); формула Word — математическими буквами, поэтому через NFKC."""
    t = unicodedata.normalize("NFKC", ru)
    return [w for w in re.findall(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű_]{2,}", t)
            if re.search(r"[áéíóöőúüűÁÉÍÓÖŐÚÜŰ]", w) and (w.islower() or w.isupper() or "_" in w)]


def _code_guess(c, ru):
    """Испорченный код в переводе: тот же префикс до последнего разделителя («UTS_A.LIS» → «UTS_A.СПИС»)."""
    cut = max(c.rfind("."), c.rfind("_"), c.rfind("/"))
    head = c[:cut + 1] if cut > 0 else c[:4]
    if len(head) < 3:
        return None
    m = re.search(rf"(?<![\w]){re.escape(head)}[\w{CYR}\-]*", ru)
    return m if m and m[0] != c else None


def fix_code(hu, ru, rules, kind="text"):
    """Правка кодом всего детерминированного. Возвращает (перевод, [правила, по которым правили])."""
    hu_t = TAGS.sub("", hu)
    t, done = ru, []

    def step(rule, new):
        nonlocal t
        if new != t:
            t = new
            done.append(rule)
    step("Я0", html.unescape(t) if re.search(r"&(?:lt|gt|amp|quot|#\d+);", t) else t)
    if kind == "text" and not CAPTION_HU.match(hu_t):
        step("Я2/Я3", _refs_fix(t))
    step("Я4", _quotes_fix(hu_t, t))
    new = t
    for k in sorted(rules.get("cyrillic_to_latin", {}), key=len, reverse=True):
        v = rules["cyrillic_to_latin"][k]
        if re.search(rf"(?<![\w]){re.escape(v.split()[0])}", hu_t, re.I):
            new = re.sub(rf"(?<![{CYR}]){re.escape(k)}[{CYR}]*", v, new)
    step("Я5", new)
    new = t
    for c in _codes(hu_t):
        if c not in new:
            m = _code_guess(c, new)
            if m:
                new = new[:m.start()] + c + new[m.end():]
    step("Я6", new)
    new = t
    for lat, cyr in rules.get("abbr", {}).items():
        if re.search(rf"(?<![\w]){re.escape(lat)}(?![\w])", hu_t) and cyr not in new:
            new = re.sub(rf"(?<![\w]){re.escape(lat)}(?![\w])", cyr, new)
    step("Я7", new)
    w = rules.get("warning") or {}
    if w and re.match(rf"\s*{w['hu']}", hu_t, re.I):
        m = re.match(r"(\s*(?:<f\d+>)?\s*)(ВНИМАНИЕ|Внимание)\s*[:!.]?\s*", t)
        if m and not t[m.end(1):].startswith(w["ru"]):
            # за словом — метки начертания («<f1>Внимание!</f1><f2> текст»): пробел и прописная — у текста за ними,
            # иначе в DOCX «ВНИМАНИЕ!  текст» с двумя пробелами и строчной
            mt = re.match(r"((?:\s*</?f\d+>)*)\s*(.*)", t[m.end():], re.S)
            step("Я8", m[1] + w["ru"] + mt[1] + ((" " + _cap(mt[2], True)) if mt[2].strip() else mt[2]))
    new = t
    for m in DATE.finditer(hu_t):
        new = new.replace(f"{m[3]}.{m[2]}.{m[1]}", m[0].rstrip("."))
    step("Я9", new)
    new = t
    for e in rules.get("terms", []):
        if not re.search(e["hu"], hu_t) or (e.get("unless_hu") and re.search(e["unless_hu"], hu_t)):
            continue
        if e.get("swap"):
            new = re.sub(e["bad"], lambda m: _cap(e["swap"] + m[0][0].lower() + m[0][1:], m[0][0].isupper()), new)
        if e.get("drop"):
            new = re.sub(rf"(^|[.!?]\s+)?{e['drop']}(\w)",
                         lambda m: (m[1] or "") + (m[2].upper() if m[1] is not None else m[2]), new)
    step("Я10", new)
    return t, done
