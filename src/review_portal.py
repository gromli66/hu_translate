# -*- coding: utf-8 -*-
"""Агент-контролёр перевода на портале: ищет СМЫСЛОВЫЕ ошибки, которые проходят формальные проверки
(не тот термин, додуманное условие, потерянное действие). Не «оцени перевод», а разбор по элементам:
оригинал → элементы (оборудование, действия, условия, ограничения, числа, отрицания) → каждый в переводе;
затем перевод → что добавлено. Простой проверяющий отвергнут 01.10 (3 из 7, выдумывал правки).
Насыщенный режим (ctx): связный текст вокруг, второй перевод, свод соглашений комплекта, разбор отрицаний.
Перед работой — configure(terms, conventions, glossary). evaluate() — замер проверяющего на размеченном наборе."""
import re, sys, json
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
sys.stdout.reconfigure(encoding="utf-8")
from llm import chat, STAT
from translate import Translator
import glossary as G
import morph


SYSTEM = """Ты — контролёр перевода эксплуатационной документации АЭС с ВВЭР-440 (АЭС «Пакш») с венгерского на русский.
Ищешь только ошибки СМЫСЛА, из-за которых специалист сделает не то: не тот термин или объект, потерянное или
добавленное условие/действие/объект, обратный смысл (отрицание), неверное число, единица, срок, ограничение.
НЕ ошибки: стиль, порядок слов, падеж, синоним с тем же смыслом, принятые русские сокращения (БРУ-К, КД, САОЗ, СУЗ,
ПГ, ГЦН, ПЭН, ВИУР и т. п.), коды оборудования латиницей, оформление номеров и дат, огрехи распознавания в оригинале."""

USER = """Оригинал (HU):
{hu}

Перевод (RU):
{ru}
{terms}
Работай по шагам.
1. Выпиши из оригинала смысловые элементы: оборудование/системы, действия, условия (ha, amennyiben, esetén…),
ограничения (legalább, max., csak, legfeljebb…), числа с единицами, отрицания (nem, sem, tilos, -talan/-telen…).
2. Для каждого элемента найди соответствие в переводе: OK / ИСКАЖЕНО / НЕТ.
3. Найди в переводе то, чего нет в оригинале (добавленные объекты, состояния, условия, уточнения): ДОБАВЛЕНО.
4. Итог: серьёзная ошибка есть, только если элемент ИСКАЖЕН или ПОТЕРЯН, или ДОБАВЛЕНО то, что меняет действие
специалиста. Если сомневаешься — серьёзной ошибки нет.

Ответ — только JSON:
{{"elements": [{{"hu": "...", "ru": "...", "v": "OK|ИСКАЖЕНО|НЕТ"}}], "added": ["..."], "serious": true|false, "reason": "кратко, по-русски"}}"""

# Насыщенный режим (02.10): то, чем пользовались слепые проверяющие Claude и чего не было у агента портала —
# связный текст вокруг, второй перевод для сравнения, свод принятых соглашений комплекта, разбор отрицаний,
# больше терминов, пересказ смысла по-английски (понимание оригинала, не зависящее от проверяемого русского).
CONVENTIONS = ""                                 # свод принятых соответствий комплекта — configure()
TERMS, GLOSSARY = None, None
SYSTEM_RICH_HEAD = SYSTEM + """

Принятые в комплекте соответствия (образец; другое написание того же смысла — не ошибка):
"""
SYSTEM_RICH_TAIL = """

Венгерские ловушки смысла:
- «лишнее» отрицание в придаточном после kizár, megakadályoz, tilt, óv, véd, tart (attól), kétség: «nem zárja ki, hogy … ne
  tudjon» = «не исключает того, что … сможет» — отрицание «ne» здесь НЕ переводится; перевод с двумя «не» переворачивает смысл;
- морфемы -talan/-telen/-tlan/-tlen и in- — отрицание («без», «не-»); nem/sem + слово — отрицание этого слова;
- -hat/-het — «можно/допускается», kell — «необходимо», tilos — «запрещается», nem kell — «не требуется»;
- контур определяется оборудованием: ПГ со стороны продувки/питательной воды/пара — второй контур; ГЦН, КД, реактор — первый."""
SYSTEM_RICH = SYSTEM_RICH_HEAD + CONVENTIONS + SYSTEM_RICH_TAIL

USER_RICH = """Текст перед фрагментом (оригинал → принятый перевод):
{prev}

ПРОВЕРЯЕМЫЙ фрагмент.
Оригинал (HU):
{hu}

Перевод (RU):
{ru}

Второй, независимый перевод того же фрагмента (может содержать свои ошибки; служит для сравнения):
{other}

Текст после фрагмента:
{next}
{terms}{morph}
Работай по шагам.
0. Кратко перескажи смысл оригинала по-английски (1–2 фразы) — до того, как смотреть на русский.
1. Выпиши из оригинала смысловые элементы: оборудование/системы, действия, условия, ограничения, числа с единицами, отрицания.
2. Для каждого элемента найди соответствие в проверяемом переводе: OK / ИСКАЖЕНО / НЕТ.
3. Найди в проверяемом переводе то, чего нет в оригинале: ДОБАВЛЕНО.
4. Где два перевода расходятся по смыслу — реши по оригиналу и контексту, какой верен.
5. Итог: серьёзная ошибка есть, только если элемент ИСКАЖЕН или ПОТЕРЯН, или ДОБАВЛЕНО то, что меняет действие
специалиста. Расхождение двух переводов само по себе ошибкой не является. Если сомневаешься — серьёзной ошибки нет.

Ответ — только JSON:
{{"en": "...", "elements": [{{"hu": "...", "ru": "...", "v": "OK|ИСКАЖЕНО|НЕТ"}}], "added": ["..."], "serious": true|false, "reason": "кратко, по-русски"}}"""

USER_COMPACT = USER_RICH.replace(
    """Ответ — только JSON:
{{"en": "...", "elements": [{{"hu": "...", "ru": "...", "v": "OK|ИСКАЖЕНО|НЕТ"}}], "added": ["..."], "serious": true|false, "reason": "кратко, по-русски"}}""",
    """Шаги 0–4 выполни в рассуждении; в ответ выпиши только найденные проблемы.
Ответ — только JSON:
{{"problems": ["элемент: что не так"], "serious": true|false, "reason": "кратко, по-русски", "fix": "исправленный перевод всего фрагмента, если serious = true; иначе пусто"}}""")

_TR = None


def configure(terms=None, conventions=None, glossary=None):
    """Глоссарий комплекта, основной глоссарий и свод соглашений (раздел «## Уже принято» или весь файл)."""
    global CONVENTIONS, SYSTEM_RICH, TERMS, GLOSSARY, _TR
    TERMS, GLOSSARY, _TR = terms, glossary, None
    if conventions and Path(conventions).exists():
        t = Path(conventions).read_text(encoding="utf-8")
        CONVENTIONS = t.split("## Уже принято", 1)[1].split("\n## ", 1)[0].strip() if "## Уже принято" in t else t.strip()
    SYSTEM_RICH = SYSTEM_RICH_HEAD + CONVENTIONS + SYSTEM_RICH_TAIL


def _terms(hu, limit=30):
    global _TR
    if _TR is None:
        _TR = Translator(extra=TERMS, glossary=GLOSSARY)
    found = G.find(hu, _TR.entries, limit=100000)
    main = [e for e in found if id(e) in _TR.main_ids]
    rest = [e for e in found if id(e) not in _TR.main_ids]
    lines = [_TR.term_line(e) for e in (main + rest)[:limit] + G.paren_abbr_hits(hu, _TR.entries, found)]
    return ("\nПринятые термины (глоссарий):\n" + "\n".join(lines) + "\n") if lines else ""


def _ctx(pairs):
    return "\n".join(f"- {h[:400]} → {r[:400]}" for h, r in pairs) or "(нет)"


def review(hu, ru, vote=0, think=True, ctx=None, compact=False):
    """ctx (насыщенный режим): {"prev": [(hu, ru)], "next": [(hu, ru)], "other": второй перевод}.
    compact: в ответ только проблемы, вердикт и исправленный перевод (fix) — короче вывод, без отдельного автоисправления."""
    if ctx is None:
        sysm, user = SYSTEM, USER.format(hu=hu, ru=ru, terms=_terms(hu))
    else:
        other, mh = ctx.get("other") or "", morph.hint(hu)
        sysm = SYSTEM_RICH
        user = (USER_COMPACT if compact else USER_RICH).format(prev=_ctx(ctx.get("prev", [])), hu=hu, ru=ru, next=_ctx(ctx.get("next", [])),
                                other=("(совпадает с проверяемым)" if other == ru else other) or "(нет)",
                                terms=_terms(hu, 50), morph=("\n" + mh + "\n") if mh else "")
    if vote:
        sysm += f"\n(проверка {vote + 1})"
    r = chat([{"role": "system", "content": sysm}, {"role": "user", "content": user}],
             max_tokens=12000 if think else 4000, think=think, temperature=0.6 if vote else 0, tag="review")
    txt = r.get("text", "")
    m = re.search(r"\{.*\}", txt, re.S)
    try:
        d = json.loads(m.group(0)) if m else {}
    except json.JSONDecodeError:
        # венгерские кавычки „…" внутри строк модель не экранирует (1,5 % ответов, 04.10) — вердикт и причина целы,
        # достаём их без разбора всего JSON
        sv = re.findall(r'"serious"\s*:\s*(true|false)', txt)
        rs = re.search(r'"reason"\s*:\s*"(.*)"\s*\}\s*$', txt.strip(), re.S)
        d = {"serious": bool(sv) and sv[-1] == "true", "reason": (rs.group(1) if rs else "ответ не разобран")[:400]}
    bad = [e for e in d.get("elements", []) if isinstance(e, dict) and e.get("v") in ("ИСКАЖЕНО", "НЕТ")]
    if not isinstance(d, dict):
        d = {}
    fx = d.get("fix") if isinstance(d.get("fix"), str) else ""
    if not fx and '"fix"' in txt:
        mfx = re.search(r'"fix"\s*:\s*"(.*)"\s*\}\s*$', txt.strip(), re.S)
        fx = mfx.group(1) if mfx else ""
    return {"serious": bool(d.get("serious")), "reason": d.get("reason", ""), "bad": bad, "added": d.get("added", []),
            "fix": fx.strip(), "problems": d.get("problems", [])}


def evaluate(path, votes, workers, think=True, rich=False, compact=False):
    items = json.loads(Path(path).read_text(encoding="utf-8"))
    jobs = [(i, v) for i in range(len(items)) for v in range(votes)]
    with ThreadPoolExecutor(workers) as ex:
        res = list(ex.map(lambda j: (j, review(items[j[0]]["hu"], items[j[0]]["ru"], j[1], think,
                                               items[j[0]] if rich else None, compact)), jobs))
    by = {}
    for (i, v), r in res:
        by.setdefault(i, []).append(r)
    out = []
    for i, it in enumerate(items):
        rs = by[i]
        out.append({**it, "flags": sum(r["serious"] for r in rs), "votes": len(rs), "reasons": [r["reason"] for r in rs if r["serious"]]})
    name = Path(path).stem + f"_portal_v{votes}{'' if think else '_fast'}{'_rich' if rich else ''}{'_compact' if compact else ''}.json"
    Path(path).with_name(name).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    for need in range(1, votes + 1):
        f = lambda lab: [x for x in out if x["label"] == lab]
        hit = lambda xs: sum(x["flags"] >= need for x in xs)
        s2, s1, s0 = f(2), f(1), f(0)
        print(f"порог {need}/{votes}: серьёзных найдено {hit(s2)}/{len(s2)}; отмечено мелких {hit(s1)}/{len(s1)}, "
              f"верных {hit(s0)}/{len(s0)}; всего отмечено {hit(out)}/{len(out)} ({hit(out) / len(out):.0%})")
    print("вызовов", STAT)
    for x in out:
        if x["label"] == 2:
            print(("✓" if x["flags"] else "✗"), x["src"], x["ru"][:90], "|", (x["reasons"] or [""])[0][:120])
    for x in out:
        if x["label"] == 0 and x["flags"]:
            print("ложная тревога:", x["src"], x["ru"][:80], "|", (x["reasons"] or [""])[0][:150])
