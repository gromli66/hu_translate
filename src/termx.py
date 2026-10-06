# -*- coding: utf-8 -*-
"""Предпроход терминов документа: частые 1–3-словные сочетания вне глоссария → перевод порталом
в контексте → дополнительный глоссарий md (тот же формат таблиц, пометка «авто»).
Зачем: глоссарий собран под другие комплекты; лексика больших документов (üzemzavar, leterhelés,
kiütni) без общего списка переводится в разных пачках по-разному. Список правит человек до прогона.
Вызывается из pipeline.terms (команда `translate_komplekt.py terms`)."""
import re, json
from collections import Counter, defaultdict
from pathlib import Path

import glossary as G
from llm import chat

STOP = set("""a az egy és vagy hogy nem kell meg ki be el fel le szerint után alatt között esetén illetve valamint
amennyiben akkor ha mint csak már még is ez azt ezt abban amely amelyek amelyet ami amit azaz pedig de sem
minden mindkét több kevesebb nagyobb kisebb legalább legfeljebb kb pl stb ill nincs van vannak volt lesz lehet
való levő lévő esetben eset alapján részére során mellett nélkül miatt keresztül szerinti tartozó vonatkozó
történő követően előtt közben után ilyen olyan ennek annak ezek azok ezen azon ott itt majd ismét tovább
további adott egyes összes mely melyek melynek amelynek amíg míg mikor amikor újra után azonnal""".split())
SUFFIX = re.compile(r"(jának|jének|ának|ének|ainak|einek|jával|jével|ával|ével|ból|ből|ról|ről|tól|től|hoz|hez|höz|"
                    r"ban|ben|nak|nek|val|vel|nál|nél|ra|re|ba|be|on|en|ön|at|et|ot|öt|ok|ek|ak|ai|ei|ja|je|t)$")
FULLWORD = re.compile(r"(?<![\w])[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]{2,}(?![\w])")
TAG = re.compile(r"</?f\d+>")

PROMPT = """Ты — терминолог, переводишь эксплуатационную документацию АЭС (ВВЭР-440) с венгерского на русский.
Ниже венгерские термины комплекта документов (в той форме, как встретились) и примеры употребления.
Для каждого дай словарную (начальную) венгерскую форму и русский термин, принятый в российской эксплуатационной документации АЭС с ВВЭР (в начальной форме).
Правила:
- Разбирай слово на морфемы. Суффиксы -talan/-telen/-tlan/-tlen и приставка in- означают отрицание/лишение: tömörtelenség = «негерметичность» (не «разрыв», не «герметичность»), üzemképtelen = «неработоспособный», inhermetikusság = «негерметичность».
- Глагольные приставки меняют смысл: fel- (вверх, увеличение), le- (вниз, снижение), ki-, be-, át-, vissza-; felbóroz = «повысить концентрацию бора» (не «разогреть»).
- Различай системы: biztonsági hűtővíz (BHV) — техническая вода ответственных потребителей, а не САОЗ (ZÜHR).
- kiégett üzemanyag/kazetta — «отработавшее топливо/ОТВС»; átrakás — «перегрузка».
- Если это общеупотребительное слово, а не термин, ставь "skip": true.
- Если не уверен в русском термине, ставь "sure": false.
Ответ — только JSON-массив объектов {"n": номер, "hu": "...", "ru": "...", "skip": false, "sure": true}.

"""


def stem(w):
    for _ in range(2):
        s = SUFFIX.sub("", w)
        if len(s) < 4:
            break
        w = s
    return w


def candidates(texts, entries, top=150):
    cnt, forms, example = Counter(), defaultdict(Counter), {}
    for t in texts:
        # целые слова; ПРОПИСНЫЕ (аббревиатуры, AZONNAL) — разрыв сочетания, остальное в нижний регистр
        toks = []
        for w in FULLWORD.findall(TAG.sub(" ", t)):
            toks.append("|" if w.isupper() else w.lower())
        for n in (1, 2, 3):
            for i in range(len(toks) - n + 1):
                gram = toks[i:i + n]
                if "|" in gram or gram[0] in STOP or gram[-1] in STOP or any(len(w) < 4 for w in gram):
                    continue
                if n == 1 and len(gram[0]) < 6:
                    continue
                key = " ".join(stem(w) for w in gram)
                cnt[key] += 1
                surf = " ".join(gram)
                forms[key][surf] += 1
                ex = example.setdefault(key, [])
                if len(t) >= 30 and t not in ex:
                    ex.append(t); ex.sort(key=len); del ex[3:]   # 3 самых коротких содержательных примера
    out = []
    for key, c in cnt.most_common(top * 6):
        if c < 4:
            break
        surf = forms[key].most_common(1)[0][0]
        if G.find(surf, entries, limit=1):          # глоссарий уже покрывает
            continue
        exs = example.get(key) or [surf]
        out.append((surf, c, " || ".join(e[:220] for e in exs[:2])))
    # сочетание, целиком входящее в более частое длинное с той же частотой, — лишнее
    keep = []
    for s, c, ex in sorted(out, key=lambda x: -len(x[0].split())):
        if any(s in k and abs(c - kc) <= max(2, c // 5) for k, kc, _ in keep):
            continue
        keep.append((s, c, ex))
    keep.sort(key=lambda x: -x[1] * (1 + 0.5 * (len(x[0].split()) - 1)))
    return keep[:top]


def translate_terms(cands, batch=25):
    out = []
    for i in range(0, len(cands), batch):
        part = cands[i:i + batch]
        body = "\n".join(f'{k+1}. «{s}» — примеры: {ex}' for k, (s, c, ex) in enumerate(part))
        r = chat([{"role": "user", "content": PROMPT + body}], max_tokens=9000, think=True, tag="termx2")
        m = re.search(r"\[.*\]", r.get("text", ""), re.S)
        try:
            arr = json.loads(m.group()) if m else []
        except json.JSONDecodeError:
            arr = []
        for o in arr:
            try:
                k = int(o.get("n")) - 1
            except (TypeError, ValueError):
                continue
            if 0 <= k < len(part) and not o.get("skip") and o.get("ru"):
                s, c, ex = part[k]
                out.append({"hu": o.get("hu") or s, "ru": o["ru"], "count": c, "surface": s, "example": ex,
                            "sure": o.get("sure", True) is not False})
    return out


def write_md(terms, path, title):
    lines = [f"# Дополнительный глоссарий (авто) — {title}", "",
             "Собран предпроходом termx.py: частые сочетания документа вне основного глоссария, перевод порталом.",
             "Проверить и поправить до полного прогона; строку можно удалить. Основной глоссарий приоритетнее.", "",
             "## Термины документа (авто)", "", "| Венгерский | Русский | Комментарий |", "|---|---|---|"]
    for t in terms:
        hu = t["hu"].replace("|", "/")
        mark = "авто" + ("" if t.get("sure", True) else ", НЕ УВЕРЕН")
        lines.append(f"| {hu} | {t['ru'].replace('|', '/')} | {mark}; встречается {t['count']} раз |")
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")
