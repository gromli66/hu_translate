# -*- coding: utf-8 -*-
"""Роли строк листа порталом для PDF без разметки (текстовый слой): где начинается абзац и что это — заголовок, пункт
списка, подпись, колонтитул, таблица, текст. Геометрию, таблицы и колонтитулы решают правила (layout, docx_layout);
портал только ставит строкам метки — текст он не трогает, потерять его не может.
Замер (work/portal_fmt_2026-10-18, 56 листов, 18.10): F1 границы абзаца — правила 0,84, портал 0,98; заголовок
0,25 / 0,69; начало пункта 0,40 / 0,90; колонтитулы портал отмечает хуже (0,62 против 0,81) — их не берём."""
import json
import re

PROMPT = """Ниже — строки одного листа документа (PDF без разметки, Word), в порядке чтения. Для каждой строки: номер | левый-правый край (pt) | верх (pt) | промежуток до предыдущей строки (pt) | кегль | Ж — полужирная | П — такая же строка есть на том же месте других листов | текст. Ширина листа {W} pt, высота {H} pt; основной левый край текста ≈ {L} pt, правый ≈ {Rr} pt.

Восстанови структуру листа, как она была в исходном Word. Для каждой строки дай:
- new: 1 — строка начинает новый абзац (или новую ячейку, новый пункт, новый заголовок, новую строку колонтитула); 0 — продолжает абзац предыдущей строки (перенос внутри абзаца);
- role — роль абзаца, к которому строка относится:
  heading — заголовок раздела/подраздела (обычно короткий, часто с номером «4.1.2», полужирный или крупнее, без точки в конце; НЕ заголовок — обычное предложение, даже если оно начинается с номера пункта «1.1. Настоящая инструкция предназначена …»);
  list — пункт списка (маркер «•», «–», «-», номер «1)», буква «a)», «б)», «ac)», «(b)» в начале — даже если пункты оформлены обычными абзацами); все строки пункта — list;
  caption — подпись рисунка или таблицы («Рисунок 3 –», «Таблица 2 –», «3. ábra»);
  header — колонтитул: строки в самом верху или внизу листа вне текста (название документа, номер листа, шифр, «Стр. N»), и рамка-таблица со сведениями о документе внизу или вверху листа; строка с пометкой П почти всегда колонтитул;
  table — текст в ячейках таблицы (строки, стоящие в несколько колонок на одной высоте, или в сетке);
  body — обычный текст.
Признаки переноса внутри абзаца: предыдущая строка доходит почти до правого края, промежуток обычный (как между строками абзаца), строка начинается у того же левого края, нет маркера и номера. Признаки нового абзаца: предыдущая строка короткая (кончается раньше правого края) и кончается точкой или двоеточием, отступ первой строки, увеличенный промежуток, маркер или номер.
Ответь только JSON-массивом без пояснений: [{{"id": 0, "new": 1, "role": "header"}}, ...] — по объекту на каждую строку, все {N} строк.

Строки:
{rows}"""
ROLES = ("heading", "list", "caption", "header", "table", "body")
MIN_COVER = 0.8                     # ответ покрыл меньше 80 % строк (оборван) — лист по правилам


def _bold(s):
    return bool(s["flags"] & 16) or "bold" in s["font"].lower()


def page_lines(page):
    """Строки листа PyMuPDF: [{id, x0, x1, top, bottom, size, bold, text}] в порядке чтения; куски одной строки
    вплотную — одной строкой."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            sp = [s for s in l["spans"] if s["text"].strip()]
            if not sp:
                continue
            t = "".join(s["text"] for s in l["spans"]).strip()
            x0, y0, x1, y1 = l["bbox"]
            if out and abs(out[-1]["top"] - y0) < 2 and 0 <= x0 - out[-1]["x1"] < 12:
                o = out[-1]
                o["text"] += " " + t
                o["x1"] = max(o["x1"], x1)
                o["bold"] = o["bold"] and all(_bold(s) for s in sp)
                continue
            out.append({"id": len(out), "x0": round(x0, 1), "x1": round(x1, 1), "top": round(y0, 1),
                        "bottom": round(y1, 1), "size": round(max(s["size"] for s in sp), 1),
                        "bold": all(_bold(s) for s in sp), "text": t})
    return out


TAIL = "\nТы ответил не на все строки. Ответь только JSON-массивом для строк с номерами {ids} (остальные не повторяй)."


def _sig(t):
    return re.sub(r"[\d\W]+", "", t.lower())[:40]


def repeats(doc):
    """Колонтитулы: (подпись строки, высота с шагом 3 pt) у края листа (верхняя или нижняя седьмая часть), что стоят на
    одной и той же высоте на 3 и больше листах и не меньше чем на 30 % листов. Высота — обязательно: «Korlátozó
    feltétel:» ÜFK стоит вверху многих листов (разделы с нового листа), но на разной высоте — это не колонтитул (18.10)."""
    cnt = {}
    for p in doc:
        H = p.rect.height
        for l in page_lines(p):
            if (l["bottom"] < H / 7 or l["top"] > H * 6 / 7) and _sig(l["text"]):
                cnt.setdefault((_sig(l["text"]), round(l["top"] / 3)), set()).add(p.number)
    need = max(3, 0.3 * len(doc))
    return {k for k, v in cnt.items() if len(v) >= need}


def is_rep(l, rep, H):
    k = _sig(l["text"])
    return len(k) >= 4 and (l["bottom"] < H / 7 or l["top"] > H * 6 / 7) and any(
        (k, round(l["top"] / 3) + d) in rep for d in (-1, 0, 1))


def _rows(lines, rep, H):
    out, prev = [], None
    for l in lines:
        gap = round(l["top"] - prev["bottom"], 1) if prev else "-"
        out.append(f'{l["id"]} | {l["x0"]}-{l["x1"]} | y{l["top"]} | {gap} | {l["size"]} | {"Ж" if l["bold"] else "-"} | '
                   f'{"П" if is_rep(l, rep, H) else "-"} | {l["text"][:160]}')
        prev = l
    return "\n".join(out)


def _parse(txt):
    m = re.search(r"\[.*\]", txt or "", re.S)
    try:
        arr = json.loads(m.group(0)) if m else []
    except json.JSONDecodeError:
        arr = []
    return {a["id"]: {"new": bool(a.get("new", 1)), "role": a["role"]} for a in arr
            if isinstance(a, dict) and isinstance(a.get("id"), int) and a.get("role") in ROLES}


def ask(lines, W, H, rep=frozenset()):
    """Метки строк порталом: {id: {"new", "role"}}; оборванный ответ — дозапрос недостающих строк (до двух раз); не
    набралось 80 % строк — пусто (лист по правилам). Строка с пометкой П (повтор у края листа) — колонтитул."""
    import llm
    if not lines:
        return {}
    xs = sorted(l["x0"] for l in lines)
    x1s = sorted(l["x1"] for l in lines)
    p = PROMPT.format(W=round(W), H=round(H), L=round(xs[len(xs) // 4]), Rr=round(x1s[len(x1s) * 3 // 4]), N=len(lines),
                      rows=_rows(lines, rep, H))
    try:
        got = _parse(llm.chat([{"role": "user", "content": p}], max_tokens=8000, tag="fmt3").get("text"))
    except RuntimeError:                                    # нет токена, отказ доступа — лист по правилам
        return {}
    for _ in range(2):
        miss = [l["id"] for l in lines if l["id"] not in got]
        if not miss:
            break
        try:
            more = _parse(llm.chat([{"role": "user", "content": p + TAIL.format(ids=", ".join(map(str, miss)))}],
                                   max_tokens=8000, tag="fmt3c").get("text"))
        except RuntimeError:
            break
        got.update({k: v for k, v in more.items() if k in miss})
    for l in lines:                                         # повтор у края листа — колонтитул (портал путал с заголовком)
        if l["id"] in got and is_rep(l, rep, H):
            got[l["id"]] = {"new": True, "role": "header"}
    return got if len(got) >= MIN_COVER * len(lines) else {}


def doc_labels(path, pages, workers=4):
    """{лист от 0: (строки, метки)} для листов pages документа."""
    import fitz
    from concurrent.futures import ThreadPoolExecutor
    with fitz.open(str(path)) as doc:
        rep = repeats(doc)
        lines = {p: page_lines(doc[p]) for p in pages}
        size = {p: (doc[p].rect.width, doc[p].rect.height) for p in pages}
    with ThreadPoolExecutor(workers) as ex:
        labs = dict(zip(pages, ex.map(lambda p: ask(lines[p], *size[p], rep), pages)))
    return {p: (lines[p], labs[p]) for p in pages}


def _line_of(bb, lines):
    """Строка листа, на которой лежит кусок bb (перекрытие по высоте и ширине)."""
    best, score = None, 0.0
    for l in lines:
        h = min(bb[3], l["bottom"]) - max(bb[1], l["top"])
        w = min(bb[2], l["x1"]) - max(bb[0], l["x0"])
        if h <= 0.4 * min(bb[3] - bb[1], l["bottom"] - l["top"]) or w <= 0:
            continue
        if h * w > score:
            best, score = l, h * w
    return best


def apply(pars, lines, labels, join_lines):
    """Абзацы листа (layout.group) по меткам портала: абзац делится на строке, которую портал начал новым абзацем
    (правила склеивали соседние абзацы — полнота границ 0,77), и склеивается со следующим, если портал продолжает им
    тот же абзац строкой подряд (ложные разрывы правил). Абзацу — роль первой строки: p["role"] (heading, list,
    caption; колонтитулы и таблицы — правилами, роль не ставится). Колонтитулы и таблицы по меткам не трогаются."""
    if not labels:
        return pars
    keep = lambda lab: lab is not None and lab["role"] not in ("header", "table")

    def lab_of(f):
        l = _line_of(f["bbox"], lines)
        return (l["id"], labels.get(l["id"])) if l else (None, None)
    out = []
    for p in pars:                                          # деление
        cur, prev_line = [], None
        for f in p["frags"]:
            lid, lab = lab_of(f)
            start = lid is not None and lid != prev_line and keep(lab) and lab["new"]
            if cur and start:
                out.append(cur); cur = []
            cur.append((f, lid, lab))
            if lid is not None:
                prev_line = lid
        if cur:
            out.append(cur)
    merged = []                                             # склейка
    for piece in out:
        if merged:
            last, first = merged[-1], piece[0]
            lid_last = next((x[1] for x in reversed(last) if x[1] is not None), None)
            if first[1] is not None and lid_last is not None and first[1] == lid_last + 1 and keep(first[2]) \
                    and not first[2]["new"] and keep(next((x[2] for x in last if x[1] is not None), None)) \
                    and first[0]["bbox"][1] > last[-1][0]["bbox"][1] + 1 \
                    and min(first[0]["bbox"][2], max(x[0]["bbox"][2] for x in last)) \
                    - max(first[0]["bbox"][0], min(x[0]["bbox"][0] for x in last)) > 0:
                last += piece
                continue
        merged.append(list(piece))
    res = []
    for piece in merged:
        frags = [x[0] for x in piece]
        bb = [min(f["bbox"][0] for f in frags), min(f["bbox"][1] for f in frags),
              max(f["bbox"][2] for f in frags), max(f["bbox"][3] for f in frags)]
        lab = next((x[2] for x in piece if x[1] is not None), None)
        fam = {}
        for f in frags:
            fam[f.get("family", "serif")] = fam.get(f.get("family", "serif"), 0) + len(f["t"])
        p = {"frags": frags, "bbox": bb, "size": max(f["size"] for f in frags),
             "bold": all(f["bold"] for f in frags), "text": join_lines([f["t"] for f in frags]),
             "family": max(fam, key=fam.get) if fam else "serif"}
        if lab is not None and lab["role"] in ("heading", "list", "caption"):
            p["role"] = lab["role"]
        res.append(p)
    return res
