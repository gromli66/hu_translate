# -*- coding: utf-8 -*-
"""Ряд — единица таблицы при нарезке (UNIFIED_DESIGN.md, П1–П2, этап E5). Граница ряда решается здесь — до склейки строк
в абзацы (layout.group) и до склейки абзацев по меткам портала (roles.apply); абзац её не пересекает. Свидетельства по
старшинству:
  1. тег TR — только у размеченного PDF (путь docx_struct, сюда не доходит);
  2. горизонтальная линейка между строками (rule_between);
  3. якорь слева: номер пункта или код оборудования — ближайший кусок слева на той же линии (row_anchor) или кусок
     столбца, где такие куски преобладают (Rows.new_row), через вертикальные линейки: они делят столбцы, а не ряды;
  4. кластер по y — строки через зазор layout.group не склеивает и так.
Столбик отметок (○ □ O ✓ ☒ ◉) — каждая отметка сама по себе (FORM_MARK)."""
import re

ANCHOR = re.compile(r"^\(?\d+(?:\.\d+)*\.?\)?$")
FORM_MARK = re.compile(r"^[○□△▲◉☒✓✔✗✘■●O]$")


def code_line(t):
    """Строка из кодов («04.TH10D001 KM1/1», «30BW07», «03/A TJ22D001»): каждое слово — с цифрой или аббревиатура до
    4 прописных, слов с цифрами не меньше половины. Такие строки — ячейки перечня без линеек, а не строки абзаца:
    столбец кодов склеивался в один абзац, код и его значение — в один кусок (3PR42 л. 112, 17.10)."""
    toks = [x for x in t.split() if re.search(r"\w", x)]
    dig = [x for x in toks if re.search(r"\d", x)]
    return bool(toks) and 2 * len(dig) >= len(toks) and all(
        re.search(r"\d", x) or (x.isupper() and len(re.sub(r"\W", "", x)) <= 4) for x in toks)


def is_anchor(t):
    """Якорь ряда: номер пункта («1.2.») или до трёх слов из кодов с цифрой («31UJ21S601», «YA-K7»)."""
    t = t.strip()
    return bool(ANCHOR.match(t)) or (code_line(t) and len(t.split()) <= 3 and bool(re.search(r"\d", t)))


def row_anchor(frags, bb):
    """Якорь строки таблицы без горизонтальных линий: ближайший слева на той же линии кусок — код оборудования или номер
    пункта (через вертикальные линейки колонок — они делят колонки, а не строки). Возвращает его рамку или None."""
    x0, y0, x1, y1 = bb
    best = None
    for f in frags:
        fb = f["bbox"]
        if fb[2] <= x0 - 5 and min(y1, fb[3]) - max(y0, fb[1]) > 0.5 * (y1 - y0) and (best is None or fb[2] > best[2]):
            best = fb
            best_t = f["t"].strip()
    if best is None:
        return None
    return tuple(best) if is_anchor(best_t) else None


def _mid(b):
    return (b[1] + b[3]) / 2


def _of_line(r, L, other):
    """Линейка — подчёркивание строки L: в пределах её текста по ширине и её рамки по высоте и не обрамляет другую строку
    с обеих сторон — или обрамляет, но выровнена с текстом L по краю. Линейка ячейки шапки «Beépítés helye» лежит и под
    широким названием таблицы над ней («Segédenergiával működtetett armatúrák», 3PR10 INFO) — граница ряда; подчёркнутый
    заголовок по центру в две строки («Reaktor és Készülék … ellenőrzési» / «terjedelme», 30PR132GT) — одна фраза."""
    if not (L[0] - 2 <= r[0] and r[1] <= L[2] + 2 and L[1] - 0.5 <= r[2] <= L[3] + 0.5):
        return False
    return not (r[0] < other[0] - 2 and r[1] > other[2] + 2) or abs(r[0] - L[0]) <= 2 or abs(r[1] - L[2]) <= 2


def rule_between(hs, a, b, soft=True):
    """Горизонтальная линейка между строкой a и строкой b под ней: в промежутке их рамок (soft=False — стороны белой
    подложки, layout.rules, не считаются) или между серединами строк, если она выходит за ширину их текста. Рамки
    моноширинных строк выше шага (15–19 pt при шаге 10–11) накрывают линейку ряда, и она выпадала из проверки:
    заголовок таблицы склеивался с шапкой столбца («2. Segédenergiával működtetett armatúrák Beépítés helye», 3SZ23 INFO;
    «1.1 … Tevékenység», 3PR10), ячейки соседних рядов — в одну («Tüzivíz gömbcsap hajtás» ×3, 3FI18; «Bórsavkoncentráció
    Átlátszóság», 02VE25). Не граница ряда — подчёркивание (_of_line) и сторона заливки без обводки (edge, layout.rules):
    серая полоса под каждой строкой заголовка («10.1 UJ21 … jelzés» / «esetén», 3SZ23 л. 18)."""
    ma, mb = _mid(a), _mid(b)
    for r in hs:
        rx0, rx1, y = r
        if min(b[2], rx1) - max(b[0], rx0) <= 5:
            continue
        hard = not getattr(r, "soft", False)
        if a[3] - 0.5 <= y <= b[1] + 0.5 and (soft or hard):
            return True
        if hard and not getattr(r, "edge", False) and ma + 1 < y < mb - 1 and not _of_line(r, a, b) \
                and not _of_line(r, b, a) and min(a[2], rx1) - max(a[0], rx0) > 5:
            return True
    return False


class Rows:
    """Границы рядов листа: frags — куски строк листа ({t, bbox, size}), hs — горизонтальные линейки (layout.rules)."""

    def __init__(self, frags, hs):
        self.frags, self.hs = frags, hs
        cand = [f for f in frags if is_anchor(f["t"])]
        self.anchors = []
        for f in cand:
            # столбец якоря — куски с тем же левым или правым краем или той же серединой (±3 pt; рамка куска бывает с
            # пробелами впереди: «32UJ21L501» с x0 = 71 вместо 114, 3SZ23 л. 26); якоря в нём преобладают. Коды внутри
            # текста столбца «Tevékenység» (3SZ23: «31(32)UJ21S603» под фразой) — содержание, а не якоря рядов
            b = f["bbox"]
            col = [g for g in frags if abs(g["bbox"][0] - b[0]) <= 3 or abs(g["bbox"][2] - b[2]) <= 3
                   or abs(g["bbox"][0] + g["bbox"][2] - b[0] - b[2]) <= 6]
            n = sum(1 for g in col if is_anchor(g["t"]))
            if n >= 2 and 2 * n >= len(col):
                self.anchors.append(tuple(b))

    def _own(self, b):
        """Якорь столбца якорей слева на уровне строки b (перекрывает больше половины её рамки)."""
        h = b[3] - b[1]
        best, ov = None, 0.5 * h
        for A in self.anchors:
            o = min(b[3], A[3]) - max(b[1], A[1])
            if A[2] <= b[0] - 3 and o > ov:
                best, ov = A, o
        return best

    def _above(self, b):
        """Последний якорь столбца якорей слева выше строки b — ряд, которому строка принадлежит без своего якоря."""
        up = [A for A in self.anchors if A[2] <= b[0] - 3 and _mid(A) < _mid(b) - 2]
        return max(up, key=_mid) if up else None

    def new_row(self, fa, fb):
        """Строка fb начинает новый ряд против строки fa над ней: у fb на её уровне свой якорь, у fa — другой, свой на
        её уровне (значения столбца без кодов против столбца кодов: «-3,6m helyszín … 9,6m helyszín» шесть рядов одним
        сегментом, 3SZ23 INFO; «126-128 bar», 3PR10 INFO) или унаследованный сверху (или никакого — шапка над первым
        рядом) через зазор больше полутора кеглей (ряды «YA-K7 ⇥ …» / «YA-K8 ⇥ …», 3PR10). Якорь по центру многострочной ячейки («4.1.9.» против второй из трёх
        строк, 30HU375GT) зазора не даёт — строки не делит."""
        a, b = fa["bbox"], fb["bbox"]
        if _mid(b) - _mid(a) < 2:
            return False
        ob = self._own(b)
        if ob is None:
            return False
        oa = self._own(a)
        if oa is not None:
            return oa != ob
        # без своего якоря строка fa — в ряду якоря выше (или в шапке над первым рядом: «Elosztó helye» над «E601/II»,
        # 3FI18 л. 38): новый ряд — только через зазор рядов
        return self._above(a) != ob and _mid(b) - _mid(a) > 1.5 * max(fa.get("size") or 0, fb.get("size") or 0)

    def apart(self, fa, fb, soft=False):
        """Строки fa (выше) и fb — в разных рядах: причина («линейка», «якорь», «отметка») или None. soft=True — граница и
        сторона белой подложки (layout.group, строка не продолжает фразу); метки портала через подложку склеивают
        (roles.apply: 30RL2126GT — плашка под каждой строкой, «…akkor a» / «30RL21S001 zárva van.» рвались)."""
        a, b = fa["bbox"], fb["bbox"]
        if rule_between(self.hs, a, b, soft):
            return "линейка"
        a1, a2 = row_anchor(self.frags, a), row_anchor(self.frags, b)
        if a1 is not None and a2 is not None and a1 != a2 or self.new_row(fa, fb):
            return "якорь"
        if FORM_MARK.match(fa["t"].strip()) or FORM_MARK.match(fb["t"].strip()):
            return "отметка"
        return None
