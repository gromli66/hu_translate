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

import patterns as PT
from lines import hyphen_break, mid_phrase

ANCHOR = re.compile(rf"^{PT.ITEM_LABEL}$")         # метка пункта: «1.», «1.2.», «(3)», «a)»
# отметка бланка; значки символьных шрифтов (частная область \uf020–\uf0ff: □ Wingdings у пунктов «30YA62P401ZQ35 p = …,»
# 30PR132GT л. 5) — тоже
FORM_MARK = re.compile(r"^[○□△▲◉☒✓✔✗✘■●O\uf020-\uf0ff]$")


def code_line(t):
    """Строка из кодов («04.TH10D001 KM1/1», «30BW07», «03/A TJ22D001»): каждое слово — с цифрой или аббревиатура до
    4 прописных, слов с цифрами не меньше половины. Такие строки — ячейки перечня без линеек, а не строки абзаца:
    столбец кодов склеивался в один абзац, код и его значение — в один кусок (3PR42 л. 112, 17.10)."""
    toks = [x for x in t.split() if re.search(r"\w", x)]
    dig = [x for x in toks if re.search(r"\d", x)]
    return bool(toks) and 2 * len(dig) >= len(toks) and all(
        re.search(r"\d", x) or (x.isupper() and len(re.sub(r"\W", "", x)) <= 4) for x in toks)


def code_enum(left, right):
    """Коды через запятую — перечисление внутри фразы, а не ячейки: «30TD97S202,  30TD98S201,  …» — промежутки растянула
    выключка по ширине (02VE25 л. 44, раздел 5.2): строка шла пятью «ячейками», коды разлетались по сегментам вразнобой,
    в переводе строки начинались с середины (NIGHT2_EYES п. 14)."""
    return code_line(left) and code_line(right) and left.rstrip().endswith(",")


def code_cells(left, right):
    """Две части строки через промежуток от CODE_GAP (layout) — разные ячейки перечня («TK41D001 ⇥ KM2/1»): обе — строки
    из кодов, и это не перечисление через запятую (code_enum)."""
    return code_line(left) and code_line(right) and not code_enum(left, right)


def code_rows(upper, lower):
    """Две строки из кодов подряд — ячейки перечня, каждая своим куском; кроме перечисления, которое продолжается на
    следующей строке: в верхней от двух кодов и запятая в конце («30TD97S202, 30TD98S201, … 40TD97S202,»)."""
    up = upper.rstrip()
    return code_line(up) and code_line(lower) and not (up.endswith(",") and len(up.split()) >= 2)


def is_anchor(t):
    """Якорь ряда: метка пункта («1.2.», «(3)», «a)») или до трёх слов из кодов с цифрой («31UJ21S601», «YA-K7»,
    «30BW07 ZÁR»). Одно определение для нарезки (Rows) и сборки (docx_layout._table_block, E9b-R): в сборке было своё —
    patterns.ROW_ANCHOR (метка с буквой, но код — только один KKS), при нарезке — номер без буквы."""
    t = t.strip()
    return bool(ANCHOR.match(t)) or (code_line(t) and len(t.split()) <= 3 and bool(re.search(r"\d", t)))


def seg_frag(group):
    """Кусок строки листа ({t, bbox, size}, как у нарезки) из сегментов JSON одной строки (docx_layout._lines): рамка —
    первые строки сегментов (loc.frags), текст — через пробел. По таким кускам Rows работает и в сборке."""
    fr = [(s["loc"].get("frags") or [s["loc"]["bbox"]])[0] for s in group]
    return {"t": " ".join(s["text"] for s in group), "size": max(s["loc"].get("size") or 0 for s in group),
            "bbox": (min(f[0] for f in fr), min(f[1] for f in fr), max(f[2] for f in fr), max(f[3] for f in fr))}


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


def _cell_edge(r, vs):
    """Горизонтальная линейка упирается концом в вертикальную — граница ячейки таблицы, а не подчёркивание. Стороны белой
    подложки и заливки (soft, edge) не в счёт: подчёркнутая строка на плашке упирается в её край."""
    return any(min(abs(v[2] - r[0]), abs(v[2] - r[1])) <= 2 and v[0] - 2 <= r[2] <= v[1] + 2
               and not getattr(v, "soft", False) and not getattr(v, "edge", False) for v in vs)


def _framed(r, L, vs):
    """Линейка — сторона рамки вокруг самой строки L: оба конца совпадают с краями её текста, а вертикали у концов (если
    есть) охватывают и саму строку (1SZ01 л. 43: каждая строка фразы в своей обведённой рамке — «…a kö-» / «vetkező …
    ELLENŐRZÉ-» / «SÉT …» резались на три сегмента). Верх таблицы под заголовком той же ширины («9./ Ellenőrizte: …» над
    «Dátum:», 3SZ17 л. 2) — вертикали уходят вниз, мимо строки: граница."""
    ends = [v for v in vs if min(abs(v[2] - r[0]), abs(v[2] - r[1])) <= 2 and v[0] - 2 <= r[2] <= v[1] + 2
            and not getattr(v, "soft", False) and not getattr(v, "edge", False)]
    return abs(r[0] - L[0]) <= 2 and abs(r[1] - L[2]) <= 2 and (not ends or any(v[0] <= _mid(L) <= v[1] for v in ends))


def _of_line(r, L):
    """Линейка — подчёркивание строки L: в пределах её текста по ширине (±4 pt: хвостовой пробел подчёркнут) и её рамки
    по высоте. Подчёркнутый заголовок по центру в две строки («Reaktor és Készülék … ellenőrzési» / «terjedelme»,
    30PR132GT) — одна фраза."""
    return L[0] - 4 <= r[0] and r[1] <= L[2] + 4 and L[1] - 0.5 <= r[2] <= L[3] + 0.5


def rule_between(hs, a, b, soft=True, vs=()):
    """Горизонтальная линейка между строкой a и строкой b под ней: в промежутке их рамок (soft=False — стороны белой
    подложки, layout.rules, не считаются) или между серединами строк, если это граница ячейки (упирается в вертикальную
    линейку) или она выходит за ширину текста обеих строк. Рамки моноширинных строк выше шага (15–19 pt при шаге 10–11)
    накрывают линейку ряда, и она выпадала из проверки: заголовок таблицы склеивался с шапкой столбца
    («2. Segédenergiával működtetett armatúrák Beépítés helye», 3SZ23 INFO; «1.1 … Tevékenység», 3PR10), ячейки соседних
    рядов — в одну («Tüzivíz gömbcsap hajtás» ×3, 3FI18; «Bórsavkoncentráció Átlátszóság», 02VE25). Не граница ряда —
    подчёркивание (_of_line) и сторона заливки без обводки (edge, layout.rules): серая полоса под каждой строкой
    заголовка («10.1 UJ21 … jelzés» / «esetén», 3SZ23 л. 18)."""
    ma, mb = _mid(a), _mid(b)
    for r in hs:
        rx0, rx1, y = r
        if min(b[2], rx1) - max(b[0], rx0) <= 5:
            continue
        hard = not getattr(r, "soft", False)
        if a[3] - 0.5 <= y <= b[1] + 0.5 and (soft or hard):
            return True
        if hard and not getattr(r, "edge", False) and ma + 1 < y < mb - 1 and min(a[2], rx1) - max(a[0], rx0) > 5 \
                and not (_framed(r, a, vs) or _framed(r, b, vs)) \
                and (_cell_edge(r, vs) or not (_of_line(r, a) or _of_line(r, b))):
            return True
    return False


def blank_between(fa, fb):
    """Между строкой fa и строкой fb под ней пустая строка — абзац кончился: середины дальше 1,8 кегля (строки абзаца — через
    1,0–1,4 кегля, через пустую строку — 2,1–2,4; замер по 7 документам). Рамки моноширинных строк выше шага, и зазор
    рамок пустую строку не видит: «02TK47S298 … zárva.» + ряд «30TD97S201» (02VE25 л. 44). Кроме фразы, что явно
    продолжается (mid_phrase): «Várakozási idő: 10 perc,» / пустая строка / «vagy ha a kezelő …» (02VE25 л. 13), и слова,
    разорванного переносом: шапка таблицы с просторным межстрочием «Megjelení-» / «tés helye» (1SZ01 л. 90, шаг 2,3 кегля)."""
    size = max(fa.get("size") or 0, fb.get("size") or 0)
    return size > 0 and _mid(fb["bbox"]) - _mid(fa["bbox"]) > 1.8 * size \
        and not (mid_phrase(fa["t"], fb["t"]) or hyphen_break(fa["t"], fb["t"]))


class Rows:
    """Границы рядов листа: frags — куски строк листа ({t, bbox, size}), hs, vs — линейки (layout.rules), gutter — x
    просвета двух колонок сплошного текста (layout.gutter) или None: через него якоря не берутся."""

    def __init__(self, frags, hs, vs=(), gutter=None):
        self.hs, self.vs, self.gutter = hs, vs, gutter
        self.frags = sorted(frags, key=lambda f: (round(f["bbox"][1] / 3), f["bbox"][0]))     # порядок layout.group
        self.marks = [tuple(f["bbox"]) for f in self.frags if FORM_MARK.match(f["t"].strip())]
        flag = [is_anchor(f["t"]) for f in self.frags]
        self.anchors = []
        for f, ok in zip(self.frags, flag):
            if not ok:
                continue
            # столбец якоря — куски с тем же левым или правым краем или той же серединой (±3 pt; рамка куска бывает с
            # пробелами впереди: «32UJ21L501» с x0 = 71 вместо 114, 3SZ23 л. 26); якоря в нём преобладают. Коды внутри
            # текста столбца «Tevékenység» (3SZ23: «31(32)UJ21S603» под фразой) — содержание, а не якоря рядов
            b = f["bbox"]
            col = [k for k, g in enumerate(self.frags) if abs(g["bbox"][0] - b[0]) <= 3 or abs(g["bbox"][2] - b[2]) <= 3
                   or abs(g["bbox"][0] + g["bbox"][2] - b[0] - b[2]) <= 6]
            n = sum(1 for k in col if flag[k])
            if n >= 2 and 2 * n >= len(col):
                self.anchors.append(tuple(b))

    def _own(self, b):
        """Якорь столбца якорей слева на уровне строки b (перекрывает больше половины её рамки), не за просветом колонок:
        нумерованный столбец левой колонки — не ряды для текста правой."""
        h = b[3] - b[1]
        best, ov = None, 0.5 * h
        for A in self.anchors:
            o = min(b[3], A[3]) - max(b[1], A[1])
            if A[2] <= b[0] - 3 and o > ov and not (self.gutter is not None and A[2] < self.gutter < b[0]):
                best, ov = A, o
        return best

    def new_row(self, fa, fb):
        """Строка fb начинает новый ряд против строки fa над ней: у fb на её уровне свой якорь, а у fa — другой свой
        (значения столбца без кодов против столбца кодов: «-3,6m helyszín … 9,6m helyszín» шесть рядов одним сегментом,
        3SZ23 INFO; «126-128 bar», 3PR10 INFO) или своего нет, но между строками зазор больше полутора кеглей (ряды
        «YA-K7 ⇥ …» / «YA-K8 ⇥ …», 3PR10; шапка над первым рядом). Не делят строки: якорь по центру многострочной ячейки
        («4.1.9.» против второй из трёх строк, 30HU375GT) и фраза или слово, продолжающиеся на следующей строке при
        обычном шаге, — ячейка во много рядов против столбца кодов, колонка текста рядом с нумерованным столбцом."""
        a, b = fa["bbox"], fb["bbox"]
        if _mid(b) - _mid(a) < 2:
            return False
        ob = self._own(b)
        if ob is None:
            return False
        size = max(fa.get("size") or 0, fb.get("size") or 0)
        oa = self._own(a)
        if oa is not None:
            # две строки с одного слова — значения столбца («csapágy szegm. fémhőfok» ×2, «főolaj sz. …» ×2, 1SZ01
            # л. 90–95), а не продолжение фразы
            cont = (mid_phrase(fa["t"], fb["t"]) or hyphen_break(fa["t"], fb["t"])) \
                and fa["t"].split()[:1] != fb["t"].split()[:1]
            return oa != ob and not (_mid(b) - _mid(a) <= 1.3 * size and cont)
        # без своего якоря строка fa — в ряду якоря выше (или в шапке над первым рядом: «Elosztó helye» над «E601/II»,
        # 3FI18 л. 38): новый ряд — только через зазор рядов
        return _mid(b) - _mid(a) > 1.5 * size

    def right_mark(self, b):
        """Отметка бланка (○ □ O …) справа от строки b на её линии (перекрывает больше половины меньшей рамки) — ближайшая,
        её рамка; нет — None."""
        best = None
        for m in self.marks:
            if m[0] >= b[2] - 1 and min(b[3], m[3]) - max(b[1], m[1]) > 0.5 * min(b[3] - b[1], m[3] - m[1]) \
                    and (best is None or m[0] < best[0]):
                best = m
        return best

    def apart(self, fa, fb, soft=False):
        """Строки fa (выше) и fb — в разных рядах или абзацах: причина («линейка», «якорь», «отметка», «пустая строка») или
        None. soft=True — граница и сторона белой подложки (layout.group, строка не продолжает фразу); метки портала через
        подложку склеивают (roles.apply: 30RL2126GT — плашка под каждой строкой, «…akkor a» / «30RL21S001 zárva van.»
        рвались). У двух строк справа свои отметки бланка — пункты перечня, каждый со своей отметкой: «30TJ13 ág: …» и
        «30TJ23 ág: …» (30PR25GT л. 13) шли одним абзацем. Пустая строка (blank_between) — и для меток портала: пункты
        «A, …» / «B, …» через пустую строку (1SZ01 л. 112) портал склеивал."""
        a, b = fa["bbox"], fb["bbox"]
        if rule_between(self.hs, a, b, soft, self.vs):
            return "линейка"
        a1, a2 = row_anchor(self.frags, a), row_anchor(self.frags, b)
        if a1 is not None and a2 is not None and a1 != a2 or self.new_row(fa, fb):
            return "якорь"
        m1, m2 = self.right_mark(a), self.right_mark(b)
        if FORM_MARK.match(fa["t"].strip()) or FORM_MARK.match(fb["t"].strip()) or m1 and m2 and m1 != m2:
            return "отметка"
        if blank_between(fa, fb):
            return "пустая строка"
        return None
