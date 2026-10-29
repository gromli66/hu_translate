# -*- coding: utf-8 -*-
"""«Ворота качества» готовых переводов: после сборки проверяют DOCX/PDF по сегментам задания и оригиналу и
печатают отчёт по-русски (документ, страница, фрагмент). Венгерский знать не нужно: сверяются числа, коды,
множества, повторы и площади. Только хост: pymupdf + lxml (ставится с python-docx), без Docker и LibreOffice.

  python -X utf8 tools/audit_out.py <папка результатов | файл *_RU.docx> ... [--work <data/jobs/uid/work>]
                                    [--orig <папка оригиналов> ...] [--only <часть имени> ...]
                                    [--json отчёт.json] [--fail-on red|yellow|never]

Рядом с `<имя>_RU.docx` берётся `<имя>_RU.pdf`; сегменты (`segs[].text`, `res[id].ru`) — из `<имя>.json`: в
`<--work>/json`, в `<папка результатов>/../json` (результаты задания лежат в work/out) или в любом
`data/jobs/*/work/json` (новые первыми — так проверяются и выложенные копии). Оригинал — `file` из JSON (путь
контейнера /app/... переводится в data/), затем `<--orig>/<имя>.PDF`. Чего нет — проверка пропускается и в отчёте
называется. Код возврата: 1, если есть 🔴 (с --fail-on yellow — и 🟡), иначе 0; 2 — нечего проверять.

Проверки (номера как в аудите 29.10.2026, work/audit_2026-10-29/S*_findings.md):
  1 потеря содержимого — сегмент с переводом, которого нет в DOCX (колонтитулы, сканы, «не требует перевода» не в счёт);
  2 повтор одной фразы ≥3 раз подряд в одной ячейке/абзаце;
  3 шапки: множество пар (идентификатор, «Номер версии») и их чередование по листам, номера листов;
  4 число листов PDF перевода к оригиналу >1,3; листы, где >70 % площади — пустые строки таблицы;
  5 коды оборудования: пропавшие, лишние, повторённые;
  6 отметки ○/□/O: число по листам, слипшиеся «O O O», привязка к строке своего кода;
  7 венгерские буквы в тексте перевода вне кодов и известных меток;
  8 картинка >25 % площади листа там, где в оригинале изображений нет («снимок листа» вместо таблицы)."""
import argparse
import bisect
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

try:
    import pymupdf
except ImportError:                                  # старое имя пакета
    import fitz as pymupdf
from lxml import etree

pymupdf.TOOLS.mupdf_display_errors(False)            # «No common ancestor in structure tree» и т. п. — шум, не результат проверки
REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
try:
    from segments import needs_translation           # то же решение, что у сборки: «не требует перевода»
except Exception:                                    # нет src/ или он занят правкой — упрощённый двойник
    def needs_translation(text):
        t = re.sub(r"</?f\d+>", "", text)
        return bool(re.search(r"[a-záéíóöőúüű]{2,}|\b[A-ZÁÉÍÓÖŐÚÜŰ]{2,}\b|[А-Яа-яЁё]", t)) and not re.fullmatch(r"[\w./-]+\d{3}[\w./-]*", t.strip())

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
RED, YELLOW = "red", "yellow"
ICON = {RED: "🔴", YELLOW: "🟡"}
# \b\d{2}[A-Z]{2}\d{2}[A-Z]\d{3}\b, но справа допускаются буквы: в оригинале код бывает слит со словом («30TK62P101szerint»),
# иначе в переводе он выглядит «лишним» (S4 п.14)
CODE_RE = re.compile(r"(?<![A-Za-z0-9])\d{2}[A-Z]{2}\d{2}[A-Z]\d{3}(?!\d)")
HU_LETTERS = "őűáéíóöúüŐŰÁÉÍÓÖÚÜ"
# отметки: O латиницей, ○/□ и знаки Wingdings (U+F06F/F070 — □, F0A1 — ○, F0A8 — □, F072 — △), как их отдаёт pymupdf
MARKS = {"O", "○", "◯", "□", "☐", "", "", "", "", ""}
# известные метки и сокращения, которые остаются латиницей с венгерскими буквами (глоссарий; 4 буквы и короче проходят и так)
KNOWN_LABELS = {"ÜFK", "ÜTJ", "ÜVFO", "ÁTJ", "ÁTVGT", "BIGT", "KÁ", "ÜBO", "BLÜGY", "HITELESÍTETT"}
RATIO_SHEETS = 1.3          # листов перевода / оригинала
EMPTY_TABLE_SHARE = 0.7     # доля площади листа — пустые строки таблицы (красный); от HALF_EMPTY_SHARE — жёлтый
HALF_EMPTY_SHARE = 0.4
BIG_IMAGE_SHARE = 0.25      # доля площади листа — одна картинка/снимок
SHOW = 8                    # сколько примеров печатать в одном замечании

_TAG = re.compile(r"</?f\d+>")
_NON_WORD = re.compile(r"[\W_]+", re.U)


def norm(s):
    """Только буквы/цифры в нижнем регистре: переводы, склеенные и разбитые по-разному, сравниваются без пробелов,
    кавычек и переносов."""
    return _NON_WORD.sub("", _TAG.sub("", s or "")).lower()


def clip(s, n=90):
    s = re.sub(r"[-]", lambda m: "\\u%04x" % ord(m.group()), s or "")   # знаки Wingdings не печатаются — показать кодом
    s = re.sub(r"\s+", " ", s).strip()
    return s if len(s) <= n else s[:n - 1] + "…"


class Finding:
    def __init__(self, check, level, msg, page=None, orig_page=None, frag=""):
        self.check, self.level, self.msg, self.page, self.orig_page, self.frag = check, level, msg, page, orig_page, frag

    def as_dict(self):
        return dict(check=self.check, level=self.level, msg=self.msg, page=self.page, orig_page=self.orig_page, frag=self.frag)


# ---------------------------------------------------------------- чтение DOCX и PDF
class DocxText:
    """Абзацы всех частей DOCX (тело, таблицы, врезки, колонтитулы). Абзац — свой текст без вложенных абзацев врезки;
    дубль врезки (mc:Fallback) пропускается. cell — id ячейки или None."""

    def __init__(self, path):
        self.paras = []                              # (часть, текст, номер ячейки или None)
        cells = {}                                   # элемент → номер: ссылка держит прокси lxml живым, id() не переиспользуется
        with zipfile.ZipFile(path) as z:
            for n in sorted(z.namelist()):
                if not re.fullmatch(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", n):
                    continue
                root = etree.fromstring(z.read(n))
                for p in root.iter(W + "p"):
                    if any(a.tag == MC + "Fallback" for a in p.iterancestors()):
                        continue
                    tc = next(p.iterancestors(W + "tc"), None)
                    self.paras.append((n, self._own_text(p), None if tc is None else cells.setdefault(tc, len(cells))))
        self.full = norm(" ".join(t for _, t, _ in self.paras))

    @staticmethod
    def _own_text(p):
        out = []

        def walk(el):
            for ch in el:
                if ch.tag == W + "t":
                    out.append(ch.text or "")
                elif ch.tag in (W + "tab", W + "br", W + "cr"):
                    out.append(" ")
                elif ch.tag not in (W + "p", MC + "Fallback", W + "del"):
                    walk(ch)
        walk(p)
        return "".join(out)

    def units(self):
        """Единицы для поиска повторов: ячейка целиком (абзацы через пробел) и абзацы вне таблиц."""
        cells, out = {}, []
        for part, t, tc in self.paras:
            if tc is None:
                out.append(t)
            else:
                cells.setdefault((part, tc), []).append(t)
        return out + [" ".join(v) for v in cells.values()]


class PdfText:
    def __init__(self, path):
        self.doc = pymupdf.open(str(path))
        self.n = len(self.doc)
        self.text = [self.doc[i].get_text() for i in range(self.n)]
        self._words = {}
        self.big = None

    def words(self, i):
        if i not in self._words:
            self._words[i] = self.doc[i].get_text("words")
        return self._words[i]

    def index(self):
        """Тексты листов без пробелов, склеенные в одну строку, + начала листов: поиск «на каком листе этот фрагмент»."""
        if self.big is None:
            parts = [norm(t) for t in self.text]
            self.starts = [0]
            for p in parts:
                self.starts.append(self.starts[-1] + len(p))
            self.big = "".join(parts)
        return self.big, self.starts

    def find_pages(self, frag, limit=40):
        key = norm(frag)[:24]
        if len(key) < 8:
            return []
        big, starts = self.index()
        out, pos = [], big.find(key)
        while pos != -1 and len(out) < limit:
            out.append(min(bisect.bisect_right(starts, pos) - 1, self.n - 1))
            pos = big.find(key, pos + 1)
        return out


# ---------------------------------------------------------------- 1. потеря содержимого
_HDR_SEG = re.compile(r"\s*(?:Azonos[ií]t[óo]|Verzi[óo]|Oldalsz[áa]m|Lapsz[áa]m|Megnevez[ée]s)", re.I)


def lost_segments(segs, res, full, polish=None, ocr_pages=(), tr=needs_translation):
    """Сегменты, у которых ни начала, ни конца перевода нет в тексте DOCX `full` (norm). Пропускаются колонтитулы
    (role=header, поля field — номер листа и т. п., обрывки шапки «Verziószám: …» — их сверяет проверка 3), сканы и
    надписи на скриншотах (ocr, shot — рисуются в картинке), «не требует перевода» и копии без перевода.
    polish — правки шлифовки: {id: новый текст}.
    Одинаковый венгерский текст в сборке переводится по памяти одним русским вариантом (res у повторов может быть иным),
    поэтому группа сегментов с одним исходником считается по числу вхождений любого их варианта в DOCX: потерянными
    признаются лишь те, кому вхождений не хватило."""
    groups = {}
    for s in segs:
        loc = s.get("loc") or {}
        r = res.get(str(s["id"])) or {}
        ru = (r.get("ru") or "").strip()
        if not ru or loc.get("role") == "header" or loc.get("ocr") or loc.get("shot") or loc.get("field")                 or _HDR_SEG.match(s.get("text") or "") or (s.get("page") or 0) + 1 in ocr_pages:
            continue
        if r.get("copied") or not tr(_TAG.sub("", s.get("text") or "")):
            continue
        cand = [c for c in [norm(ru)] + ([norm(polish[s["id"]])] if polish and s["id"] in polish else []) if len(c) >= 6]
        if cand:
            groups.setdefault(norm(s["text"]), []).append((s, [(c[:20], c[-14:]) for c in cand]))
    out = []
    for items in groups.values():
        own_missing = [s for s, cands in items if not any(h in full or t in full for h, t in cands)]
        if not own_missing:
            continue
        variants = {v for _, cands in items for v in cands}
        found = sum(max(full.count(h), full.count(t)) for h, t in variants)
        out += own_missing[:max(0, len(items) - found)]
    return sorted(out, key=lambda s: s["id"])


def check_lost(d):
    if not d.js or not d.docx:
        return d.skip("1 потеря содержимого", "нужны JSON задания и DOCX")
    js = d.js
    polish = {e["id"]: e["now"] for e in (js.get("polish") or {}).get("edits", []) if "id" in e and e.get("now")}
    ocr = set((js.get("info") or {}).get("ocr_pages") or [])
    lost = [s for s in lost_segments(js["segs"], js["res"], d.docx.full, polish, ocr) if not d.in_picture(s)]
    if not lost:
        return []
    ex ="; ".join("стр.%d #%s «%s»→«%s»" % ((s.get("page") or 0) + 1, s["id"], clip(s["text"], 40),
                                            clip(js["res"][str(s["id"])]["ru"], 50)) for s in lost[:SHOW])
    pg = (lost[0].get("page") or 0) + 1
    lvl = RED if len(lost) >= 3 else YELLOW
    return [Finding("1 потеря содержимого", lvl,
                    "%d сегм. с готовым переводом нет в DOCX (первые: %s%s)" % (len(lost), ex, "; …" if len(lost) > SHOW else ""),
                    orig_page=pg, frag=clip(lost[0]["text"]))]


# ---------------------------------------------------------------- 2. повтор фразы
def find_repeat(text, min_chars=10, min_reps=3, max_words=30, cap=20000):
    """Первая фраза (≥ min_chars знаков), идущая подряд ≥ min_reps раз: (фраза, повторов) или None."""
    words = [w.strip(",;.:") for w in text.split()][:cap]
    words = [w for w in words if w]
    n = len(words)
    for i in range(n):
        for p in range(1, max_words + 1):
            if i + p * min_reps > n:
                break
            if words[i:i + p] != words[i + p:i + 2 * p] or sum(map(len, words[i:i + p])) + p - 1 < min_chars:
                continue
            if sum(c.isalpha() for w in words[i:i + p] for c in w) < 3:
                continue                              # «○ ○ ○ ○ ○ ○», «0,0 0,0» — стопка значков/чисел, её ловят проверки 5–6
            reps = 2
            while words[i:i + p] == words[i + reps * p:i + (reps + 1) * p]:
                reps += 1
            if reps >= min_reps:
                return " ".join(words[i:i + p]), reps
    return None


def check_repeats(d):
    if not d.docx:
        return d.skip("2 повтор фразы", "нет DOCX")
    out, seen = [], set()
    for u in d.docx.units():
        if len(u) < 30:
            continue
        r = find_repeat(u)
        if r and r[0] not in seen:
            seen.add(r[0])
            pg = (d.pdf.find_pages(r[0]) if d.pdf else [])
            out.append(Finding("2 повтор фразы", RED, "«%s» ×%d подряд в одной ячейке/абзаце" % (clip(r[0], 70), r[1]),
                               page=pg[0] + 1 if pg else None, frag=clip(r[0])))
            if len(out) >= 40:
                break
    return out


# ---------------------------------------------------------------- 3. шапки
_HDR = {
    "hu": dict(id=re.compile(r"Azonos[ií]t[óo]\s*:?\s*([^\s,;]+)", re.I),
               ver=re.compile(r"Verzi[óo](?:sz[áa]m)?\s*:?\s*(\d+)", re.I),
               pg=re.compile(r"(?:Oldalsz[áa]m|Lapsz[áa]m)\s*:?\s*(\d+)", re.I)),
    "ru": dict(id=re.compile(r"Идентификатор\s*:?\s*([^\s,;]+)", re.I),
               ver=re.compile(r"Номер\s+версии\s*:?\s*(\d+)", re.I),
               pg=re.compile(r"(?:Номер\s+страницы|Лист)\s*:?\s*(\d+)", re.I)),
}


def header_of(text, lang, head=700):
    """(идентификатор, версия, номер листа) шапки в верху листа; None, если шапки с идентификатором и версией нет."""
    t = text[:head]
    r = _HDR[lang]
    i, v, p = r["id"].search(t), r["ver"].search(t), r["pg"].search(t)
    if not (i and v):
        return None
    return i.group(1).strip(".").upper(), int(v.group(1)), int(p.group(1)) if p else None


def runs(seq):
    out = []
    for x in seq:
        if not out or out[-1] != x:
            out.append(x)
    return out


def compare_headers(ho, hr):
    """ho/hr — списки шапок (идентификатор, версия, номер листа) оригинала и перевода по листам. Замечания (уровень, текст)."""
    out = []
    po, pr = [(i, v) for i, v, _ in ho], [(i, v) for i, v, _ in hr]
    so, sr = set(po), set(pr)
    if not so:
        return []                                      # у оригинала шапки нет в текстовом слое (титул-скан): сравнивать не с чем
    if so and not sr:
        return [(RED, "в оригинале шапка есть (%s), в переводе не найдена «Идентификатор/Номер версии» ни на одном листе" %
                 ", ".join("%s v%d" % x for x in sorted(so)[:4]))]
    for i, v in sorted(so - sr):
        n = [k + 1 for k, x in enumerate(po) if x == (i, v)]
        out.append((RED, "шапка оригинала «%s, версия %d» (листы ориг. %s) в переводе отсутствует; в переводе: %s" %
                    (i, v, ",".join(map(str, n[:8])), ", ".join("%s v%d" % x for x in sorted(sr)) or "—")))
    for i, v in sorted(sr - so):
        n = [k + 1 for k, x in enumerate(pr) if x == (i, v)]
        out.append((RED, "в переводе шапка «%s, версия %d» (листы PDF %s), которой нет в оригинале" %
                    (i, v, ",".join(map(str, n[:8])))))
    if so == sr and runs(po) != runs(pr):
        out.append((YELLOW, "пары те же, но чередование по листам иное: оригинал %s, перевод %s" %
                    (" → ".join("%s v%d" % x for x in runs(po)[:8]), " → ".join("%s v%d" % x for x in runs(pr)[:8]))))
    no = [p for _, _, p in ho if p is not None]
    nr = [p for _, _, p in hr if p is not None]
    if no and nr and abs(max(no) - max(nr)) >= 2:
        out.append((YELLOW, "номера листов: в оригинале до %d, в переводе до %d — ссылки «стр.», «Измененные страницы» могут указывать не туда" %
                    (max(no), max(nr))))
    return out


def check_headers(d):
    if not d.pdf or not d.orig:
        return d.skip("3 шапки", "нужны PDF перевода и оригинала")
    ho = [h for h in (header_of(t, "hu") for t in d.orig.text) if h]
    hr = [h for h in (header_of(t, "ru") for t in d.pdf.text) if h]
    if not ho and not hr:
        return []
    out = []
    for lvl, msg in compare_headers(ho, hr):
        out.append(Finding("3 шапки", lvl, msg))
    return out


# ---------------------------------------------------------------- 4. листы
def _hrules(page):
    """Горизонтальные линии листа [(x0, x1, y)]: тонкие прямоугольники, отрезки и верх/низ рамок без заливки; отрезки на одной
    высоте (±1 pt) сливаются — линия таблицы в PDF часто нарезана по ячейкам."""
    segs = []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l" and abs(it[1].y - it[2].y) < 0.6:
                segs.append((min(it[1].x, it[2].x), max(it[1].x, it[2].x), it[1].y))
            elif it[0] == "re":
                r = it[1]
                if r.height <= 1.6:
                    segs.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
                elif not d.get("fill") and r.height < 400:
                    segs += [(r.x0, r.x1, r.y0), (r.x0, r.x1, r.y1)]
    merged = []
    for x0, x1, y in sorted(segs, key=lambda s: s[2]):
        if merged and abs(merged[-1][2] - y) < 1.0:
            m = merged[-1]
            merged[-1] = (min(m[0], x0), max(m[1], x1), m[2], m[3] + (x1 - x0))
        else:
            merged.append((x0, x1, y, x1 - x0))
    return [(x0, x1, y) for x0, x1, y, tot in merged if x1 - x0 >= 0.3 * page.rect.width and tot >= 0.5 * (x1 - x0)]


def empty_row_share(page, words):
    """Доля площади листа, занятая строками таблицы без единого слова: полосы между соседними горизонтальными линиями
    (ловят и полностью пустую сетку) и строки, которые находит find_tables (ему нужен хоть какой-то текст в таблице)."""
    W, H = page.rect.width, page.rect.height
    empty = lambda r: not any(r.x0 <= (w[0] + w[2]) / 2 <= r.x1 and r.y0 <= (w[1] + w[3]) / 2 <= r.y1 for w in words)
    area = 0.0
    ls = _hrules(page)
    for (a0, a1, ay), (b0, b1, by) in zip(ls, ls[1:]):
        x0, x1 = max(a0, b0), min(a1, b1)
        if 8 <= by - ay <= 200 and x1 - x0 >= 0.3 * W and empty(pymupdf.Rect(x0, ay, x1, by)):
            area += (by - ay) * (x1 - x0)
    try:
        tabs = page.find_tables().tables
    except Exception:
        tabs = []
    area2 = sum(pymupdf.Rect(row.bbox).get_area() for t in tabs for row in t.rows if empty(pymupdf.Rect(row.bbox)))
    return min(1.0, max(area, area2) / (W * H))


def check_sheets(d):
    if not d.pdf:
        return d.skip("4 листы", "нет PDF перевода")
    out = []
    if d.orig and d.orig.n:
        ratio = d.pdf.n / d.orig.n
        if ratio > RATIO_SHEETS:
            out.append(Finding("4 листы", RED, "листов в переводе %d при оригинале %d (×%.2f > %.1f): ссылки на листы и «стр.» уезжают" %
                               (d.pdf.n, d.orig.n, ratio, RATIO_SHEETS)))
    wc = sorted(len(d.pdf.words(i)) for i in range(d.pdf.n))
    med = wc[len(wc) // 2] if wc else 0
    bad = []
    for i in range(d.pdf.n):
        w = d.pdf.words(i)
        if len(w) > max(120, med * 0.6) and d.pdf.n > 3:
            continue                                  # плотная страница: пустых строк на 70 % не бывает
        s = empty_row_share(d.pdf.doc[i], w)
        if s > HALF_EMPTY_SHARE:
            bad.append((i + 1, s))
    for lvl, lo, hi, what in ((RED, EMPTY_TABLE_SHARE, 2, ">%d%%" % (EMPTY_TABLE_SHARE * 100)),
                              (YELLOW, HALF_EMPTY_SHARE, EMPTY_TABLE_SHARE, "%d–%d%%" % (HALF_EMPTY_SHARE * 100, EMPTY_TABLE_SHARE * 100))):
        sel = [(p, s) for p, s in bad if lo < s <= hi]
        if sel:
            out.append(Finding("4 листы", lvl, "листы, где %s площади — пустые строки таблицы: %s" %
                               (what, ", ".join("%d (%d%%)" % (p, s * 100) for p, s in sel[:15]) +
                                (" …всего %d" % len(sel) if len(sel) > 15 else "")), page=sel[0][0]))
    return out


# ---------------------------------------------------------------- 5. коды оборудования
def codes_in(texts):
    """Counter кодов и первая страница (1-based) каждого."""
    c, first = Counter(), {}
    for i, t in enumerate(texts):
        for m in CODE_RE.findall(t):
            c[m] += 1
            first.setdefault(m, i + 1)
    return c, first


def compare_codes(co, cr):
    """(пропали, лишние, повторены лишний раз, потеряны вхождения) — словари код→число."""
    gone = {k: v for k, v in co.items() if k not in cr}
    extra = {k: v for k, v in cr.items() if k not in co}
    more = {k: cr[k] - co[k] for k in co if k in cr and cr[k] > co[k]}
    less = {k: co[k] - cr[k] for k in co if k in cr and cr[k] < co[k]}
    return gone, extra, more, less


def check_codes(d):
    if not d.pdf or not d.orig:
        return d.skip("5 коды оборудования", "нужны PDF перевода и оригинала")
    co, fo = codes_in(d.orig.text)
    cr, fr = codes_in(d.pdf.text)
    gone, extra, more, less = compare_codes(co, cr)
    n = len(gone) + len(extra) + len(more) + len(less)
    if not n:
        return []
    parts = []
    if gone:
        parts.append("пропали %d: %s" % (len(gone), ", ".join("%s (ориг. стр.%d)" % (k, fo[k]) for k in sorted(gone)[:SHOW])))
    if extra:
        parts.append("лишние %d: %s" % (len(extra), ", ".join("%s (стр.%d)" % (k, fr[k]) for k in sorted(extra)[:SHOW])))
    if more:
        parts.append("повторены %d: %s" % (len(more), ", ".join("%s +%d" % (k, more[k]) for k in sorted(more)[:SHOW])))
    if less:
        parts.append("вхождений меньше %d: %s" % (len(less), ", ".join("%s −%d" % (k, less[k]) for k in sorted(less)[:SHOW])))
    lvl = RED if n >= 3 or sum(gone.values()) + sum(extra.values()) >= 3 else YELLOW
    return [Finding("5 коды оборудования", lvl, "кодов в оригинале %d (вхождений %d), в переводе %d (%d); %s" %
                    (len(co), sum(co.values()), len(cr), sum(cr.values()), "; ".join(parts)))]


# ---------------------------------------------------------------- 6. отметки
def marks_on(words):
    return [w for w in words if w[4] in MARKS]


def glued_rows(marks, gap=8.0):
    """Строки, где две отметки подряд с зазором меньше gap pt («O O O» в одной ячейке)."""
    rows = defaultdict(list)
    for w in marks:
        rows[round((w[1] + w[3]) / 4)].append(w)
    n = 0
    for r in rows.values():
        r.sort(key=lambda w: w[0])
        if any(b[0] - a[2] < gap for a, b in zip(r, r[1:])):
            n += 1
    return n


def code_row_share(words):
    """(кодов на листах с отметками, из них с отметкой на своей строке) для одного листа."""
    marks = marks_on(words)
    codes = [w for w in words if CODE_RE.fullmatch(w[4].strip(",;.:()"))]
    if not marks or not codes:
        return 0, 0
    cy = lambda w: (w[1] + w[3]) / 2
    return len(codes), sum(1 for c in codes if any(abs(cy(m) - cy(c)) < 3.5 for m in marks))


def groups(edges, o_pages, r_pages):
    """Связные группы листов оригинала и перевода по рёбрам (o, r): листы, поделившиеся/слившиеся, считаются вместе."""
    par = {}

    def f(x):
        while par.setdefault(x, x) != x:
            par[x] = par[par[x]]
            x = par[x]
        return x
    for o, r in edges:
        par[f(("o", o))] = f(("r", r))
    g = defaultdict(lambda: ([], []))
    for o in o_pages:
        if ("o", o) in par:
            g[f(("o", o))][0].append(o)
    for r in r_pages:
        if ("r", r) in par:
            g[f(("r", r))][1].append(r)
    return [v for v in g.values() if v[0] and v[1]]


def check_marks(d):
    if not d.pdf or not d.orig:
        return d.skip("6 отметки", "нужны PDF перевода и оригинала")
    mo = [len(marks_on(d.orig.words(i))) for i in range(d.orig.n)]
    mr = [len(marks_on(d.pdf.words(i))) for i in range(d.pdf.n)]
    if not sum(mo):
        return []                                     # у оригинала отметок нет в текстовом слое (рисованные кружки)
    out = []
    to, tr = sum(mo), sum(mr)
    pm = d.page_map()
    if pm:
        gs = sorted(groups(pm, range(d.orig.n), range(d.pdf.n)), key=lambda g: min(g[0]))
        ab = [(sum(mo[i] for i in og), sum(mr[i] for i in rg)) for og, rg in gs]
        bad = []
        for k, (og, rg) in enumerate(gs):
            a, b = ab[k]
            win = ab[max(0, k - 1):k + 2]                 # переезд отметок на соседний лист при перекладке — не потеря:
            wd = abs(sum(x for x, _ in win) - sum(y for _, y in win))   # разница должна держаться и в окне из трёх групп
            if abs(a - b) >= 5 and abs(a - b) >= 0.3 * a and wd >= 0.6 * abs(a - b):
                bad.append((og, rg, a, b))
        if bad:
            lvl = RED if abs(to - tr) >= max(5, 0.05 * to) else YELLOW   # итог цел, но отметки разъехались по листам — жёлтый
            out.append(Finding("6 отметки", lvl, "число отметок ○/□/O по листам не совпало (оригинал %d, перевод %d): %s" %
                               (to, tr, "; ".join("ориг. стр.%s → стр.%s: %d→%d" % (_rng(og), _rng(rg), a, b)
                                                  for og, rg, a, b in bad[:SHOW]) + ("; …" if len(bad) > SHOW else "")),
                               page=bad[0][1][0] + 1, orig_page=bad[0][0][0] + 1))
    elif abs(to - tr) >= max(3, 0.05 * to):
        out.append(Finding("6 отметки", RED, "число отметок ○/□/O: оригинал %d, перевод %d (листы не сопоставлены, сверка по документу)" % (to, tr)))
    go = sum(glued_rows(marks_on(d.orig.words(i))) for i in range(d.orig.n))
    gl = [(i + 1, glued_rows(marks_on(d.pdf.words(i)))) for i in range(d.pdf.n)]
    g = sum(n for _, n in gl)
    if g - go >= 1:
        pages = [p for p, n in gl if n]
        out.append(Finding("6 отметки", RED if g - go >= 2 else YELLOW,
                           "слипшиеся отметки «O O O» в одной строке: в переводе %d строк на листах %s%s, в оригинале %d" %
                           (g, ",".join(map(str, pages[:12])), "…" if len(pages) > 12 else "", go), page=pages[0]))
    so = [code_row_share(d.orig.words(i)) for i in range(d.orig.n)]
    sr = [code_row_share(d.pdf.words(i)) for i in range(d.pdf.n)]
    co, ho = sum(c for c, _ in so), sum(h for _, h in so)
    cr, hr = sum(c for c, _ in sr), sum(h for _, h in sr)
    if co >= 10 and cr >= 10 and ho / co >= 0.8 and hr / cr < ho / co - 0.25:
        out.append(Finding("6 отметки", RED if hr / cr < ho / co - 0.4 else YELLOW,
                           "отметка на строке своего кода: оригинал %d%% (%d из %d), перевод %d%% (%d из %d) — отметки уехали от строк" %
                           (100 * ho // co, ho, co, 100 * hr // cr, hr, cr)))
    return out


def _rng(pages):
    p = sorted(x + 1 for x in pages)
    return str(p[0]) if len(p) == 1 else "%d–%d" % (p[0], p[-1])


# ---------------------------------------------------------------- 7. венгерские буквы
_PUNCT = ".,;:()[]{}«»\"'„“”-–—!?"


def hu_leftovers(text, allowed=frozenset()):
    """Слова с венгерскими буквами в русском тексте, не коды (нет цифр), не сокращения до 4 букв, не имена."""
    toks = text.split()
    cores = [t.strip(_PUNCT) for t in toks]
    out = []
    for i, c in enumerate(cores):
        if not c or not any(ch in HU_LETTERS for ch in c):
            continue
        if re.search(r"\d|[@/\\_]", c) or c.upper() in KNOWN_LABELS or c.lower() in allowed:
            continue
        if c.isupper() and (len(c) <= 6 or "." in c):
            continue                                   # сокращение («ÜFK», «BLÜGY», «TNÉVL»): строчной венгерщины в нём нет
        if re.search(r"[А-Яа-яЁё]", c):
            continue                                   # слитно с русским («AdorjánJános», «ÜVIGПОДГОТОВ») — мусор OCR скана, не недоперевод
        cap = lambda x: x[:1].isupper() and x[1:2].islower() and x.isalpha()
        if cap(c) and (len(cores) <= 2 or any(0 <= j < len(cores) and cap(cores[j]) for j in (i - 1, i + 1))):
            continue                                   # «Bóth Zoltán», ячейка «Péter» — имя, а не недоперевод
        out.append(c)
    return out


def check_hungarian(d):
    if not d.docx:
        return d.skip("7 венгерские буквы", "нет DOCX")
    allowed = set()
    if d.js:                                           # слова из сегментов, которые сборка оставляет как есть (имена, сокращения)
        for s in d.js["segs"]:
            r = d.js["res"].get(str(s["id"])) or {}
            if r.get("copied") or not needs_translation(_TAG.sub("", s.get("text") or "")):
                allowed |= {w.lower() for w in hu_leftovers(s.get("text") or "")}
    found = {}
    for _, t, _ in d.docx.paras:
        for w in hu_leftovers(t, allowed):
            found.setdefault(w.lower(), (w, t))
    if not found:
        return []
    ex = []
    for k, (w, t) in list(found.items())[:SHOW]:
        pg = d.pdf.find_pages(t) if d.pdf else []
        ex.append("«%s»%s в «%s»" % (w, " стр.%d" % (pg[0] + 1) if pg else "", clip(t, 60)))
    pg0 = d.pdf.find_pages(next(iter(found.values()))[1]) if d.pdf else []
    return [Finding("7 венгерские буквы", RED if len(found) >= 8 else YELLOW,
                    "%d слов с буквами ő ű á é í ó ö ú ü в тексте перевода: %s%s" % (len(found), "; ".join(ex), "; …" if len(found) > SHOW else ""),
                    page=pg0[0] + 1 if pg0 else None)]


# ---------------------------------------------------------------- 8. снимок листа вместо таблицы
def image_share(page, min_area=400):
    """Доля площади листа под картинками (без мелких значков)."""
    area = sum(pymupdf.Rect(i["bbox"]).get_area() for i in page.get_image_info() if pymupdf.Rect(i["bbox"]).get_area() >= min_area)
    return min(1.0, area / (page.rect.width * page.rect.height))


def check_images(d):
    if not d.pdf or not d.orig:
        return d.skip("8 снимок листа", "нужны PDF перевода и оригинала")
    so = [d._img("orig", i) for i in range(d.orig.n)]
    sr = [(i, d._img("ru", i)) for i in range(d.pdf.n)]
    big = [(i, s) for i, s in sr if s > BIG_IMAGE_SHARE]
    if not big:
        return []
    r2o = d.page_map(by="r2o")
    ratio = d.pdf.n / d.orig.n
    bad = []
    for i, s in big:
        if not any(x > 0.02 for x in so):               # в оригинале картинок нет вовсе
            bad.append((i, s))
            continue
        near = r2o.get(i) or {o for o in range(round(i / ratio) - 1, round(i / ratio) + 2) if 0 <= o < d.orig.n}
        if all(so[o] <= 0.02 for o in near):            # лист не сопоставлен — берём соседей по линии пропорции
            bad.append((i, s))
    if not bad:
        return []
    return [Finding("8 снимок листа", RED, "листов-картинок (>%d%% площади) там, где в оригинале изображений нет: %d — %s; текстового слоя и правки нет" %
                    (BIG_IMAGE_SHARE * 100, len(bad), ", ".join("%d (%d%%)" % (i + 1, s * 100) for i, s in bad[:15]) +
                     (" …" if len(bad) > 15 else "")), page=bad[0][0] + 1)]


CHECKS = [check_lost, check_repeats, check_headers, check_sheets, check_codes, check_marks, check_hungarian, check_images]


# ---------------------------------------------------------------- документ
class Doc:
    """Один результат: DOCX + PDF перевода + JSON сегментов + PDF оригинала (что найдено)."""

    def __init__(self, docx_path, work=None, orig_dirs=()):
        p = Path(docx_path)
        self.docx_path, self.stem = p, re.sub(r"_RU$", "", p.stem)
        self.skipped = []
        self.docx = DocxText(p)
        pdf = p.with_suffix(".pdf")
        self.pdf = PdfText(pdf) if pdf.exists() else None
        self.js, self.js_path = self._find_json(work)
        self.orig = None
        self.orig_path = self._find_orig(orig_dirs)
        if self.orig_path:
            self.orig = PdfText(self.orig_path)
        self._pm, self._imgs = None, {}

    def _find_json(self, work):
        name = self.stem + ".json"
        cands = []
        if work:
            cands.append(Path(work) / "json" / name)
        cands.append(self.docx_path.parent.parent / "json" / name)
        cands += sorted((REPO / "data" / "jobs").glob("*/work/json/" + name), key=lambda x: x.stat().st_mtime, reverse=True)
        for c in cands:
            if c.exists():
                return json.loads(c.read_text(encoding="utf-8")), c
        return None, None

    def _find_orig(self, orig_dirs):
        c = []
        f = (self.js or {}).get("file") or ""
        m = re.match(r"/app/(.*)$", f)
        if m:
            c.append(REPO / m.group(1))
        elif f:
            c.append(Path(f))
        for dd in orig_dirs:
            c += [Path(dd) / (self.stem + e) for e in (".PDF", ".pdf")]
        for x in c:
            if x.suffix.lower() == ".pdf" and x.exists():
                return x
        for pat in ("*/in/%s.[pP][dD][fF]", "*/work/unpacked/*/%s.[pP][dD][fF]"):      # вложения портфолио — в unpacked задания
            g = sorted((REPO / "data" / "jobs").glob(pat % re.sub(r"([\[\]*?])", r"[\1]", self.stem)), key=lambda x: x.stat().st_mtime, reverse=True)
            if g:
                return g[0]
        return None

    def skip(self, check, why):
        self.skipped.append("%s: пропущено — %s" % (check, why))
        return []

    def page_map(self, by="edges"):
        """Листы оригинала ↔ листы перевода по тексту сегментов: рёбра (o, r) | {r: {o}} (0-based). Пусто, если не вышло."""
        if self._pm is None:
            self._pm = self._build_page_map()
        if by in ("r2o", "o2r"):
            out = defaultdict(set)
            for o, r in self._pm:
                out[r if by == "r2o" else o].add(o if by == "r2o" else r)
            return out
        return self._pm

    def _build_page_map(self):
        """Голосование: сегмент с различимым текстом (на ≤3 листах перевода) голосует за свои листы; лист перевода
        берётся, если собрал голосов не меньше 1 и 30 % от лучшего и лежит у линии пропорции (оригинал→перевод)."""
        if not (self.js and self.pdf and self.orig):
            return []
        votes = defaultdict(Counter)
        for s in self.js["segs"]:
            loc = s.get("loc") or {}
            if loc.get("role") == "header" or loc.get("ocr") or loc.get("shot"):
                continue
            ru = (self.js["res"].get(str(s["id"])) or {}).get("ru") or ""
            pg = sorted(set(self.pdf.find_pages(ru, limit=8)))
            if not pg or len(pg) > 3:
                continue
            for p in pg:
                votes[min(s.get("page") or 0, self.orig.n - 1)][p] += 1 / len(pg)
        ratio = self.pdf.n / self.orig.n
        edges = []
        for o, c in votes.items():
            top = max(c.values())
            edges += [(o, r) for r, v in c.items()
                      if v >= max(1.0, 0.3 * top) and abs(r - o * ratio) <= 0.25 * self.pdf.n + 3]
        mapped = {o for o, _ in edges}
        return edges if len(mapped) >= 0.6 * self.orig.n else []

    def _img(self, which, i):
        pdf = self.pdf if which == "ru" else self.orig
        key = (which, i)
        if key not in self._imgs:
            self._imgs[key] = image_share(pdf.doc[i])
        return self._imgs[key]

    def in_picture(self, s):
        """Текст сегмента — внутри картинки-листа: страница оригинала с изображением, а соответствующий лист перевода
        почти целиком картинка (надписи нарисованы на ней — потерю по тексту DOCX не засчитываем; снимок ловит проверка 8)."""
        if not self.orig:
            return False
        o = min(s.get("page") or 0, self.orig.n - 1)
        if self._img("orig", o) <= 0.02:
            return False
        c = round(o * self.pdf.n / self.orig.n)         # листы-картинки не сопоставляются по тексту — берём ещё и соседей по пропорции
        rs = set(self.page_map(by="o2r").get(o, ())) | {r for r in range(c - 1, c + 2) if 0 <= r < self.pdf.n}
        return any(self._img("ru", r) > BIG_IMAGE_SHARE for r in rs)

    def run(self):
        out = []
        for chk in CHECKS:
            try:
                out += chk(self)
            except Exception as e:                      # проверка упала — это тоже сигнал, а не тишина
                out.append(Finding(chk.__name__, YELLOW, "проверка не выполнена: %s: %s" % (type(e).__name__, e)))
        return out


# ---------------------------------------------------------------- отчёт
def collect(targets, only):
    files = []
    for t in targets:
        p = Path(t)
        files += [p] if p.is_file() else sorted(p.rglob("*_RU.docx"))
    files = [f for f in files if not f.name.startswith("~$") and (not only or any(o.lower() in f.name.lower() for o in only))]
    return files


def render(d, findings):
    head = "══ %s  (перевод: %s л., оригинал: %s л.%s)" % (
        d.stem, d.pdf.n if d.pdf else "—", d.orig.n if d.orig else "—", "" if d.js else "; JSON задания не найден")
    lines = [head]
    for f in sorted(findings, key=lambda x: (x.level != RED, x.check)):
        pg = " [стр. %s%s]" % (f.page or "—", ", ориг. %s" % f.orig_page if f.orig_page else "") if f.page or f.orig_page else ""
        lines.append("  %s %s%s: %s" % (ICON[f.level], f.check, pg, f.msg))
    if not findings:
        lines.append("  ✔ замечаний нет")
    for s in d.skipped:
        lines.append("  · " + s)
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Ворота качества готовых переводов (см. docstring модуля)")
    ap.add_argument("targets", nargs="+", help="папка с *_RU.docx (ищется рекурсивно) или сами файлы")
    ap.add_argument("--work", help="рабочая папка задания (data/jobs/<uid>/work): json/ внутри")
    ap.add_argument("--orig", action="append", default=[], help="папка с оригиналами (можно несколько)")
    ap.add_argument("--only", action="append", default=[], help="только документы, в имени которых есть эта строка")
    ap.add_argument("--json", help="записать результат в JSON")
    ap.add_argument("--fail-on", choices=["red", "yellow", "never"], default="red")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    files = collect(a.targets, a.only)
    if not files:
        print("Нечего проверять: *_RU.docx не найдены")
        return 2
    allf, summary = {}, []
    for f in files:
        d = Doc(f, a.work, a.orig)
        fs = d.run()
        allf[d.stem] = [x.as_dict() for x in fs]
        print(render(d, fs))
        summary.append((d.stem, sum(x.level == RED for x in fs), sum(x.level == YELLOW for x in fs)))
    nr, ny = sum(s[1] for s in summary), sum(s[2] for s in summary)
    print("\nИТОГО: документов %d, %s %d, %s %d" % (len(summary), ICON[RED], nr, ICON[YELLOW], ny))
    for stem, r, y in summary:
        print("  %-60s %s %2d  %s %2d" % (stem[:60], ICON[RED], r, ICON[YELLOW], y))
    if a.json:
        Path(a.json).write_text(json.dumps(allf, ensure_ascii=False, indent=1), encoding="utf-8")
    if a.fail_on == "never":
        return 0
    return 1 if nr or (a.fail_on == "yellow" and ny) else 0


if __name__ == "__main__":
    sys.exit(main())
