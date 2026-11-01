# -*- coding: utf-8 -*-
"""«Галерея страниц»: фиксированный набор ~40 страниц корпуса (tools/gallery_pages.yaml, по семи классам дефектов из
work/audit_2026-10-29/visual/CATALOG.md) в виде картинок «оригинал | A | B» — смотреть вёрстку глазами до и после каждого
этапа единой модели, а не только по числам ворот.

КАК ЗАПУСКАТЬ (хост, без Docker; сборки A и B — папки/метки `tools/gate_baseline.py build`, либо `база` = final_branch):

  до этапа / проверка набора:    python -X utf8 tools/gallery.py база --out work/gallery/base
  после этапа (до | после):      python -X utf8 tools/gallery.py before_E3 after_E3 --out work/gallery/E3
  один класс или документ:       python -X utf8 tools/gallery.py база after_E3 --out work/gallery/E3 --only 3PR10 --only класс:2

Результат в --out: `class<N>_<имя>_<k>.png` — листы-сводки по классам (до 3 пар на картинку), `pages/NN_<документ>_p<стр>.png` —
каждая пара отдельно (крупнее) и `index.md` — что где смотреть. Колонки: оригинал | A | B (B можно не давать — тогда две).

Как подбираются страницы. В yaml номер страницы — страница ПЕРЕВОДА в базовой сборке (`база`). Страница оригинала берётся по
совпадению кодов оборудования (как в work/audit_2026-10-29/visual/pairs.py: коэффициент Жаккара кодов страницы перевода и каждой
страницы оригинала) либо задана вручную `orig_page`. Страница сборки A/B — та, чей текст (слова и коды) ближе всего к базовой
странице (тот же русский текст, другая вёрстка). Подбор с малой уверенностью (мало общих кодов/слов, у сканов — по пропорции)
помечен в подписи красным «НЕНАДЁЖНО» и в index.md: выводы по такой паре делать нельзя — только вид страницы."""
import argparse
import json
import re
import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
import gate_baseline as G  # noqa: E402

pymupdf.TOOLS.mupdf_display_errors(False)
PAGES_FILE = REPO / "tools" / "gallery_pages.yaml"
CLASSES = {1: ("fields", "поля, листы, пустышки"), 2: ("rows", "таблицы: ряды и столбцы"), 3: ("headers", "шапки и номера листов"),
           4: ("frames", "рамки, фон, списки"), 5: ("marks", "значки и отметки"), 6: ("schemes", "схемы и сканы"),
           7: ("portfolio", "портфолио")}
PER_SHEET = 3                                   # пар на листе-сводке
CELL_W, CELL_H = 560, 760                       # ячейка страницы, px
MIN_CODES = 3                                   # кодов на странице, чтобы подбор по кодам считался осмысленным
MIN_J_ORIG = 0.25                               # Жаккар кодов страницы перевода и оригинала: меньше — «ненадёжно»
MIN_J_BUILD = 0.30                              # то же для слов+кодов базовой страницы и страницы сборки A/B
CODE = re.compile(r"[A-Z0-9]{2,}[A-Z]{1,3}\d{2,}[A-Z0-9]*|\d{2,}[A-Z]{2,}\d{2,}[A-Z0-9]*|\d+[,.]\d+")
WORD = re.compile(r"[^\W\d_]{5,}", re.U)
SHEET = re.compile(r"(?:Oldalsz[aá]m|Lapsz[aá]m|Номер\s+страницы)\s*:?\s*(\d+)", re.I)
RED, GREY, BLACK = (200, 30, 30), (90, 90, 90), (0, 0, 0)


# ---------------------------------------------------------------- список страниц
def load_pages(path=PAGES_FILE):
    try:
        import yaml
    except ImportError:
        raise SystemExit("нужен PyYAML: pip install pyyaml")
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    rows = data.get("pages") if isinstance(data, dict) else data
    if not rows:
        raise SystemExit("в %s нет списка pages" % path)
    out = []
    for i, r in enumerate(rows, 1):
        if not isinstance(r, dict) or not r.get("doc") or not isinstance(r.get("page"), int) or r["page"] < 1:
            raise SystemExit("%s, элемент %d: нужны doc и page (целое ≥ 1): %r" % (path, i, r))
        if r.get("cls") not in CLASSES:
            raise SystemExit("%s, элемент %d (%s): cls — число 1–7, не %r" % (path, i, r["doc"], r.get("cls")))
        if not str(r.get("look") or "").strip():
            raise SystemExit("%s, элемент %d (%s стр. %s): не сказано, что смотреть (look)" % (path, i, r["doc"], r["page"]))
        out.append(dict(doc=str(r["doc"]), page=r["page"], cls=r["cls"], look=str(r["look"]).strip(), orig_page=r.get("orig_page")))
    return out


# ---------------------------------------------------------------- подбор страниц
def jaccard(a, b):
    return len(a & b) / len(a | b) if a and b else 0.0


def codes_of(page):
    return set(CODE.findall(page.get_text()))


def sheet_no(page):
    """Номер листа из шапки («Oldalszám: 9» / «Номер страницы: 9»); None, если шапки нет."""
    m = SHEET.search(page.get_text()[:1200])
    return int(m.group(1)) if m else None


def words_of(page):
    t = page.get_text()
    return set(CODE.findall(t)) | {w.lower() for w in WORD.findall(t)}


def best_index(target, cands, hint, min_j):
    """Лучший кандидат по Жаккару; при нулевом пересечении — по пропорции (hint). → (индекс, J, откуда, надёжно)."""
    best, bi = 0.0, None
    for i, c in enumerate(cands):
        j = jaccard(target, c)
        if j > best or (j == best and bi is not None and j > 0 and abs(i - hint) < abs(bi - hint)):
            best, bi = j, i
    if bi is not None and best > 0:
        return bi, best, "по совпадению", best >= min_j
    return max(0, min(len(cands) - 1, hint)), 0.0, "по пропорции", False


def orig_for(base_page, base_idx, base_n, orig_doc, pinned=None):
    """Страница оригинала для страницы перевода базовой сборки: (индекс, подпись, надёжно). Сначала коды оборудования; когда их
    мало или совпадение слабое — номер листа из шапки (в базовой сборке он как у оригинала), если он указывает ровно на один лист."""
    if pinned:
        return pinned - 1, "задана вручную", True
    if len(orig_doc) == 1:
        return 0, "единственная страница оригинала", True
    target = codes_of(base_page)
    hint = round(base_idx * len(orig_doc) / max(1, base_n))
    i, j, how, ok = best_index(target, [codes_of(p) for p in orig_doc], hint, MIN_J_ORIG)
    if len(target) < MIN_CODES:
        ok = False
    if not ok and sheet_no(base_page) is not None:
        hits = [k for k, p in enumerate(orig_doc) if sheet_no(p) == sheet_no(base_page)]
        if len(hits) == 1:                              # номера листов в базе местами сдвинуты и листы переливаются: надёжно, лишь если есть общий код
            return hits[0], "по номеру листа %d" % sheet_no(base_page), jaccard(target, codes_of(orig_doc[hits[0]])) > 0
    return i, ("по кодам, J=%.2f, кодов %d" % (j, len(target))) if how == "по совпадению" else "по пропорции", ok


def page_in_build(base_page, base_idx, base_n, doc):
    """Страница сборки по тексту базовой страницы: (индекс, подпись, надёжно)."""
    if len(doc) == 1:
        return 0, "единственная страница", True
    hint = round(base_idx * len(doc) / max(1, base_n))
    target = words_of(base_page)
    i, j, how, ok = best_index(target, [words_of(p) for p in doc], hint, MIN_J_BUILD)
    if len(target) < MIN_CODES:
        ok = False
    return i, ("по тексту, J=%.2f" % j) if how == "по совпадению" else "по пропорции", ok


# ---------------------------------------------------------------- файлы
def pdfs_of(build_dir):
    return {re.sub(r"_(RU|EN|HU)$", "", p.stem): p for p in Path(build_dir).rglob("*.pdf")}


def resolve_stem(name, pdfs):
    """Полное имя документа сборки по имени из списка: точное, иначе единственное, начинающееся с него."""
    if name in pdfs:
        return name
    hit = sorted(s for s in pdfs if s.startswith(name))
    if len(hit) == 1:
        return hit[0]
    raise SystemExit("документ «%s»: %s" % (name, "под него подходит несколько: " + ", ".join(hit[:6]) if hit else "нет PDF в сборке"))


def find_orig(root, stem, orig_dir, json_by_stem):
    """PDF оригинала: поле `file` JSON корпуса (/app/data/ → data/ основного репозитория), иначе <orig_dir>/<имя>.PDF."""
    j = json_by_stem.get(stem)
    if j:
        f = json.loads(Path(j).read_text(encoding="utf-8")).get("file") or ""
        m = re.match(r"/app/(.*)$", f)
        p = Path(root) / m.group(1) if m else Path(f)
        if p.suffix.lower() == ".pdf" and p.exists():
            return p
    for ext in (".PDF", ".pdf"):
        if (Path(orig_dir) / (stem + ext)).exists():
            return Path(orig_dir) / (stem + ext)
    return None


# ---------------------------------------------------------------- рисунок
def font(size):
    for f in (r"C:\Windows\Fonts\arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/Library/Fonts/Arial.ttf"):
        if Path(f).exists():
            return ImageFont.truetype(f, size)
    return ImageFont.load_default()


def render_page(page, w=CELL_W, h=CELL_H):
    r = page.rect
    z = min(w / r.width, h / r.height)
    pm = page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
    return Image.frombytes("RGB", (pm.width, pm.height), pm.samples)


def wrap(draw, text, fnt, width):
    lines, cur = [], ""
    for word in text.split():
        t = (cur + " " + word).strip()
        if draw.textlength(t, font=fnt) <= width or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = word
    return lines + ([cur] if cur else [])


def compose_row(row, n, cells):
    """Картинка пары: заголовок (номер, класс, документ, что смотреть) и ячейки [(подпись, надёжно, страница|None), ...]."""
    fb, fs = font(15), font(13)
    gap, W = 10, len(cells) * CELL_W + (len(cells) + 1) * 10
    probe = ImageDraw.Draw(Image.new("RGB", (W, 10)))
    head = "№%02d · класс %d (%s) · %s, стр. %d" % (n, row["cls"], CLASSES[row["cls"]][1], row["stem"], row["page"])
    look = wrap(probe, "Смотреть: " + row["look"], fs, W - 2 * gap)
    top = 8 + 20 + 16 * len(look) + 6
    im = Image.new("RGB", (W, top + 36 + CELL_H + gap), "white")
    d = ImageDraw.Draw(im)
    d.text((gap, 6), head, fill=BLACK, font=fb)
    for k, ln in enumerate(look):
        d.text((gap, 28 + 16 * k), ln, fill=GREY, font=fs)
    for k, (cap, ok, page) in enumerate(cells):
        x = gap + k * (CELL_W + gap)
        lines = wrap(d, cap, fs, CELL_W)
        for m, ln in enumerate(lines[:2]):
            d.text((x, top + 16 * m), ln, fill=BLACK, font=fs)
        if not ok:
            d.text((x, top + 32), "НЕНАДЁЖНО", fill=RED, font=fb)
        if page is not None:
            pg = render_page(page)
            im.paste(pg, (x, top + 36))
            d.rectangle([x, top + 36, x + pg.width - 1, top + 36 + pg.height - 1], outline=(150, 150, 150))
    return im


def stack(images, gap=12):
    W = max(i.width for i in images)
    im = Image.new("RGB", (W, sum(i.height for i in images) + gap * (len(images) - 1)), (235, 235, 235))
    y = 0
    for i in images:
        im.paste(i, (0, y))
        y += i.height + gap
    return im


# ---------------------------------------------------------------- галерея
def build_row(row, n, root, base_pdfs, builds, orig_dir, json_by_stem, cache):
    """Пара для строки списка: (описание строки, картинка, замечания). builds — [(имя, {stem: путь})] для A и B."""
    stem = resolve_stem(row["doc"], base_pdfs)
    opened = lambda p: cache.setdefault(str(p), pymupdf.open(str(p)))
    base = opened(base_pdfs[stem])
    if row["page"] > len(base):
        raise SystemExit("%s: в базовой сборке %d листов, а в списке стр. %d" % (stem, len(base), row["page"]))
    bp = base[row["page"] - 1]
    cells, notes = [], []
    op = find_orig(root, stem, orig_dir, json_by_stem)
    info = dict(row, stem=stem)
    if op:
        od = opened(op)
        oi, how, ok = orig_for(bp, row["page"] - 1, len(base), od, row.get("orig_page"))
        cells.append(("ОРИГИНАЛ стр. %d/%d (%s)" % (oi + 1, len(od), how), ok, od[oi]))
        info.update(orig_page_found=oi + 1, orig_how=how, orig_ok=ok)
        if not ok:
            notes.append("оригинал: " + how)
    else:
        cells.append(("ОРИГИНАЛ не найден", False, None))
        notes.append("оригинал не найден")
    for name, pdfs in builds:
        p = pdfs.get(stem)
        if p is None:
            cells.append(("%s: нет документа" % name, False, None))
            notes.append("%s: нет документа" % name)
            continue
        dd = opened(p)
        if Path(p) == Path(base_pdfs[stem]):
            i, how, ok = row["page"] - 1, "страница из списка", True
        else:
            i, how, ok = page_in_build(bp, row["page"] - 1, len(base), dd)
        cells.append(("%s стр. %d/%d (%s)" % (name, i + 1, len(dd), how), ok, dd[i]))
        if not ok:
            notes.append("%s: %s" % (name, how))
    info["notes"] = notes
    return info, compose_row(info, n, cells), notes


def slug(s):
    return re.sub(r"[^\w.+()\-]+", "_", s, flags=re.A)[:40]


def make_gallery(names, out_dir, rows, root, orig_dir, say=print):
    """names — [A] или [A, B] (папки/метки сборок). Пишет листы-сводки, пары и index.md; возвращает (пути листов, число «ненадёжно»)."""
    base_pdfs = pdfs_of(G.base_dir())
    builds = [(Path(n).name if Path(n).exists() else n, pdfs_of(G.locate_build(n))) for n in names]
    json_by_stem = {f.stem: f for f in G.corpus(root)}
    out = Path(out_dir)
    (out / "pages").mkdir(parents=True, exist_ok=True)
    for old in list(out.glob("class*.png")) + list((out / "pages").glob("*.png")):
        old.unlink()
    cache, made = {}, []
    for n, row in enumerate(rows, 1):
        info, im, notes = build_row(row, n, root, base_pdfs, builds, orig_dir, json_by_stem, cache)
        f = out / "pages" / ("%02d_%s_p%d.png" % (n, slug(info["stem"]), row["page"]))
        im.save(f)
        made.append((info, f, im))
        say("%02d класс %d  %-44s стр. %-3d%s" % (n, row["cls"], info["stem"][:44], row["page"], "  НЕНАДЁЖНО: " + "; ".join(notes) if notes else ""))
    sheets = []
    for cls, (sl, title) in CLASSES.items():
        group = [m for m in made if m[0]["cls"] == cls]
        for k in range(0, len(group), PER_SHEET):
            f = out / ("class%d_%s_%d.png" % (cls, sl, k // PER_SHEET + 1))
            stack([m[2] for m in group[k:k + PER_SHEET]]).save(f)
            sheets.append(f)
    write_index(out, made, sheets, [n for n, _ in builds])
    return sheets, sum(1 for m in made if m[0]["notes"])


def write_index(out, made, sheets, names):
    L = ["# Галерея страниц", "", "Колонки: оригинал | %s. Страницы списка — страницы перевода базовой сборки (`база`)." % " | ".join(names), ""]
    for cls, (_, title) in CLASSES.items():
        grp = [(i, f) for i, f, _ in made if i["cls"] == cls]
        if not grp:
            continue
        L += ["## Класс %d. %s" % (cls, title), "", "| № | документ | стр. перевода | стр. оригинала | что смотреть | картинка |", "|---|---|---|---|---|---|"]
        for i, f in grp:
            num = f.name.split("_")[0]
            L.append("| %s | %s | %d | %s | %s | pages/%s |" % (num, i["stem"], i["page"],
                                                              ("%s%s" % (i.get("orig_page_found", "—"), "" if i.get("orig_ok", True) else " ⚠ ненадёжно")),
                                                              i["look"].replace("|", "/"), f.name))
        L.append("")
    L += ["Листы-сводки: " + ", ".join(s.name for s in sheets), ""]
    (out / "index.md").write_text("\n".join(L), encoding="utf-8")


def filter_rows(rows, only):
    """--only: часть имени документа или `класс:N`."""
    if not only:
        return rows
    keep = []
    for r in rows:
        for o in only:
            m = re.fullmatch(r"класс:(\d)", o)
            if (m and r["cls"] == int(m.group(1))) or (not m and o.lower() in r["doc"].lower()):
                keep.append(r)
                break
    return keep


def main(argv=None):
    ap = argparse.ArgumentParser(description="Галерея страниц «оригинал | A | B» (см. docstring модуля)")
    ap.add_argument("builds", nargs="+", help="сборка A и, по желанию, B: папка, метка из work/gate или `база`")
    ap.add_argument("--out", required=True, help="папка результата")
    ap.add_argument("--list", default=str(PAGES_FILE), help="список страниц (yaml)")
    ap.add_argument("--orig", default=str(G.ORIG_DIR), help="папка с оригиналами")
    ap.add_argument("--only", action="append", default=[], help="часть имени документа или класс:N")
    a = ap.parse_args(argv)
    if len(a.builds) > 2:
        raise SystemExit("не больше двух сборок (A и B)")
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    rows = filter_rows(load_pages(a.list), a.only)
    if not rows:
        raise SystemExit("под --only не подошла ни одна строка списка")
    sheets, bad = make_gallery(a.builds, a.out, rows, G.data_root(), a.orig)
    print("\nСтраниц: %d, ненадёжных подборов: %d. Листы-сводки:" % (len(rows), bad))
    for s in sheets:
        print("  " + str(s))
    print("Каждая пара отдельно: %s; что где смотреть: %s" % (Path(a.out) / "pages", Path(a.out) / "index.md"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
