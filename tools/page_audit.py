# -*- coding: utf-8 -*-
"""Постраничная сверка пересборки DOCX→PDF с оригиналом: что потерялось, что задвоилось, где картинки и векторная
графика оригинала, пустые листы. Листы сопоставляются по тексту (монотонно), у перевода — по словам с цифрами (коды,
номера, значения переводом не меняются). Плюс сводные листы пар «оригинал | пересборка» для просмотра глазами.

  python tools/page_audit.py <оригинал.pdf> <пересборка.pdf> <папка отчёта> [--codes] [--sheets]

--codes  — пересборка переведена: сверка только по словам с цифрами.
--sheets — сводки пар по 4 на картинку (sheet_NNN.png) с номерами листов и отметками.
Отчёт — audit.md (отмеченные листы) и audit.json (все листы)."""
import json
import re
import sys
from collections import Counter
from pathlib import Path

import fitz

WORD = re.compile(r"[^\W_]+", re.U)


def page_words(page, codes=False):
    t = page.get_text("text", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_DEHYPHENATE)
    ws = [w.lower() for w in WORD.findall(t) if len(w) > 1 or w.isdigit()]
    return [w for w in ws if re.search(r"\d", w)] if codes else ws


def colors(page):
    """Цвета текста, кроме чёрного и серого (r≈g≈b): {(r, g, b) грубо}."""
    out = set()
    for bl in page.get_text("dict")["blocks"]:
        for ln in bl.get("lines", []):
            for sp in ln["spans"]:
                c = sp["color"]
                r, g, b = c >> 16 & 255, c >> 8 & 255, c & 255
                if sp["text"].strip() and max(r, g, b) - min(r, g, b) > 40:
                    out.add((r // 64, g // 64, b // 64))
    return out


def graphics(page):
    """Картинки (площадь, pt²) и векторные фигуры оригинала: не тонкие линии, не белое, не во весь лист."""
    W, H = page.rect.width, page.rect.height
    img = sum(fitz.Rect(i["bbox"]).get_area() for i in page.get_image_info()
              if fitz.Rect(i["bbox"]).get_area() >= 400 and fitz.Rect(i["bbox"]).get_area() < 0.9 * W * H)
    shapes = 0
    for d in page.get_drawings():
        r = d["rect"]
        white = d.get("fill") and min(d["fill"]) > 0.95 and not d.get("color")
        if r.width > 3 and r.height > 3 and r.width < 0.8 * W and r.height < 0.8 * H and not white:
            shapes += 1
    return img, shapes


def align(A, B):
    """Лист оригинала → лист пересборки с наибольшей долей общих слов, монотонно (окно 8 листов вперёд)."""
    out, j = [], 0
    for ws in A:
        s = set(ws)
        if len(s) < 5:
            out.append(j)
            continue
        # доля общих слов, со штрафом за прыжок от ожидаемого соседнего листа: у ÜFK листы с одинаковыми рубриками
        # («Korlátozó feltétel», «Érvényesség») уводили сопоставление на 10 листов вперёд
        exp = min(j + 1, len(B) - 1) if out else 0
        score = lambda k: len(s & set(B[k])) / len(s) - 0.04 * abs(k - exp)
        best = max(range(max(0, j - 1), min(j + 8, len(B))), key=score, default=j)
        if len(s & set(B[best])) >= 0.2 * len(s):
            j = best
        out.append(j)
    return out


def audit(orig, built, codes=False):
    a, b = fitz.open(orig), fitz.open(built)
    A = [page_words(p, codes) for p in a]
    B = [page_words(p, codes) for p in b]
    # слова колонтитулов (на половине листов и чаще) — не в счёт: листов у пересборки другое число
    run = {w for w, n in Counter(w for ws in A for w in set(ws)).items() if n >= max(3, 0.5 * len(A))}
    A = [[w for w in ws if w not in run] for ws in A]
    B = [[w for w in ws if w not in run] for ws in B]
    m = align(A, B)
    miss_all = Counter(w for ws in A for w in ws) - Counter(w for ws in B for w in ws)
    extra_all = Counter(w for ws in B for w in ws) - Counter(w for ws in A for w in ws)
    rows = []
    for i, ws in enumerate(A):
        lo = max(0, m[i] - 1)
        hi = min(len(B) - 1, (m[i + 1] if i + 1 < len(m) else len(B) - 1) + 1)
        near = Counter(w for k in range(lo, hi + 1) for w in B[k])
        local = Counter(ws) - near
        lost = [w for w, n in local.items() for _ in range(min(n, miss_all[w]))]
        img_a, shp_a = graphics(a[i])
        img_b = sum(graphics(b[k])[0] for k in range(lo, hi + 1))
        flags = []
        if lost:                                            # любой потерянный кусок — на просмотр
            flags.append(f"нет {len(lost)} слов: {' '.join(lost[:12])}")
        ca = colors(a[i])
        cb = set().union(*(colors(b[k]) for k in range(lo, hi + 1))) if hi >= lo else set()
        if ca - cb:
            flags.append(f"цвет текста оригинала не перенесён: {sorted(ca - cb)}")
        if img_a >= 2000 and img_b == 0:
            flags.append(f"картинки оригинала ({img_a:.0f} pt²) — в пересборке рядом нет")
        if shp_a >= 15 and img_b == 0:
            flags.append(f"векторная графика ({shp_a} фигур) — в пересборке рядом нет картинки")
        rows.append({"orig": i + 1, "built": m[i] + 1, "words": len(ws), "lost": lost[:40], "flags": flags})
    mapped = set(m)
    for k in range(len(B)):
        extra = [w for w in B[k] if extra_all[w] > 0]
        for w in extra:
            extra_all[w] -= 1
        flags = []
        if len(extra) >= 12 and len(extra) > 0.3 * max(len(B[k]), 1):
            flags.append(f"лишние {len(extra)} слов: {' '.join(extra[:12])}")
        img_b, _ = graphics(b[k])
        if len(B[k]) < 3 and img_b == 0 and k not in (0, len(B) - 1):
            flags.append("пустой лист")
        rb = b[k].rect
        out = [w for w in b[k].get_text("words") if w[2] > rb.width + 1 or w[3] > rb.height + 1]
        if out:
            flags.append(f"текст за краем листа: {' '.join(w[4] for w in out[:6])}")
        if flags:
            rows.append({"built_only": k + 1, "mapped": k in mapped, "flags": flags})
    return rows, m, len(a), len(b)


def sheets(orig, built, m, rows, folder, dpi=40):
    """Пары «оригинал | пересборка» по 4 на картинку; лист пересборки без пары — с пустым местом слева."""
    from PIL import Image, ImageDraw
    a, b = fitz.open(orig), fitz.open(built)
    flagged_o = {r["orig"] for r in rows if r.get("flags") and "orig" in r}
    flagged_b = {r["built_only"] for r in rows if "built_only" in r}
    pairs, used = [], set()
    for i, j in enumerate(m):
        for k in range(max(used, default=-1) + 1, j):          # листы пересборки между парами — тоже
            pairs.append((None, k))
            used.add(k)
        pairs.append((i, j))
        used.add(j)
    for k in range(max(used, default=-1) + 1, len(b)):
        pairs.append((None, k))

    def img(doc, n):
        pix = doc[n].get_pixmap(dpi=dpi)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples) if pix.n == 3 else \
            Image.frombytes("RGB", (pix.width, pix.height), fitz.Pixmap(fitz.csRGB, pix).samples)
    w0 = int(a[0].rect.width * dpi / 72) + 4
    h0 = int(max(a[0].rect.height, b[0].rect.height) * dpi / 72) + 18
    n = 0
    for s in range(0, len(pairs), 4):
        sheet = Image.new("RGB", (4 * w0 + 12, 2 * h0), "white")
        d = ImageDraw.Draw(sheet)
        for q, (i, j) in enumerate(pairs[s:s + 4]):
            x, y = (q % 2) * (2 * w0 + 12), (q // 2) * h0
            if i is not None:
                sheet.paste(img(a, i), (x, y + 16))
            sheet.paste(img(b, j), (x + w0, y + 16))
            mark = "!" if (i is not None and i + 1 in flagged_o) or j + 1 in flagged_b else ""
            d.text((x + 2, y + 2), f"{mark}o{i + 1 if i is not None else '-'} | r{j + 1}", fill=(200, 0, 0) if mark else (0, 0, 0))
            d.line([(x + w0 - 2, y), (x + w0 - 2, y + h0)], fill=(180, 180, 180))
        sheet.save(Path(folder) / f"sheet_{n:03d}.png")
        n += 1
    return n


def main():
    args = [x for x in sys.argv[1:] if not x.startswith("--")]
    codes, want = "--codes" in sys.argv, "--sheets" in sys.argv
    orig, built, folder = args[0], args[1], Path(args[2])
    folder.mkdir(parents=True, exist_ok=True)
    rows, m, na, nb = audit(orig, built, codes)
    (folder / "audit.json").write_text(json.dumps({"map": m, "rows": rows}, ensure_ascii=False, indent=1),
                                       encoding="utf-8")
    flagged = [r for r in rows if r.get("flags")]
    lines = [f"# {Path(orig).name}: {na} листов → {nb}", "", f"отмечено {len(flagged)}", ""]
    for r in flagged:
        where = f"о{r['orig']} → п{r['built']}" if "orig" in r else f"п{r['built_only']}"
        lines.append(f"- {where}: " + "; ".join(r["flags"]))
    (folder / "audit.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    n = sheets(orig, built, m, rows, folder) if want else 0
    print(f"{Path(orig).name}: {na} → {nb} листов, отмечено {len(flagged)}, сводок {n}")


if __name__ == "__main__":
    main()
