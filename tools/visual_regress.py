# Визуальная регрессия (договор правил, шаг 1): две сборки одних документов — по листам, что изменилось.
#   python tools/visual_regress.py <папка A> <папка B> <выход> [--dpi 30] [--pairs 4]
# PDF с одинаковым относительным путём в A и B сравниваются по листам (оттенки серого, низкое разрешение): число листов,
# доля изменившихся, первый изменившийся лист. В <выход>/report.md — таблица; для каждого изменившегося документа —
# картинка пар «A | B» первых изменившихся листов (<выход>/<имя>.png) — смотреть глазами, а не верить числу.
import argparse
import sys
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw


def pages(path, dpi):
    out = []
    with fitz.open(str(path)) as d:
        for pg in d:
            pix = pg.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
            out.append(np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width].copy())
    return out


def differ(a, b, thr=0.004):
    """Листы различаются: размеры разные или доля пикселей, отличных больше чем на 64 уровня, выше thr."""
    if a.shape != b.shape:
        return True
    return float((np.abs(a.astype(np.int16) - b.astype(np.int16)) > 64).mean()) > thr


def sheet(pa, pb, idx, out):
    tiles = []
    for i in idx:
        ims = [Image.fromarray(p[i]) if i < len(p) else Image.new("L", (10, 10), 128) for p in (pa, pb)]
        h = max(im.height for im in ims)
        t = Image.new("L", (sum(im.width for im in ims) + 6, h), 90)
        t.paste(ims[0], (0, 0)); t.paste(ims[1], (ims[0].width + 6, 0))
        ImageDraw.Draw(t).text((3, 3), str(i + 1), fill=0)
        tiles.append(t)
    w = sum(t.width for t in tiles) + 10 * len(tiles)
    s = Image.new("L", (w, max(t.height for t in tiles)), 60)
    x = 0
    for t in tiles:
        s.paste(t, (x, 0)); x += t.width + 10
    s.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("a"); ap.add_argument("b"); ap.add_argument("out")
    ap.add_argument("--dpi", type=int, default=30); ap.add_argument("--pairs", type=int, default=4)
    o = ap.parse_args()
    A, B, out = Path(o.a), Path(o.b), Path(o.out)
    out.mkdir(parents=True, exist_ok=True)
    rows = ["| документ | листов A → B | изменилось | первый | пары |", "|---|---|---|---|---|"]
    for fa in sorted(A.rglob("*.pdf")):
        rel = fa.relative_to(A)
        fb = B / rel
        if not fb.exists():
            rows.append(f"| {rel} | нет в B | | | |")
            continue
        pa, pb = pages(fa, o.dpi), pages(fb, o.dpi)
        diff = [i for i in range(max(len(pa), len(pb))) if i >= len(pa) or i >= len(pb) or differ(pa[i], pb[i])]
        img = ""
        if diff:
            name = str(rel).replace("/", "_").replace("\\", "_")[:-4] + ".png"
            sheet(pa, pb, diff[:o.pairs], out / name)
            img = name
        rows.append(f"| {rel} | {len(pa)} → {len(pb)} | {len(diff)} | {diff[0] + 1 if diff else ''} | {img} |")
        print(rows[-1], flush=True)
    (out / "report.md").write_text("\n".join(rows) + "\n", encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
