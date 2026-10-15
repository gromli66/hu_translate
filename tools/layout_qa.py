# -*- coding: utf-8 -*-
"""Проверка вёрстки перевода порталом по картинкам: оригинал и перевод страницы рядом → список дефектов.
Результат — подсказка вычитчику, какие страницы смотреть глазами (не замена просмотра): на старом выводе UTS
(15.10) отмечены 7 из 8 испорченных страниц (пропуск — разнобой кегля в таблице, стр. 17), ложных тревог нет.

    python tools/layout_qa.py <оригинал.pdf> <перевод.pdf> <out.json> [страницы через запятую]"""
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

PROMPT = (
    "Слева — страница технического документа на венгерском, справа — та же страница после перевода на русский "
    "с сохранением вёрстки. Сравни вёрстку правой страницы с левой и найди дефекты оформления (не перевода):\n"
    "- overlap: текст наезжает на другой текст, линии таблицы, картинку или значок;\n"
    "- tiny: текст заметно мельче, чем в том же месте слева, или разного размера в соседних ячейках/пунктах;\n"
    "- cut: текст обрезан или вылезает за рамку ячейки/страницы;\n"
    "- untranslated: на правой странице остался венгерский текст (кроме кодов, имён, адресов);\n"
    "- artifact: зачёркивания, белые заплатки, мусорные символы, испорченные картинки или логотипы;\n"
    "- missing: на правой странице пропал текст или элемент, который есть слева.\n"
    "Отвечай только JSON: {\"ok\": true|false, \"issues\": [{\"type\": тип, \"where\": \"коротко, где на странице\"}]}. "
    "Если дефектов нет — {\"ok\": true, \"issues\": []}.")


def pair_png(a, b, pno, dpi=90):
    import fitz
    pa, pb = a[pno].get_pixmap(dpi=dpi), b[pno].get_pixmap(dpi=dpi)
    out = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, pa.width + pb.width + 10, max(pa.height, pb.height)), False)
    out.clear_with(255)
    for p, x in ((pa, 0), (pb, pa.width + 10)):
        p = p if p.n == 3 else fitz.Pixmap(fitz.csRGB, p)
        p.set_origin(x, 0); out.copy(p, p.irect)
    return out.tobytes("png")


def main(src, ru, out, pages=None, workers=2):
    import fitz
    from llm import chat, img_b64
    a, b = fitz.open(src), fitz.open(ru)
    pages = pages if pages is not None else list(range(min(len(a), len(b))))
    imgs = {p: pair_png(a, b, p) for p in pages}          # MuPDF — в этом потоке

    def one(p):
        r = chat([{"role": "user", "content": PROMPT}], images=[img_b64(imgs[p])], max_tokens=3000, tag=f"qa{p}")
        m = re.search(r"\{.*\}", r.get("text") or "", re.S)
        try:
            v = json.loads(m.group()) if m else {"ok": None, "issues": [], "raw": (r.get("text") or r.get("error"))[:300]}
        except ValueError:
            v = {"ok": None, "issues": [], "raw": (r.get("text") or "")[:300]}
        return p, v

    with ThreadPoolExecutor(workers) as ex:
        res = dict(ex.map(one, pages))
    Path(out).write_text(json.dumps({str(p + 1): v for p, v in sorted(res.items())}, ensure_ascii=False, indent=1),
                         encoding="utf-8")
    # замечание — и при «ok: true» с непустым списком: портал так отвечал на страницах с реальными дефектами
    # (калибровка на старом выводе UTS 15.10: 7 из 8 плохих страниц, ложных тревог нет)
    bad = [p + 1 for p, v in sorted(res.items()) if v.get("ok") is False or v.get("issues")]
    print(f"страниц {len(pages)}, с замечаниями {len(bad)}: {bad}")


if __name__ == "__main__":
    pg = [int(x) - 1 for x in sys.argv[4].split(",")] if len(sys.argv) > 4 else None
    main(sys.argv[1], sys.argv[2], sys.argv[3], pg)
