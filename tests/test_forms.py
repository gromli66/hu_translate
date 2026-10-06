# -*- coding: utf-8 -*-
"""Сканы форм: линейки на картинке, печать против рукописи, сверка с чтением портала, строки и вписывание.
Каждый случай — дефект, найденный на 30RDGT, 3PR42 и 3SZ19 (06.10)."""
import numpy as np
import fitz

import layout as L
import segments as S
from assemble import _free


def word(t, x, y, w=30, h=6, line=(1, 1, 1), c=95, ink=0.2):
    return {"t": t, "box": [x, y, x + w, y + h], "h": h, "line": line, "key": line[:2], "c": c, "ink": ink}


def test_rules_from_image_finds_lines_not_dots():
    dpi, k = 300, 300 / 72
    img = np.full((int(200 * k), int(400 * k)), 255, np.uint8)
    img[int(100 * k):int(100 * k) + 4, int(20 * k):int(380 * k)] = 0           # сплошная линейка
    for x in range(int(20 * k), int(380 * k), int(4 * k)):                      # пунктир под подпись
        img[int(150 * k):int(150 * k) + 3, x:x + 4] = 0
    img[int(20 * k):int(180 * k), int(10 * k):int(10 * k) + 4] = 0            # вертикаль рамки
    hs, vs = L.rules_from_image(img, dpi)
    assert len(hs) == 1 and abs(hs[0][2] - 100) < 2
    assert len(vs) == 1 and abs(vs[0][2] - 10.5) < 2


def test_handwriting_is_left_alone_print_and_labels_kept():
    ws = [word("Tamási", 10, 10, c=96, ink=0.21),
          word("Vale", 50, 10, c=0, ink=0.05),                 # рукопись: неуверенно и бледно
          word("242.", 90, 10, c=75, ink=0.042),               # рукописные цифры с высокой «уверенностью»
          word("Aláírás:..........sss", 130, 10, c=0, ink=0.07),   # подпись поля с пунктиром
          word("L]", 200, 10, c=86, ink=0.3)]                  # флажок
    out, _ = S.ocr_words(ws, "")
    assert [w["t"] for w in out] == ["Tamási", "Aláírás…"]


def test_portal_fixes_codes_and_signs():
    ws = [word("felé", 10, 10), word("v", 45, 10, w=5), word("5", 55, 10, w=5, c=70),      # «>» прочитан как «5»
          word("1,0", 65, 10, w=12), word("mm,", 80, 10, w=15),
          word("5090-nál", 10, 30, line=(1, 2, 1), c=63, ink=0.25)]                   # «50%» прочитано как «5090»
    out, _ = S.ocr_words(ws, "felé v > 1,0 mm,\n50%-nál")
    assert [w["t"] for w in out] == ["felé", "v", ">", "1,0", "mm,", "50%-nál"]


def test_portal_inserts_only_within_its_line():
    ws = [word("fennáll:", 10, 10), word("A", 10, 30, line=(1, 2, 1)), word("turbina", 30, 30, line=(1, 2, 1))]
    out, lost = S.ocr_words(ws, "fennáll:\n3/2 A turbina")                # «3/2» из рамки строкой ниже
    assert out[0]["t"] == "fennáll:" and lost == ["3/2"]


def test_frags_split_by_rule_and_fix_inflated_box():
    ws = [word("Tamási", 10, 10, w=24, h=6), word("Zoltán", 37, 10, w=22, h=6),
          word("TO", 70, 10, w=12, h=6),                                         # за вертикальной линейкой x=65
          word("Aláírás:", 120, 6, w=16, h=14)]                                  # рамку раздула подпись от руки
    fr = S.ocr_frags(ws, vs=[(0, 30, 65)])
    assert [f["t"] for f in fr] == ["Tamási Zoltán", "TO", "Aláírás:"]
    assert fr[2]["bbox"][3] - fr[2]["bbox"][1] < 8 and fr[2]["size"] < 7       # высота и кегль — как у соседей
    assert all(f["family"] == "serif" for f in fr)


def test_monospace_page_detected():
    # Courier: буква ≈ 0,6 кегля; у шрифта с засечками ≈ 0,43
    ws = [word("tengelyiranyu elmozdulasa", 10, 10 + 12 * i, w=25 * 0.6 * 8.3, h=6, line=(1, i + 1, 1)) for i in range(5)]
    assert all(f["family"] == "monospace" for f in S.ocr_frags(ws))


def test_label_closes_paragraph_on_forms():
    frags = [{"t": "Alkalmasság ellenőrzés szükséges:", "bbox": [10, 10, 200, 18], "size": 8, "bold": False},
             {"t": "Alkalmasság ellenőrzés módja:", "bbox": [10, 20, 190, 28], "size": 8, "bold": False}]
    assert len(L.group(frags, [], [], form=True)) == 2


def test_free_keeps_own_line_and_avoids_neighbours():
    busy = [fitz.Rect(380, 116, 440, 135)]                    # «Номер версии», перенесённое на вторую строку
    r = _free(fitz.Rect(116, 126, 480, 140), busy, [116, 126, 320, 134])
    assert r.x1 <= 380 and r.y0 == 126                        # длинный перевод строки ниже не залезает под него


def test_list_line_avail_stops_at_next_line():
    """Рамки строк тесного списка перекрываются: перенос строки не должен наезжать на следующую (32RDGT, п. 2.2)."""
    pars = [{"frags": [{"bbox": [203, 262, 299, 276.9]}], "bbox": [203, 262, 299, 276.9]},
            {"frags": [{"bbox": [203, 274.6, 306, 289.5]}], "bbox": [203, 274.6, 306, 289.5]}]
    L.avail_rects(pars, [], [], fitz.Rect(0, 0, 595, 842))
    assert pars[0]["avail"][3] == 276.9


import pytest


@pytest.mark.parametrize("rot", [0, 90, 180, 270])
def test_rotated_page_translation_lands_upright(tmp_path, rot):
    """Страница с /Rotate: координаты сегментов — как страница видна, перевод встаёт на место исходной строки
    и читается слева направо (32RDGT: листы с поворотом 270° получали перевод поперёк листа, 07.10)."""
    from assemble import build_pdf
    src, out = tmp_path / "src.pdf", tmp_path / "ru.pdf"
    d = fitz.open(); p = d.new_page(width=595, height=842); p.set_rotation(rot)
    p.insert_htmlbox(fitz.Rect(100, 100, 400, 130) * p.derotation_matrix, "Alma körte szilva",
                     css="* {font-family: serif; font-size: 12pt;}", rotate=rot)
    d.save(src)
    page = fitz.open(src)[0]
    pars = L.group(L.frags_from_text(page), [], [])
    L.avail_rects(pars, [], [], page.rect)
    assert [q["text"] for q in pars] == ["Alma körte szilva"]
    x0, y0, x1, y1 = pars[0]["bbox"]
    assert abs(x0 - 100) < 4 and 95 <= y0 < 115
    seg = {"id": 0, "page": 0, "text": pars[0]["text"],
           "loc": {"bbox": pars[0]["bbox"], "avail": [x0, y0, 500, y1], "frags": [f["bbox"] for f in pars[0]["frags"]],
                   "size": 12, "bold": False, "family": "serif", "ocr": False}}
    build_pdf(str(src), str(out), [seg], {"0": {"ru": "Яблоко груша слива"}})
    q = fitz.open(out)[0]
    rot_m = q.rotation_matrix
    lines = [ln for b in q.get_text("dict")["blocks"] for ln in b.get("lines", [])]
    assert "".join(s["text"] for ln in lines for s in ln["spans"]).split() == ["Яблоко", "груша", "слива"]
    for ln in lines:
        v = fitz.Point(ln["dir"]) * rot_m - fitz.Point(0, 0) * rot_m
        r = fitz.Rect(ln["bbox"]) * rot_m
        assert v.x > 0.9 and abs(v.y) < 0.1                    # слева направо на видимой странице
        assert abs(r.x0 - x0) < 4 and abs(r.y0 - y0) < 6       # на месте исходной строки


def test_glued_label_split_from_handwriting():
    """«oílNév:» — рукописная дата слиплась с подписью поля: подпись отдельно, обрывок линейки «I» — прочь."""
    out = S._split_glued([word("oílNév:", 150, 280, w=41), word("INév:", 172, 294, w=20), word("Megnevezés:", 10, 10)])
    assert [w["t"] for w in out] == ["oíl", "Név:", "Név:", "Megnevezés:"]
    assert out[0]["box"][2] == out[1]["box"][0] and out[1]["box"][2] == 191


def test_checkbox_and_bullet_glyphs_are_not_text():
    """Флажок ☑ («VI», «[5») и буллет • («e») — картинки: не переводятся и не стирают соседний код (32RDGT)."""
    ws = [word("2./", 68, 100, w=10, h=7), word("VI", 230, 100, w=8.4, h=8.6, c=45),
          word("új", 245, 100, w=12, h=7), word("[5", 308, 100, w=8.6, h=8.6, c=37),
          word("ciklikus", 320, 100, h=7),
          word("e", 120, 380, w=3.8, h=3.8, line=(2, 1, 1), ink=0.66, c=87),
          word("RD3.1.;", 133, 378, w=34, h=8.4, line=(2, 1, 1), c=44),
          word("RD", 170, 380, w=13, h=6.7, line=(2, 1, 1)), word("3.2.", 186, 380, w=14, h=6.7, line=(2, 1, 1))]
    out, _ = S.ocr_words(ws, "2./ ☑ új ☑ ciklikus\n• RD 3.1.; RD 3.2.")
    assert [w["t"] for w in out] == ["2./", "új", "ciklikus", "RD 3.1.;", "RD", "3.2."]


def test_numbered_item_starts_paragraph():
    """«3./ …» под строкой 2 — новый пункт, а не продолжение строки 2 (рамка абзаца накрывала подпись «2./»)."""
    f = lambda t, x0, y0, x1: {"t": t, "bbox": [x0, y0, x1, y0 + 8], "size": 10, "bold": False}
    pars = L.group([f("új utasítás ciklikus felülvizsgálat", 190, 100, 430), f("3./ A tesztelési utasítás", 69, 109, 300)],
                   [], [], form=True)
    assert len(pars) == 2


def test_center_only_in_middle_of_own_cell_and_not_for_lists():
    """По центру — посередине своей ячейки; строка списка, делящая левый край с соседом, — влево; штамп у кромки
    листа края не сбивает (32RDGT, титул)."""
    vs = [(0, 300, 69.0), (0, 300, 520.0)]
    p = lambda x0, y0, x1: {"bbox": [x0, y0, x1, y0 + 10], "frags": [{"bbox": [x0, y0, x1, y0 + 10]}]}
    stamp, title = p(0, 0, 60), p(110, 50, 480)
    li1, li2 = p(203, 150, 300), p(203, 163, 386)
    L.aligns([stamp, title, li1, li2], vs)
    assert title["align"] == "center" and title["cell"] == [69.0, 520.0]
    assert li1["align"] == "left" and li2["align"] == "left"
