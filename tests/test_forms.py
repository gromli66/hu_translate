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
