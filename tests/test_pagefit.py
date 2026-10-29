# -*- coding: utf-8 -*-
"""Перевод на своё место оригинала: повёрнутые подписи схем, блоки схем, сканы форм (аудит 29.10, причина №5).
Случаи — 30GF(135)51GT / 30GF(246)51GT (лист-схема, титул-скан), 30GF118GT (címl, nyilv)."""
import fitz

import assemble as A
import pagefit as PF
import segments as S
from test_forms import word


# ------------------------------------------------------------------ скан формы
def test_pixels_only_names_numbers_dates_stay_pixels():
    for t in ("Schnepp Péter TO 2025.01.15.", "Bucsi Gábor", "2023 ÁPR 21.", "10135772", "Bárdos Zoltán ÜBO",
              "Horváth Zoltán 142506"):
        assert PF.pixels_only(t, set()), t


def test_pixels_only_labels_are_redrawn():
    known = {"oldal", "verzió", "übo", "dátum"}
    for t in ("Oldal Í Verzió", "Verziószám:1", "Dátum:2025.05.", "Készítette:", "név, szervezeti egység/ aláírás/dátum",
              "AZ OLDALSZÁMOK AKTUÁLIS VERZIÓI:", "ÜBO", "Ver-", "Oldal"):
        assert not PF.pixels_only(t, known), t


def test_checkbox_glyph_read_as_letters_is_left_as_pixels():
    # ☒ прочитан как «G4» с уверенностью 32, портал значок не перепечатал (30GF(135), титул): слово не читается надёжно
    ws = [word("NUKLEÁRIS", 10, 10, w=60, h=11), word("ÉRINTI:", 75, 10, w=40, h=11, c=93),
          word("G4", 120, 8, w=12, h=16, c=32, ink=0.22), word("igen;", 140, 10, w=27, h=11, c=97)]
    pix = []
    out, _ = S.ocr_words(ws, "NUKLEÁRIS ÉRINTI: ☒ igen;", pix)
    assert [w["t"] for w in out] == ["NUKLEÁRIS", "ÉRINTI:", "igen;"]
    assert pix == [[120, 8, 132, 24]]                    # место значка — препятствие для перевода справа от абзаца


def test_code_merge_does_not_swallow_label_far_to_the_right():
    # «30GEF(135)51GT» (Tesseract) против «30GF(135)51GT» (портал): слияние кода — только соседних слов; «Verziószám: 2»
    # через 200 pt — другая подпись, она не стирается вместе с кодом (30GF(135), титул)
    ws = [word("30GEF(135)51GT", 10, 10, w=70, c=37), word("Verziószám:", 300, 10, w=50, c=96), word("2", 380, 10, w=5, c=97)]
    out, _ = S.ocr_words(ws, "30GF(135)51GT")
    assert "Verziószám:" in [w["t"] for w in out]


def test_frag_does_not_continue_leftwards_or_across_a_column_edge():
    # Tesseract держит «Verziószám: 2» и «Megnevezés: …» одной строкой блока; правый столбец шапки — отдельные фрагменты
    ws = [word("Azonosító:", 10, 10, w=40, line=(1, 1, 1)),
          word("Verziószám:", 300, 10, w=50, line=(1, 1, 2)), word("2", 360, 10, w=5, line=(1, 1, 2)),
          word("Megnevezés:", 10, 22, w=45, line=(1, 1, 2)), word("ellenőrzése.", 60, 22, w=60, line=(1, 1, 2)),
          word("Oldalszám:", 300, 22, w=48, line=(1, 1, 2))]
    got = [f["t"] for f in S.ocr_frags(ws)]
    assert "Megnevezés: ellenőrzése." in got and "Oldalszám:" in got
    assert not any("Verziószám" in t and "Megnevezés" in t for t in got)


# ------------------------------------------------------------------ повёрнутые подписи
def test_rotated_lines_report_reading_direction():
    doc = fitz.open(); pg = doc.new_page(width=200, height=200)
    pg.insert_text((50, 150), "sikeres", fontsize=10, rotate=90)         # снизу вверх
    pg.insert_text((100, 50), "sikertelen", fontsize=10, rotate=270)     # сверху вниз
    pg.insert_text((20, 20), "horizontal", fontsize=10)
    got = {ln["t"]: ln["rot"] for ln in PF.rotated_lines(pg)}
    assert got == {"sikeres": 90, "sikertelen": 270}


def test_rotated_avail_grows_both_ways_but_stops_at_blocks():
    page = fitz.Rect(0, 0, 200, 300)
    block = fitz.Rect(80, 0, 130, 115)                                   # блок выше подписи
    av = PF.rotated_avail([90, 120, 100, 150], 270, page, [block])
    assert av[1] == 116 and av[3] > 150 and av[0] == 90 and av[2] == 100


def _scheme_pdf(tmp_path, label, rotated=None):
    doc = fitz.open(); pg = doc.new_page(width=700, height=200)
    for i in range(4):                                                                       # блоки схемы 80×40
        pg.draw_rect(fitz.Rect(100 + 120 * i, 80, 180 + 120 * i, 120), color=(0, 0, 1), fill=(0.4, 0.6, 0.9))
    pg.insert_text((115, 98), label[0], fontsize=8)
    pg.insert_text((115, 108), label[1], fontsize=8)
    for i in range(1, 4):
        pg.insert_text((130 + 120 * i, 103), "5.5. pont", fontsize=8)
    if rotated:
        pg.insert_text((60, 150), rotated, fontsize=10, rotate=90)
    path = tmp_path / "scheme.pdf"
    doc.save(path)
    return path


def _seg(i, t, bb, **loc):
    return {"id": i, "text": t, "src": t, "page": 0,
            "loc": {"bbox": bb, "avail": [bb[0], bb[1], 690, 190], "align": "left", "cell": [0, 700], "frags": [bb],
                    "size": 8.0, "bold": False, "family": "sans-serif", "ocr": False, **loc}}


def test_label_is_fitted_into_its_block_and_lines_of_one_block_go_together(tmp_path):
    src = _scheme_pdf(tmp_path, ("Új 132 baros", "tömörségpróba"))
    segs = [_seg(0, "Új 132 baros", [115, 91, 160, 99]), _seg(1, "tömörségpróba", [115, 101, 165, 109])]         + [_seg(2 + i, "5.5. pont", [130 + 120 * (i + 1), 95, 170 + 120 * (i + 1), 103]) for i in range(3)]
    res = {"0": {"ru": "Новое испытание на плотность"}, "1": {"ru": "при давлении 132 бар"},
           **{str(2 + i): {"ru": "п. 5.5."} for i in range(3)}}
    out = tmp_path / "out.pdf"
    A.build_pdf(str(src), str(out), segs, res)
    ws = fitz.open(str(out))[0].get_text("words")
    ru = [w for w in ws if any("а" <= c.lower() <= "я" for c in w[4]) and w[0] < 190]
    assert len(ru) >= 6
    assert all(100 <= w[0] and w[2] <= 180 and 80 <= w[1] and w[3] <= 120 for w in ru), ru     # всё внутри блока


def test_rotated_label_is_drawn_rotated(tmp_path):
    src = _scheme_pdf(tmp_path, ("x", "y"), rotated="sikeres")
    rl = PF.rotated_lines(fitz.open(str(src))[0])
    assert [ln["t"] for ln in rl] == ["sikeres"] and rl[0]["rot"] == 90
    bb = rl[0]["bbox"]
    seg = _seg(0, "sikeres", bb, rot=90, avail=PF.rotated_avail(bb, 90, fitz.Rect(0, 0, 300, 200)))
    out = tmp_path / "out.pdf"
    A.build_pdf(str(src), str(out), [seg], {"0": {"ru": "успешно"}})
    d = fitz.open(str(out))[0].get_text("dict")
    lines = [ln for b in d["blocks"] if b["type"] == 0 for ln in b["lines"]
             if "успешно" in "".join(s["text"] for s in ln["spans"])]
    assert lines and abs(lines[0]["dir"][0]) < 0.1 and abs(lines[0]["dir"][1]) > 0.9            # строка вертикальная
    assert "sikeres" not in fitz.open(str(out))[0].get_text()


def test_pix_paragraph_is_not_redrawn(tmp_path):
    doc = fitz.open(); pg = doc.new_page(width=200, height=100)
    pg.insert_text((20, 50), "Bucsi Gábor", fontsize=9)
    src = tmp_path / "s.pdf"; doc.save(src)
    seg = _seg(0, "Bucsi Gábor", [20, 42, 80, 52], pix=True)
    out = tmp_path / "o.pdf"
    A.build_pdf(str(src), str(out), [seg], {"0": {"ru": "Бучи Габор"}})
    assert "Бучи" not in fitz.open(str(out))[0].get_text()
