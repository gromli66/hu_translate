# -*- coding: utf-8 -*-
"""PDF с текстовым слоем (руководство UTS, 15.10): страница со скриншотом не скан, невидимый текст не переводится,
маркер пункта отделён от текста, ячейка таблицы — одним абзацем, надписи скриншота перерисовываются без порчи
соседних пикселей. (Проверки DOCX из оригинала через pdf2docx — docx_src — сняты вместе с ним, 01.11.)"""
import numpy as np
import fitz

import layout as L
import segments as S
import shots as SH

LOREM = ("Az UTS alkalmazás grafikus felhasználói felülete a következő fő részekre bontható a felhasználó szerint, "
         "amelyek a menüsor alatt helyezkednek el és a tesztelés során folyamatosan láthatók maradnak.")


def png(w, h, gray=200):
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, w, h), False)
    pix.clear_with(gray)
    return pix.tobytes("png")


def word(t, box, c=95):
    return {"t": t, "box": list(box), "h": box[3] - box[1], "c": c}


def test_screenshot_words_do_not_make_page_a_scan():
    doc = fitz.open(); page = doc.new_page()
    tw = fitz.TextWriter(page.rect)
    y = 60
    for _ in range(6):
        tw.fill_textbox(fitz.Rect(50, y, 550, y + 40), LOREM, fontsize=9); y += 45
    tw.write_text(page)
    page.insert_image(fitz.Rect(50, 400, 550, 600), stream=png(500, 200))
    words = [word("Eredmények", (60, 410, 110, 420)), word("Adminisztráció", (120, 410, 180, 420)),
             word("Tesztelés", (190, 410, 230, 420)), word("felhasználói", (60, 430, 120, 440))]
    words += [word(w, (50 + i * 10, 62, 58 + i * 10, 70)) for i, w in enumerate(LOREM.split()[:20])]
    share, outside = S.page_missing_share(page, words)
    assert share == 0.0
    assert all(w["box"][1] < 400 for w in outside)          # слова скриншота отброшены


def test_curve_text_still_counts_as_missing():
    doc = fitz.open(); page = doc.new_page()
    page.insert_text((50, 60), "Azonosító: 3PR42 " + " ".join(["Verziószám"] * 8), fontsize=9)
    words = [word("Azonosító", (50, 52, 90, 62))] + [word("blokkvezénylőben", (50 + i * 60, 100, 100 + i * 60, 110))
                                                      for i in range(5)]
    share, _ = S.page_missing_share(page, words)
    assert share > 0.15


def test_hidden_text_under_full_page_image_is_skipped():
    doc = fitz.open(); page = doc.new_page()
    page.insert_text((50, 800), "Létesítmény: Paksi Atomerőmű", fontsize=9)       # под фоном
    page.insert_image(page.rect, stream=png(60, 80, 240))
    page.insert_text((50, 100), "UTS felhasználói kézikönyv", fontsize=14)        # поверх фона
    texts = [f["t"] for f in L.frags_from_text(page)]
    assert texts == ["UTS felhasználói kézikönyv"]


def test_bullet_marker_is_own_fragment_and_text_is_item():
    doc = fitz.open(); page = doc.new_page()
    put = lambda x, y, t: page.insert_htmlbox(fitz.Rect(x, y, 560, y + 14), t, css="* {font-size: 10pt;}")
    put(80, 100, "•"); put(98, 100, "Alfanumerika: a vizsgálandó jel alfanumerikája"); put(98, 113, "folytatás")
    put(80, 140, "•"); put(98, 140, "Művelet: összehasonlító logikai művelet")
    hs, vs = L.rules(page)
    pars = L.group(L.frags_from_text(page, vs), hs, vs)
    texts = [p["text"] for p in pars]
    assert "•" in texts
    assert "Művelet: összehasonlító logikai művelet" in texts       # пункт не приклеился к предыдущему
    assert not any(t.startswith("•") and len(t) > 1 for t in texts)


def test_narrow_cell_lines_stay_one_paragraph():
    doc = fitz.open(); page = doc.new_page()
    for x in (60, 150):
        page.draw_line((x, 80), (x, 200))
    page.insert_text((65, 100), "Teszt", fontsize=10)
    page.insert_text((65, 113), "szüneteltetése", fontsize=10)
    hs, vs = L.rules(page)
    pars = L.group(L.frags_from_text(page, vs), hs, vs)
    assert [p["text"] for p in pars] == ["Teszt szüneteltetése"]


def test_screenshot_label_repaint_keeps_neighbours():
    img = np.full((40, 200, 3), 120, np.uint8)            # серая шапка
    img[30:32, :] = 255                                   # белая линия под шапкой
    img[12:22, 20:70] = 250                               # «буквы» надписи
    doc = fitz.open(); page = doc.new_page()
    pix = fitz.Pixmap(fitz.csRGB, 200, 40, img.tobytes(), False)
    xref = page.insert_image(fitz.Rect(50, 50, 250, 90), pixmap=pix)
    out = SH.render(doc, xref, [({"box": [20, 12, 69, 21], "h": 9, "t": "Kategória"}, "Категория")])
    a = np.frombuffer(out.samples, np.uint8).reshape(out.height, out.width, 3)
    assert (a[30:32, :] == 255).all()                     # линия цела
    assert (np.abs(a[2:8, :].astype(int) - 120) <= 2).all()   # фон выше надписи не тронут
    assert (a[10:24, 15:120] > 200).any()                 # перевод нарисован светлым, как оригинал


# ---------------------------------------------------------------- находки красной команды 15.10
def test_watermark_over_whole_page_does_not_hide_text():
    doc = fitz.open(); page = doc.new_page()
    for i in range(10):
        page.insert_text((50, 80 + 20 * i), f"Sor {i}: ellenőrizni kell a nyomást", fontsize=10)
    page.insert_image(page.rect, stream=png(60, 80, 250))          # во весь лист, поверх всего текста
    assert len(L.frags_from_text(page)) == 10


def test_stencil_image_is_not_a_screenshot():
    doc = fitz.open(); page = doc.new_page()
    a = (np.random.rand(200, 400) > 0.5).astype(np.uint8) * 255
    x = page.insert_image(fitz.Rect(100, 450, 400, 600), pixmap=fitz.Pixmap(fitz.csGRAY, 400, 200, a.tobytes(), False))
    doc.xref_set_key(x, "ImageMask", "true")
    assert x not in SH.find_shots(doc)


def test_label_that_does_not_fit_keeps_original():
    img = np.full((30, 60, 3), 230, np.uint8)
    img[10:20, 5:25] = 20                                             # короткая надпись «Ki»
    img[:, 28:30] = 0                                                 # граница кнопки сразу справа
    doc = fitz.open(); page = doc.new_page()
    xref = page.insert_image(fitz.Rect(50, 50, 110, 80), pixmap=fitz.Pixmap(fitz.csRGB, 60, 30, img.tobytes(), False))
    out = SH.render(doc, xref, [({"box": [5, 10, 24, 19], "h": 9, "t": "Ki"}, "Выключено полностью и навсегда")])
    a = np.frombuffer(out.samples, np.uint8).reshape(out.height, out.width, 3)
    assert (a[10:20, 5:25] == 20).all()                               # оригинал не стёрт
