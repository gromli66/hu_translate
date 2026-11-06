# -*- coding: utf-8 -*-
"""Дефекты сканов по просмотру глазами 06.11 (EYES_2026-11-06: N6, п. 5, 13, штамп 30RYBSZGT, красная плашка формы). Слова — как их
отдаёт Tesseract (segments._tess_words): t, box, h, line, key, c (уверенность), ink (плотность тёмной краски), ink_c (красная краска)."""
import io

import fitz
import numpy as np
from PIL import Image

import pagefit as PF
import segments as S
from assemble import build_pdf


def word(t, x, y, w=30, h=6, line=(1, 1, 1), c=95, ink=0.2, **kw):
    return dict({"t": t, "box": [x, y, x + w, y + h], "h": h, "line": line, "key": line[:2], "c": c, "ink": ink}, **kw)


def test_red_print_is_not_handwriting():
    """Красная печать (предупреждение «AZ UT_ÜVIG_25 … FIGYELEMBEVÉTELÉVEL KÉSZÜLT!», красные заголовки форм) Tesseract читает
    уверенно (89–97), а тёмной краски в сером изображении у неё почти нет (красный (251,61,49) светлее порога): слово шло как
    «бледная рукопись», не переводилось и не закрывалось плашкой — венгерский остался на листе. Красная краска — тоже краска
    (PF.mark_red_ink поднимает плотность слова); рукопись (синяя паста, бледно, неуверенно) — по-прежнему рукопись."""
    k = 300 / 72
    img = np.full((int(120 * k), int(300 * k), 3), 255, np.uint8)
    img[int(70 * k):int(80 * k), int(150 * k):int(277 * k)] = (251, 61, 49)               # красная строка
    img[int(20 * k):int(30 * k), int(50 * k):int(100 * k)] = (60, 70, 170)                # синяя паста
    red = word("FIGYELEMBEVÉTELÉVEL", 150, 70, w=127, h=10, c=93, ink=0.02)
    pen = word("Vale", 50, 20, w=50, h=10, c=0, ink=0.05)
    for w in (red, pen):
        PF.mark_red_ink(w, img, 300)
    assert red["ink_c"] > 0.9 and red["ink"] > 0.9 and pen["ink_c"] == 0.0 and pen["ink"] == 0.05
    out, _ = S.ocr_words([red, pen, word("Tamási", 10, 90, c=96, ink=0.21)], "")
    assert [w["t"] for w in out] == ["FIGYELEMBEVÉTELÉVEL", "Tamási"]                   # рукопись — не слово, красная печать — слово


def test_red_line_of_a_scan_page_becomes_a_segment(tmp_path, monkeypatch):
    """Через pdf_segments: красная строка скана (Tesseract прочитал её уверенно) — сегмент, синяя подпись рядом — нет."""
    import llm
    k = 150 / 72
    im = np.full((round(842 * k), round(595 * k), 3), 255, np.uint8)
    im[round(70 * k):round(80 * k), round(150 * k):round(277 * k)] = (251, 61, 49)
    buf = io.BytesIO()
    Image.fromarray(im).save(buf, "PNG")
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(tmp_path / "red.pdf")
    monkeypatch.setattr(S, "_tess_words", lambda png, dpi: [
        {"t": "FIGYELEMBEVÉTELÉVEL", "key": (1, 1), "line": (1, 1, 1), "box": [150.0, 70.0, 277.0, 80.0], "h": 10.0, "c": 93.0}])
    monkeypatch.setattr(llm, "chat", lambda *a, **kw: {"text": "FIGYELEMBEVÉTELÉVEL"})
    monkeypatch.setattr(S, "OCR_DIR", tmp_path / "ocr")
    segs, _ = S.pdf_segments(str(tmp_path / "red.pdf"), ocr=True, shots=False)
    assert [s["text"] for s in segs] == ["FIGYELEMBEVÉTELÉVEL"]


def test_ink_red_counts_only_saturated_red_pixels():
    k = 300 / 72
    img = np.full((int(60 * k), int(200 * k), 3), 255, np.uint8)
    img[int(10 * k):int(20 * k), int(10 * k):int(60 * k)] = (251, 61, 49)                 # красный блок
    img[int(10 * k):int(20 * k), int(100 * k):int(150 * k)] = (40, 60, 170)               # синяя паста
    img[int(30 * k):int(40 * k), int(10 * k):int(60 * k)] = (132, 132, 132)               # серая печать
    ink = lambda x0, x1, y0=8, y1=22: PF.ink_red(img, [x0, y0, x1, y1], 300)
    assert ink(10, 60) > 0.5 and ink(100, 150) == 0.0 and ink(10, 60, 28, 42) == 0.0


def test_cell_rule_read_as_letters_is_not_a_word():
    """Вертикальная линейка ячейки таблицы версий, прочитанная как «[d» (рамка 2,4 × 4,6 pt, краски 0,64, уверенность 92): знаки
    в такую щель не вмещаются — слово шло в перевод («[d Страница» на титуле 3PR10, п. 5 EYES). Теперь это картинка (линия остаётся):
    правило общее — любые два знака и больше в рамке уже 0,3 высоты на знак при плотной краске («|1», «l|», «[d»)."""
    ws = [word("[d", 240.5, 373.9, w=2.4, h=4.6, c=92, ink=0.637), word("Oldal", 247.9, 373.7, w=20.7, h=4.8, c=96, ink=0.3)]
    pix = []
    assert [f["t"] for f in S.ocr_frags(ws, pix=pix)] == ["Oldal"] and pix == [[240.5, 373.9, 242.9, 378.5]]
    for bar in ("|1", "l|", "[l", "I|I"):                                                       # другие чтения линейки
        pix = []
        got = S.ocr_frags([word(bar, 100, 50, w=0.4 * len(bar), h=6, ink=0.8), word("Verzió", 106, 50, w=20, h=6)], pix=pix)
        assert [f["t"] for f in got] == ["Verzió"] and len(pix) == 1, bar
    ok = [word("Id", 10, 10, w=8, h=5, c=90, ink=0.3), word("il", 30, 10, w=4, h=7, c=90, ink=0.3),         # настоящие короткие слова
          word("li", 50, 10, w=5, h=7, c=90, ink=0.5)]
    assert [f["t"] for f in S.ocr_frags(ok, pix=[])] == ["Id il li"]
    assert PF.ruleword(word("[d", 0, 0, w=2.4, h=4.6, ink=0.637) | {"raw": "[d"}) and not PF.ruleword(word("Id", 0, 0, w=8, h=5, ink=0.3))


def test_hyphenated_bold_word_is_not_a_cell_rule():
    """Слово с переносом склеивается (_dehyphen) с рамкой первой половины: «mű-» (10 pt) + «ködik» = «működik» — 7 знаков в рамке
    шириной на 3, и «обрывок линейки» мерил по 7 (0,3 × 6 × 7 = 12,6 > 10): у жирного слова (краска > 0,45) оно уходило в картинку
    целиком, вместе со второй половиной (на кэше Tesseract набора у 259 склеек из 950 так; спасал только порог краски, ≤ 0,355)."""
    ws = [word("A", 10, 10, w=4, line=(1, 1, 1), ink=0.5), word("szivattyú", 18, 10, w=40, line=(1, 1, 1), ink=0.5),
          word("mű-", 62, 10, w=10, line=(1, 1, 1), ink=0.5), word("ködik.", 10, 20, w=26, line=(1, 1, 2), ink=0.5)]
    words, _ = S.ocr_words(S._dehyphen(ws), "")
    assert " ".join(f["t"] for f in S.ocr_frags(words, pix=[])) == "A szivattyú működik."


def test_piece_of_word_cut_at_cell_rule_is_measured_by_its_own_letters():
    """Слово поперёк линейки ячейки («szerintijA.2.2.») режется на куски (_split_at_rules), а raw остаётся чтением целого слова:
    15 знаков на кусок в 26 pt — оба куска жирного слова сходили за «обрывок линейки» и пропадали."""
    ws = [word("szerintijA.2.2.", 100, 10, w=50, h=6, ink=0.5, raw="szerintijA.2.2.")]
    assert [f["t"] for f in S.ocr_frags(ws, [(5, 20, 127)], pix=[])] == ["szerinti", "A.2.2."]


def _prose_page():
    """Строка машинописного абзаца (моноширинный, пробелы по 10–12 pt) и три другие строки листа, где слово после зазора ≥ 12 pt
    стоит у x ≈ 122–127 (таблица версий, штамп даты): по ним «граница столбца» 125 pt находилась и у слова прозы."""
    line = [("A", 54.0, 6.2), ("KU", 70.3, 10.8), ("2.6.", 92.2, 20.5), ("pontjában", 124.8, 53.5), ("az", 189.1, 10.3),
            ("ideiglenes", 210.7, 58.1), ("módosítással", 279.4, 71.2), ("történő", 361.7, 40.8)]
    ws = [word(t, x, 236, w=w, h=8.4, line=(15, 2, 1)) for t, x, w in line]
    for i, (t, x, y) in enumerate((("Utasítás", 54, 181), ("1,", 122.2, 181), ("INFObank", 54, 192), ("3,", 121.9, 192),
                                   ("Dátum:", 60, 460), ("JUN", 127.0, 460))):
        ws.append(word(t, x, y, w=42 if len(t) > 3 else 8, h=7, line=(16 + i // 2, 1, 1)))
    return ws


def test_prose_line_is_not_cut_at_false_column_edge():
    """«A KU 2.6. pontjában…» рвалась на «A KU 2.6.» и «pontjában…»: зазор 12,0 pt (в этой строке все пробелы 10–12 pt) прошёл
    за «пустое место перед столбцом», а левый край 124,8 совпал с краем у трёх чужих строк листа — «A KU 2.6.» уходила в абзац ниже, его
    перевод рисовался поверх (3PR10 стр. 2, N6 EYES). Зазор строки в обычных пределах её пробелов столбцом не считается."""
    got = [f["t"] for f in S.ocr_frags(_prose_page()) if f["bbox"][1] > 230 and f["bbox"][1] < 240]
    assert got == ["A KU 2.6. pontjában az ideiglenes módosítással történő"]


def _signature_row(y, org):
    """Строка подписей формы (30RVRWVTBIGT стр. 2, «6./ Készítette»): ячейки «Dátum: …», «Név: …», «Szervezet: …», «Aláírás:» без
    линеек, пробелы 9–10 pt, перед «Aláírás:» — 12,0 pt."""
    line = (11, 1, y)
    return [word(t, x, y, w=w, h=7, line=line) for t, x, w in (
        ("Dátum:", 78.5, 28.8), ("2026.04.16.", 117.4, 46.5), ("Név:", 173.0, 22.8), ("Balogh", 206.2, 28.0),
        ("Béla", 237.6, 17.5), ("Szervezet:", 289.7, 45.1), (org, 344.2, 17.0), ("Aláírás:", 373.2, 35.0))]


def test_signature_row_cells_are_not_glued_across_cell_border():
    """В строке подписей больше трёх пробелов по 9–10 pt, и правило N6 («зазор 12–14 pt в строке с пробелами машинописи — пробел»)
    склеивало ячейки через границу: «Szervezet: ÜIO Aláírás:» — одна плашка на две ячейки, линейка между ними под ней (3SZ23,
    30RVRWVTBIGT, 30RVRYVTBIGT, 30RA+SGT, 01-02FMP0BIGT). Слово с прописной после такого зазора — начало ячейки (подпись поля)."""
    pix = []
    got = [f["t"] for f in S.ocr_frags(_signature_row(237.6, "ÜIO") + _signature_row(261.6, "RTO"), pix=pix)]
    assert "Aláírás:" in got and not any("Szervezet" in t and "Aláírás" in t for t in got), got
    assert "Dátum:" in got and not any("2026" in t for t in got)            # дата между подписями — пиксели, как прежде


def test_table_header_column_edge_survives_typewriter_spaces():
    """Шапка таблицы версий «Oldal Verzió Oldal Í Verzió Oldal» идёт ячейками через 11,5–12,2 pt (и 3–5 pt у «Í»): по правилу N6 она
    сходила за строку машинописи, «Verzió» переставало быть краем столбца, и «TU» ⇥ «-ban» под ним шли одним куском — а за ним
    в абзац уходил номер страницы строки ниже («TU -ban 1», 30RR2126GT, 30YBGT, 01-02FMP0BIGT, 1SZ01, 3SZ23 стр. 3)."""
    head = [("Oldal", 75.1, 16.3), ("Verzió", 103.4, 19.5), ("Oldal", 134.4, 16.1), ("Í", 155.5, 3.6), ("Verzió", 162.5, 19.2),
            ("Oldal", 193.4, 16.1)]
    ws = [word(t, x, 473.3, w=w, h=5, line=(28, 1, 1)) for t, x, w in head]
    ws += [word(t, x, 490.6, w=w, h=7, line=(30, 1, 1)) for t, x, w in (("TU", 76.8, 13.2), ("-ban", 103.4, 19.2),
                                                                         ("9", 140.4, 4.3), ("0", 170.2, 4.3))]
    got = [f["t"] for f in S.ocr_frags(ws, pix=[])]
    assert "TU" in got and not any("TU" in t and "-ban" in t for t in got), got


def test_real_column_gap_still_cuts_header_line():
    """Столбцы шапки через пустое место (25 pt при пробелах 3–6 pt) режутся по-прежнему."""
    ws = [word("Azonosító:", 62, 10, w=45, line=(1, 1, 1)), word("3PR10", 110, 10, w=25, line=(1, 1, 1)),
          word("Verziószám:", 420, 10, w=50, line=(1, 1, 1)), word("1", 473, 10, w=4, line=(1, 1, 1)),
          word("Megnevezés:", 62, 20, w=50, line=(1, 2, 1)), word("Főkeringtető", 115, 20, w=55, line=(1, 2, 1)),
          word("szivattyú.", 173, 20, w=40, line=(1, 2, 1)), word("Oldalszám:", 420, 20, w=48, line=(1, 2, 1)),
          word("1", 471, 20, w=4, line=(1, 2, 1))]
    assert [f["t"] for f in S.ocr_frags(ws)] == ["Azonosító:", "Verziószám:", "Megnevezés: Főkeringtető szivattyú.",
                                                  "Oldalszám:"]                  # значения «3PR10», «1» — пиксели (хвост строки)


def test_value_between_two_labels_stays_pixels():
    """Значение после подписи в середине строки («Dátum: 71127 JAN 2 Név: Bucsi Gábor…») — пиксели оригинала, как у хвоста строки
    (E7 §6.4): штамп «2022 JAN 25» Tesseract прочитал «71127 JAN 2», и неверное чтение рисовалось поверх штампа (30YCGT NYILV, п. 13
    EYES). Подпись — своим куском, остаток строки — своим, рамка значения — в pix."""
    ws = [word("Dátum:", 68, 634, w=34), word("71127", 108, 634, w=26), word("JAN", 138, 634, w=18), word("2", 160, 634, w=5),
          word("Név:", 190, 634, w=20), word("Bucsi", 214, 634, w=26), word("Gábor", 244, 634, w=30),
          word("Szervezet:", 280, 634, w=48), word("ÜVFO", 332, 634, w=24), word("|", 360, 634, w=3), word("Aláírás:", 366, 634, w=40),
          word("MÉR", 410, 634, w=24)]
    pix = []
    got = S.ocr_frags(ws, pix=pix)
    assert [f["t"] for f in got] == ["Dátum:", "Név: Bucsi Gábor Szervezet: ÜVFO | Aláírás:"]
    assert [round(v) for v in pix[0]] == [108, 634, 165, 640]                           # «71127 JAN 2» — рамка значения
    plain = [word("Verziószám:", 10, 10, w=50), word("1", 64, 10, w=4), word("Oldalszám:", 100, 10, w=48), word("2", 152, 10, w=4)]
    assert [f["t"] for f in S.ocr_frags(plain, pix=[])] == ["Verziószám:", "Oldalszám:"]  # короткие значения-числа — как хвост


def _scan_pdf(path, k=150 / 72, stamp=None):
    """Страница-скан: белое поле с чёрным «штампом» (если задан) и строкой обычной печати."""
    im = Image.new("RGB", (round(595 * k), round(842 * k)), "white")
    if stamp:
        x0, y0, x1, y1 = (round(v * k) for v in stamp)
        px = np.array(im)
        px[y0:y1, x0:x1] = 30
        im = Image.fromarray(px)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(path)


def _ocr_seg(i, text, frags, size=11.0):
    bb = [min(f[0] for f in frags), min(f[1] for f in frags), max(f[2] for f in frags), max(f[3] for f in frags)]
    return {"id": i, "page": 0, "text": text, "loc": {"bbox": bb, "avail": [*bb[:2], 520, bb[3]], "frags": frags, "size": size,
                                                       "bold": False, "family": "serif", "ocr": True}}


def test_plate_does_not_erase_table_rules_at_its_edges(tmp_path):
    """Белая плашка строки скана стирала линейки таблицы, лежащие у её верхнего и нижнего края: рамка красной плашки «AZ UT_ÜVIG_25 …»
    (верхняя и нижняя линии в 1,5 pt от букв) оставалась без линий на всю ширину надписи. Линейка у края плашки (длинная, тонкая,
    сплошная) остаётся — плашка обрезается по ней."""
    k = 150 / 72
    im = Image.new("RGB", (round(595 * k), round(842 * k)), "white")
    px = np.array(im)
    for y in (300, 316):                                                            # две линейки рамки: сверху и снизу надписи
        px[round(y * k):round(y * k) + 2, round(50 * k):round(400 * k)] = 0
    px[round(305 * k):round(311 * k), round(100 * k):round(160 * k)] = 0              # буквы между ними
    src = tmp_path / "frame.pdf"
    buf = io.BytesIO()
    Image.fromarray(px).save(buf, "PNG")
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(src)
    seg = _ocr_seg(0, "KÉSZÜLT!", [[100, 299.5, 160, 317.5]], 9.0)                    # рамка строки захватывает обе линейки
    out = tmp_path / "ru.pdf"
    build_pdf(str(src), str(out), [seg], {"0": {"ru": "СОСТАВЛЕНО!"}}, check=False)
    page = fitz.open(out)[0]
    assert "СОСТАВЛЕНО" in page.get_text()
    for y in (300, 316):                                                            # обе линейки целы по всей ширине плашки
        pm = page.get_pixmap(dpi=150, clip=fitz.Rect(100, y - 0.5, 160, y + 1.5), colorspace=fitz.csGRAY)
        cols = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width).min(axis=0)
        assert (cols < 100).mean() > 0.9, y                                         # в каждом столбце — тёмная линия


def test_rotated_scan_page_is_built_without_rule_trimming(tmp_path):
    """Повёрнутый лист скана (/Rotate): краска листа для линеек и края слов не читается (плашки и чужая краска — только без поворота) —
    плашка рисуется как была, сборка не падает."""
    src = tmp_path / "rot.pdf"
    _scan_pdf(src)
    d = fitz.open(src)
    d[0].set_rotation(90)
    d.save(tmp_path / "rot90.pdf")
    seg = _ocr_seg(0, "Budapest, 2020.", [[60, 300, 160, 312]])
    out = tmp_path / "ru.pdf"
    build_pdf(str(tmp_path / "rot90.pdf"), str(out), [seg], {"0": {"ru": "Будапешт, 2020."}}, check=False)
    assert "Будапешт" in fitz.open(out)[0].get_text()
