# -*- coding: utf-8 -*-
"""Дефекты сканов по просмотру глазами 06.11 (EYES_2026-11-06: N6, п. 5, 13, штамп 30RYBSZGT, красная плашка формы). Слова — как их
отдаёт Tesseract (segments._tess_words): t, box, h, line, key, c (уверенность), ink (плотность тёмной краски), ink_c (красная краска)."""
import io

import fitz
import numpy as np
from PIL import Image

import segments as S
from assemble import build_pdf


def word(t, x, y, w=30, h=6, line=(1, 1, 1), c=95, ink=0.2, **kw):
    return dict({"t": t, "box": [x, y, x + w, y + h], "h": h, "line": line, "key": line[:2], "c": c, "ink": ink}, **kw)


def test_red_print_is_not_handwriting():
    """Красная печать (предупреждение «AZ UT_ÜVIG_25 … FIGYELEMBEVÉTELÉVEL KÉSZÜLT!», красные заголовки форм) Tesseract читает
    уверенно (89–97), а тёмной краски в сером изображении у неё почти нет (красный (251,61,49) светлее порога): слово шло как
    «бледная рукопись», не переводилось и не закрывалось плашкой — венгерский остался на листе. Красная краска — тоже краска."""
    red = [word("AZ", 150, 72, w=11, h=6.5, c=96, ink=0.0, ink_c=0.38),
           word("UTASÍTÁS", 229, 70, w=48, h=8.6, c=93, ink=0.041, ink_c=0.27),
           word("FIGYELEMBEVÉTELÉVEL", 281, 70, w=126, h=9.1, c=93, ink=0.024, ink_c=0.27)]
    hand = [word("Vale", 50, 10, c=0, ink=0.05, ink_c=0.0), word("Tamási", 10, 10, c=96, ink=0.21)]
    out, _ = S.ocr_words(red + hand, "")
    assert [w["t"] for w in out] == ["AZ", "UTASÍTÁS", "FIGYELEMBEVÉTELÉVEL", "Tamási"]        # рукопись — по-прежнему не слово


def test_ink_red_counts_only_saturated_red_pixels():
    k = 300 / 72
    img = np.full((int(60 * k), int(200 * k), 3), 255, np.uint8)
    img[int(10 * k):int(20 * k), int(10 * k):int(60 * k)] = (251, 61, 49)                 # красный блок
    img[int(10 * k):int(20 * k), int(100 * k):int(150 * k)] = (40, 60, 170)               # синяя паста
    img[int(30 * k):int(40 * k), int(10 * k):int(60 * k)] = (132, 132, 132)               # серая печать
    ink = lambda x0, x1, y0=8, y1=22: S._ink_red(img, [x0, y0, x1, y1], 300)
    assert ink(10, 60) > 0.5 and ink(100, 150) == 0.0 and ink(10, 60, 28, 42) == 0.0


def test_cell_rule_read_as_letters_is_not_a_word():
    """Вертикальная линейка ячейки таблицы версий, прочитанная как «[d» (рамка 2,4 × 4,6 pt, краски 0,64, уверенность 92): знаки
    в такую щель не вмещаются — слово шло в перевод («[d Страница» на титуле 3PR10, п. 5 EYES). Теперь это картинка (линия остаётся)."""
    ws = [word("[d", 240.5, 373.9, w=2.4, h=4.6, c=92, ink=0.637), word("Oldal", 247.9, 373.7, w=20.7, h=4.8, c=96, ink=0.3)]
    pix = []
    out, _ = S.ocr_words(ws, "", pix)
    assert [w["t"] for w in out] == ["Oldal"] and pix == [[240.5, 373.9, 242.9, 378.5]]
    ok = [word("Id", 10, 10, w=8, h=5, c=90, ink=0.3), word("il", 30, 10, w=4, h=7, c=90, ink=0.3)]   # настоящие короткие слова
    assert [w["t"] for w in S.ocr_words(ok, "")[0]] == ["Id", "il"]


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


def test_overlapping_scan_segments_are_a_stamp_and_stay_pixels(tmp_path):
    """Штамп/наклонная надпись: Tesseract читает его мусором («ömű Zrt. Minőségfelügyeleti 97», «pakol Atomer tégy»), рамки двух
    сегментов лежат друг в друге — плашки стирали штамп, а два перевода рисовались друг на друге («АЭС ЗАО. Контроль качества 97» поверх
    «АЭС «Пакш»», 30RYBSZGT PG-144760 стр. 1, п. 4 EYES). Рамки строк, перекрывающие рамки чужих, — не текст листа: остаются пикселями."""
    src = tmp_path / "scan.pdf"
    _scan_pdf(src, stamp=(391, 745, 500, 778))
    segs = [_ocr_seg(0, "ömű Zrt. Minőségfelügyeleti 97", [[391, 745, 500, 778]], 11.1),
            _ocr_seg(1, "pakol Atomer tégy", [[408, 752, 497, 765]], 11.1),
            _ocr_seg(2, "Budapest, 2020.", [[60, 300, 160, 312]])]
    res = {"0": {"ru": "АЭС ЗАО. Контроль качества 97"}, "1": {"ru": "АЭС «Пакш»"}, "2": {"ru": "Будапешт, 2020."}}
    out = tmp_path / "ru.pdf"
    build_pdf(str(src), str(out), segs, res, check=False)
    page = fitz.open(out)[0]
    assert "Будапешт" in page.get_text() and "Пакш" not in page.get_text() and "Контроль" not in page.get_text()
    pm = page.get_pixmap(dpi=72, clip=fitz.Rect(391, 745, 500, 778), colorspace=fitz.csGRAY)
    assert int((np.frombuffer(pm.samples, np.uint8) < 100).mean() * 100) > 90                       # штамп цел, не закрыт плашкой


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
