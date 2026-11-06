# -*- coding: utf-8 -*-
"""Этап E7 (UNIFIED_DESIGN.md, П6): один классификатор фигур, снимок листа без колонтитула, формы-шаблоны, сверка
перевода поверх скана вторым распознаванием. Каждый тест падает на коде без своей правки."""
import io
import json

import fitz
import pytest

import docx_layout as DL
import figures as FIG


def seg(i, text, x0, y0, x1, y1=None, **loc):
    return {"id": i, "page": 0, "text": text, "loc": dict({"bbox": [x0, y0, x1, y1 or y0 + 10]}, **loc)}


# ---------------------------------------------------------------- флаг
def test_flag_default_on_and_zero_means_old_path(monkeypatch):
    monkeypatch.delenv("HU_U_FIG", raising=False)
    assert FIG.enabled()
    monkeypatch.setenv("HU_U_FIG", "0")
    assert not FIG.enabled()


# ---------------------------------------------------------------- схема: один набор признаков
def _arrows_page(with_arrows=True):
    page = fitz.open().new_page(width=595, height=842)
    sh = page.new_shape()
    xs, ys = (100, 198, 296), (200, 248, 296)                      # блоки 80×30 через 18 pt
    for y in ys:
        for x in xs:
            sh.draw_rect(fitz.Rect(x, y, x + 80, y + 30))
            sh.finish(color=(0, 0, 0), width=0.7)
    if with_arrows:                                                # связки: и горизонтальные, и вертикальные
        for y in ys:
            sh.draw_line((180, y + 15), (198, y + 15))
            sh.finish(color=(0, 0, 0), width=0.7)
        for y in (230, 278):
            sh.draw_line((140, y), (140, y + 18))
            sh.finish(color=(0, 0, 0), width=0.7)
    sh.commit()
    segs = [seg(i, "címke %d" % i, xs[i % 3] + 10, ys[i // 3] + 10, xs[i % 3] + 70) for i in range(9)]
    return page, segs


def test_scheme_needs_arrows_both_ways_and_captions():
    page, segs = _arrows_page()
    got = FIG.scheme_regions(page, 0, page.get_cdrawings(), segs, [], [], [])
    assert len(got) == 1 and got[0][1]["page_number"] == 1 and got[0][0].width > 250
    page, segs = _arrows_page(with_arrows=False)                   # рамки без связок — не схема (поля бланка, перечень)
    assert FIG.scheme_regions(page, 0, page.get_cdrawings(), segs, [], [], []) == []


def test_scheme_same_for_text_layer_and_recognized_page():
    """Раньше схему текстового слоя и распознанного листа искали разными признаками (_vector_figures / _scheme_figures):
    один и тот же лист находился по-разному. Теперь вызов один — masked пуст (распознанный) или нет (слой)."""
    page, segs = _arrows_page()
    drw = page.get_cdrawings()
    a = FIG.scheme_regions(page, 0, drw, segs, segs, [], [])
    b = FIG.scheme_regions(page, 0, drw, segs, [], [], [])
    assert [tuple(r) for r, _ in a] == [tuple(r) for r, _ in b] and len(a) == 1


def test_scheme_not_in_header_band_when_no_mask():
    """У распознанного листа колонтитула как сегментов нет: схема не набирается из линий рамки шапки/подвала."""
    page = fitz.open().new_page(width=595, height=842)
    sh = page.new_shape()
    for x in range(60, 540, 40):
        sh.draw_line((x, 30), (x, 60))
        sh.draw_line((x, 30), (x + 30, 30))
    sh.finish(color=(0, 0, 0), width=0.7)
    sh.commit()
    segs = [seg(0, "Fejléc", 60, 35, 100), seg(1, "Oldal", 200, 35, 240)]
    assert FIG.scheme_regions(page, 0, page.get_cdrawings(), segs, [], [], []) == []


# ---------------------------------------------------------------- классификатор листа
def test_figure_kind_scan_scheme_chart_none():
    page = fitz.open().new_page(width=842, height=595)
    big = [(fitz.Rect(20, 20, 800, 560), {})]
    labels = [seg(i, "címke", 100 + 40 * i, 200, 130 + 40 * i) for i in range(10)]
    assert FIG.figure_kind(page, labels, big, False, 0, True) == "scheme"
    ocr = [dict(s, loc=dict(s["loc"], ocr=True)) for s in labels]
    assert FIG.figure_kind(page, ocr, [], True, 0, True) == "scan"             # скан с распознанными надписями
    assert FIG.figure_kind(page, labels, [], False, 0.4 * page.rect.get_area(), True) == "chart"
    prose = [seg(i, "szöveg " * 20, 60, 100 + 14 * i, 780) for i in range(30)]
    assert FIG.figure_kind(page, prose, [(fitz.Rect(60, 300, 500, 560), {})], False, 0, True) is None
    assert FIG.figure_kind(page, ocr, [], True, 0, False) is None              # распознанный PDF: снимков листа нет


def test_figure_decisions_live_in_figures_only():
    """E9.1: прежние делегаты docx_layout (is_scan, _graphic_page, _ocr_layer, _chart_table) сняты — решение одно, в figures."""
    page = fitz.open().new_page(width=842, height=595)
    labels = [seg(i, "címke", 100 + 40 * i, 200, 130 + 40 * i) for i in range(10)]
    big = [(fitz.Rect(20, 20, 800, 560), {})]
    assert FIG.graphic_page(page, labels, big) is True
    assert FIG.is_scan(page) is False
    assert not any(hasattr(DL, n) for n in ("is_scan", "_graphic_page", "_ocr_layer", "_chart_table"))


# ---------------------------------------------------------------- снимок листа: без колонтитула оригинала
def test_snapshot_rect_stops_below_header_frame_and_above_footer():
    R = fitz.Rect(8, 57, 834, 578)                                   # снимок GF: шапка 34–59, линия рамки на 64
    FIG.clamp_zones(R, [[52, 34, 166, 47], [739, 45, 802, 59]], [], [(29, 42, 830), (64, 42, 830), (300, 10, 20)], 842)
    assert R.y0 == 66 and R.y1 == 578                                # ниже линии рамки шапки (64) + 2
    R = fitz.Rect(8, 57, 834, 578)
    FIG.clamp_zones(R, [[739, 45, 802, 59]], [], [(120, 42, 830)], 842)
    assert R.y0 == 61                                                # линия далеко (≥ FRAME_GAP) — рамка шапки не она
    R = fitz.Rect(8, 57, 834, 578)
    FIG.clamp_zones(R, [], [[300, 560, 400, 570]], [(555, 40, 800)], 842)
    assert R.y1 == 553
    R = fitz.Rect(8, 57, 834, 578)
    FIG.clamp_zones(R, [], [], [(64, 42, 830)], 842)
    assert tuple(R) == (8, 57, 834, 578)                             # колонтитулов нет — как есть


def _poster_pdf(path, poster=2):
    """Лист-плакат (рисунок во весь лист с подписями) под рамкой шапки с номером листа — как последний лист GF; вокруг —
    листы с той же шапкой (колонтитул — то, что повторяется)."""
    from PIL import Image
    png = io.BytesIO()
    Image.new("RGB", (40, 40), (60, 120, 200)).save(png, "PNG")
    doc = fitz.open()
    for n in range(4):
        pg = doc.new_page(width=842, height=595)
        pg.draw_rect(fitz.Rect(40, 30, 800, 64), color=(0, 0, 0), width=1)               # рамка шапки
        pg.insert_text((52, 44), "Azonosító: X", fontsize=9)
        pg.insert_text((739, 58), "Oldalszám: 1%d." % (7 + n), fontsize=9)
        if n == poster:
            for k in range(70):                                                         # блок-схема Word: десятки картинок
                x, y = 30 + (k % 10) * 78, 90 + (k // 10) * 60
                pg.insert_image(fitz.Rect(x, y, x + 70, y + 30), stream=png.getvalue())
                pg.insert_text((x + 4, y + 18), "szöveg", fontsize=7)
        else:
            pg.insert_text((60, 200), "Rendes szöveg a lapon.", fontsize=10)
    doc.save(path)


def test_sheet_snapshot_of_poster_excludes_header_frame(tmp_path, monkeypatch):
    f = tmp_path / "poster.pdf"
    _poster_pdf(f)
    segs = []
    for pno in (0, 1, 2, 3):
        segs.append({"id": len(segs), "page": pno, "text": "Azonosító: X",
                     "loc": {"bbox": [52, 36, 120, 46], "frags": [[52, 36, 120, 46]], "size": 9, "family": "serif"}})
        segs.append({"id": len(segs), "page": pno, "text": "Oldalszám: 1%d." % (7 + pno),
                     "loc": {"bbox": [739, 50, 802, 60], "frags": [[739, 50, 802, 60]], "size": 9, "family": "serif"}})
        if pno != 2:
            segs.append({"id": len(segs), "page": pno, "text": "Rendes szöveg a lapon.",
                         "loc": {"bbox": [60, 190, 180, 202], "frags": [[60, 190, 180, 202]], "size": 10,
                                 "family": "serif"}})
    for k in range(70):
        x, y = 30 + (k % 10) * 78, 90 + (k // 10) * 60
        segs.append({"id": len(segs), "page": 2, "text": "szöveg",
                     "loc": {"bbox": [x + 4, y + 10, x + 40, y + 22], "frags": [[x + 4, y + 10, x + 40, y + 22]],
                             "size": 7, "family": "serif"}})
    clips = {}
    for flag in ("0", "1"):
        monkeypatch.setenv("HU_U_FIG", flag)
        blocks, info = DL.load(str(f), segs, text=True)
        figs = [b for b in blocks if b.get("figure")]
        assert len(figs) == 1, flag
        clips[flag] = figs[0]["runs"][0]["img"]["clip"]
    assert clips["0"][1] < 64 < clips["1"][1]                        # раньше — с обрывком шапки, теперь — ниже её рамки
    assert clips["1"][1] == 66 and clips["0"][3] == clips["1"][3]


# ---------------------------------------------------------------- формы-шаблоны
def test_known_form_by_short_title_in_short_document_only():
    title = lambda t, page=0: {"id": 0, "page": page, "text": t, "loc": {}}
    assert FIG.known_form([title("NYILVÁNTARTÓ LAP")]) == "nyilv"
    assert FIG.known_form([title("OPERATÍV ÜZEMVITELI TESZTELÉSI UTASÍTÁS")]) == "ciml"
    assert FIG.known_form([title("MÓDOSÍTÁS JEGYZÉK")]) == "modj"
    assert FIG.known_form([title("Lásd a nyilvántartó lap kitöltési útmutatóját az 5.3. pontban, a mellékletek szerint.")]) is None
    assert FIG.known_form([title("MÓDOSÍTÁS JEGYZÉK"), title("szöveg", 2)]) is None   # раздел длинной инструкции — не бланк


def test_form_labels_applied_but_not_over_editor_tm_or_copied(monkeypatch):
    monkeypatch.setitem(FIG._labels, "d", {"nyilv": {"Dátum:": "Дата:", "igen": "да"}})
    segs = [{"id": i, "page": 0, "text": t, "loc": {}} for i, t in
            enumerate(["NYILVÁNTARTÓ LAP", "Dátum: |", "Dátum:", "Dátum:", "igen", "igen"])]
    res = {"0": {"ru": "ЛИСТ РЕГИСТРАЦИИ"}, "1": {"ru": "Дата: |"}, "2": {"ru": "Дата:", "edited": True},
           "3": {"ru": "Число:", "tm": True}, "4": {"ru": "Да"}, "5": {"ru": "igen", "copied": True}}
    out, n, form = FIG.apply_form_labels(segs, res)
    assert form == "nyilv" and n == 2
    assert out["1"]["ru"] == "Дата:" and out["4"]["ru"] == "да"                 # «|» по краю и вариант регистра
    assert out["2"]["ru"] == "Дата:" and out["3"]["ru"] == "Число:" and out["5"]["ru"] == "igen"
    assert res["1"]["ru"] == "Дата: |"                                           # исходный словарь не меняется
    assert FIG.apply_form_labels([segs[1]], res)[1] == 0                         # не форма — ничего


def test_form_labels_file_is_consistent():
    labels = json.loads(FIG.LABELS_FILE.read_text(encoding="utf-8"))
    assert set(labels) == {"ciml", "modj", "nyilv"}
    for form, rows in labels.items():
        for hu, ru in rows.items():
            assert hu == FIG.label_key(hu) and ru == FIG.label_key(ru) and FIG.is_label(hu), (form, hu)
            assert "|" not in ru and ru.strip()


# ---------------------------------------------------------------- перевод поверх скана
def test_drawn_text_drops_cell_line_artifacts_only_at_edges():
    assert FIG.drawn_text("Дата: 2022-01-11 |") == "Дата: 2022-01-11"
    assert FIG.drawn_text("| Новые страницы |") == "Новые страницы"
    assert FIG.drawn_text("а | б") == "а | б"                                     # посреди строки — текст
    assert FIG.drawn_text("|") == "|"


def test_codes_and_numbers_checked_with_ocr_confusions():
    want = FIG.codes_of("Идентификатор: 30YCGT, дата 2022-01-11, код 31UJ21S601 и 36 бар (UT_ÜVIG_25)")
    assert want == {FIG._tok(t) for t in ("30YCGT", "2022-01-11", "31UJ21S601", "36", "UT_ÜVIG_25")} and len(want) == 5
    seen = ["30YCGT", "2022-01-1I", "3IUJ21S6O1", "36", "UT_ÜVIG_25"]            # I/1, O/0 путает и распознавание
    assert FIG.missing_codes(want, seen) == []
    assert FIG.missing_codes(want, ["30YCGT", "36"]) == sorted(want - {FIG._tok("30YCGT"), "36"})


def test_verify_page_loud_only_for_real_loss_minor_as_note(monkeypatch):
    page = fitz.open().new_page()
    drawn = ["Идентификатор: 30YCGT", "Дата: 2022-01-11", "Блок 3-В, лист 55"]
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["Дата:"])                  # все четыре кода потеряны
    msg, loud = FIG.verify_page(page, drawn, 2)
    assert loud and "лист 3" in msg and "2022-01-11" in msg and "30YCGT" in msg and "4 из 4" in msg
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["30YCGT", "Дата:", "55", "3-B"])   # потеряна одна дата
    msg, loud = FIG.verify_page(page, drawn, 2)
    assert not loud and "1 из 4" in msg and "2022-01-11" in msg
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["30YCGT", "55"])               # два из четырёх — ещё сведения
    assert FIG.verify_page(page, drawn, 2)[1] is False
    many = ["код %d%d" % (i, i) for i in range(10)]
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: [t for t in "11 22 33 44 55 66 77 88 99".split()])
    assert FIG.verify_page(page, many, 0)[1] is False                                  # одно слово из десяти — сведения
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["11", "22", "33", "44", "55", "66"])
    assert FIG.verify_page(page, many, 0)[1] is True                                   # четыре из десяти — громко
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["30YCGT", "2022-01-11", "3-B", "55"])
    assert FIG.verify_page(page, drawn, 2) is None
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: None)                        # нет Tesseract — сверка пропущена
    assert FIG.verify_page(page, ["30YCGT"], 0) is None
    assert FIG.verify_page(page, ["Дата:", "Подпись:"], 0) is None                      # кодов в переводе нет
    big = "A1" * 20000
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: [])
    assert len(FIG.verify_page(page, [big], 0)[0]) < 200                                # токен в сообщении обрезан


def test_build_pdf_scan_page_strips_cell_line_and_reports_lost_code(tmp_path, monkeypatch):
    import numpy as np
    from PIL import Image
    from assemble import build_pdf
    k = 150 / 72
    buf = io.BytesIO()
    Image.new("RGB", (round(595 * k), round(842 * k)), "white").save(buf, "PNG")
    src, out = tmp_path / "scan.pdf", tmp_path / "ru.pdf"
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(src)
    s = {"id": 0, "page": 0, "text": "Dátum: 2022-01-11 |",
         "loc": {"bbox": [60, 300, 180, 312], "avail": [60, 300, 520, 312], "frags": [[60, 300, 180, 312]], "size": 10,
                 "bold": False, "family": "serif", "ocr": True}}
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **kw: ["Дата:"])            # дата на листе второй раз не прочитана
    info = build_pdf(str(src), str(out), [s], {"0": {"ru": "Дата: 2022-01-11 |"}})
    assert fitz.open(out)[0].get_text().strip() == "Дата: 2022-01-11"
    assert "2022-01-11" in info["verify_notes"][0] and "verify" not in info        # одно слово из одного — сведения, не тревога
    s2 = dict(s, id=1, text="Verzió 30YCGT 2022-01-11 10135772", loc=dict(s["loc"], bbox=[60, 330, 180, 342],
              avail=[60, 330, 520, 342], frags=[[60, 330, 180, 342]]))
    info = build_pdf(str(src), str(out), [s, s2], {"0": {"ru": "Дата: 2022-01-11"},
                                                    "1": {"ru": "Версия 30YCGT 2022-01-11 10135772"}})
    assert info["verify"] and "30YCGT" in info["verify"][0]                        # три из трёх — громко
    monkeypatch.setenv("HU_U_FIG", "0")                                         # выключено — как раньше: «|» рисуется, сверки нет
    info = build_pdf(str(src), str(out), [s], {"0": {"ru": "Дата: 2022-01-11 |"}})
    assert "|" in fitz.open(out)[0].get_text() and "verify" not in info


def test_split_checkboxes_cuts_mark_letters_and_keeps_articles_and_edits():
    seg_ = lambda i, t: {"id": i, "text": t, "loc": {"ocr": True, "frags": [[100, 10, 160, 20]],
                                                    "bbox": [100, 10, 160, 20], "avail": [100, 10, 300, 20]}}
    segs = [seg_(0, "CI nem"), seg_(1, "a nem"), seg_(2, "X IGEN"), seg_(3, "igen"), seg_(4, "Fr igen"), seg_(5, "KI igen")]
    res = {"0": {"ru": "CI нет"}, "1": {"ru": "не"}, "2": {"ru": "X ДА"}, "3": {"ru": "да"},
           "4": {"ru": "Фр да", "edited": True}, "5": {"ru": "ОТКЛ. да"}}
    s2, r2 = FIG.split_checkboxes(segs, res)
    assert [x["text"] for x in s2] == ["nem", "a nem", "IGEN", "igen", "Fr igen", "igen"]
    assert {k: v["ru"] for k, v in r2.items()} == {"0": "нет", "1": "не", "2": "ДА", "3": "да", "4": "Фр да", "5": "да"}
    assert s2[0]["loc"]["frags"][0][0] == pytest.approx(126) and s2[0]["loc"]["avail"][0] == pytest.approx(126)
    assert s2[1]["loc"]["frags"][0][0] == 100                                      # «a nem» — слова, не флажок
    assert segs[0]["text"] == "CI nem" and res["0"]["ru"] == "CI нет"             # исходные не меняются


def test_build_pdf_keeps_checkbox_pixels_and_draws_only_yes_no(tmp_path, monkeypatch):
    import numpy as np
    from PIL import Image, ImageDraw
    from assemble import build_pdf
    k = 150 / 72
    im = Image.new("RGB", (round(595 * k), round(842 * k)), "white")
    ImageDraw.Draw(im).rectangle([100 * k, 301 * k, 108 * k, 309 * k], outline="black", width=3)     # флажок «☐»
    buf = io.BytesIO()
    im.save(buf, "PNG")
    src, out = tmp_path / "scan.pdf", tmp_path / "ru.pdf"
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(src)
    s = {"id": 0, "page": 0, "text": "CI nem",
         "loc": {"bbox": [100, 300, 160, 312], "avail": [100, 300, 400, 312], "frags": [[100, 300, 160, 312]], "size": 10,
                 "bold": False, "family": "serif", "ocr": True}}
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **kw: None)
    for flag in ("1", "0"):
        monkeypatch.setenv("HU_U_FIG", flag)
        out = tmp_path / ("ru%s.pdf" % flag)
        build_pdf(str(src), str(out), [s], {"0": {"ru": "CI нет"}})
        q = fitz.open(out)[0]
        pm = q.get_pixmap(dpi=150, clip=fitz.Rect(99, 300, 109, 310), colorspace=fitz.csGRAY)
        dark = int((np.frombuffer(pm.samples, np.uint8) < 100).sum())
        text = q.get_text().strip()
        if flag == "1":
            assert dark > 20 and text == "нет"                                  # флажок цел, «CI» не рисуется
        else:
            assert text == "CI нет"                                              # прежнее: на месте флажка — буквы «CI»


# ---------------------------------------------------------------- по ревью: схема, формы, сверка
def test_scheme_grid_equals_brute_force_pairs_and_is_fast():
    import random
    import time
    rnd = random.Random(7)
    items = []
    for _ in range(400):
        x, y = rnd.uniform(0, 500), rnd.uniform(0, 780)
        items.append((rnd.random() < 0.8, fitz.Rect(x, y, x + rnd.uniform(1, 90), y + rnd.uniform(1, 30))))
    brute = {(i, j) for i in range(len(items)) for j in range(i + 1, len(items))
             if FIG._linked(*items[i], *items[j])}
    assert brute and brute <= set(FIG._candidates(items))                  # сетка ничего не теряет
    page = fitz.open().new_page(width=595, height=842)
    sh = page.new_shape()
    for _ in range(5000):                                                    # лист-карта: тысячи мелких фигур
        x, y = rnd.uniform(0, 560), rnd.uniform(0, 800)
        sh.draw_rect(fitz.Rect(x, y, x + 3, y + 3))
        sh.finish(color=(0, 0, 0), width=0.2)
    sh.commit()
    t = time.time()
    FIG.scheme_regions(page, 0, page.get_cdrawings(), [seg(0, "a", 10, 10, 30), seg(1, "b", 50, 10, 70)], [], [], [])
    assert time.time() - t < 10                                              # было 109 с на 6000 фигур (прямой перебор пар)


def test_text_near_scheme_does_not_join_or_spoil_it():
    """Подводка-абзац над схемой и тело под ней — не подписи схемы: не втягиваются в группу и не отбрасывают её."""
    page, segs = _arrows_page()
    n = len(segs)
    segs += [seg(n, "Hosszú bevezető szöveg " * 8, 96, 170, 500),                      # 190 знаков, в 20 pt над схемой
             seg(n + 1, "Alatta a törzs első sora " * 3, 96, 345, 480),
             seg(n + 2, "Alatta a törzs második sora " * 3, 96, 358, 480)]
    got = FIG.scheme_regions(page, 0, page.get_cdrawings(), segs, segs, [], [])
    assert len(got) == 1 and got[0][0].y0 > 190 and got[0][0].y1 < 340


def test_clamp_zones_leaves_rect_alone_when_nothing_would_remain():
    R = fitz.Rect(8, 57, 834, 70)
    FIG.clamp_zones(R, [[739, 45, 802, 59]], [[300, 62, 400, 70]], [], 842)
    assert (R.y0, R.y1) == (57, 70)                                           # от снимка осталась бы полоска — как есть


def test_known_form_needs_known_pages_and_one_form():
    t = lambda text, page: {"id": 0, "page": page, "text": text, "loc": {}}
    assert FIG.known_form([t("NYILVÁNTARTÓ LAP", None), t("szöveg", None)]) is None          # DOCX: листов не знаем
    assert FIG.known_form([t("NYILVÁNTARTÓ LAP", 0), t("OPERATÍV ÜZEMVITELI TESZTELÉSI UTASÍTÁS", 1)]) is None


def test_form_labels_keep_polished_and_shortened_translations(monkeypatch):
    monkeypatch.setitem(FIG._labels, "d", {"nyilv": {"Dátum:": "Дата:"}})
    segs = [{"id": i, "page": 0, "text": t, "loc": {}} for i, t in enumerate(["NYILVÁNTARTÓ LAP", "Dátum:", "Dátum:"])]
    res = {"0": {"ru": "ЛИСТ РЕГИСТРАЦИИ"}, "1": {"ru": "Число:", "polish": {"of": "Число:"}},
           "2": {"ru": "Числа:", "ru_pdf": "Ч.:", "ru_pdf_of": "Числа:"}}
    assert FIG.apply_form_labels(segs, res)[1] == 0


def test_codes_survive_cyrillic_lookalikes_and_are_deterministic():
    assert FIG._tok("B2") == FIG._tok("В2")                                   # латинская и кириллическая «В»
    assert FIG.missing_codes(FIG.codes_of("Код B2"), ["В2"]) == []
    assert FIG.code_map("Код 31UJ21S601")["31UJ215601"] == "31UJ21S601"        # в сообщении — как в тексте
    for _ in range(3):
        assert FIG.missing_codes({"36"}, ["Лист", "A3", "6B", "12"]) == ["36"]  # раньше зависело от PYTHONHASHSEED
    assert FIG.missing_codes({"36"}, ["2036"]) == ["36"]                       # «36» — не часть «2036»
    assert FIG.missing_codes({"30YCGT"}, ["30Y", "CGT"]) == []                 # длинный код, распознан частями


def test_verify_page_survives_tesseract_errors(monkeypatch):
    page = fitz.open().new_page()
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: (_ for _ in ()).throw(ValueError("tsv")))
    assert FIG.verify_page(page, ["30YCGT 2022-01-11"], 0) is None


def test_checkbox_split_only_for_russian_translation():
    seg_ = {"id": 0, "text": "CI nem", "loc": {"ocr": True, "frags": [[100, 10, 160, 20]], "bbox": [100, 10, 160, 20]}}
    s2, r2 = FIG.split_checkboxes([seg_], {"0": {"ru": "CI no"}})              # перевод не русский — «да/нет» не пишем
    assert s2[0]["text"] == "CI nem" and r2["0"]["ru"] == "CI no"
    s2, _ = FIG.split_checkboxes([dict(seg_, text="Ki nem")], {"0": {"ru": "Кто не"}})
    assert s2[0]["text"] == "Ki nem"                                           # «Ki nem» — слова
    s2, _ = FIG.split_checkboxes([dict(seg_, text="KI igen")], {"0": {"ru": "ОТКЛ. да"}})
    assert s2[0]["text"] == "igen"                                             # «KI» заглавными — флажок


def test_clamp_zones_keeps_body_title_between_header_text_and_line():
    """Заголовок рисунка под шапкой с подчёркиванием — тело листа, а не рамка шапки: снимок остаётся с ним."""
    R = fitz.Rect(8, 45, 834, 578)
    FIG.clamp_zones(R, [[52, 34, 166, 47]], [], [(58, 40, 800)], 842, body=[[52, 49, 200, 56]])
    assert R.y0 == 49                                                           # только над текстом шапки (47 + 2)
    R = fitz.Rect(8, 45, 834, 578)
    FIG.clamp_zones(R, [[52, 34, 166, 47]], [], [(58, 40, 800)], 842, body=[[52, 80, 200, 90]])
    assert R.y0 == 60                                                           # тело ниже линии — линия была рамкой шапки


def test_checkbox_edge_survives_when_gap_is_small(tmp_path, monkeypatch):
    """Флажок 8×8 pt, зазор 3 pt, слово «igen»: плашка перевода не задевает границу флажка (красная команда)."""
    import numpy as np
    from PIL import Image, ImageDraw
    from assemble import build_pdf
    k = 150 / 72
    im = Image.new("RGB", (round(595 * k), round(842 * k)), "white")
    dr = ImageDraw.Draw(im)
    dr.rectangle([100 * k, 301 * k, 108 * k, 309 * k], outline="black", width=2)
    dr.rectangle([111 * k, 302 * k, 138 * k, 308 * k], fill="black")                 # слово «igen» (27 pt)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    src = tmp_path / "scan.pdf"
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(src)
    s = {"id": 0, "page": 0, "text": "CI igen",
         "loc": {"bbox": [100, 300, 138, 310], "avail": [100, 300, 400, 310], "frags": [[100, 300, 138, 310]], "size": 10,
                 "bold": False, "family": "serif", "ocr": True}}
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **kw: None)
    build_pdf(str(src), str(tmp_path / "ru.pdf"), [s], {"0": {"ru": "CI да"}})
    pm = fitz.open(tmp_path / "ru.pdf")[0].get_pixmap(dpi=150, clip=fitz.Rect(99, 300, 109.5, 310), colorspace=fitz.csGRAY)
    a = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width)
    cols = (a < 100).sum(0)
    assert cols[:4].max() >= 6 and cols[15:20].max() >= 6                      # обе границы рамки флажка (x=100 и 108) целы


def test_short_phrases_are_not_checkboxes_and_bars_are_linear():
    import time
    seg_ = lambda t: {"id": 0, "text": t, "loc": {"ocr": True, "frags": [[100, 10, 160, 20]], "bbox": [100, 10, 160, 20]}}
    for t, ru in (("Én nem", "Я не"), ("ő nem", "Он не"), ("se nem", "ни не")):
        assert FIG.split_checkboxes([seg_(t)], {"0": {"ru": ru}})[0][0]["text"] == t, t   # фраза, а не флажок
    s2, r2 = FIG.split_checkboxes([seg_("CI Igen")], {"0": {"ru": "CI да"}})
    assert r2["0"]["ru"] == "Да"                                               # регистр слова сохранён
    t0 = time.time()
    assert FIG.drawn_text("x" + " |" * 40000 + "x") == "x" + " |" * 40000 + "x" and time.time() - t0 < 2
    assert FIG.label_key("| | Név: | |") == "Név:"


def test_russian_words_with_digit_are_not_codes_to_verify(monkeypatch):
    """«Коды и числа» для сверки скана — слова с цифрой, кроме русских слов («7-барной», «3-го»): второе распознавание
    читает обычные слова иначе, и три промаха из десяти давали ложное «ВНИМАНИЕ» (30RDGT nyilv 06.10; ревью 03.11, 🟡4).
    Коды, числа, даты и русские обозначения заглавными («ТЭН-3») остаются."""
    text = "Манометр 7-барной линии 3-го блока, 30RDGT, 140387, 2017.04.20, 10:30, ТЭН-3, 5-й и 10-кратный"
    assert FIG.codes_of(text) == {FIG._tok(t) for t in ("30RDGT", "140387", "2017.04.20", "10:30", "ТЭН-3")}
    page = fitz.open().new_page()
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **k: ["30RDGT", "140387", "2017.04.20", "10:30", "ТЭН-3"])
    assert FIG.verify_page(page, [text], 0) is None                                   # слов с цифрой в кодах нет — не искали


def test_build_pdf_step_failure_is_loud_and_page_is_still_built(tmp_path, monkeypatch, capsys):
    """Новые шаги листа скана (флажки-буквы, чужая краска, край краски слова) вне try роняли build_pdf, а с ним и сборку
    всего задания (ревью 03.11, 🟡2). Сбой шага — громко (лог с листом и причиной, info["steps_failed"]), лист собирается
    без шага; один и тот же сбой на листе не повторяется строкой на каждую строку."""
    from PIL import Image
    import assemble as A
    from assemble import build_pdf
    k = 150 / 72
    buf = io.BytesIO()
    Image.new("RGB", (round(595 * k), round(842 * k)), "white").save(buf, "PNG")
    src, out = tmp_path / "scan.pdf", tmp_path / "ru.pdf"
    d = fitz.open()
    d.new_page(width=595, height=842).insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    d.save(src)
    mk = lambda i, y: {"id": i, "page": 0, "text": "AZONOSÍTÓ:", "loc": {
        "bbox": [60, y, 100, y + 12], "avail": [60, y, 520, y + 12], "frags": [[60, y, 100, y + 12]], "size": 10,
        "bold": False, "family": "monospace", "ocr": True}}
    segs = [mk(0, 300), mk(1, 330)]                              # моноширинный, слово шире рамки: _ink_span вызывается
    res = {"0": {"ru": "ИДЕНТИФИКАТОР:"}, "1": {"ru": "ИДЕНТИФИКАТОР:"}}
    monkeypatch.setattr(FIG, "ocr_words", lambda pg, **kw: None)

    def boom(*a, **kw):
        raise ValueError("сбой шага")
    monkeypatch.setattr(FIG, "split_checkboxes", boom)
    monkeypatch.setattr(A, "_foreign_ink", boom)
    monkeypatch.setattr(A, "_ink_span", boom)
    info = build_pdf(str(src), str(out), segs, res)
    assert fitz.open(out)[0].get_text().split() == ["ИДЕНТИФИКАТОР:", "ИДЕНТИФИКАТОР:"]      # перевод на листе есть
    fails = info["steps_failed"]
    assert len(fails) == 3                                                                 # _ink_span — раз на лист
    assert all("ValueError: сбой шага" in f for f in fails)
    assert any("split_checkboxes" in f for f in fails) and any("_foreign_ink" in f and "лист 1" in f for f in fails)
    assert any("_ink_span" in f and "лист 1" in f for f in fails)
    log = capsys.readouterr().out
    assert "Traceback" in log and "_foreign_ink" in log and "scan.pdf, лист 1" in log          # в лог — причина и место


def test_form_label_replaces_translation_whose_short_form_was_made_for_the_label():
    """Сокращение для PDF, сделанное pipeline.fit для подписи по словарю формы (ru_pdf_of — подпись), перевод записи не
    защищает: подпись по словарю подставляется, и сокращение к ней подходит (ревью 03.11, 🟡3); своё сокращение — защищает."""
    import assemble as A
    segs = [{"id": i, "page": 0, "text": t, "loc": {}} for i, t in enumerate(["NYILVÁNTARTÓ LAP", "Dátum:", "Dátum:"])]
    res = {"0": {"ru": "ЛИСТ РЕГИСТРАЦИИ"},
           "1": {"ru": "Число:", "ru_pdf": "Дат.:", "ru_pdf_of": "Дата:"},                # сделано для подписи словаря
           "2": {"ru": "Числа:", "ru_pdf": "Ч.:", "ru_pdf_of": "Числа:"}}                 # своё — защищает
    out, n, form = FIG.apply_form_labels(segs, res)
    assert form == "nyilv" and n == 1 and out["2"]["ru"] == "Числа:"
    assert out["1"]["ru"] == "Дата:" and A.pdf_text(out["1"]) == "Дат.:"
