# -*- coding: utf-8 -*-
"""Ворота качества готовых переводов (tools/audit_out.py): каждая проверка на малом примере + сквозной прогон по
собранному в tmp_path заданию (JSON, DOCX, PDF перевода, PDF оригинала). Данных проекта и Docker не требуется."""
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import audit_out as A  # noqa: E402

FONT = next((f for f in (r"C:\Windows\Fonts\arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf") if Path(f).exists()), None)
needs_font = pytest.mark.skipif(FONT is None, reason="нет шрифта с кириллицей для синтетических PDF")


def W(x0, y0, x1, y1, t):
    return (x0, y0, x1, y1, t, 0, 0, 0)


# ---------------------------------------------------------------- 2. повтор фразы
def test_repeat_phrase_three_times():
    r = A.find_repeat("МАШИННЫЙ ЗАЛ ряд 28КА " * 3 + "конец")
    assert r == ("МАШИННЫЙ ЗАЛ ряд 28КА", 3)


def test_repeat_twice_is_not_a_defect():
    assert A.find_repeat("МАШИННЫЙ ЗАЛ ряд 28КА МАШИННЫЙ ЗАЛ ряд 28КА конец") is None


def test_repeat_ignores_marks_and_short_words():
    assert A.find_repeat("○ ○ ○ ○ ○ ○ ○ ○ ○ ○") is None            # стопка значков — дело проверки 6
    assert A.find_repeat("Да Да Да Да") is None                     # короче 10 знаков
    assert A.find_repeat("0,0 м 0,0 м 0,0 м 0,0 м") is None         # числа без букв-слов


def test_repeat_inside_long_text():
    r = A.find_repeat("Начало абзаца. " + "30UY01/1 закрыт, " * 4 + "дальше текст")
    assert r and r[1] == 4 and "закрыт" in r[0]


def test_repeat_by_table_rows_is_not_a_defect():
    """Одинаковые значения по рядам таблицы (INFO 3PR10: «A301» в 40 рядах, каждое — свой сегмент и свой абзац ячейки)
    — не повтор фразы; значения, слитые нарезкой в один сегмент, и размноженное сборкой значение — повтор."""
    rows = ["A301"] * 40 + ["GA01"]
    u = " ".join(rows)
    assert A.find_repeat(u)[1] == 13                                 # без сегментов — повтор, как прежде
    assert A.find_repeat(u, legit=A.row_values(rows, {"a301": 40})) is None
    assert A.find_repeat(u, legit=A.row_values(rows, {"a301": 5})) is not None     # сегментов меньше рядов
    merged = ["126-128 бар " * 6, "x"]                                # нарезка слила шесть рядов в один сегмент
    assert A.find_repeat(" ".join(merged), legit=A.row_values(merged, {A.norm(merged[0]): 1, "126128бар": 50}))[1] == 6
    sig = ["сигнал 11GB02 сигнал 11GB02 сигнал 11GB02"] * 7         # в каждом абзаце — по три ряда
    assert A.find_repeat(" ".join(sig), legit=A.row_values(sig, {A.norm(sig[0]): 7}))[1] == 21
    tail = ["A301"] * 9 + ["Привод крана Привод крана Привод крана"]  # законные ряды пропущены, поиск идёт дальше
    r = A.find_repeat(" ".join(tail), legit=A.row_values(tail, {"a301": 9}))
    assert r == ("Привод крана", 3)
    codes = ["JM03 JM03 JM03"] * 3                                    # в каждом абзаце — три ряда, слитые нарезкой
    assert A.find_repeat(" ".join(codes), legit=A.row_values(codes, {A.norm(codes[0]): 3})) is not None


def test_row_counts_from_the_cells_own_sheets():
    """Сегменты считаются на листах самой ячейки (окно из двух соседних листов, где её абзацы находятся лучше всего), а
    не где угодно в документе: размноженное сборкой «Не используется» ×3 не прячется за тремя такими строками на
    других листах (ред-тим 02.11)."""
    page_n = {"a301": {30: 40}, "ga01": {30: 1}, "неиспользуется": {3: 1, 12: 3}, "клапан": {3: 1}}
    assert A.cell_counts(["A301"] * 40 + ["GA01"], page_n) == {"a301": 40, "ga01": 1}
    cnt = A.cell_counts(["Клапан"] + ["Не используется"] * 3, page_n)
    assert cnt == {"клапан": 1, "неиспользуется": 1}


# ---------------------------------------------------------------- 1. потеря содержимого
def seg(i, text, page=0, **loc):
    return {"id": i, "text": text, "page": page, "loc": loc}


def test_lost_segments_found_and_filtered():
    segs = [seg(1, "Zárva vannak az alábbi armatúrák:"), seg(2, "Nyitás (bar)"), seg(3, "Azonosító: 3SZ23"),
            seg(4, "Lapszám: 1."), seg(5, "30RR21S002"), seg(6, "Hőmérséklet", role="header"),
            seg(7, "Állapot tölt.", shot={"t": "x"}), seg(8, "Oldalszám: 4"), seg(9, "Szabályozó szelep zárva", ocr=True),
            seg(10, "Új javítás szükséges", 4)]
    res = {"1": {"ru": "Следующая арматура закрыта:"}, "2": {"ru": "Открытие (бар)"}, "3": {"ru": "Идентификатор: 3SZ23"},
           "4": {"ru": "Лист: 1."}, "5": {"ru": "30RR21S002", "copied": True}, "6": {"ru": "Температура"},
           "7": {"ru": "Сост. заправки"}, "8": {"ru": "Номер страницы: 4"}, "9": {"ru": "Регулирующий клапан закрыт"},
           "10": {"ru": "Требуется новый ремонт"}}
    full = A.norm("Открытие (бар) и ещё текст")                     # нашлось только сегмент 2
    lost = A.lost_segments(segs, res, full)
    assert [s["id"] for s in lost] == [1, 10]                       # заголовок шапки, копия, колонтитул, скрин, скан не в счёт


def test_lost_segment_found_by_polished_text():
    segs = [seg(1, "Ismételt új „0” verziós KU kiadása")]
    res = {"1": {"ru": "Повторный выпуск ИЭ новой версии „0“"}}
    full = A.norm("Повторный выпуск ИЭ новой версии «0»")
    assert A.lost_segments(segs, res, full) == []                   # кавычки не отличают
    assert [s["id"] for s in A.lost_segments(segs, res, A.norm("совсем другое"))] == [1]
    assert A.lost_segments(segs, res, A.norm("Первичный выпуск ИЭ версии 0"), polish={1: "Первичный выпуск ИЭ версии 0"}) == []


def test_lost_segments_repeated_source_counts_occurrences():
    # один венгерский текст дважды; в res у повтора иная формулировка, а сборка (память) ставит везде первую
    segs = [seg(1, "Zárva vannak az alábbi armatúrák:", 0), seg(2, "Zárva vannak az alábbi armatúrák:", 3)]
    res = {"1": {"ru": "Закрыты следующие арматуры:"}, "2": {"ru": "Следующая арматура закрыта:"}}
    assert A.lost_segments(segs, res, A.norm("Закрыты следующие арматуры: ... Закрыты следующие арматуры:")) == []
    lost = A.lost_segments(segs, res, A.norm("Закрыты следующие арматуры: и только один раз"))
    assert [s["id"] for s in lost] == [2]                           # вхождений 1 на два сегмента — не хватило одному
    assert [s["id"] for s in A.lost_segments(segs, res, A.norm("пусто совсем"))] == [1, 2]


# ---------------------------------------------------------------- 3. шапки
HU_HDR = "Azonosító: 3SZ23 \nVerziószám: 1\nMegnevezés: TURBINA\nOldalszám: 5 \n"
RU_HDR = "Идентификатор: 3SZ23 Номер версии: 1\nНаименование: ТУРБИНА Номер \nстраницы: 5\nтекст"


def test_header_of_both_languages():
    assert A.header_of(HU_HDR, "hu") == ("3SZ23", 1, 5)
    assert A.header_of(RU_HDR, "ru") == ("3SZ23", 1, 5)
    assert A.header_of("просто текст листа", "ru") is None
    assert A.header_of("Azonosító:\n 30RL2126GT\n Verziószám:\n 5\n", "hu") == ("30RL2126GT", 5, None)


def test_compare_headers_version_lost():
    ho = [("3SZ23", 0, 1), ("3SZ23", 1, 2), ("3SZ23/INFO", 0, 1)]
    hr = [("3SZ23", 0, 1), ("3SZ23", 0, 2), ("3SZ23", 0, 3)]
    out = A.compare_headers(ho, hr)
    msgs = [m for lvl, m in out if lvl == A.RED]
    assert any("3SZ23, версия 1" in m for m in msgs) and any("3SZ23/INFO" in m for m in msgs)


def test_compare_headers_same_is_clean_and_order_is_yellow():
    ho = [("A", 0, 1), ("A", 1, 2), ("A", 0, 3)]
    assert A.compare_headers(ho, list(ho)) == []
    out = A.compare_headers(ho, [("A", 0, 1), ("A", 0, 2), ("A", 1, 3)])
    assert [lvl for lvl, _ in out] == [A.YELLOW]
    assert A.compare_headers([], [("A", 1, 1)]) == []              # у оригинала шапки нет — сравнивать не с чем
    assert A.compare_headers(ho, [])[0][0] == A.RED                 # а в переводе шапка пропала


def test_compare_headers_sheet_numbers():
    out = A.compare_headers([("A", 0, 1), ("A", 0, 17)], [("A", 0, 1), ("A", 0, 26)])
    assert [lvl for lvl, _ in out] == [A.YELLOW] and "до 17" in out[0][1] and "до 26" in out[0][1]


# ---------------------------------------------------------------- 5. коды
def test_compare_codes():
    co, _ = A.codes_in(["30RR21S002 30RR21S003 30RR21S003", "31UJ21S601"])
    cr, _ = A.codes_in(["30RR21S002 30RR21S002", "30XX99S999 30RR21S003"])
    gone, extra, more, less = A.compare_codes(co, cr)
    assert gone == {"31UJ21S601": 1} and extra == {"30XX99S999": 1}
    assert more == {"30RR21S002": 1} and less == {"30RR21S003": 1}


def test_code_glued_to_word_counts_like_clean():
    # «30TK62P101szerint» слито в оригинале — в переводе «30TK62P101 согласно» не должно быть «лишним»
    co, _ = A.codes_in(["a 30TK62P101szerint p=3 bar"])
    cr, _ = A.codes_in(["по 30TK62P101 согласно p=3 бар"])
    assert A.compare_codes(co, cr) == ({}, {}, {}, {})
    assert A.codes_in(["30TK62P1011"])[0] == {}                      # лишняя цифра — уже не код


# ---------------------------------------------------------------- 6. отметки
def test_marks_and_glued_rows():
    row = [W(100, 100, 107, 108, "O"), W(111, 100, 118, 108, "O"), W(122, 100, 129, 108, "O")]
    other = [W(100, 200, 107, 208, "O"), W(300, 200, 307, 208, "O")]            # две отметки в разных колонках — не слиплись
    assert len(A.marks_on(row + [W(0, 0, 5, 5, "слово")])) == 3
    assert A.glued_rows(A.marks_on(row)) == 1
    assert A.glued_rows(A.marks_on(other)) == 0
    assert A.glued_rows(A.marks_on([W(100, 10 * i, 107, 10 * i + 8, "\uf0a1") for i in range(5)])) == 0   # столбик по строкам


def test_code_row_share():
    words = [W(10, 100, 70, 108, "30RR21S002"), W(300, 101, 307, 109, "○"),
             W(10, 120, 70, 128, "30RR22S002"), W(300, 140, 307, 148, "○")]      # у второго кода отметка на чужой строке
    assert A.code_row_share(words) == (2, 1)
    assert A.code_row_share([W(10, 100, 70, 108, "30RR21S002")]) == (0, 0)


def test_groups_merge_split_pages():
    gs = A.groups([(0, 0), (1, 1), (1, 2), (2, 2)], range(3), range(3))
    assert len(gs) == 1 and sorted(gs[0][0]) == [1, 2] or len(gs) == 2
    only = A.groups([(0, 0), (1, 1)], range(2), range(2))
    assert sorted(map(lambda g: (g[0], g[1]), only)) == [([0], [0]), ([1], [1])]


# ---------------------------------------------------------------- 7. венгерские буквы
def test_hungarian_leftovers():
    f = A.hu_leftovers("Проверить, что процесс nem sikerült начат; Ф.И.О. Bóth Zoltán, НСБ (BLÜGY), ÜFK, 30RR0102ÁTJGT")
    assert f == ["sikerült"]
    assert "Bóth" not in f and "BLÜGY" not in f and "ÜFK" not in f and "30RR0102ÁTJGT" not in f
    assert A.hu_leftovers("Давление в петле: tömörségvizsgálat прошло") == ["tömörségvizsgálat"]
    assert A.hu_leftovers("Péter") == []                                        # ячейка с одним именем
    assert A.hu_leftovers("tömörségvizsgálat", allowed={"tömörségvizsgálat"}) == []


# ---------------------------------------------------------------- DOCX
def test_docx_text_cells_and_units(tmp_path):
    import docx
    d = docx.Document()
    d.add_paragraph("Абзац вне таблицы")
    t = d.add_table(rows=2, cols=2)
    for r, text in enumerate(["Ячейка один", "Ячейка один", "Ячейка один"][:2]):
        t.cell(r, 0).text = text
    t.cell(0, 1).text = "Верх"
    t.cell(0, 1).add_paragraph("Низ")
    d.sections[0].header.paragraphs[0].text = "Идентификатор: X Номер версии: 0"
    p = tmp_path / "a_RU.docx"
    d.save(p)
    dt = A.DocxText(p)
    assert "абзацвнетаблицы" in dt.full and "идентификатор" in dt.full
    units = dt.units()
    assert "Верх Низ" in units and "Абзац вне таблицы" in units


# ---------------------------------------------------------------- сквозной прогон
def make_pdf(path, pages, images=()):
    """pages — [[(x, y, текст), …], …]; images — {индекс листа: доля площади}: картинка-заливка."""
    import pymupdf
    doc = pymupdf.open()
    for i, items in enumerate(pages):
        pg = doc.new_page(width=595, height=842)
        for x, y, t in items:
            pg.insert_text((x, y), t, fontsize=10, fontname="F0", fontfile=FONT)
        if i in dict(images):
            share = dict(images)[i]
            h = 842 * share
            pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 40), False)
            pix.clear_with(200)
            pg.insert_image(pymupdf.Rect(0, 200, 595, 200 + h), pixmap=pix)
    doc.save(path)
    doc.close()


def make_job(tmp_path, ru_paras, ru_pages, orig_pages, segs, res, ru_images=(), ru_name="doc"):
    import docx
    work = tmp_path / "work"
    for sub in ("json", "out", "in"):
        (work / sub).mkdir(parents=True, exist_ok=True)
    d = docx.Document()
    for t in ru_paras:
        d.add_paragraph(t)
    d.save(work / "out" / f"{ru_name}_RU.docx")
    make_pdf(str(work / "out" / f"{ru_name}_RU.pdf"), ru_pages, ru_images)
    make_pdf(str(work / "in" / f"{ru_name}.pdf"), orig_pages)
    (work / "json" / f"{ru_name}.json").write_text(json.dumps({"file": "x.pdf", "segs": segs, "res": res, "info": {}}), encoding="utf-8")
    return work


def hdr_hu(v, n):
    return [(40, 40, f"Azonosító: 3SZ23 Verziószám: {v}"), (40, 55, f"Oldalszám: {n}")]


def hdr_ru(v, n):
    return [(40, 40, f"Идентификатор: 3SZ23 Номер версии: {v}"), (40, 55, f"Номер страницы: {n}")]


@needs_font
def test_end_to_end_clean(tmp_path):
    segs = [seg(1, "Zárva vannak az armatúrák", 0), seg(2, "Nyitott szelepek vannak", 1)]
    res = {"1": {"ru": "Арматура закрыта"}, "2": {"ru": "Клапаны открыты"}}
    work = make_job(tmp_path, ["Арматура закрыта", "Клапаны открыты"],
                    [hdr_ru(0, 1) + [(40, 100, "Арматура закрыта 30RR21S002")], hdr_ru(1, 2) + [(40, 100, "Клапаны открыты 30RR22S003")]],
                    [hdr_hu(0, 1) + [(40, 100, "Zárva 30RR21S002")], hdr_hu(1, 2) + [(40, 100, "Nyitott 30RR22S003")]], segs, res)
    d = A.Doc(work / "out" / "doc_RU.docx", work, [work / "in"])
    assert d.js and d.orig and d.pdf
    assert [f.as_dict() for f in d.run()] == []
    assert A.main([str(work / "out"), "--work", str(work), "--orig", str(work / "in")]) == 0


@needs_font
def test_end_to_end_defects(tmp_path, capsys):
    segs = [seg(1, "Zárva vannak az armatúrák", 0), seg(2, "Nyitott szelepek vannak", 1), seg(3, "Vezérlő szelep zárva", 1)]
    res = {"1": {"ru": "Арматура закрыта"}, "2": {"ru": "Клапаны открыты"}, "3": {"ru": "Регулирующий клапан закрыт"}}
    boom = "МАШИННЫЙ ЗАЛ ряд 28КА " * 4
    work = make_job(
        tmp_path, ["Арматура закрыта", boom, "Клапаны открыты tömörség"],                       # сегмента 3 в DOCX нет
        [hdr_ru(0, 1) + [(40, 100, "Арматура закрыта 30RR21S002")], hdr_ru(0, 2) + [(40, 100, "Клапаны открыты tömörség")],
         hdr_ru(0, 3) + [(40, 100, "хвост")], [(40, 100, "ещё")]],
        [hdr_hu(0, 1) + [(40, 100, "Zárva 30RR21S002 30RR22S003")], hdr_hu(1, 2) + [(40, 100, "Nyitott")]], segs, res)
    d = A.Doc(work / "out" / "doc_RU.docx", work, [work / "in"])
    found = {f.check: f for f in d.run()}
    assert "3 сегм" not in found["1 потеря содержимого"].msg and "1 сегм" in found["1 потеря содержимого"].msg
    assert "×4" in found["2 повтор фразы"].msg
    assert "версия 1" in found["3 шапки"].msg and found["3 шапки"].level == A.RED
    assert "×2.00" in found["4 листы"].msg                                                       # 4 листа при 2
    assert "30RR22S003" in found["5 коды оборудования"].msg
    assert "tömörség" in found["7 венгерские буквы"].msg
    assert A.main([str(work / "out"), "--work", str(work), "--orig", str(work / "in")]) == 1
    out = capsys.readouterr().out
    assert "🔴" in out and "стр." in out and "doc" in out
    assert A.main([str(work / "out"), "--work", str(work), "--orig", str(work / "in"), "--fail-on", "never"]) == 0


@needs_font
def test_end_to_end_picture_page(tmp_path):
    segs = [seg(1, "Zárva vannak az armatúrák", 0)]
    work = make_job(tmp_path, ["Арматура закрыта"], [[(40, 100, "Арматура закрыта")], [(40, 100, "x")]],
                    [[(40, 100, "Zárva vannak")]], segs, {"1": {"ru": "Арматура закрыта"}}, ru_images=[(1, 0.5)])
    d = A.Doc(work / "out" / "doc_RU.docx", work, [work / "in"])
    f = [x for x in d.run() if x.check == "8 снимок листа"]
    assert f and f[0].page == 2 and f[0].level == A.RED


@needs_font
def test_empty_table_rows_page(tmp_path):
    import pymupdf
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)
    pg.insert_text((60, 70), "Заголовок", fontsize=10, fontname="F0", fontfile=FONT)
    for i in range(16):                                             # сетка 16×3 без текста почти на весь лист
        for j in range(3):
            pg.draw_rect(pymupdf.Rect(40 + j * 170, 100 + i * 44, 40 + (j + 1) * 170, 100 + (i + 1) * 44), width=0.8)
    assert A.empty_row_share(pg, pg.get_text("words")) > A.EMPTY_TABLE_SHARE
    pg2 = doc.new_page(width=595, height=842)
    assert A.empty_row_share(pg2, []) == 0.0


def test_main_without_documents(tmp_path, capsys):
    assert A.main([str(tmp_path)]) == 2
    assert "Нечего проверять" in capsys.readouterr().out


# ---------------------------------------------------------------- без JSON задания (git-worktree без data/jobs)
def _nojson_job(tmp_path):
    """Задание с дефектами (как test_end_to_end_defects); имя редкое — не найдётся в data/jobs основной копии."""
    segs = [seg(1, "Zárva vannak az armatúrák", 0), seg(2, "Nyitott szelepek vannak", 1), seg(3, "Vezérlő szelep zárva", 1)]
    res = {"1": {"ru": "Арматура закрыта"}, "2": {"ru": "Клапаны открыты"}, "3": {"ru": "Регулирующий клапан закрыт"}}
    return make_job(
        tmp_path, ["Арматура закрыта", "Клапаны открыты"],
        [hdr_ru(0, 1) + [(40, 100, "Арматура закрыта 30RR21S002")], hdr_ru(0, 2) + [(40, 100, "Клапаны открыты")]],
        [hdr_hu(0, 1) + [(40, 100, "Zárva 30RR21S002 30RR22S003")], hdr_hu(1, 2) + [(40, 100, "Nyitott")]], segs, res,
        ru_name="zz_audit_nojson_7f3a")


@needs_font
def test_missing_job_json_is_loud_and_listed(tmp_path, capsys):
    """Без JSON задания часть проверок молча отключалась (1 потеря содержимого; без пути к оригиналу — 3, 5, 6, 8):
    тот же набор давал 70 🔴 вместо 112, и сравнение с базой показывало ложное улучшение."""
    work = _nojson_job(tmp_path)
    (work / "json" / "zz_audit_nojson_7f3a.json").unlink()
    rep = tmp_path / "r.json"
    code = A.main([str(work / "out"), "--json", str(rep)])                     # ни --work, ни --orig
    out, err = capsys.readouterr()
    assert code in (0, 1) and "JSON задания не найден" in out and "НЕ сравнимо" in out and "JSON задания" in err
    gaps = json.loads(rep.read_text(encoding="utf-8"))["проверки_пропущены"]["zz_audit_nojson_7f3a"]
    assert {g.split(":")[0] for g in gaps} >= {"1 потеря содержимого", "3 шапки", "5 коды оборудования", "6 отметки",
                                               "8 снимок листа"}
    assert A.main([str(work / "out"), "--require-json"]) == 2
    assert "--require-json" in capsys.readouterr().out
    # оригинал указан: пропущена только проверка по сегментам, сопоставление листов — неполное
    A.main([str(work / "out"), "--orig", str(work / "in"), "--json", str(rep)])
    gaps = json.loads(rep.read_text(encoding="utf-8"))["проверки_пропущены"]["zz_audit_nojson_7f3a"]
    assert gaps[0].startswith("1 потеря содержимого: пропущено") and not any(g.startswith("3 шапки") for g in gaps)
    assert any(g.startswith("6 отметки: неполно") for g in gaps)


@needs_font
def test_with_job_json_output_unchanged_by_new_flag(tmp_path, capsys):
    work = _nojson_job(tmp_path)
    args = [str(work / "out"), "--work", str(work), "--orig", str(work / "in")]
    code = A.main(args + ["--json", str(tmp_path / "a.json")])
    plain = capsys.readouterr()
    assert A.main(args + ["--json", str(tmp_path / "b.json"), "--require-json"]) == code
    strict = capsys.readouterr()
    assert plain == strict and "JSON задания не найден" not in plain.out and plain.err == ""
    a = json.loads((tmp_path / "a.json").read_text(encoding="utf-8"))
    assert list(a) == ["zz_audit_nojson_7f3a"] and (tmp_path / "a.json").read_text() == (tmp_path / "b.json").read_text()


def _scheme(pg, x=150, y=300):
    """Векторная схема: кружок-арматура, труба, косая выноска к рамке надписи (как 3FI18 стр. 7)."""
    import pymupdf
    pg.draw_circle((x + 150, y + 150), 30)
    pg.draw_line((x + 150, y + 120), (x + 150, y))
    pg.draw_line((x + 180, y + 150), (x + 230, y + 250))
    pg.draw_rect(pymupdf.Rect(x + 200, y + 250, x + 300, y + 300))
    pg.draw_line((x, y), (x + 300, y))


def test_figure_rects_scheme_not_table_grid():
    import pymupdf
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)
    for i in range(5):                                              # сетка таблицы — только оси
        pg.draw_line((40, 100 + 20 * i), (500, 100 + 20 * i))
    for x in (40, 200, 500):
        pg.draw_line((x, 100), (x, 180))
    assert A.figure_rects(pg) == []
    _scheme(pg)
    figs = A.figure_rects(pg)
    assert len(figs) == 1 and figs[0].contains(pymupdf.Rect(352, 552, 448, 598)) and figs[0].y0 > 250   # с рамкой надписи


def test_figure_rects_not_text_outlines_table_marks_or_page_frames():
    """Ред-тим 02.11: знаки текста контурами (3PR42 — весь лист кривыми), кружок-отметка у сетки таблицы, рамка листа со
    скруглёнными углами и линии во всю высоту — не рисунок (иначе снимок листа и потеря текста таблицы не замечались)."""
    import pymupdf
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)
    for k in range(20):                                             # «буквы» — кривые внутри рамок строк текста
        pg.draw_circle((100 + 12 * k, 300), 3)
    assert A.figure_rects(pg, [[90, 290, 350, 310]]) == []
    pg = doc.new_page(width=595, height=842)
    for r in range(6):                                              # таблица по ячейкам и кружок-отметка у линии
        for c in range(3):
            pg.draw_rect(pymupdf.Rect(80 + 130 * c, 150 + 30 * r, 210 + 130 * c, 180 + 30 * r))
    pg.draw_circle((210, 165), 7)
    figs = A.figure_rects(pg)
    assert len(figs) == 1 and A.figure_share(figs[0], pg) < 0.01   # рисунок — сам кружок с запасом, не таблица
    pg = doc.new_page(width=595, height=842)
    pg.draw_rect(pymupdf.Rect(20, 20, 575, 822), radius=0.05)        # рамка листа со скруглёнными углами
    pg.draw_line((30, 20), (30, 822))
    pg.draw_circle((300, 400), 5)
    assert all(A.figure_share(f, pg) < 0.01 for f in A.figure_rects(pg))


@needs_font
@pytest.mark.parametrize("share, snapshot", [(0.3, False), (0.95, True)])
def test_vector_scheme_as_picture(tmp_path, share, snapshot):
    """Схема, нарисованная векторами, в переводе — картинкой с надписями (3FI18 стр. 7, 26–27): надписи в ней — не
    потеря, картинка — не снимок листа. Картинка во весь лист при небольшой схеме — по-прежнему снимок."""
    import pymupdf
    segs = [seg(1, "Szórtvizes armatúra működtetése", 0, bbox=[60, 90, 300, 104]),
            seg(2, "Áramlásjelző nincs használva", 0, bbox=[355, 560, 445, 590])]
    res = {"1": {"ru": "Управление арматурой распыления"}, "2": {"ru": "Датчик расхода не используется"}}
    work = make_job(tmp_path, ["Управление арматурой распыления"], [[(40, 100, "Управление арматурой распыления")]],
                    [[(60, 100, "Szórtvizes armatúra működtetése")]], segs, res, ru_images=[(0, share)])
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)
    pg.insert_text((60, 100), "Szórtvizes armatúra működtetése", fontsize=10, fontname="F0", fontfile=FONT)
    pg.insert_text((358, 580), "Áramlásjelző nincs használva", fontsize=7, fontname="F0", fontfile=FONT)
    _scheme(pg)
    doc.save(str(work / "in" / "doc.pdf"))
    doc.close()
    d = A.Doc(work / "out" / "doc_RU.docx", work, [work / "in"])
    found = {f.check for f in d.run()}
    assert "1 потеря содержимого" not in found
    assert ("8 снимок листа" in found) == snapshot


def test_page_background_fill_is_not_a_picture_page(tmp_path):
    """Фон листа — картинка 1×1 цвета во весь лист за текстом (листы ÜTJ, docx_build._page_bg): заливка, не снимок;
    настоящая картинка во весь лист — снимок."""
    import io
    import pymupdf
    from PIL import Image

    def png(w, h):
        buf = io.BytesIO()
        Image.new("RGB", (w, h), "#FFFF99").save(buf, "PNG")
        return buf.getvalue()
    doc = pymupdf.open()
    for w, h in ((1, 1), (60, 80)):
        pg = doc.new_page(width=595, height=842)
        pg.insert_image(pg.rect, stream=png(w, h), overlay=False)
        pg.insert_text((60, 70), "Text", fontsize=10)
    assert A.image_share(doc[0]) == 0.0
    assert A.image_share(doc[1]) > 0.9                              # картинка 60×80 — во весь лист с полями пропорций
