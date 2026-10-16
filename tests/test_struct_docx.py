# -*- coding: utf-8 -*-
"""Структурный DOCX по профилю оригинала и шлифовка по языковым соглашениям (UTS, 15.10, утверждено Максимом):
списки по месту и номерам из текста (номер не дублируется, нумерация продолжается после рисунка), память без
номеров пунктов, стили и нумерация заголовков по профилю, значок в скобках, соглашения комплекта (ссылки,
латиница, коды, предупреждение, термины документа), цикл шлифовки «код → портал → проверка»; в конвейере —
размеченный PDF узнаётся по покрытию разметкой, абзацы разметки без пары переводятся своими сегментами, шлифовка
правит переводы задачи, сборка даёт DOCX по разметке и таблицу вычитки с листами «шлифовка» и «нарушения»."""
import json
import re

import fitz
import pytest
from docx import Document

import conventions as C
import docx_build as B
import docx_struct as D
import polish as PL

W = B.W


def p(text, page=1, top=100.0, x0=80.0, **kw):
    g = {"page": page, "top": top, "bottom": top + 10, "x0": x0, "x1": 500.0, "first_x0": x0, "size": 10, "lines": 1}
    return dict({"k": "p", "runs": [{"t": text, "font": "Tahoma", "size": 10, "b": False, "i": False, "color": "000000"}],
                 "geom": g}, **kw)


def lst(kind, items):
    return {"k": "list", "kind": kind, "items": items}


RULES = {
    "latin": ["UTS", "Ethernet"], "cyrillic_to_latin": {"УТС": "UTS", "Этернет": "Ethernet"},
    "abbr": {"TU": "ИИ"}, "warning": {"hu": "FIGYELEM", "ru": "ВНИМАНИЕ!"},
    "terms": [{"hu": "[Ii]ndít", "ru": "запуск", "swap": "за",
               "bad": "(?<![А-Яа-яЁё])[Пп](?:уск\\w*|уст(?:ить|ит|ил|ила)\\b|ущен\\w*)"},
              {"hu": "[Tt]esztlépés", "ru": "шаг испытания", "bad": "[Тт]естов\\w*\\s+шаг"}],
}


# ---------------------------------------------------------------- чтение
def test_relist_follows_page_order_and_labels():
    """Дерево Word на стыке страниц: подпункты п. 3 лежат в подсписке п. 2, а п. 3–4 — отдельным списком следом."""
    sub2 = lst("LowerAlpha", [[p("a. első", 44, 600, 116)], [p("b. második", 44, 620, 116)],
                              [p("a. harmadik alatt", 45, 180, 116)]])
    l1 = lst("Decimal", [[p("1. egy", 44, 500)], [p("2. kettő", 44, 580), sub2]])
    l2 = lst("Decimal", [[p("3. három", 45, 150)], [p("4. négy", 45, 260)]])
    out = D.relist([l1, l2])
    assert len(out) == 1 and out[0]["start"] == 1
    items = out[0]["items"]
    assert [D.text_of(it[0]["runs"]) for it in items] == ["egy", "kettő", "három", "négy"]   # номер — не текстом
    assert [D.text_of(x[0]["runs"]) for x in items[1][1]["items"]] == ["első", "második"]
    assert [D.text_of(x[0]["runs"]) for x in items[2][1]["items"]] == ["harmadik alatt"]
    assert items[2][1]["start"] == 1


def test_numbering_continues_after_figure():
    fig = {"k": "p", "runs": [{"img": {"top": 0, "bottom": 100}}], "geom": {"page": 9, "top": 300}}
    out = D.relist([lst("Decimal", [[p("6. hat", 9, 100)]]), fig, lst("Decimal", [[p("7. hét", 9, 500)]])])
    assert out[2]["start"] == 7 and D.text_of(out[2]["items"][0][0]["runs"]) == "hét"


def test_memory_finds_item_without_label():
    mem = B.Memory([("a. BER panel állapota", "a. состояние панели BER"),
                    ("<f1>b. az UTS kiadott </f1><f2>ping x_can</f2>", "<f1>b. команда UTS </f1><f2>ping x_can</f2>")])
    assert mem("BER panel állapota") == "состояние панели BER"
    assert mem("<f1>az UTS kiadott </f1><f2>ping x_can</f2>") == "<f1>команда UTS </f1><f2>ping x_can</f2>"


def test_math_letters_become_plain_text():
    runs = D._runs([{"ch": {"text": c, "fontname": "CambriaMath", "size": 10, "top": 0, "x0": i * 5, "x1": i * 5 + 4,
                            "non_stroking_color": [0]}} for i, c in enumerate("𝑏𝑖𝑛á𝑟𝑖𝑠")])
    assert D.text_of(runs) == "bináris"


# ---------------------------------------------------------------- сборка по профилю
PROFILE = {"page": {"right": 62.3}, "roles": {"body": {"font": "Tahoma", "size": 10.0, "b": False},
                                              "H1": {"font": "Tahoma", "size": 14.0, "b": True}},
           "align": {"body": "both", "list": "left"}, "line": 13.9, "space": {"body": 6.1},
           "heads": {"1": {"page_break": True, "tab": 21.6}, "2": {"page_break": False, "tab": 28.8}},
           "lists": {"0": {"marker": 18.0, "text": 36.0}, "1": {"marker": 54.0, "text": 72.0}},
           "figure": {"center": True}, "caption": {"center": True}, "table": {"pad_left": 5.4}}
LAYOUT = {"W": 595.3, "H": 841.9, "left": 62.4, "right": 59.7, "top": 119, "bottom": 123, "header": 36, "footer": 18,
          "title_left": 45, "head_imgs": [], "head_lines": [], "stamp": [], "stamp_widths": []}


def png():
    return fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 4, 4), False).tobytes("png")


def test_builder_follows_profile(tmp_path):
    icon = {"k": "p", "runs": [{"img": {"x0": 300, "x1": 310, "top": 100, "bottom": 110}}],
            "geom": {"page": 2, "top": 100, "x0": 300, "x1": 310, "first_x0": 300}}
    blocks = [{"k": "h", "level": 1, "runs": [{"t": "1 Bevezetés", "font": "Tahoma", "size": 14, "b": True}],
               "geom": {"page": 2, "top": 90}},
              p("Az indító () gombbal.", 2, 100, 62.4), icon,
              dict(lst("Decimal", [[p("hét", 2, 200)]]), start=7)]
    mem = B.Memory([("Bevezetés", "Введение"), ("Az indító () gombbal.", "Кнопкой запуска ( )."), ("hét", "семь")])
    out = tmp_path / "x.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), LAYOUT)
    doc = Document(str(out))
    st = doc.styles
    assert st["Heading 1"].paragraph_format.page_break_before and int(st["Normal"].paragraph_format.alignment) == 3
    assert abs(st["Normal"].paragraph_format.line_spacing.pt - 13.9) < 0.1
    xml = doc.part.numbering_part.element.xml
    assert 'w:pos="432"' in xml and 'w:hanging="432"' in xml                 # 21,6 pt номер → текст у H1
    assert 'w:startOverride w:val="7"' in xml
    par = next(x for x in doc.paragraphs if "запуска" in x.text)
    seq = "".join("\0" if e.tag == W + "drawing" else (e.text or "") for e in par._p.iter(W + "t", W + "drawing"))
    assert "(\0)" in seq                                                     # значок — в скобках, не в конце


def test_monospace_part_command_verbatim_placeholder_in_brackets():
    assert B.code_part("sync", "синхронизировать") == "sync"
    assert B.code_part("gatewayconfig ipaddr ", "настройка IP-адреса шлюза") == "gatewayconfig ipaddr "
    assert B.code_part(" <beállítani kívánt IP cím>", "IP-адрес для установки") == " <IP-адрес для установки>"
    assert B.keep_code("<f1>Adjuk ki: </f1><f2>sync</f2>", "<f1>Введите: </f1><f2>синхронизация</f2>", [2]) \
        == "<f1>Введите: </f1><f2>sync</f2>"


def test_title_bottom_text_goes_to_title_footer(tmp_path):
    """Юридический текст внизу титула — в подвал титула: иначе русский адрес длиннее и текст уезжал на стр. 2."""
    blocks = [p("UTS felhasználói kézikönyv", 1, 300, 150), p("A dokumentum átadója", 1, 780, 20),
              {"k": "h", "level": 1, "runs": [{"t": "1 Bevezetés", "font": "Tahoma", "size": 14, "b": True}],
               "geom": {"page": 2, "top": 90}}]
    mem = B.Memory([("UTS felhasználói kézikönyv", "Руководство пользователя UTS"),
                    ("A dokumentum átadója", "Передающей стороной"), ("Bevezetés", "Введение")])
    out = tmp_path / "t.docx"
    B.Builder(mem, lambda img: None, PROFILE).build(blocks, str(out), LAYOUT)
    doc = Document(str(out))
    assert [x.text for x in doc.sections[0].footer.paragraphs] == ["Передающей стороной"]
    assert "Передающей стороной" not in [x.text for x in doc.paragraphs]


# ---------------------------------------------------------------- соглашения
def test_refs_quotes_names_codes_warning_terms():
    fix = lambda hu, ru: C.fix_code(hu, ru, RULES)[0]
    assert fix("lásd: 5.5", "(см.: 5.5)") == "(см. раздел 5.5)"
    assert fix("lásd: 7.2 és 7.3", "см. 7.2 и 7.3") == "см. разделы 7.2 и 7.3"
    assert fix("(4. fejezet)", "Запуск UTS (глава 4)") == "Запуск UTS (см. раздел 4)"
    assert fix("a 8. fejezetben", "описано в главе 8") == "описано в разделе 8"
    assert fix("az 5. ábrán", "показано на рис. 5") == "показано на рисунке 5"
    assert fix("Ethernet kábel", "Кабель Этернет") == "Кабель Ethernet"
    assert fix("UTS_A.LIS", "UTS_A.СПИС") == "UTS_A.LIS"
    assert fix("FIGYELEM: újra", "ВНИМАНИЕ: при повторе") == "ВНИМАНИЕ! При повторе"
    assert fix("<f1>FIGYELEM!</f1><f2> újra</f2>", "<f1>Внимание!</f1><f2> при повторе</f2>") \
        == "<f1>ВНИМАНИЕ!</f1><f2> При повторе</f2>"                                # без двух пробелов в DOCX
    assert fix("a TU-k listája", "список TU") == "список ИИ"
    assert fix("az UTS indítása", "Пуск UTS") == "Запуск UTS"
    assert fix("indítás után üres", "после пуска поле пустое") == "после запуска поле пустое"   # «пустое» не трогать
    assert fix("„B” szerver", 'сервер "B"') == "сервер «B»"
    assert fix('Filter=„0x0”', 'Filter="0x0"') == 'Filter="0x0"'                  # значение настройки — как есть
    assert fix("<dátum>", "TxsAttachLog_&lt;дата&gt;") == "TxsAttachLog_<дата>"
    assert C.check("5.sz. melléklet", "Приложение 5", RULES) == []                   # «5.sz» — не код
    v = C.check("Az UTS tesztlépés", "тестовый шаг UTS", RULES)
    assert [(r, code) for r, _, code in v] == [("Я10", False)]
    assert C.check("PARAMÉTERNÉV=PARAMÉTERÉRTÉK", "пара PARAMÉTERNÉV=PARAMÉTERÉRTÉK", RULES)[0][0] == "Я11"
    assert C.check("BLÜGY laptop", "ноутбук НСБ (BLÜGY)", RULES) == []           # первое употребление сокращения


def test_polish_cycle_code_then_portal_and_log(tmp_path):
    items = {"a": {"hu": "Az UTS indítása (lásd: 4.2)", "ru": "Пуск УТС (см.: 4.2)", "kind": "text", "page": 3, "id": 1},
             "b": {"hu": "A tesztlépés vége", "ru": "Конец тестового шага", "kind": "text", "page": 4, "id": 2},
             "c": {"hu": "A tesztlépés eleje", "ru": "Начало тестового шага", "kind": "text", "page": 5, "id": 3}}
    calls = []

    def chat(msgs, **kw):
        data = json.loads(msgs[-1]["content"])
        calls.append(len(data))
        # «b» портал исправил, «c» вернул с лишним числом — правка не принимается
        fixed = {"Конец тестового шага": "Конец шага испытания", "Начало тестового шага": "Начало 2 шага испытания"}
        return {"text": json.dumps([{"id": d["id"], "ru": fixed[d["ru"]]} for d in data], ensure_ascii=False)}
    edits, left = PL.run(items, RULES, chat=chat, log=lambda *a: None)
    assert items["a"]["ru"] == "Запуск UTS (см. раздел 4.2)"
    assert items["b"]["ru"] == "Конец шага испытания" and items["c"]["ru"] == "Начало тестового шага"
    assert {(e["key"], e["who"]) for e in edits} == {("a", "код"), ("b", "портал")}
    assert [x["key"] for x in left] == ["c"] and calls[0] == 2
    out = tmp_path / "v.xlsx"
    PL.write_xlsx(str(out), edits, left)
    from openpyxl import load_workbook
    wb = load_workbook(out)
    assert wb["шлифовка"].max_row == 3 and wb["нарушения"]["B2"].value == "Я10"


def test_polish_accepts_collapsing_expanded_name_rejects_cut_answer():
    """Портал сворачивает раскрытое название в «UTS» — перевод короче в разы, это не обрыв (сервис 15.10: все 12 таких
    правок отбрасывала проверка длины против старого перевода); обрыв ответа по-прежнему не принимается."""
    hu, old = "Kapcsolja be az UTS számítógépet!", "Включите ЭВМ системы оперативной информации!"
    assert PL._ok(old, "Включите ЭВМ UTS!", hu, RULES, "text")
    assert PL._ok("3 Система оперативной информации", "3 Система UTS", "3 Az UTS rendszer", RULES, "text")
    assert not PL._ok(old, "UTS", hu, RULES, "text")


def test_vocab_skips_paths_and_abbreviation_written_once_lowercase():
    """«uts» в путях и именах файлов и один раз строчными в тексте при сотнях «UTS» — не венгерское слово (иначе голое
    «UTS» переводилось и не ловилась потерянная латиница, сервис 15.10); слова через косую и венгерское слово
    прописными, которое встречается и строчными, остаются."""
    import pipeline as P

    class T:
        entries = []
    segs = [{"text": t} for t in ["Az UTS alkalmazás indítása"] * 12 + ["MINDEN JEL"] * 20 + [
        "(/etc/uts/shellenv)", "az uts.ini fájl", "uts_<éééé>.tar.gz", "egyébként pedig az uts naplófájl",
        "a fogyasztás kg/óra", "minden", "minden", "minden"]]
    tr = T()
    P.set_vocab(tr, segs)
    assert "uts" not in tr.vocab and "ini" not in tr.vocab
    assert {"kg", "óra", "minden", "alkalmazás", "naplófájl"} <= tr.vocab


def test_item_number_as_text_but_not_initials():
    import docx_profile as DP
    assert DP.ITEM_NUM_TEXT.match("1. Включите ЭВМ") and DP.ITEM_NUM_TEXT.match("a) пункт")
    assert not DP.ITEM_NUM_TEXT.match("Ф. И. О. оператора: имя пользователя UTS")


def run(t, x, top=140.0, page=31, **kw):
    return dict({"t": t, "font": "Tahoma", "size": 10, "b": False, "i": False, "color": "000000",
                 "x": x, "top": top, "page": page}, **kw)


def test_icon_in_line_goes_to_its_place_not_paragraph_end(tmp_path):
    """Word рисует значок строки после текста: «● : не подключено» становилось «: не подключено●» (UTS, 27 пунктов,
    16.10). Значок встаёт перед текстом своей строки по листу; значок без MCID (оформление) возвращается в абзац своей
    строки; в DOCX значок после маркера пункта — в начале пункта."""
    icon = {"img": {"x0": 98, "x1": 106, "top": 141, "bottom": 149, "page_number": 31}}
    runs = [run("•", 90), run("\t", 94), run(": az UTS nincs felcsatlakozva", 108)]
    got = D.insert_imgs(runs, [icon])
    assert ["img" in r for r in got] == [False, False, True, False]
    far = {"img": {"x0": 400, "x1": 408, "top": 300, "bottom": 308, "page_number": 31}}
    assert "img" in D.insert_imgs(runs, [far])[-1]                       # не на строке абзаца — в конец, как было
    para = {"k": "p", "runs": [run(": nem meghatározható", 108, 300)],
            "geom": {"page": 31, "top": 300, "x0": 108, "x1": 200, "size": 10}}
    D.attach_icons([para], [{"x0": 98, "x1": 106, "top": 301, "bottom": 309, "page_number": 31},
                            {"x0": 62, "x1": 120, "top": 40, "bottom": 80, "page_number": 31}], 842)   # логотип шапки — нет
    assert ["img" in r for r in para["runs"]] == [True, False]
    blocks = [lst("Disc", [[{"k": "p", "runs": got, "geom": {"page": 31, "top": 140, "x0": 90, "x1": 300}}]])]
    mem = B.Memory([(": az UTS nincs felcsatlakozva", ": UTS не подключено")])
    out = tmp_path / "i.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), LAYOUT)
    par = next(x for x in Document(str(out)).paragraphs if "подключено" in x.text)
    seq = "".join("\0" if e.tag == W + "drawing" else (e.text or "") for e in par._p.iter(W + "t", W + "drawing"))
    assert seq.lstrip().startswith("\0") and seq.rstrip().endswith("подключено")


def test_table_continued_with_repeated_header_is_joined():
    """Продолжение таблицы на новом листе, начатое повтором шапки («Повторять как заголовок»), — та же таблица, повтор
    прочь; другая таблица с другой шапкой — своя (синтетика, 16.10)."""
    cell = lambda th, t: {"blocks": [p(t)], "th": th}
    head = lambda: [cell(True, "№"), cell(True, "Параметр")]
    t1 = {"k": "table", "rows": [head(), [cell(False, "1"), cell(False, "a")]], "geom": {"page": 4}}
    t2 = {"k": "table", "rows": [head(), [cell(False, "2"), cell(False, "b")]], "geom": {"page": 5}}
    t3 = {"k": "table", "rows": [[cell(True, "Код"), cell(True, "Имя")], [cell(False, "x"), cell(False, "y")]],
          "geom": {"page": 6}}
    got = D.join_tables([t1, t2, t3])
    assert [len(t["rows"]) for t in got] == [3, 2]


def test_table_cut_by_pages_is_joined():
    """Word при экспорте режет таблицу по листам: продолжение — таблица без шапки сразу за ней (UTS стр. 75–78)."""
    cell = lambda th: {"blocks": [], "th": th}
    t1 = {"k": "table", "rows": [[cell(True), cell(True)], [cell(False), cell(False)]], "geom": {"page": 75}}
    t2 = {"k": "table", "rows": [[cell(False), cell(False)]] * 3, "geom": {"page": 76}}
    t3 = {"k": "table", "rows": [[cell(True), cell(True)], [cell(False), cell(False)]], "geom": {"page": 77}}
    got = D.join_tables([t1, t2, t3])
    assert [len(t["rows"]) for t in got] == [5, 2]


def test_label_tab_kept_and_column_alignment_from_original():
    assert B._keep_tab("Műszaki változás száma\t10776", "Номер технического изменения 10776") == \
        "Номер технического изменения\t10776"
    assert B._keep_tab("Tervező\tSinger Gábor", "Проектант\tСинглер Габор") == "Проектант\tСинглер Габор"
    g = lambda x0, x1: {"page": 2, "x0": x0, "x1": x1, "lines": 1}
    rows = [[{"th": True, "geom": g(160, 202)}, {"th": True, "geom": g(410, 450)}],
            [{"th": False, "geom": g(66, 120)}, {"th": False, "geom": g(500, 555)}]]
    assert B._col_align(rows, [62, 300, 560], 2) == {("head", 0): "center", ("head", 1): "center", ("body", 1): "right"}


def test_latin_name_counted_paths_aside_and_verbatim_style_kept():
    """«…UTS… при запуске СОИ другой среды» проходило проверку: UTS хоть раз есть (или «uts_kmod.conf»)."""
    hu = "Blokki UTS esetén az UTS alkalmazás indítása előtt, lásd /etc/modprobe.d/uts_kmod.conf"
    assert any(v[0] == "Я5" for v in C.check(hu, "В случае СОИ блока перед запуском приложения UTS, см. "
                                                 "/etc/modprobe.d/uts_kmod.conf", RULES))
    assert not any(v[0] == "Я5" for v in C.check(hu, "В случае UTS блока перед запуском приложения UTS, см. "
                                                     "/etc/modprobe.d/uts_kmod.conf", RULES))
    assert not any(v[0] == "Я5" for v in C.check("<UTS környezet>/<év>", "<среда UTS>/<год>", RULES))
    plain, ital = (False, False, False, None), (False, True, False, None)
    parts = ["Ezen kívül a kernel modul ", "/etc/modprobe.d/uts_kmod.conf", " fájlban meg kell adni"]
    groups = [[plain, []], [ital, []], [plain, []]]
    out = B._restyle("Кроме того, в файле /etc/modprobe.d/uts_kmod.conf необходимо указать", groups, parts)
    assert (ital, "/etc/modprobe.d/uts_kmod.conf") in out and "".join(t for _, t in out).startswith("Кроме")


def test_slanted_tahoma_is_italic_and_verbatim_parts_need_no_portal():
    """У Tahoma нет курсива — Word наклоняет знаки матрицей (c = 0,33): это курсив (пути в UTS). Если перевод без меток
    есть и курсивный кусок стоит в нём дословно — перевод с метками у портала не просим."""
    ch = {"fontname": "BCDEEE+Tahoma", "size": 10, "non_stroking_color": (0,), "matrix": (1, 0, 0.33, 1, 200, 600)}
    assert D._fmt(ch)["i"] and not D._fmt(dict(ch, matrix=(1, 0, 0, 1, 200, 600)))["i"]
    plain, ital = (False, False, False, None), (False, True, False, None)
    groups = [[plain, []], [ital, []], [plain, []]]
    parts = ["Ezen kívül a kernel modul ", "/etc/modprobe.d/uts_kmod.conf", " fájlban"]
    assert B._verbatim("Кроме того, в файле /etc/modprobe.d/uts_kmod.conf", groups, parts)
    assert not B._verbatim("Отметьте опцию «Шаг испытания успешен»", groups, ["Jelölje be a ", "„A tesztlépés sikeres”", ""])


def test_toc_like_original_levels_bold_title_and_unnumbered_entry(tmp_path):
    """Оглавление — по оригиналу: уровней столько же (UTS: 3, а собиралось 4), 1-й полужирный, отступы и шаг строк,
    заголовок своим кеглем и цветом, пункт без номера раздела (шапка таблицы) — с уровнем структуры."""
    info = {"title": {"t": "Tartalom", "font": "Tahoma", "size": 14, "b": True, "color": "365f91", "center": False},
            "levels": 3, "bold": [True, False, False], "indent": [0, 21, 40], "pitch": 19.6, "figures": False,
            "extra": [("Dokumentumhoz kapcsolódó szerepek", 1)]}
    th = {"blocks": [p("Dokumentumhoz kapcsolódó szerepek", 2, 100)], "th": True}
    blocks = [{"k": "table", "rows": [[th], [{"blocks": [p("Singler Gábor", 2, 120)], "th": False}]],
               "geom": {"page": 2, "top": 100}}, {"k": "toc", "info": info}]
    mem = B.Memory([("Tartalom", "Содержание"), ("Dokumentumhoz kapcsolódó szerepek", "Роли, связанные с документом"),
                    ("Singler Gábor", "Синглер Габор")])
    out = tmp_path / "t.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), LAYOUT)
    doc = Document(str(out))
    body = doc.element.body.xml
    assert 'TOC \\o "1-3"' in body and '<w:outlineLvl w:val="0"/>' in body
    st = {s.style_id: s for s in doc.styles}
    assert st["TOC1"].font.bold and not st["TOC2"].font.bold and st["TOC2"].paragraph_format.left_indent.pt == 21
    assert st["TOC1"].element.get(W + "customStyle") is None      # встроенный «toc 1», иначе Word берёт свой
    assert str(st["TOCTitleUTS"].font.color.rgb) == "365F91"


# ---------------------------------------------------------------- в конвейере: извлечение → структура → шлифовка → сборка
def tagged_pdf(path, pages, tagged=True, producer="Microsoft® Word 2019"):
    """PDF «как из Word»: pages — [[(тег, текст, x, y, кегль, жирный)]], каждая строка — свой элемент дерева разметки
    со своим MCID. Возвращает рамки строк [(стр., текст, Rect, кегль, жирный)] — из них сегменты страниц."""
    doc = fitz.open()
    elems, lines = [], []
    for pno, items in enumerate(pages):
        page = doc.new_page(width=595, height=842)
        for tag, text, x, y, size, bold in items:
            page.insert_text((x, y), text, fontsize=size, fontname="cour" if bold == "code" else "hebo" if bold else "helv")
            lines.append((pno, text, page.search_for(text)[0], size, bold))
        if not tagged:
            continue
        page.clean_contents()
        xref = page.get_contents()[0]
        s = doc.xref_stream(xref).decode("latin-1")
        out, pos = [], 0
        for i, (b, it) in enumerate(zip(re.findall(r"BT.*?ET", s, re.S), items)):
            k = s.index(b, pos)
            out.append(s[pos:k] + f"/{it[0]} <</MCID {i}>> BDC " + b + " EMC")
            pos = k + len(b)
        doc.update_stream(xref, ("".join(out) + s[pos:]).encode("latin-1"))
        elems += [(it[0], page.xref, i) for i, it in enumerate(items)]
    if tagged:
        root, sect = doc.get_new_xref(), doc.get_new_xref()
        kids = []
        for tag, pxref, mcid in elems:
            x = doc.get_new_xref()
            doc.update_object(x, f"<< /Type /StructElem /S /{tag} /P {sect} 0 R /Pg {pxref} 0 R /K {mcid} >>")
            kids.append(x)
        doc.update_object(sect, f"<< /Type /StructElem /S /Sect /P {root} 0 R /K [{' '.join(f'{k} 0 R' for k in kids)}] >>")
        doc.update_object(root, f"<< /Type /StructTreeRoot /K [{sect} 0 R] >>")
        doc.xref_set_key(doc.pdf_catalog(), "StructTreeRoot", f"{root} 0 R")
    doc.set_metadata({"producer": producer, "creator": producer})
    doc.save(path)
    return lines


def page_segs(lines):
    return [{"id": i, "text": t, "src": t, "page": pno,
             "loc": {"bbox": list(r), "avail": list(r), "frags": [list(r)], "align": "left", "cell": [0, 595],
                     "size": size, "bold": bold, "family": "sans-serif", "ocr": False}}
            for i, (pno, t, r, size, bold) in enumerate(lines)]


PAGES = [[("P", "UTS felhasználói kézikönyv", 150, 300, 18, True)],
         [("H1", "1 Bevezetés", 62, 130, 14, True),
          ("P", "Az UTS indítása után a felület üres.", 62, 160, 10, False),
          ("P", "FIGYELEM!", 62, 190, 10, True), ("P", "A teszt leállítása nem lehetséges.", 115, 190, 10, False),
          ("P", "Adjuk ki az alábbi parancsot:", 62, 220, 10, False), ("P", "sync", 210, 220, 10, "code")]]
# перевод «портала»: с нарушениями соглашений (УТС, «пуск», «ВНИМАНИЕ:»), их правит шлифовка кодом; команда
# моноширинным шрифтом переведена — её возвращает сборка
RU = {"UTS felhasználói kézikönyv": "Руководство пользователя UTS", "1 Bevezetés": "1 Введение",
      "Az UTS indítása után a felület üres.": "После пуска УТС интерфейс пуст.", "FIGYELEM!": "ВНИМАНИЕ!",
      "A teszt leállítása nem lehetséges.": "Остановить испытание нельзя.",
      "<f1>FIGYELEM!</f1><f2> A teszt leállítása nem lehetséges.</f2>":
          "<f1>ВНИМАНИЕ:</f1><f2> остановить испытание нельзя.</f2>",
      "Adjuk ki az alábbi parancsot:": "Введите следующую команду:", "sync": "sync",
      "<f1>Adjuk ki az alábbi parancsot:</f1><f2> sync</f2>": "<f1>Введите следующую команду:</f1><f2> синхронизировать</f2>"}


def test_tagged_pdf_is_recognized_by_coverage(tmp_path):
    lines = tagged_pdf(tmp_path / "t.pdf", PAGES)
    assert B.coverage(str(tmp_path / "t.pdf"), page_segs(lines)) == 1.0
    plain = tagged_pdf(tmp_path / "p.pdf", PAGES, tagged=False)
    assert B.coverage(str(tmp_path / "p.pdf"), page_segs(plain)) == 0.0      # без разметки — прежним путём
    # разметка не из Word (Acrobat ставит её сам, дерево иное: 31RDGT развалился) — прежним путём
    acro = tagged_pdf(tmp_path / "a.pdf", PAGES, producer="Adobe PDF Library 20.1.68")
    assert B.coverage(str(tmp_path / "a.pdf"), page_segs(acro)) == 0.0


def test_pipeline_struct_fill_polish_assemble(tmp_path, monkeypatch):
    import pipeline as P
    import translate as T
    from openpyxl import load_workbook
    pdf = tmp_path / "doc.pdf"
    segs = page_segs(tagged_pdf(pdf, PAGES))
    work = tmp_path / "work"
    (work / "json").mkdir(parents=True)
    res = {str(s["id"]): {"ru": RU[s["text"]], "issues": []} for s in segs}
    (work / "json" / "doc.json").write_text(json.dumps({"file": str(pdf), "segs": segs, "res": res, "info": {},
                                                        "docx": None, "struct": {"coverage": 1.0}}), encoding="utf-8")
    g = tmp_path / "g.md"
    g.write_text("## Т\n\n| Венгерский | Русский | Комментарий |\n|---|---|---|\n| szelep | клапан | |\n", encoding="utf-8")
    rules = tmp_path / "conventions.json"
    rules.write_text(json.dumps({k: RULES[k] for k in ("latin", "cyrillic_to_latin", "warning")}), encoding="utf-8")
    (tmp_path / "doc_terms").mkdir()
    (tmp_path / "doc_terms" / "doc_terms.json").write_text(json.dumps({"terms": RULES["terms"][:1]}), encoding="utf-8")
    prj = {"glossary": str(g), "rules": str(rules), "doc_terms": str(tmp_path / "doc_terms")}
    asked = []

    def call(self, units, *a, **k):
        asked.extend(u["text"] for u in units)
        return {i: RU[u["text"]] for i, u in enumerate(units)}, {}
    monkeypatch.setattr(T.Translator, "_call", call)

    # абзац «FIGYELEM!» жирным + текст на той же строке — в разметке два элемента, в DOCX — один абзац с метками:
    # среди сегментов страниц его нет, он переводится своим сегментом (без меток — не спрашивается)
    P.struct_fill(work, prj)
    d = json.loads((work / "json" / "doc.json").read_text(encoding="utf-8"))
    own = [s for s in d["segs"] if s["loc"].get("struct_only")]
    assert asked == ["<f1>FIGYELEM!</f1><f2> A teszt leállítása nem lehetséges.</f2>",
                     "<f1>Adjuk ki az alábbi parancsot:</f1><f2> sync</f2>"]
    assert [s["page"] for s in own] == [1, 1] and d["struct"]["filled"] == 2
    assert d["res"][str(own[1]["id"])]["ru"] == "<f1>Введите следующую команду:</f1><f2> sync</f2>"   # команда — дословно

    P.polish(work, prj, chat=lambda *a, **k: pytest.fail("всё правится кодом — портал не нужен"))
    d = json.loads((work / "json" / "doc.json").read_text(encoding="utf-8"))
    r = d["res"][str(segs[2]["id"])]
    assert r["ru"] == "После запуска UTS интерфейс пуст." and r["polish"]["who"] == "код"
    assert d["res"][str(own[0]["id"])]["ru"] == "<f1>ВНИМАНИЕ!</f1><f2> Остановить испытание нельзя.</f2>"
    assert d["polish"]["remaining"] == [] and {e["rule"] for e in d["polish"]["edits"]} >= {"Я5, Я10", "Я8"}

    P.assemble(work)
    out = work / "out"
    doc = Document(str(out / "doc_RU.docx"))
    pars = [(p.style.name, p.text) for p in doc.paragraphs if p.text.strip()]
    assert ("Heading 1", "Введение") in pars and ("Normal", "После запуска UTS интерфейс пуст.") in pars
    warn = next(p for p in doc.paragraphs if p.text.startswith("ВНИМАНИЕ!"))
    assert warn.text == "ВНИМАНИЕ! Остановить испытание нельзя." and warn.runs[0].bold
    cmd = next(p for p in doc.paragraphs if p.text.startswith("Введите"))
    assert cmd.text == "Введите следующую команду: sync" and cmd.runs[-1].font.name == "Consolas"
    assert (out / "doc_RU.pdf").exists()
    wb = load_workbook(out / "doc_вычитка.xlsx")
    assert wb.sheetnames == ["вычитка", "шлифовка", "нарушения"]
    main = {row[2]: row for row in wb["вычитка"].iter_rows(min_row=2, values_only=True)}
    assert main["Az UTS indítása után a felület üres."][3] == "После запуска UTS интерфейс пуст."
    assert "Шлифовка (Я5, Я10; код), было: После пуска УТС" in main["Az UTS indítása után a felület üres."][7]
    assert "FIGYELEM! A teszt leállítása nem lehetséges." in main                # абзац разметки — строкой вычитки
    assert wb["шлифовка"].max_row == 1 + len(d["polish"]["edits"])


def test_editor_sheet_found_by_name_when_other_sheet_active(tmp_path):
    """Редактор сохранил таблицу, стоя на листе «шлифовка»: правки всё равно читаются с листа «вычитка»."""
    import apply_review as AR
    from assemble import review_xlsx
    from openpyxl import load_workbook
    src = tmp_path / "d.docx"
    dx = Document()
    dx.add_paragraph("A szelep zárva.")
    dx.save(str(src))
    segs = [{"id": 0, "text": "A szelep zárva.", "src": "A szelep zárva.", "page": None,
             "loc": {"part": "word/document.xml", "p": 0, "groups": [[0]]}}]
    res = {"0": {"ru": "Клапан закрыт.", "issues": []}}
    jf = tmp_path / "d.json"
    jf.write_text(json.dumps({"file": str(src), "segs": segs, "res": res}, ensure_ascii=False), encoding="utf-8")
    xf = tmp_path / "d_вычитка.xlsx"
    review_xlsx(xf, segs, res, "d", polish={"edits": [], "remaining": []})
    wb = load_workbook(xf)
    wb["вычитка"]["G2"] = "Клапан закрыт полностью."
    wb.active = wb.sheetnames.index("шлифовка")
    wb.save(xf)
    assert AR.apply(jf, xf, tmp_path / "tm.json", tmp_path / "out") == [("A szelep zárva.", "Клапан закрыт полностью.")]


def test_runner_weights_cover_new_steps():
    from server.runner import W_REVIEW, W_FAST
    for st in ("извлечение", "перевод", "структура", "шлифовка", "сборка", "термины"):
        assert st in W_REVIEW and st in W_FAST
    assert "проверка" in W_REVIEW


# ---------------------------------------------------------------- без зашивок под UTS (ÜFK, 16.10) и распознанный PDF
def test_spans_autonum_text_labels_and_continuation_note():
    """Шапка над «A. | текст» — по столбцам; автонумерация — только если повторяет номера оригинала; номера и тире
    в тексте пунктов — вид списка; пометка о переносе между частями таблицы — прочь."""
    row = [{"geom": {"x0": 110, "x1": 174}}, {"geom": {"x0": 291, "x1": 398}}, {"geom": {"x0": 471, "x1": 531}}]
    assert B._span_starts(row, [58, 80, 231, 267, 464, 537], 5) == [0, 2, 4]
    h = lambda t, lv: {"k": "h", "level": lv, "runs": [{"t": t}]}
    assert B._autonum_ok([h("1 Bevezetés", 1), h("1.1 Cél", 2), h("2 Rendszer", 1)])
    assert not B._autonum_ok([h("1. Bevezetés", 1), h("DF 1.1. Cél", 2), h("DF 1.2. Más", 2)])
    assert not B._autonum_ok([h("Bevezetés", 1)])                      # номеров нет — автонумерация их не добавит
    item = lambda t: [p(t)]
    alpha = {"k": "list", "kind": "Disc", "items": [item("a. első"), item("b. második"), item("c. harmadik")]}
    dash = {"k": "list", "kind": "Disc", "items": [item("– ha az első"), item("– ha a második")]}
    D.text_labels([alpha, dash])
    assert (alpha["kind"], alpha["suffix"], dash["kind"]) == ("LowerAlpha", ".", "Dash")
    cell = lambda th: {"blocks": [], "th": th}
    t1 = {"k": "table", "rows": [[cell(True)] * 2, [cell(False)] * 2], "geom": {"page": 90}}
    note = {"k": "p", "runs": [{"t": "Folytatás a következő oldalon."}], "geom": {"page": 90}}
    t2 = {"k": "table", "rows": [[cell(False)] * 2], "geom": {"page": 91}}
    got = D.join_tables([t1, note, t2])
    assert len(got) == 1 and len(got[0]["rows"]) == 3


def ocr_doc(tmp_path, pages=8):
    """Распознанный PDF: текст — в сегментах OCR, в текстовом слое — только тире пунктов (как «Печать в PDF» 3SZ19)."""
    doc = fitz.open()
    segs = []
    sid = 0

    def seg(pno, text, x0, y0, x1, align="left", bold=False):
        nonlocal sid
        segs.append({"id": sid, "page": pno, "text": text,
                     "loc": {"bbox": [x0, y0, x1, y0 + 9], "frags": [[x0, y0, x1, y0 + 9]], "size": 8.3 + (sid % 3) / 10,
                             "bold": bold, "align": align, "family": "monospace", "ocr": True}})
        sid += 1
    for pno in range(pages):
        page = doc.new_page(width=595, height=842)
        seg(pno, "Azonosító: 3SZ19", 77, 35, 170)
        seg(pno, f"Oldalszám: {pno + 2}", 460, 48, 530)
        title = ["1. A turbina üzemzavarai"] * 3 + ["2. Szekunderköri berendezések"] * 3 + ["3. Egyéb"] * 2
        seg(pno, title[pno], 250, 86, 395, "center", True)     # бегущее название раздела — на листах раздела
        page.insert_text((100, 140 + 7), "-", fontsize=9)
        seg(pno, f"Ellenőrizni a turbinát {pno}.", 112, 140, 400)
        seg(pno, "1.4.", 100, 200, 120)
        seg(pno, "Tengelyelmozdulás megnövekedése", 160, 200, 330)
        seg(pno, "13", 500, 200, 512)
    doc.save(tmp_path / "ocr.pdf")
    return str(tmp_path / "ocr.pdf"), segs


def test_ocr_pdf_blocks_header_page_number_dash_items_toc_line(tmp_path):
    import docx_layout as DL
    src, segs = ocr_doc(tmp_path)
    blocks, info = DL.load(src, segs)
    assert info["font"] == "Courier New" and info["size"] == 8.5             # один кегль на документ, не оценка OCR
    heads = [b for b in blocks if b["k"] == "h"]
    assert len(heads) == 3                                                   # название раздела — один раз
    lists = [b for b in blocks if b["k"] == "list"]
    assert lists and all(b["kind"] == "Dash" for b in lists)
    toc = next(b for b in blocks if b["k"] == "p" and "Tengely" in "".join(r["t"] for r in b["runs"]))
    assert [r["t"] for r in toc["runs"]] == ["1.4.\t", "Tengelyelmozdulás megnövekedése\t", "13"]
    assert toc["runs"][1].get("rtab") and toc.get("seg")
    L = DL.layout(src, blocks, segs, info)
    assert L["page_tok"] and L["head_grid"]["rows"][0]["cells"][0] == "Azonosító: 3SZ19"
    assert L["first_is_body"] and not any("Azonosító" in "".join(r["t"] for r in b.get("runs", [])) for b in blocks)


# ---------------------------------------------------------------- 16.10, вечер: разделы, линии таблиц, распознанные PDF
def test_front_matter_section_and_page_numbers_restart(tmp_path):
    """Вводная часть ÜFK (оглавление и лист изменений, подвал «Tartalomjegyzék» без номера) — своим разделом;
    нумерация тела — с числа оригинала («1. oldal» на 7-м листе)."""
    blocks = [p("ÜZEMELTETÉSI FELTÉTELEK", 1, 300), p("Módosítás Nyilvántartás", 2, 100), p("Bevezetés", 3, 100)]
    mem = B.Memory([("ÜZEMELTETÉSI FELTÉTELEK", "ПРЕДЕЛЫ"), ("Módosítás Nyilvántartás", "Лист изменений"),
                    ("Bevezetés", "Введение"), ("Tartalomjegyzék", "Содержание"), ("85. oldal", "Страница 85")])
    L = dict(LAYOUT, stamp=[{"cells": ["85. oldal"], "bold": False}], stamp_widths=[400], stamp_aligns=["right"],
             page_tok="85.", num_start={"page": 3, "value": 1},
             front={"to": 2, "stamp": [{"cells": ["Tartalomjegyzék"], "bold": False}], "stamp_widths": [400],
                    "stamp_aligns": ["center"], "stamp_box": True, "stamp_inner": False})
    out = tmp_path / "f.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(blocks, str(out), L)
    doc = Document(str(out))
    assert len(doc.sections) == 3
    front, body = doc.sections[1], doc.sections[2]
    assert "Содержание" in front.footer._element.xml and "PAGE" not in front.footer._element.xml
    assert "PAGE" in body.footer._element.xml
    assert body._sectPr.find(W + "pgNumType").get(W + "start") == "1"
    assert front._sectPr.find(W + "pgNumType") is None


def test_toc_split_by_pages_is_one_and_headings_outside_it_have_no_level(tmp_path):
    """Оглавление, размеченное по листам, — одно поле (было четыре: 556 пунктов вместо 206); заголовок, которого в
    оглавлении оригинала нет, — без уровня структуры (регистр не важен: «1.1. BEVEZETÉS» / «1.1. Bevezetés»)."""
    toc = lambda keys, lv: {"k": "toc", "info": {"levels": lv, "keys": keys, "extra": [], "bold": [True] * lv,
                                                 "indent": [0] * lv, "figures": False, "page": 2},
                            "geom": {"page": 2}}
    got = D.join_tocs([toc(["1.1.BEVEZETÉS"], 2), toc(["4.1.1.Areaktor"], 3)])
    assert len(got) == 1 and got[0]["info"]["levels"] == 3 and len(got[0]["info"]["keys"]) == 2
    h = lambda t, lv: {"k": "h", "level": lv, "runs": [{"t": t, "font": "Tahoma", "size": 12, "b": True}],
                       "geom": {"page": 3, "top": 100}}
    mem = B.Memory([("1.1. Bevezetés", "1.1. Введение"), ("1.1.1. Megjegyzés", "1.1.1. Примечание"),
                    ("A reaktor", "Реактор")])
    heads = [h("1.1. Bevezetés", 2), h("A reaktor", 3), h("1.1.1. Megjegyzés", 3)]   # «A reaktor» — номер у автонумерации
    out = tmp_path / "t.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build(got + heads, str(out), LAYOUT)
    doc = Document(str(out))
    styles = {x.text.split()[-1]: x.style.name for x in doc.paragraphs if x.text}
    assert styles["Введение"] == "Heading 2" and styles["Реактор"] == "Heading 3"
    assert styles["Примечание"] == "Heading 3 NoTOC"
    assert doc.element.body.xml.count("TOC \\o") == 1
    # пунктов оглавления среди заголовков нашлось меньше 60 % — сверка выключена: поле TOC не пустеет
    B.Builder(mem, lambda img: png(), PROFILE).build(got + [h("1.1.1. Megjegyzés", 3)], str(out), LAYOUT)
    assert not any(x.style.name.endswith("NoTOC") for x in Document(str(out)).paragraphs)
    # а лишние заголовки (абзацы стилем заголовка, ИН-АСУТП-1172) сверку не выключают
    many = heads + [h(f"Abszatz {i} szövege", 2) for i in range(6)]
    B.Builder(mem, lambda img: png(), PROFILE).build(got + many, str(out), LAYOUT)
    styles = [x.style.name for x in Document(str(out)).paragraphs if x.text.startswith("Abszatz")]
    assert styles and all(st.endswith("NoTOC") for st in styles)


def test_cell_edges_follow_original_rules_and_builder_draws_them(tmp_path):
    """Линии вокруг ячеек — как в оригинале: у ÜFK нет внешних вертикалей и линии между «E.1.» и текстом; Word
    рисовал полную сетку."""
    g = lambda x0, x1, top: {"page": 5, "x0": x0, "x1": x1, "top": top, "bottom": top + 10, "first_x0": x0,
                             "size": 10, "lines": 1}
    rows = [[{"blocks": [p("Ellenőrzés", 5, 100, 150)], "th": True, "geom": g(150, 230, 100)},
             {"blocks": [p("Ciklusidő", 5, 100, 420)], "th": True, "geom": g(420, 500, 100)}],
            [{"blocks": [p("E.1.", 5, 120, 80)], "th": False, "geom": g(80, 100, 120)},
             {"blocks": [p("Ellenőrizni kell", 5, 120, 120)], "th": False, "geom": g(120, 300, 120)},
             {"blocks": [p("8 óra", 5, 120, 420)], "th": False, "geom": g(420, 450, 120)}]]
    hs = [(96, 75, 520), (116, 75, 520), (135, 75, 520)]                  # над шапкой, между строками, под таблицей
    vs = [(410, 96, 135)]                                                  # одна внутренняя вертикаль
    D.cell_edges(rows, lambda pn: (hs, vs))
    e = rows[1]
    assert e[0]["edges"] == {"top": True, "bottom": True, "left": False, "right": False}
    assert e[1]["edges"]["right"] and e[2]["edges"]["left"] and not e[2]["edges"]["right"]
    mem = B.Memory([("Ellenőrzés", "Контроль"), ("Ciklusidő", "Периодичность"), ("E.1.", "E.1."),
                    ("Ellenőrizni kell", "Необходимо проверить"), ("8 óra", "8 ч")])
    rows[0][0]["span"], rows[0][1]["span"] = 2, 1
    for c in rows[1]:
        c["span"] = 1
    tbl = {"k": "table", "rows": rows, "geom": {"page": 5, "top": 100}, "borders": True,
           "frame": {"xs": [75, 110, 410, 520]}}
    out = tmp_path / "e.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build([p("x", 5, 80), tbl], str(out), LAYOUT)
    xml = Document(str(out)).tables[0]._tbl.xml
    assert '<w:left w:val="nil"/>' in xml and "w:tcBorders" in xml and '<w:insideV w:val="nil"/>' in xml


def test_ocr_lines_reordered_and_wrapped_toc_number_joined():
    """Строка листа — слева направо, хоть OCR отдал «Turbina…» раньше «1.1.»; номер листа у перенесённого названия —
    на его второй строке — в той же строке оглавления."""
    import docx_layout as DL

    def s(i, t, x0, y0, x1, frags=None):
        frags = frags or [[x0, y0, x1, y0 + 7]]
        return {"id": i, "page": 0, "text": t, "loc": {"bbox": [x0, y0, x1, frags[-1][3]], "frags": frags}}
    segs = [s(11, "Turbina üzemzavari leállítása", 145, 214, 388), s(12, "1.1.", 93, 215, 112),
            s(13, "5", 500, 215, 504), s(17, "1.3.", 93, 238, 112),
            s(18, "Generátor teherledobás üresjáratig", 145, 238, 398, [[145, 238, 398, 245], [145, 250, 230, 257]]),
            s(19, "11", 500, 250, 509)]
    got = [[x["text"] for x in g] for g in DL._lines(segs)]
    assert got == [["1.1.", "Turbina üzemzavari leállítása", "5"], ["1.3.", "Generátor teherledobás üresjáratig", "11"]]


def ocr_page_doc(tmp_path):
    """Лист распознанного PDF: обычный и полужирный набор и курсив (текст тут шрифтом, но сегменты — как от OCR),
    таблица линиями, график — рамкой и кривой с «мусором» OCR."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    segs, y = [], 100
    lines = [[("Check the oil pump and ", "cour"), ("START", "cobo"), (" it then.", "cour")]] + \
        [[(f"Plain words written here as usual {i}.", "cour")] for i in range(12)] + \
        [[("Megjegyzes: ", "cobo"), ("the reactor will be slanted text here", "coit")]]
    for i, parts in enumerate(lines):
        x = 90
        for t, f in parts:
            page.insert_text((x, y + 8), t, fontname=f, fontsize=9)
            x += fitz.get_text_length(t, fontname=f, fontsize=9)
        segs.append({"id": i, "page": 0, "text": "".join(t for t, _ in parts),
                     "loc": {"bbox": [90, y, x, y + 10], "frags": [[90, y, x, y + 10]], "size": 9, "family": "monospace"}})
        y += 14
    for yy in (330, 350, 370):                                             # таблица 2×2
        page.draw_line((90, yy), (400, yy), width=0.7)
    for xx in (90, 250, 400):
        page.draw_line((xx, 330), (xx, 370), width=0.7)
    for i, (t, x0, y0) in enumerate((("Alfa", 100, 336), ("Beta", 260, 336), ("Gamma", 100, 356), ("Delta |", 260, 356))):
        page.insert_text((x0, y0 + 8), t, fontname="cour", fontsize=9)
        segs.append({"id": 100 + i, "page": 0, "text": t,
                     "loc": {"bbox": [x0, y0, x0 + 40, y0 + 10], "frags": [[x0, y0, x0 + 40, y0 + 10]], "size": 9,
                             "family": "monospace"}})
    page.draw_rect(fitz.Rect(150, 420, 450, 700), width=0.8)               # график
    page.draw_polyline([fitz.Point(160 + i * 5, 690 - (i * i) % 250) for i in range(58)], width=1.5)
    segs.append({"id": 200, "page": 0, "text": "oO", "loc": {"bbox": [440, 430, 447, 436], "frags": [[440, 430, 447, 436]],
                                                             "size": 5, "family": "monospace"}})
    doc.save(tmp_path / "o.pdf")
    return str(tmp_path / "o.pdf"), segs


def test_ocr_bold_italic_table_and_figure_from_page(tmp_path):
    """Распознанный лист: полужирное и курсив — по штриху на картинке (текстовый слой у 3SZ19 неполный), таблица — по
    линиям, ячейки — по сетке («|» вертикали прочь), график — картинкой, «мусор» OCR внутри него — не абзац."""
    import docx_layout as DL
    src, segs = ocr_page_doc(tmp_path)
    st = DL._pixel_style(fitz.open(src), segs)
    assert [(t.strip(), b) for t, b, _ in st[0]] == [("Check the oil pump and", False), ("START", True), ("it then.", False)]
    assert st[13][0][1] and not st[13][0][2] and st[13][-1][2] and not st[13][-1][1]
    assert 1 not in st                                                      # обычная строка — без выделения
    blocks, info = DL.load(src, segs)
    tables = [b for b in blocks if b["k"] == "table"]
    assert len(tables) == 1 and [[c["blocks"][0]["runs"][0]["t"] for c in r] for r in tables[0]["rows"]] == \
        [["Alfa", "Beta"], ["Gamma", "Delta"]]
    figs = [b for b in blocks if b.get("figure")]
    assert len(figs) == 1 and figs[0]["runs"][0]["img"]["png"] and figs[0]["geom"]["top"] <= 421
    assert not any("oO" in B.text_of(b.get("runs", [])) for b in blocks)


def test_label_paragraphs_become_list():
    """Абзацы «a.)», «b.)» подряд (дерево Word — простые абзацы) — список с той же меткой; продолжение пункта с
    отступом — в пункт; один «a.)» с отступом под «Korlátozó feltétel:» — тоже список; «f.)» после «a.)–e.)» — в него."""
    bs = [p("Korlátozó feltétel:", 9, 100, 57), p("a.) A fűtőelem tartaléka", 9, 120, 85),
          p("folytatás a pont alatt", 9, 135, 110), p("b.) A kazetták tartaléka", 9, 150, 85), p("Más szöveg", 9, 170, 57),
          p("Korlátozó feltétel:", 10, 100, 57), p("a.) Egyetlen pont", 10, 120, 85)]
    got = D.para_lists(bs)
    assert [b["k"] for b in got] == ["p", "list", "p", "p", "list"]
    assert got[1]["kind"] == "LowerAlpha" and got[1]["suffix"] == ".)" and len(got[1]["items"]) == 2
    assert len(got[1]["items"][0]) == 2                                     # продолжение — в пункте
    lst = {"k": "list", "kind": "LowerAlpha", "suffix": ".)", "items": [[p(f"{c}.) x")] for c in "abcde"]}
    got = D.para_lists([lst, p("f.) Az állapot", 14, 300, 57)])
    assert len(got) == 1 and len(got[0]["items"]) == 6


def test_roman_text_labels_make_roman_list():
    """«• i. …», «• ii. …» — двойной маркер: пункты с римскими метками в тексте — римский список (ÜFK I, EF 2.2.2)."""
    b = lst("Disc", [[p(f"{x}. ellenőrzés")] for x in ("i", "ii", "iii", "iv", "v")])
    D.text_labels([b])
    assert (b["kind"], b["suffix"]) == ("LowerRoman", ".")
    b = lst("Disc", [[p("h. első")], [p("i. második")], [p("j. harmadik")]])
    D.text_labels([b])
    assert b["kind"] == "LowerAlpha"


def test_ocr_line_split_paragraph_is_joined():
    """OCR разрезал абзац по строкам — один абзац: тот же край, вплотную ниже, без точки в конце, со строчной буквы;
    после точки или с заглавной — новый абзац."""
    import docx_layout as DL
    prev = {"k": "p", "runs": [{"t": "célja, hogy a Duna vízhőmérséklete ne", "b": False, "i": False}],
            "geom": {"page": 145, "x0": 127.4, "x1": 547, "top": 274.6, "bottom": 285.6, "lines": 1}, "seg": True}
    seg = lambda t, y: {"id": 1, "page": 144, "text": t, "loc": {"bbox": [127.4, y, 446, y + 11]}}
    assert DL._continues(prev, seg("haladja meg a 29,5 °C-ot", 286.6), 11)
    assert not DL._continues(prev, seg("Haladja meg", 286.6), 11)
    assert not DL._continues(prev, seg("haladja meg", 320), 11)            # не вплотную
    assert DL._continues(prev, seg("2.2 és 2.3 pontok figyelembe vétele mellett:", 286.6), 11)
    assert not DL._continues(prev, seg("2.2. Ha a terhelés", 286.6), 11)    # номер пункта — новый абзац
    DL._join(prev, [{"t": "haladja meg a 29,5 °C-ot", "b": False, "i": False}], {"bottom": 297.6, "x1": 446, "lines": 1})
    assert prev["runs"][0]["t"] == "célja, hogy a Duna vízhőmérséklete ne haladja meg a 29,5 °C-ot"
    assert prev["geom"]["lines"] == 2 and prev["geom"]["bottom"] == 297.6


def test_table_without_cell_padding_in_profile(tmp_path):
    """Профиль без полей ячеек (у документа нет таблиц с текстом в рамке: 3PR42) — таблица собирается, поля обычные."""
    prof = dict(PROFILE, table={"pad_left": None, "pad_top": None, "pad_bottom": None})
    rows = [[{"blocks": [p("Mért érték", 5, 100, 130)], "th": True, "geom": None, "span": 1,
              "edges": {"top": True, "bottom": True, "left": True, "right": True}},
             {"blocks": [p("80 MW", 5, 100, 260)], "th": True, "geom": None, "span": 1,
              "edges": {"top": True, "bottom": True, "left": True, "right": True}}]]
    tbl = {"k": "table", "rows": rows, "geom": {"page": 5, "top": 100}, "borders": True, "frame": {"xs": [120, 250, 400]}}
    out = tmp_path / "t.docx"
    B.Builder(B.Memory([("Mért érték", "Измеренное значение"), ("80 MW", "80 МВт")]), lambda img: png(), prof).build(
        [p("x", 5, 80), tbl], str(out), LAYOUT)
    assert "Измеренное значение" in Document(str(out)).tables[0]._tbl.xml


def test_grid_table_without_full_row_and_courier_body_is_not_code(tmp_path):
    """Сетка, где ни в одной строке нет всех столбцов (разрез в 1-й строке на 200, во 2-й на 250), — столбцов по сетке,
    не падает; основной шрифт Courier (3SZ19) — не «код»: перевод с метками берётся целиком, без Consolas."""
    cell = lambda t, span: {"blocks": [p(t, 5, 100, 100)], "th": False, "geom": None, "span": span,
                            "edges": {"top": True, "bottom": True, "left": True, "right": True}}
    tbl = {"k": "table", "rows": [[cell("Egy", 1), cell("Kettő", 2)], [cell("Három", 2), cell("Négy", 1)]],
           "geom": {"page": 5, "top": 100}, "borders": True, "frame": {"xs": [90, 200, 250, 400]}}
    mono = lambda t, b=False: {"t": t, "font": "Courier New", "size": 9, "b": b, "i": False, "color": "000000",
                               "body_mono": True}
    para = {"k": "p", "runs": [mono("Ha a nyomas magas, "), mono("ZÁRNI", True), mono(" a szelepet.")],
            "geom": {"page": 5, "top": 300, "bottom": 310, "x0": 80, "x1": 400, "first_x0": 80, "size": 9}, "seg": True}
    mem = B.Memory([("Egy", "Один"), ("Kettő", "Два"), ("Három", "Три"), ("Négy", "Четыре"),
                    ("<f1>Ha a nyomas magas, </f1><f2>ZÁRNI</f2><f3> a szelepet.</f3>",
                     "<f1>Если давление высокое, </f1><f2>ЗАКРЫТЬ</f2><f3> клапан.</f3>")])
    out = tmp_path / "g.docx"
    B.Builder(mem, lambda img: png(), PROFILE).build([p("x", 5, 80), tbl, para], str(out), LAYOUT)
    doc = Document(str(out))
    assert len(doc.tables[0].columns) == 3
    par = next(x for x in doc.paragraphs if "ЗАКРЫТЬ" in x.text)
    assert par.text == "Если давление высокое, ЗАКРЫТЬ клапан." and "Consolas" not in par._p.xml
    assert [r.bold for r in par.runs if r.text.strip()] == [None, True, None]


def test_floating_text_box_goes_back_to_its_page():
    """Надпись Word, лежащая в дереве у места привязки (лист 40 между листами 2 и 3, руководство ТЭ357), — к своему
    листу: оглавление снова одно поле."""
    toc = lambda pg: {"k": "toc", "info": {"levels": 2, "keys": [f"k{pg}"], "extra": [], "bold": [True, False],
                                           "indent": [0, 0], "figures": False, "page": pg}, "geom": {"page": pg}}
    bs = [p("Cím", 1, 100), toc(2), p("", 40, 250), p("Текст строки", 40, 254), p("", None), toc(3),
          p("Bevezetés", 4, 100), p("Szöveg", 40, 100), p("Vége", 41, 100)]
    for b in bs:
        if b.get("geom") and b["geom"]["page"] is None:
            b["geom"] = None
    got = D.join_tocs(D.reflow(bs))
    pages = [(b.get("geom") or {}).get("page") for b in got if b["k"] != "toc"]
    assert sum(1 for b in got if b["k"] == "toc") == 1
    assert [x for x in pages if x] == sorted(x for x in pages if x)        # листы по порядку
    texts = [B.text_of(b["runs"]) for b in got if b["k"] == "p"]
    assert texts.index("Текст строки") > texts.index("Bevezetés")


def test_scanned_pdf_keeps_old_path_outline_pdf_goes_by_segments(tmp_path):
    """Распознанный PDF — по сегментам листа, только если это не сканы: у скана формы линии — пиксели, рукопись в
    сегменты не попадает, и форма разваливалась (30RDGT, 30RDIT)."""
    import docx_layout as DL
    scan = fitz.open()
    for _ in range(2):
        pg = scan.new_page(width=595, height=842)
        pg.insert_image(fitz.Rect(40, 60, 555, 800), stream=fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 50, 70), False)
                        .tobytes("png"))
    scan.save(tmp_path / "scan.pdf")
    outline = fitz.open()
    for _ in range(2):
        pg = outline.new_page(width=595, height=842)
        pg.draw_rect(fitz.Rect(100, 100, 300, 120))
    outline.save(tmp_path / "outline.pdf")
    assert not DL.suits(tmp_path / "scan.pdf", [0, 1])
    assert DL.suits(tmp_path / "outline.pdf", [0, 1])
    assert not DL.suits(tmp_path / "outline.pdf", [])                     # не распознан — прежний путь


def _cell(text, page, x0, th=False):
    g = {"page": page, "top": 100.0, "bottom": 110.0, "x0": x0, "x1": x0 + 90, "first_x0": x0, "size": 10, "lines": 1}
    return {"blocks": [p(text, page, x0=x0)] if text else [], "th": th, "geom": g, "attrs": {}}


def test_row_cut_by_page_tail_goes_back_to_its_cells():
    """Строка, разрезанная листом: хвост («необходимо контроль» под 4-й ячейкой шапки) — в ту же ячейку; следующая
    строка «1 | …» нового листа — своя, хотя ячейка выше без точки (синтетика, 16.10)."""
    head = [_cell("Система", 1, 57, True), _cell("Расход", 1, 177, True), _cell("Вода", 1, 298, True),
            _cell("Выполнение уровень", 1, 418, True)]
    tail = [_cell("необходимо контроль", 2, 418)]
    row1 = [_cell("1", 2, 57, True), _cell("Частота датчик.", 2, 177), _cell("Ручной.", 2, 418)]
    got = D._cut_rows([head, tail, row1])
    assert len(got) == 2
    assert B.text_of(got[0][3]["blocks"][0]["runs"]) == "Выполнение уровень необходимо контроль"
    assert B.text_of(got[1][0]["blocks"][0]["runs"]) == "1"


def test_paragraph_cut_by_page_is_one():
    """Абзац, разрезанный листом, Word размечает двумя P — снова один; начатый с прописной после точки — свой."""
    a = p("Пуск регулятор блок уставка время выполняется расход", 3, 700)
    b = p("температура производится частота.", 4, 60)
    c = p("Давление пуск агрегат.", 5, 60)
    got = D.join_pages([a, b, c])
    assert [B.text_of(x["runs"]) for x in got] == [
        "Пуск регулятор блок уставка время выполняется расход температура производится частота.",
        "Давление пуск агрегат."]


def test_numbered_and_bullet_items_in_one_word_list_split():
    """Нумерованный и следом маркированный пункты Word кладёт одним L: вид пункта — по маркеру «•» в его тексте."""
    bullet = p("x", 2, 300)
    bullet["runs"] = [{"t": "•", "font": "Symbol", "size": 10, "b": False, "i": False, "color": "000000"},
                      {"t": " Канал подача", "font": "Arial", "size": 10, "b": False, "i": False, "color": "000000"}]
    got = D.relist([lst("Decimal", [[p("Турбина работа", 2, 260)], [p("Уставка документ", 2, 280)], [bullet]])])
    assert [b["kind"] for b in got] == ["Decimal", "Disc"]
    assert [len(b["items"]) for b in got] == [2, 1]


def test_justified_two_word_line_is_not_tab():
    """Строка из двух слов, растянутая выключкой до правого края ячейки, — не табуляция."""
    def ch(t, x0, top):
        return {"text": t, "x0": x0, "x1": x0 + 5, "top": top, "size": 10, "fontname": "Arial", "ncs": "DeviceGray",
                "non_stroking_color": (0,), "page_number": 1}
    line1 = [ch(c, 177 + 5 * i, 118) for i, c in enumerate("Агрегат")] + [ch(" ", 212, 118)] + \
        [ch(c, 257 + 5 * i, 118) for i, c in enumerate("работа")]
    line2 = [ch(c, 177 + 5 * i, 132) for i, c in enumerate("измерение.")]
    runs = D._runs([{"ch": c} for c in line1 + line2])
    assert "\t" not in B.text_of(runs)


def test_text_from_next_page_nested_in_table_cell_goes_out():
    """Подпись с листа 13, которую Word вложил в список последней ячейки таблицы листа 12, — в поток на своё место."""
    cap = p("Рисунок 1 – Данные", 13, 608)
    cell = {"blocks": [lst("Disc", [[p("Строка таблицы", 12, 190), lst("Disc", [[cap]])]])], "th": False,
            "geom": {"page": 12, "top": 190.0, "bottom": 200.0, "x0": 80, "x1": 300, "first_x0": 80, "size": 10}}
    table = {"k": "table", "rows": [[cell]], "geom": {"page": 12, "top": 180.0, "bottom": 200.0}}
    got = D.evict([table, p("Данные в табличном виде", 13, 76), p("Назначение элементов", 13, 630)])
    texts = [B.text_of(b["runs"]) for b in got if b["k"] == "p"]
    assert texts == ["Данные в табличном виде", "Рисунок 1 – Данные", "Назначение элементов"]
    inner = got[0]["rows"][0][0]["blocks"][0]["items"][0]
    assert [x["k"] for x in inner] == ["p"]                  # пустой вложенный список-обёртка убран
