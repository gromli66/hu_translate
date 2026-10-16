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
def tagged_pdf(path, pages, tagged=True):
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
