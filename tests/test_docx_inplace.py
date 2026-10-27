# -*- coding: utf-8 -*-
"""Перевод DOCX «на месте» (segments.docx_segments → assemble.build_docx): поля и прогоны."""
import zipfile

from lxml import etree

import assemble as A
import segments as S

W = S.W


def _docx(tmp_path, body):
    f = tmp_path / "f.docx"
    xml = (f'<w:document xmlns:w="{W[1:-1]}"><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(f, "w") as z:
        z.writestr("word/document.xml", xml)
    return f


def test_page_field_result_is_not_text(tmp_path):
    # «Стр.» + PAGE: результат поля «24» шёл в сегмент, перевод «Page 24» + пересчитанное поле = «Page 2402»
    f = _docx(tmp_path, '<w:p><w:r><w:t>Стр.</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                        '<w:r><w:instrText> PAGE \\* MERGEFORMAT </w:instrText></w:r>'
                        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>24</w:t></w:r>'
                        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
                        '<w:p><w:r><w:t>см. стр. </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                        '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>'
                        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>'
                        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    segs = S.docx_segments(f)
    assert [s["text"] for s in segs] == ["Стр.", "см. стр. 7"]       # ссылка на страницу (PAGEREF) — как раньше


def _runs_after(tmp_path, body, ru):
    """Сборка одного абзаца с переводом ru: [(текст прогона, его rPr-теги)]."""
    f = _docx(tmp_path, body)
    segs = S.docx_segments(f)
    out = tmp_path / "o.docx"
    A.build_docx(f, out, segs, {"0": {"ru": ru}})
    with zipfile.ZipFile(out) as z:
        p = next(etree.fromstring(z.read("word/document.xml")).iter(W + "p"))
    return [(S.run_text(r), [c.tag.split("}")[1] for c in (r.find(W + "rPr") if r.find(W + "rPr") is not None else [])])
            for r in S._own_runs(p)]


def test_highlighted_first_word_does_not_paint_whole_paragraph(tmp_path):
    # выделено одно первое слово — прежде перевод всего абзаца ложился в него, и жёлтым выходил весь абзац
    body = ('<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>печатном виде</w:t></w:r>'
            '<w:r><w:t>. Документация должна быть составлена в Microsoft Word.</w:t></w:r></w:p>')
    assert _runs_after(tmp_path, body, "printed form. Documentation shall be prepared in Microsoft Word.") == [
        ("", ["highlight"]), ("printed form. Documentation shall be prepared in Microsoft Word.", [])]


def test_leading_tab_in_other_size_does_not_shrink_line(tmp_path):
    # табуляция кеглем 11 перед текстом кеглем по умолчанию — прежде вся строка печаталась кеглем табуляции
    body = ('<w:p><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:tab/></w:r>'
            '<w:r><w:t>управления должно быть соизмеримым</w:t></w:r></w:p>')
    assert _runs_after(tmp_path, body, "\tcontrol shall be commensurate") == [
        ("", ["sz"]), ("\tcontrol shall be commensurate", [])]


def test_text_does_not_move_into_form_field(tmp_path):
    # подпись + серая подсказка поля формы (SDT) длиннее подписи — перевод остаётся в подписи, а не уходит в поле
    body = ('<w:p><w:r><w:t>Engedélyes: </w:t></w:r><w:sdt><w:sdtContent><w:r><w:rPr><w:color w:val="808080"/></w:rPr>'
            '<w:t>Szöveg beírásához kattintson vagy koppintson ide.</w:t></w:r></w:sdtContent></w:sdt></w:p>')
    assert _runs_after(tmp_path, body, "Лицензиат: Нажмите здесь для ввода текста.") == [
        ("Лицензиат: Нажмите здесь для ввода текста.", []), ("", ["color"])]


def test_same_look_keeps_first_run(tmp_path):
    # оформление одинаковое (цвет «auto» = без цвета) — как прежде, в первый прогон
    body = ('<w:p><w:r><w:t>1.</w:t></w:r><w:r><w:tab/></w:r>'
            '<w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>A blokk üzemeltetése</w:t></w:r></w:p>')
    assert _runs_after(tmp_path, body, "1.\tЭксплуатация блока") == [
        ("1.\tЭксплуатация блока", []), ("", []), ("", ["color"])]


def test_document_code_is_not_text():
    # обозначение документа одним словом — идентификатор: модель «переводила» ДСША…УТЗ в DSNS…TSH (ru-en, 27.10)
    assert not S.needs_translation("ДСША.161458.1300-A.УТЗ.Д15")
    assert S.needs_translation("ГОСТ Р 50.05.01-2018")             # с пробелами — текст
    assert S.needs_translation("п/п")
    assert S.needs_translation("www.get-sim.ru")                   # без номера из трёх цифр — как раньше


def test_numbering_caption_text_is_translated(tmp_path):
    # «Рисунок %8 –» — текст номера списка Word (numbering.xml): у всех подписей рисунков оставался русским
    f = _docx(tmp_path, '<w:p><w:r><w:t>Текст</w:t></w:r></w:p>')
    num = (f'<w:numbering xmlns:w="{W[1:-1]}"><w:abstractNum w:abstractNumId="0">'
           '<w:lvl w:ilvl="0"><w:lvlText w:val="%1."/></w:lvl>'
           '<w:lvl w:ilvl="1"><w:lvlText w:val="Рисунок %2 –"/></w:lvl>'
           '<w:lvl w:ilvl="2"><w:lvlText w:val="П %1.%3"/></w:lvl></w:abstractNum></w:numbering>')
    with zipfile.ZipFile(f, "a") as z:
        z.writestr("word/numbering.xml", num)
    segs = S.docx_segments(f)
    assert [s["text"] for s in segs] == ["Текст", "Рисунок %2 –", "П %1.%3"]
    out = tmp_path / "o.docx"
    A.build_docx(f, out, segs, {"0": {"ru": "Text"}, "1": {"ru": "Figure %2 –"}, "2": {"ru": "Appendix %1"}})
    with zipfile.ZipFile(out) as z:
        vals = [l.get(W + "val") for l in etree.fromstring(z.read("word/numbering.xml")).iter(W + "lvlText")]
    assert vals == ["%1.", "Figure %2 –", "П %1.%3"]             # потерян номер %3 — остаётся оригинал


def test_bound_form_field_loses_binding_in_translated_paragraph(tmp_path):
    # «Редакция » + поле, привязанное к свойству «Состояние» = «1.0»: перевод ложится в первый прогон, поле пустеет,
    # LibreOffice/Word ставят в него свойство заново — «Revision 1.01.0» в шапке каждого листа
    body = ('<w:p><w:r><w:t xml:space="preserve">Редакция </w:t></w:r><w:sdt><w:sdtPr><w:dataBinding w:xpath="/x"/>'
            '</w:sdtPr><w:sdtContent><w:r><w:t>1.0</w:t></w:r></w:sdtContent></w:sdt></w:p>'
            '<w:p><w:sdt><w:sdtPr><w:dataBinding w:xpath="/y"/></w:sdtPr><w:sdtContent><w:r><w:t>2.0</w:t></w:r>'
            '</w:sdtContent></w:sdt></w:p>')
    f = _docx(tmp_path, body)
    segs = S.docx_segments(f)
    out = tmp_path / "o.docx"
    A.build_docx(f, out, segs, {"0": {"ru": "Revision 1.0"}})
    with zipfile.ZipFile(out) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    assert [p.xpath("string()") for p in root.iter(W + "p")] == ["Revision 1.0", "2.0"]
    assert [x.get(W + "xpath") for x in root.iter(W + "dataBinding")] == ["/y"]   # чужой абзац — с привязкой


def test_text_after_field_or_symbol_keeps_its_place(tmp_path):
    # «(рисунок » + REF «1» + «).»: перевод всей группы ложился в первый прогон, поле оставалось после «).» — после
    # пересчёта «(Figure 1).1»; стрелка шрифта Symbol посреди строки уезжала в конец
    body = ('<w:p><w:r><w:t xml:space="preserve">окно (рисунок </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText> REF _Ref1 \r \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>).</w:t></w:r></w:p>'
            '<w:p><w:r><w:t xml:space="preserve">Назначить </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr>'
            '<w:t>→</w:t></w:r><w:r><w:t xml:space="preserve"> Выбрать</w:t></w:r></w:p>')
    f = _docx(tmp_path, body)
    segs = S.docx_segments(f)
    assert [s["text"] for s in segs] == ["<f1>окно (рисунок </f1><f2>).</f2>", "<f1>Назначить </f1><f2> Выбрать</f2>"]
    out = tmp_path / "o.docx"
    A.build_docx(f, out, segs, {"0": {"ru": "<f1>window (Figure </f1><f2>).</f2>"},
                                "1": {"ru": "<f1>Assign </f1><f2> Select</f2>"}})
    with zipfile.ZipFile(out) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    assert ["".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p")] == ["window (Figure 1).",
                                                                                         "Assign → Select"]


def test_text_after_sym_or_inline_picture_keeps_its_place(tmp_path):
    # стрелка — элемент w:sym (текста у прогона нет): перевод «] Обучаемый…» ложился до стрелки, она уезжала в конец
    body = ('<w:p><w:r><w:t xml:space="preserve">] </w:t></w:r><w:r><w:sym w:font="Symbol" w:char="F0AE"/></w:r>'
            '<w:r><w:t xml:space="preserve"> Обучаемый выполняет</w:t></w:r></w:p>')
    f = _docx(tmp_path, body)
    assert [s["text"] for s in S.docx_segments(f)] == ["<f1>] </f1><f2> Обучаемый выполняет</f2>"]
    out = tmp_path / "o.docx"
    A.build_docx(f, out, S.docx_segments(f), {"0": {"ru": "<f1>]</f1><f2>The trainee performs</f2>"}})
    with zipfile.ZipFile(out) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    # пробелы по краям стрелки — как в оригинале (модель их срезала: «]→The trainee»)
    assert [S.run_text(r) for r in root.iter(W + "r")] == ["] ", "", " The trainee performs"]


def test_group_edge_space_not_doubled():
    # пробел на стыке групп уже есть у соседа — второй не добавляется
    runs = [etree.fromstring(f'<w:r xmlns:w="{W[1:-1]}"><w:t xml:space="preserve">{t}</w:t></w:r>') for t in ("a ", "b")]
    assert A._pad(["x ", "y"], [[0], [1]], runs) == ["x ", "y"]
    assert A._pad(["x", "y"], [[0], [1]], runs) == ["x ", "y"]
