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
