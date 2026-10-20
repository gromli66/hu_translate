# -*- coding: utf-8 -*-
"""Перевод DOCX «на месте» (segments.docx_segments → assemble.build_docx): поля и прогоны."""
import zipfile

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

