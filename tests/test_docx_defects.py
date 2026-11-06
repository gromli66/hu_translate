# Дефекты сборки DOCX по просмотру глазами 06.11 (EYES_2026-11-06.md, E9_DESIGN.md §5): один дефект — один тест.
import fitz
from docx import Document
from docx.oxml.ns import qn

import docx_build as B
import docx_layout as DL
import docx_struct as D

from test_assembly_sheets import _sheets_pdf
from test_struct_docx import LAYOUT, PROFILE, png

KEYS = ("head_grid", "variants", "page_marked", "page_tok", "stamp", "stamp_widths", "stamp_box", "stamp_inner",
        "stamp_paras", "stamp_aligns", "stamp_merge", "head_rule", "head_imgs", "head_lines", "first_is_body",
        "sheet_no", "num_start")


def _build_sheets(tmp_path, pages=3, sizes=None, name="sheets", **kw):
    src = tmp_path / f"{name}.pdf"
    segs = _sheets_pdf(src, pages=pages, **kw)
    blocks, info = DL.load(str(src), segs, text=True)
    L = D.layout(str(src), blocks, content=info.get("content_boxes"))
    blocks = D.split_at_zones(blocks, L.get("variants"), D.sheet_closed(L, blocks))
    out = tmp_path / f"{name}.docx"
    bld = B.Builder(B.Memory([(s["text"], s["text"]) for s in segs]), lambda img: png(), PROFILE)
    lay = dict(LAYOUT, **{k: v for k, v in L.items() if k in KEYS})
    if sizes:
        lay["page_sizes"] = sizes
    bld.build(blocks, str(out), lay)
    return bld, Document(str(out)), L


def _text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def test_section_break_after_table_takes_no_room(tmp_path):
    """Разрыв раздела после таблицы (лист изменений 30THHAGT, каждый лист пути А): python-docx ставит разрыв в новый
    пустой абзац кеглем тела, а после таблицы уже стоит пустой абзац Word. Два пустых абзаца не влезали под таблицу,
    кончавшуюся у низа листа (722 pt при нижнем поле 733), и уходили на свой лист — пустая стр. 2 (просмотр 06.11).
    Теперь разрыв несёт сам абзац после таблицы, в 1 pt; а там, где таблицы нет, — абзац разрыва в 1 pt."""
    _, doc, _ = _build_sheets(tmp_path, pages=3)
    body = doc.element.body
    brks = [p for p in body.iterchildren(qn("w:p")) if p.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None]
    assert len(brks) >= 2
    for p in brks:
        assert not _text(p).strip()
        sp = p.find(qn("w:pPr") + "/" + qn("w:spacing"))
        assert sp is not None and sp.get(qn("w:lineRule")) == "exact" and int(sp.get(qn("w:line"))) <= 20
        assert int(sp.get(qn("w:before")) or 0) == 0 and int(sp.get(qn("w:after")) or 0) == 0
        prev = p.getprevious()
        # перед разрывом — не ещё один пустой абзац (тот, что Word ставит после таблицы)
        assert not (prev is not None and prev.tag == qn("w:p") and not _text(prev).strip()
                    and prev.find(".//" + qn("w:drawing")) is None)



def _sec_head(sec):
    return "".join(t.text for t in sec.header._element.iter(qn("w:t")))


def test_last_sheet_before_blank_tail_page_is_one_sheet(tmp_path):
    """Последний лист перед пустой страницей в конце оригинала (30THHAGT: стр. 19 пустая, без номера): граница
    последнего раздела — «листов в документе + 1», и пустой хвост считался листом раздела — раздел выходил «в два
    листа»: без шапки «N (прод.)» и без ужатия docx_fit; перелив последнего листа печатался с номером «17» (просмотр
    06.11). Пустой хвост теперь не лист раздела: раздел в один лист, как у остальных."""
    bld, doc, L = _build_sheets(tmp_path, pages=3, sizes=[(595, 842)] * 4)     # 4-я страница оригинала — пустая
    assert L["sheet_no"] == {1: 1, 2: 2, 3: 3}
    last = len(doc.sections) - 1
    assert bld.fit_want.get(last) == 1
    sec = doc.sections[last]
    assert sec.different_first_page_header_footer and "(прод.)" in _sec_head(sec)
