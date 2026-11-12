# -*- coding: utf-8 -*-
"""Копия скана чистым набором (src/rucopy.py): HTML портала → Word — абзацы, полужирное, таблица с объединёнными
ячейками, печать вырезкой скана; задание с «только копия» идёт без перевода (12.11)."""
import io
import json

import fitz
import numpy as np
from docx import Document

import rucopy as RC
from test_server import env, client, login, SECRET, last_uid, wait      # noqa: F401 — фикстуры сервиса
from server import jobs

HTML = """```html
<body>
<h1 align="center">ДЕКЛАРАЦИЯ О СООТВЕТСТВИИ</h1>
<p><b>Заявитель</b> Общество с ограниченной ответственностью "Авангард-НСК"</p>
<p>Место нахождения: Россия, 630099,
город Новосибирск</p>
<table border="1">
<tr><th>Наименование</th><th colspan="2">Размеры, мм</th></tr>
<tr><td rowspan="2">Лента</td><td>6</td><td>16</td></tr>
<tr><td>9</td><td>8</td></tr>
</table>
<p align="right">Сливкин Д. Н. <img alt="подпись" data-box="300,700,420,760"></p>
<img alt="печать" data-box="100,800,250,950">
<img alt="QR">
</body>
```"""


def scan(tmp_path):
    """Лист-«скан»: белая картинка, синий круг печати в рамке (100–250, 800–950) картинки портала 595×1000."""
    img = np.full((1000, 595, 3), 255, np.uint8)
    yy, xx = np.ogrid[:1000, :595]
    img[(yy - 875) ** 2 + (xx - 175) ** 2 < 60 ** 2] = (40, 60, 200)
    img[720:740, 310:410] = (20, 20, 20)                          # подпись
    d = fitz.open(); p = d.new_page(width=595, height=1000)
    p.insert_image(p.rect, pixmap=fitz.Pixmap(fitz.csRGB, 595, 1000, img.tobytes(), False))
    f = tmp_path / "scan.pdf"; d.save(f)
    return f


def test_html_to_word_text_table_and_figures(tmp_path):
    pdf = scan(tmp_path)
    out = tmp_path / "copy.docx"
    n = RC.build(pdf, [(HTML, 595, 1000, None)], out)
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs if p.text.strip()]
    assert texts[0] == "ДЕКЛАРАЦИЯ О СООТВЕТСТВИИ" and doc.paragraphs[0].alignment == 1
    assert "Место нахождения: Россия, 630099, город Новосибирск" in texts                   # строки — один абзац
    p = next(p for p in doc.paragraphs if p.text.startswith("Заявитель"))
    assert p.runs[0].bold and not p.runs[1].bold
    t = doc.tables[0]
    assert len(t.rows) == 3 and len(t.columns) == 3
    assert t.cell(0, 1).text == t.cell(0, 2).text == "Размеры, мм"                         # colspan
    assert t.cell(1, 0).text == t.cell(2, 0).text == "Лента"                               # rowspan
    assert t.cell(0, 0).paragraphs[0].runs[0].bold                                         # шапка th — полужирная
    assert len(doc.inline_shapes) == 2                                                     # подпись и печать
    assert any(p.text.endswith("[QR]") for p in doc.paragraphs)                           # без рамки — пометка
    assert n >= 5


def test_crop_tightens_to_ink(tmp_path):
    pdf = scan(tmp_path)
    with fitz.open(pdf) as d:
        png, w, h = RC.crop(d[0], (0.1, 0.75, 0.5, 0.99))          # рамка портала с большим запасом
    assert 120 <= w <= 135 and 120 <= h <= 135                     # круг 120 pt + поля краски


def test_run_writes_copy_per_pdf_without_translation(tmp_path, monkeypatch):
    src = tmp_path / "in"; src.mkdir()
    pdf = scan(tmp_path)
    (src / "Декларация.pdf").write_bytes(pdf.read_bytes())
    (src / "notes.docx").write_bytes(b"x")
    monkeypatch.setattr(RC, "read_page", lambda f, pno, lang: (HTML, 595, 1000, None))
    n = RC.run(src, tmp_path / "work", "ru")
    assert n >= 5 and [f.name for f in (tmp_path / "work" / "out").iterdir()] == ["Декларация.docx"]


def test_page_not_read_is_marked_loudly(tmp_path, capsys):
    pdf = scan(tmp_path)
    out = tmp_path / "copy.docx"
    RC.build(pdf, [("", 595, 1000, "HTTP 500")], out)
    assert "[лист 1 не прочитан: HTTP 500]" in [p.text for p in Document(str(out)).paragraphs]
    assert "не прочитан" in capsys.readouterr().out


def test_only_copy_option_goes_to_job(client, env):
    login(client, "alice")
    client.post("/profile/token", data={"token": SECRET})
    r = client.post("/jobs", files=[("files", ("a.pdf", io.BytesIO(b"doc"), "application/pdf"))],
                    data={"project": "demo", "only_copy": "true"}, follow_redirects=False)
    assert r.headers["location"] == "/?m=queued"
    uid = last_uid()
    assert json.loads((jobs.job_dir(uid) / "job.json").read_text(encoding="utf-8"))["only_copy"] is True
    wait(uid, "done", "failed")
