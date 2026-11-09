# Четыре дефекта вне сборки (задание 09.11, DEFECTS4_REPORT.md): один дефект — один тест.
import importlib
import sys
import types


def _fake_uno(monkeypatch, calls):
    """uno и com.sun.star без LibreOffice: Desktop пишет в calls, что открыто и куда сохранено каким фильтром."""
    class Doc:
        def __init__(self, url):
            self.url = url

        def storeToURL(self, url, props):
            calls.append(("store", self.url, url, dict((p.Name, p.Value) for p in props)["FilterName"]))

        def getDocumentIndexes(self):
            return types.SimpleNamespace(getCount=lambda: 0)

        def getTextFields(self):
            return types.SimpleNamespace(refresh=lambda: calls.append(("refresh",)))

        def close(self, _):
            calls.append(("close", self.url))

    class Desktop:
        def loadComponentFromURL(self, url, *_):
            calls.append(("load", url))
            return Doc(url)

        def terminate(self):
            pass

    sm = types.SimpleNamespace(createInstanceWithContext=lambda name, ctx: (
        Desktop() if name.endswith("Desktop") else types.SimpleNamespace(resolve=lambda u: ctx)))
    ctx = types.SimpleNamespace(ServiceManager=sm)
    uno = types.ModuleType("uno")
    uno.getComponentContext = lambda: ctx
    uno.systemPathToFileUrl = lambda p: "file://" + p.replace("\\", "/")
    mods = {"uno": uno}
    for name in ("com", "com.sun", "com.sun.star", "com.sun.star.beans", "com.sun.star.connection",
                 "com.sun.star.lang"):
        mods[name] = types.ModuleType(name)
    mods["com.sun.star.beans"].PropertyValue = type("PropertyValue", (), {})
    mods["com.sun.star.connection"].NoConnectException = type("NoConnectException", (Exception,), {})
    mods["com.sun.star.lang"].DisposedException = type("DisposedException", (Exception,), {})
    for name, mod in mods.items():
        monkeypatch.setitem(sys.modules, name, mod)
    monkeypatch.delitem(sys.modules, "lo_pdf", raising=False)
    lo = importlib.import_module("lo_pdf")
    monkeypatch.setattr(lo.subprocess, "Popen", lambda *a, **k: types.SimpleNamespace(wait=lambda t: 0, kill=lambda: 0))
    return lo


def _docx(path, body):
    import zipfile
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", '<w:document xmlns:w="w"><w:body>' + body + "</w:body></w:document>")
    return str(path)


def test_n2_pdf_printed_like_convert_to(monkeypatch, tmp_path):
    """N2 (01-02FMP0BIGT стр. 25 «22 (прод.)»): открытый через UNO DOCX LibreOffice раскладывает по частям, и первый
    лист раздела с titlePg оставался с шапкой продолжения прежнего раздела; «--convert-to» печатал верно. PDF — как
    «--convert-to»; через UNO — только DOCX с оглавлением/указателем (их пересчитывает только UNO) и .doc → DOCX."""
    import os
    calls = []
    lo = _fake_uno(monkeypatch, calls)

    def run(cmd, **kw):
        calls.append(("run", cmd))
        out = cmd[cmd.index("--outdir") + 1]
        os.makedirs(out, exist_ok=True)
        open(os.path.join(out, os.path.splitext(os.path.basename(cmd[-1]))[0] + ".pdf"), "wb").write(b"%PDF")
        return types.SimpleNamespace(returncode=0, stdout="", stderr="")
    monkeypatch.setattr(lo.subprocess, "run", run)

    plain = _docx(tmp_path / "a.docx", "<w:p><w:r><w:t>Номер страницы</w:t></w:r></w:p>")
    lo.convert(plain, str(tmp_path / "a.pdf"))
    assert [c[0] for c in calls] == ["run"] and "--convert-to" in calls[0][1], calls
    assert (tmp_path / "a.pdf").read_bytes() == b"%PDF"

    calls.clear()
    toc = _docx(tmp_path / "b.docx", r'<w:p><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r></w:p>')
    lo.convert(toc, str(tmp_path / "b.pdf"))
    stores = [c for c in calls if c[0] == "store"]
    assert "run" not in [c[0] for c in calls] and calls[0] == ("load", stores[0][1]), calls
    assert stores[0][1].endswith("b.docx") and stores[0][3] == "writer_pdf_Export", calls

    calls.clear()
    lo.convert(str(tmp_path / "c.doc"), str(tmp_path / "c.docx"))
    assert [c[0] for c in calls] == ["load", "store", "close"] and calls[1][3] == "MS Word 2007 XML", calls


def _stamp_page(angle=8.0):
    """Лист скана: шесть ровных строк текста и внизу штамп в две строки, повёрнутый на angle градусов; слова — как у
    Tesseract (t, key, line, box, h, c, ink), картинка — серая 300 dpi."""
    import fitz
    import numpy as np
    import segments as S
    doc = fitz.open()
    pg = doc.new_page(width=595, height=842)
    body = ["A határozatot a FEL007 sz eljárásrend alapján adtuk ki",
            "Az engedélyezés során figyelembe vett dokumentumok listája",
            "A Turbina Osztály beadványa alapján a tesztelési utasítások",
            "jóváhagyásra kerülnek a mellékelt dokumentumok szerint",
            "A határozat érvényes visszavonásig minden érintett számára",
            "Paks 2020 szeptember huszonkilenc osztályvezető aláírásával"]
    for i, t in enumerate(body):
        pg.insert_text((72, 100 + 16 * i), t, fontsize=11)
    piv = fitz.Point(400, 760)
    for i, t in enumerate(["Paksi Atomerőmű Zrt.", "Minőségfelügyeleti Osztály 1."]):
        pg.insert_text((400, 760 + 14 * i), t, fontsize=11, morph=(piv, fitz.Matrix(-angle)))
    pix = pg.get_pixmap(dpi=300, colorspace=fitz.csGRAY)
    gray = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
    words = []
    for x0, y0, x1, y1, t, b, ln, wn in pg.get_text("words"):
        box = [x0, y0, x1, y1]
        words.append({"t": t, "key": (b + 1, 1), "line": (b + 1, 1, ln + 1), "box": box, "h": y1 - y0, "c": 90.0,
                      "ink": round(S._ink(gray, box, 300), 3)})
    return words, gray


def test_stamp_tilted_block_stays_pixels():
    """Штамп 30RYBSZGT PG-144760 (наклон 8°): Tesseract читал его мусором, плашки стирали штамп, и два перевода
    («АЭС ЗАО. Контроль качества 97», «АЭС «Пакш»») ложились друг на друга. Блок слов, повёрнутый к строкам листа
    (PF.mark_tilt: угол самой резкой проекции краски на вертикаль, ≥ 5° и резче в ≥ 1,2 раза), — пиксели, как рукопись.
    Ровные строки листа — текст, как прежде."""
    import pagefit as PF
    import segments as S
    words, gray = _stamp_page()
    PF.mark_tilt(words, gray, 300)
    stamp = [w for w in words if w["box"][1] > 700]
    body = [w for w in words if w["box"][1] < 700]
    assert stamp and all(PF.tilted(w) for w in stamp), [(w["t"], w["tilt"]) for w in stamp]
    assert not any(PF.tilted(w) for w in body), [(w["t"], w["tilt"]) for w in body if PF.tilted(w)]
    pix = []
    kept = {tuple(w["box"]) for w in S.ocr_words(words, "", pix)[0]}
    assert not kept & {tuple(w["box"]) for w in stamp} and {tuple(w["box"]) for w in body} <= kept
    assert {tuple(w["box"]) for w in stamp} <= {tuple(b) for b in pix}       # штамп остаётся пикселями

    # ровный блок той же формы — не штамп (наклон 0,5°: скан чуть скошен)
    words, gray = _stamp_page(angle=0.5)
    PF.mark_tilt(words, gray, 300)
    assert not any(PF.tilted(w) for w in words)
