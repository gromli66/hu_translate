# Четыре дефекта вне сборки (задание 09.11, DEFECTS4_REPORT.md): один дефект — один тест.


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
