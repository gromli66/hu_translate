# -*- coding: utf-8 -*-
"""Поиск надписей на скриншотах (UTS, 15.10): детектор строк находит светлое на тёмной кнопке, слова одной строки
из разных проходов склеиваются по порядку, оборванный ответ портала не теряет картинку."""
import json

import numpy as np
import fitz

import llm
import shots as SH


def button_shot(size=10):
    """Светлая панель: тёмная кнопка с белой надписью «Mentés» и тёмная подпись «Jogosultság» рядом."""
    doc = fitz.open(); pg = doc.new_page(width=260, height=40)
    pg.draw_rect(pg.rect, color=None, fill=(0.93, 0.93, 0.93))
    pg.draw_rect(fitz.Rect(10, 8, 110, 32), color=None, fill=(0.2, 0.22, 0.22))
    pg.insert_htmlbox(fitz.Rect(10, 13, 110, 32), "Mentés",
                      css=f"* {{font-family: sans-serif; font-size: {size}px; color: white; font-weight: bold; text-align: center}}")
    pg.insert_htmlbox(fitz.Rect(130, 13, 250, 32), "Jogosultság",
                      css=f"* {{font-family: sans-serif; font-size: {size}px; color: rgb(30,30,30); font-weight: bold}}")
    pix = pg.get_pixmap(dpi=72, alpha=False)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3).copy()


def test_detector_finds_light_label_on_dark_button(tmp_path):
    # без детектора Tesseract на этой картинке не находит ни одной надписи
    img = button_shot()
    boxes = SH.det_boxes(img)
    assert SH._det["error"] is None and any(10 <= b[0] and b[2] <= 110 and b[1] < 20 < b[3] for b in boxes)   # внутри кнопки
    texts = [lb["t"] for lb in SH.tess_labels(img, tmp_path / "w.png")]
    assert "Mentés" in texts and "Jogosultság" in texts


def test_words_of_one_line_join_in_reading_order():
    # детектор и Tesseract дают середины строки на 1–2 px врозь: «lezár» не должен уйти в отдельную надпись или вперёд
    words = [{"t": "lezár", "box": [148.6, 335.6, 183.1, 347.1], "h": 11.5, "c": 96},
             {"t": "Ideiglenesen", "box": [53.2, 335.6, 142.9, 350.0], "h": 14.4, "c": 95}]
    assert [lb["t"] for lb in SH._labels(words)] == ["Ideiglenesen lezár"]


def test_truncated_portal_answer_keeps_parsed_and_asks_rest(monkeypatch):
    labels = [{"t": t, "box": [0, i * 20, 50, i * 20 + 10], "h": 10, "c": 60} for i, t in enumerate(["Mentés", "Törlés", "Módosítás"])]
    asked = []

    def chat(messages, images=None, max_tokens=0, tag=""):
        asked.append(messages[0]["content"])
        if len(asked) == 1:                  # модель ушла в рассуждения: массив оборван на второй надписи
            return {"text": '[\n{"n": 1, "t": "Mentés", "lang": "hu"},\n{"n": 2, "t": "Tör', "finish": "length"}
        return {"text": json.dumps([{"n": 1, "t": "Törlés", "lang": "hu"}, {"n": 2, "t": "Módosítás", "lang": "hu"}]),
                "finish": "stop"}

    monkeypatch.setattr(llm, "chat", chat)
    out = SH.check_labels(b"png", labels, "t")
    assert [lb["t"] for lb in out] == ["Mentés", "Törlés", "Módosítás"]
    assert "1: Törlés\n2: Módosítás" in asked[1]       # дозапрос — только недостающие


def test_followup_repeats_while_portal_keeps_truncating(monkeypatch):
    # оборвался и дозапрос (UTS p26: 51 надпись, 10 + 19 за два ответа — остаток терялся): остаток переспрашивается снова
    names = ["Mentés", "Törlés", "Módosítás", "Kilépés"]
    labels = [{"t": t, "box": [0, i * 20, 50, i * 20 + 10], "h": 10, "c": 60} for i, t in enumerate(names)]
    answers = ['[\n{"n": 1, "t": "Mentés", "lang": "hu"},\n{"n": 2, "t": "Tö',
               '[\n{"n": 1, "t": "Törlés", "lang": "hu"},\n{"n": 2, "t": "Mó',
               json.dumps([{"n": 1, "t": "Módosítás", "lang": "hu"}, {"n": 2, "t": "Kilépés", "lang": "hu"}])]
    asked = []

    def chat(messages, images=None, max_tokens=0, tag=""):
        asked.append(messages[0]["content"])
        return {"text": answers[len(asked) - 1], "finish": "stop" if len(asked) == len(answers) else "length"}

    monkeypatch.setattr(llm, "chat", chat)
    assert [lb["t"] for lb in SH.check_labels(b"png", labels, "t")] == names
    assert len(asked) == 3 and "1: Módosítás\n2: Kilépés" in asked[2]


def test_followup_stops_when_it_adds_nothing(monkeypatch):
    # дозапрос снова ушёл в рассуждения и не дал ни одной надписи — следующий дал бы то же: не повторять
    labels = [{"t": t, "box": [0, i * 20, 50, i * 20 + 10], "h": 10, "c": 60} for i, t in enumerate(["Mentés", "Törlés", "Módosítás"])]
    asked = []

    def chat(messages, images=None, max_tokens=0, tag=""):
        asked.append(tag)
        return {"text": '[\n{"n": 1, "t": "Mentés", "lang": "hu"},\n{"n' if len(asked) == 1 else '[\n{"n', "finish": "length"}

    monkeypatch.setattr(llm, "chat", chat)
    assert [lb["t"] for lb in SH.check_labels(b"png", labels, "t")] == ["Mentés"]
    assert len(asked) == 2


def test_unescaped_quotes_and_braces_inside_label_text(monkeypatch):
    # модель пишет прямые кавычки внутри строки JSON (UTS p43: «„W aktiv", „Y aktiv" lassú…») — объект не разбирался
    # и надпись пропадала; фигурные скобки в тексте (GUID, p67) рвали поиск объекта
    quoted = '„W aktiv", „Y aktiv" lassú villogóra vált'
    labels = [{"t": t, "box": [0, i * 20, 50, i * 20 + 10], "h": 10, "c": 60}
              for i, t in enumerate(['„W aktív", „Y aktiv" lassu villogóra vált', "1 {31580b92-d15e}", "YP23 működött"])]
    text = ('```json\n[\n  {"n": 1, "t": "' + quoted + '", "lang": "hu"},\n'
            '  {"n": 2, "t": "1 {31580b92-d15e}", "lang": "hu"},\n  {"n": 3, "t": "YP23 működött", "lang": "hu"}\n]\n```')
    monkeypatch.setattr(llm, "chat", lambda *a, **k: {"text": text, "finish": "stop"})
    assert [lb["t"] for lb in SH.check_labels(b"png", labels, "t")] == [quoted, "1 {31580b92-d15e}", "YP23 működött"]


def test_only_hungarian_labels_go_to_translation(monkeypatch):
    # «переводить ли» портал решал через раз (английские окна VirtualBox, SUSE — UTS стр. 95, 19): решает язык надписи;
    # венгерское без диакритики (PARANCS) проходит, код («-») и ответ без lang — нет
    labels = [{"t": t, "box": [0, i * 20, 50, i * 20 + 10], "h": 10, "c": 90}
              for i, t in enumerate(["Machine Tools", "PARANCS", "Lomtár", "X0YZ11K500", "Shut Down"])]
    ans = [{"n": 1, "t": "Machine Tools", "lang": "en"}, {"n": 2, "t": "PARANCS", "lang": "hu"},
           {"n": 3, "t": "Lomtár", "lang": "hu"}, {"n": 4, "t": "X0YZ11K500", "lang": "-"}, {"n": 5, "t": "Shut Down"}]
    monkeypatch.setattr(llm, "chat", lambda *a, **k: {"text": json.dumps(ans), "finish": "stop"})
    assert [lb["t"] for lb in SH.check_labels(b"png", labels, "t")] == ["PARANCS", "Lomtár"]
