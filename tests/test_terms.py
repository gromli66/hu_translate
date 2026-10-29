# Аудит 29.10, причина №7: значения без слов берут перевод из res, идентификаторы кода не переводятся,
# единые термины документа (детектор и подстановка под флагом).
import docx_build as B
import segments as S


def test_value_without_words_takes_res():
    """«1,7 kW» не требует перевода, но в res лежит «1,7 кВт» (fix_units): память отдаёт res, а не исходный текст."""
    assert not S.needs_translation("1,7 kW")
    m = B.Memory([("1,7 kW", "1,7 кВт"), ("6 t/h", "6 т/ч"), ("30÷42oC", "30÷42 °C")])
    assert m.get("1,7 kW") == "1,7 кВт"
    assert m.get("6 t/h") == "6 т/ч"
    assert m.get("30÷42oC") == "30÷42 °C"


def test_value_without_words_not_in_memory_stays():
    """Значения нет в памяти — как раньше: исходный текст, а не None (в missing не попадает)."""
    m = B.Memory([])
    assert m.get("02TD90F601") == "02TD90F601"
    assert m("02TD90F601") == "02TD90F601" and m.missing == []


def test_copied_segment_gets_units():
    """Копируемый сегмент в конвейере: единица по глоссарию при переводе на русский."""
    from translate import fix_units
    assert fix_units("1,7 kW") == "1,7 кВт" and fix_units("36 m3/h") == "36 м³/ч" and fix_units("40÷42 t/h") == "40÷42 т/ч"
    assert fix_units("30TK62P101") == "30TK62P101"
