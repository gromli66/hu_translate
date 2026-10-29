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


def test_caps_word_from_code_names_is_identifier():
    """RUN рядом с RUN_XSTEP, MODE рядом с get_mode: слово, из которого документ строит имена, — не текст (IS_reko_09)."""
    texts = ["RUN_XSTEP", "REPLAY_ON", "Char400", "RUN", "MODE", "get_mode_val", "END", "KIESETT", "Állapot", "BYTE"]
    cw = S.code_words(texts)
    assert S.is_code_word("RUN", cw) and S.is_code_word("MODE", cw)
    assert not S.is_code_word("END", cw) and not S.is_code_word("KIESETT", cw)    # END нет среди имён — слово как слово
    assert not S.is_code_word("Run", cw)                                          # только прописными
    assert not S.needs_translation("BYTE") and S.needs_translation("Byte")        # тип данных кода


def test_kks_signal_suffix_is_not_a_code_name():
    """TH21S801_NYIT — сигнал с кодом KKS; NYIT в таблице сигналов остаётся венгерским словом (30RVR..VT)."""
    cw = S.code_words(["TH21S801_NYIT", "TJ23S803_NYIT", "NYIT"])
    assert not S.is_code_word("NYIT", cw)
    assert S.code_words(["Zárva", "NYITVA", "A szivattyú"]) == frozenset()
