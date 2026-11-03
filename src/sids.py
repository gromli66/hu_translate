# -*- coding: utf-8 -*-
"""E6 «Блок помнит свои сегменты» (UNIFIED_DESIGN §2 П5/П8, §5.3): блок сборки несёт `sids` — id сегментов, из текста
которых он собран. Память сборки (docx_build.Memory) берёт перевод блока по id, по тексту — только когда блок составной или
его текст не равен тексту сегмента; сборка помнит, какие id спрашивала (`used`); потеря = id с переводом, который сборка
не спросила и который никто не снял с причиной (колонтитул, рисунок, маркер пункта). Потеря — громко: `info["lost_ids"]`,
«ВНИМАНИЕ» в `info["attention"]` и строка листа «нарушения» (assemble.lost_id_rows).
Включение — HU_U_SIDS (по умолчанию включено; 0 — прежнее поведение побайтно: ключа `sids` в блоках нет, память
ищет перевод по тексту)."""
import os


def enabled():
    return os.environ.get("HU_U_SIDS", "1") != "0"


def of(group):
    """{"sids": [id, …]} для блока, собранного из сегментов group, — пустой словарь при выключенном флаге (блок как прежде)."""
    return {"sids": [s["id"] for s in group]} if enabled() else {}


def join(*lists):
    """Id нескольких блоков одним списком — в порядке появления, без повторов (склейка блоков)."""
    seen = []
    for lst in lists:
        for i in lst or ():
            if i not in seen:
                seen.append(i)
    return seen


def add(block, more):
    """Блок получает ещё и id more (его текст склеен из нескольких сегментов); без флага — не трогает."""
    if enabled() and more:
        block["sids"] = join(block.get("sids"), more)
    return block


def stamp(L, info):
    """Раскладка листа L несёт id сегментов, снятых нарочно (docx_layout.load, info["exempt"]: {id: причина})."""
    if enabled() and "exempt" in info:
        L["exempt_ids"] = info["exempt"]
    return L


def ids(blocks):
    """Id всех блоков списка, со вложенными списками и ячейками таблиц."""
    out = set()
    for b in blocks:
        out.update(b.get("sids") or ())
        if b["k"] == "list":
            for it in b["items"]:
                out |= ids(it)
        elif b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    out |= ids(c.get("blocks") or [])
    return out


def dropped(blocks):
    """Id сегментов, снятых при склейке блоков нарочно (повтор шапки таблицы на новом листе: docx_struct.join_tables)."""
    out = set()
    for b in blocks:
        out.update(b.get("sids_dropped") or ())
        if b["k"] == "list":
            for it in b["items"]:
                out |= dropped(it)
        elif b["k"] == "table":
            for r in b["rows"]:
                for c in r:
                    out |= dropped(c.get("blocks") or [])
    return out
