# -*- coding: utf-8 -*-
"""Глоссарий HU→RU: разбор md, поиск терминов в куске текста, список аббревиатур.

Поиск — по венгерским основам, не по эмбеддингам: венгерский агглютинативный
(kazetta → kazettá-t, megszorulás → megszorulásának), поэтому слово термина
сравнивается как префикс слова текста, у основы срезается конечная a/e
(удлиняется в á/é перед суффиксом). Прописные аббревиатуры (SZBV, KU, ÜFK)
ищутся с учётом регистра и границы: SZBV не находится внутри SZBVR."""
import re
import threading
from functools import lru_cache
from dataclasses import dataclass, field
from pathlib import Path

GLOSS_MD = None                                  # основной глоссарий задаётся проектом (projects/<имя>/project.json)

WORD = re.compile(r"[0-9A-Za-zÀ-ɏ]+")


@dataclass
class Entry:
    hu: str
    ru: str
    note: str
    section: str
    alts: list = field(default_factory=list)   # варианты-последовательности слов для поиска

    def line(self):
        s = f"{self.hu} → {self.ru}"
        return s + (f"  ({self.note})" if self.note else "")


def _is_abbr(w):
    letters = [c for c in w if c.isalpha()]
    return len(letters) >= 2 and all(c.isupper() for c in letters)


def _stem(w):
    lw = w.lower()
    if len(lw) >= 5 and lw[-1] in "ae":
        return lw[:-1]
    return lw


def _alts(hu, ru=""):
    """«A próba érvényessége / Érvényessége» → две последовательности (list); скобки — пояснения.
    Многоточие — пропуск внутри термина: «villamos létfontosságú ... rendszerek» → tuple, слова ищутся окном.
    Раньше многоточие резало термин на куски, и голое «rendszerek» срабатывало на любой строке (v3, 02.10)."""
    out = []
    core = re.sub(r"\([^)]*\)", " ", hu)          # пояснения в скобках — не часть термина
    if not WORD.search(core):                      # «(31AM)»-подобное целиком в скобках — оставить
        core = hu.replace("(", " ").replace(")", " ")
    parts = re.split(r"\s+/\s+", core)
    # «zárt / nyitott primerkör → уплотненный / разуплотненный первый контур»: косая — выбор слова, а не термина
    if len(parts) >= 2 and "/" in re.sub(r"\([^)]*\)", "", ru):
        a, b = WORD.findall(parts[0]), WORD.findall(parts[1])
        if len(a) == 1 and len(b) >= 2 and a[0].lower() != b[0].lower():
            parts[0] = " ".join(a + b[1:])
    for part in parts:
        ws = WORD.findall(part)
        if not ws:
            continue
        # одиночное короткое строчное слово («a», «az», «és») даст шум
        if len(ws) == 1 and len(ws[0]) < 3 and not _is_abbr(ws[0]):
            continue
        out.append(tuple(ws) if ("…" in part or "..." in part) and len(ws) > 1 else ws)
    # «nem oldott / oldott → нерастворенный»: вариант без «nem» — противоположный смысл, не тот же термин
    low = [tuple(w.lower() for w in s) for s in out]
    return [s for s, l in zip(out, low) if ("nem",) + l not in low]


def load(path=None):
    path = path or GLOSS_MD
    if not path:
        raise ValueError("не задан основной глоссарий (project.json: glossary)")
    text = Path(path).read_text(encoding="utf-8")
    rules = text.split("## Правила применения", 1)[1].split("\n## ", 1)[0].strip() if "## Правила применения" in text else ""
    entries, section = [], ""
    for ln in text.splitlines():
        if ln.startswith("## "):
            section = ln[3:].strip()
            continue
        if not ln.startswith("|") or ln.startswith("|---") or ln.startswith("| Венгерский"):
            continue
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) < 2:
            continue
        hu, ru, note = cells[0], cells[1], (cells[2] if len(cells) > 2 else "")
        e = Entry(hu, ru, note, section)
        e.alts = _alts(hu, ru)
        if e.alts:
            entries.append(e)
    return rules, entries


def _tokens(text):
    return [(m.group(), m.start()) for m in WORD.finditer(text)]


@lru_cache(maxsize=500000)
def _word_hit(tw, pw):
    """tw — слово термина, pw — слово текста."""
    if _is_abbr(tw):
        # аббревиатура: точное совпадение токена (SZBV-t → токены SZBV, t); от 4 букв — и с заглавной буквы:
        # «Blügy» = «BLÜGY → НСБ» (без этого модель писала НСС, 05.10); короткие («KI», «ÉS») — только точно
        return pw == tw or (len(tw) >= 4 and pw[:1].isupper() and pw.upper() == tw)
    if tw[0].isdigit():
        return pw == tw
    st, lp = _stem(tw), pw.lower()
    if len(st) < 3:
        return lp == tw.lower()
    # хвост после основы — цепочка суффиксов, а не второе слово сложного (üzem|állapotba, üzem|mód, üzemel|tetési)
    return lp.startswith(st) and _is_suffix_chain(lp[len(st):])


SUFFIXES = sorted(set("""nak nek ban ben ból ből ról ről tól től hoz hez höz val vel nál nél ért ként kor ig ba be ra re
on en ön n t at et ot öt ök ok ek ak k i ai ei a e á é ja je já jé uk ük unk ünk juk jük ság ség ú ű s os es ös as ul ül
vá vé ának ének jának jének ait eit jait jeit ra re stul""".split()), key=len, reverse=True)


@lru_cache(maxsize=200000)
def _is_suffix_chain(tail, max_pieces=3):
    """Хвост раскладывается не более чем на max_pieces суффиксов (megszorulás|ának, kazett|á|t, kapcsoló|já|t)."""
    if not tail:
        return True
    best = {0: 0}
    for i in range(len(tail)):
        if i not in best or best[i] >= max_pieces:
            continue
        for s in SUFFIXES:
            if tail.startswith(s, i):
                j = i + len(s)
                if best.get(j, 99) > best[i] + 1:
                    best[j] = best[i] + 1
    return best.get(len(tail), 99) <= max_pieces


_INDEX = {}                                      # id(список) -> (список, длина, индекс)
_INDEX_LOCK = threading.Lock()


def _ekey(w):
    """Ключ индекса слова термина: первые 4 буквы основы (аббревиатура/число — слово целиком в нижнем регистре)."""
    if _is_abbr(w) or w[:1].isdigit():
        return w.lower()
    st = _stem(w)
    return st[:4] if len(st) >= 4 else st


def _candidates(toks, entries):
    """Индекс по ключу первого слова термина: с плотным глоссарием (~6200 записей) полный перебор на каждом
    сегменте занимал часы (02.10). Токен текста даёт ключи — сам токен и его префиксы длиной 2–4."""
    # кэш узнаёт список по ссылке на сам объект, а не по id(): id удалённого временного списка Python выдаёт
    # новому — индекс брался от чужого списка той же длины и термины молча терялись (тест, 06.10).
    # Мест несколько: проверка строки ходит по двум спискам (весь глоссарий и строгие термины), с одним местом
    # индекс на 7 тыс. записей строился заново на каждой строке — перевод из кэша шёл 105 с вместо секунд (06.10)
    with _INDEX_LOCK:
        c = _INDEX.get(id(entries))
        if c and c[0] is entries and c[1] == len(entries):
            idx = c[2]
        else:
            idx = {}
            for pos, e in enumerate(entries):
                for seq in e.alts:
                    idx.setdefault(_ekey(seq[0]), set()).add(pos)
            if len(_INDEX) >= 8:
                _INDEX.pop(next(iter(_INDEX)))
            _INDEX[id(entries)] = (entries, len(entries), idx)    # сильная ссылка: пока список в кэше, его id занят
    keys = set()
    for t in toks:
        lt = t.lower()
        keys.add(lt)
        keys.update(lt[:k] for k in (2, 3, 4))
    pos = set()
    for k in keys:
        pos |= idx.get(k, set())
    # многословный термин: каждое его слово должно иметь ключ в тексте — иначе окна можно не сканировать
    out = []
    for p in sorted(pos):
        e = entries[p]
        if any(all(_ekey(w) in keys for w in seq) for seq in e.alts):
            out.append(e)
    return out


def seq_spans(seq, toks, first=False):
    """Где в токенах текста стоит последовательность термина: [(i, j)] — полуинтервалы токенов."""
    n, out = len(seq), []
    if n > len(toks):
        return out
    for i in range(len(toks) - n + 1):
        if _word_hit(seq[0], toks[i]) and all(_word_hit(seq[k], toks[i + k]) for k in range(1, n)):
            out.append((i, i + n))
            if first:
                return out
    # термин из 3+ слов в другом порядке / с вставкой («biztonsági létfontosságú villamos betáplálási
    # rendszer» против записи «villamos biztonsági létfontosságú rendszer»): все слова в окне n+2;
    # термин с многоточием (tuple) — окном при любой длине
    if not out and (n >= 3 or isinstance(seq, tuple)):
        win = n + 2
        for i in range(max(1, len(toks) - win + 1)):
            chunk = toks[i:i + win]
            if all(any(_word_hit(w, t) for t in chunk) for w in seq):
                out.append((i, min(i + win, len(toks))))
                if first:
                    return out
    return out


def spans(text, e):
    """Все места записи e в тексте (по всем её вариантам)."""
    toks = [t for t, _ in _tokens(text)]
    return [sp for seq in e.alts for sp in seq_spans(seq, toks)]


def find(text, entries, limit=90):
    """Термины глоссария, встретившиеся в тексте. Длинные термины вперёд."""
    toks = [t for t, _ in _tokens(text)]
    hits = []
    for e in _candidates(toks, entries):
        for seq in e.alts:
            if seq_spans(seq, toks, first=True):
                hits.append((len(seq), e))
                break
    hits.sort(key=lambda h: -h[0])
    seen, out = set(), []
    for _, e in hits:
        if id(e) not in seen:
            seen.add(id(e)); out.append(e)
    return out[:limit]


def paren_abbr_hits(text, entries, found=()):
    """Записи, где сокращение из текста стоит в скобках: «turbinagenerátor (TG) → турбогенератор» при «TG» в тексте.
    Поиск по вариантам такие скобки отбрасывает, и модель видела только «TG → охлаждение БВ» (ложная тревога
    смысловой проверки 3SZ19, 02.10). Только подсказка — второе значение сокращения, не обязательный термин."""
    abbrs = {t for t, _ in _tokens(text) if _is_abbr(t) and len(t) <= 6}
    if not abbrs:
        return []
    have = {id(e) for e in found}
    pat = re.compile(r"\((%s)\)" % "|".join(map(re.escape, sorted(abbrs))))
    return [e for e in entries if id(e) not in have and pat.search(e.hu)][:6]


def abbreviations(entries):
    """Аббревиатуры раздела «Аббревиатуры» — для правила первого употребления."""
    out = {}
    for e in entries:
        if e.section != "Аббревиатуры":
            continue
        key = e.alts[0][0] if e.alts else ""
        if _is_abbr(key):
            out[key] = e
    return out


def abbr_in(text, abbrs):
    toks = {t for t, _ in _tokens(text)}
    return [a for a in abbrs if a in toks]


if __name__ == "__main__":
    import sys
    rules, ents = load()
    print("терминов:", len(ents), "аббревиатур:", len(abbreviations(ents)))
    sample = sys.argv[1] if len(sys.argv) > 1 else (
        "2.1. SZBV kazetta megszorulásának gyanúja esetén el kell végezni az alábbiakat: a reaktort 3 órán belül "
        "- M3 - üzemállapotba kell vinni. Az irányítástechnikai szolgálatnak az FI101 KU alapján az SZBVR "
        "szakcsoport. ÜFK 4.3.1. A reaktivitás szabályozás mechanikus rendszere")
    for e in find(sample, ents):
        print("  ", e.section[:12], "|", e.line()[:140])
