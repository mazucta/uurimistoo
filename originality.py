"""Оригинальность и честность текста, то, что считает код без модели:
совпадения с другими загруженными работами, цитирование, несомненные следы чат-бота.

Процент «написано ИИ» здесь не считается и не показывается: по тексту его надёжно не установить.
Сильные признаки — это только то, что нельзя объяснить ничем, кроме копирования из чат-бота.
"""
import re

SHINGLE = 8                 # совпадение — это 8 слов подряд: общие фразы короче
LOCAL_MIN = 0.03            # меньше 3% совпавшего текста — шум (термины, названия, шаблонные фразы)
WORD = re.compile(r"[^\W\d_]+(?:[-'’][^\W\d_]+)*|\d+", re.U)


def words(text):
    return [w.lower() for w in WORD.findall(text)]


def _covered(mine, other):
    """Какие слова этой работы попали хоть в одно совпадение из SHINGLE слов подряд с другим текстом."""
    theirs = words(other)
    pool = {tuple(theirs[i:i + SHINGLE]) for i in range(len(theirs) - SHINGLE + 1)}
    covered = [False] * len(mine)
    for i in range(len(mine) - SHINGLE + 1):
        if tuple(mine[i:i + SHINGLE]) in pool:
            covered[i:i + SHINGLE] = [True] * SHINGLE
    return covered


def _runs(mine, mask, shortest=1):
    """До трёх первых сплошных кусков, где mask истинна, по 30 слов, не короче shortest слов."""
    pieces, i = [], 0
    while i < len(mask) and len(pieces) < 3:
        if mask[i]:
            j = i
            while j < len(mask) and mask[j]:
                j += 1
            if j - i >= shortest:
                pieces.append(" ".join(mine[i:min(j, i + 30)]))
            i = j
        else:
            i += 1
    return pieces


def overlap(text, others, min_share=LOCAL_MIN):
    """Совпадения работы с другими текстами: работами учителя, источниками, найденными в сети работами.
    others: [(подпись, текст)]. Ответ: [(подпись, доля совпавшего текста этой работы, до трёх совпавших кусков)]
    по убыванию доли; тексты с долей меньше min_share не попадают."""
    mine = words(text)
    if len(mine) < SHINGLE:
        return []
    out = []
    for label, other in others:
        covered = _covered(mine, other)
        share = sum(covered) / len(mine)
        if share >= min_share:
            out.append((label, share, _runs(mine, covered)))
    return sorted(out, key=lambda x: -x[1])


def changes(text, old):
    """Новая версия работы против прошлой того же ученика: (доля текста, оставшаяся как была, до трёх новых кусков).
    Новые куски короче SHINGLE слов не показываем: это правка пары слов, а не новый текст."""
    mine = words(text)
    if not mine:
        return 1.0, []
    covered = _covered(mine, old)
    return sum(covered) / len(mine), _runs(mine, [not c for c in covered], SHINGLE)


SOURCE_MIN = 0.01           # из источника: два-три переписанных предложения уже стоит показать учителю
ANY_QUOTE = re.compile(r"[«„“\"][^«»„“”\"\n]{1,600}[»“”\"]")


def unquoted(text):
    """Текст без кусков в кавычках: дословная цитата с кавычками — это цитирование, а не копирование."""
    return ANY_QUOTE.sub(" … ", text)


QUOTE = re.compile(r"[«„“\"]([^«»„“”\"]{40,400})[»“”\"]")


def citations(text, sources, cite_pattern):
    """Проблемы цитирования. sources: строки списка источников; cite_pattern: регулярка ссылок в тексте.
    Ответ: [(что не так, фрагмент)]."""
    out = []
    for m in QUOTE.finditer(text):  # длинная цитата должна иметь ссылку рядом: в той же фразе или сразу после
        after = text[m.end():m.end() + 160]
        if not cite_pattern.search(text[max(0, m.start() - 20):m.end()] + after.split(".")[0] + "."):
            out.append(("Цитата без ссылки на источник", m.group(1)[:160]))
    cited_numbers, cited_names = set(), []
    for c in cite_pattern.findall(text):
        nums = re.findall(r"^\[(\d{1,3})", c)
        if nums:
            n = int(nums[0])
            cited_numbers.add(n)
            if n > len(sources):
                out.append(("Ссылка на источник, которого нет в списке", c))
            continue
        # автор — слово с большой буквы («составлено автором по данным Elering, 2026a» → Elering)
        names = [w.lower() for w in re.findall(r"\b[^\W\d_a-zа-яёіїєґõäöü][^\W\d_]{2,}", c, re.U)]
        if names:
            cited_names.append((c, names))
    lowered = [s.lower() for s in sources]
    for c, names in cited_names:
        if lowered and not any(n in s for s in lowered for n in names):
            out.append(("Ссылка на источник, которого нет в списке", c))
    used = {n - 1 for n in cited_numbers} | {i for i, s in enumerate(lowered) for _c, names in cited_names
                                             if any(n in s for n in names)}
    if cited_numbers or cited_names:  # если ссылок по тексту нет вовсе, об этом и так скажут условия учителя
        for i, s in enumerate(sources):
            if i not in used:
                out.append(("Источник из списка ни разу не упомянут в тексте", s[:160]))
    return list(dict.fromkeys(out))  # одна и та же ссылка под каждой таблицей — одно замечание


NO_DATE = re.compile(r"\((?:n\.\s?d\.|н\.\s?д\.|б\.\s?д\.|б\.\s?г\.|s\.\s?a\.|i\.\s?a\.)\)", re.I | re.U)
YEAR = re.compile(r"\b(1[5-9]\d\d|20\d\d)\b")


ORDER = "абвгґдеёєжзиіїйклмнопрстуфхцчшщъыьэюя" + "abcdefghijklmnopqrsšzžtuvwõäöüxy"  # укр./рус. и эст. алфавиты


def alpha(text):
    """Ключ сортировки по алфавиту, а не по кодам букв: иначе украинская «і» окажется после «п»."""
    return [ORDER.index(ch) if ch in ORDER else 1000 + ord(ch) for ch in text]


def bibliography(sources):
    """Оформление списка источников, то, что видно без стандарта: порядок, год, единообразие «без даты»,
    пробелы, повторы. Ответ: [(что не так, запись)]."""
    out = []
    keys = [re.sub(r"^[\d\s.)]+", "", s).strip().lower() for s in sources]
    for script in (r"[a-zõäöüšž]", r"[а-яёіїєґ]"):  # по алфавиту внутри каждой азбуки: латиница и кириллица отдельно
        group = [(k, s) for k, s in zip(keys, sources) if re.match(script, k)]
        for (prev, _), (cur, entry) in zip(group, group[1:]):
            if alpha(cur) < alpha(prev):
                out.append(("Список не по алфавиту", entry[:160]))
                break
    marks = {m.group(0).lower().replace(" ", "")[1:2] for s in sources for m in NO_DATE.finditer(s)}
    if len(marks) > 1:
        out.append(("Разные обозначения «без даты» в одном списке", ", ".join(sorted(
            {m.group(0) for s in sources for m in NO_DATE.finditer(s)}))))
    for s in sources:
        if not YEAR.search(s) and not NO_DATE.search(s):
            out.append(("Нет года издания или пометки «без даты»", s[:160]))
        if re.search(r"[^\s(\[]https?://|\s[.,;:](?!\d)|\S {2,}\S", s):
            out.append(("Пробелы и знаки: слитно перед ссылкой, пробел перед точкой или двойной пробел", s[:160]))
    seen = set()
    for k, s in zip(keys, sources):
        short = re.sub(r"\W+", "", k)[:60]
        if short in seen:
            out.append(("Источник повторяется в списке", s[:160]))
        seen.add(short)
    return out


# Следы чат-бота: так ученик не пишет, так отвечает модель, а ученик скопировал ответ целиком.
BOT = re.compile("|".join([
    r"\bas an ai\b", r"\bas a language model\b", r"\bi hope this helps\b", r"\bcertainly!? here(?: is|'s)\b",
    r"\bsure!? here(?: is|'s)\b", r"\bknowledge cutoff\b",
    r"как (?:ии|искусственный интеллект|языковая модель)", r"\bконечно!? вот\b", r"надеюсь,? это поможет",
    r"як (?:ші|штучний інтелект|мовна модель)", r"\bзвісно!? ось\b", r"сподіваюся,? це допоможе",
    r"tehisintellektina", r"keelemudelina", r"\bloomulikult!? siin\b", r"loodan,? et see aitab",
]), re.I | re.U)
MARKDOWN = re.compile(r"\*\*[^*\n]{2,80}\*\*|^#{2,4} \S|^```", re.M)
PLACEHOLDER = re.compile(r"\[(?:вставьте|вставте|укажите|вкажіть|ваше|ваш|ваша|your|insert|sisesta|lisa)[^\]]{0,60}\]", re.I | re.U)
INVISIBLE = re.compile("[​‌‍⁠﻿]")
MIXED = re.compile(r"\b(?=[^\W\d_]*[а-яіїєґё])(?=[^\W\d_]*[a-z])[^\W\d_]{4,}\b", re.I | re.U)


def ai_signs(text):
    """Несомненные следы: фраза чат-бота, остатки разметки, незаполненная заготовка, невидимые символы,
    слова из смеси кириллицы и латиницы (так обходят проверку на совпадения). Ответ: [(признак, фрагмент)]."""
    out = []
    for pattern, label in ((BOT, "Фраза чат-бота"), (PLACEHOLDER, "Незаполненная заготовка из ответа чат-бота"),
                           (MARKDOWN, "Остатки разметки чат-бота (**, ###, ```)")):
        for m in list(pattern.finditer(text))[:3]:
            out.append((label, text[max(0, m.start() - 40):m.end() + 40].replace("\n", " ").strip()))
    hidden = INVISIBLE.findall(text)
    if hidden:
        out.append(("Невидимые символы в тексте", f"{len(hidden)}"))
    mixed = [m.group(0) for m in MIXED.finditer(text)]
    if len(mixed) >= 3:  # одно-два слова бывают опечаткой раскладки, много — уже подмена букв
        out.append(("Слова из смеси кириллицы и латиницы: так прячут текст от проверки на совпадения", ", ".join(mixed[:6])))
    return out


if __name__ == "__main__":
    base = ("Отрицательные цены на электроэнергию возникают в часы, когда солнечные станции производят "
            "больше энергии, чем может потребить сеть, и производитель доплачивает за отбор мощности")
    hits = overlap(base + " и ещё моя собственная мысль про Эстонию", [("Иванов", "Вступ. " + base + ". Конец."), ("Другой", "совсем другой текст")])
    assert len(hits) == 1 and hits[0][0] == "Иванов" and hits[0][1] > 0.7 and hits[0][2], hits
    assert overlap("коротко", [("x", "коротко")]) == []
    source = base + ". Дальше в источнике идёт совсем другой текст про рынок"
    assert overlap(unquoted(f"Как пишет автор, «{base}» (Hirth, 2013)."), [("1", source)], SOURCE_MIN) == []
    copied = overlap(unquoted(f"Как пишет автор, {base} (Hirth, 2013)."), [("1", source)], SOURCE_MIN)
    assert copied and copied[0][1] > 0.5, copied
    same, fresh = changes(base + ". Новый раздел про то, как я сам опросил тридцать одноклассников в школе.", base)
    assert 0.5 < same < 1 and fresh == ["новый раздел про то как я сам опросил тридцать одноклассников в школе"], (same, fresh)
    assert changes(base, base) == (1.0, []) and changes("", base) == (1.0, [])
    assert overlap(base, [("чисто", "ничего общего у этих текстов нет совсем никак вообще")], 0)[0][1] == 0

    cite = re.compile(r"\([^()]{2,90}?\d{4}[^()]{0,25}\)|\[\d{1,3}\]")
    srcs = ["Hirth L. The market value of solar power. 2013.", "Elering. Andmed. 2025.", "Никто не цитирует 2020."]
    assert citations("Таблица (составлено автором по данным Hirth, 2013). Ещё (составлено автором по данным Hirth, 2013).",
                     srcs, cite) == [("Источник из списка ни разу не упомянут в тексте", s) for s in srcs[1:]]
    text = ("По данным Hirth (Hirth, 2013) стоимость падает. «Солнечная энергия дешевеет быстрее, чем ожидали аналитики "
            "всех крупных агентств» — и это важно. Таблица [2]. Ещё [7] и (Петров, 2019).")
    problems = citations(text, srcs, cite)
    kinds = [p[0] for p in problems]
    assert kinds.count("Цитата без ссылки на источник") == 1, problems
    assert "[7]" in [p[1] for p in problems] and "(Петров, 2019)" in [p[1] for p in problems], problems
    assert any(p[0].startswith("Источник из списка ни разу") and "Никто" in p[1] for p in problems), problems
    assert not any(p[0].startswith("Источник из списка ни разу") and "Hirth" in p[1] for p in problems), problems
    assert citations("Без ссылок вовсе.", srcs, cite) == []

    biblio = bibliography(["1. Petrov P. Kiberport. 2020.", "2. Abels A. (n.d.)https://a.ee", "3. Борисов Б. (н.д.). Текст .",
                           "4. Агеев А. Без года.", "5. Abels A. (n.d.)https://a.ee"])
    kinds = [b[0] for b in biblio]
    for need in ("Список не по алфавиту", "Разные обозначения «без даты» в одном списке", "Нет года издания или пометки «без даты»",
                 "Источник повторяется в списке"):
        assert need in kinds, (need, biblio)
    assert any(k.startswith("Пробелы") for k in kinds), biblio
    assert bibliography(["Aaa A. 2020. https://a.ee (vaadatud 01.10.2026).", "Bbb B. 2021."]) == []
    assert bibliography(["Імас Є. 2021.", "Подрігало Л. 2024.", "Шинкарук О. 2021."]) == []  # і перед п по-украински
    assert bibliography(["Saar S. 2020.", "Šmidt Š. 2020.", "Zeiger Z. 2020.", "Tamm T. 2020."]) == []  # эстонский порядок
    assert bibliography(["Aaa A. 2020. https://a.ee", "Bbb B. 2021. www.b.ee"]) == []  # дату обращения не требуем

    signs = ai_signs("Конечно! Вот ваш реферат.\n## Введение\nТекст [вставьте название школы]​. "
                     "Слова pуcский, cвoй, нaш, тeкст.")
    labels = [s[0] for s in signs]
    assert "Фраза чат-бота" in labels and any("разметки" in x for x in labels) and any("заготовка" in x for x in labels), labels
    assert "Невидимые символы в тексте" in labels and any("кириллицы" in x for x in labels), labels
    assert ai_signs("Обычный текст ученика без странностей, с одним словом Elering.") == []
    print("ok")
