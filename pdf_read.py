"""Чтение .pdf: тот же разбор работы, что и у .docx, только оформление приходится измерять.

В PDF нет стилей и абзацев: есть буквы в точках страницы. Поэтому строки собираются по высоте,
абзацы по отступам и разрывам между строками, а кегль, шрифт, поля и интервал измеряются по ним.
Чего в PDF узнать нельзя (отбивка абзаца), перечислено в UNKNOWN: такие правила не проверяются.
"""
import re, statistics
from pypdf import PdfReader

from docx_read import APPENDIX, CAPTION, Paper, SOURCE_HEADS

PT_CM = 28.35            # точек PDF в сантиметре
SINGLE = 1.15            # одинарный интервал Word примерно такой от кегля
FOOTER = 50              # всё ниже этой высоты считаем колонтитулом: текст с полем 2 см кончается на 57
MAX_PAGES = 300
INDENT = (20, 60)        # так выглядит абзацный отступ: примерно от 0,7 до 2 см
EDGE = 12                # на столько строка может не дотянуть до правого поля и всё ещё считаться полной
# Отбивку абзаца, выравнивание по ширине и правое поле по PDF не измерить: в нём нет
# ни ширины букв конкретного шрифта, ни границ абзаца, и правый край строки приходится угадывать.
UNKNOWN = frozenset({"Отбивка абзаца 6 пт до и после", "Выравнивание по ширине"})
LIST_ITEM = re.compile(r"^\s*(?:\d{1,2}\s*[.)]|[•·–-])\s*")


# Ширина буквы в долях кегля: в PDF её нет, шрифт подставляет subset без таблицы на Unicode.
# Значения для Times New Roman, ошибка на строке в пределах пары процентов.
NARROW, WIDE, CAPS = "ijlt.,:;!'|()[]", "mwMWÜŠŽ", "ABCDEFGHIJKLMNOPQRSTUVWXYZАБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЭЮЯЄІЇҐÕÄÖÜ"


def text_width(text, size):
    em = 0.0
    for ch in text:
        em += 0.25 if ch == " " else 0.3 if ch in NARROW else 0.78 if ch in WIDE else 0.68 if ch in CAPS else 0.5
    return em * size


def median(values, default=0.0):
    return statistics.median(values) if values else default


def common(values, step=1.0, default=0.0):
    """Самое частое значение с точностью до шага: так находится левое поле."""
    if not values:
        return default
    counts = {}
    for v in values:
        counts[round(v / step) * step] = counts.get(round(v / step) * step, 0) + 1
    return max(counts, key=counts.get)


def mark_blocks(paragraphs):
    """Список источников и подписи к таблицам оформлены по-своему: под правила абзаца они не идут."""
    sources = False
    for p in paragraphs:
        if p["level"]:
            sources = p["text"].strip().lower().startswith(SOURCE_HEADS)
        elif sources or p["text"].lower().startswith(CAPTION):
            p["block"] = True
    return paragraphs


class Doc(Paper):
    """Работа, прочитанная из PDF."""

    unknown = UNKNOWN

    def __init__(self, data):
        reader = PdfReader(data)
        if not reader.pages:
            raise ValueError("в файле нет страниц")
        self.page_size = (float(reader.pages[0].mediabox.width), float(reader.pages[0].mediabox.height))
        self.lines, self.footers = [], []
        for number, page in enumerate(reader.pages[:MAX_PAGES], 1):
            self._read_page(page, number)
        if not self.lines:
            raise ValueError("в файле нет текста, возможно это скан")
        inner = [l for l in self.lines if l["page"] > 1] or self.lines  # титульный лист стоит по центру
        self.left = common([l["x"] for l in inner], 0.5)
        self.right = common([l["x2"] for l in inner if l["x2"] > self.left + 100], 2.0, self.left + 400)
        self.gap = median([g for g in self._gaps()], 20.7)
        self.paragraphs = self._paragraphs()
        self.margins = self._margins()

    def _read_page(self, page, number):
        pieces = []

        def visit(text, _cm, tm, font, size):
            if text.strip():
                pieces.append({"text": text, "x": tm[4], "y": tm[5], "size": round(size, 1),
                               "font": (font or {}).get("/BaseFont", "") or ""})
        page.extract_text(visitor_text=visit)
        rows = {}
        for piece in pieces:
            rows.setdefault(round(piece["y"], 0), []).append(piece)
        for y, row in sorted(rows.items(), reverse=True):
            row.sort(key=lambda p: p["x"])
            text = "".join(p["text"] for p in row).strip()
            if not text:
                continue
            biggest = max(row, key=lambda p: len(p["text"].strip()))
            font = biggest["font"].split("+")[-1].lstrip("/")
            line = {"page": number, "y": y, "x": round(row[0]["x"], 1),
                    "x2": round(row[-1]["x"] + text_width(row[-1]["text"], row[-1]["size"]), 1),
                    "size": biggest["size"], "font": font, "bold": "bold" in font.lower(), "text": text}
            (self.footers if y < FOOTER else self.lines).append(line)

    def _gaps(self):
        for a, b in zip(self.lines, self.lines[1:]):
            if a["page"] == b["page"] and a["size"] == b["size"] and 0 < a["y"] - b["y"] < 40:
                yield a["y"] - b["y"]

    def _paragraphs(self):
        """Новый абзац: другая страница, другой кегль, увеличенный разрыв или отступ первой строки."""
        out, current = [], []
        for line in self.lines:
            previous = current[-1] if current else None
            # строки одного абзаца стоят вплотную: разрыв больше полуторного это уже новый абзац
            inside = max(self.gap, line["size"] * SINGLE * 1.5) * 1.2
            split = previous is None or line["page"] != previous["page"] \
                or line["size"] != previous["size"] or line["bold"] != previous["bold"] \
                or previous["y"] - line["y"] > inside \
                or (line["x"] > self.left + 3 and previous["x2"] > self.right - EDGE)
            if split and current:
                out.append(self._paragraph(current))
                current = []
            current.append(line)
        if current:
            out.append(self._paragraph(current))
        return mark_blocks(out)

    def _paragraph(self, lines):
        first, size = lines[0], lines[0]["size"]
        text = " ".join(l["text"] for l in lines)
        level = 0
        # «ПРИЛОЖЕНИЕ 1» это подпись приложения, а не заголовок, и стоит она справа
        if first["bold"] and len(text) < 200 and first["page"] > 1 \
                and not text.strip().lower().startswith(APPENDIX):
            level = 1 if size >= 15.5 else 2 if size >= 13.5 else 3 if size >= 11.5 else 0
        return {
            "text": text, "size": size, "font": first["font"], "bold": first["bold"],
            "jc": self._align(lines), "level": level, "in_table": False,
            # список, цитата или ячейка таблицы: у них своё оформление, под правила основного текста не идут.
            # Абзацный отступ это сдвиг примерно в сантиметр, остальные сдвиги значат именно блок.
            "block": bool(LIST_ITEM.match(text)) or not (
                first["x"] <= self.left + 3 or INDENT[0] <= first["x"] - self.left <= INDENT[1]),
            "line": round(self._line_gap(lines) / (size * SINGLE) * 240) if len(lines) > 1 else 360,
            # висячий отступ списка, сдвинутый блок и продолжение абзаца с прошлой страницы
            # (оно начинается со строчной буквы) это не абзацный отступ
            "first_line": round(first["x"] - self.left)
                          if first["x"] > self.left + 3 and not LIST_ITEM.match(text)
                          and not text[:1].islower()
                          and (len(lines) == 1 or lines[1]["x"] >= first["x"] - 1) else None,
            "page_break": first["y"] > 700 and first["page"] > 1,  # абзац начинает страницу
            "page": first["page"],
        }

    def _line_gap(self, lines):
        return median([a["y"] - b["y"] for a, b in zip(lines, lines[1:])], self.gap)

    def _align(self, lines):
        """Ровный правый край у всех строк кроме последней это выравнивание по ширине."""
        middle = (self.left + self.right) / 2
        if len(lines) > 1:
            full = [l for l in lines[:-1] if abs(l["x2"] - self.right) < EDGE]
            if len(full) == len(lines) - 1:
                return "both"
        for line in lines[:1]:
            if abs((line["x"] + line["x2"]) / 2 - middle) < 12 and line["x"] > self.left + 12:
                return "center"
            if line["x2"] > self.right - EDGE and line["x"] > middle:
                return "right"
        return "left"

    def _margins(self):
        width, height = self.page_size
        top = max((l["y"] + l["size"] for l in self.lines), default=0)
        bottom = min((l["y"] for l in self.lines), default=0)
        # правого поля здесь нет: конец строки в PDF приходится оценивать, ошибка больше допуска
        return {"left": round(self.left / PT_CM, 2),
                "top": round((height - top) / PT_CM, 2), "bottom": round(bottom / PT_CM, 2)}

    def cover(self):
        return [p for p in self.paragraphs if p["page"] == 1]

    def page_numbering(self):
        """Номер страницы это число в колонтитуле. Выравнивание считаем по колонке текста."""
        middle = (self.left + self.right) / 2
        numbers = [l for l in self.footers if l["text"].strip().isdigit()]
        if not numbers:
            return None
        line = numbers[0]
        if abs((line["x"] + line["x2"]) / 2 - middle) < 20:
            return "center"
        return "left" if line["x"] < middle else "right"


def read(data):
    return Doc(data)


if __name__ == "__main__":
    import io, zlib

    def pdf(pages):
        """Самый простой PDF: страницы A4 с кусками текста (x, y, кегль, жирный, строка)."""
        objects, kids = {}, []
        for i, items in enumerate(pages):
            stream = "BT\n" + "".join(
                f"/F{2 if bold else 1} {size} Tf 1 0 0 1 {x} {y} Tm ({text}) Tj\n"
                for x, y, size, bold, text in items) + "ET"
            data = zlib.compress(stream.encode("cp1251"))
            objects[10 + i * 2] = (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.3 841.9] /Contents "
                                   + f"{11 + i * 2} 0 R /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> >>".encode())
            objects[11 + i * 2] = f"<< /Length {len(data)} /Filter /FlateDecode >>\nstream\n".encode() + data + b"\nendstream"
            kids.append(f"{10 + i * 2} 0 R")
        objects[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
        objects[2] = f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(pages)} >>".encode()
        objects[3] = b"<< /Type /Font /Subtype /Type1 /BaseFont /TimesNewRomanPSMT >>"
        objects[4] = b"<< /Type /Font /Subtype /Type1 /BaseFont /TimesNewRomanPS-BoldMT >>"
        out, offsets = bytearray(b"%PDF-1.4\n"), {}
        for number in sorted(objects):
            offsets[number] = len(out)
            out += f"{number} 0 obj\n".encode() + objects[number] + b"\nendobj\n"
        start = len(out)
        last = max(objects)
        out += f"xref\n0 {last + 1}\n0000000000 65535 f \n".encode()
        for number in range(1, last + 1):
            out += (f"{offsets[number]:010d} 00000 n \n".encode() if number in offsets else b"0000000000 65535 f \n")
        out += f"trailer\n<< /Size {last + 1} /Root 1 0 R >>\nstartxref\n{start}\n%%EOF".encode()
        return io.BytesIO(bytes(out))

    left = 85.2
    wide = "x" * 75              # строка во всю колонку: правый край на 2 см от края листа
    d = read(pdf([
        [(200, 600, 20, True, "Tema raboty"), (200, 500, 14, False, "Gimnaziya")],
        [(left, 770, 16, True, "Vvedenie"),
         (left, 740, 12, False, wide), (left, 719.3, 12, False, wide), (left, 698.6, 12, False, "konec abzaca"),
         (left, 670, 12, False, "Novyi abzac so ssylkoi (Ivanov, 2021)"),
         (left, 100, 16, True, "Kasutatud allikad"),
         (left, 57, 12, False, "Ivanov I. Energiya zdanii. 2021. https://err.ee/uurimus"),
         (300, 45, 12, False, "2")],
    ]))
    assert round(d.margins["left"]) == 3 and "right" not in d.margins, d.margins
    assert round(d.margins["top"]) == 2 and round(d.margins["bottom"]) == 2, d.margins
    assert [p["text"][:9] for p in d.cover()] == ["Tema rabo", "Gimnaziya"], d.cover()
    assert [h["text"] for h in d.headings()] == ["Vvedenie", "Kasutatud allikad"], d.headings()
    assert all(h["level"] == 1 and h["size"] == 16 for h in d.headings())
    first = d.body()[1]
    assert first["font"] == "TimesNewRomanPSMT" and first["size"] == 12, first
    assert abs(first["line"] - 360) < 15, first["line"]   # 20,7 пт при кегле 12 это интервал 1,5
    assert d.page_numbering() == "center", d.page_numbering()
    assert len(d.citations()) == 1 and d.sources() == ["Ivanov I. Energiya zdanii. 2021. https://err.ee/uurimus"], d.sources()
    assert d.cover()[0]["size"] == 20 and d.cover()[1]["size"] == 14
    print("ok")
