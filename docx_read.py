"""Чтение .docx: текст работы, её структура и оформление каждого абзаца.

Word хранит оформление тремя слоями: умолчания документа, стиль абзаца (со своим родителем)
и прямые настройки в самом абзаце. Здесь они собираются в один словарь на абзац, по нему
потом проверяются правила оформления.
"""
import re, zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
TWIP_CM = 567.0          # твипов в сантиметре
MAX_ZIP = 60 * 1024 * 1024   # распакованный .docx больше этого не читаем
CAPTION = ("таблица", "рисунок", "tabel", "joonis", "фото", "схема", "диаграмма")  # подписи стоят по центру
CITATION = re.compile(r"\([^()]{2,90}?(?:\d{4}|lk\.?\s*\d|с\.\s*\d)[^()]{0,25}\)|\[\d{1,3}(?:[,;]\s*[^\]]{0,20})?\]")
URL = re.compile(r"https?://[^\s,;)\]]+|www\.[^\s,;)\]]+")
SOURCE_HEADS = ("список", "kasutatud", "allika", "літератур", "литератур", "kirjandus",
                "references", "використ", "джерел")


def attr(el, name, default=None):
    return el.get(W + name, default) if el is not None else default


def num(value, default=None):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def para_props(ppr):
    """Настройки абзаца: выравнивание, интервалы, отступ, разрыв страницы."""
    if ppr is None:
        return {}
    spacing, ind = ppr.find(W + "spacing"), ppr.find(W + "ind")
    out = {
        "jc": attr(ppr.find(W + "jc"), "val"),
        "line": num(attr(spacing, "line")),
        "line_rule": attr(spacing, "lineRule"),
        "before": num(attr(spacing, "before")),
        "after": num(attr(spacing, "after")),
        "first_line": num(attr(ind, "firstLine")),  # висячий отступ списка это не абзацный отступ
        "left": num(attr(ind, "left")),
        "page_break": ppr.find(W + "pageBreakBefore") is not None
                      and attr(ppr.find(W + "pageBreakBefore"), "val", "1") not in ("0", "false"),
        "outline": num(attr(ppr.find(W + "outlineLvl"), "val")),
        "style": attr(ppr.find(W + "pStyle"), "val"),
    }
    return {k: v for k, v in out.items() if v is not None and v is not False}


def run_props(rpr):
    """Настройки букв: шрифт, кегль в пунктах, полужирность."""
    if rpr is None:
        return {}
    fonts, sz, b = rpr.find(W + "rFonts"), rpr.find(W + "sz"), rpr.find(W + "b")
    out = {
        "font": attr(fonts, "ascii") or attr(fonts, "hAnsi") or attr(fonts, "cs"),
        "size": (num(attr(sz, "val")) or 0) / 2 or None,
        "bold": None if b is None else attr(b, "val", "1") not in ("0", "false"),
    }
    return {k: v for k, v in out.items() if v is not None}


class Doc:
    """Разобранный .docx. paragraphs: текст и оформление, margins: поля страницы в сантиметрах."""

    def __init__(self, data):
        with zipfile.ZipFile(data) as z:
            if sum(i.file_size for i in z.infolist()) > MAX_ZIP:
                raise ValueError("файл слишком большой")
            part = lambda name: ET.fromstring(z.read(name)) if name in z.namelist() else None
            doc = part("word/document.xml")
            if doc is None:
                raise ValueError("это не документ Word")
            self.styles_xml, settings = part("word/styles.xml"), part("word/settings.xml")
            rels = part("word/_rels/document.xml.rels")
            self.footers = [part(f"word/{f}") for f in sorted(
                n[5:] for n in z.namelist() if re.fullmatch(r"word/footer\d+\.xml", n))]
        self.defaults = self._defaults()
        self.styles = {attr(s, "styleId"): s for s in self.styles_xml.iter(W + "style")} if self.styles_xml is not None else {}
        self.resolved = {}
        in_table = {id(p) for t in doc.iter(W + "tbl") for p in t.iter(W + "p")}
        self.paragraphs = [self._paragraph(p, id(p) in in_table) for p in doc.find(W + "body").iter(W + "p")]
        self.margins = self._margins(doc)
        self.updates_fields = settings is not None and settings.find(W + "updateFields") is not None
        self.rel_targets = {attr(r, "Id"): r.get("Target") for r in rels} if rels is not None else {}

    def _defaults(self):
        ppr = rpr = None
        if self.styles_xml is not None:
            d = self.styles_xml.find(W + "docDefaults")
            if d is not None:
                pd, rd = d.find(W + "pPrDefault"), d.find(W + "rPrDefault")
                ppr, rpr = (pd.find(W + "pPr") if pd is not None else None), (rd.find(W + "rPr") if rd is not None else None)
        return {**para_props(ppr), **run_props(rpr)}

    def _style(self, style_id, seen=()):
        """Настройки стиля вместе с унаследованными от родителя."""
        if style_id in self.resolved:
            return self.resolved[style_id]
        s = self.styles.get(style_id)
        if s is None or style_id in seen:
            return {}
        parent = self._style(attr(s.find(W + "basedOn"), "val"), seen + (style_id,))
        own = {**para_props(s.find(W + "pPr")), **run_props(s.find(W + "rPr")),
               "style_name": (attr(s.find(W + "name"), "val") or "").lower()}
        out = {**parent, **own}
        self.resolved[style_id] = out
        return out

    def _paragraph(self, p, in_table=False):
        ppr = p.find(W + "pPr")
        direct = para_props(ppr)
        style = self._style(direct.get("style"))
        runs = [(("".join(t.text or "" for t in r.iter(W + "t"))),
                 {**run_props(ppr.find(W + "rPr") if ppr is not None else None), **run_props(r.find(W + "rPr"))})
                for r in p.findall(W + "r")]
        text = "".join(t for t, _ in runs)
        sized = [(len(t.strip()), rp) for t, rp in runs if t.strip()]
        main = max(sized, key=lambda x: x[0])[1] if sized else {}
        out = {**self.defaults, **style, **direct, **{k: v for k, v in main.items()},
               "text": text.strip(), "runs": runs, "in_table": in_table,
               "page_break": direct.get("page_break", style.get("page_break", False))
                             or any(attr(br, "type") == "page" for br in p.iter(W + "br"))}
        name = out.get("style_name", "")
        level = out.get("outline")
        if re.fullmatch(r"heading [1-9]", name):
            level = int(name[-1]) - 1
        out["level"] = int(level) + 1 if level is not None and level < 9 and out["text"] else 0
        return out

    def _margins(self, doc):
        sect = next((s for s in doc.iter(W + "sectPr")), None)
        m = sect.find(W + "pgMar") if sect is not None else None
        return {side: round(num(attr(m, side), 0) / TWIP_CM, 2) for side in ("top", "right", "bottom", "left")}

    # ---------- то, что нужно проверкам ----------

    def body(self):
        """Текст работы: от первого заголовка раздела. До него идут титул, декларация и содержание."""
        first = next((i for i, p in enumerate(self.paragraphs) if p["level"] == 1), None)
        return self.paragraphs[first:] if first is not None else self.paragraphs[len(self.cover()):]

    def cover(self):
        """Титульный лист: всё до первого разрыва страницы."""
        rest = self.paragraphs[1:]
        first = next((i for i, p in enumerate(rest) if p["page_break"]), None)
        return self.paragraphs[:first + 1] if first is not None else []

    def text(self):
        return "\n".join(p["text"] for p in self.paragraphs if p["text"])

    def headings(self):
        return [p for p in self.paragraphs if p["level"]]

    def page_numbering(self):
        """Есть ли в колонтитуле поле PAGE и как оно выровнено."""
        for f in self.footers:
            if f is None:
                continue
            marks = "".join(t.text or "" for t in f.iter(W + "instrText"))
            if "PAGE" in marks.upper():
                p = next((p for p in f.iter(W + "p") if "PAGE" in "".join(t.text or "" for t in p.iter(W + "instrText")).upper()), None)
                jc = attr(p.find(W + "pPr/" + W + "jc") if p is not None else None, "val")
                return jc or "left"
        return None

    def sources(self):
        """Строки из раздела «Список использованных источников»."""
        out, inside = [], False
        for p in self.paragraphs:
            if p["level"]:
                inside = p["text"].strip().lower().startswith(SOURCE_HEADS)
                continue
            if inside and len(p["text"]) > 8:
                out.append(re.sub(r"^\s*\d{1,3}\s*[.)]\s*", "", p["text"]))  # свой номер списка не дублируем
        return out

    def citations(self):
        return CITATION.findall(self.text())


def read(data):
    return Doc(data)


if __name__ == "__main__":
    import io, zipfile as zf

    def make(body, styles=""):
        buf = io.BytesIO()
        with zf.ZipFile(buf, "w") as z:
            z.writestr("word/document.xml", f'<w:document xmlns:w="{W[1:-1]}"><w:body>{body}</w:body></w:document>')
            z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W[1:-1]}">{styles}</w:styles>')
        return io.BytesIO(buf.getvalue())

    styles = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:pageBreakBefore/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>')
    d = read(make(
        '<w:p><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>Тема работы</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Введение</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Обычный текст (Иванов, 2021) и ещё [3].</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Список использованных источников</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>1. Иванов И. Энергия зданий. 2021. https://err.ee/uurimus</w:t></w:r></w:p>'
        '<w:sectPr><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701"/></w:sectPr>', styles))
    assert d.margins == {"top": 2.0, "right": 2.0, "bottom": 2.0, "left": 3.0}, d.margins
    assert [p["text"] for p in d.cover()] == ["Тема работы"], d.cover()
    assert [p["text"][:20] for p in d.body()][:2] == ["Введение", "Обычный текст (Ивано"], d.body()
    assert [h["text"] for h in d.headings()] == ["Введение", "Список использованных источников"]
    assert all(h["level"] == 1 and h["bold"] and h["size"] == 16 and h["page_break"] for h in d.headings())
    body = d.body()[1]
    assert body["font"] == "Times New Roman" and body["size"] == 12 and body["jc"] == "both" and body["line"] == 360
    assert d.cover()[0]["size"] == 20, d.cover()[0]
    assert len(d.citations()) == 2, d.citations()
    assert d.sources() == ["Иванов И. Энергия зданий. 2021. https://err.ee/uurimus"], d.sources()
    assert d.page_numbering() is None
    assert not any(p["in_table"] for p in d.paragraphs)
    print("ok")
