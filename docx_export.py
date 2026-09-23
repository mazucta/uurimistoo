"""Работа в Word по школьным правилам оформления.

Поля 3/2/2/2 см, Times New Roman 12, интервал 1,5, по ширине, 6 пт до и после абзаца, без отступа.
Заголовок 1: 16 пт, полужирный, с новой страницы. Заголовок 2: 14 пт. Заголовок 3: 12 пт.
Разделы основной части нумеруются, введение, выводы, список источников и приложения нет.
Титул, декларация и содержание в первой секции без номеров страниц, дальше номер по центру снизу.
"""
import io, re, struct, zipfile
from html import escape, unescape
from html.parser import HTMLParser

CM = 567            # twips в сантиметре
EMU_CM = 360000     # EMU в сантиметре
EMU_PX = 9525       # EMU в пикселе при 96 dpi
MAX_W, MAX_H = 16 * EMU_CM, 22 * EMU_CM
SOURCES_KEYS = ("список", "kasutatud", "allikad", "використ")  # ученик уже написал список сам

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {NS}>
<w:docDefaults>
 <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>
  <w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault>
 <w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="left"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9061"/></w:tabs><w:spacing w:before="0" w:after="0"/><w:jc w:val="left"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="TOC1"/><w:pPr><w:ind w:left="284"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/><w:basedOn w:val="TOC1"/><w:pPr><w:ind w:left="567"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>"""

FOOTER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr {NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>"""

# Word при открытии предложит обновить поля: так в содержании появятся номера страниц.
SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {NS}><w:updateFields w:val="true"/><w:defaultTabStop w:val="708"/></w:settings>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def image_size(data):
    """Ширина и высота PNG или JPEG в пикселях, без Pillow."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", data[16:24])
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker, length = data[i + 1], struct.unpack(">H", data[i + 2:i + 4])[0]
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return w, h
        i += 2 + length
    return 800, 600


# Порядок важен: Word не откроет файл, если свойства идут не по схеме.
RUN_PROPS = (("b", "<w:b/>"), ("i", "<w:i/>"), ("s", "<w:strike/>"), ("u", '<w:u w:val="single"/>'),
             ("sup", '<w:vertAlign w:val="superscript"/>'), ("sub", '<w:vertAlign w:val="subscript"/>'))


def run_xml(text, fmt, link=None):
    props = "".join(xml for f, xml in RUN_PROPS if f in fmt)
    if link:
        props = '<w:rStyle w:val="Hyperlink"/>' + props
    parts = text.split("\n")
    body = "<w:br/>".join(f'<w:t xml:space="preserve">{escape(p, quote=False)}</w:t>' for p in parts)
    return f"<w:r><w:rPr>{props}</w:rPr>{body}</w:r>"


def para(runs, style=None, jc=None, extra=""):
    ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + extra + (f'<w:jc w:val="{jc}"/>' if jc else "")
    return f"<w:p><w:pPr>{ppr}</w:pPr>{''.join(runs)}</w:p>"


def text_para(text, size=28, bold=False, jc="center", extra=""):
    rpr = ("<w:b/>" if bold else "") + f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return para([f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text, quote=False)}</w:t></w:r>'],
                jc=jc, extra=extra)


class Body(HTMLParser):
    """Переводит очищенный HTML редактора в абзацы Word."""

    def __init__(self, image, labels):
        super().__init__(convert_charrefs=True)
        self.image, self.labels = image, labels
        self.out, self.runs, self.block = [], [], None
        self.fmt, self.link, self.lists, self.quote = set(), None, [], 0
        self.num, self.appendix = [0, 0, 0], 0
        self.rels, self.media, self.headings = [], [], []
        self.first_appendix = None  # сюда встаёт список источников: после выводов, перед приложениями
        self.rows, self.body_out = None, None  # таблица: строки ячеек и абзацы вне таблицы

    def rel(self, kind, target, external=False):
        rid = f"rId{len(self.rels) + 10}"
        self.rels.append(f'<Relationship Id="{rid}" Type="{REL}/{kind}" Target="{escape(target)}"'
                         + (' TargetMode="External"/>' if external else "/>"))
        return rid

    def open(self, **block):
        self.close_block()
        self.block, self.runs = block, []

    def close_block(self):
        if self.block is None:
            return
        b, self.block = self.block, None
        if b.get("level"):
            text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", "".join(self.runs)))
            self.heading(b, text)
            return
        if not self.runs and not b.get("prefix"):
            return
        indent = ""
        if b.get("cell"):
            b["jc"] = "center" if self.rows and self.rows[-1] else "left"  # первый столбец это подписи строк
        elif b.get("li"):
            depth = len(self.lists)
            indent = f'<w:ind w:left="{360 * depth + 360}" w:hanging="360"/>'
        elif self.quote:
            indent = '<w:ind w:left="567"/>'
        self.out.append(para(b.get("prefix", []) + self.runs, b.get("style"), b.get("jc"), indent))

    def heading(self, b, text):
        text = unescape(text).strip()
        level, cls = b["level"], b.get("cls")
        prefix = ""
        if cls == "appendix":
            if self.first_appendix is None:
                self.first_appendix = (len(self.out), len(self.headings))
            self.appendix += 1
            self.out.append(text_para(f"{self.labels['appendix']} {self.appendix}", 24, True, "right",
                                      "<w:keepNext/><w:pageBreakBefore/>"))
            extra = '<w:pageBreakBefore w:val="0"/>'
        else:
            extra = ""
            if cls != "nonum":
                self.num[level - 1] += 1
                self.num[level:] = [0] * (3 - level)
                prefix = ".".join(str(n) for n in self.num[:level]) + " "
        self.headings.append((level, prefix + text))
        self.out.append(para([run_xml(prefix + text, set())], f"Heading{level}", extra=extra))

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("p", "figcaption"):
            if self.block is None or not self.block.get("li"):
                self.open(style="Caption" if tag == "figcaption" else None)
        elif tag in ("h2", "h3", "h4"):
            self.open(level=int(tag[1]) - 1, cls=a.get("class"))
        elif tag in ("ul", "ol"):
            self.close_block()
            self.lists.append([tag, 0])
        elif tag == "li":
            kind = self.lists[-1] if self.lists else ["ul", 0]
            kind[1] += 1
            mark = f"{kind[1]}." if kind[0] == "ol" else "•"
            self.open(li=True, prefix=[run_xml(mark + "\t", set())])
        elif tag == "blockquote":
            self.close_block()
            self.quote += 1
            self.open()
        elif tag == "table":
            self.close_block()
            self.rows, self.body_out = [], self.out
        elif tag == "tr" and self.rows is not None:
            self.rows.append([])
        elif tag in ("td", "th") and self.rows:
            self.close_block()
            self.out = []
            if tag == "th":
                self.fmt.add("b")
            self.open(cell=True)
        elif tag == "br":
            self.data("\n")
        elif tag == "img":
            self.picture(a.get("src", ""))
        elif tag == "a":
            self.link = a.get("href")
        elif tag in ("b", "i", "u", "s", "sup", "sub"):
            self.fmt.add(tag)

    def handle_endtag(self, tag):
        if tag in ("p", "figcaption", "h2", "h3", "h4", "li"):
            self.close_block()
        elif tag in ("ul", "ol"):
            self.close_block()
            if self.lists:
                self.lists.pop()
        elif tag == "blockquote":
            self.close_block()
            self.quote = max(0, self.quote - 1)
        elif tag == "a":
            self.link = None
        elif tag in ("td", "th") and self.rows:
            self.close_block()
            self.fmt.discard("b")
            self.rows[-1].append(self.out or ["<w:p/>"])
            self.out = []
        elif tag == "table" and self.rows is not None:
            self.close_block()
            self.out = self.body_out
            self.out.append(table_xml(self.rows))
            self.out.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>')  # Word не терпит таблицу вплотную к таблице
            self.rows = None
        elif tag in self.fmt:
            self.fmt.discard(tag)

    def handle_data(self, data):
        self.data(data)

    def data(self, text):
        if self.block is None:
            if not text.strip():
                return
            self.open()
        r = run_xml(text, self.fmt, self.link)
        if self.link:
            rid = self.rel("hyperlink", self.link, external=True)
            r = f'<w:hyperlink r:id="{rid}">{r}</w:hyperlink>'
        self.runs.append(r)

    def picture(self, src):
        m = re.fullmatch(r"/images/(\d+)", src)
        found = m and self.image(int(m[1]))
        if not found:
            return
        mime, data = found
        ext = "png" if mime == "image/png" else "jpeg"
        name = f"media/image{len(self.media) + 1}.{ext}"
        self.media.append((name, data))
        rid = self.rel("image", name)
        w, h = image_size(data)
        cx, cy = w * EMU_PX, h * EMU_PX
        k = min(1, MAX_W / cx, MAX_H / cy)
        cx, cy, n = int(cx * k), int(cy * k), len(self.media)
        drawing = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                   f'<wp:docPr id="{n}" name="Picture {n}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="image{n}"/><pic:cNvPicPr/></pic:nvPicPr>'
                   f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                   f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
                   f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
        self.close_block()
        self.out.append(para([drawing], jc="center", extra="<w:keepNext/>"))


def table_xml(rows):
    """Таблица во всю ширину текста, с рамкой. Первый столбец (подписи строк) вдвое шире, шапка повторяется на каждой странице."""
    cols = max((len(r) for r in rows if r), default=1)
    unit = 16 * CM // (cols + 1)
    widths = [unit * 2] + [unit] * (cols - 1)
    border = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="{16 * CM}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>{border}</w:tblBorders></w:tblPr>'
           f'<w:tblGrid>{grid}</w:tblGrid>']
    for i, row in enumerate(r for r in rows if r):
        cells = row + [["<w:p/>"]] * (cols - len(row))
        head = "<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>" if i == 0 else "<w:trPr><w:cantSplit/></w:trPr>"
        out.append("<w:tr>" + head + "".join(
            f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{"".join(c)}</w:tc>' for w, c in zip(widths, cells)) + "</w:tr>")
    return "".join(out) + "</w:tbl>"


def sources_block(sources, labels):
    """Список источников по алфавиту, если ученик не написал его сам."""
    rows = sorted(sources, key=lambda s: (s["author"] or s["title"]).lower())
    out = [para([run_xml(labels["sources"], set())], "Heading1")]
    for i, s in enumerate(rows, 1):
        line = ". ".join(x.rstrip(". ") for x in (s["author"], s["title"], s["year"]) if x) + "."
        if s["url"]:
            line += f" {s['url']}"
        out.append(para([run_xml(f"{i}.\t{line}", set())], jc="left", extra='<w:ind w:left="360" w:hanging="360"/>'))
    return out


def build(work, labels, sources, image):
    """work: title, html, student, teacher, school, grade, city, year, date. labels: подписи на языке работы."""
    body = Body(image, labels)
    body.feed(work["html"] or "")
    body.close()
    body.close_block()
    content = body.out
    if sources and not any(t.lower().startswith(SOURCES_KEYS) for _, t in body.headings):
        at, at_heading = body.first_appendix or (len(content), len(body.headings))
        content[at:at] = sources_block(sources, labels)
        body.headings.insert(at_heading, (1, labels["sources"]))

    gap = text_para("", 28)
    cover = [text_para(work["school"], 28)] + [gap] * 7 + [
        text_para(work["title"] or "…", 40, True),
        text_para(labels["kind"], 28),
    ] + [gap] * 5 + [
        text_para(f"{labels['author']}: {work['student']}" + (f", {work['grade']}" if work["grade"] else ""), 28, jc="right"),
        text_para(f"{labels['teacher']}: {work['teacher']}", 28, jc="right"),
    ] + [gap] * 5 + [text_para(", ".join(x for x in (work["city"], work["year"]) if x), 28)]

    declaration = [
        text_para(labels["declaration"], 32, True, "left", "<w:pageBreakBefore/>"),
        para([run_xml(labels["declaration_text"], set())]),
        para([run_xml(f"{work['student']}\t\t{work['date']}\t\t____________", set())], jc="left"),
    ]

    toc_lines = "".join(para([run_xml(t, set())], f"TOC{lvl}") for lvl, t in body.headings)
    contents = [
        text_para(labels["contents"], 32, True, "left", "<w:pageBreakBefore/>"),
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>' + toc_lines +
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    ]
    margins = (f'<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="{2 * CM}" w:right="{2 * CM}" w:bottom="{2 * CM}" '
               f'w:left="{3 * CM}" w:header="708" w:footer="708" w:gutter="0"/>')
    # Первая секция без колонтитула: номера страниц считаются, но не печатаются.
    front_end = f'<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>{margins}</w:sectPr></w:pPr></w:p>'
    footer_rid = body.rel("footer", "footer1.xml")
    main_sect = f'<w:sectPr><w:footerReference w:type="default" r:id="{footer_rid}"/><w:type w:val="nextPage"/>{margins}</w:sectPr>'

    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>'
           + "".join(cover + declaration + contents) + front_end + "".join(content) + main_sect
           + "</w:body></w:document>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>'
            + "".join(body.rels) + "</Relationships>")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footer1.xml", FOOTER)
        for name, data in body.media:
            z.writestr("word/" + name, data)
    return buf.getvalue()
