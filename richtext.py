"""Оформленный текст ученика: чистка HTML из редактора и перевод его в обычный текст."""
import re
from html import escape
from html.parser import HTMLParser

ALLOWED = {"p", "br", "h2", "h3", "h4", "b", "i", "u", "s", "ul", "ol", "li", "blockquote", "sup", "sub", "a",
           "figure", "figcaption"}
SAME = {"div": "p", "h1": "h2", "h5": "h4", "h6": "h4", "strong": "b", "em": "i",
        "strike": "s", "del": "s", "ins": "u", "mark": "u", "pre": "p", "section": "p", "article": "p"}
BLOCKS = {"p", "h2", "h3", "h4", "li", "blockquote", "ul", "ol", "br", "table", "tr", "figure", "figcaption"}
IMAGE_SRC = re.compile(r"/images/\d+")
# Заголовки самостоятельных частей работы оформляются как раздел, но без номера.
UNNUMBERED = ("введен", "вступ", "sissejuhatus", "заключ", "вывод", "висновк", "kokkuvõte", "список",
              "kasutatud", "allika", "приложен", "додат", "lisa", "аннотац", "анотац", "resümee", "summary",
              "содержан", "зміст", "sisukord")
HEADINGS = re.compile(r'<h([234])(?: class="(\w+)")?>(.*?)</h\1>', re.S)
SKIP_CONTENT = {"script", "style", "head", "title"}
MAX_HTML = 300_000


class Cleaner(HTMLParser):
    """Оставляет только простое оформление. Теги не из списка выбрасывает, текст внутри сохраняет."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in SKIP_CONTENT:
            self.skip += 1
            return
        tag = SAME.get(tag, tag)
        if tag == "br":
            self.out.append("<br>")
        elif tag == "img":
            src = dict(attrs).get("src") or ""
            if IMAGE_SRC.fullmatch(src):
                self.out.append(f'<img src="{src}">')
        elif tag == "h2" and dict(attrs).get("class") == "appendix":
            self.out.append('<h2 class="appendix">')
        elif tag == "a":
            href = dict(attrs).get("href") or ""
            ok = re.match(r"https?://", href, re.I)
            self.out.append(f'<a href="{escape(href, quote=True)}" target="_blank" rel="noopener noreferrer">' if ok else "<a>")
        elif tag in ALLOWED:
            self.out.append(f"<{tag}>")

    def handle_endtag(self, tag):
        if tag in SKIP_CONTENT:
            self.skip = max(0, self.skip - 1)
            return
        tag = SAME.get(tag, tag)
        if tag in ALLOWED and tag not in ("br", "img"):
            self.out.append(f"</{tag}>")

    def handle_data(self, data):
        if not self.skip:
            self.out.append(escape(data, quote=False))


class ToText(HTMLParser):
    """Текст с переносами на границах абзацев: так работу читает ИИ."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in SKIP_CONTENT:
            self.skip += 1
        elif tag in BLOCKS:
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in SKIP_CONTENT:
            self.skip = max(0, self.skip - 1)
        elif tag in BLOCKS:
            self.out.append("\n")

    def handle_data(self, data):
        if not self.skip:
            self.out.append(data)


def run(parser, html):
    parser.feed(html or "")
    parser.close()
    return "".join(parser.out)


def mark_headings(html):
    """Класс nonum получают введение, выводы, список источников и подразделы под ними. Как в editor.js."""
    numbered = False

    def one(m):
        nonlocal numbered
        level, cls, inner = m[1], m[2], m[3]
        if level == "2":
            if cls == "appendix":
                numbered = False
                return m[0]
            numbered = not re.sub(r"<[^>]+>", "", inner).strip().lower().startswith(UNNUMBERED)
        return f"<h{level}>{inner}</h{level}>" if numbered else f'<h{level} class="nonum">{inner}</h{level}>'
    return HEADINGS.sub(one, html)


def sanitize(html):
    return mark_headings(run(Cleaner(), (html or "")[:MAX_HTML]))


def plain(html):
    """Текст ровно так, как его видит браузер в textContent: по нему считаются позиции пометок."""
    return run(ToText(), html).replace("\n", "")


def readable(html):
    text = re.sub(r"\n{3,}", "\n\n", run(ToText(), html))
    return "\n".join(line.strip() for line in text.split("\n")).strip()


if __name__ == "__main__":
    dirty = ('<div style="x">Привет<script>alert(1)</script> <B>мир</B>!</div>'
             '<p>Ссылка <a href="javascript:alert(1)">злая</a> и <a href="https://err.ee/uurimus">добрая</a></p>'
             '<img src=x onerror=alert(1)><ul><li>раз</li><li>два</li></ul>')
    clean = sanitize(dirty)
    assert "<script" not in clean and "onerror" not in clean and "style" not in clean, clean
    assert clean.count("<a>") == 1 and 'href="https://err.ee/uurimus"' in clean, clean
    assert clean.startswith("<p>Привет <b>мир</b>!</p>"), clean
    assert plain(clean) == "Привет мир!Ссылка злая и добраяраздва", plain(clean)
    assert readable(clean) == "Привет мир!\n\nСсылка злая и добрая\n\nраз\n\nдва", repr(readable(clean))
    assert plain(sanitize("<p>a&amp;b</p>")) == "a&b"
    assert sanitize("<p>1 < 2 и 3 > 2</p>") == "<p>1 &lt; 2 и 3 &gt; 2</p>", sanitize("<p>1 < 2 и 3 > 2</p>")
    assert plain(sanitize("<p>раз</p><p>два</p>")) == "раздва"  # как textContent в браузере
    marked = sanitize('<h2 class="nonum">Методика</h2><h2 class="x">Введение</h2><h2 class="appendix">Анкета</h2>'
                      '<h4>Пункт</h4><h5>мелко</h5><figure><img src="/images/7" onerror="x"><figcaption>Фото</figcaption></figure>'
                      '<img src="https://evil.example/x.png"><img src="/images/7?x">')
    assert marked == ('<h2>Методика</h2><h2 class="nonum">Введение</h2><h2 class="appendix">Анкета</h2>'
                      '<h4 class="nonum">Пункт</h4><h4 class="nonum">мелко</h4><figure><img src="/images/7"><figcaption>Фото</figcaption></figure>'), marked
    assert sanitize("<h3>до разделов</h3><h2>Теория</h2><h3>a</h3><h2>Выводы</h2><h3>b</h3>") == (
        '<h3 class="nonum">до разделов</h3><h2>Теория</h2><h3>a</h3><h2 class="nonum">Выводы</h2><h3 class="nonum">b</h3>')
    print("ok")
