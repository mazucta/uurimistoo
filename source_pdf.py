"""Источники из списка и найденные в сети работы: программа скачивает их сама.
PDF отдаётся модели текстом, чтобы сверка цитат и цифр шла по самому источнику; текст любых скачанных
страниц потом сравнивается с работой кодом — так видно, что переписано слово в слово.

Адреса пишет ученик, поэтому скачивание осторожное: только http(s), только публичные адреса в интернете
(ни этот компьютер, ни локальная сеть, ни служебные адреса облака), каждый переход по редиректу проверяется
заново, размер и время ограничены, берётся только PDF или страница HTML.
"""
import html.parser
import io
import ipaddress
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request

from pypdf import PdfReader

MAX_BYTES = 20 * 1024 * 1024   # учебник или статья целиком, но не архив на гигабайт
TIMEOUT = 15                   # секунд на один файл
BUDGET = 60                    # секунд на все PDF одной работы
MAX_PDFS = 6
MAX_SOURCES = 12               # адресов из списка за одну проверку
HTML_MAX = 3 * 1024 * 1024
TEXT_MAX = 300_000             # текста страницы для сравнения с работой
MAX_PAGES = 80
MAX_CHARS = 40_000             # текста одного источника в модель: этого хватает для сверки
# как у обычного браузера: часть сайтов отвечает 403 на незнакомую подпись, хотя у ученика страница открывается
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/126.0 Safari/537.36",
           "Accept": "text/html,application/xhtml+xml,application/pdf,*/*;q=0.8",
           "Accept-Language": "et,uk,ru;q=0.9,en;q=0.8"}
URL = re.compile(r"https?://[^\s,;)\]»”\"]+")


def public(host):
    """Адрес хоста только публичный: ни один из его IP не ведёт внутрь компьютера или сети."""
    try:
        infos = socket.getaddrinfo(host, None)
    except (OSError, UnicodeError):
        return False
    return bool(infos) and all(ipaddress.ip_address(i[4][0].split("%")[0]).is_global for i in infos)


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):  # редирект разбираем сами, чтобы проверить новый адрес
        return None


def fetch(url):
    """('pdf' или 'html', байты, кодировка) по адресу или None: не PDF и не HTML, закрытый адрес, слишком большой файл."""
    # ponytail: между проверкой адреса и запросом DNS может смениться (rebinding); для пилота хватает,
    # на сервере для настоящих учеников — скачивать через прокси, который сам не пускает во внутреннюю сеть
    opener = urllib.request.build_opener(_NoRedirect)
    for _hop in range(4):
        parts = urllib.parse.urlsplit(url)
        if parts.scheme not in ("http", "https") or not parts.hostname or not public(parts.hostname):
            return None
        request = urllib.request.Request(url, headers=HEADERS)
        try:
            response = opener.open(request, timeout=TIMEOUT)
        except urllib.error.HTTPError as ex:
            if ex.code in (301, 302, 303, 307, 308) and ex.headers.get("Location"):
                url = urllib.parse.urljoin(url, ex.headers["Location"])
                continue
            return None
        except (OSError, ValueError):
            return None
        with response:
            head = response.read(5)
            if head == b"%PDF-":
                kind, limit = "pdf", MAX_BYTES
            elif response.headers.get_content_type() in ("text/html", "application/xhtml+xml"):
                kind, limit = "html", HTML_MAX
            else:
                return None
            try:
                data = head + response.read(limit)
            except (OSError, ValueError):
                return None
        return (kind, data, response.headers.get_content_charset() or "utf-8") if len(data) <= limit else None
    return None


class _Text(html.parser.HTMLParser):
    SKIP = {"script", "style", "noscript", "nav", "header", "footer", "aside", "form", "svg", "template"}

    def __init__(self):
        super().__init__()
        self.out, self.skip = [], 0

    def handle_starttag(self, tag, attrs):
        self.skip += tag in self.SKIP
        self.out.append(" ")  # иначе слова соседних блоков слипаются

    def handle_endtag(self, tag):
        if tag in self.SKIP and self.skip:
            self.skip -= 1
        self.out.append(" ")

    def handle_data(self, data):
        if not self.skip:
            self.out.append(data)


def html_text(data, charset="utf-8"):
    """Видимый текст страницы: без скриптов, меню, шапки и подвала."""
    parser = _Text()
    try:
        parser.feed(data.decode(charset, errors="replace"))
    except LookupError:  # кодировку из заголовка Python не знает
        parser.feed(data.decode("utf-8", errors="replace"))
    return re.sub(r"\s+", " ", "".join(parser.out)).strip()[:TEXT_MAX]


def page_text(url, fetch_=None):
    """(вид, страниц, текст) по адресу или None. У HTML страниц 0."""
    got = (fetch_ or fetch)(url)
    if not got:
        return None
    kind, data, charset = got
    try:
        return ("pdf", *pdf_text(data, TEXT_MAX)) if kind == "pdf" else ("html", 0, html_text(data, charset))
    except Exception:  # битый PDF: источник просто идёт к модели как обычно
        return None


def pdf_text(data, limit=MAX_CHARS):
    """(страниц, текст) из PDF. У скана текста нет — тогда пустая строка."""
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        reader.decrypt("")
    text = "\n".join((page.extract_text() or "") for page in reader.pages[:MAX_PAGES])
    return len(reader.pages), re.sub(r"[ \t]+", " ", text).strip()[:limit]


def source_texts(sources, fetch_=None):
    """{номер источника: (вид, страниц, текст)} для источников, у которых есть адрес PDF или страницы.
    sources: [(номер, строка списка)]."""
    out, start, tried = {}, time.monotonic(), 0
    for number, entry in sources:
        if tried >= MAX_SOURCES or time.monotonic() - start > BUDGET:
            break
        for url in URL.findall(entry):
            tried += 1
            got = page_text(url.rstrip(".,"), fetch_)
            if got and (got[2] or got[0] == "pdf"):  # скан без текста: источник всё равно существует
                out[number] = got
                break
    return out


def urls_texts(urls, fetch_=None):
    """[(адрес, текст)] для найденных в сети работ: что скачалось за отведённое время."""
    out, start = [], time.monotonic()
    for url in urls:
        if time.monotonic() - start > BUDGET:
            break
        got = page_text(url, fetch_)
        if got and got[2]:
            out.append((url, got[2]))
    return out


if __name__ == "__main__":
    for blocked in ("http://127.0.0.1:8000/x.pdf", "http://localhost/x.pdf", "http://169.254.169.254/latest/x.pdf",
                    "http://10.0.0.5/a.pdf", "http://192.168.1.1/a.pdf", "file:///etc/passwd", "ftp://example.com/a.pdf",
                    "http://[::1]/a.pdf"):
        assert fetch(blocked) is None, blocked
    assert public("example.com") and not public("localhost") and not public("no-such-host.invalid")

    from PIL import Image, ImageDraw
    img = Image.new("RGB", (400, 200), "white")
    ImageDraw.Draw(img).text((10, 10), "x", fill="black")
    buf = io.BytesIO()
    img.save(buf, "PDF")
    pages, text = pdf_text(buf.getvalue())
    assert pages == 1 and text == "", (pages, text)  # скан: страница есть, текста нет

    page = ("<html><head><style>p{}</style><script>var x=1</script></head><body><nav>Меню Главная</nav>"
            "<article><h1>Заголовок</h1><p>Первый абзац<br>второй</p></article><footer>©</footer></body></html>").encode()
    assert html_text(page) == "Заголовок Первый абзац второй", html_text(page)
    assert html_text("<p>Ülikool</p>".encode("cp1257"), "windows-1257") == "Ülikool"
    assert html_text(b"<p>x</p>", "no-such-charset") == "x"

    fake = lambda url: (("pdf", buf.getvalue(), "") if url.endswith(".pdf") else ("html", page, "utf-8")
                        if "page" in url else None)
    got = source_texts([(1, "Автор. Статья. https://site.ee/a.pdf"), (2, "Сайт. https://site.ee/page"),
                        (3, "Книга. https://repo.ee/bitstream/1/k.pdf."), (4, "Без адреса. 2020."),
                        (5, "Мёртвая ссылка. https://site.ee/404")], fake)
    assert set(got) == {1, 2, 3} and got[1][0] == "pdf" and got[2] == ("html", 0, html_text(page)), got
    assert urls_texts(["https://x.ee/page", "https://x.ee/404"], fake) == [("https://x.ee/page", html_text(page))]
    print("ok")
