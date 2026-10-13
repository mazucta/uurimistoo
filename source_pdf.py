"""PDF из списка источников: программа скачивает его сама и отдаёт модели текст, чтобы сверка цитат и цифр
шла по самому источнику, а не по тому, получилось ли у модели открыть файл.

Адреса пишет ученик, поэтому скачивание осторожное: только http(s), только публичные адреса в интернете
(ни этот компьютер, ни локальная сеть, ни служебные адреса облака), каждый переход по редиректу проверяется
заново, размер и время ограничены, берётся только то, что действительно PDF.
"""
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
MAX_PAGES = 80
MAX_CHARS = 40_000             # текста одного источника в модель: этого хватает для сверки
LOOKS_PDF = re.compile(r"\.pdf(?:$|[?#])|/pdf/|download|/file|bitstream|getfile|viewcontent", re.I)
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


def fetch_pdf(url):
    """Байты PDF по адресу или None. Не PDF, закрытый адрес, слишком большой файл — None."""
    # ponytail: между проверкой адреса и запросом DNS может смениться (rebinding); для пилота хватает,
    # на сервере для настоящих учеников — скачивать через прокси, который сам не пускает во внутреннюю сеть
    opener = urllib.request.build_opener(_NoRedirect)
    for _hop in range(4):
        parts = urllib.parse.urlsplit(url)
        if parts.scheme not in ("http", "https") or not parts.hostname or not public(parts.hostname):
            return None
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (uurimistoo: source check)"})
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
            if head != b"%PDF-":
                return None
            data = head + response.read(MAX_BYTES)
        return data if len(data) <= MAX_BYTES else None
    return None


def pdf_text(data):
    """(страниц, текст) из PDF. У скана текста нет — тогда пустая строка."""
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        reader.decrypt("")
    text = "\n".join((page.extract_text() or "") for page in reader.pages[:MAX_PAGES])
    return len(reader.pages), re.sub(r"[ \t]+", " ", text).strip()[:MAX_CHARS]


def source_pdfs(sources, fetch=None):
    """{номер источника: (страниц, текст)} для источников, у которых адрес ведёт на PDF.
    sources: [(номер, строка списка)]."""
    fetch = fetch or fetch_pdf
    out, start = {}, time.monotonic()
    for number, entry in sources:
        if len(out) >= MAX_PDFS or time.monotonic() - start > BUDGET:
            break
        for url in URL.findall(entry):
            if not LOOKS_PDF.search(url):
                continue
            data = fetch(url.rstrip(".,"))
            if data:
                try:
                    out[number] = pdf_text(data)
                except Exception:  # битый PDF: источник просто идёт к модели как обычно
                    pass
                break
    return out


if __name__ == "__main__":
    for blocked in ("http://127.0.0.1:8000/x.pdf", "http://localhost/x.pdf", "http://169.254.169.254/latest/x.pdf",
                    "http://10.0.0.5/a.pdf", "http://192.168.1.1/a.pdf", "file:///etc/passwd", "ftp://example.com/a.pdf",
                    "http://[::1]/a.pdf"):
        assert fetch_pdf(blocked) is None, blocked
    assert public("example.com") and not public("localhost") and not public("no-such-host.invalid")

    from pypdf import PdfWriter
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (400, 200), "white")
    ImageDraw.Draw(img).text((10, 10), "x", fill="black")
    buf = io.BytesIO()
    img.save(buf, "PDF")
    pages, text = pdf_text(buf.getvalue())
    assert pages == 1 and text == "", (pages, text)  # скан: страница есть, текста нет

    got = source_pdfs([(1, "Автор. Статья. https://site.ee/a.pdf"), (2, "Сайт. https://site.ee/page"),
                       (3, "Книга. https://repo.ee/bitstream/1/k.pdf.")],
                      fetch=lambda url: buf.getvalue() if url.endswith(".pdf") else None)
    assert set(got) == {1, 3}, got  # обычную страницу не качаем, точку в конце записи отрезаем
    print("ok")
