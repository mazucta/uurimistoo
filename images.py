"""Фотографии из работы и их метаданные: снято камерой или нарисовано ИИ.

Генераторы оставляют следы: C2PA-манифест (Adobe, OpenAI, Google), XMP-поле DigitalSourceType,
имя программы в EXIF или текстовые куски PNG, куда Stable Diffusion пишет промпт целиком.
Камера оставляет марку, модель и дату съёмки. Чистый файл не значит ничего: метаданные слетают
при пересохранении, скриншоте и загрузке в мессенджер, поэтому такие снимки помечаются «не понять».
"""
import io, re, zipfile

MAX_IMAGES = 60
HEAD = 300_000          # метаданные лежат в начале файла, дальше не читаем
AI_WORDS = ("midjourney", "dall-e", "dalle", "stable diffusion", "stablediffusion", "comfyui", "automatic1111",
            "firefly", "adobe firefly", "imagen", "gemini", "sora", "grok", "flux", "leonardo.ai", "ideogram",
            "recraft", "nano banana", "playground ai", "novelai", "krea", "runway", "luma", "generative fill")
AI_MARKS = ("trainedalgorithmicmedia", "compositewithtrainedalgorithmicmedia", "c2pa.created_by_ai",
            "ai_generative", "digitalsourcetype")
CAMERA_HINT = ("iphone", "samsung", "xiaomi", "huawei", "canon", "nikon", "sony", "olympus", "fujifilm", "pixel")
EXIF_TAGS = {0x010F: "Make", 0x0110: "Model", 0x0131: "Software", 0x9003: "DateTimeOriginal", 0x0132: "DateTime"}


def jpeg_meta(data):
    """EXIF, XMP и C2PA из JPEG: идём по маркерам APP1…APP11."""
    out, i = {}, 2
    while i + 4 < min(len(data), HEAD):
        if data[i] != 0xFF:
            i += 1
            continue
        marker, size = data[i + 1], int.from_bytes(data[i + 2:i + 4], "big")
        block = data[i + 4:i + 2 + size]
        if marker == 0xE1 and block.startswith(b"Exif\x00\x00"):
            out.update(exif(block[6:]))
        elif marker == 0xE1 and b"adobe:ns:meta" in block[:200]:
            out["xmp"] = block.decode("utf-8", "replace")
        elif marker == 0xEB or b"jumb" in block[:40] or b"c2pa" in block[:200].lower():
            out["c2pa"] = True
        if marker == 0xDA:  # пошли данные картинки
            break
        i += 2 + max(size, 1)
    return out


def exif(data):
    """Марка, модель, программа и дата съёмки. Разбираем только нужные поля IFD0 и Exif IFD."""
    if len(data) < 8 or data[:2] not in (b"II", b"MM"):
        return {}
    order = "little" if data[:2] == b"II" else "big"
    out, seen = {}, set()

    def read(offset, depth=0):
        if offset + 2 > len(data) or depth > 2 or offset in seen:
            return
        seen.add(offset)
        count = int.from_bytes(data[offset:offset + 2], order)
        for n in range(count):
            at = offset + 2 + n * 12
            if at + 12 > len(data):
                return
            tag = int.from_bytes(data[at:at + 2], order)
            kind = int.from_bytes(data[at + 2:at + 4], order)
            length = int.from_bytes(data[at + 4:at + 8], order)
            value = data[at + 8:at + 12]
            if tag == 0x8769:  # ссылка на Exif IFD, там дата съёмки
                read(int.from_bytes(value, order), depth + 1)
            elif tag in EXIF_TAGS and kind == 2 and length > 1:
                start = int.from_bytes(value, order) if length > 4 else at + 8
                text = data[start:start + length].split(b"\x00")[0].decode("utf-8", "replace").strip()
                if text:
                    out[EXIF_TAGS[tag]] = text[:100]
    read(int.from_bytes(data[4:8], order))
    return out


def png_meta(data):
    """Текстовые куски PNG: туда Stable Diffusion и ComfyUI пишут промпт и настройки."""
    out, i = {}, 8
    while i + 8 < min(len(data), HEAD):
        size = int.from_bytes(data[i:i + 4], "big")
        kind = data[i + 4:i + 8]
        body = data[i + 8:i + 8 + min(size, 20000)]
        if kind in (b"tEXt", b"iTXt", b"zTXt"):
            out.setdefault("text", "")
            out["text"] += " " + body.decode("utf-8", "replace")
        elif kind == b"eXIf":
            out.update(exif(body))
        elif kind == b"caBX":
            out["c2pa"] = True
        if kind == b"IDAT":
            break
        i += 12 + size
    return out


def verdict(meta):
    """Что показать учителю: ИИ, камера или не понять."""
    blob = " ".join(str(v) for v in meta.values()).lower()
    if meta.get("c2pa") or any(w in blob for w in AI_MARKS) or any(w in blob for w in AI_WORDS):
        found = [w for w in AI_WORDS + AI_MARKS if w in blob]
        note = ", ".join(sorted(set(found))[:3]) or "C2PA"
        return "ai", note
    camera = " ".join(filter(None, (meta.get("Make"), meta.get("Model")))).strip()
    if camera:
        when = meta.get("DateTimeOriginal") or meta.get("DateTime") or ""
        return "camera", f"{camera[:60]}{', ' + when[:19] if when else ''}"
    if meta.get("Software"):
        return "unclear", meta["Software"][:60]
    return "unclear", ""


def describe(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        meta = png_meta(data)
    elif data[:3] == b"\xff\xd8\xff":
        meta = jpeg_meta(data)
    else:
        meta = {}
    status, note = verdict(meta)
    return {"status": status, "note": note, "size": len(data)}


def from_docx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = [n for n in z.namelist() if n.startswith("word/media/")][:MAX_IMAGES]
        return [(n.rsplit("/", 1)[-1], z.read(n)) for n in names]


def from_pdf(data):
    """Берём только потоки JPEG: они лежат в PDF как есть, вместе с метаданными.
    Остальные картинки PDF пересобирает в сырые пиксели, метаданных там уже нет."""
    from pypdf import PdfReader
    out = []
    for page in PdfReader(io.BytesIO(data)).pages:
        xobjects = (page.get("/Resources") or {}).get("/XObject")
        for name, ref in (xobjects.get_object().items() if xobjects else []):
            obj = ref.get_object()
            if obj.get("/Subtype") != "/Image":
                continue
            kinds = obj.get("/Filter")
            kinds = [str(k) for k in (kinds if isinstance(kinds, list) else [kinds])]
            if "/DCTDecode" in kinds and getattr(obj, "_data", None):
                out.append((str(name).lstrip("/"), bytes(obj._data)))
                if len(out) >= MAX_IMAGES:
                    return out
    return out


def photos(data, filename):
    """Список фотографий работы: имя, вердикт и пояснение."""
    try:
        found = from_pdf(data) if filename.lower().endswith(".pdf") else from_docx(data)
    except Exception:
        return []
    out, seen = [], set()
    for name, blob in found:
        if len(blob) < 4000:  # иконки и линейки в счёт не идут
            continue
        key = hash(blob)  # одна и та же картинка на нескольких страницах это одна картинка
        if key in seen:
            continue
        seen.add(key)
        out.append({"name": name, **describe(blob)})
    return out


if __name__ == "__main__":
    import struct, zlib

    def png(chunks=b""):
        head = struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0)
        body = zlib.compress(b"\x00" + b"\xff" * 24 * 8)
        def chunk(kind, payload):
            return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload))
        return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", head) + chunks + chunk(b"IDAT", body) + chunk(b"IEND", b"") + b"\x00" * 5000

    def chunk(kind, payload):
        return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload))

    drawn = png(chunk(b"tEXt", b"parameters\x00masterpiece, school building, Steps: 30, Sampler: DPM++ 2M, Model: stable diffusion xl"))
    assert describe(drawn)["status"] == "ai", describe(drawn)
    assert "stable diffusion" in describe(drawn)["note"], describe(drawn)

    def jpeg(exif_payload=b""):
        app1 = b"\xff\xe1" + struct.pack(">H", len(exif_payload) + 8) + b"Exif\x00\x00" + exif_payload if exif_payload else b""
        return b"\xff\xd8\xff\xe0" + struct.pack(">H", 16) + b"JFIF\x00" + b"\x00" * 9 + app1 + b"\xff\xda" + b"\x00" * 6000

    def tiff(tags):
        head = b"II*\x00" + struct.pack("<I", 8)
        entries, blob, offset = b"", b"", 8 + 2 + len(tags) * 12 + 4
        for tag, text in sorted(tags.items()):
            payload = text.encode() + b"\x00"
            entries += struct.pack("<HHI", tag, 2, len(payload)) + struct.pack("<I", offset + len(blob))
            blob += payload
        return head + struct.pack("<H", len(tags)) + entries + struct.pack("<I", 0) + blob

    shot = jpeg(tiff({0x010F: "Apple", 0x0110: "iPhone 14", 0x0132: "2026:05:01 12:00:00"}))
    got = describe(shot)
    assert got["status"] == "camera" and "iPhone 14" in got["note"] and "2026:05:01" in got["note"], got
    made = jpeg(tiff({0x0131: "Midjourney v6"}))
    assert describe(made)["status"] == "ai", describe(made)
    assert describe(png())["status"] == "unclear", describe(png())
    assert describe(b"not an image at all")["status"] == "unclear"
    print("ok")
