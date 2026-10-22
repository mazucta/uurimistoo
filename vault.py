"""Шифрование работ учеников в базе: файл, текст, имя, найденные цитаты.

Кто унесёт data.db (или её копию), получит шум: ключ лежит не рядом с базой, а в связке ключей macOS
или в переменной окружения DATA_KEY на сервере. AES-256-GCM: подделанное значение не расшифруется, а упадёт.
Ключ потерян — данные не восстановить, поэтому его стоит сохранить отдельно (python vault.py --show-key).
"""
import base64
import os
import subprocess
import sys

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

TEXT = "enc1:"
BLOB = b"ENC1"
KEYCHAIN = ("school-editor", "data-key")  # служба и запись в связке ключей macOS

# Колонка шифруется только при условии: в ai_note у незаконченной проверки служебная пометка, её ищут по равенству
ONLY_IF = {("papers", "ai_note"): "ai_status='done'"}

# Что шифруем: всё, где лежит сама работа или куски из неё. Остальное (статусы, числа, настройки) открыто:
# по нему идут запросы. Записывают значения триггеры базы, поэтому ни одна запись в коде их не обойдёт.
COLUMNS = {
    "papers": ("student", "title", "filename", "text", "data", "verdict_note", "sources_edit", "teacher", "ai_note"),
    "findings": ("label", "note", "extra"),
    "canva_tokens": ("access", "refresh"),
    "lessons": ("lines",),
}


def _keychain(create):
    service, account = KEYCHAIN
    got = subprocess.run(["security", "find-generic-password", "-s", service, "-a", account, "-w"],
                         capture_output=True, text=True)
    if got.returncode == 0 and got.stdout.strip():
        return got.stdout.strip()
    if not create:
        return ""
    key = base64.b64encode(os.urandom(32)).decode()
    subprocess.run(["security", "add-generic-password", "-s", service, "-a", account, "-w", key, "-U"], check=True)
    return key


def load_key():
    """Ключ из DATA_KEY или из связки ключей macOS (там он создаётся при первом запуске)."""
    raw = os.environ.get("DATA_KEY", "").strip() or (_keychain(create=True) if sys.platform == "darwin" else "")
    if not raw:
        raise RuntimeError("Нет ключа шифрования: задайте DATA_KEY (python vault.py --new-key)")
    key = base64.b64decode(raw)
    if len(key) != 32:
        raise RuntimeError("DATA_KEY должен быть 32 байта в base64")
    return AESGCM(key)


class Vault:
    def __init__(self, aead):
        self.aead = aead

    def seal(self, value):
        """Зашифровать; уже зашифрованное, пустое и не строку/байты оставить как есть."""
        if isinstance(value, str) and value and not value.startswith(TEXT):
            nonce = os.urandom(12)
            return TEXT + base64.b64encode(nonce + self.aead.encrypt(nonce, value.encode(), None)).decode()
        if isinstance(value, bytes) and value and not value.startswith(BLOB):
            nonce = os.urandom(12)
            return BLOB + nonce + self.aead.encrypt(nonce, value, None)
        return value

    def open(self, value):
        if isinstance(value, str) and value.startswith(TEXT):
            raw = base64.b64decode(value[len(TEXT):])
            return self.aead.decrypt(raw[:12], raw[12:], None).decode()
        if isinstance(value, bytes) and value.startswith(BLOB):
            return self.aead.decrypt(value[4:16], value[16:], None)
        return value

    def install(self, con):
        """Подключить к соединению: функция seal() для триггеров и расшифровка при чтении."""
        con.create_function("seal", 1, self.seal, deterministic=False)
        con.row_factory = lambda cursor, row: Row(cursor, row, self.open)
        return con


class Row(tuple):
    """Как sqlite3.Row (row[0], row["name"], dict(row)), только значения уже расшифрованы."""

    def __new__(cls, cursor, row, open_):
        obj = super().__new__(cls, [open_(v) for v in row])
        obj._names = [d[0] for d in cursor.description]
        return obj

    def __getitem__(self, key):
        return tuple.__getitem__(self, self._names.index(key) if isinstance(key, str) else key)

    def keys(self):
        return list(self._names)


def _when(table, column, prefix=""):
    cond = ONLY_IF.get((table, column))
    return f"{prefix}{cond}" if cond else "1"


def triggers():
    """Триггеры: после любой вставки и правки нужных колонок значение в базе заменяется зашифрованным."""
    out = []
    for table, cols in COLUMNS.items():
        key = "id" if table != "canva_tokens" else "user_id"
        sets = ", ".join(f"{c}=CASE WHEN {_when(table, c, 'NEW.')} THEN seal(NEW.{c}) ELSE NEW.{c} END" for c in cols)
        watched = ", ".join(cols + tuple(w.split("=")[0] for (t, _c), w in ONLY_IF.items() if t == table))
        out.append(f"DROP TRIGGER IF EXISTS {table}_seal_ins; DROP TRIGGER IF EXISTS {table}_seal_upd;")  # список колонок мог измениться
        out.append(f"CREATE TRIGGER {table}_seal_ins AFTER INSERT ON {table} BEGIN "
                   f"UPDATE {table} SET {sets} WHERE {key}=NEW.{key}; END;")
        out.append(f"CREATE TRIGGER {table}_seal_upd AFTER UPDATE OF {watched} ON {table} BEGIN "
                   f"UPDATE {table} SET {sets} WHERE {key}=NEW.{key}; END;")
    return "\n".join(out)


def seal_existing(con):
    """Однократно зашифровать то, что лежало открыто до шифрования. Ответ: сколько строк тронуто."""
    n = 0
    for table, cols in COLUMNS.items():
        cond = " OR ".join(f"({_when(table, c)} AND ((typeof({c})='text' AND {c}<>'' AND substr({c},1,5)<>'{TEXT}') OR "
                           f"(typeof({c})='blob' AND length({c})>0 AND substr({c},1,4)<>CAST('ENC1' AS BLOB))))" for c in cols)
        sets = ", ".join(f"{c}=CASE WHEN {_when(table, c)} THEN seal({c}) ELSE {c} END" for c in cols)
        n += con.execute(f"UPDATE {table} SET {sets} WHERE {cond}").rowcount
    return n


if __name__ == "__main__":
    if "--new-key" in sys.argv:  # ключ для сервера: положить в DATA_KEY
        print(base64.b64encode(os.urandom(32)).decode())
        sys.exit()
    if "--show-key" in sys.argv:  # резервная копия ключа: без неё зашифрованное не прочитать
        print(_keychain(create=False) or os.environ.get("DATA_KEY", "") or "ключа нет")
        sys.exit()
    import sqlite3
    v = Vault(AESGCM(os.urandom(32)))
    con = v.install(sqlite3.connect(":memory:"))
    con.executescript("CREATE TABLE papers(id INTEGER PRIMARY KEY, student TEXT, title TEXT, filename TEXT, text TEXT,"
                      " data BLOB, verdict_note TEXT, sources_edit TEXT, teacher TEXT, ai_note TEXT, ai_status TEXT);"
                      "CREATE TABLE findings(id INTEGER PRIMARY KEY, label TEXT, note TEXT, extra TEXT);"
                      "CREATE TABLE canva_tokens(user_id INTEGER PRIMARY KEY, access TEXT, refresh TEXT);"
                      "CREATE TABLE lessons(id INTEGER PRIMARY KEY, lines TEXT);")
    con.execute("INSERT INTO papers(student,text,data,ai_status) VALUES('Петренко','старый открытый текст',?, 'x')",
                (b"%PDF old",))
    con.executescript(triggers())
    assert seal_existing(con) == 1 and seal_existing(con) == 0  # второй раз шифровать нечего
    con.execute("INSERT INTO papers(student,title,text,data) VALUES(?,?,?,?)", ("Іваненко", "", "Текст роботи", b"PK\x03\x04"))
    con.execute("UPDATE papers SET text='Новий текст', ai_status='done' WHERE id=2")
    con.execute("INSERT INTO findings(label,note,extra) VALUES('цитата','','')")
    con.execute("INSERT INTO papers(ai_status, ai_note) VALUES('pending', '↻ пометка')")  # служебная: остаётся как есть
    assert con.execute("SELECT id FROM papers WHERE ai_note='↻ пометка'").fetchone() == (3,)
    con.execute("UPDATE papers SET ai_status='done', ai_note='Підсумок про роботу' WHERE id=3")
    dump = "\n".join(con.iterdump())  # то, что увидит укравший файл
    assert "Іваненко" not in dump and "Новий текст" not in dump and "Петренко" not in dump and "цитата" not in dump \
        and "Підсумок" not in dump, dump
    row = con.execute("SELECT id, student, title, text, data, ai_status FROM papers WHERE id=2").fetchone()
    assert row == (2, "Іваненко", "", "Новий текст", b"PK\x03\x04", "done") and row["text"] == "Новий текст"
    assert dict(row)["student"] == "Іваненко" and con.execute("SELECT text FROM papers WHERE id=1").fetchone()[0] == "старый открытый текст"
    con.row_factory = None  # подмена байтов не должна расшифроваться
    sealed = con.execute("SELECT text FROM papers WHERE id=2").fetchone()[0]
    try:
        v.open(sealed[:-4] + ("AAAA" if not sealed.endswith("AAAA") else "BBBB"))
        raise AssertionError("подменённое значение не должно расшифроваться")
    except Exception as ex:
        assert not isinstance(ex, AssertionError), ex
    print("ok")
