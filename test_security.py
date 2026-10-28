"""Проверки из аудита от 30.09.2026: сессии, лимиты, транзакции, изоляция отчёта.

Запуск: `.venv/bin/python test_security.py`. Работает на своей временной базе,
внешних запросов не делает: потоки проверки подменяются заглушкой.
"""
import os, sqlite3, sys, tempfile, time

DB = os.path.join(tempfile.mkdtemp(), "test.db")
os.environ["DB"], os.environ["DEMO"] = DB, "0"
os.environ["DATA_KEY"] = "MDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWY="  # свой ключ: связку ключей не трогаем
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app as A
_RealThread = A.threading.Thread
A.threading.Thread = type("T", (), {"__init__": lambda self, **k: None, "start": lambda self: None})  # без ИИ-запросов
c = A.app.test_client()
failed = []
real_ask, real_cli, real_backend = A.ask_claude, A.ask_claude_cli, A.AI_BACKEND  # до заглушек ниже


def ok(name, good, extra=""):
    print(("✓" if good else "✗"), name, extra)
    if not good:
        failed.append(name)

r = c.post("/api/register", json={"name": "Учитель", "login": "teach1", "password": "korotkiy"})
tok = r.json.get("token")
H = {"Authorization": "Bearer " + tok}
ok("BUG-03 профиль есть сразу после регистрации",
   len(c.get("/api/papers", headers=H).json["profiles"]) == 1)
ok("SEC-06 пароль короче 8 не принимается",
   c.post("/api/register", json={"name": "x", "login": "teach2", "password": "1234567"}).status_code == 400)
ok("BUG-02 условие через JSON добавляется",
   c.post("/api/requirements", headers=H, json={"text": "Новое условие", "rule": "ai", "profile": 1}).status_code == 200)

# SEC-02: старая сессия отвергается
con = A.connect()
con.execute("UPDATE sessions SET created_at=0")
con.commit(); con.close()
ok("SEC-02 сессия 1970 года отклонена", c.get("/api/papers", headers=H).status_code == 401)

tok = c.post("/api/login", json={"login": "teach1", "password": "korotkiy"}).json["token"]
H = {"Authorization": "Bearer " + tok}
con = A.connect()
stored = con.execute("SELECT token FROM sessions").fetchone()[0]
ok("SEC-02 в базе лежит хеш, а не токен", stored != tok and len(stored) == 64)
con.close()

# BUG-01: профиль с работой не удаляется, данные целы
sample = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo.db")  # любой файл: он всё равно не читается
pid = c.post("/api/papers", headers=H, data={"profile": "1", "file": (open(sample, "rb"), "s.docx")}).json.get("id")
if not pid:  # файл не разобрался, заводим работу напрямую: тестам нужен только её id
    con = A.connect()
    pid = con.execute("INSERT INTO papers(teacher_id,profile_id,uploaded_at) VALUES(1,1,strftime('%s','now'))").lastrowid
    con.commit(); con.close()
before = c.get("/api/papers", headers=H).json["requirements"]
resp = c.post("/api/profiles/1/delete", headers=H, json={})
after = c.get("/api/papers", headers=H).json["requirements"]
ok("BUG-01 профиль с работой не удаляется", resp.status_code == 400 and len(before) == len(after),
   f"код {resp.status_code}, условий было {len(before)} стало {len(after)}")

# BUG-04: открытие отчёта ничего не пишет
con = A.connect()
snap = con.execute("SELECT COUNT(*), SUM(LENGTH(note)) FROM findings").fetchone()
c.get(f"/api/papers/{pid}?lang=et", headers=H); c.get(f"/api/papers/{pid}?lang=uk", headers=H)
same = con.execute("SELECT COUNT(*), SUM(LENGTH(note)) FROM findings").fetchone()
ok("BUG-04 открытие отчёта не меняет базу", snap == same, f"{snap} → {same}")
con.close()

# SEC-03: повтор не создаёт вторую задачу
started = []
A.threading.Thread = type("T", (), {"__init__": lambda self, **k: started.append(k), "start": lambda self: None})
con = A.connect(); con.execute("UPDATE papers SET ai_status='done'"); con.commit(); con.close()
for _ in range(3):
    c.post(f"/api/papers/{pid}/recheck", headers=H, json={})
ok("SEC-03 три повтора создают одну задачу", len(started) == 1, f"задач {len(started)}")

# квота в час
con = A.connect()
con.execute("DELETE FROM limits"); con.commit()
for _ in range(A.AI_PER_HOUR + 3):
    con.execute("UPDATE papers SET ai_status='done'"); con.commit()
    c.post(f"/api/papers/{pid}/recheck", headers=H, json={})
note = con.execute("SELECT ai_note FROM papers WHERE id=?", (pid,)).fetchone()[0]
ok("SEC-03 сверх квоты приходит понятный отказ", "слишком много" in (note or ""), repr(note))
con.close()

# Решение учителя: баллы зажимаются, чужая работа недоступна, вывод об авторстве только из трёх значений
c.post(f"/api/papers/{pid}/marks", headers=H, json={"score-1": 3, "note-1": "слабая цель", "score-2": 99})
rep = c.get(f"/api/papers/{pid}", headers=H).json
ok("учитель: баллы сохранены и зажаты до 5", rep["teacher_score"] == 8 and rep["teacher"]["marks"]["1"]["note"] == "слабая цель",
   repr(rep.get("teacher")))
other = {"Authorization": "Bearer " + c.post("/api/register", json={"name": "Чужой", "login": "teach9",
                                                                     "password": "korotkiy"}).json["token"]}
ok("учитель: чужую работу не оценить и не скачать",
   c.post(f"/api/papers/{pid}/marks", headers=other, json={"score-1": 1}).status_code == 404
   and c.post(f"/api/papers/{pid}/verdict", headers=other, json={"verdict": "ai"}).status_code == 404
   and c.post(f"/api/papers/{pid}/sources", headers=other, json={"sources": "x"}).status_code == 404
   and c.get(f"/api/papers/{pid}/file", headers=other).status_code == 404)
ok("учитель: вывод об авторстве только из трёх значений",
   c.post(f"/api/papers/{pid}/verdict", headers=H, json={"verdict": "guilty"}).status_code == 400
   and c.post(f"/api/papers/{pid}/verdict", headers=H, json={"verdict": "student", "note": "видел черновики"}).status_code == 200
   and c.get(f"/api/papers/{pid}", headers=H).json["teacher"]["verdict"] == "student")
resp = c.get(f"/api/papers/{pid}/file", headers=H)
ok("отчёт выгружается в .docx", resp.status_code == 200 and resp.data[:2] == b"PK", f"код {resp.status_code}")

# Очередь: оборванная задача поднимается один раз, на втором обрыве становится понятной ошибкой
con = A.connect()
con.execute("UPDATE papers SET ai_status='pending', ai_note=NULL WHERE id=?", (pid,)); con.commit()
first = A.resume_jobs()
state1 = con.execute("SELECT ai_status, ai_note FROM papers WHERE id=?", (pid,)).fetchone()
second = A.resume_jobs()
state2 = con.execute("SELECT ai_status, ai_note FROM papers WHERE id=?", (pid,)).fetchone()
ok("очередь: задача поднята один раз, потом ошибка", first == 1 and state1 == ("pending", A.RESUMED)
   and second == 0 and state2 == ("error", A.INTERRUPTED), f"{state1} {state2}")

# Срок хранения: старая работа уходит вместе с отчётом, свежая остаётся
old = con.execute("INSERT INTO papers(teacher_id,profile_id,uploaded_at) VALUES(1,1,1)").lastrowid
con.execute("INSERT INTO findings(paper_id,kind,label) VALUES(?,'format','x')", (old,)); con.commit()
A._purged_at = 0  # чистка идёт не чаще раза в час: здесь время пришло
ids = [p["id"] for p in c.get("/api/papers", headers=H).json["papers"]]
left = con.execute("SELECT COUNT(*) FROM findings WHERE paper_id=?", (old,)).fetchone()[0]
ok("хранение: работа старше срока удалена с отчётом", old not in ids and pid in ids and left == 0, f"{ids} {left}")
con.close()

con = A.connect()
con.execute("DELETE FROM limits"); con.commit(); con.close()  # квоту выбрал SEC-03

# Canva: вход по OAuth с одноразовым state, без открытого редиректа, импорт кладёт ссылку на дизайн
from urllib.parse import parse_qs, urlparse as _up
calls = []
def fake_canva(method, path, data=None, headers=None):
    calls.append((method, path, headers or {}))
    if path == "/oauth/token":
        return {"access_token": "acc", "refresh_token": "ref", "expires_in": 3600}
    if method == "POST" and path == "/imports":
        return {"job": {"id": "j1", "status": "in_progress"}}
    return {"job": {"id": "j1", "status": "success",
                    "result": {"designs": [{"urls": {"edit_url": "https://www.canva.com/design/X/edit"}}]}}}
A.canva_http, A.CANVA_CLIENT_ID, A.CANVA_CLIENT_SECRET = fake_canva, "cid", "secret"
A.time.sleep = lambda s: None
go = c.get("/api/canva/connect?next=//evil.example/x", headers=H)
query = parse_qs(_up(go.headers["Location"]).query)
ok("canva: вход с PKCE S256", go.status_code == 302 and query["code_challenge_method"] == ["S256"]
   and query["scope"] == ["design:content:write"], go.headers.get("Location", "")[:80])
state = query["state"][0]
ok("canva: чужой state не принимается",
   c.get(f"/api/canva/callback?state={state}&code=z", headers=other).status_code == 400)
back = c.get(f"/api/canva/callback?state={state}&code=z", headers=H)
ok("canva: после входа только свой адрес", back.status_code == 302 and back.headers["Location"].endswith("/materials"),
   back.headers.get("Location"))
ok("canva: state одноразовый", c.get(f"/api/canva/callback?state={state}&code=z", headers=H).status_code == 400)
ok("canva: секрет уходит только в заголовке Basic",
   any(p == "/oauth/token" and h.get("Authorization", "").startswith("Basic ") for _m, p, h in calls))
con = A.connect()
slides = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'slides','Тема','done',?,1)",
                     ('{"title": "Тема", "slides": [{"title": "Первый", "points": ["раз"], "question": "", "notes": ""}]}',)).lastrowid
con.commit()
resp = c.post(f"/api/materials/{slides}/canva", headers=H, json={})
saved = con.execute("SELECT canva_url FROM materials WHERE id=?", (slides,)).fetchone()[0]
imp = next(h for m_, p, h in calls if p == "/imports")
ok("canva: презентация импортирована, ссылка сохранена", resp.json.get("edit_url") == saved == "https://www.canva.com/design/X/edit"
   and "presentation" in imp["Import-Metadata"] and imp["Authorization"] == "Bearer acc", repr(resp.json))
ok("canva: чужую презентацию не отправить", c.post(f"/api/materials/{slides}/canva", headers=other, json={}).status_code == 404)
con.execute("UPDATE canva_tokens SET expires=0"); con.commit()
A.canva_http = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("Canva 400: invalid_grant"))
resp = c.post(f"/api/materials/{slides}/canva", headers=H, json={})
left = con.execute("SELECT COUNT(*) FROM canva_tokens").fetchone()[0]
ok("canva: отозванный доступ стирается и просит подключить заново", left == 0, f"токенов {left}")
con.close()

# Помощник: мусор отклоняется, чужие значения зажимаются, готовый ответ заводит материал через общую квоту
ok("чат: мусор вместо разговора отклонён",
   c.post("/api/materials/chat", headers=H, json={"messages": "x"}).status_code == 400
   and c.post("/api/materials/chat", headers=H, json={"messages": [1]}).status_code == 400)
A.ask_claude = lambda system, content, schema, **k: {"reply": "Какой класс?", "options": ["a"] * 9, "ready": False}
resp = c.post("/api/materials/chat", headers=H, json={"messages": [{"role": "user", "text": "презентация"}]}).json
ok("чат: вопрос возвращается, вариантов не больше четырёх", resp.get("reply") and len(resp["options"]) == 4 and "id" not in resp)
A.ask_claude = lambda system, content, schema, **k: {"reply": "Готовлю", "ready": True, "kind": "quiz", "topic": "Т" * 999,
                                                "grade": "9", "count": 500, "extra": "х" * 5000}
mid = c.post("/api/materials/chat", headers=H, json={"messages": [{"role": "user", "text": "готовить"}]}).json.get("id")
con = A.connect()
row = con.execute("SELECT kind, LENGTH(topic), count, LENGTH(extra), teacher_id FROM materials WHERE id=?", (mid,)).fetchone()
con.close()
ok("чат: материал заведён с теми же пределами, что у формы", row == ("quiz", 300, 30, 1000, 1), repr(row))
ok("материал: подготовка действительно запущена", {"target": A.queued, "args": (A.make_material, mid), "daemon": True} in started, repr(started[-2:]))
fid = c.post("/api/materials", headers=H, json={"kind": "kahoot", "topic": "Фотосинтез"}).json.get("id")
ok("материал из формы тоже запускается", {"target": A.queued, "args": (A.make_material, fid), "daemon": True} in started, repr(started[-2:]))

# Приложения к помощнику: тип по содержимому, чужие не берутся, в модель уходят блоками, к материалу привязываются
import io as _io
png = b"\x89PNG\r\n\x1a\n" + b"0" * 100
up = lambda headers, data, name: c.post("/api/materials/files", headers=headers,
                                        data={"file": (_io.BytesIO(data), name)}, content_type="multipart/form-data")
fake = up(H, b"MZ\x90\x00 exe", "shot.png").status_code
big = up(H, b"\x89PNG\r\n\x1a\n" + b"0" * (A.UPLOAD_MAX + 1), "big.png").status_code
mine_id = up(H, png, "Знімок.png").json["files"][0]["id"]
their_id = up(other, b"Chuzhoi tekst", "chuzhoe.txt").json["files"][0]["id"]
seen = {}
def fake_claude(system, content, schema, **k):
    seen["files"] = [f["name"] for f in k.get("files") or []]
    return {"reply": "Готую", "options": [], "ready": True, "kind": "quiz", "topic": "Фото", "grade": "", "count": 5, "extra": ""}
A.ask_claude = fake_claude
mid_files = c.post("/api/materials/chat", headers=H, json={"messages": [{"role": "user", "text": "ось"}],
                                                            "files": [mine_id, their_id, "1; DROP"]}).json.get("id")
con = A.connect()
linked = con.execute("SELECT name FROM uploads WHERE material_id=?", (mid_files,)).fetchall()
con.close()
ok("приложения: подделка и большой файл отклонены", (fake, big) == (400, 400), f"{fake} {big}")
ok("приложения: в модель и в материал идут только свои", seen.get("files") == ["Знімок.png"] and linked == [("Знімок.png",)],
   f"{seen} {linked}")
blocks = A.file_blocks([{"name": "a.png", "mime": "image/png", "data": png},
                        {"name": "b.pdf", "mime": "application/pdf", "data": b"%PDF-1.4"},
                        {"name": "c.txt", "mime": "text/plain", "data": "Тема: дроби".encode()}])
ok("приложения: картинка, PDF и текст становятся блоками", [b["type"] for b in blocks] == ["image", "document", "text"]
   and "Тема: дроби" in blocks[2]["text"], repr([b["type"] for b in blocks]))
c.post(f"/api/materials/{mid_files}/delete", headers=H, json={})
con = A.connect()
left = con.execute("SELECT COUNT(*) FROM uploads WHERE material_id=?", (mid_files,)).fetchone()[0]
con.close()
ok("приложения: удаляются вместе с материалом", left == 0, f"осталось {left}")

# Учебник в PDF: обрезка до нужных страниц, проверка диапазона, целый учебник в модель не уходит
from pypdf import PdfWriter as _W
book = _W()
for _ in range(60):
    book.add_blank_page(width=200, height=200)
_buf = _io.BytesIO(); book.write(_buf)
book_id = up(H, _buf.getvalue(), "Підручник.pdf").json["files"][0]
whole = c.post("/api/materials/chat", headers=H, json={"messages": [{"role": "user", "text": "ось"}], "files": [book_id["id"]]})
bad = c.post(f"/api/materials/files/{book_id['id']}/pages", headers=H, json={"pages": "12-99"}).status_code
cut = c.post(f"/api/materials/files/{book_id['id']}/pages", headers=H, json={"pages": "12-14, 20"}).json
alien = c.post(f"/api/materials/files/{book_id['id']}/pages", headers=other, json={"pages": "1"}).status_code
ok("учебник: 60 страниц сразу не отправить, кривой диапазон и чужой файл отклонены",
   book_id["pages"] == 60 and whole.status_code == 400 and bad == 400 and alien == 404, f"{whole.status_code} {bad} {alien}")
ok("учебник: остаются только выбранные страницы", cut.get("pages") == 4 and cut["name"].endswith("(с. 12–14, 20)"), repr(cut))
ok("страницы разбираются", A.parse_pages("3, 1-2, 2", 5) == [3, 1, 2] and A.parse_pages("0", 5) is None
   and A.parse_pages("abc", 5) is None and A.ranges([1, 2, 3, 7, 9, 10]) == "1–3, 7, 9–10")

# Просьба переделать материал: модель видит текущий материал и просьбу, старая версия возвращается кнопкой
con = A.connect()
rev = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'quiz','Дроби','done',?,1)",
                  ('{"title": "Дроби", "questions": [{"question": "Ручна правка вчителя"}]}',)).lastrowid
con.execute("DELETE FROM limits"); con.commit()
ok("переделка: чужой материал не переделать",
   c.post(f"/api/materials/{rev}/revise", headers=other, json={"request": "x"}).status_code == 404)
c.post(f"/api/materials/{rev}/revise", headers=H, json={"request": "зроби простіше"})
seen = {}
def reviser(system, content, schema, **k):
    seen["prompt"] = content
    return {"title": "Дроби", "questions": [{"question": "Простіше"}]}
A.ask_claude = reviser
A.make_material(rev)
row = con.execute("SELECT status, content, prev_content, revise FROM materials WHERE id=?", (rev,)).fetchone()
ok("переделка: просьба и ручные правки ушли в модель, старая версия сохранена",
   "зроби простіше" in seen["prompt"] and "Ручна правка вчителя" in seen["prompt"] and row[0] == "done"
   and "Простіше" in row[1] and "Ручна правка" in row[2] and row[3] == "", repr(row))
c.post(f"/api/materials/{rev}/undo", headers=H, json={})
back = con.execute("SELECT content FROM materials WHERE id=?", (rev,)).fetchone()[0]
con.close()
ok("переделка: «вернуть как было» возвращает прежний вариант", "Ручна правка" in back, back[:60])
con = A.connect()
con.execute("UPDATE materials SET lang='uk', status='done' WHERE id=?", (rev,)); con.commit()
c.post(f"/api/materials/{rev}/revise?lang=et", headers=H, data={"request": "додай приклади",
       "file": (_io.BytesIO(png), "Приклад.png")}, content_type="multipart/form-data")
row = con.execute("SELECT lang, revise FROM materials WHERE id=?", (rev,)).fetchone()
example = con.execute("SELECT name FROM uploads WHERE material_id=?", (rev,)).fetchall()
con.close()
ok("переделка: язык материала не меняется, пример прикладывается к материалу",
   row == ("uk", "додай приклади") and example == [("Приклад.png",)], f"{row} {example}")

# Модели по задачам: разговор помощника на дешёвой модели, всё, что создаёт результат, на основной
models = []
A.ask_claude_cli = lambda system, content, schema, blocks=(), model=None: models.append(model) or {}
A.AI_BACKEND = "cli"
real_ask("s", "c", {}, light=True); real_ask("s", "c", {})
A.ask_claude_cli, A.AI_BACKEND = real_cli, real_backend
ok("модели: разговор на CHAT_MODEL, генерация на SUBSCRIPTION_MODEL",
   models == [A.CHAT_MODEL, A.SUBSCRIPTION_MODEL] and A.CHAT_MODEL != A.SUBSCRIPTION_MODEL, repr(models))

# Картинки в материале: модель ставит только свои картинки, правка принимает скриншот, выдача только владельцу,
# в презентацию картинка попадает
import base64 as _b64, zipfile as _zip
tiny = _b64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC")
con = A.connect()
pic_mid = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'slides','Фото','new','',1)").lastrowid
pic_id = con.execute("INSERT INTO uploads(teacher_id,material_id,name,mime,data,created_at) VALUES(1,?,'фото.png','image/png',?,1)",
                     (pic_mid, tiny)).lastrowid
con.commit()
A.ask_claude = lambda system, content, schema, **k: {"title": "Фото", "slides": [
    {"title": "Свій", "points": ["а"], "question": "", "notes": "", "image": pic_id},
    {"title": "Вигаданий", "points": ["б"], "question": "", "notes": "", "image": 999999}]}
A.make_material(pic_mid)
placed = [s_["image"] for s_ in __import__("json").loads(con.execute("SELECT content FROM materials WHERE id=?", (pic_mid,)).fetchone()[0])["slides"]]
ok("картинки: модель ставит только картинки этого материала", placed == [pic_id, 0], repr(placed))
resp = c.post(f"/api/materials/{pic_mid}/edit", headers=H, content_type="multipart/form-data", data={
    "0-title": "Свій", "0-points": "а", "0-image": str(pic_id),
    "1-title": "Новий", "1-points": "в", "1-image": "", "1-image_file": (_io.BytesIO(tiny), "screenshot.png")})
slides = __import__("json").loads(con.execute("SELECT content FROM materials WHERE id=?", (pic_mid,)).fetchone()[0])["slides"]
con.close()
ok("картинки: скриншот в правке становится картинкой пункта", resp.status_code == 200 and slides[0]["image"] == pic_id
   and slides[1]["image"] not in (0, pic_id), f"{resp.status_code} {[x.get('image') for x in slides]}")
ok("картинки: чужой учитель картинку не получит", c.get(f"/api/materials/{pic_mid}/images/{pic_id}", headers=other).status_code == 404
   and c.get(f"/api/materials/{pic_mid}/images/{pic_id}", headers=H).data == tiny)
deck = c.get(f"/api/materials/{pic_mid}/file", headers=H).data
ok("картинки: в презентации .pptx есть картинки", any(n.startswith("ppt/media/") for n in _zip.ZipFile(_io.BytesIO(deck)).namelist()))

# Рабочий лист: страница открывается, таблица настоящая, номера не двоятся, Word собирается
con = A.connect()
ws = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'worksheet','A1','done',?,1)",
                 (__import__("json").dumps({"title": "Kontrolltöö", "tasks": [
                     {"variant": "I variant", "title": "1. KÜSIMUSED", "kind": "questions", "items": ["1. Mis sinu nimi on?"],
                      "header": "", "rows": [], "lines": 0, "score": 2, "answers": ["nimi"], "image": 0},
                     {"variant": "I variant", "title": "2. TEGUSÕNAD", "kind": "table", "items": [], "header": "ISTUMA | MAGAMA",
                      "rows": ["Ma istu____ | "], "lines": 0, "score": 4, "answers": [], "image": 0}]}),)).lastrowid
con.commit(); con.close()
page = c.get(f"/materials/{ws}", headers=H)
html = page.get_data(as_text=True)
ok("рабочий лист: страница с таблицей и линиями, без «1. 1.»", page.status_code == 200 and "sheet-table" in html
   and "answer-line" in html and "Ma istu____" in html and ">1. Mis" not in html, f"код {page.status_code}")
ok("рабочий лист: выгружается в Word", c.get(f"/api/materials/{ws}/file", headers=H).data[:2] == b"PK")
con = A.connect()
ws_pic = con.execute("INSERT INTO uploads(teacher_id,material_id,name,mime,data,created_at) VALUES(1,?,'риба','image/png',?,1)",
                     (ws, tiny)).lastrowid
con.execute("UPDATE materials SET status='new' WHERE id=?", (ws,)); con.commit()
A.ask_claude = lambda system, content, schema, **k: {"title": "Pildid", "tasks": [
    {"variant": "", "title": "1. Mis see on?", "instruction": "", "kind": "questions", "items": ["", ""], "header": "",
     "rows": [], "lines": 0, "score": 2, "answers": ["kala"], "image": 0, "pictures": [ws_pic, 424242, ws_pic]}]}
A.make_material(ws)
row = __import__("json").loads(con.execute("SELECT content FROM materials WHERE id=?", (ws,)).fetchone()[0])["tasks"][0]["pictures"]
con.close()
ok("рабочий лист: ряд картинок только из своих, без повторов", row == [ws_pic], repr(row))
sheet = c.get(f"/api/materials/{ws}/file", headers=H).data
ok("рабочий лист: ряд картинок попадает в Word", any(n.startswith("word/media/") for n in _zip.ZipFile(_io.BytesIO(sheet)).namelist()))

# Учебники: загружаются в библиотеку, не стираются как брошенные, отдают страницы и рисунки, чужому недоступны
from PIL import Image as _Img
pic_buf = _io.BytesIO(); _Img.new("RGB", (400, 300), (200, 40, 40)).save(pic_buf, "PDF")  # «скан»: страница-картинка
book_pdf = _W()
for _ in range(3):
    book_pdf.append(_io.BytesIO(pic_buf.getvalue()))
_buf = _io.BytesIO(); book_pdf.write(_buf)
r_book = c.post("/api/materials/books", headers=H, content_type="multipart/form-data",
                data={"name": "Біологія 7", "file": (_io.BytesIO(_buf.getvalue()), "bio.pdf")})
con = A.connect()
book_id = con.execute("SELECT id FROM uploads WHERE library=1 AND name='Біологія 7'").fetchone()[0]
con.execute("UPDATE uploads SET created_at=0 WHERE id=?", (book_id,)); con.commit()
A.purge_old(con); con.commit()
kept = con.execute("SELECT COUNT(*) FROM uploads WHERE id=?", (book_id,)).fetchone()[0]
con.close()
took = c.post(f"/api/materials/books/{book_id}/take", headers=H, json={"pages": "2-3"}).json
alien_take = c.post(f"/api/materials/books/{book_id}/take", headers=other, json={"pages": "1"}).status_code
ok("учебник: хранится в библиотеке и не стирается сроком", r_book.status_code == 200 and kept == 1, f"{r_book.status_code} {kept}")
ok("учебник: страницы и рисунки становятся вложениями, чужому недоступен",
   took["files"][0]["name"] == "Біологія 7 (с. 2–3)" and took["files"][0]["pages"] == 2
   and len(took["files"]) == 4 and "с. 2" in took["files"][1]["name"] and took["files"][3]["name"].startswith("Біологія 7, с. 3") and alien_take == 404, repr(took))

# PDF материала: рисунки из него достаются отдельными картинками один раз; переделка меняет вид материала
con = A.connect()
with_pdf = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'quiz','Kell','done',?,1)",
                       ('{"title": "Kell", "questions": []}',)).lastrowid
con.execute("INSERT INTO uploads(teacher_id,material_id,name,mime,data,pages,created_at) VALUES(1,?,'Tööraamat.pdf','application/pdf',?,3,1)",
            (with_pdf, _buf.getvalue()))
con.execute("DELETE FROM limits"); con.commit()
c.post(f"/api/materials/{with_pdf}/revise", headers=H, json={"request": "встав малюнки з підручника", "kind": "worksheet"})
got = {}
def sheet_maker(system, content, schema, **k):
    got["files"] = [f["name"] for f in k.get("files") or []]
    got["system"] = system
    return {"title": "Kell", "tasks": []}
A.ask_claude = sheet_maker
A.make_material(with_pdf); A.make_material(with_pdf)  # второй раз рисунки не дублируются
pics = con.execute("SELECT COUNT(*) FROM uploads WHERE material_id=? AND parent>0", (with_pdf,)).fetchone()[0]
kind_now = con.execute("SELECT kind FROM materials WHERE id=?", (with_pdf,)).fetchone()[0]
con.close()
ok("PDF материала: рисунки достаются один раз и уходят в модель отдельными картинками",
   pics == 4 and any(", с. 1," in n for n in got["files"]) and sum(n.endswith("сторінка") for n in got["files"]) == 3, f"{pics} {got.get('files')}")
ok("переделка: вид меняется на рабочий лист, рисовать символами запрещено",
   kind_now == "worksheet" and "Не рисуй символами" in got["system"], kind_now)

# Снимки страниц и вырезы: страница становится картинкой, модель называет края куска, программа вырезает
shots = A.page_snapshots(_buf.getvalue(), [1, 2])
con = A.connect()
snap_mid = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'slides','Сторінка','new','',1)").lastrowid
snap_id = con.execute("INSERT INTO uploads(teacher_id,material_id,name,mime,data,created_at) VALUES(1,?,'с. 1, сторінка','image/png',?,1)",
                      (snap_mid, shots[0][1])).lastrowid
con.commit()
A.ask_claude = lambda system, content, schema, **k: {"title": "x", "cuts": [
    {"source": snap_id, "left": 10, "top": 20, "right": 60, "bottom": 70, "name": "завдання 5"},
    {"source": 777777, "left": 0, "top": 0, "right": 50, "bottom": 50, "name": "чужий"}],
    "slides": [{"title": "a", "points": ["b"], "question": "", "notes": "", "image": -1},
               {"title": "c", "points": ["d"], "question": "", "notes": "", "image": -2}]}
A.make_material(snap_mid)
placed = [x["image"] for x in __import__("json").loads(con.execute("SELECT content FROM materials WHERE id=?", (snap_mid,)).fetchone()[0])["slides"]]
piece = con.execute("SELECT data FROM uploads WHERE id=?", (placed[0],)).fetchone()
con.close()
from PIL import Image as _I
w0, h0 = _I.open(_io.BytesIO(shots[0][1])).size
w1, h1 = _I.open(_io.BytesIO(piece[0])).size if piece else (0, 0)
ok("снимки: страница рендерится в картинку", len(shots) == 2 and shots[0][1][:4] == b"\x89PNG")
ok("вырез: кусок вырезан по краям, чужой источник отброшен", placed[1] == 0 and abs(w1 - w0 * 0.5) <= 2 and abs(h1 - h0 * 0.5) <= 2,
   f"{placed} {w0}x{h0} -> {w1}x{h1}")

# Оригинальность: совпадения с другими работами учителя, цитирование и сильные признаки ИИ считает код
class _Doc:
    def __init__(self, text): self._t = text
    def text(self): return self._t
    def sources(self): return ["Hirth L. 2013."]
    def body(self): return [{"text": self._t, "level": 0}]
    def headings(self): return []
    def citations(self): return []
    def meta(self): return {}
    unknown, margins = frozenset(), {}
common = ("Отрицательные цены на электроэнергию возникают в часы, когда солнечные станции производят больше "
          "энергии, чем может потребить сеть, и производитель доплачивает за отбор мощности.")
con = A.connect()
twin = con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,uploaded_at) VALUES(1,1,'Робота Петра',?,?)",
                   (common + " Інший кінець.", __import__("time").time())).lastrowid
alien_twin = con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,uploaded_at) VALUES(2,1,'Чужа',?,?)",
                         (common, __import__("time").time())).lastrowid
mine = con.execute("INSERT INTO papers(teacher_id,profile_id,title,uploaded_at) VALUES(1,1,'Моя',?)", (__import__("time").time(),)).lastrowid
con.commit(); con.close()
with A.app.test_request_context("/"):
    A.check_paper(mine, _Doc(common + " Конечно! Вот ваш текст. Ещё [7]. Потом длинная мысль. «Длинная цитата без всякой ссылки на источник в этом месте работы» конец."),
                  1, "uk", profile={"rules": '{"do_format": 0, "do_req": 0, "do_photos": 0, "do_sources": 0}'})
    got = {k: [(r["label"], r["status"], r["note"]) for r in A.findings_of(mine, k)] for k in ("match", "cite", "aisign")}
ok("совпадения: найдена своя работа с процентом, чужого учителя не видно",
   len(got["match"]) == 1 and got["match"][0][0].startswith("Робота Петра") and got["match"][0][2].endswith("%"), repr(got["match"]))
ok("цитирование и сильные признаки ИИ попадают в отчёт",
   any("[7]" in n for _l, _s, n in got["cite"]) and any("Цитата" in l for l, _s, _n in got["cite"])
   and any("чат-бота" in l for l, _s, _n in got["aisign"]), repr(got["cite"] + got["aisign"]))
page = c.get(f"/papers/{mine}", headers=H)
with A.app.test_request_context("/"):
    orig = A.originality_pct({"text": common + " ещё десять своих слов про энергию и цены"}, mine)
ok("кольца: оригинальность в процентах", 'class="ring' in page.get_data(as_text=True)
   and orig == 100 - int(got["match"][0][2][:-1]), f"{orig} {got['match'][0][2]}")
ok("отчёт с новыми разделами открывается", page.status_code == 200 and 'id="matches"' in page.get_data(as_text=True)
   and 'id="cites"' in page.get_data(as_text=True), f"код {page.status_code}")

# «Сохранить оформление» не выключает проверки; проверка удалённой работы не пишет в новую с тем же номером
con = A.connect()
prof = con.execute("INSERT INTO profiles(teacher_id,name,rules,created_at) VALUES(1,'Оформлення','{}',1)").lastrowid
con.commit(); con.close()
c.post(f"/api/profiles/{prof}/rules", headers=H, json={"font": "Arial", "justify": "0"})
with A.app.test_request_context("/"):
    saved = A.rules_of(A.profile_of(1, prof))
ok("профиль: «Сохранить оформление» не выключает проверки", saved["font"] == "Arial" and saved["do_sources"] == 1
   and saved["do_citations"] == 1 and saved["justify"] == 0, repr({k: saved[k] for k in ("font", "do_sources", "justify")}))
# Профиль из файла школы: модель разбирает требования, старые условия профиля заменяются, выключатели остаются
def school_rules(system, content, schema, **k):
    seen["import"] = [f["name"] for f in k.get("files") or []]
    return {"format": {**{key: -1 for key in A.IMPORT_NUMBERS}, "size": 14, "line": 1, "left": 999,
                       "font": "Arial", "justify": "no", "page_numbers": "unknown", "sources": "yes"},
            "requirements": [{"text": "Обсяг від 10 сторінок", "rule": "chars_min", "value": "20 000 знаків"},
                             {"text": "Є висновки", "rule": "section", "value": "висновк,kokkuvõte"},
                             {"text": "Є гіпотеза", "rule": "ai", "value": "щось"},
                             {"text": "Кривий", "rule": "sources_min", "value": "багато"},
                             {"text": "", "rule": "ai", "value": ""}]}
A.ask_claude = school_rules
con = A.connect(); con.execute("DELETE FROM limits"); con.commit(); con.close()
doc_up = lambda headers, data, name: c.post(f"/api/profiles/{prof}/import", headers=headers,
                                           data={"file": (_io.BytesIO(data), name)}, content_type="multipart/form-data")
bad_file = doc_up(H, b"MZ\x90 exe", "vimogi.exe").status_code
alien_import = doc_up(other, "Вимоги".encode(), "vimogi.txt").status_code
got = doc_up(H, "Вимоги школи: 10 сторінок, Arial 14.".encode(), "vimogi.txt").json
with A.app.test_request_context("/"):
    after = A.rules_of(A.profile_of(1, prof))
    reqs = [(x["text"], x["rule"], x["value"]) for x in A.requirements_of(1, prof)]
ok("профиль из файла: чужой профиль и не тот файл отклонены", (bad_file, alien_import) == (400, 404), f"{bad_file} {alien_import}")
ok("профиль из файла: оформление по документу, остальное по умолчанию, выключатели прежние",
   got.get("imported") == 4 and after["size"] == 14 and after["line"] == 1 and after["left"] == 3.0 and after["font"] == "Arial"
   and after["justify"] == 0 and after["page_numbers"] == 1 and after["do_sources"] == 1 and seen["import"] == ["vimogi.txt"],
   repr({k: after[k] for k in ("size", "line", "left", "font", "justify")}))
ok("профиль из файла: условия заменены, числа и правила проверены",
   reqs == [("Обсяг від 10 сторінок", "chars_min", "20000"), ("Є висновки", "section", "висновк,kokkuvõte"),
            ("Є гіпотеза", "ai", ""), ("Кривий", "ai", "")], repr(reqs))

con = A.connect()
ghost = con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,ai_status,uploaded_at) VALUES(1,1,'Стара','текст','pending',1)").lastrowid
con.commit()
def swap_paper(system, content, schema, **k):  # пока модель думает, работу удалили, номер достался новой
    db = A.connect()
    db.execute("DELETE FROM papers WHERE id=?", (ghost,))
    db.execute("INSERT INTO papers(id,teacher_id,profile_id,title,ai_status,uploaded_at) VALUES(?,1,1,'Нова','pending',2)", (ghost,))
    db.commit(); db.close()
    return {"summary": "старий результат", "claims": [], "sources": []}
A.ask_claude = swap_paper
A.run_ai(ghost)
row = con.execute("SELECT title, ai_status, ai_note FROM papers WHERE id=?", (ghost,)).fetchone()
con.close()
ok("проверка удалённой работы не пишет в новую с тем же номером", row == ("Нова", "pending", None), repr(row))

# Обучение на правках учителя: его баллы и комментарии идут в следующую проверку, чужие — нет
con = A.connect()
marked = con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,ai_status,uploaded_at,verdict,teacher) "
                     "VALUES(1,1,'Оцінена','т','done',6,'ai',?)",
                     ('{"marks": {"3": {"score": 4, "note": "теорію пов’язано з опитуванням"}}, "verdict": "student"}',)).lastrowid
con.execute("INSERT INTO findings(paper_id,kind,position,ref,label,status) VALUES(?,'rubric',3,3,'x','2')", (marked,))
con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,ai_status,uploaded_at,teacher) VALUES(2,1,'Чужа','т','done',7,?)",
            ('{"marks": {"1": {"score": 1, "note": "ЧУЖИЙ КОМЕНТАР"}}}',))
A.remember(con, marked); A.remember(con, marked + 1)  # так их записывают /marks и /verdict
con.commit()
lessons = A.teacher_lessons(con, 1, -1, {**A.CHECK_DEFAULTS})
none_ = A.teacher_lessons(con, 1, -1, {**A.CHECK_DEFAULTS, "do_rubric": 0, "do_authorship": 0})
con.close()
ok("правки учителя: балл, комментарий и вывод об авторстве идут в модель",
   "ты предлагала 2, учитель поставил 4 — «теорію пов’язано з опитуванням»" in lessons
   and "ты решила ai, учитель решил student" in lessons, lessons)
con = A.connect()
con.execute("DELETE FROM findings WHERE paper_id=?", (marked,)); con.execute("DELETE FROM papers WHERE id=?", (marked,))
con.commit()
kept = A.teacher_lessons(con, 1, -1, {**A.CHECK_DEFAULTS})
own = A.teacher_lessons(con, 1, marked, {**A.CHECK_DEFAULTS})  # номер удалённой достался бы новой: её строки не свои
sealed = con.execute("SELECT 1 FROM lessons WHERE lines LIKE '%опитуванням%'").fetchone()
con.close()
ok("копилка: исправления переживают удаление работы и лежат зашифрованными", kept == lessons and own == lessons
   and not sealed, kept)
ok("правки учителя: чужие не попадают, без оценки и авторства ничего не уходит", "ЧУЖИЙ" not in lessons and none_ == "", lessons)

# PDF-источник: программа скачивает его сама, текст уходит в модель, источник отмечен как существующий
from pypdf import PdfWriter as _PW
con = A.connect()
src_paper = con.execute("INSERT INTO papers(teacher_id,profile_id,title,text,ai_status,uploaded_at) VALUES(1,1,'PDF','текст','pending',5)").lastrowid
con.execute("INSERT INTO findings(paper_id,kind,position,label,status) VALUES(?,'source',1,'Автор. Звіт. https://repo.ee/files/report.pdf','unclear')", (src_paper,))
con.commit()
A.source_pdf.fetch = lambda url: ("pdf", pic_buf.getvalue(), "")  # «скан» из теста учебников
seen_pdf = {}
def pdf_reader(system, content, schema, **k):
    seen_pdf["content"] = content
    return {"summary": "ok", "sources": [{"source": 1, "status": "not_found", "supports": "unclear", "note": "x"}], "claims": []}
A.ask_claude = pdf_reader
A.run_ai(src_paper)
row = con.execute("SELECT status, extra FROM findings WHERE paper_id=? AND kind='source'", (src_paper,)).fetchone()
con.close()
ok("PDF-источник: скачан, текст ушёл в модель, отмечен как существующий",
   "PDF скачан программой" in seen_pdf.get("content", "") and row == ("exists", "pdf:1"), repr(row))

# Оригинальность: кусок, переписанный из источника без кавычек, и работа на ту же тему из сети
import zipfile as _z
W_ = A.docx_read.W[1:-1]
copied = ("Отрицательные цены на электроэнергию возникают в часы когда солнечные станции производят больше энергии "
          "чем может потребить сеть и производитель доплачивает за отбор мощности")
para = lambda text, style="": (f'<w:p>{f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>" if style else ""}'
                               f'<w:r><w:t>{text}</w:t></w:r></w:p>')
_buf = _io.BytesIO()
with _z.ZipFile(_buf, "w") as z:
    z.writestr("word/document.xml", f'<w:document xmlns:w="{W_}"><w:body>' + para("Вступ", "H") + para(
        "Моя думка про ринок Естонії зовсім інша. " + copied + " (Elering, 2024). Цитата «" + copied + "» (Elering, 2024).")
        + para("Список використаних джерел", "H") + "<w:p><w:r><w:t>1. Elering. Звіт. 2024. </w:t></w:r><w:hyperlink><w:r>"
        "<w:t>https://elering.ee/report</w:t></w:r></w:hyperlink></w:p></w:body></w:document>")
    z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W_}"><w:style w:styleId="H"><w:name w:val="heading 1"/></w:style></w:styles>')
con = A.connect()
orig = con.execute("INSERT INTO papers(teacher_id,profile_id,title,filename,text,data,ai_status,uploaded_at) "
                   "VALUES(1,1,'Ціни','w.docx','текст',?,'pending',6)", (_buf.getvalue(),)).lastrowid
con.execute("INSERT INTO findings(paper_id,kind,position,label,status) VALUES(?,'source',1,'Elering. Звіт. 2024. https://elering.ee/report','unclear')", (orig,))
con.commit()
pages_ = {"https://elering.ee/report": "<p>Вступ звіту. " + copied + ". Далі інше.</p>",
          "https://school.ee/uurimus": "<article>Чужа робота: " + copied + "</article>",
          "https://other.ee/clean": "<p>Зовсім інший текст без жодного збігу з роботою учня про щось своє</p>"}
A.source_pdf.fetch = lambda url: ("html", pages_[url].encode(), "utf-8") if url in pages_ else None
seen_html = {}
A.ask_claude = lambda system, content, schema, **k: seen_html.update(content=content) or {"summary": "ok", "claims": [], "matches": [],
    "sources": [{"source": 1, "status": "not_found", "supports": "unclear", "note": "web_fetch: 403"}],
    "similar": [{"title": "Uurimus", "url": "https://school.ee/uurimus"}, {"title": "Чисто", "url": "https://other.ee/clean"},
                {"title": "Свій же", "url": "https://elering.ee/report"}, {"title": "x", "url": "javascript:alert(1)"}]}
A.run_ai(orig)
rows = con.execute("SELECT status, ref, label, note, extra FROM findings WHERE paper_id=? AND kind='match' ORDER BY position", (orig,)).fetchall()
src_status, src_extra = con.execute("SELECT status, extra FROM findings WHERE paper_id=? AND kind='source'", (orig,)).fetchone()
con.close()
by = {r[0]: r for r in rows}
ok("оригинальность: переписанное из источника без кавычек найдено, цитата в кавычках — нет",
   "source" in by and by["source"][1] == 1 and 30 <= int(by["source"][3][:-1]) <= 70 and src_extra == "html", repr(rows))
ok("страница-источник: скачана программой, текст ушёл в модель, «не найден» от модели не ставится",
   src_status == "exists" and "страница скачана программой" in seen_html.get("content", ""), src_status)
ok("оригинальность: работы на ту же тему сравнены, свой источник и не-http пропущены",
   [r[2] for r in rows if r[0] == "similar"] == ["Uurimus", "Чисто"] and rows[-1][3] == "0%"
   and rows[1][4].startswith("https://school.ee/uurimus\nотрицательные"), repr(rows))
page = c.get(f"/papers/{orig}?lang=uk", headers=H)
ok("оригинальность: отчёт показывает совпадения и процент", page.status_code == 200
   and "Переписано з джерела" in page.text and "Робота на ту саму тему" in page.text and "https://school.ee/uurimus" in page.text,
   page.status_code)
with A.app.test_request_context("/"):
    pct = A.originality_pct({"text": "слово " * 40}, orig)
ok("оригинальность: процент учитывает источники и сеть", pct < 100, pct)
ok("оригинальность: отчёт выгружается в .docx", c.get(f"/papers/{orig}/file", headers=H).status_code == 200)

# Версии одной работы: прошлая версия того же ученика — отдельная строка, а не списывание; копия у другого — списывание
def paper_docx(body):
    b = _io.BytesIO()
    with _z.ZipFile(b, "w") as z:
        z.writestr("word/document.xml", f'<w:document xmlns:w="{W_}"><w:body>' + para("Вступ", "H") + para(body)
                   + "</w:body></w:document>")
        z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W_}"><w:style w:styleId="H"><w:name w:val="heading 1"/></w:style></w:styles>')
    return b.getvalue()
A.ask_claude = lambda system, content, schema, **k: {"summary": "ok"}
draft = ("Мій перший варіант дослідження про те, як учні гімназії в Тарту обирають гуртки після уроків і скільки "
         "часу на тиждень вони проводять у шкільній бібліотеці, спортзалі та музичному класі. Тут я пишу про свої висновки.")
up = lambda body, who: c.post("/api/papers", headers=H, data={"profile": "1", "student": who,
                                                            "file": (_io.BytesIO(paper_docx(body)), "v.docx")}).json["id"]
v1 = up(draft, "Марія Т.")
v2 = up(draft + " Новий розділ про те як я сама опитала тридцять однокласників у школі.", " марія т. ")
v3 = up(draft, "Інший учень")
con = A.connect()
m2, m3 = (con.execute("SELECT status, ref, note, extra FROM findings WHERE paper_id=? AND kind='match' ORDER BY position",
                      (x,)).fetchall() for x in (v2, v3))
con.close()
ok("версии: прошлая версия того же ученика сравнена отдельно, новое показано, в списывание не попала",
   m2 and m2[0][:2] == ("previous", v1) and 0 < int(m2[0][2][:-1]) < 100 and "опитала тридцять" in m2[0][3]
   and len(m2) == 1, repr(m2))
ok("версии: такая же работа у другого ученика — совпадение", [r[0] for r in m3][:1] == ["local"]
   and "previous" not in [r[0] for r in m3], repr(m3))
con = A.connect(); con.execute("PRAGMA wal_checkpoint(FULL)"); con.close()
on_disk = b"".join(open(f, "rb").read() for f in (DB, DB + "-wal") if os.path.exists(f))
ok("шифрование: текста работы, имени ученика и цитат нет в файле базы", all(
    w.encode() not in on_disk for w in ("шкільній бібліотеці", "Марія Т.", "опитала тридцять", "Інший учень")))
ok("шифрование: файл базы читает только владелец", oct(os.stat(DB).st_mode)[-3:] == "600", oct(os.stat(DB).st_mode))
page = c.get(f"/papers/{v2}?lang=uk", headers=H)
with A.app.test_request_context("/"):
    pct = A.originality_pct({"text": draft}, v2)
ok("версии: отчёт выгружается в .docx", c.get(f"/papers/{v2}/file", headers=H).status_code == 200)
ok("версии: отчёт показывает прошлую версию, оригинальность от неё не падает",
   "Попередня версія цієї роботи" in page.text and f'href="/papers/{v1}"' in page.text and pct == 100, pct)

# Кто за что платит: хозяин через подписку Claude Code, остальные через ключ API; без Claude Code все через ключ
con = A.connect()
owner, guest = (con.execute("SELECT id FROM users WHERE login=?", (x,)).fetchone()[0] for x in ("teach1", "teach9"))
con.close()
which, A.SUBSCRIPTION_LOGINS, A.AI_BACKEND = A.shutil.which, {"teach1"}, ""
os.environ["ANTHROPIC_API_KEY"] = os.environ["CLAUDE_CODE_OAUTH_TOKEN"] = "test"  # на Linux (CI, Render) без токена Claude Code не считается
A.shutil.which = lambda name: "/usr/local/bin/claude"
routes = (A.backend(owner), A.backend(guest), A.backend())
A.shutil.which = lambda name: None
routes += (A.backend(owner),)
A.shutil.which = which
del os.environ["ANTHROPIC_API_KEY"], os.environ["CLAUDE_CODE_OAUTH_TOKEN"]
ok("оплата: хозяин подпиской, другие ключом, без Claude Code все ключом", routes == ("cli", "api", "api", "api"), repr(routes))

# Без регистрации: страницы нет, учитель из настроек заводится сам, пароль берётся из настроек
A.REGISTRATION, A.TEACHER_LOGIN, A.TEACHER_NAME, A.TEACHER_PASSWORD = False, "lisa", "Ліза", "pervyi-parol"
A.ensure_teacher(); A.ensure_teacher()  # второй запуск не плодит учителя
first = c.post("/api/login", json={"login": "lisa", "password": "pervyi-parol"}).status_code
A.TEACHER_PASSWORD = "vtoroi-parol"; A.ensure_teacher()
old = c.post("/api/login", json={"login": "lisa", "password": "pervyi-parol"}).status_code
new = c.post("/api/login", json={"login": "lisa", "password": "vtoroi-parol"}).status_code
A.TEACHERS = [["Kati", "Kati2", "kati-parol:x"], ["", "short", "1234"]]; A.ensure_teacher()
kati = c.post("/api/login", json={"login": "kati2", "password": "kati-parol:x"}).status_code
con = A.connect()
ok("учителя из TEACHERS заводятся, короткий пароль пропущен", kati == 200
   and not con.execute("SELECT 1 FROM users WHERE login='short'").fetchone())
count = con.execute("SELECT COUNT(*) FROM users WHERE login='lisa'").fetchone()[0]
reqs = con.execute("SELECT COUNT(*) FROM requirements r JOIN users u ON u.id=r.teacher_id WHERE u.login='lisa'").fetchone()[0]
con.close()
ok("без регистрации: страницы нет", c.get("/register").status_code == 404
   and c.post("/api/register", json={"name": "x", "login": "zzz", "password": "12345678"}).status_code == 404)
ok("учитель из настроек: заводится один раз с условиями, пароль меняется настройкой",
   (first, old, new, count) == (200, 400, 200, 1) and reqs > 0, f"{first} {old} {new} {count} {reqs}")
A.REGISTRATION = True

# Очередь: при одном слоте вторая задача ждёт, пока не закончится первая
import threading as _th
_Th = _RealThread
A.AI_SLOTS, _order, _gate = _th.BoundedSemaphore(1), [], _th.Event()
_t1 = _Th(target=A.queued, args=(lambda n: (_gate.wait(5), _order.append(n)), 1)); _t1.start()
time.sleep(0.1)
_t2 = _Th(target=A.queued, args=(_order.append, 2)); _t2.start()
time.sleep(0.1); _waited = _order == []; _gate.set(); _t1.join(); _t2.join()
ok("очередь: задачи по одной", _waited and _order == [1, 2], repr(_order))
A.AI_SLOTS = _th.BoundedSemaphore(2)

# Входа по ссылке нет: адрес /k/ не открывает кабинет
ok("вход по ссылке убран", A.app.test_client().get("/k/sekret-ssylki").status_code == 404)

# повторная проверка во время идущей: список источников не меняется под ногами у модели
con = A.connect()
con.execute("UPDATE papers SET ai_status='pending' WHERE id=?", (pid,))
con.execute("INSERT INTO findings(paper_id,kind,position,label) VALUES(?,'source',1,'Старий А. 2019.')", (pid,)); con.commit()
before = con.execute("SELECT id FROM findings WHERE paper_id=? AND kind='source' ORDER BY id", (pid,)).fetchall()
c.post(f"/api/papers/{pid}/sources", headers=H, json={"sources": "Новий А. 2020.\nІнший Б. 2021."})
c.post(f"/api/papers/{pid}/recheck", headers=H, json={})
after = con.execute("SELECT id FROM findings WHERE paper_id=? AND kind='source' ORDER BY id", (pid,)).fetchall()
edit = con.execute("SELECT sources_edit FROM papers WHERE id=?", (pid,)).fetchone()[0]
con.execute("UPDATE papers SET ai_status='done' WHERE id=?", (pid,)); con.commit()
ok("пока идёт проверка, «заново» и правка источников ничего не трогают", before == after and edit == "" and len(before) == 1,
   repr((before, after, edit)))

# срок хранения: файлы готового материала живут вместе с ним
old_t = time.time() - (A.RETENTION_DAYS + 1) * 86400
mat = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,created_at) VALUES(1,'slides','старий','done',?)", (old_t,)).lastrowid
kept = con.execute("INSERT INTO uploads(teacher_id,material_id,name,mime,data,created_at) VALUES(1,?,'a.png','image/png',x'00',?)", (mat, old_t)).lastrowid
gone = con.execute("INSERT INTO uploads(teacher_id,name,mime,data,created_at) VALUES(1,'b.png','image/png',x'00',?)", (old_t,)).lastrowid
A.purge_old(con); con.commit()
left = {r[0] for r in con.execute("SELECT id FROM uploads WHERE id IN (?,?)", (kept, gone))}
con.close()
ok("срок хранения не стирает картинки готового материала", left == {kept}, repr(left))

# Тесты для учеников: из готовых вопросов, ссылка без аккаунта, баллы, уходы со вкладки, копирование
import json as _json
con = A.connect()
quiz = con.execute("INSERT INTO materials(teacher_id,kind,topic,status,content,created_at) VALUES(1,'quiz','Фотосинтез','done',?,1)",
                   (_json.dumps({"questions": [
                       {"question": "Що поглинає рослина?", "kind": "choice", "options": ["CO2", "O2"], "answer": "CO2", "level": "easy"},
                       {"question": "Де йде фотосинтез?", "kind": "short", "options": [], "answer": "Хлоропласти", "level": "easy"},
                       {"question": "Чому листя зелене?", "kind": "open", "options": [], "answer": "хлорофіл", "level": "hard"}]}),)).lastrowid
con.commit(); con.close()
made = c.post("/api/tests", headers=H, json={"material": quiz}).json
tid = made["id"]
code = A.connect().execute("SELECT code FROM tests WHERE id=?", (tid,)).fetchone()[0]
kid = A.app.test_client()
start = kid.post(f"/t/{code}", data={"student": "  Іван   Петренко "})
page_url = start.headers["Location"]
page = kid.get(page_url).text
ok("тест: ученик входит по ссылке без аккаунта, правильных ответов на странице нет", start.status_code == 302
   and "Що поглинає рослина?" in page and "Хлоропласти" not in page and "хлорофіл" not in page)
kid.post(page_url + "/away", json={"state": "hidden"}); kid.post(page_url + "/away", json={"state": "hidden"})
_c = A.connect(); _c.execute("UPDATE attempts SET away_since=away_since-5 WHERE test_id=?", (tid,)); _c.commit(); _c.close()
kid.post(page_url + "/away", json={"state": "visible"}); kid.post(page_url + "/away", json={"copy": 1})
kid.post(page_url, data={"a0": "CO2", "a1": "хлоропласти!", "a2": "бо хлорофіл"})
kid.post(page_url + "/away", json={"state": "hidden"})  # после сдачи уже не считается
again = kid.post(page_url, data={"a0": "O2", "a1": "", "a2": ""})
row = A.connect().execute("SELECT student, score, total, away_count, away_seconds, copy_tries, answers FROM attempts "
                          "WHERE test_id=?", (tid,)).fetchone()
ok("тест: баллы сами за выбор и короткий ответ, уходы и копирование посчитаны, пересдать нельзя",
   tuple(row)[:4] == ("Іван Петренко", 2, 2, 1) and row[4] >= 5 and row[5] == 1 and "CO2" in row[6]
   and again.status_code == 302, repr(tuple(row)))
raw = A.connect(); raw.row_factory = None
sealed = raw.execute("SELECT student, answers FROM attempts WHERE test_id=?", (tid,)).fetchone()
ok("тест: имя ученика и ответы лежат зашифрованными", all(str(v).startswith("enc1:") for v in sealed), repr(sealed))
seen = c.get(f"/api/tests/{tid}", headers=H).json
other = c.get(f"/api/tests/{tid}", headers={"Authorization": "Bearer " + c.post("/api/login",
    json={"login": "lisa", "password": "vtoroi-parol"}).json["token"]})
ok("тест: учитель видит результат, чужой учитель — нет", seen["attempts"][0]["score"] == 2
   and seen["link"].endswith("/t/" + code) and other.status_code == 404)
qr = c.get(f"/api/tests/{tid}/qr.svg", headers=H)
ok("тест: QR-код ссылки отдаётся учителю, чужим нет", qr.status_code == 200 and qr.mimetype == "image/svg+xml"
   and b"<svg" in qr.data and A.app.test_client().get(f"/tests/{tid}/qr.svg").status_code == 302)
# отметка учителя: развёрнутый ответ засчитан, итог 3 из 3; «как было» возвращает
aid = seen["attempts"][0]["id"]
c.post(f"/api/tests/{tid}/attempts/{aid}/mark", headers=H, json={"q": "2", "mark": "1"})
after = c.get(f"/api/tests/{tid}", headers=H).json["attempts"][0]
c.post(f"/api/tests/{tid}/attempts/{aid}/mark", headers=H, json={"q": "2", "mark": ""})
back = c.get(f"/api/tests/{tid}", headers=H).json["attempts"][0]
bad = c.post(f"/api/tests/{tid}/attempts/{aid}/mark", headers=H, json={"q": "9", "mark": "1"}).status_code
ok("тест: учитель засчитывает развёрнутый ответ и может вернуть как было", (after["score"], after["total"], after["unmarked"])
   == (3, 3, 0) and (back["score"], back["total"], back["unmarked"]) == (2, 2, 1) and bad == 400, repr(after))
xl = c.get(f"/api/tests/{tid}/results", headers=H)
import zipfile as _zf, io as _io2
sheet = _zf.ZipFile(_io2.BytesIO(xl.data)).read("xl/worksheets/sheet1.xml").decode()
ok("тест: результаты выгружаются в Excel", xl.status_code == 200 and "Іван Петренко" in sheet, xl.status_code)
# перемешивание: порядок свой у ученика, ответ пишется под исходным номером
orders = [_json.loads(r[0]) for r in A.connect().execute("SELECT orders FROM attempts WHERE test_id=?", (tid,))]
ok("тест: у попытки свой порядок вопросов и вариантов", all(sorted(o["q"]) == [0, 1, 2] and sorted(o["o"]["0"]) == [0, 1]
                                                            for o in orders), repr(orders))
# таймер: тест на 1 минуту показывает остаток, сдача после лимита помечена
con = A.connect()
tq = con.execute("INSERT INTO tests(teacher_id,code,title,questions,minutes,created_at) VALUES(1,'timed1','T',?,1,1)",
                 (_json.dumps([{"q": "a", "kind": "short", "options": [], "answer": "b"}]),)).lastrowid
con.commit(); con.close()
tk = A.app.test_client(); turl = tk.post("/t/timed1", data={"student": "Повільний"}).headers["Location"]
timed_page = tk.get(turl).text
con = A.connect(); con.execute("UPDATE attempts SET started_at=started_at-200 WHERE test_id=?", (tq,)); con.commit(); con.close()
tk.post(turl, data={"a0": "b"})
late_row = c.get(f"/api/tests/{tq}", headers=H).json["attempts"][0]
ok("тест: таймер на странице, сдача после лимита помечена", 'data-left="' in timed_page and late_row["late"], late_row["late"])
# резервная копия: зашифрованная копия базы, владельцу скачивается, остальным нет
_bdir = tempfile.mkdtemp()
_copy = A.backup(_bdir)
_raw = open(_copy, "rb").read()
ok("копия базы: создаётся, права 600, имён учеников открытым текстом нет", os.stat(_copy).st_mode & 0o777 == 0o600
   and "Іван Петренко".encode() not in _raw and len(_raw) > 0)
A.TEACHER_LOGIN = "lisa"
lisa_h = {"Authorization": "Bearer " + c.post("/api/login", json={"login": "lisa", "password": "vtoroi-parol"}).json["token"]}
ok("копия базы: скачивает только владелец", c.get("/api/backup", headers=lisa_h).status_code == 200
   and c.get("/api/backup", headers=H).status_code == 404)
c.post(f"/api/tests/{tid}/open", headers=H, json={})
late = A.app.test_client().post(f"/t/{code}", data={"student": "Запізнився"})
ok("тест: закрытый не принимает новых", late.status_code == 400 and "Тест закрит" in late.text, late.status_code)
ok("тест: чужая ссылка и чужая попытка 404", A.app.test_client().get("/t/nema").status_code == 404
   and kid.get(f"/t/{code}/chuzhyi-token").status_code == 404)

# SEC-07 и SEC-04
resp = c.get(f"/api/papers/{pid}", headers=H)
ok("SEC-07 у отчёта Cache-Control: no-store", resp.headers.get("Cache-Control") == "no-store")
cookie = c.post("/login", data={"login": "teach1", "password": "korotkiy"})
resp = c.post("/papers/%d/recheck" % pid, headers={"Origin": "https://evil.example"})
ok("SEC-04 POST с чужого сайта отклонён", resp.status_code == 403, f"код {resp.status_code}")
resp = c.post("/papers/%d/recheck" % pid, headers={"Origin": "http://localhost"})
ok("SEC-04 POST со своего сайта проходит", resp.status_code in (302, 200), f"код {resp.status_code}")

# SEC-01: демонстрационный вход только смотрит
con = A.connect()
con.execute("UPDATE users SET readonly=1 WHERE id=1")
con.commit(); con.close()
ok("SEC-01 демо не может загружать и удалять",
   c.post(f"/api/papers/{pid}/delete", headers=H, json={}).status_code == 403
   and c.post("/api/materials", headers=H, json={"kind": "slides", "topic": "тест"}).status_code == 403
   and c.post("/api/materials/chat", headers=H, json={"messages": [{"role": "user", "text": "x"}]}).status_code == 403)
ok("SEC-01 демо по-прежнему читает", c.get("/api/papers", headers=H).status_code == 200)

if failed:
    sys.exit("не прошло: " + ", ".join(failed))
print("ok")
