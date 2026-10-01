"""Проверки из аудита от 30.09.2026: сессии, лимиты, транзакции, изоляция отчёта.

Запуск: `.venv/bin/python test_security.py`. Работает на своей временной базе,
внешних запросов не делает: потоки проверки подменяются заглушкой.
"""
import os, sqlite3, sys, tempfile

DB = os.path.join(tempfile.mkdtemp(), "test.db")
os.environ["DB"], os.environ["DEMO"] = DB, "0"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app as A
A.threading.Thread = type("T", (), {"__init__": lambda self, **k: None, "start": lambda self: None})  # без ИИ-запросов
c = A.app.test_client()
failed = []


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
con = sqlite3.connect(DB)
con.execute("UPDATE sessions SET created_at=0")
con.commit(); con.close()
ok("SEC-02 сессия 1970 года отклонена", c.get("/api/papers", headers=H).status_code == 401)

tok = c.post("/api/login", json={"login": "teach1", "password": "korotkiy"}).json["token"]
H = {"Authorization": "Bearer " + tok}
con = sqlite3.connect(DB)
stored = con.execute("SELECT token FROM sessions").fetchone()[0]
ok("SEC-02 в базе лежит хеш, а не токен", stored != tok and len(stored) == 64)
con.close()

# BUG-01: профиль с работой не удаляется, данные целы
sample = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo.db")  # любой файл: он всё равно не читается
pid = c.post("/api/papers", headers=H, data={"profile": "1", "file": (open(sample, "rb"), "s.docx")}).json.get("id")
if not pid:  # файл не разобрался, заводим работу напрямую: тестам нужен только её id
    con = sqlite3.connect(DB)
    pid = con.execute("INSERT INTO papers(teacher_id,profile_id,uploaded_at) VALUES(1,1,strftime('%s','now'))").lastrowid
    con.commit(); con.close()
before = c.get("/api/papers", headers=H).json["requirements"]
resp = c.post("/api/profiles/1/delete", headers=H, json={})
after = c.get("/api/papers", headers=H).json["requirements"]
ok("BUG-01 профиль с работой не удаляется", resp.status_code == 400 and len(before) == len(after),
   f"код {resp.status_code}, условий было {len(before)} стало {len(after)}")

# BUG-04: открытие отчёта ничего не пишет
con = sqlite3.connect(DB)
snap = con.execute("SELECT COUNT(*), SUM(LENGTH(note)) FROM findings").fetchone()
c.get(f"/api/papers/{pid}?lang=et", headers=H); c.get(f"/api/papers/{pid}?lang=uk", headers=H)
same = con.execute("SELECT COUNT(*), SUM(LENGTH(note)) FROM findings").fetchone()
ok("BUG-04 открытие отчёта не меняет базу", snap == same, f"{snap} → {same}")
con.close()

# SEC-03: повтор не создаёт вторую задачу
started = []
A.threading.Thread = type("T", (), {"__init__": lambda self, **k: started.append(k), "start": lambda self: None})
con = sqlite3.connect(DB); con.execute("UPDATE papers SET ai_status='done'"); con.commit(); con.close()
for _ in range(3):
    c.post(f"/api/papers/{pid}/recheck", headers=H, json={})
ok("SEC-03 три повтора создают одну задачу", len(started) == 1, f"задач {len(started)}")

# квота в час
con = sqlite3.connect(DB)
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
con = sqlite3.connect(DB)
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
ids = [p["id"] for p in c.get("/api/papers", headers=H).json["papers"]]
left = con.execute("SELECT COUNT(*) FROM findings WHERE paper_id=?", (old,)).fetchone()[0]
ok("хранение: работа старше срока удалена с отчётом", old not in ids and pid in ids and left == 0, f"{ids} {left}")
con.close()

con = sqlite3.connect(DB)
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
con = sqlite3.connect(DB)
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
con = sqlite3.connect(DB)
row = con.execute("SELECT kind, LENGTH(topic), count, LENGTH(extra), teacher_id FROM materials WHERE id=?", (mid,)).fetchone()
con.close()
ok("чат: материал заведён с теми же пределами, что у формы", row == ("quiz", 300, 30, 1000, 1), repr(row))
ok("материал: подготовка действительно запущена", {"target": A.make_material, "args": (mid,), "daemon": True} in started, repr(started[-2:]))
fid = c.post("/api/materials", headers=H, json={"kind": "kahoot", "topic": "Фотосинтез"}).json.get("id")
ok("материал из формы тоже запускается", {"target": A.make_material, "args": (fid,), "daemon": True} in started, repr(started[-2:]))

# Кто за что платит: хозяин через подписку Claude Code, остальные через ключ API; без Claude Code все через ключ
con = sqlite3.connect(DB)
owner, guest = (con.execute("SELECT id FROM users WHERE login=?", (x,)).fetchone()[0] for x in ("teach1", "teach9"))
con.close()
which, A.SUBSCRIPTION_LOGINS, A.AI_BACKEND = A.shutil.which, {"teach1"}, ""
os.environ["ANTHROPIC_API_KEY"] = "test"
A.shutil.which = lambda name: "/usr/local/bin/claude"
routes = (A.backend(owner), A.backend(guest), A.backend())
A.shutil.which = lambda name: None
routes += (A.backend(owner),)
A.shutil.which = which
del os.environ["ANTHROPIC_API_KEY"]
ok("оплата: хозяин подпиской, другие ключом, без Claude Code все ключом", routes == ("cli", "api", "api", "api"), repr(routes))

# Без регистрации: страницы нет, учитель из настроек заводится сам, пароль берётся из настроек
A.REGISTRATION, A.TEACHER_LOGIN, A.TEACHER_NAME, A.TEACHER_PASSWORD = False, "lisa", "Ліза", "pervyi-parol"
A.ensure_teacher(); A.ensure_teacher()  # второй запуск не плодит учителя
first = c.post("/api/login", json={"login": "lisa", "password": "pervyi-parol"}).status_code
A.TEACHER_PASSWORD = "vtoroi-parol"; A.ensure_teacher()
old = c.post("/api/login", json={"login": "lisa", "password": "pervyi-parol"}).status_code
new = c.post("/api/login", json={"login": "lisa", "password": "vtoroi-parol"}).status_code
con = sqlite3.connect(DB)
count = con.execute("SELECT COUNT(*) FROM users WHERE login='lisa'").fetchone()[0]
reqs = con.execute("SELECT COUNT(*) FROM requirements r JOIN users u ON u.id=r.teacher_id WHERE u.login='lisa'").fetchone()[0]
con.close()
ok("без регистрации: страницы нет", c.get("/register").status_code == 404
   and c.post("/api/register", json={"name": "x", "login": "zzz", "password": "12345678"}).status_code == 404)
ok("учитель из настроек: заводится один раз с условиями, пароль меняется настройкой",
   (first, old, new, count) == (200, 400, 200, 1) and reqs > 0, f"{first} {old} {new} {count} {reqs}")
A.REGISTRATION = True

# SEC-07 и SEC-04
resp = c.get(f"/api/papers/{pid}", headers=H)
ok("SEC-07 у отчёта Cache-Control: no-store", resp.headers.get("Cache-Control") == "no-store")
cookie = c.post("/login", data={"login": "teach1", "password": "korotkiy"})
resp = c.post("/papers/%d/recheck" % pid, headers={"Origin": "https://evil.example"})
ok("SEC-04 POST с чужого сайта отклонён", resp.status_code == 403, f"код {resp.status_code}")
resp = c.post("/papers/%d/recheck" % pid, headers={"Origin": "http://localhost"})
ok("SEC-04 POST со своего сайта проходит", resp.status_code in (302, 200), f"код {resp.status_code}")

# SEC-01: демонстрационный вход только смотрит
con = sqlite3.connect(DB)
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
