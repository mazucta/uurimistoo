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
    pid = con.execute("INSERT INTO papers(teacher_id,profile_id,uploaded_at) VALUES(1,1,0)").lastrowid
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
   and c.post("/api/materials", headers=H, json={"kind": "slides", "topic": "тест"}).status_code == 403)
ok("SEC-01 демо по-прежнему читает", c.get("/api/papers", headers=H).status_code == 200)

if failed:
    sys.exit("не прошло: " + ", ".join(failed))
print("ok")
