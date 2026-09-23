"""Демо-данные: один учитель со списком условий. Работы загружаются через страницу проверок."""
import sqlite3, sys, time
from werkzeug.security import generate_password_hash
import app as A

con = sqlite3.connect(A.DB)
con.row_factory = sqlite3.Row
if con.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]:
    sys.exit("База не пустая. Удалите data.db и запустите снова.")

uid = con.execute("INSERT INTO users(name,login,pw_hash,lang,created_at) VALUES(?,?,?,?,?)",
                  ("Мария Петровна Соколова", "teacher", generate_password_hash("teacher123"), "ru", time.time())).lastrowid
con.executemany("INSERT INTO requirements(teacher_id,position,text,rule,value) VALUES(?,?,?,?,?)",
                [(uid, i, text, rule, value) for i, (text, rule, value) in enumerate(A.DEFAULT_REQUIREMENTS, 1)])
con.commit()
con.close()
print("Готово: учитель teacher / teacher123")
