"""Демо-данные: один учитель со списком условий. Работы загружаются через страницу проверок."""
import os, sqlite3, sys, time

os.environ["DEMO"] = "0"  # чистая база, демонстрационную не подставляем
from werkzeug.security import generate_password_hash
import app as A

con = sqlite3.connect(A.DB)
con.row_factory = sqlite3.Row
if con.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]:
    sys.exit("База не пустая. Удалите data.db и запустите снова.")

uid = con.execute("INSERT INTO users(name,login,pw_hash,lang,created_at) VALUES(?,?,?,?,?)",
                  ("Марія Соколова", "teacher", generate_password_hash("teacher123"), "uk", time.time())).lastrowid
prid = con.execute("INSERT INTO profiles(teacher_id,name,rules,created_at) VALUES(?,?,'{}',?)",
                   (uid, A.translate("Исследовательская работа", "uk"), time.time())).lastrowid
con.executemany("INSERT INTO requirements(teacher_id,position,text,rule,value,profile_id) VALUES(?,?,?,?,?,?)",
                [(uid, i, A.translate(text, "uk"), rule, value, prid)
                 for i, (text, rule, value) in enumerate(A.DEFAULT_REQUIREMENTS, 1)])
con.commit()
con.close()
print("Готово: учитель teacher / teacher123")
