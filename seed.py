"""Демо-данные на пустую базу. Руководитель: teacher / teacher123. Ученики: s1…s6 / student123."""
import html as html_mod
import random, re, sqlite3, sys, time, zlib
from werkzeug.security import generate_password_hash
import app as A

con = sqlite3.connect(A.DB)
if con.execute("SELECT COUNT(*) FROM users").fetchone()[0]:
    sys.exit("База не пустая. Удалите data.db и запустите снова.")
rnd, now, DAY = random.Random(7), time.time(), 86400


def ins(sql, *args):
    return con.execute(sql, args).lastrowid


TOPICS = [
    ("Влияние удалённого обучения на успеваемость гимназистов", [
        ("h2", "Введение"),
        ("p", "Цель работы: выяснить, как переход на удалённое обучение отразился на успеваемости гимназистов Эстонии."),
        ("p", "Работа опирается на статистику Департамента статистики, исследования Тартуского университета и собственный опрос."),
        ("h2", "Обзор источников"),
        ("p", "По данным Департамента статистики, в 2020 году на дистанционное обучение перешли почти все школы страны."),
        ("p", "Исследование Тартуского университета показало, что сильнее всего пострадали ученики из семей с низким доходом."),
        ("p", "Результаты PISA говорят о снижении среднего балла по математике примерно на десять пунктов."),
        ("h2", "Методика"),
        ("p", "Я провёл анкету среди 84 гимназистов и сравнил средние оценки своей школы за 2019 и 2021 годы."),
        ("p", "Анкета состояла из двенадцати вопросов, восемь из них с выбором ответа, четыре открытые."),
        ("h2", "Результаты"),
        ("p", "Две трети опрошенных считают удалённое обучение менее эффективным, почти половина отметила трудности с пониманием объяснений."),
        ("p", "Сравнение оценок показало разницу в 0,4 балла, причём у слабых учеников разрыв оказался заметнее."),
        ("h2", "Заключение"),
        ("p", "Удалённое обучение усилило разрыв между сильными и слабыми учениками, поэтому школе нужна адресная поддержка."),
    ]),
    ("Энергосбережение в школьных зданиях Эстонии", [
        ("h2", "Введение"),
        ("p", "Цель работы: понять, какие меры энергосбережения дают школам наибольшую экономию."),
        ("p", "Большая часть школьных зданий Эстонии построена до 1990 года и теряет тепло через фасады и окна."),
        ("h2", "Обзор источников"),
        ("p", "По данным Департамента статистики, потребление энергии школами за пять лет заметно снизилось."),
        ("p", "Европейские программы поддержки покрывают часть расходов на ремонт школьных зданий."),
        ("h2", "Методика"),
        ("p", "Я сравнил счета за отопление двух школ Нарвы и поговорил с директором одной из них."),
        ("h2", "Результаты"),
        ("p", "В утеплённом здании расход ниже почти на треть, окупаемость утепления составила около семи лет."),
        ("p", "Замена окон дала меньший эффект, но обошлась дешевле и быстрее окупилась."),
        ("h2", "Заключение"),
        ("p", "Утепление фасадов и замена окон дают самый заметный эффект, но без привычки экономить энергию техника мало что меняет."),
    ]),
]
SOURCE_POOL = [
    ("Департамент статистики Эстонии", "Образование и обучение", "2021", "https://www.stat.ee/et/haridus", "статистика", "supports"),
    ("Тартуский университет", "Kaugõppe mõju õpitulemustele", "2022", "https://www.ut.ee/et/uuringud", "научная статья", "supports"),
    ("OECD", "PISA 2022 Results", "2023", "https://www.oecd.org/pisa/", "научная статья", "supports"),
    ("ERR", "Гимназисты о дистанционном обучении", "2021", "https://rus.err.ee/haridus", "сайт", "contradicts"),
    ("Министерство образования и науки", "Haridusvaldkonna arengukava 2035", "2021", "https://www.hm.ee/arengukava", "документ", "unrelated"),
    ("Личный опрос автора", "Анкета среди гимназистов", "2026", "", "интервью", None),
    ("Блог учителя", "Как мы пережили карантин", "2020", "https://example.invalid/blog", "сайт", "unreachable"),
]
SOURCE_NOTES = {"supports": "Цифры в работе совпадают с источником.",
                "contradicts": "В источнике речь о другом срезе, вывод ученика шире.",
                "unrelated": "Документ о стратегии, к выводам работы прямого отношения не имеет.",
                "unreachable": "Страница не открылась, ссылку нужно проверить."}
AI_SUMMARY = ("Работа построена по плану, но часть цифр дана без ссылок на источники. "
              "Выводы местами шире того, что подтверждают источники.")
NOTES = [("источник", "Цифра без ссылки на источник"), ("источник", "В источнике другой показатель"),
         ("язык", "Тяжёлая формулировка, лучше разбить на два предложения"), ("язык", "Повтор слова")]
TEACHER_NOTES = ["Добавьте ссылку на источник рядом с цифрой.", "Уточните, как вы считали эту разницу.",
                 "Этот вывод шире, чем данные в источнике.", "Перепишите фразу проще."]
NAMES = ["Кару Мартин", "Лийв Анна", "Мяги Роберт", "Саар Кристина", "Тамм Даниэль", "Хансен Мария"]


def as_html(blocks):
    return "".join(f"<{tag}>{html_mod.escape(t)}</{tag}>" for tag, t in blocks)


def typing(blocks, t, paste_share):
    """Правдоподобный лог: набор по буквам с опечатками и паузами, часть абзацев вставлена целиком."""
    ev, pos = [{"t": t, "type": "focus", "pos": None, "del": None, "ins": None}], 0

    def add(kind, ins="", dele=""):
        nonlocal pos
        p = pos - len(dele)
        ev.append({"t": t, "type": kind, "pos": p, "del": dele, "ins": ins})
        pos = p + len(ins)

    for tag, part in blocks:
        if tag == "p" and len(part) > 40 and rnd.random() < paste_share:
            t += rnd.randint(4000, 30000)
            add("insertFromPaste", part)
            continue
        if rnd.random() < .1:
            t += rnd.randint(15, 40) * 60000
        for ch in part:
            t += rnd.randint(150, 700) + (rnd.randint(5000, 60000) if rnd.random() < .04 else 0)
            if ch.isalpha() and rnd.random() < .03:
                wrong = rnd.choice("аоеинтсрвл")
                add("insertText", wrong)
                t += rnd.randint(200, 700)
                add("deleteContentBackward", dele=wrong)
            add("insertText", ch)
    ev.append({"t": t + 2000, "type": "blur", "pos": None, "del": None, "ins": None})
    return ev


pw = generate_password_hash("student123")
T = ins("INSERT INTO users(role,name,login,pw_hash,lang,invite,created_at) VALUES('teacher',?,?,?,'ru','demo',?)",
        "Мария Петровна Соколова", "teacher", generate_password_hash("teacher123"), now)
con.executemany("INSERT INTO requirements(teacher_id,position,text,rule,value) VALUES(?,?,?,?,?)",
                [(T, i, text, rule, value) for i, (text, rule, value) in enumerate(A.DEFAULT_REQUIREMENTS, 1)])
con.execute("UPDATE requirements SET value='6000' WHERE teacher_id=? AND rule='chars_min'", (T,))
con.execute("UPDATE requirements SET value='3' WHERE teacher_id=? AND rule='sources_min'", (T,))

# статус, доля вставок, сколько абзацев написано, сколько источников
PLAN = [("accepted", .0, 1.0, 5), ("review", .0, 1.0, 5), ("review", .6, .9, 3),
        ("revise", .0, .8, 4), ("draft", .0, .5, 2), ("draft", .3, .2, 0)]

for i, (name, (status, paste, share, n_sources)) in enumerate(zip(NAMES, PLAN), 1):
    uid = ins("INSERT INTO users(role,name,login,pw_hash,lang,teacher_id,created_at) VALUES('student',?,?,?,'et',?,?)",
              name, f"s{i}", pw, T, now)
    title, blocks = TOPICS[i % 2]
    blocks = blocks[:max(2, round(len(blocks) * share))]
    text = "".join(t for _, t in blocks)
    started = now - rnd.uniform(20, 40) * DAY
    ev = typing(blocks, started * 1000, paste)
    assert A.replay(ev)[0] == text
    st = A.process_stats(ev, text)
    submitted = min(ev[-1]["t"] / 1000 + 600, now - 3600) if status != "draft" else None
    wid = ins("""INSERT INTO works(student_id,title,html,text,status,target_chars,deadline,active_sec,paste_chars,
        ai_status,ai_note,created_at,submitted_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
              uid, title, as_html(blocks), text, status, 6000,
              time.strftime("%Y-%m-%d", time.localtime(now + (7 if i % 2 else -3) * DAY)),
              st["active_sec"], st["paste_chars"],
              "done" if status != "draft" else None, AI_SUMMARY if status != "draft" else None,
              started, submitted)
    con.executemany("INSERT INTO events(work_id,t,type,pos,deleted,inserted) VALUES(?,?,?,?,?,?)",
                    [(wid, e["t"], e["type"], e["pos"], e["del"], e["ins"]) for e in ev])
    # история: несколько сохранений по ходу работы, последнее совпадает с текущим текстом
    steps = sorted({max(2, round(len(blocks) * k / 4)) for k in range(1, 5)} | {len(blocks)})
    previous, moment = "", started
    for n, upto in enumerate(steps):
        part_html = as_html(blocks[:upto])
        part_text = "".join(t for _, t in blocks[:upto])
        added, removed = A.diff_counts(previous, part_text)
        moment += rnd.uniform(.5, 3) * DAY
        last = n == len(steps) - 1
        ins("""INSERT INTO versions(work_id,created_at,html,chars,added,removed,reason) VALUES(?,?,?,?,?,?,?)""",
            wid, min(moment, submitted or now - 3600), zlib.compress(part_html.encode()), len(part_text),
            added, removed, "submit" if last and status != "draft" else "save")
        previous = part_text
    for n, (author, stitle, year, url, kind, src_status) in enumerate(rnd.sample(SOURCE_POOL, n_sources), 1):
        checked = status in ("review", "revise", "accepted")
        ins("""INSERT INTO sources(work_id,position,author,title,year,url,kind,status,note,created_at)
            VALUES(?,?,?,?,?,?,?,?,?,?)""", wid, n, author, stitle, year, url, kind,
            src_status if checked else None, SOURCE_NOTES.get(src_status) if checked else None, started)
    if status == "draft":
        continue
    words = [m.span() for m in re.finditer(r"\w+", text)]
    for n, (start, end) in enumerate(rnd.sample(words, 5)):
        kind, note = NOTES[n % len(NOTES)]
        if status == "review":
            row = ("ai", "suggested", kind, note)
        elif status == "revise":
            row = ("teacher", "open" if n % 3 else "done", "", rnd.choice(TEACHER_NOTES))
        else:
            row = ("teacher", "done", "", rnd.choice(TEACHER_NOTES))
        ins("""INSERT INTO comments(work_id,start,end,quote,text,author,kind,status,created_at)
            VALUES(?,?,?,?,?,?,?,?,?)""", wid, start, end, text[start:end], row[3], row[0], row[2], row[1],
            (submitted or started) + 600)

con.commit()
print("Готово.\n  руководитель: teacher / teacher123 (интерфейс по-русски)\n"
      "  ученики: s1 … s6 / student123 (интерфейс по-эстонски, s3 и s6 часто вставляют текст)\n"
      "  приглашение: /join/demo")
