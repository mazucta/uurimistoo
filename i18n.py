"""Интерфейс на русском и эстонском. Ключ перевода это русская строка из шаблона."""
import re
from pathlib import Path

LANGS = ("ru", "et")

ET = {
    # общее
    "Исследовательская работа": "Uurimistöö",
    "Ученик пишет работу, ведёт список источников и получает замечания. Руководитель видит, как появился текст и на что работа опирается.":
        "Õpilane kirjutab tööd, peab allikate nimekirja ja saab märkusi. Juhendaja näeb, kuidas tekst tekkis ja millele töö tugineb.",
    "ученик": "õpilane",
    "Руководитель": "Juhendaja",
    "Ученик": "Õpilane",
    "Выйти": "Logi välja",
    "Вход": "Sisselogimine",
    "Логин": "Kasutajanimi",
    "Пароль": "Parool",
    "Войти": "Logi sisse",
    "Регистрация руководителя": "Juhendaja registreerimine",
    "Имя и отчество": "Nimi",
    "Создать аккаунт": "Loo konto",
    "Уже есть аккаунт?": "Konto on juba olemas?",
    "Руководитель без аккаунта?": "Juhendaja ilma kontota?",
    "Зарегистрироваться": "Registreeru",
    "Ученики входят по ссылке-приглашению от руководителя.": "Õpilased liituvad juhendaja kutselingi kaudu.",
    "Неверный логин или пароль": "Vale kasutajanimi või parool",
    "Укажите имя": "Sisestage nimi",
    "Логин: от 3 до 40 символов, латиница, цифры, точка, дефис": "Kasutajanimi: 3–40 märki, ladina tähed, numbrid, punkt, sidekriips",
    "Пароль не короче 6 символов": "Parool vähemalt 6 märki",
    "Этот логин уже занят": "See kasutajanimi on juba võetud",
    "Слишком много попыток входа. Попробуйте через четверть часа": "Liiga palju sisselogimiskatseid. Proovige veerand tunni pärast",
    # список подопечных
    "Работы": "Tööd",
    "Тема": "Teema",
    "Объём": "Maht",
    "Источники": "Allikad",
    "пишет": "kirjutab",
    "ИИ": "TI",
    # рабочее место ученика
    "Тема работы": "Töö teema",
    "Оформление": "Vormindus",
    "Обычный текст": "Tavaline tekst",
    "Текст": "Tekst",
    "Заголовок раздела": "Peatüki pealkiri",
    "Текст работы": "Töö tekst",
    "Введение": "Sissejuhatus",
    "документ": "dokument",
    "Удалить": "Kustuta",
    "слов": "sõna",
    "знаков": "tähemärki",
    # страница чтения работы
    "из": "/",
    "мин": "min",
    "ч": "h",
    "Проверка ИИ": "TI kontroll",
    "Проверить заново": "Kontrolli uuesti",
    "не открылся": "ei avanenud",
    # условия и история
    "Условия": "Tingimused",
    "Условия к работе": "Töö tingimused",
    "Условий пока нет": "Tingimusi veel pole",
    "Условие, как его прочитает ученик": "Tingimus nii, nagu õpilane seda loeb",
    "одни считает программа, остальные проверяет ИИ": "osa arvutab programm, ülejäänud kontrollib TI",
    "программа": "programm",
    "проверяет ИИ": "kontrollib TI",
    "объём, знаков не меньше": "maht, tähemärke vähemalt",
    "источников не меньше": "allikaid vähemalt",
    "ссылок в тексте не меньше": "viiteid tekstis vähemalt",
    "есть раздел, слова через запятую": "peatükk olemas, sõnad komaga",
    "Значение": "Väärtus",
    "Добавить": "Lisa",
    "выполнено": "täidetud",
    # виды замечаний и сообщения сервера
    "язык": "keel",
    "источник": "allikas",
    "не задан ключ ANTHROPIC_API_KEY на сервере": "serveris pole ANTHROPIC_API_KEY määratud",
    "неверный ключ API": "vale API võti",
    "у ключа API нет доступа к модели": "API võtmel pole mudelile ligipääsu",
    "слишком много запросов, повторите через минуту": "liiga palju päringuid, proovige minuti pärast",
    "нет связи с сервисом ИИ": "TI teenusega puudub ühendus",
    "сервис ИИ временно недоступен": "TI teenus pole hetkel kättesaadav",
    # оформление по правилам и выгрузка в Word
    "Титульный лист": "Tiitelleht",
    "Список использованных источников": "Kasutatud allikad",
    # проверка загруженной работы
    "Проверки": "Kontrollid",
    "Проверки работ": "Tööde kontrollid",
    "Загрузите работу в формате .docx и получите отчёт: оформление, условия и источники.":
        "Laadige töö üles .docx-failina ja saate aruande: vormistus, tingimused ja allikad.",
    "Новая проверка": "Uus kontroll",
    "Чья работа, для себя": "Kelle töö, enda jaoks",
    "Проверить": "Kontrolli",
    "Принимается только .docx: по нему видно настоящее оформление. Проверка с ИИ занимает несколько минут.":
        "Vastu võetakse ainult .docx: sellest on näha tegelik vormistus. TI kontroll võtab mõne minuti.",
    "Проверенные работы": "Kontrollitud tööd",
    "Работа": "Töö",
    "Нарушений": "Rikkumisi",
    "Спорных мест": "Kahtlasi kohti",
    "Проверка ИИ": "TI kontroll",
    "Загружено": "Üles laaditud",
    "Отчёт": "Aruanne",
    "Пока ничего не проверено": "Midagi pole veel kontrollitud",
    "идёт": "käib",
    "ошибка": "viga",
    "готово": "valmis",
    "Оформление по школьным правилам проверяется всегда, его в список добавлять не нужно.":
        "Vormistust kooli reeglite järgi kontrollitakse alati, seda pole vaja nimekirja lisada.",
    "Что говорит ИИ": "Mida ütleb TI",
    "идёт проверка": "kontroll käib",
    "Проверка не удалась": "Kontroll ebaõnnestus",
    "Нажмите «Проверить заново».": "Vajutage „Kontrolli uuesti”.",
    "Модель открывает источники и сверяет их с работой. Страница обновится сама.":
        "Mudel avab allikad ja võrdleb neid tööga. Leht uueneb ise.",
    "Проверка ИИ не запускалась": "TI kontrolli pole käivitatud",
    "Проверить заново": "Kontrolli uuesti",
    "Удалить проверку вместе с файлом?": "Kas kustutada kontroll koos failiga?",
    "Оформление": "Vormistus",
    "Оформление не проверено": "Vormistust pole kontrollitud",
    "Условия учителя": "Õpetaja tingimused",
    "Спорные места": "Kahtlased kohad",
    "текст со ссылкой, который источник не подтверждает": "viitega tekst, mida allikas ei kinnita",
    "Таких мест не нашлось": "Selliseid kohti ei leitud",
    "расходится с источником": "on allikaga vastuolus",
    "нет в источнике": "allikas puudub",
    "существует": "on olemas",
    "не открылся": "ei avanenud",
    "не найден": "ei leitud",
    "не проверен": "pole kontrollitud",
    "В работе нет раздела со списком источников": "Töös puudub kasutatud allikate loetelu",
    # правила оформления
    "Поля: левое 3 см, правое, верхнее и нижнее 2 см": "Veerised: vasak 3 cm, parem, ülemine ja alumine 2 cm",
    "левое": "vasak", "правое": "parem", "верхнее": "ülemine", "нижнее": "alumine", "см": "cm", "пт": "pt",
    "и ещё": "ja veel", "из": "-st", "выравнивание": "joondus",
    "Шрифт Times New Roman": "Kirjatüüp Times New Roman",
    "Кегль 12": "Kirja suurus 12",
    "Междустрочный интервал 1,5": "Reavahe 1,5",
    "Выравнивание по ширине": "Rööpjoondus",
    "Отбивка абзаца 6 пт до и после": "Lõigu vahe 6 pt enne ja pärast",
    "Абзацного отступа нет": "Taandrida puudub",
    "Заголовок раздела: 16 пт, полужирный, с новой страницы": "Peatüki pealkiri: 16 pt, rasvane, uuelt lehelt",
    "Заголовок подраздела: 14 пт, полужирный": "Alapeatüki pealkiri: 14 pt, rasvane",
    "Заголовок пункта: 12 пт, полужирный": "Punkti pealkiri: 12 pt, rasvane",
    "таких заголовков нет": "selliseid pealkirju pole",
    "Заголовки выровнены по левому полю": "Pealkirjad on vasakule joondatud",
    "Номер страницы внизу по центру": "Leheküljenumber all keskel",
    "нумерации страниц нет": "leheküljenumbreid pole",
    "Титульный лист: 14 пт, название работы 20 пт": "Tiitelleht: 14 pt, töö pealkiri 20 pt",
    "титульного листа нет": "tiitellehte pole",
    "Список использованных источников оформлен отдельным разделом": "Kasutatud allikad on eraldi peatükina",
    "раздел со списком источников не найден": "kasutatud allikate peatükki ei leitud",
    "Нужен файл .docx": "Vaja on .docx-faili",
    "Файл не читается как документ Word": "Faili ei õnnestu Wordi dokumendina lugeda",
}


def translate(text, lang):
    return ET.get(text, text) if lang == "et" else text


if __name__ == "__main__":
    used = set()
    for path in Path(__file__).with_name("templates").glob("*.html"):
        for a, b in re.findall(r"_\(\s*'([^']+)'\s*\)|_\(\s*\"([^\"]+)\"\s*\)", path.read_text()):
            used.add(a or b)
    missing = sorted(t for t in used if t not in ET)
    for t in missing:
        print("НЕТ ПЕРЕВОДА:", t)
    for t in sorted(t for t in ET if t not in used):
        print("не из шаблонов (сообщения сервера, виды замечаний):", t)
    assert not missing, f"без перевода: {len(missing)}"
    print(f"ok, строк: {len(used)}")
