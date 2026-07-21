# Технічне завдання v4.2: спрощений MVP для автоматизації заявок на аутстаффінг (Smart Solutions)

> **Версія:** 4.2 (MVP Lite + Інтегрована технічна специфікація та кодові контракти)
> **Оновлено:** 20.07.2026
> **Принцип версії:** Зберегти цінні технічні напрацювання та кодові контракти попередніх ітерацій, але повністю підпорядкувати їх спрощеній бізнес-логіці першого запуску (MVP Lite). Процес максиально лінеаризовано: `форма → Excel у пам'яті → email менеджеру → ручний експорт у PDF`. Усі складні архітектурні надбудови (багатоклієнтська динамічна маршрутизація, збереження копій у тимчасові каталоги сервера, автоматична конвертація) свідомо винесені за межі релізу.
> **Статус:** Остаточний документ для розробки та приймального тестування першої робочої версії.

---

## 1. Мета MVP

Автоматизувати первинний збір даних за заявкою на аутстаффінг: клієнт самостійно заповнює вебформу, а відповідальний менеджер Smart Solutions миттєво отримує готовий Excel-файл на електронну пошту. 

Застосунок **не** будує CRM/HRM-систему, не веде серверний реєстр чи базу заявок, не зберігає історію на сервері і не має окремого інтерфейсу чи кабінету менеджера. Весь подальший життєвий цикл документа (перевірка, внесення службових реквізитів, експорт у PDF та накладання КЕП) менеджер виконує самостійно у своєму штатному настільному середовищі (MS Excel або LibreOffice).

---

## 2. Цільовий процес MVP Lite

1. Клієнт відкриває вебформу за єдиним наданим йому службовим посиланнями, яке містить унікальний ключ маршруту `route_key`.
2. Клієнт заповнює видимі поля форми та натискає «Надіслати».
3. Застосунок валідовує дані, **генерує Excel-файл виключно в пам'яті** (через `BytesIO`) та надсилає його вкладенням на server-side адресу менеджера через API обраного поштового провайдера.
4. Якщо поштовий сервіс успішно підтвердив прийняття листа, клієнт бачить повідомлення про успіх та унікальний номер звернення (`submission_id`). Кнопка завантаження файлу в цьому сценарії ховається для запобігання дублювання каналів передачі.
5. Якщо поштовий API повернув помилку або сплив таймаут з'єднання, застосунок переходить у fallback-режим: показує погоджену інструкцію для ручної передачі та відображає `st.download_button` для скачувания згенерованого XLSX-файлу з пам'яті поточної сесії.
6. Менеджер отримує файл на корпоративну пошту, відкриває його, заповнює внутрішні поля (номер заявки, службову дату заявки), візуально перевіряє друкований лист і самостійно експортує його в PDF для подальшого підписання поза межами системи.

---

## 3. Що свідомо не входить у MVP

Нижче наведено функціонал, який **заборонено** реалізовувати в поточній ітерації розробки:
* Окремий інтерфейс, веб-режим або окремий застосунок для менеджера (`manager_app.py`).
* База даних, серверний реєстр заявок, історія звернень, дашборди або фонове логування вмісту полів.
* Багаторівнева складна маршрутизація виду `manager_key + client_key`.
* Зберігання копій файлів у директорії `/tmp` на сервері та механізми їхнього очищення за TTL.
* Складна ідемпотентна логіка submit з фіксацією ключів у зовнішніх кеш-системах.
* Автоматичний серверний рендер або конвертація Excel-листа у PDF.

---

## 4. Спрощена маршрутизація

Для MVP використовується **один** стабільний, непрозорий `route_key` у URL-адресі:

```text
https://<app>.streamlit.app/?r=<route_key>
```

Цього параметра повністю достатньо для пілотного запуску з одним менеджером та одним погодженим клієнтським шаблоном друкованої форми. Сервер виконує розрізнення та маппінг ключа виключно всередині `st.secrets`. Усі конфіденційні метадані (email менеджера, назва клієнта, реквізити договору) ніколи не передаються через URL-параметри та не доступні у клієнтському браузері — вони підставляються сервером під час збирання документа.

### Код: `core/routes.py`

```python
import streamlit as st

def resolve_route(route_key: str) -> dict:
    if not route_key:
        raise ValueError("Ключ маршруту відсутній")
        
    routes = st.secrets.get("routes", {})
    route = routes.get(route_key)
    if not route or not route.get("active"):
        raise ValueError("Посилання форми недійсне або деактивоване")

    return {
        "route_id": route["route_id"],
        "email": st.secrets["managers"][route["email_secret_key"]],
        "client_template": route["client_template"],  # Конфігурація реквізитів та юридичного тексту
    }
```

### Код: використання в `client_app.py`

```python
import streamlit as st
from core.routes import resolve_route

route_key = st.query_params.get("r", "")
try:
    route = resolve_route(route_key)
except ValueError:
    st.error("Посилання форми недійсне. Зверніться до менеджера Smart Solutions.")
    st.stop()
```

---

## 5. Незмінний контракт Листа 1

Файл-шаблон містить один лист із 27 колонками A–AA. Генератор відкриває копію оригінального шаблону через `openpyxl.load_workbook()`, записує значення тільки в Рядок 2 і не змінює жодного заголовка, порядку, кількості колонок, ширини чи назви Листа 1 (`Реєстр_Даних`).

### Код: заголовки шаблону та маппінг

```python
TEMPLATE_HEADERS = [
    "Заявка", "заявка", "П.І.Б.", "ІПН", "Дата прийома", "Термін дії СТД",
    "Посада згідно кп", "Випробувальний термін", "Номер телефону", "E-mail",
    "Оклад гросс", "Бонус", "Періодичність бонуса", "Робота",
    "Дата виплата авансу та ЗП", "Графік роботи (з-по)",
    "Графік роботи количество годин", "Кількість дней відпустки",
    "Місто та місце роботи", "Формат роботи", "Матеріальна відповідальність",
    "Інші компенсації", "Коментарі", "Реквізити картки",
    "Функціональні обов'язки", "Примітка", "документи",
]

HEADER_TO_COL = {header: index + 1 for index, header in enumerate(TEMPLATE_HEADERS)}
```

### Код: перевірка контракту шаблону перед записом

```python
def validate_template_headers(ws) -> None:
    actual = [ws.cell(row=1, column=i).value for i in range(1, len(TEMPLATE_HEADERS) + 1)]
    if actual != TEMPLATE_HEADERS:
        raise ValueError("Порушено суворий контракт шаблону Excel: заголовки або їхній порядок не збігаються")
```

### Особливості обробки полів:
* **Колонка A (`Заявка`):** Клієнтська форма не містить цього поля. Воно генерується порожнім, виділяється жовтим кольором і заповнюється безпосередньо менеджером в Excel.
* **Колонка B (`заявка`):** Випадковий дубль колонки A у вихідному бізнес-шаблоні. Структурно зберігається для сумісності, але розробником не використовується і не заповнюється.
* **Місто та місце роботи:** Для зручності UX клієнтська вебформа містить два роздільних текстових поля: «Місто» та «Адреса / місце роботи». Перед записом у комірку стовпця S (19) вони автоматично конкатенуються в один рядок: `", ".join(part for part in [city.strip(), address.strip()] if part)`.
* **Термін дії СТД / проєкту:** Клієнт може ввести або конкретну дату (наприклад, `31.12.2027`), або вказати текстовий шаблон `до завершення Заявки`. Поле записується в існуючу колонку без зміни типу даних усього стовпця.

---

## 6. Валідація і безпечний запис

У першій версії потрібна тільки базова, практична валідація форматів на рівні Streamlit. Всі текстові значення від клієнта перед записом у Excel проходять обов'язкове екранування від ін'єкцій формул.

### Код: `core/security.py`

```python
import re

FORMULA_PREFIXES = ("=", "+", "-", "@")

def escape_excel_text(value):
    if value is None:
        return ""
    value = str(value)
    # Запобігання Excel Injection (додавання апострофа, якщо рядок починається з формульного префікса)
    return "'" + value if value.lstrip().startswith(FORMULA_PREFIXES) else value

def normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if digits.startswith("380") and len(digits) == 12:
        return "+" + digits
    if digits.startswith("0") and len(digits) == 10:
        return "+38" + digits
    raise ValueError("Некоректний формат номера телефону. Формат: +380XXXXXXXXX")

def validate_ipn(value: str) -> str:
    value = re.sub(r"\D", "", value or "")
    if not re.fullmatch(r"\d{10}", value):
        raise ValueError("ІПН має містити рівно 10 цифр")
    return value

def validate_iban(value: str) -> str:
    normalized = re.sub(r"\s+", "", (value or "").upper())
    if not re.fullmatch(r"UA\d{27}", normalized):
        raise ValueError("Некоректний формат IBAN (має бути UA + 27 цифр)")
    return normalized

def write_client_value(ws, row: int, col: int, value) -> None:
    # Дати, числа та булеві значення записуються як рідні типи Python; строки екрануються
    ws.cell(row=row, column=col).value = escape_excel_text(value) if isinstance(value, str) else value
```

---

## 7. Генерація Excel-пакета

На одну заявку генерується один `.xlsx` файл з трьома листами: `Реєстр_Даних`, `Форма_Печати` та `Інструкція`. Застосунок не записує файл у постійне сховище сервера, а повертає масив байтів.

### Код: `core/export.py` — базовий каркас

```python
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from core.schema import HEADER_TO_COL, validate_template_headers
from core.security import write_client_value

YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")

def build_workbook(template_path: str, form_data: dict, client_template: dict) -> bytes:
    wb = load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]  # Перший лист 'Реєстр_Даних'
    validate_template_headers(ws)

    # Запис клієнтських даних у Рядок 2
    for header, value in form_data.items():
        if header not in HEADER_TO_COL:
            continue
        col = HEADER_TO_COL[header]
        write_client_value(ws, 2, col, value)

    # Оформлення службової комірки А2 для менеджера
    ws["A2"].value = None
    ws["A2"].fill = YELLOW_FILL
    ws["A2"].comment = Comment(
        "Поле заповнює менеджер Smart Solutions після отримання заявки",
        "system",
    )

    # Збирання суміжних листів
    from core.export_print import build_print_sheet
    from core.export_instruction import build_instruction_sheet
    
    build_print_sheet(wb, ws, client_template)
    build_instruction_sheet(wb)

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
```

---

## 8. Лист `Форма_Печати`

Друкований лист має бути простим і надійним бланком, пристосованим для ручного експорту оператором. Базові реквізити, які менеджер може дозаповнити або виправити, підтягуються через **прямі формули-посилання** на Рядок 2 Листа 1. 

Довгі багаторядкові тексти (`Примітка` як джерело вимог до кандидатів та `Функціональні обов'язки`) за бізнес-логікою MVP формуються на `Форма_Печати` як **статичні нумеровані рядки** під час генерації файлу за допомогою Python. Це виключає використання несумісних формул типу `TEXTSPLIT`.

### Код: розрахунок безпечної висоти об'єднаного рядка

Оскільки Excel не підтримує AutoFit для об'єднаних комірок із прапорцем `wrap_text=True`, висота кожного такого рядка розраховується за математичною евристикою із коефіцієнтом запасу:

```python
import math

CHARS_PER_LINE = 75       # Розрахункова ширина текстової зони колонок B:C
LINE_HEIGHT_PT = 18       # Висота одного рядка тексту в пунктах
MIN_ROW_HEIGHT_PT = 24    # Мінімальна базова висота рядка бланка

def set_merged_row_height(ws, row: int, text: str) -> None:
    if not text:
        ws.row_dimensions[row].height = MIN_ROW_HEIGHT_PT
        return

    explicit_paragraphs = str(text).replace("
", "
").split("
")
    total_lines = 0
    for paragraph in explicit_paragraphs:
        paragraph_len = len(paragraph) if paragraph else 1
        total_lines += max(1, math.ceil(paragraph_len / CHARS_PER_LINE))

    # Коефіцієнт 1.30 додає необхідний запас висоти для запобігання обрізанню нижніх елементів літер
    ws.row_dimensions[row].height = max(
        MIN_ROW_HEIGHT_PT,
        math.ceil(total_lines * LINE_HEIGHT_PT * 1.30),
    )
```

### Код: `core/export_print.py` — каркас друкованого листа

Лист блокується від випадкового ручного руйнування макета, сітка вимикається, а параметри сторінки підганяються під формат А4 за шириною. Для обов'язкового реквізиту дати заявки виділяється окрема необ'єднана комірка `C5`, яка залишається **розблорованою** та фарбується у жовтий колір.

```python
from openpyxl.styles import Alignment, Protection, PatternFill
from core.schema import HEADER_TO_COL
from core.security import escape_excel_text
import datetime

YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")

def split_lines(value: str) -> list[str]:
    return [line.strip() for line in str(value or "").replace("
", "
").split("
") if line.strip()]

def build_print_sheet(wb, ws_data, client_template: dict) -> None:
    ws_p = wb.create_sheet("Форма_Печати")
    ws_p.views.sheetView[0].showGridLines = False

    # Логічні зв'язки через формули-посилання на Лист 1
    ws_p["A2"] = "№ заявки"
    ws_p["C2"] = "='Реєстр_Даних'!A2"

    ws_p["A4"] = "Дата прийома"
    ws_p["C4"] = "='Реєстр_Даних'!E2"
    ws_p["C4"].number_format = "DD.MM.YYYY"
    ws_p["C4"].alignment = Alignment(horizontal="right")

    # Службова розблокована зона дати заявки (заповнює менеджер перед PDF)
    ws_p["A5"] = "Дата заявки"
    ws_p["C5"] = datetime.date.today().strftime("«%d» %m %Y року")
    ws_p["C5"].fill = YELLOW_FILL
    ws_p["C5"].alignment = Alignment(horizontal="right")
    ws_p["C5"].protection = Protection(locked=False)  # Залишається доступною для редагування

    # Статичний рендер багаторядкових полів
    duties = split_lines(ws_data.cell(row=2, column=HEADER_TO_COL["Функціональні обов'язки"]).value)
    row_ptr = 20
    
    font_regular = ws_data.cell(row=1, column=1).font # Або створення кастомного шрифту Arial
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center")

    for index, duty in enumerate(duties, start=1):
        ws_p.cell(row=row_ptr, column=1, value=f"{index}.").alignment = align_center
        ws_p.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr, end_column=3)
        
        cell = ws_p.cell(row=row_ptr, column=2, value=escape_excel_text(duty))
        cell.font = font_regular
        cell.alignment = align_left
        
        from core.export_print import set_merged_row_height
        set_merged_row_height(ws_p, row_ptr, duty)
        row_ptr += 1

    # Захист структури листа
    ws_p.protection.sheet = True
    ws_p.protection.enableSelection = "unlockedCells"
    
    # Налаштування області друку
    ws_p.page_setup.orientation = "portrait"
    ws_p.page_setup.paperSize = ws_p.PAPERSIZE_A4
    ws_p.sheet_properties.pageSetUpPr.fitToPage = True
    ws_p.page_setup.fitToWidth = 1
    ws_p.page_setup.fitToHeight = 0
    ws_p.print_area = f"A1:C{max(row_ptr, 40)}"
```

---

## 9. Лист `Інструкція`

Лист містить короткий, безальтернативний операційний регламент для менеджера.

### Код: `core/export_instruction.py`

```python
def build_instruction_sheet(wb) -> None:
    ws_i = wb.create_sheet("Інструкція")
    ws_i.views.sheetView[0].showGridLines = True
    
    instructions = [
        "1. Заповніть номер заявки у жовтій комірці A2 на листі 'Реєстр_Даних'.",
        "2. За потреби скоригуйте службову дату заявки у жовтій комірці C5 на листі 'Форма_Печати'.",
        "3. Візуально перевірте коректність відображення бланка на листі 'Форма_Печати'.",
        "4. Експортуйте / збережіть лише лист 'Форма_Печати' як PDF через засоби Excel або LibreOffice.",
        "5. УВАГА: Якщо ви редагували довгі тексти обов'язків чи вимог на першому листі, синхронізуйте їх на Листі 2 вручную.",
        "6. Збережіть робочий файл у папку клієнта за правилом: Клієнт → Обмін → Оформлення_дата оформлення."
    ]
    
    for index, text in enumerate(instructions, start=1):
        ws_i.cell(row=index, column=1, value=text)
```

---

## 10. Email-відправка

Логіка надсилання абстрагована через протокол, що дозволяє залишатися незалежними від конкретного провайдера (SMTP, Brevo, Resend) і спрощує тестування.

### Код: `core/mailer.py`

```python
from typing import Protocol
from email.message import EmailMessage

class MailProviderProtocol(Protocol):
    def send_email(self, msg: EmailMessage) -> bool:
        ...

class OutstaffingMailer:
    def __init__(self, provider: MailProviderProtocol, from_addr: str):
        self.provider = provider
        self.from_addr = from_addr

    def send_document_package(self, file_bytes: bytes, filename: str, to_addr: str, submission_id: str) -> bool:
        msg = EmailMessage()
        msg["Subject"] = f"Smart Solutions | Нова заявка на аутстаффінг [ID: {submission_id}]"
        msg["From"] = self.from_addr
        msg["To"] = to_addr
        
        msg.set_content(
            f"Отримано нову заявку на оформлення співробітника.
"
            f"Номер звернення (Submission ID): {submission_id}

"
            f"Файл згенеровано автоматично та додано до вкладення."
        )
        
        msg.add_attachment(
            file_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )
        return self.provider.send_email(msg)
```

---

## 11. Сценарій Submit та Fallback

Обробка результату надсилання виконується виключно в рамках поточної сесії через стан оперативної пам'яті (`st.session_state`), без задіяння локального диска сервера.

### Код: submit-flow у `client_app.py`

```python
import streamlit as st
from uuid import uuid4
from datetime import datetime
from core.export import build_workbook

def make_filename(pib: str) -> str:
    # Нормалізація імені для файлової системи
    safe_pib = "_".join((pib or "Без_ПІБ").split())
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Smart_Solutions_Заявка_Нова_{safe_pib}_{ts}.xlsx"

# Блокування повторних натискань кнопки на час транзакції
if st.button("Надіслати форму", disabled=st.session_state.get("sending", False)):
    try:
        st.session_state["sending"] = True
        submission_id = f"SS-{datetime.now().strftime('%Y%m%d')}-{uuid4().hex[:4].upper()}"
        
        # Збирання даних форми (припускаємо наявність заповненого словника form_data)
        file_bytes = build_workbook("assets/client_template.xlsx", form_data, route["client_template"])
        filename = make_filename(form_data.get("П.І.Б.", ""))

        # Виклик поштового модуля (mailer ініціалізовано на рівні старту додатка)
        ok = mailer.send_document_package(
            file_bytes=file_bytes,
            filename=filename,
            to_addr=route["email"],
            submission_id=submission_id,
        )

        if ok:
            st.success(f"Форму успішно надіслано менеджеру. Номер звернення: {submission_id}")
            # Кнопка скачування НЕ показується згідно з бізнес-вимогами пілота
        else:
            st.error("Не вдалося автоматично надіслати файл поштою. Завантажте згенерований файл самостійно та передайте менеджеру вручну.")
            st.download_button(
                "Завантажити сформований Excel",
                data=file_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        # Маскування технічних деталей системних помилок для кінцевого клієнта
        st.error("Під час обробки форми виникла технічна помилка. Будь ласка, зверніться до підтримки.")
    finally:
        st.session_state["sending"] = False
```

---

## 12. Чернетка форми и збереження станів

Зберігання проміжних значень елементів вводу здійснюється за допомогою стандартних механізмів `st.session_state`. Це дозволяє користувачеві уникати втрати даних при випадкових перезавантаженнях віджетів в рамках однієї сесії, але повністю нівелює ризики витоку конфіденційної інформації, оскільки дані безповоротно видаляються при закритті вкладки.

---

## 13. Антиспам і захист

* **Шифрування даних:** Робота веб-застосунку дозволена виключно через захищений протокол HTTPS.
* **Суворий Rate Limit:** Захист за замовчуванням реалізується на рівні пам'яті інстансу сервера (через словник фіксації IP-адрес користувачів): дозволено **не більше 5 submit-спроб на 1 IP-адресу протягом 1 години** для поточної пари ключів маршруту. При перевищенні ліміту видається нейтральний блок: *"Перевищено ліміт запитів. Спробуйте пізніше"*.
* **Конфіденційність технічного журналу:** Категорично **заборонено** виводити в системні логи payload запиту, значення ІПН, IBAN, ПІБ, номери телефонів або вміст коментарів. Дозволено фіксувати лише метадані транзакції (час, `submission_id`, `route_id`, статус відправки).

---

## 14. Мінімальна структура репозиторію

```text
├── client_app.py              # Основний вхідний файл та інтерфейс Streamlit
├── requirements.txt           # Залежності проекту (streamlit, openpyxl, pydantic)
├── README.md                  # Документація з налаштування та копіювання st.secrets
├── assets/
│   └── client_template.xlsx   # Оригінальний незмінний 27-колонковий бізнес-шаблон
└── core/
    ├── schema.py              # Контракт та масиви заголовків, верифікація Листа 1
    ├── routes.py              # Роутинг та розбір параметрів URL
    ├── export.py              # Головний диспетчер збирання книги
    ├── export_print.py        # Рендер друкованої форми бланка та розрахунок висоти рядків
    ├── export_instruction.py  # Наповнення листа інструкцій оператора
    ├── security.py            # Очищення від ін'єкцій формул, нормалізація, валідація
    └── mailer.py              # Поштовий адаптер та протокол взаємодії
```

---

## 15. Обов'язкові тести для MVP Lite

Запуск системи в промислову експлуатацію можливий лише після 100% успішного виконання наступних тестів:

1. **Тест суворого контракту заголовків:** Перевірка, що згенерований файл містить рівно 27 колонок, а їхні імена у Рядку 1 повністю та посимвольно відповідають масиву `TEMPLATE_HEADERS`. Поява 28-ї або суміжних колонок розцінюється як критичний збій.
2. **Тест цілісності комірки A2:** Перевірка, що перша комірка даних таблиці реєстру порожня, пофарбована в колір `FFF2CC` і містить текстовий коментар для оператора.
3. **Тест очищення префіксів (Excel Injection):** Перевірка, що при введенні клієнтом у текстовому полі обов'язків формульних символів (`=`, `+`, `-`, `@`), у комірку записується екранована строка з апострофом (наприклад, `'=1+1`).
4. **Тест формул друкованої форми:** Перевірка, що при ручному внесенні номера заявки у `Реєстр_Даних!A2` движок Excel або LibreOffice автоматично відображає це значення на аркуші `Форма_Печати` у полі номера.
5. **Тест автоматичного розрахунку висоти рядків:** Перевірка, що довгі багаторядкові блоки тексту обов'язків коректно розділяються на статичні пункти на другому листі, а функція `set_merged_row_height()` пропорційно збільшує висоту об'єднаних комірок B:C, виключаючи зрізання тексту в PDF.
6. **Тест ізоляції захисту:** Перевірка, що всі елементи верстки бланка заблоковані від випадкового пошкодження, а службова комірка дати заявки `C5` доступна для введення даних.
7. **Тест успішного сценарію:** Перевірка, що при отриманні коду успіху від поштового API інтерфейс Streamlit приховує кнопку скачування файлу.
8. **Тест fallback-сценарію:** Емуляція таймауту поштового провайдера. Перевірка, що користувачеві виводиться заклик до ручної відправки файлу на адресу `O.Korolova@smart-hr.com.ua`, а кнопка скачування віддає валідний файл безпосередньо з оперативної пам'яті сесії.
9. **Тест ізоляції secrets:** Перевірка, що передача вигаданого, некоректного чи неактивного ключа `route_key` повністю блокує відображення інтерфейсу форми за допомогою `st.stop()`.

---

## 16. Post-MVP

Після успішної перевірки працездатності базового ланцюжка та накопичення первинного досвіду експлуатації, до розробки можуть бути прийняті наступні розширення:
* Перехід на роздільну багатоклієнтську маршрутизацію виду `manager_key + client_key`.
* Підтримка кількох кастомних клієнтських шаблонів друку на рівні server-side конфігурацій.
* Інтеграція фонового збереження копій файлів у директорію `/tmp` з очищенням по TTL для підвищення надійності fallback-сценаріїв.
* Автоматичний серверний рендер друкованого листа у PDF за допомогою спеціалізованих headless-бібліотек.

---

## 17. Головний принцип версії 4.2

У поточній версії повністю збережено надійну технічну базу — ізольовані кодові модулі, суворі перевірки контрактів структури, безпечний запис, рендер друкованого бланка за допомогою openpyxl, provider-agnostic mailer — проте всі вони жорстко підпорядковані **спрощеній та лінійній бізнес-логіці MVP Lite**, а не навпаки.

Якщо під час розробки виникає дилема між «красивою гнучкою архітектурою на майбутнє» та «простою стабільною робочою першою поставкою», пріоритет завжди віддається **простій та стабільній робочій першій поставці**.