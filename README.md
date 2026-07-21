# Smart Solutions — Автоматизація заявок на аутстаффінг (MVP Lite v4.2)

Застосунок для автоматизації первинного збору даних за заявкою на аутстаффінг.
Клієнт заповнює вебформу Streamlit, після чого система генерує `.xlsx` файл у пам'яті (`BytesIO`) з трьома листами (`Реєстр_Даних`, `Форма_Печати`, `Інструкція`) та надсилає його менеджеру електронною поштою (або надає можливість скачування у разі fallback).

---

## 1. Встановлення залежностей

```bash
python -m pip install -r requirements.txt
```

---

## 2. Конфігурація `.streamlit/secrets.toml`

Перед запуском переконайтеся, що файл `.streamlit/secrets.toml` містить налаштування маршрутів та пошти (зразок у `.streamlit/secrets.toml.template`):

```toml
[routes.demo_route_123]
route_id = "DEMO-001"
active = true
email_secret_key = "manager_olena"
client_template = "standard"

[managers]
manager_olena = "O.Korolova@smart-hr.com.ua"

[smtp]
host = "smtp.example.com"
port = 587
username = "notifications@example.com"
password = "secret_password"
from_addr = "notifications@example.com"
use_tls = true
use_mock = true  # При true використовується MockMailProvider для розробки
```

---

## 3. Запуск веб-застосунку

```bash
streamlit run client_app.py
```

Для тестування форми відкрийте у браузері посилання з тестовим роутинг-ключем `r`:
[http://localhost:8501/?r=demo_route_123](http://localhost:8501/?r=demo_route_123)

---

## 4. Запуск автоматичних pytest-тестів

```bash
python -m pytest -v tests/test_all_requirements.py
```

---

## 5. Структура проекту

```text
├── client_app.py              # Основний вхідний файл та веб-інтерфейс Streamlit
├── requirements.txt           # Залежності проекту
├── README.md                  # Інструкція з налаштування та запуску
├── .streamlit/
│   ├── secrets.toml.template  # Шаблон секретів
│   └── secrets.toml           # Робочий файл секретів
├── assets/
│   └── client_template.xlsx   # Оригінальний 27-колонковий бізнес-шаблон
├── core/
│   ├── schema.py              # Контракт та масиви заголовків, верифікація Листа 1
│   ├── security.py            # Очищення від ін'єкцій формул, нормалізація, валідація
│   ├── routes.py              # Роутинг та розбір параметрів URL
│   ├── export_instruction.py  # Наповнення листа інструкцій оператора
│   ├── export_print.py        # Рендер друкованої форми бланка та розрахунок висоти рядків
│   ├── export.py              # Головний диспетчер збирання книги
│   └── mailer.py              # Поштовий адаптер та протокол взаємодії
├── scripts/
│   └── prepare_template.py    # Скрипт підготовки шаблону з вихідного XLSX
└── tests/
    └── test_all_requirements.py # Комплект з 9 обов'язкових pytest-тестів
```
