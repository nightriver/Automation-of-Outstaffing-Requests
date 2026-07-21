# План реалізації MVP Lite v4.2: Автоматизація заявок на аутстаффінг (Smart Solutions)

На основі детального аналізу [technical-specification-v4.2-smart-solutions.md](file:///d:/!work/smart-solutions/Automation/docs/technical-specification-v4.2-smart-solutions.md) та узгоджених в інтерв'ю рішеннях, нижче подано технічний план розробки першої версії системи.

---

## 1. Основні технічні та архітектурні рішення

* **Архітектура без БД та серверних файлів**: Дані вебформи обробляються виключно в пам'яті (`BytesIO`), згенерований `.xlsx` надсилається електронною поштою менеджеру або віддається через `st.download_button` у разі fallback.
* **Структура Excel (3 листи)**:
  1. `Реєстр_Даних`: Суворий контракт з 27 колонок (`TEMPLATE_HEADERS`). Запис клієнтських даних у рядок 2. Комірка `A2` виділяється коліром `#FFF2CC` з коментарем для менеджера.
  2. `Форма_Печати`: Бланк для друку/експорту в PDF. Зв'язок реквізитів з Листом 1 через Excel-формули (`='Реєстр_Даних'!E2`). Статичне формування нумерованих списків для довгого тексту. Автоматичний розрахунок висоти об'єднаних комірок (`set_merged_row_height`). Комірка дати заявки `C5` розблокована (`Protection(locked=False)`).
  3. `Інструкція`: Текстовий регламент дій оператора з 6 пунктів.
* **Роутинг та безпека**:
  * Однозначний параметр `?r=<route_key>` в URL з перевіркою через `st.secrets["routes"]`.
  * Екранування Excel Injection (` escape_excel_text ` для знаків `=`, `+`, `-`, `@`).
  * Простий rate limiting через `st.session_state` (до 5 спроб).
* **Поштовий модуль (`core/mailer.py`)**:
  * Реалізація через `MailProviderProtocol`.
  * Підтримка Console/Mock поштового провайдера для локального dev/тестування + SMTP/API реалізація.
* **Тестування (`tests/`)**:
  * Повний pytest-комплект з 9 обов'язкових тестів згідно з Розділом 15 ТЗ.

---

## 2. Узгоджені параметри (Результат /grill-me)

1. **Поштовий провайдер**: Консольний / Mock mailer для локального dev із можливістю відправки поштою через SMTP конфігурацію в `st.secrets`.
2. **Шаблон Excel**: Створення `assets/client_template.xlsx` на основі наявного файлу [Приклад таблички от клієнта.xlsx](file:///d:/!work/smart-solutions/Automation/docs/%D0%9F%D1%80%D0%B8%D0%BA%D0%BB%D0%B0%D0%B4%20%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%87%D0%BA%D0%B8%20%D0%BE%D1%82%20%D0%BA%D0%BB%D1%96%D1%96%D0%BD%D1%82%D0%B0.xlsx).
3. **UX Web-форми**: Єдина структурована Streamlit-форма (`st.form`) з тематичними секціями (Загальні дані, Кандидат, Умови та графік, Оплата, Обов'язки та коментарі).
4. **Антиспам**: Сесійний лічильник спроб надсилання у `st.session_state`.
5. **Тестування**: 100% покриття pytest-тестами всіх 9 обов'язкових сценаріїв.

---

## 3. Запропоновані зміни та структура файлів

### [Структура репозиторію]

#### [NEW] [assets/client_template.xlsx](file:///d:/!work/smart-solutions/Automation/assets/client_template.xlsx)
Основа шаблону з 27 колонками для генерації реєстру.

#### [NEW] [.streamlit/secrets.toml.template](file:///d:/!work/smart-solutions/Automation/.streamlit/secrets.toml.template)
Шаблон конфігурації секретів для роутингу, ключів та пошти.

#### [NEW] [core/schema.py](file:///d:/!work/smart-solutions/Automation/core/schema.py)
Специфікація 27 заголовків `TEMPLATE_HEADERS`, маппінг `HEADER_TO_COL` та функція `validate_template_headers(ws)`.

#### [NEW] [core/security.py](file:///d:/!work/smart-solutions/Automation/core/security.py)
Функції `escape_excel_text`, `normalize_phone`, `validate_ipn`, `validate_iban`, `write_client_value`.

#### [NEW] [core/routes.py](file:///d:/!work/smart-solutions/Automation/core/routes.py)
Функція `resolve_route(route_key)` для розбору параметрів URL з `st.secrets`.

#### [NEW] [core/export_print.py](file:///d:/!work/smart-solutions/Automation/core/export_print.py)
Генерація листа `Форма_Печати`, формульні посилання, розрахунок `set_merged_row_height`, захист листа та налаштування А4.

#### [NEW] [core/export_instruction.py](file:///d:/!work/smart-solutions/Automation/core/export_instruction.py)
Генерація листа `Інструкція` з 6 регламентними пунктами для оператора.

#### [NEW] [core/export.py](file:///d:/!work/smart-solutions/Automation/core/export.py)
Головний диспетчер генерації `.xlsx` книги в пам'яті (`build_workbook`).

#### [NEW] [core/mailer.py](file:///d:/!work/smart-solutions/Automation/core/mailer.py)
Клас `OutstaffingMailer`, протокол `MailProviderProtocol`, консольний `MockMailProvider` та `SMTPMailProvider`.

#### [NEW] [client_app.py](file:///d:/!work/smart-solutions/Automation/client_app.py)
Основний веб-інтерфейс Streamlit з розділами форми, обробкою `r` параметра, submit-кнопкою, генерацією назв файлів та fallback завантаженням.

#### [NEW] [requirements.txt](file:///d:/!work/smart-solutions/Automation/requirements.txt)
Залежності проекту (`streamlit`, `openpyxl`, `pytest`).

#### [NEW] [tests/test_all_requirements.py](file:///d:/!work/smart-solutions/Automation/tests/test_all_requirements.py)
Тестовий комплект pytest для 9 обов'язкових перевірок з Розділу 15 ТЗ.

---

## 4. План перевірки (Verification Plan)

### Автоматизовані тести
Запуск pytest:
```bash
pytest tests/
```
Перевірятиме:
1. Тест 1: Перевірка 27 колонок у `TEMPLATE_HEADERS`.
2. Тест 2: Перевірка комірки `A2` (жовта заливка + коментар).
3. Тест 3: Екранування Excel Injection (`=`, `+`, `-`, `@`).
4. Тест 4: Зв'язок формул друкованої форми.
5. Тест 5: Розрахунок `set_merged_row_height`.
6. Тест 6: Ізоляція захисту листа (`locked=False` для C5).
7. Тест 7: Успішна відправка поштою та приховування download_button.
8. Тест 8: Fallback сценарій з появою download_button.
9. Тест 9: Блокування невалідного `route_key`.

### Ручне тестування
- Запуск локального Streamlit сервера `streamlit run client_app.py --server.port 8501`.
- Перевірка заповнення форми та завантаження згенерованого `.xlsx` у fallback режимі.
