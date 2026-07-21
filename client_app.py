"""
Головний вхідний файл застосунку Streamlit (client_app.py)
З нативними st.toast() сповіщеннями (тривалість 10 сек) без передчасного st.rerun(),
що гарантує відображення виринаючих підказок та червоних рамок полів.
"""

from datetime import date, datetime
import logging
import os
from uuid import uuid4
import streamlit as st

from core.export import build_workbook
from core.mailer import MockMailProvider, OutstaffingMailer, SMTPMailProvider
from core.routes import resolve_route
from core.security import normalize_phone, validate_ipn  # validate_iban тимчасово виключено

logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="Smart Solutions — Заявка на аутстаффінг",
    page_icon="📋",
    layout="centered",
)

# Стилізація нативного st.toast() з тривалістю 10 секунд та червоних рамок помилок
st.markdown(
    """
    <style>
    /* Головний контейнер з падінгом 3rem */
    div[data-testid="stMainBlockContainer"], .stMainBlockContainer {
        padding: 3rem !important;
    }

    /* Головний фон сторінки */
    .stApp {
        background-color: #FFFFFF;
        color: #1A1A1A;
        font-family: 'Arial', sans-serif;
    }

    /* Нативна стилізація блоків st.toast() з тривалістю 10 секунд */
    div[data-testid="stToast"] {
        border-radius: 8px !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 14px rgba(0,0,0,0.15) !important;
        animation: stToastHold 10s ease-in-out forwards !important;
    }

    @keyframes stToastHold {
        0% { opacity: 0; transform: translateY(15px); }
        5% { opacity: 1; transform: translateY(0); }
        92% { opacity: 1; transform: translateY(0); }
        100% { opacity: 0; transform: translateY(-10px); }
    }

    /* Приховуємо маркерні посилання */
    div[data-testid="stToast"] a[href*="invalid"] {
        display: none !important;
    }

    /* Блідо-червоний для помилок (посилання https://error.invalid) */
    div[data-testid="stToast"]:has(a[href="https://error.invalid"]) {
        background-color: #F8D7DA !important;
        color: #721C24 !important;
        border: 1px solid #F5C6CB !important;
    }

    /* Блідо-зелений колір для успішного toast (посилання https://success.invalid) */
    div[data-testid="stToast"]:has(a[href="https://success.invalid"]) {
        background-color: #D4EDDA !important;
        color: #155724 !important;
        border: 1px solid #C3E6CB !important;
    }

    /* Зробити колір плейсхолдерів світлішим */
    input::placeholder, textarea::placeholder {
        color: #A9A9A9 !important;
        opacity: 0.85 !important;
    }



    /* Картка/секції форми з сірим фоном #F3F3F3 та жовтим акцентним бордером #FFD100 */
    div[data-testid="stForm"] {
        background-color: #F3F3F3 !important;
        border-left: 6px solid #FFD100 !important;
        border-radius: 8px !important;
        padding: 24px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }

    /* Секційні заголовки з відступом 0.8rem до полів */
    .stForm h3, h3 {
        color: #1A1A1A !important;
        border-bottom: 2px solid #FFD100 !important;
        padding-bottom: 4px !important;
        margin-top: 18px !important;
        margin-bottom: 0.8rem !important;
    }

    /* 1. Dropdown selectbox: білий фон для всього контейнера включаючи стрілку */
    div[data-testid="stSelectbox"] div[data-baseweb="select"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    div[data-baseweb="select"],
    div[data-baseweb="select"] * {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
    }

    /* 2. Усунення подвійної рамки навколо DateInput */
    div[data-testid="stDateInput"] div[data-baseweb="input"] {
        border: none !important;
        background-color: transparent !important;
        box-shadow: none !important;
    }

    div[data-baseweb="base-input"] {
        border: none !important;
        background-color: transparent !important;
    }

    /* Єдиний стилізований контейнер з рамкою для всіх полів вводу */
    div[data-testid="stTextInput"] > div > div,
    div[data-testid="stTextArea"] > div > div,
    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stDateInput"] > div > div,
    div[data-baseweb="input"],
    div[data-baseweb="textarea"] {
        background-color: #FFFFFF !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 6px !important;
    }

    input, textarea {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
    }

    /* Приховання підказок "Press Enter to submit form" */
    div[data-testid="InputInstructions"], 
    small[data-testid="stFormSubmitButtonInstructions"],
    .stInputInstructions,
    div[data-baseweb="input"] + small {
        display: none !important;
    }

    /* Червона CTA кнопка відправки (#E8312A) */
    div[data-testid="stFormSubmitButton"] button {
        background-color: #E8312A !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border-radius: 6px !important;
        padding: 12px 28px !important;
        width: 100% !important;
        cursor: pointer;
        transition: background-color 0.2s ease;
    }
    div[data-testid="stFormSubmitButton"] button:hover {
        background-color: #C62822 !important;
    }

    /* Кнопка завантаження Excel (контурна) */
    div[data-testid="stDownloadButton"] button {
        background-color: transparent !important;
        color: #1A1A1A !important;
        border: 2px solid #1A1A1A !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        border-radius: 6px !important;
        padding: 10px 22px !important;
        width: 100% !important;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #FFD100 !important;
        border-color: #FFD100 !important;
        color: #1A1A1A !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 1. Розрізнення маршруту за r параметром в URL (тимчасово вимкнено обов'язковість для тестів)
query_params = st.query_params
route_key = query_params.get("r", "demo_route_123")

try:
    route = resolve_route(route_key)
except ValueError:
    route = {
        "route_id": "TEST-ROUTE",
        "email": "nightriver@gmail.com",
        "client_template": {}
    }

# 2. Ініціалізація поштового сервісу
smtp_secrets = st.secrets.get("smtp", {})
use_mock = smtp_secrets.get("use_mock", True)
provider_type = smtp_secrets.get("provider", "")

if use_mock or provider_type == "mock":
    provider = MockMailProvider(should_succeed=True)
elif provider_type == "brevo" or (provider_type == "" and smtp_secrets.get("api_key")):
    from core.mailer import BrevoMailProvider
    provider = BrevoMailProvider(api_key=smtp_secrets.get("api_key", ""))
else:
    provider = SMTPMailProvider(
        host=smtp_secrets.get("host", ""),
        port=int(smtp_secrets.get("port", 587)),
        username=smtp_secrets.get("username", ""),
        password=smtp_secrets.get("password", ""),
        use_tls=bool(smtp_secrets.get("use_tls", True)),
    )

mailer = OutstaffingMailer(
    provider=provider,
    from_addr=smtp_secrets.get("from_addr", "notifications@smart-solutions.ua"),
)

# 3. Антиспам rate-limit у session_state
if "submit_count" not in st.session_state:
    st.session_state["submit_count"] = 0

if "field_errors" not in st.session_state:
    st.session_state["field_errors"] = []

# Динамічне виділення рамок полів з помилкою у червоний колір (#E8312A)
if st.session_state["field_errors"]:
    err_rules = []
    if "pib" in st.session_state["field_errors"]:
        err_rules.append('div[data-testid="stTextInput"]:has(input[aria-label*="П.І.Б."]) > div > div { border: 2px solid #E8312A !important; box-shadow: 0 0 6px rgba(232,49,42,0.3) !important; }')
    if "phone" in st.session_state["field_errors"]:
        err_rules.append('div[data-testid="stTextInput"]:has(input[aria-label*="телефон"]) > div > div { border: 2px solid #E8312A !important; box-shadow: 0 0 6px rgba(232,49,42,0.3) !important; }')
    if "ipn" in st.session_state["field_errors"]:
        err_rules.append('div[data-testid="stTextInput"]:has(input[aria-label*="ІПН"]) > div > div { border: 2px solid #E8312A !important; box-shadow: 0 0 6px rgba(232,49,42,0.3) !important; }')
    if err_rules:
        st.markdown(f"<style>{' '.join(err_rules)}</style>", unsafe_allow_html=True)

# Відображення логотипу Smart Solutions
logo_path = os.path.join("assets", "logo.svg")
if os.path.exists(logo_path):
    st.image(logo_path, width=220)

st.title("📋 Заявка на оформлення співробітника (Аутстаффінг)")

if st.session_state["submit_count"] >= 5:
    st.warning("Перевищено ліміт запитів для цієї сесії. Спробуйте пізніше або зверніться до менеджера.")
    st.stop()


def make_filename(pib: str) -> str:
    safe_pib = "_".join((pib or "Без_ПІБ").split())
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Smart_Solutions_Заявка_Нова_{safe_pib}_{ts}.xlsx"


with st.form("outstaffing_form"):
    st.subheader("1. Загальні відомості про Замовника та заявку")
    col0a, col0b = st.columns(2)
    with col0a:
        client_name_input = st.text_input("Назва компанії Замовника*", placeholder="")
    with col0b:
        client_director_input = st.text_input("П.І.Б. Директора Замовника*", placeholder="")

    col1, col2 = st.columns(2)
    with col1:
        city = st.text_input("Місто*", placeholder="")
        format_work = st.selectbox("Формат роботи*", ["Офіс", "Віддалено", "Гібрид"])
        start_date = st.date_input("Дата прийому*", value=date.today())
    with col2:
        address = st.text_input("Адреса / місце роботи*", placeholder="")
        term_std = st.text_input("Термін дії СТД / проєкту*", placeholder="напр. 31.12.2027 або до завершення Заявки")

    st.subheader("2. Відомості про кандидата")
    col3, col4 = st.columns(2)
    with col3:
        pib = st.text_input("П.І.Б. кандидата*", placeholder="")
        ipn = st.text_input("ІПН (10 цифр)*", placeholder="")
        phone = st.text_input("Номер телефону*", placeholder="+380 XXX-XX-XX")
    with col4:
        email = st.text_input("E-mail кандидата*", placeholder="")
        iban = st.text_input("Реквізити картки / IBAN", placeholder="")

    st.subheader("3. Умови праці та графік")
    col5, col6 = st.columns(2)
    with col5:
        position = st.text_input("Посада згідно КП*", placeholder="")
        probation = st.text_input("Випробувальний термін", value="3 місяці")
        work_type = st.selectbox("Тип роботи*", ["Повна зайнятість", "Часткова зайнятість"])
        mat_resp = st.selectbox("Матеріальна відповідальність*", ["Ні", "Так"])
    with col6:
        schedule_hours = st.text_input("Графік роботи (з-по)*", value="09:00 - 18:00")
        hours_count = st.text_input("Кількість годин на тиждень*", value="40")
        vacation_days = st.text_input("Кількість днів відпустки*", value="24")

    st.subheader("4. Оплата праці та компенсації")
    col7, col8 = st.columns(2)
    with col7:
        gross_salary = st.text_input("Оклад гросс (грн)*", placeholder="")
        bonus = st.text_input("Бонус", placeholder="")
        bonus_period = st.selectbox("Періодичність бонуса", ["Щомісячно", "Щоквартально", "Щорічно", "Без бонуса"])
    with col8:
        advance_salary_date = st.text_input("Дата виплати авансу та ЗП*", value="15 та 30 числа")
        other_comp = st.text_input("Інші компенсації", placeholder="Мобільний зв'язок, медогляд")

    st.subheader("5. Функціональні обов'язки та примітки")
    duties = st.text_area("Функціональні обов'язки*", placeholder="1. Проведення переговорів\n2. Підготовка звітів")
    notes = st.text_area("Примітка / Вимоги до кандидатів", placeholder="Наявність водійського посвідчення категорії B")
    comments = st.text_area("Коментарі")
    documents = st.text_input("Необхідні документи", value="Паспорт, ІПН, Трудова книжка")

    submitted = st.form_submit_button("Надіслати форму")

# Завжди доступний блок завантаження Excel
st.markdown("---")
st.caption("Додаткова дія: збереження згенерованого Excel-файлу на ваш пристрій")

location = ", ".join(part for part in [city.strip(), address.strip()] if part)
str_start_date = start_date.strftime("%d.%m.%Y") if isinstance(start_date, (date, datetime)) else str(start_date)

form_data = {
    "Замовник": client_name_input.strip(),
    "Директор замовника": client_director_input.strip(),
    "П.І.Б.": pib.strip(),
    "ІПН": ipn.strip(),
    "Дата прийома": str_start_date,
    "Термін дії СТД": term_std.strip(),
    "Посада згідно кп": position.strip(),
    "Випробувальний термін": probation.strip(),
    "Номер телефону": phone.strip(),
    "E-mail": email.strip(),
    "Оклад гросс": gross_salary.strip(),
    "Бонус": bonus.strip(),
    "Періодичність бонуса": bonus_period,
    "Робота": work_type,
    "Дата виплата авансу та ЗП": advance_salary_date.strip(),
    "Графік роботи (з-по)": schedule_hours.strip(),
    "Графік роботи количество годин": hours_count.strip(),
    "Кількість дней відпустки": vacation_days.strip(),
    "Місто та місце роботи": location,
    "Формат роботи": format_work,
    "Матеріальна відповідальність": mat_resp,
    "Інші компенсації": other_comp.strip(),
    "Коментарі": comments.strip(),
    "Реквізити картки": iban.strip(),
    "Функціональні обов'язки": duties.strip(),
    "Примітка": notes.strip(),
    "документи": documents.strip(),
}

filename = make_filename(pib)

# Безпечна генерація масиву байтів із фіксацією помилок у лог
try:
    current_bytes = build_workbook("assets/client_template.xlsx", form_data, route.get("client_template"))
except Exception as ex:
    logger.error(f"Помилка генерації Excel у пам'яті: {ex}", exc_info=True)
    current_bytes = b""

st.download_button(
    "📥 Завантажити Excel",
    data=current_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if submitted:
    st.session_state["submit_count"] += 1

    validation_errors = []
    field_errors = []

    if not pib.strip():
        validation_errors.append("Поле П.І.Б. є обов'язковим.")
        field_errors.append("pib")

    try:
        norm_phone = normalize_phone(phone)
        form_data["Номер телефону"] = norm_phone
    except ValueError as e:
        validation_errors.append(str(e))
        field_errors.append("phone")

    try:
        norm_ipn = validate_ipn(ipn)
        form_data["ІПН"] = norm_ipn
    except ValueError as e:
        validation_errors.append(str(e))
        field_errors.append("ipn")

    st.session_state["field_errors"] = field_errors

    if validation_errors:
        for err in validation_errors:
            st.toast(f"⚠️ {err} [err](https://error.invalid)", icon="⚠️")
    else:
        st.session_state["field_errors"] = []
        try:
            submission_id = f"SS-{datetime.now().strftime('%Y%m%d')}-{uuid4().hex[:4].upper()}"
            final_bytes = build_workbook("assets/client_template.xlsx", form_data, route.get("client_template"))

            ok = mailer.send_document_package(
                file_bytes=final_bytes,
                filename=filename,
                to_addr=route["email"],
                submission_id=submission_id,
            )

            if ok:
                st.toast(f"✅ Форму успішно надіслано менеджеру! [ID: {submission_id}] [ok](https://success.invalid)", icon="✅")
            else:
                st.toast("❌ Помилка надсилання поштою. Завантажте файл кнопкою нижче. [err](https://error.invalid)", icon="❌")
        except Exception as ex:
            st.toast("❌ Під час обробки форми виникла технічна помилка. [err](https://error.invalid)", icon="❌")
