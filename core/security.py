"""
Модуль санітизації, екранування від Excel Injection та валідації введених даних.
"""

import re

FORMULA_PREFIXES = ("=", "+", "-", "@")


def escape_excel_text(value):
    """
    Запобігає Excel Injection. Якщо рядок починається з формульного префікса (=, +, -, @),
    додається апостроф перед значенням.
    """
    if value is None:
        return ""
    value = str(value)
    return "'" + value if value.lstrip().startswith(FORMULA_PREFIXES) else value


def normalize_phone(value: str) -> str:
    """
    Нормалізує номер телефону до формату +380XXXXXXXXX.
    """
    digits = re.sub(r"\D", "", value or "")
    if digits.startswith("380") and len(digits) == 12:
        return "+" + digits
    if digits.startswith("0") and len(digits) == 10:
        return "+38" + digits
    raise ValueError("Некоректний формат номера телефону. Формат: +380XXXXXXXXX")


def validate_ipn(value: str) -> str:
    """
    Перевіряє, що ІПН містить рівно 10 цифр.
    """
    digits = re.sub(r"\D", "", value or "")
    if not re.fullmatch(r"\d{10}", digits):
        raise ValueError("ІПН має містити рівно 10 цифр")
    return digits


def validate_iban(value: str) -> str:
    """
    Перевіряє формат IBAN (UA + 27 цифр).
    """
    normalized = re.sub(r"\s+", "", (value or "").upper())
    if not re.fullmatch(r"UA\d{27}", normalized):
        raise ValueError("Некоректний формат IBAN (має бути UA + 27 цифр)")
    return normalized


def write_client_value(ws, row: int, col: int, value) -> None:
    """
    Безапеляційно та безпечно записує значення клієнта в комірку Excel.
    Дати, числа та булеві значення записуються як рідні типи Python; строки екрануються.
    """
    ws.cell(row=row, column=col).value = escape_excel_text(value) if isinstance(value, str) else value
