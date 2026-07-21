"""
Модуль розрізнення та маршрутизації параметрів URL за допомогою st.secrets.
"""

import streamlit as st


def resolve_route(route_key: str) -> dict:
    """
    Отримує конфігурацію маршруту з st.secrets за роутинг-ключем r.
    """
    if not route_key:
        raise ValueError("Ключ маршруту відсутній")

    routes = st.secrets.get("routes", {})
    route = routes.get(route_key)
    if not route or not route.get("active"):
        raise ValueError("Посилання форми недійсне або деактивоване")

    managers = st.secrets.get("managers", {})
    email_key = route.get("email_secret_key")
    manager_email = managers.get(email_key, "")

    return {
        "route_id": route.get("route_id", "UNKNOWN"),
        "email": manager_email,
        "client_template": route.get("client_template", "standard"),
    }
