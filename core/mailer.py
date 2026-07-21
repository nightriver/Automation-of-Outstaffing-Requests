"""
Поштовий модуль відправки Excel-пакета електронною поштою з абстракцією через MailProviderProtocol.
"""

from email.message import EmailMessage
import logging
import smtplib
from typing import Protocol

logger = logging.getLogger(__name__)


class MailProviderProtocol(Protocol):
    def send_email(self, msg: EmailMessage) -> bool:
        ...


class MockMailProvider:
    """
    Mock-провайдер для локальної розробки та тестування.
    """

    def __init__(self, should_succeed: bool = True):
        self.should_succeed = should_succeed
        self.sent_messages: list[EmailMessage] = []

    def send_email(self, msg: EmailMessage) -> bool:
        if self.should_succeed:
            self.sent_messages.append(msg)
            logger.info(f"[MockMailProvider] Успішно імітовано відправку листа до {msg['To']}")
            return True
        else:
            logger.warning("[MockMailProvider] Імітація помилки відправки для тестування fallback-режиму")
            return False


class SMTPMailProvider:
    """
    Стандартний SMTP-провайдер надсилання електронних листів.
    """

    def __init__(self, host: str, port: int, username: str, password: str, use_tls: bool = True):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.use_tls = use_tls

    def send_email(self, msg: EmailMessage) -> bool:
        try:
            with smtplib.SMTP(self.host, self.port, timeout=10) as server:
                if self.use_tls:
                    server.starttls()
                if self.username and self.password:
                    server.login(self.username, self.password)
                server.send_message(msg)
            return True
        except Exception as e:
            logger.error(f"[SMTPMailProvider] Помилка надсилання листа: {e}")
            return False


class OutstaffingMailer:
    """
    Головний сервіс надсилання заявок на пошту менеджера.
    """

    def __init__(self, provider: MailProviderProtocol, from_addr: str):
        self.provider = provider
        self.from_addr = from_addr

    def send_document_package(
        self, file_bytes: bytes, filename: str, to_addr: str, submission_id: str
    ) -> bool:
        """
        Формує EmailMessage та передає обраному поштовому провайдеру.
        """
        msg = EmailMessage()
        msg["Subject"] = f"Smart Solutions | Нова заявка на аутстаффінг [ID: {submission_id}]"
        msg["From"] = self.from_addr
        msg["To"] = to_addr

        msg.set_content(
            "Отримано нову заявку на оформлення співробітника.\n\n"
            f"Номер звернення (Submission ID): {submission_id}\n\n"
            "Файл згенеровано автоматично та додано до вкладення."
        )

        msg.add_attachment(
            file_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

        return self.provider.send_email(msg)
