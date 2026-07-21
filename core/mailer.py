import base64
from email.message import EmailMessage
import json
import logging
import smtplib
from typing import Protocol
import urllib.error
import urllib.request

logger = logging.getLogger(__name__)


class MailProviderProtocol(Protocol):
    def send_email(self, msg: EmailMessage) -> bool:
        ...


class BrevoMailProvider:
    """
    Поштовий провайдер надсилання листів через Brevo Transactional Emails API v3.
    """

    def __init__(self, api_key: str):
        self.api_key = api_key

    def send_email(self, msg: EmailMessage) -> bool:
        if not self.api_key:
            logger.error("[BrevoMailProvider] API ключ відсутній")
            return False

        # Витягуємо текстовий вміст листа
        text_content = ""
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    text_content = part.get_payload(decode=True).decode("utf-8")
                except Exception:
                    text_content = part.get_payload()
                break

        # Витягуємо вкладені файли (Excel)
        attachments = []
        for part in msg.iter_attachments():
            file_name = part.get_filename()
            file_bytes = part.get_payload(decode=True)
            if file_name and file_bytes:
                base64_content = base64.b64encode(file_bytes).decode("utf-8")
                attachments.append({
                    "content": base64_content,
                    "name": file_name
                })

        url = "https://api.brevo.com/v3/smtp/email"
        headers = {
            "accept": "application/json",
            "api-key": self.api_key,
            "content-type": "application/json"
        }
        
        payload = {
            "sender": {"email": msg["From"]},
            "to": [{"email": msg["To"]}],
            "subject": msg["Subject"],
            "textContent": text_content,
            "attachment": attachments
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST"
        )

        try:
            with urllib.request.urlopen(req, timeout=12) as response:
                res_data = response.read()
                logger.info(f"[BrevoMailProvider] Лист успішно надіслано через Brevo API: {res_data}")
                return True
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8")
            logger.error(f"[BrevoMailProvider] HTTP Error: {e.code} - {err_body}")
            return False
        except Exception as e:
            logger.error(f"[BrevoMailProvider] Connection/API Error: {e}")
            return False



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
