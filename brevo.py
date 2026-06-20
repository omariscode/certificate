import os
import base64
from dotenv import load_dotenv
import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException
from flask import render_template

load_dotenv()


def get_brevo_api():
    api_key = (os.getenv("BREVO_API_KEY") or "").strip()
    if not api_key:
        raise ValueError("BREVO_API_KEY não configurada no .env")

    config = sib_api_v3_sdk.Configuration()
    config.api_key["api-key"] = api_key
    client = sib_api_v3_sdk.ApiClient(config)
    return sib_api_v3_sdk.TransactionalEmailsApi(client)


def send_certificate_email(username, email, pdf_content, certificate_type):
    api = get_brevo_api()

    html = render_template(
        "email_certificate.html",
        username=username,
        certificate_type=certificate_type,
    )

    pdf_b64 = base64.b64encode(pdf_content).decode("utf-8") if isinstance(pdf_content, bytes) else pdf_content

    email_data = sib_api_v3_sdk.SendSmtpEmail(
        to=[{"email": email, "name": username}],
        sender={
            "email": "omarscode007@gmail.com",
            "name": "Academia Metanoia",
        },
        subject=f"Certificado - {certificate_type}",
        html_content=html,
        attachment=[
            {
                "name": "certificado.pdf",
                "content": pdf_b64,
            }
        ],
    )

    try:
        api.send_transac_email(email_data)
        print(f"E-mail enviado com sucesso para {email}")
    except ApiException as exc:
        if exc.status == 401:
            raise ValueError(
                "Brevo API key não autorizada. Gera uma nova chave e atualiza o .env."
            ) from exc
        raise


def send_verification_email(email, code):
    api = get_brevo_api()

    html = render_template("email_verify.html", {"code": code})

    email_data = sib_api_v3_sdk.SendSmtpEmail(
        to=[{"email": email}],
        sender={
            "email": "omarscode007@gmail.com",
            "name": "Academia Metanoia",
        },
        subject="Verify Email",
        html_content=html,
    )

    try:
        api.send_transac_email(email_data)
    except ApiException as exc:
        if exc.status == 401:
            raise ValueError(
                "Brevo API key não autorizada. Gera uma nova chave e atualiza o .env."
            ) from exc
        raise
