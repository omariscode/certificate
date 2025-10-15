import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_email_with_pdf(username, email, pdf_content, certificate_type):
    try:
        sender_email = "omarscode007@gmail.com"
        sender_password = 'onys bhbi rgou kmlu'
        subject = f'Certificado - {certificate_type}'
        body = f"Olá {username},\n\nAqui está o seu certificado de {certificate_type}\n\nParabéns!"

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        attachment = MIMEApplication(pdf_content, _subtype="pdf")
        attachment.add_header('Content-Disposition', 'attachment', filename='certificado.pdf')
        msg.attach(attachment)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Erro ao enviar e-mail: {e}')