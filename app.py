import PyPDF2
import smtplib
import pandas as pd
from io import BytesIO
from flask_cors import CORS
from reportlab.pdfgen import canvas
from email.mime.text import MIMEText
from reportlab.lib.pagesizes import letter
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from flask import Flask, request, jsonify, make_response

app = Flask(__name__)
CORS(app)

def verify_name(username):
    try:
        df = pd.read_excel('students.xlsx')

        if 'Nome' in df.columns:
            partes_nome = username.strip().lower().split()
            if len(partes_nome) < 2:
                print('Por favor, insira um nome válido com pelo menos um primeiro e último nome.', 'error')
                return False
            
            primeiro_nome = partes_nome[0]
            ultimo_nome = partes_nome[-1]
            
            nome_formatado = f"{primeiro_nome} {ultimo_nome}"
            print("Nome formatado para verificação:", nome_formatado)

            df_cleaned = df['Nome'].str.strip().str.lower()

            print("Nomes no Excel:", df_cleaned.values)

            if nome_formatado in df_cleaned.values:
                return True
            else:
                print('Seu nome não faz parte dos estudantes do CAF.', 'error')
                return False
        else:
            print('Coluna "Nome" não encontrada no arquivo Excel.', 'error')
            return False
    except Exception as e:
        print(f'Erro ao ler o arquivo Excel: {e}', 'danger')
        return False    

def format_full_name(username):
    partes_nome = username.split()
    
    if len(partes_nome) < 2:
        return None
    
    primeiro_nome = partes_nome[0]
    ultimo_nome = partes_nome[-1]

    iniciais_nomes_meio = " ".join([parte[0] + '.' for parte in partes_nome[1:-1]])

    if iniciais_nomes_meio:
        nome_formatado = f"{primeiro_nome} {iniciais_nomes_meio} {ultimo_nome}"
    else:
        nome_formatado = f"{primeiro_nome} {ultimo_nome}"

    return nome_formatado

def send_email_with_pdf(username, email, pdf_content):
    try:
        sender_email = "omarscode007@gmail.com"
        sender_password = 'onys bhbi rgou kmlu'
        subject = 'Seu Certificado de Participação'
        body = f"Olá {username},\n\nAqui está o seu certificado de participação."

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

@app.route('/', methods=['POST'])
def index():
    if request.method == 'POST':
        user_data = request.get_json()
        username = user_data.get('name')
        phone = user_data.get('number')
        email = user_data.get('email')


        if not username or not phone or not email:
            return jsonify({'error': 'Todos os campos devem ser preenchidos.'}), 400
        else:
            if verify_name(username):
                print(f'Certificado adquirido com sucesso., {email}', 'success')
            else:
                print('Nome não encontrado na lista dos estudantes do CAF.', 'error')
                return jsonify({'success': False, 'message': 'Nome não encontrado na lista dos estudantes do CAF.'}), 404

        with open("./Index.pdf", "rb") as f:
            pdf_bytes = f.read()

            pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
            pdf_writer = PyPDF2.PdfWriter()

            page = pdf_reader.pages[0]

            nome_certificado = format_full_name(username)

            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)

            can.setFont("Helvetica", 35)
            can.drawString(50, 350, nome_certificado)
            can.save()

            packet.seek(0)

            new_pdf = PyPDF2.PdfReader(packet)

            page.merge_page(new_pdf.pages[0])
            pdf_writer.add_page(page)

            final_pdf = BytesIO()
            pdf_writer.write(final_pdf)
            final_pdf.seek(0)

            send_email_with_pdf(username, email, final_pdf.read())

            return jsonify({'success': True, 'message': 'Certificado enviado por e-mail com sucesso!'}), 201

    return make_response(jsonify({
        "detail": "doc created"
    }),
    201)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5032)
