import PyPDF2
import pandas as pd
from io import BytesIO
from flask_cors import CORS
from reportlab.pdfgen import canvas
from flask_mail import Message, Mail
from reportlab.lib.pagesizes import letter
from flask import Flask, render_template, redirect, url_for, request, flash, jsonify, make_response

app = Flask(__name__)

app.config['SECRET_KEY'] = 'djnjendfu31nd'
app.config['MAIL_SERVER'] = 'smtp.gmail.com'  
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'omarscode007@gmail.com'  
app.config['MAIL_PASSWORD'] = 'onys bhbi rgou kmlu' 
app.config['MAIL_DEFAULT_SENDER'] = 'omarscode007@gmail.com'

mail = Mail(app)
CORS(app)

def verify_name(username):
    try:
        df = pd.read_excel('students.xlsx')

        if 'Nome' in df.columns:
            partes_nome = username.strip().lower().split()
            if len(partes_nome) < 2:
                flash('Por favor, insira um nome válido com pelo menos um primeiro e último nome.', 'error')
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
                print('Certificado adquirido com sucesso.', 'success')
                return jsonify({'success': True, 'message': 'Certificado gerado com sucesso!'}), 201
            else:
                print('Nome não encontrado na lista dos estudantes do CAF.', 'error')
                return jsonify({'success': False, 'message': 'Nome não encontrado na lista dos estudantes do CAF.'}), 404

    with open("./White Gold Elegant Modern Certificate of Participation.pdf", "rb") as f:
        pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()

        page = pdf_reader.pages[0]

        nome_certificado = format_full_name(username)

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        can.setFont("Helvetica", 50)
        can.drawString(200, 300, nome_certificado)
        can.save()

        packet.seek(0)

        new_pdf = PyPDF2.PdfReader(packet)

        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        final_pdf = BytesIO()
        pdf_writer.write(final_pdf)
        final_pdf.seek(0)

        msg = Message('Seu Certificado de Participação',
                      recipients=[email])

        msg.body = f"Olá {username},\n\nAqui está o seu certificado de participação."
        
        msg.attach('certificado.pdf', 'application/pdf', final_pdf.read())

        try:
            mail.send(msg)
            print('E-mail enviado com sucesso!')
        except Exception as e:
            print(f'Falha ao enviar o e-mail: {str(e)}', 'error')

    return make_response(jsonify({
        "detail": "doc created"
    }),
    201)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5032)
