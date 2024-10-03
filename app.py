import PyPDF2
import pandas as pd
from io import BytesIO
from flask_cors import CORS
from reportlab.pdfgen import canvas
from flask_mail import Message, Mail
from reportlab.lib.pagesizes import letter
from flask import Flask, render_template, redirect, url_for, request, flash, get_flashed_messages

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
            name = username.strip().lower()
            df_cleaned = df['Nome'].str.strip().str.lower()

            print("Username fornecido:", name)
            print("Nomes no Excel:", df_cleaned.values)

            if name in df_cleaned.values:
                return True
            else:
                flash('Seu nome não faz parte dos estudantes do CAF.', 'error')
                return False
        else:
            flash('Coluna "Nome" não encontrada no arquivo Excel.', 'error')
            return False
    except Exception as e:
        flash(f'Erro ao ler o arquivo Excel: {e}', 'danger')
        return False    

@app.route('/', methods=['POST', 'GET'])
def index():
    if request.method == 'POST':
        username = request.form.get('username')
        phone = request.form.get('phone_number')
        email = request.form.get('email')

        print(f"Nome fornecido: {username}, Telefone: {phone}, Email: {email}")

        if not username or not phone or not email:
            flash('Todos os campos devem ser preenchidos.', 'error')
            return redirect(url_for('index'))  
        else:
            if verify_name(username):  
                flash('Certificado adquirido com sucesso.', 'success')
            else:
                flash('Nome não encontrado na lista dos estudantes do CAF.', 'error')
                return redirect(url_for('index'))

        with open("./White Gold Elegant Modern Certificate of Participation.pdf", "rb") as f:
            pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()

        page = pdf_reader.pages[0]

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        can.setFont("Helvetica", 50)
        can.drawString(200, 300, username)
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
            flash('E-mail enviado com sucesso!')
        except Exception as e:
            flash(f'Falha ao enviar o e-mail: {str(e)}', 'error')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5032)