import PyPDF2
from io import BytesIO
from flask_cors import CORS
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from send_email import send_email_with_pdf
from flask import Flask, request, jsonify, make_response

app = Flask(__name__)
app.config['SECRET_KEY'] = 'scodilson'
CORS(app)

pdfmetrics.registerFont(TTFont('GreatVibes', 'fonts/GreatVibes-Regular.ttf'))

pdf_path = None
font_name = "GreatVibes"
font_size = 40
x, y = 300, 325
fill_color = (1, 1, 1)  
certificate_type = None

@app.post('/best-project')
def best_project():
    user_data = request.get_json()
    username = user_data.get('name')
    classe = user_data.get('classe')
    curso = user_data.get('curso')
    email = user_data.get('email')

    if not username or not email or not classe or not curso:
        return jsonify({'error': 'Todos os campos devem ser preenchidos.'}), 400

    if curso not in ['Informática', 'Eletrônica']:
        return jsonify({'error': 'Curso inválido. Deve ser Informática ou Eletrônica.'}), 400

    if curso == 'Informática':
        if classe == '10ª':
            pdf_path = "certificates/projects/bestProjectInfo10.pdf"
            certificate_type = 'Melhor projeto'
        elif classe == '11ª':
            pdf_path = "certificates/projects/bestProjectInfo11.pdf"
            certificate_type = 'Melhor projeto'
        elif classe == '12ª':
            pdf_path = "certificates/projects/bestProject12.pdf"
            certificate_type = 'Melhor projeto'
        else:
            return jsonify({'error': 'Classe inválida para Informática.'}), 400

    elif curso == 'Eletrônica':
        if classe == '10ª':
            pdf_path = "certificates/projects/bestProjectEletro10.pdf"
            certificate_type = 'Melhor projeto'
        elif classe == '11ª':
            pdf_path = "certificates/projects/bestProjectEletro11.pdf"
            certificate_type = 'Melhor projeto'
        elif classe == '12ª':
            pdf_path = "certificates/projects/bestProject12.pdf"
            certificate_type = 'Melhor projeto'
        else:
            return jsonify({'error': 'Classe inválida para Eletrônica.'}), 400

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()
        page = pdf_reader.pages[0]

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont(font_name, font_size)

        if fill_color:
            can.setFillColorRGB(*fill_color)

        can.drawString(x, y, username)
        can.save()

        packet.seek(0)
        new_pdf = PyPDF2.PdfReader(packet)
        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        final_pdf = BytesIO()
        pdf_writer.write(final_pdf)
        final_pdf.seek(0)

        send_email_with_pdf(username, email, final_pdf.read(), certificate_type=certificate_type)

    return jsonify({'success': True, 'message': 'Certificado enviado por e-mail com sucesso!'}), 201

@app.post('/best-expo')
def best_Expo():
    user_data = request.get_json()
    username = user_data.get('name')
    email = user_data.get('email')

    if not username or not email:
        return jsonify({'error': 'Todos os campos devem ser preenchidos.'}), 400

    with open("certificates/expositors/bestExpo.pdf", "rb") as f:
        pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()

        page = pdf_reader.pages[0]

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont(font_name, font_size)

        if fill_color:
            can.setFillColorRGB(*fill_color)

        can.drawString(x, y, username)
        can.save()

        packet.seek(0)

        new_pdf = PyPDF2.PdfReader(packet)

        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        final_pdf = BytesIO()
        pdf_writer.write(final_pdf)
        final_pdf.seek(0)

        send_email_with_pdf(username, email, final_pdf.read(), certificate_type='Melhor Expositor')

        return jsonify({'success': True, 'message': 'Certificado enviado por e-mail com sucesso!'}), 201

    return make_response(jsonify({"detail": "doc created"}), 201)

@app.post('/best-stand')
def best_stand():
    user_data = request.get_json()
    username = user_data.get('name')
    email = user_data.get('email')

    with open("certificates/stand/bestStand.pdf", "rb") as f:
        pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()

        page = pdf_reader.pages[0]

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        can.setFont(font_name, font_size)

        if fill_color:
            can.setFillColorRGB(*fill_color)

        can.drawString(x, y, username)
        can.save()

        packet.seek(0)

        new_pdf = PyPDF2.PdfReader(packet)

        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        final_pdf = BytesIO()
        pdf_writer.write(final_pdf)
        final_pdf.seek(0)

        send_email_with_pdf(username, email, final_pdf.read(), certificate_type='Melhor Stand')

        return jsonify({'success': True, 'message': 'Certificado enviado por e-mail com sucesso!'}), 201
    
@app.post('/participants')
def partcipants():
    user_data = request.get_json()
    username = user_data.get('name')
    email = user_data.get('email')

    with open("certificates/stand/bestStand.pdf", "rb") as f:
        pdf_bytes = f.read()

        pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_bytes))
        pdf_writer = PyPDF2.PdfWriter()

        page = pdf_reader.pages[0]

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        can.setFont(font_name, font_size)

        if fill_color:
            can.setFillColorRGB(*fill_color)

        can.drawString(x, y, username)
        can.save()

        packet.seek(0)

        new_pdf = PyPDF2.PdfReader(packet)

        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        final_pdf = BytesIO()
        pdf_writer.write(final_pdf)
        final_pdf.seek(0)

        send_email_with_pdf(username, email, final_pdf.read(), certificate_type='ParticipanteC')

        return jsonify({'success': True, 'message': 'Certificado enviado por e-mail com sucesso!'}), 201


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5032)
