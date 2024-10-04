# app.py
from flask import Flask, render_template, request, send_file
from docx import Document
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_certificate', methods=['POST'])
def generate_certificate():
    name = request.form['name']
    last_name = request.form['last_name']
    birth_date = request.form['birth_date']
    level = request.form['level']
    registration_number = request.form['reg_number']

    # تحميل قالب الشهادة
    doc = Document('template.docx')
    
    # استبدال القيم داخل المستند
    for para in doc.paragraphs:
        if 'FIRST_NAME' in para.text:
            para.text = para.text.replace('FIRST_NAME', name)
        if 'LAST_NAME' in para.text:
            para.text = para.text.replace('LAST_NAME', last_name)
        if 'BIRTH_DATE' in para.text:
            para.text = para.text.replace('BIRTH_DATE', birth_date)
        if 'LEVEL' in para.text:
            para.text = para.text.replace('LEVEL', level)
        if 'REGISTRATION_NUMBER' in para.text:
            para.text = para.text.replace('REGISTRATION_NUMBER', registration_number)

    # حفظ المستند في الذاكرة باستخدام io.BytesIO
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)  # إعادة المؤشر إلى بداية الملف

    # إرسال الملف مباشرة إلى المتصفح لتحميله
    return send_file(file_stream, as_attachment=True, download_name=f'certificate_{name}_{last_name}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    app.run(debug=True)
