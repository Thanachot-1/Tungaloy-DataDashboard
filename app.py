from flask import Flask, render_template, request, redirect, url_for, session
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # เปลี่ยนเป็นคีย์ลับของคุณ

ALLOWED_EXTENSIONS = {'xlsx'}
excel_cache = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def format_date(raw):
    if isinstance(raw, str):
        digits = ''.join(filter(str.isdigit, raw))
        if len(digits) == 6:
            d = digits[:2]
            m = digits[2:4]
            y = digits[4:6]
            return f"{d}/{m}/{y}"
        elif len(digits) == 5:
            d = digits[0]
            m = digits[1:3]
            y = digits[3:5]
            return f"0{d}/{m}/{y}"
    return raw

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join('uploads', filename)
            os.makedirs('uploads', exist_ok=True)
            file.save(filepath)
            session['uploaded_file'] = filepath
            return redirect(url_for('index'))
        else:
            return render_template('upload.html', error="กรุณาเลือกไฟล์ .xlsx ที่ถูกต้อง")
    return render_template('upload.html', error=None)

@app.route('/index')
def index():
    filepath = session.get('uploaded_file')
    if not filepath or not os.path.exists(filepath):
        return redirect(url_for('upload'))

    # ใช้ cache ถ้ามี
    if filepath not in excel_cache:
        df = pd.read_excel(filepath, header=None, engine='openpyxl')
        excel_cache[filepath] = df
    else:
        df = excel_cache[filepath]

    a_column = df.iloc[3:, 2]
    options = [(i, str(val)) for i, val in a_column.items() if pd.notna(val)]

    query = request.args.get('q', '').strip().lower()
    if query:
        filtered_options = [(i, val) for i, val in options if query in val.lower()]
    else:
        filtered_options = options
    return render_template('index.html', options=filtered_options, query=query)

@app.route('/detail')
def detail():
    filepath = session.get('uploaded_file')
    if not filepath or not os.path.exists(filepath):
        return redirect(url_for('upload'))

    df = pd.read_excel(filepath, header=None, engine='openpyxl')
    selected_index = request.args.get('index')
    selected_text = request.args.get('text')
    
    if not selected_index or not selected_text:
        return "ข้อมูลไม่ครบ"

    try:
        row_index = int(selected_index)
        value_b = df.iloc[row_index, 12]
        value_c = df.iloc[row_index, 1]
        value_d = df.iloc[row_index, 13]
        value_b = format_date(str(value_b))
        value_description2 = df.iloc[row_index, 3]
    except Exception as e:
        return f"เกิดข้อผิดพลาด: {e}"

    return render_template('detail.html', a_text=selected_text, b_text=value_b, c_text=value_c, d_text=value_d, description2_text=value_description2)

if __name__ == '__main__':
    app.run(debug=True)
