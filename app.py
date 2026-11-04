import os
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import io, base64
from flask import Flask, render_template_string, request, redirect, flash, session, send_file
from datetime import datetime
import uuid
import numpy as np
from PIL import Image
from werkzeug.utils import secure_filename

# مكتبات PDF والعربية
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

# مكتبة Word
from docx import Document

app = Flask(__name__)
app.secret_key = 'super_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

data_store = {
    'excel_file': None, 
    'sheets': [], 
    'current_df': None, 
    'filename': '', 
    'current_sheet_name': '',
    'reports': {},
    'cat_cols': [],
    'num_cols': []
}

# --- إعداد الخط العربي لـ PDF ---
FONT_PATH = '/system/fonts/NotoNaskhArabic-Regular.ttf'
try:
    pdfmetrics.registerFont(TTFont('Arabic', FONT_PATH))
except Exception as e:
    print(f"تحذير: لم يتم العثور على الخط العربي. {e}")

def ar(text):
    if pd.isna(text): return ""
    return get_display(arabic_reshaper.reshape(str(text)))

def get_image_base64(path):
    try:
        with open(path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    except:
        return ""

# ==========================================
# دالة المعالجة المسبقة للبيانات
# ==========================================
def process_dataframe(df):
    cat_cols = []
    num_cols = []
    for col in df.columns:
        if pd.isna(col) or str(col).strip() == '' or 'Unnamed' in str(col):
            continue
        series = df[col]
        if isinstance(series, pd.DataFrame):
            continue 
        
        numeric_series = pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce')
        if numeric_series.notna().sum() > (len(series) * 0.5):
            df[col] = numeric_series
            num_cols.append(col)
        else:
            cat_cols.append(col)
    return df, cat_cols, num_cols

def save_report(data, title):
    report_id = str(uuid.uuid4())
    data_store['reports'][report_id] = {'data': data, 'title': title}
    return report_id

# ==========================================
# قالب شاشة الدخول
# ==========================================
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تسجيل الدخول</title>
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; font-family: 'Tajawal', sans-serif; }
        body { margin: 0; padding: 0; background: linear-gradient(135deg, #1e4d33 0%, #0d2618 100%); color: #fff; display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
        .login-container { width: 100%; max-width: 420px; text-align: center; background: rgba(255, 255, 255, 0.08); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 25px; padding: 35px 25px; box-shadow: 0 20px 50px rgba(0,0,0,0.6), inset 0 0 20px rgba(255,255,255,0.1); transform: perspective(1000px) rotateX(2deg); }
        .logo-img { width: 100%; max-width: 240px; margin-bottom: 25px; filter: drop-shadow(0 6px 10px rgba(0,0,0,0.6)); }
        h2 { color: #d4af37; margin: 0; font-size: 24px; text-shadow: 2px 2px 4px #000; }
        p { color: #e0e0e0; font-size: 15px; margin-bottom: 25px; text-shadow: 1px 1px 2px #000; }
        .info-box { border: 1px solid rgba(212, 175, 55, 0.7); border-radius: 15px; padding: 15px; margin-bottom: 25px; background: rgba(0, 0, 0, 0.3); box-shadow: 0 4px 15px rgba(0,0,0,0.4); }
        .info-box h3 { color: #d4af37; margin: 0 0 8px 0; font-size: 16px; }
        .info-box p { margin: 3px 0; color: #fff; font-size: 15px; font-weight: bold; }
        input { width: 100%; padding: 15px; margin-bottom: 15px; border-radius: 25px; border: 2px solid #d4af37; background: rgba(255, 255, 255, 0.95); font-size: 16px; text-align: center; color: #333; outline: none; box-shadow: 0 4px 10px rgba(0,0,0,0.3); }
        .btn { background: #00e676; color: #000; padding: 15px; border: none; border-radius: 10px; width: 100%; cursor: pointer; font-size: 18px; font-weight: bold; box-shadow: 0 4px 15px rgba(0, 230, 118, 0.4); transition: transform 0.2s; }
        .btn:active { transform: scale(0.98); }
        .footer { color: rgba(255,255,255,0.6); font-size: 12px; margin-top: 20px; text-shadow: 1px 1px 2px #000;}
    </style>
</head>
<body>
    <div class="login-container">
        {% if logo_image %}<img src="data:image/png;base64,{{ logo_image }}" class="logo-img" alt="Logo">{% endif %}
        <h2>الوحدة التنفيذية للمشاريع والإنشاءات</h2>
        <p>نظام تحليل إدارة الأداء ودعم القرار</p>
        <div class="info-box">
            <h3>إعداد وتطوير</h3>
            <p>السيد عبد الملك علي عبد الله الحوثي</p>
            <p style="color:#d4af37; direction: ltr;">+967 777 008 785</p>
        </div>
        <form method="POST" action="/login">
            <input type="password" name="password" placeholder="أدخل كلمة المرور (admin)" required>
            <button type="submit" class="btn">← دخول</button>
        </form>
        <div class="footer">© 2026 - جميع الحقوق محفوظة</div>
    </div>
</body>
</html>
"""

# ==========================================
# قالب لوحة التحكم (مع إضافة خيار رفع ملف)
# ==========================================
DASHBOARD_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>لوحة التحكم</title>
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; font-family: 'Tajawal', sans-serif; }
        body { margin: 0; padding-bottom: 20px; background: linear-gradient(rgba(10, 42, 26, 0.85), rgba(10, 42, 26, 0.95)), url('data:image/jpeg;base64,{{ bg_data }}') no-repeat center center fixed; background-size: cover; color: #333; }
        .header { background: rgba(10, 42, 26, 0.85); backdrop-filter: blur(10px); padding: 15px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid rgba(212, 175, 55, 0.7); color: #d4af37; }
        .container { padding: 15px; max-width: 600px; margin: auto; }
        .card { background: rgba(255, 255, 255, 0.95); padding: 15px; border-radius: 15px; margin-bottom: 15px; box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37); display: flex; align-items: center; cursor: pointer; text-decoration: none; color: inherit; border: 1px solid rgba(255,255,255,0.18); transition: transform 0.2s; }
        .card:active { transform: scale(0.98); }
        
        .card-icon { width: 60px; height: 60px; border-radius: 15px; display: flex; justify-content: center; align-items: center; font-size: 30px; margin-left: 15px; }
        .icon-blue { background: #e3f2fd; color: #1976d2; }
        .icon-green { background: #e8f5e9; color: #388e3c; }
        .icon-orange { background: #fff3e0; color: #f57c00; }
        .icon-purple { background: #f3e5f5; color: #7b1fa2; }
        .icon-teal { background: #e0f2f1; color: #00796b; }
        .icon-brown { background: #efebe9; color: #5d4037; }
        .icon-red { background: #ffebee; color: #c62828; }
        .icon-indigo { background: #e8eaf6; color: #283593; }
        .icon-pink { background: #fce4ec; color: #c2185b; }
        
        .card-content h3 { margin: 0; font-size: 17px; color: #0a2a1a; }
        .card-content p { margin: 5px 0 0 0; font-size: 13px; color: #777; }
        
        .file-info { background: rgba(255, 255, 255, 0.95); padding: 15px; border-radius: 10px; margin-bottom: 15px; font-size: 14px; text-align: center; color: #2e7d32; font-weight: bold; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        .form-card { background: rgba(255, 255, 255, 0.95); padding: 20px; border-radius: 15px; margin-bottom: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); }
        
        /* تنسيق حقل رفع الملف */
        input[type="file"] { width: 100%; padding: 10px; margin-bottom: 10px; border: 1px dashed #0a2a1a; border-radius: 8px; background: #f9f9f9; }
        
        .btn-submit { width: 100%; padding: 12px; background: #0a2a1a; color: #d4af37; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; display:block; text-align:center; text-decoration:none; margin-bottom: 10px; }
        .btn-secondary { background: #2e7d32; color: white; }
        
        .loader { display: none; border: 6px solid #f3f3f3; border-top: 6px solid #0a2a1a; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 20px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
    <script>
        function showLoader() {
            document.getElementById('main-form').style.display = 'none';
            document.getElementById('loading-div').style.display = 'block';
        }
    </script>
</head>
<body>
    <div class="header">
        <a href="/logout" style="text-decoration:none; color:#d4af37;">🚪 خروج</a>
        <h2>أقسام التطبيق</h2>
    </div>
    <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div style="background: #ffebee; color: #c62828; padding: 10px; border-radius: 8px; margin-bottom: 15px; text-align: center; font-weight: bold; border: 1px solid #c62828;">{{ message }}</div>
            {% endfor %}
          {% endif %}
        {% endwith %}

        {% if not filename %}
            <div class="form-card" id="main-form">
                <h3>1️⃣ استيراد ملف البيانات</h3>
                <p style="font-size:14px; color:#555;">اختر ملف الإكسيل من هاتفك لرفعه وتحليله فوراً:</p>
                
                <!-- خيار 1: رفع ملف من الهاتف -->
                <form method="POST" action="/upload" enctype="multipart/form-data" onsubmit="showLoader()">
                    <input type="file" name="file" accept=".xlsx, .xls" required>
                    <button type="submit" class="btn-submit">⬆️ رفع وتحليل الملف الآن</button>
                </form>

                <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">

                <!-- خيار 2: قراءة الملف الموجود مسبقاً -->
                <p style="font-size:12px; color:#777;">أو قراءة الملف الموجود مسبقاً في مجلد uploads باسم uploaded_file.xlsx:</p>
                <a href="/upload_existing" onclick="showLoader()" class="btn-submit btn-secondary">📂 قراءة الملف الموجود</a>
            </div>
            <div id="loading-div" style="display:none; text-align:center; background:#fff; padding:20px; border-radius:15px; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                <h3 style="color:#0a2a1a;">⏳ جاري قراءة الملف وتحليله...</h3>
                <p>قد يستغرق هذا بضع ثوانٍ نظراً لحجم البيانات.</p>
                <div class="loader"></div>
            </div>
        {% else %}
            <div class="file-info">الملف: {{ filename }}<br>الصفوف: {{ rows }} | الأعمدة: {{ cols }}</div>
            <form method="POST" action="/select_sheet" style="margin-bottom:15px;">
                <select name="sheet" onchange="this.form.submit()" style="width:100%; padding:10px; border-radius:8px; border:1px solid #ddd;">
                    {% for sheet in sheets %}<option value="{{ sheet }}" {% if sheet == current_sheet %}selected{% endif %}>ورقة: {{ sheet }}</option>{% endfor %}
                </select>
            </form>
            
            <a href="/check" class="card"><div class="card-icon icon-blue">🔍</div><div class="card-content"><h3>2️⃣ فحص الجاهزية</h3><p>فحص القيم الفارغة والمكررة</p></div></a>
            <a href="/describe" class="card"><div class="card-icon icon-green">📊</div><div class="card-content"><h3>3️⃣ التحليل الوصفي</h3><p>الإحصاءات الوصفية والاتجاهات</p></div></a>
            <a href="/performance" class="card"><div class="card-icon icon-orange">📈</div><div class="card-content"><h3>4️⃣ تحليل الأداء</h3><p>مقارنة الإنجاز والمؤشرات</p></div></a>
            <a href="/visual" class="card"><div class="card-icon icon-teal">📉</div><div class="card-content"><h3>5️⃣ التحليل البصري</h3><p>رسوم بيانية (أعمدة، خطوط)</p></div></a>
            <a href="/advanced" class="card"><div class="card-icon icon-indigo">🧠</div><div class="card-content"><h3>6️⃣ التحليل المتقدم</h3><p>الارتباطات والتوزيعات</p></div></a>
            <a href="/insights" class="card"><div class="card-icon icon-purple">💡</div><div class="card-content"><h3>7️⃣ الرؤى الاستراتيجية</h3><p>استخراج أفضل وأسوأ النتائج</p></div></a>
            <a href="/summary" class="card"><div class="card-icon icon-brown">🧾</div><div class="card-content"><h3>8️⃣ الخلاصات العامة</h3><p>تجميع الإجماليات</p></div></a>
            <a href="/recommendations" class="card"><div class="card-icon icon-pink">🎯</div><div class="card-content"><h3>9️⃣ التوصيات والمقترحات</h3><p>دعم القرار</p></div></a>
            
            <a href="/reset" class="card" style="border-color:#ffcdd2; margin-top:20px;"><div class="card-icon icon-red">🗑️</div><div class="card-content"><h3>تغيير الملف</h3><p>العودة لقراءة ملف جديد</p></div></a>
        {% endif %}
    </div>
</body>
</html>
"""

RESULT_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>النتيجة</title>
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; font-family: 'Tajawal', sans-serif; }
        body { background: #f4f6f5; color: #333; padding: 15px; }
        .card { background: #fff; padding: 20px; border-radius: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border: 1px solid #eee; overflow-x: auto; }
        h2 { color: #0a2a1a; text-align: center; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background: #0a2a1a; color: #fff; }
        .btn-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 20px; }
        .btn { flex: 1; min-width: 120px; padding: 12px; background: #0a2a1a; color: #d4af37; border: none; border-radius: 8px; font-size: 14px; cursor: pointer; text-align: center; text-decoration: none; font-weight: bold; }
        .btn-pdf { background: #c62828; color: #fff; }
        .btn-excel { background: #2e7d32; color: #fff; }
        .btn-word { background: #1565c0; color: #fff; }
        .insight-box { background: #e8f5e9; border-right: 5px solid #388e3c; padding: 15px; margin-bottom: 15px; border-radius: 8px; }
        .insight-box h3 { margin: 0 0 10px 0; color: #388e3c; }
        .insight-box p { margin: 5px 0; font-size: 15px; }
    </style>
</head>
<body>
    <div class="card">
        <h2>{{ title }}</h2>
        {{ result | safe }}
        
        {% if export_id %}
        <div class="btn-group">
            <a href="/export/pdf/{{ export_id }}" class="btn btn-pdf">📄 PDF</a>
            <a href="/export/excel/{{ export_id }}" class="btn btn-excel">📊 Excel</a>
            <a href="/export/word/{{ export_id }}" class="btn btn-word">📝 Word</a>
        </div>
        {% endif %}
        
        <a href="/" class="btn" style="margin-top: 15px; display: block;">← العودة للرئيسية</a>
    </div>
</body>
</html>
"""

# ==========================================
# المسارات
# ==========================================

@app.route('/login', methods=['POST'])
def login():
    if request.form.get('password') == 'admin':
        session['logged_in'] = True
        return redirect('/')
    flash('كلمة المرور غير صحيحة', 'error')
    return redirect('/')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect('/')

@app.route('/reset')
def reset():
    data_store['current_df'] = None
    data_store['filename'] = ''
    return redirect('/')

@app.route('/')
def index():
    if not session.get('logged_in'):
        logo_image = get_image_base64('logo_yt.png')
        return render_template_string(LOGIN_TEMPLATE, logo_image=logo_image)
    
    bg_data = get_image_base64('bg_data.jpg') 
    
    if data_store.get('excel_file') is not None and data_store.get('current_df') is not None:
        df = data_store['current_df']
        return render_template_string(DASHBOARD_TEMPLATE, 
            filename=data_store['filename'], 
            sheets=data_store['sheets'],
            current_sheet=data_store.get('current_sheet_name', data_store['sheets'][0]),
            rows=df.shape[0], cols=df.shape[1],
            bg_data=bg_data)
    return render_template_string(DASHBOARD_TEMPLATE, filename=None, bg_data=bg_data)

# --- مسار رفع ملف جديد من الهاتف ---
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('لم يتم اختيار ملف', 'error')
        return redirect('/')
    
    file = request.files['file']
    if file.filename == '':
        flash('لم يتم اختيار ملف', 'error')
        return redirect('/')
    
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], "uploaded_file.xlsx")
        file.save(filepath)
        
        try:
            excel = pd.ExcelFile(filepath)
            data_store['excel_file'] = excel
            data_store['sheets'] = excel.sheet_names
            data_store['filename'] = file.filename
            data_store['current_sheet_name'] = data_store['sheets'][0]
            
            df = pd.read_excel(filepath, sheet_name=data_store['current_sheet_name'], header=None)
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            df = df.dropna(axis=1, how='all')
            df = df.loc[:, ~df.columns.duplicated()]
            
            df, cat_cols, num_cols = process_dataframe(df)
            data_store['current_df'] = df
            data_store['cat_cols'] = cat_cols
            data_store['num_cols'] = num_cols
            
            flash('تم رفع الملف وتحليله بنجاح!', 'success')
        except Exception as e:
            flash(f'خطأ في القراءة: {str(e)}', 'error')
    else:
        flash('صيغة الملف غير مدعومة. يرجى رفع ملف .xlsx أو .xls', 'error')
    return redirect('/')

# --- مسار قراءة الملف الموجود مسبقاً ---
@app.route('/upload_existing', methods=['GET'])
def upload_existing():
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], "uploaded_file.xlsx")
    if not os.path.exists(filepath):
        flash('خطأ: لم يتم العثور على الملف الموجود.', 'error')
        return redirect('/')
    try:
        excel = pd.ExcelFile(filepath)
        data_store['excel_file'] = excel
        data_store['sheets'] = excel.sheet_names
        data_store['filename'] = "uploaded_file.xlsx"
        data_store['current_sheet_name'] = data_store['sheets'][0]
        
        df = pd.read_excel(filepath, sheet_name=data_store['current_sheet_name'], header=None)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        df = df.loc[:, ~df.columns.duplicated()]
        
        df, cat_cols, num_cols = process_dataframe(df)
        data_store['current_df'] = df
        data_store['cat_cols'] = cat_cols
        data_store['num_cols'] = num_cols
        
        flash('تم تحميل الملف بنجاح!', 'success')
    except Exception as e:
        flash(f'خطأ في القراءة: {str(e)}', 'error')
    return redirect('/')

@app.route('/select_sheet', methods=['POST'])
def select_sheet():
    sheet = request.form.get('sheet')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], "uploaded_file.xlsx")
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        df = df.loc[:, ~df.columns.duplicated()]
        
        df, cat_cols, num_cols = process_dataframe(df)
        data_store['current_df'] = df
        data_store['cat_cols'] = cat_cols
        data_store['num_cols'] = num_cols
        data_store['current_sheet_name'] = sheet
        
        flash(f'تم تحميل الورقة: {sheet}', 'success')
    except Exception as e:
        flash(f'خطأ: {str(e)}', 'error')
    return redirect('/')

@app.route('/check')
def check():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    missing = df.isnull().sum().sum()
    duplicates = df.duplicated().sum()
    
    result = f"<p>إجمالي الصفوف: {df.shape[0]:,}</p>"
    result += f"<p>الأعمدة المفيدة: {df.shape[1]:,}</p>"
    result += f"<p>القيم الفارغة المتبقية: {missing:,}</p>"
    result += f"<p>الصفوف المكررة: {duplicates:,}</p>"
    result += "<p style='color:green; margin-top:15px;'>✅ البيانات جاهزة للتحليل.</p>"
    return render_template_string(RESULT_TEMPLATE, title="2. تقرير فحص الجاهزية", result=result, export_id=None)

@app.route('/describe')
def describe():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    try:
        num_cols = data_store.get('num_cols', [])
        if not num_cols:
             return render_template_string(RESULT_TEMPLATE, title="التحليل الوصفي", result="<p>لا توجد أعمدة رقمية.</p>", export_id=None)

        stats = df[num_cols].describe().T
        stats = stats.round(2)
        stats.columns = ['العدد', 'المتوسط', 'الانحراف المعياري', 'أقل قيمة', 'الربع الأول', 'الوسيط', 'الربع الثالث', 'أعلى قيمة']
        
        result_table = stats.to_html(classes='table', border=1)
        result_text = "<p style='color:green; font-weight:bold;'>تم عرض إحصاءات الأعمدة الرقمية.</p>" + result_table
        
        report_id = save_report(stats, "التحليل الوصفي")
        return render_template_string(RESULT_TEMPLATE, title="3. التحليل الوصفي", result=result_text, export_id=report_id)
    except Exception as e:
        return render_template_string(RESULT_TEMPLATE, title="خطأ", result=f"<pre>{str(e)}</pre>", export_id=None)

@app.route('/performance')
def performance():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    num_cols = data_store.get('num_cols', [])
    progress_col = next((c for c in num_cols if 'إنجاز' in str(c) or 'نسبة' in str(c)), None)
    cost_col = next((c for c in num_cols if 'تكلفة' in str(c)), None)
    
    result_text = "<h3>مؤشرات الأداء الرئيسية (KPIs)</h3>"
    
    if progress_col:
        avg_progress = df[progress_col].mean()
        result_text += f"<p>متوسط نسبة الإنجاز العام: <b>{avg_progress:.2f}%</b></p>"
    else:
        result_text += "<p>لم يتم العثور على عمود لنسب الإنجاز.</p>"
        
    if cost_col:
        total_cost = df[cost_col].sum()
        result_text += f"<p>إجمالي التكاليف: <b>{total_cost:,.2f}</b></p>"
    else:
        result_text += "<p>لم يتم العثور على عمود للتكاليف.</p>"
        
    report_id = save_report(pd.DataFrame([{'المؤشر': 'نسبة الإنجاز', 'القيمة': avg_progress if progress_col else 0}]), "تحليل الأداء")
    return render_template_string(RESULT_TEMPLATE, title="4. تحليل الأداء", result=result_text, export_id=report_id)

@app.route('/visual', methods=['GET', 'POST'])
def visual():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    cat_cols = data_store.get('cat_cols', [])
    num_cols = data_store.get('num_cols', [])
    
    if request.method == 'POST':
        group_col = request.form.get('group_col')
        sum_col = request.form.get('sum_col')
        
        summary_df = df.groupby(group_col)[sum_col].sum().reset_index().sort_values(by=sum_col, ascending=False)

        plt.figure(figsize=(8, 5))
        plt.bar(summary_df[group_col].head(10).astype(str), summary_df[sum_col].head(10), color='#0a2a1a')
        plt.xticks(rotation=45, ha='right')
        plt.title(f'أعلى 10 نتائج: {sum_col} حسب {group_col}')
        plt.tight_layout()
        
        img = io.BytesIO()
        plt.savefig(img, format='png')
        img.seek(0)
        plot_url = base64.b64encode(img.getvalue()).decode()
        plt.close()
        
        result_text = f'<img src="data:image/png;base64,{plot_url}" style="width:100%; border-radius:10px;">'
        report_id = save_report(summary_df, "التحليل البصري")
        return render_template_string(RESULT_TEMPLATE, title="5. التحليل البصري", result=result_text, export_id=report_id)

    form_html = "<p>اختر البيانات لإنشاء الرسم البياني:</p>"
    form_html += "<form method='POST'>"
    form_html += "<select name='group_col' style='margin-bottom:10px; width:100%;'>" + "".join([f"<option value='{c}'>{c}</option>" for c in cat_cols]) + "</select>"
    form_html += "<select name='sum_col' style='margin-bottom:10px; width:100%;'>" + "".join([f"<option value='{c}'>{c}</option>" for c in num_cols]) + "</select>"
    form_html += "<button type='submit' class='btn'>توليد الرسم</button>"
    form_html += "</form>"
    
    return render_template_string(RESULT_TEMPLATE, title="5. التحليل البصري", result=form_html, export_id=None)

@app.route('/advanced')
def advanced():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    num_cols = data_store.get('num_cols', [])
    result = "<h3>تحليل الارتباطات والتوزيعات</h3>"
    
    if len(num_cols) > 1:
        temp_df = df[num_cols[:5]]
        corr = temp_df.corr().round(2)
        result += "<p>مصفوفة الارتباط (Correlation Matrix) لأول 5 أعمدة رقمية:</p>"
        result += corr.to_html(classes='table', border=1)
    else:
        result += "<p>لا توجد أعمدة رقمية كافية لحساب الارتباطات.</p>"
        
    return render_template_string(RESULT_TEMPLATE, title="6. التحليل المتقدم", result=result, export_id=None)

@app.route('/insights')
def insights():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    cat_cols = data_store.get('cat_cols', [])
    num_cols = data_store.get('num_cols', [])
    
    if not cat_cols or not num_cols:
        return render_template_string(RESULT_TEMPLATE, title="الرؤى", result="<p>بيانات غير كافية.</p>", export_id=None)

    group_col = cat_cols[0]
    num_col = next((c for c in num_cols if 'إنجاز' in str(c) or 'نسبة' in str(c)), num_cols[0])
            
    best_row = df.loc[df[num_col].idxmax()]
    worst_row = df.loc[df[num_col].idxmin()]
    
    result_html = f"""
    <div class="insight-box">
        <h3>🌟 أفضل أداء</h3>
        <p>بناءً على مؤشر <b>({num_col})</b>، الأفضل هو:</p>
        <p><b>{best_row[group_col]}</b> بقيمة <b>{best_row[num_col]:,.2f}</b></p>
    </div>
    <div class="insight-box" style="background: #ffebee; border-right-color: #c62828;">
        <h3 style="color: #c62828;">⚠️ أدنى أداء</h3>
        <p>بناءً على مؤشر <b>({num_col})</b>، الأقل هو:</p>
        <p><b>{worst_row[group_col]}</b> بقيمة <b>{worst_row[num_col]:,.2f}</b></p>
    </div>
    """
    report_id = save_report(pd.DataFrame([{'الأفضل': best_row[group_col], 'الأقل': worst_row[group_col]}]), "الرؤى")
    return render_template_string(RESULT_TEMPLATE, title="7. الرؤى الاستراتيجية", result=result_html, export_id=report_id)

@app.route('/summary', methods=['GET', 'POST'])
def summary():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    cat_cols = data_store.get('cat_cols', [])
    num_cols = data_store.get('num_cols', [])
    
    if request.method == 'POST':
        group_col = request.form.get('group_col')
        sum_col = request.form.get('sum_col')
        
        summary_df = df.groupby(group_col)[sum_col].sum().reset_index().sort_values(by=sum_col, ascending=False)
        
        result_table = summary_df.head(20).to_html(classes='table', index=False, border=1)
        result_text = f"<p>تم التجميع بنجاح.</p>" + result_table
        
        report_id = save_report(summary_df, "الخلاصات")
        return render_template_string(RESULT_TEMPLATE, title="8. الخلاصات العامة", result=result_text, export_id=report_id)

    form_html = "<p>اختر البيانات لتوليد الخلاصة:</p>"
    form_html += "<form method='POST'>"
    form_html += "<select name='group_col' style='margin-bottom:10px; width:100%;'>" + "".join([f"<option value='{c}'>{c}</option>" for c in cat_cols]) + "</select>"
    form_html += "<select name='sum_col' style='margin-bottom:10px; width:100%;'>" + "".join([f"<option value='{c}'>{c}</option>" for c in num_cols]) + "</select>"
    form_html += "<button type='submit' class='btn'>توليد الخلاصة</button>"
    form_html += "</form>"
    
    return render_template_string(RESULT_TEMPLATE, title="8. الخلاصات العامة", result=form_html, export_id=None)

@app.route('/recommendations')
def recommendations():
    df = data_store.get('current_df')
    if df is None: return redirect('/')
    
    result = "<h3>التوصيات بناءً على المعايير الدولية</h3>"
    result += "<ul>"
    result += "<li>مراجعة المشاريع التي تقل نسبة إنجازها عن 50% واتخاذ إجراءات تصحيحية.</li>"
    result += "<li>تطبيق مبادئ الحوكمة في إدارة التكاليف للمشاريع التي تجاوزت الميزانية.</li>"
    result += "<li>توحيد أنماط البيانات (Data Standardization) لضمان دقة التقارير المستقبلية.</li>"
    result += "</ul>"
    report_id = save_report(pd.DataFrame([{'التوصية': 'مراجعة المشاريع المتعثرة'}]), "التوصيات")
    return render_template_string(RESULT_TEMPLATE, title="9. التوصيات والمقترحات", result=result, export_id=report_id)

# ==========================================
# التصدير الموحد (PDF, Excel, Word)
# ==========================================
@app.route('/export/<format_type>/<report_id>')
def export_report(format_type, report_id):
    report = data_store['reports'].get(report_id)
    if not report:
        return "التقرير غير موجود أو انتهت صلاحيته.", 404
        
    title = report['title']
    data = report['data']
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    if format_type == 'pdf':
        filename = f"{title}_{timestamp}.pdf"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        doc = SimpleDocTemplate(filepath, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(name='ArabicTitle', parent=styles['Title'], fontName='Arabic', fontSize=16, alignment=1)
        cell_style = ParagraphStyle(name='ArabicCell', parent=styles['Normal'], fontName='Arabic', fontSize=8, alignment=1)
        
        elements.append(Paragraph(ar(title), title_style))
        elements.append(Spacer(1, 15))
        
        if isinstance(data, pd.DataFrame):
            formatted_data = []
            header = [Paragraph(ar(str(col)), cell_style) for col in data.columns]
            formatted_data.append(header)
            
            for _, row in data.iterrows():
                row_data = []
                for val in row.values:
                    if isinstance(val, (int, float, np.number)):
                        if pd.isna(val):
                            row_data.append("-")
                        elif abs(val) >= 1000000:
                            row_data.append(f"{val:,.0f}")
                        else:
                            row_data.append(f"{val:,.2f}")
                    else:
                        row_data.append(Paragraph(ar(str(val)), cell_style))
                formatted_data.append(row_data)
            
            num_cols = len(data.columns)
            available_width = 555
            col_width = available_width / num_cols
            
            table = Table(formatted_data, colWidths=[col_width]*num_cols)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0a2a1a')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,-1), 'Arabic'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0f4f8')]),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            elements.append(table)
        
        doc.build(elements)
        return send_file(filepath, as_attachment=True, download_name=filename)
        
    elif format_type == 'excel':
        filename = f"{title}_{timestamp}.xlsx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if isinstance(data, pd.DataFrame):
            data.to_excel(filepath, index=False)
        return send_file(filepath, as_attachment=True, download_name=filename)
        
    elif format_type == 'word':
        filename = f"{title}_{timestamp}.docx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        doc = Document()
        doc.add_heading(title, 0)
        
        if isinstance(data, pd.DataFrame):
            table = doc.add_table(rows=1, cols=len(data.columns))
            hdr_cells = table.rows[0].cells
            for i, col in enumerate(data.columns):
                hdr_cells[i].text = str(col)
            for _, row in data.head(50).iterrows():
                row_cells = table.add_row().cells
                for i, val in enumerate(row.values):
                    row_cells[i].text = str(val)
        doc.save(filepath)
        return send_file(filepath, as_attachment=True, download_name=filename)

    return "صيغة غير مدعومة", 400

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)