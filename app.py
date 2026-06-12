from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, abort
import pymysql
import os
import time
import re
import secrets
import string
import logging
from datetime import timedelta
from collections import defaultdict
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
from PIL import Image

app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# إعداد الاتصال بـ MySQL (بدون الحاجة لملف db.py)
def get_db_connection():
    return pymysql.connect(
        host=os.environ.get('DB_HOST'),
        user=os.environ.get('DB_USER'),
        password=os.environ.get('DB_PASSWORD'),
        database=os.environ.get('DB_NAME'),
        cursorclass=pymysql.cursors.DictCursor,
        autocommit=True
    )

# ... [بقية الدوال كما هي، تأكدي فقط من أنكِ قمتِ بضبط المتغيرات في Render] ...

# دالة تهيئة الجداول (تأكدي من تشغيلها مرة واحدة)
def ensure_schema():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        # إضافة أعمدة المشاهدات والتحميل إذا لم تكن موجودة
        cursor.execute("SHOW COLUMNS FROM handouts LIKE 'view_count'")
        if not cursor.fetchone():
            cursor.execute('ALTER TABLE handouts ADD COLUMN view_count INT DEFAULT 0')
        cursor.execute("SHOW COLUMNS FROM handouts LIKE 'download_count'")
        if not cursor.fetchone():
            cursor.execute('ALTER TABLE handouts ADD COLUMN download_count INT DEFAULT 0')
    finally:
        conn.close()

# تشغيل التطبيق
if __name__ == '__main__':
    # تأكدي من إعداد المتغيرات في الريندر قبل التشغيل
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
