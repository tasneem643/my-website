# -*- coding: utf-8 -*-
"""
طبقة وصول مزدوجة لقاعدة البيانات (SQLite افتراضياً / MySQL عبر PyMySQL).

Dual-driver database access layer for the Musaid Portal.

- Default backend is SQLite (zero behavior change vs. the original app).
- When MySQL is configured (DATABASE_URL=mysql+pymysql://...  OR MUSAID_DB_BACKEND=mysql),
  connections are created through a SQLAlchemy engine (connection pooling) using the
  PyMySQL driver, wrapped so the existing raw-SQL routes keep working unchanged:
    * "?" placeholders are translated to "%s" (PyMySQL paramstyle).
    * Result rows behave like sqlite3.Row — support BOTH row[0] and row['col'],
      dict(row), iteration, .keys(), and .get().
    * conn.execute(...), conn.cursor(), .fetchone(), .fetchall(), .lastrowid,
      .commit(), .close() all mirror the sqlite3 API the app relies on.

The app code (app.py) calls db.connect(DATABASE) and otherwise stays the same.
"""
import os
import sqlite3

_engine = None  # lazily-created SQLAlchemy engine (MySQL only)


# --------------------------------------------------------------------------
# اختيار المحرك من البيئة
# --------------------------------------------------------------------------
def backend():
    """يُعيد 'mysql' أو 'sqlite' بحسب متغيرات البيئة."""
    url = os.environ.get('DATABASE_URL', '').strip()
    if url.startswith('mysql'):
        return 'mysql'
    if os.environ.get('MUSAID_DB_BACKEND', '').strip().lower() in ('mysql', 'mariadb'):
        return 'mysql'
    return 'sqlite'


def is_mysql():
    return backend() == 'mysql'


def database_url():
    """يبني عنوان اتصال SQLAlchemy لـ MySQL من DATABASE_URL أو من المتغيرات المنفصلة."""
    url = os.environ.get('DATABASE_URL', '').strip()
    if url:
        return url
    user = os.environ.get('MYSQL_USER', 'musaid')
    pw = os.environ.get('MYSQL_PASSWORD', '')
    host = os.environ.get('MYSQL_HOST', 'localhost')
    port = os.environ.get('MYSQL_PORT', '3306')
    name = os.environ.get('MYSQL_DB', 'musaid_portal')
    return f'mysql+pymysql://{user}:{pw}@{host}:{port}/{name}?charset=utf8mb4'


def get_engine():
    """ينشئ (مرة واحدة) محرك SQLAlchemy لـ MySQL مع تجميع الاتصالات."""
    global _engine
    if _engine is None:
        from sqlalchemy import create_engine
        _engine = create_engine(
            database_url(),
            pool_pre_ping=True,   # يتفادى الاتصالات الميتة
            pool_recycle=3600,    # يعيد تدوير الاتصال كل ساعة
            future=True,
        )
    return _engine


# --------------------------------------------------------------------------
# ترجمة صيغة المعاملات: "?"  ->  "%s"  (مع تهريب "%" الحرفية)
# --------------------------------------------------------------------------
def translate(sql):
    """يحوّل صيغة معاملات SQLite (?) إلى صيغة PyMySQL (%s).

    ملاحظة: التطبيق لا يستخدم رموز % حرفية داخل SQL ولا يضع "?" داخل سلاسل نصية،
    لذا هذا التحويل البسيط آمن لهذه القاعدة البرمجية تحديداً.
    """
    if '%' in sql:
        sql = sql.replace('%', '%%')
    return sql.replace('?', '%s')


# --------------------------------------------------------------------------
# صف متوافق مع sqlite3.Row (فهرسة رقمية ونصية معاً)
# --------------------------------------------------------------------------
class Row:
    """يغلّف صفاً من قاموس DictCursor ليتصرف مثل sqlite3.Row."""
    __slots__ = ('_d', '_v')

    def __init__(self, d):
        self._d = d
        self._v = list(d.values())

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._v[key]
        return self._d[key]

    def get(self, key, default=None):
        return self._d.get(key, default)

    def keys(self):
        return list(self._d.keys())

    def __iter__(self):
        # sqlite3.Row يكرّر القيم
        return iter(self._v)

    def __len__(self):
        return len(self._v)

    def __contains__(self, key):
        return key in self._d


# --------------------------------------------------------------------------
# مؤشر/اتصال MySQL يحاكيان واجهة sqlite3 التي يعتمدها التطبيق
# --------------------------------------------------------------------------
class _ResultCursor:
    def __init__(self, raw_cursor):
        self._cur = raw_cursor

    def execute(self, sql, params=()):
        self._cur.execute(translate(sql), params)
        return self

    def fetchone(self):
        r = self._cur.fetchone()
        return Row(r) if r is not None else None

    def fetchall(self):
        return [Row(r) for r in self._cur.fetchall()]

    def __iter__(self):
        return (Row(r) for r in self._cur.fetchall())

    @property
    def lastrowid(self):
        return self._cur.lastrowid

    @property
    def rowcount(self):
        return self._cur.rowcount

    def close(self):
        try:
            self._cur.close()
        except Exception:
            pass


class MySQLConnection:
    """يغلّف اتصال PyMySQL ليحاكي sqlite3.Connection بالقدر الذي يستخدمه التطبيق."""
    def __init__(self, raw_conn):
        self._raw = raw_conn

    def _dict_cursor(self):
        import pymysql.cursors
        return self._raw.cursor(pymysql.cursors.DictCursor)

    def execute(self, sql, params=()):
        cur = self._dict_cursor()
        cur.execute(translate(sql), params)
        return _ResultCursor(cur)

    def cursor(self):
        return _ResultCursor(self._dict_cursor())

    def commit(self):
        self._raw.commit()

    def rollback(self):
        try:
            self._raw.rollback()
        except Exception:
            pass

    def close(self):
        # raw_connection من SQLAlchemy: close() يعيد الاتصال إلى التجمّع
        try:
            self._raw.close()
        except Exception:
            pass


# --------------------------------------------------------------------------
# نقطة الدخول الموحّدة
# --------------------------------------------------------------------------
def connect(sqlite_path):
    """يفتح اتصالاً حسب المحرك المختار.

    SQLite: يعيد sqlite3.Connection (سلوك مطابق للأصل تماماً).
    MySQL : يعيد MySQLConnection المتوافق.
    """
    if is_mysql():
        raw = get_engine().raw_connection()  # اتصال PyMySQL من تجمّع SQLAlchemy
        return MySQLConnection(raw)
    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    return conn
