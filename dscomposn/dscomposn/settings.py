"""
Django settings for dscomposn
รวมระบบ COMPOSN (สิทธิ์ค่าย) + Food (เวลาเรียน / คะแนน / สั่งอาหาร)
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# โหลดค่าจาก .env (ถ้ามี)
try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / '.env')
except ImportError:  # pragma: no cover
    pass


def env_bool(name, default=False):
    return os.environ.get(name, str(default)).lower() in ('1', 'true', 'yes', 'on')


SECRET_KEY = os.environ.get(
    'DJANGO_SECRET_KEY',
    'django-insecure-dscomposn-dev-key-change-me-in-production-!@#',
)
DEBUG = env_bool('DJANGO_DEBUG', True)
ALLOWED_HOSTS = [h.strip() for h in os.environ.get('DJANGO_ALLOWED_HOSTS', '*').split(',') if h.strip()]
CSRF_TRUSTED_ORIGINS = [h.strip() for h in os.environ.get('DJANGO_CSRF_TRUSTED_ORIGINS', '').split(',') if h.strip()]

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django.contrib.humanize',
    'composn.apps.ComposnConfig',
    'food.apps.FoodConfig',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    # ระบบความปลอดภัยของ dscomposn
    'composn.middleware.SingleSessionMiddleware',
    'composn.middleware.ForcePasswordChangeMiddleware',
    # จดจำตัวเลือก/ตัวกรองล่าสุดของแต่ละหน้า เมื่อกลับมาจะอยู่ที่เดิม
    'composn.middleware.RememberPageStateMiddleware',
]

ROOT_URLCONF = 'dscomposn.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'composn.context_processors.site_context',
            ],
        },
    },
]

WSGI_APPLICATION = 'dscomposn.wsgi.application'

# ฐานข้อมูล: SQLite (พัฒนา) / PostgreSQL (Production ตั้งค่า DATABASE_ENGINE=postgresql)
if os.environ.get('DATABASE_ENGINE', 'sqlite').lower().startswith('postgres'):
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.postgresql',
            'NAME': os.environ.get('DATABASE_NAME', 'dscomposn'),
            'USER': os.environ.get('DATABASE_USER', 'postgres'),
            'PASSWORD': os.environ.get('DATABASE_PASSWORD', ''),
            'HOST': os.environ.get('DATABASE_HOST', 'localhost'),
            'PORT': os.environ.get('DATABASE_PORT', '5432'),
        }
    }
else:
    DATABASES = {
        'default': {
            'ENGINE': 'django.db.backends.sqlite3',
            'NAME': BASE_DIR / 'db.sqlite3',
            'OPTIONS': {'timeout': 20},
        }
    }

# Custom User Model: camp1_code เป็น USERNAME_FIELD
AUTH_USER_MODEL = 'composn.User'

# ไม่บังคับความซับซ้อนของรหัสผ่าน (รหัสเริ่มต้นคือ 1234 / รหัสค่าย) แต่ต้องยาวอย่างน้อย 4
AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator', 'OPTIONS': {'min_length': 4}},
]

LOGIN_URL = 'composn:login'
LOGIN_REDIRECT_URL = 'composn:home'
LOGOUT_REDIRECT_URL = 'composn:login'

# Session: Auto Logout หลังไม่มีการทำงาน 5 นาที
AUTO_LOGOUT_SECONDS = int(os.environ.get('AUTO_LOGOUT_SECONDS', 300))
SESSION_COOKIE_AGE = AUTO_LOGOUT_SECONDS
SESSION_SAVE_EVERY_REQUEST = True
SESSION_EXPIRE_AT_BROWSER_CLOSE = True
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = 'Lax'
CSRF_COOKIE_SAMESITE = 'Lax'
SECURE_CONTENT_TYPE_NOSNIFF = True
X_FRAME_OPTIONS = 'SAMEORIGIN'
SESSION_COOKIE_SECURE = env_bool('DJANGO_SECURE_COOKIES', False)
CSRF_COOKIE_SECURE = SESSION_COOKIE_SECURE

# การป้องกันการเดารหัสผ่านซ้ำ (Brute-force Lockout)
LOGIN_MAX_FAILURES_PER_USER = int(os.environ.get('LOGIN_MAX_FAILURES_PER_USER', 5))
LOGIN_MAX_FAILURES_PER_IP = int(os.environ.get('LOGIN_MAX_FAILURES_PER_IP', 15))
LOGIN_LOCKOUT_WINDOW_MINUTES = int(os.environ.get('LOGIN_LOCKOUT_WINDOW_MINUTES', 15))
LOGIN_LOCKOUT_MINUTES = int(os.environ.get('LOGIN_LOCKOUT_MINUTES', 15))

LANGUAGE_CODE = 'th'
TIME_ZONE = 'Asia/Bangkok'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
STATIC_ROOT = BASE_DIR / 'staticfiles'
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

BACKUP_DIR = BASE_DIR / 'backups'
AUDIT_LOG_DIR = BASE_DIR / 'audit_logs'

DATA_UPLOAD_MAX_NUMBER_FIELDS = 20000
FILE_UPLOAD_MAX_MEMORY_SIZE = 10 * 1024 * 1024

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ค่าคงที่ของระบบ
DEFAULT_ADMIN_USERNAME = 'admin'
DEFAULT_ADMIN_PASSWORD = '1234'
