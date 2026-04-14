import socket
from pathlib import Path


def _discover_local_hosts():
    hosts = {'127.0.0.1', 'localhost'}
    try:
        hostname = socket.gethostname()
        if hostname:
            hosts.add(hostname)
        for info in socket.getaddrinfo(hostname, None):
            address = info[4][0]
            if address and ':' not in address:
                hosts.add(address)
    except OSError:
        pass
    return sorted(hosts)

BASE_DIR = Path(__file__).resolve().parent.parent
SECRET_KEY = 'hdec-1100mw-alhenakiya-secret-2025-xK9p'
DEBUG = False
LOCAL_HOSTS = _discover_local_hosts()
ALLOWED_HOSTS = list(dict.fromkeys([
    'hdec-om.live',
    'www.hdec-om.live',
    '93.127.141.7',
    *LOCAL_HOSTS,
]))

# ── HTTPS / PROXY SETTINGS ───────────────────────────────────────────────────
# Tell Django it's behind Caddy (HTTPS reverse proxy)
SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')
USE_X_FORWARDED_HOST = True

# Secure cookies over HTTPS
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True
CSRF_TRUSTED_ORIGINS = [
    'https://hdec-om.live',
    'https://www.hdec-om.live',
    *[f'http://{host}:8000' for host in LOCAL_HOSTS],
]

INSTALLED_APPS = [
    'django.contrib.staticfiles',
    'core',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'enterprise_hub.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.template.context_processors.csrf',
            ],
        },
    },
]

# Signed-cookie sessions — zero setup, no DB, no files needed
SESSION_ENGINE = 'django.contrib.sessions.backends.signed_cookies'
SESSION_COOKIE_AGE = 86400        # 24 hours
SESSION_COOKIE_NAME = 'hdec_sid'
SESSION_COOKIE_HTTPONLY = True

STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ── MEDIA FILES (uploaded photos, checklists) ──────────────────────────────
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

# ── EMAIL CONFIGURATION ─────────────────────────────────────────────────────
# Loads SMTP settings from cmms_email_config.json if it exists (set via admin panel).
# Falls back to console backend when not configured.
import json as _json
_email_cfg_file = BASE_DIR / 'cmms_email_config.json'
if _email_cfg_file.exists():
    try:
        _cfg = _json.loads(_email_cfg_file.read_text())
        if _cfg.get('host') and _cfg.get('username') and _cfg.get('password'):
            EMAIL_BACKEND   = 'django.core.mail.backends.smtp.EmailBackend'
            EMAIL_HOST      = _cfg['host']
            EMAIL_PORT      = int(_cfg.get('port', 587))
            EMAIL_USE_TLS   = bool(_cfg.get('use_tls', True))
            EMAIL_HOST_USER = _cfg['username']
            EMAIL_HOST_PASSWORD = _cfg['password']
            DEFAULT_FROM_EMAIL  = _cfg.get('from_email', f"HDEC CMMS <{_cfg['username']}>")
        else:
            EMAIL_BACKEND = 'django.core.mail.backends.console.EmailBackend'
            DEFAULT_FROM_EMAIL = 'HDEC CMMS <noreply@hdec.sa>'
    except Exception:
        EMAIL_BACKEND = 'django.core.mail.backends.console.EmailBackend'
        DEFAULT_FROM_EMAIL = 'HDEC CMMS <noreply@hdec.sa>'
else:
    EMAIL_BACKEND = 'django.core.mail.backends.console.EmailBackend'
    DEFAULT_FROM_EMAIL = 'HDEC CMMS <noreply@hdec.sa>'

# ── CMMS SETTINGS ───────────────────────────────────────────────────────────
CMMS_DATA_DIR = BASE_DIR / 'cmms_data'

# ── HSE SETTINGS ────────────────────────────────────────────────────────────
HSE_DATA_DIR = BASE_DIR / 'data' / 'hse'

# ── PORTAL URL ───────────────────────────────────────────────────────────────
HDEC_PORTAL_URL = 'https://hdec-om.live'

# ── LIVE CHECKLIST DATA SOURCE ───────────────────────────────────────────────
CMMS_CHECKLIST_DATA_SOURCE_URL = (
    'https://docs.google.com/spreadsheets/d/'
    '1HNbHa021efx6cly2vUqMGavlCMFJqwc_86UfTqn9Yrs/edit?usp=sharing'
)
CMMS_CHECKLIST_HTTP_TIMEOUT = 10

