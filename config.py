import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Native .env loader (avoids third-party dependencies)
env_path = os.path.join(BASE_DIR, '.env')
if os.path.exists(env_path):
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if '=' in line and not line.startswith('#'):
                key, val = line.split('=', 1)
                os.environ[key.strip()] = val.strip()

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'yss-anantapur-spiritual-2026-secret-key-ravi')
    
    # Database Configuration
    database_url = os.environ.get('DATABASE_URL')
    if database_url and database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    
    # Fallback to local SQLite if DATABASE_URL is not set
    SQLALCHEMY_DATABASE_URI = database_url or f"sqlite:///{os.path.join(BASE_DIR, 'instance', 'database.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": 300,
    }
    EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')

    # Flask-Mail config
    MAIL_SERVER = os.environ.get('MAIL_SERVER', 'smtp.gmail.com')
    MAIL_PORT = int(os.environ.get('MAIL_PORT', 587))
    MAIL_USE_TLS = os.environ.get('MAIL_USE_TLS', 'True').lower() == 'true'
    MAIL_USE_SSL = os.environ.get('MAIL_USE_SSL', 'False').lower() == 'true'
    MAIL_USERNAME = os.environ.get('MAIL_USERNAME', 'anantapur@ysscenters.org')
    MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD', 'zmyevudryraylggx')
    MAIL_DEFAULT_SENDER = (
        os.environ.get('MAIL_SENDER_NAME', 'YSS Anantapur'),
        os.environ.get('MAIL_USERNAME', 'anantapur@ysscenters.org')
    )

    # Event config
    EVENT_NAME = '3-Day Spiritual Program – Anantapur'
    EVENT_DATES = '24 – 26 July 2026'
    EVENT_DATE_RANGE = 'July 24, 25, 26 – 2026'
    EVENT_VENUE = 'Krishna Kala Mandir, Anantapur, Andhra Pradesh'
    EVENT_CONTACT_EMAIL = 'anantapur@ysscenters.org'
    EVENT_CONTACT_MOBILE = '9441665181'
    EVENT_MAPS_URL = 'https://maps.app.goo.gl/WQxUo86SVYtabmMP9'

    # Admin defaults
    ADMIN_EMAIL = 'anantapur@ysscenters.org'
    ADMIN_PASSWORD = 'yssatp2026'
