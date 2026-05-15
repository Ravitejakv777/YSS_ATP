import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'yss-anantapur-spiritual-2026-secret-key-ravi')
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(BASE_DIR, 'instance', 'yss.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    EXPORTS_DIR = os.path.join(BASE_DIR, 'exports')

    # Flask-Mail config
    MAIL_SERVER = 'smtp.gmail.com'
    MAIL_PORT = 587
    MAIL_USE_TLS = True
    MAIL_USERNAME = 'ravitejakona2007@gmail.com'
    MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD', '')  # Set via environment variable
    MAIL_DEFAULT_SENDER = ('YSS Anantapur', 'ravitejakona2007@gmail.com')

    # Event config
    EVENT_NAME = '3-Day Spiritual Program – Anantapur'
    EVENT_DATES = '24 – 26 July 2026'
    EVENT_DATE_RANGE = 'July 24, 25, 26 – 2026'
    EVENT_VENUE = 'Krishna Kala Mandir, Anantapur, Andhra Pradesh'
    EVENT_CONTACT_EMAIL = 'ravitejakona2007@gmail.com'
    EVENT_CONTACT_MOBILE = '8019682209'
    EVENT_MAPS_URL = 'https://maps.app.goo.gl/WQxUo86SVYtabmMP9'

    # Admin defaults
    ADMIN_EMAIL = 'ravitejakona2007@gmail.com'
    ADMIN_PASSWORD = 'YSS@Ravi2026'
