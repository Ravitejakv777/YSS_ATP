from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import uuid

db = SQLAlchemy()

def generate_id(prefix='REG'):
    return f"{prefix}-{uuid.uuid4().hex[:8].upper()}"

class Admin(UserMixin, db.Model):
    __tablename__ = 'admins'
    id = db.Column(db.Integer, primary_key=True)
    lesson_no = db.Column(db.String(20), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    mobile = db.Column(db.String(15), nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<Admin {self.name}>'


class Registration(db.Model):
    __tablename__ = 'registrations'
    id = db.Column(db.Integer, primary_key=True)
    reg_id = db.Column(db.String(20), unique=True, nullable=False, default=lambda: generate_id('REG'))
    lesson_no = db.Column(db.String(20), nullable=False)
    full_name = db.Column(db.String(100), nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    age = db.Column(db.Integer, nullable=False)
    place = db.Column(db.String(100), nullable=False)
    state = db.Column(db.String(100), nullable=True)
    email = db.Column(db.String(120), nullable=True)
    country_code = db.Column(db.String(10), default='+91')
    whatsapp = db.Column(db.String(15), nullable=False)
    is_kriyaban = db.Column(db.Boolean, default=False)
    accommodation = db.Column(db.Boolean, default=False)
    volunteer = db.Column(db.Boolean, default=False)
    arrival_date = db.Column(db.String(20), nullable=False)
    departure_date = db.Column(db.String(20), nullable=False)
    payment_mode = db.Column(db.String(30), nullable=False)
    amount = db.Column(db.Float, nullable=True)
    transaction_id = db.Column(db.String(100), nullable=True)
    payment_screenshot = db.Column(db.String(255), nullable=True)
    payment_status = db.Column(db.String(20), default='Pending')  # Pending, Paid
    reg_status = db.Column(db.String(20), default='Pending')  # Pending, Approved, Rejected
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'reg_id': self.reg_id,
            'lesson_no': self.lesson_no,
            'full_name': self.full_name,
            'gender': self.gender,
            'age': self.age,
            'place': self.place,
            'state': self.state,
            'email': self.email,
            'country_code': self.country_code,
            'whatsapp': self.whatsapp,
            'is_kriyaban': 'Yes' if self.is_kriyaban else 'No',
            'accommodation': 'Yes' if self.accommodation else 'No',
            'volunteer': 'Yes' if self.volunteer else 'No',
            'arrival_date': self.arrival_date,
            'departure_date': self.departure_date,
            'payment_mode': self.payment_mode,
            'amount': self.amount,
            'transaction_id': self.transaction_id,
            'payment_screenshot': self.payment_screenshot,
            'payment_status': self.payment_status,
            'reg_status': self.reg_status,
            'created_at': self.created_at.strftime('%d %b %Y') if self.created_at else ''
        }


class Donation(db.Model):
    __tablename__ = 'donations'
    id = db.Column(db.Integer, primary_key=True)
    donation_id = db.Column(db.String(20), unique=True, nullable=False, default=lambda: generate_id('DON'))
    lesson_no = db.Column(db.String(20), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    age = db.Column(db.Integer, nullable=False)
    place = db.Column(db.String(100), nullable=False)
    whatsapp = db.Column(db.String(15), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    payment_mode = db.Column(db.String(30), nullable=False)
    transaction_id = db.Column(db.String(100), nullable=True)
    payment_screenshot = db.Column(db.String(255), nullable=True)
    payment_status = db.Column(db.String(20), default='Pending')  # Pending, Completed
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            'id': self.id,
            'donation_id': self.donation_id,
            'lesson_no': self.lesson_no,
            'name': self.name,
            'age': self.age,
            'place': self.place,
            'whatsapp': self.whatsapp,
            'amount': self.amount,
            'payment_mode': self.payment_mode,
            'transaction_id': self.transaction_id,
            'payment_screenshot': self.payment_screenshot,
            'payment_status': self.payment_status,
            'created_at': self.created_at.strftime('%d %b %Y') if self.created_at else ''
        }


class EventSchedule(db.Model):
    __tablename__ = 'event_schedules'
    id = db.Column(db.Integer, primary_key=True)
    day_number = db.Column(db.Integer, nullable=False)  # 1, 2, 3
    day_label = db.Column(db.String(50), nullable=False)  # Day 1 – 20 June 2026
    start_time = db.Column(db.String(10), nullable=False)  # 05:00 AM
    end_time = db.Column(db.String(10), nullable=False)    # 06:00 AM
    activity = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(30), nullable=False)  # meditation, food, talk, bhajan, volunteer, other
    sort_order = db.Column(db.Integer, default=0)

    def to_dict(self):
        return {
            'id': self.id,
            'day_number': self.day_number,
            'day_label': self.day_label,
            'start_time': self.start_time,
            'end_time': self.end_time,
            'activity': self.activity,
            'category': self.category,
            'sort_order': self.sort_order
        }
