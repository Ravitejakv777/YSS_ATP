from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file, make_response
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_mail import Mail, Message
from models import db, Admin, Registration, Donation, EventSchedule, Room, RoomAllotment, ActivityLog, WhatsAppTemplate, GalleryImage
from config import Config
from datetime import datetime, timedelta
import os, openpyxl, uuid

app = Flask(__name__)
app.config.from_object(Config)
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'gallery'), exist_ok=True) # Ensure gallery upload directory exists

def to_ist(dt, format_str=None):
    if not dt:
        return ""
    ist_dt = dt + timedelta(hours=5, minutes=30)
    if format_str:
        return ist_dt.strftime(format_str)
    return ist_dt

app.jinja_env.filters['to_ist'] = to_ist

db.init_app(app)
mail = Mail(app)

login_manager = LoginManager(app)
login_manager.login_view = 'admin_login'
login_manager.login_message = 'Please login to access admin panel.'

# Initialize Secondary Database Engine (Railway DB) if configured
from sqlalchemy import create_engine, event, insert, update, delete
secondary_engine = None
secondary_url = app.config.get('SQLALCHEMY_SECONDARY_URI')
if secondary_url:
    try:
        secondary_engine = create_engine(
            secondary_url,
            pool_pre_ping=True,
            pool_recycle=300
        )
        print("Secondary database engine (Railway DB) initialized.")
    except Exception as e:
        print(f"Error initializing secondary database engine: {e}")

# Function to replicate changes to the secondary database
def replicate_changes(session, target_engine):
    new_objs = list(session.new)
    dirty_objs = list(session.dirty)
    deleted_objs = list(session.deleted)
    
    if not (new_objs or dirty_objs or deleted_objs):
        return
        
    try:
        with target_engine.begin() as conn:
            # 1. Insert new objects
            for obj in new_objs:
                table = obj.__table__
                values = {}
                for col in table.columns:
                    val = getattr(obj, col.key)
                    if val is not None:
                        values[col.name] = val
                conn.execute(insert(table).values(values))
                
            # 2. Update dirty objects
            for obj in dirty_objs:
                table = obj.__table__
                pks = [c.name for c in table.primary_key]
                values = {}
                filters = {}
                for col in table.columns:
                    val = getattr(obj, col.key)
                    if col.name in pks:
                        filters[col.name] = val
                    else:
                        values[col.name] = val
                stmt = update(table)
                for pk_name, pk_val in filters.items():
                    stmt = stmt.where(table.c[pk_name] == pk_val)
                conn.execute(stmt.values(values))
                
            # 3. Delete objects
            for obj in deleted_objs:
                table = obj.__table__
                pks = [c.name for c in table.primary_key]
                filters = {}
                for col in table.columns:
                    if col.name in pks:
                        filters[col.name] = getattr(obj, col.key)
                stmt = delete(table)
                for pk_name, pk_val in filters.items():
                    stmt = stmt.where(table.c[pk_name] == pk_val)
                conn.execute(stmt)
    except Exception as exc:
        print(f"!!! DUAL WRITE ERROR: Failed to replicate changes to secondary database: {exc}")
        raise exc

# Register the session listener if the secondary database is active
if secondary_engine:
    @event.listens_for(db.session, 'after_flush')
    def after_flush_listener(session, flush_context):
        replicate_changes(session, secondary_engine)

# Initialize Database on Startup (Required for Render)
with app.app_context():
    try:
        # Ensure required directories exist
        os.makedirs(app.config.get('EXPORTS_DIR', os.path.join(app.root_path, 'exports')), exist_ok=True)
        os.makedirs(os.path.join(app.root_path, 'instance'), exist_ok=True)
        
        print("Checking database connection...")
        db.create_all()
        
        # Ensure secondary database tables exist
        if secondary_engine:
            print("Checking secondary database connection and tables...")
            db.metadata.create_all(bind=secondary_engine)
        
        # Dynamic database schema verification and migration helper
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        secondary_inspector = inspect(secondary_engine) if secondary_engine else None
        
        def ensure_column_on_engine(engine, insp, table_name, col_name, col_type_sql, default_sql=None):
            if not insp or not insp.has_table(table_name):
                return
            columns = [c['name'] for c in insp.get_columns(table_name)]
            if col_name not in columns:
                print(f"Migration ({engine.name}): Column '{col_name}' is missing in table '{table_name}'. Adding it...")
                with engine.begin() as conn:
                    try:
                        # Try with IF NOT EXISTS (PostgreSQL 9.6+)
                        sql = f"ALTER TABLE {table_name} ADD COLUMN IF NOT EXISTS {col_name} {col_type_sql}"
                        if default_sql is not None:
                            sql += f" DEFAULT {default_sql}"
                        conn.execute(db.text(sql))
                    except Exception:
                        try:
                            # Fallback for SQLite
                            sql = f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type_sql}"
                            if default_sql is not None:
                                sql += f" DEFAULT {default_sql}"
                            conn.execute(db.text(sql))
                        except Exception as e:
                            print(f"Skipping migration on {engine.name} for {table_name}.{col_name}: {e}")

        def ensure_column(table_name, col_name, col_type_sql, default_sql=None):
            # Migrate primary database
            ensure_column_on_engine(db.engine, inspector, table_name, col_name, col_type_sql, default_sql)
            # Migrate secondary database
            if secondary_engine:
                ensure_column_on_engine(secondary_engine, secondary_inspector, table_name, col_name, col_type_sql, default_sql)

        # List of columns to ensure exist in registrations table:
        ensure_column('registrations', 'state', 'VARCHAR(100)')
        ensure_column('registrations', 'email', 'VARCHAR(120)')
        ensure_column('registrations', 'country_code', 'VARCHAR(10)', "'+91'")
        ensure_column('registrations', 'amount', 'FLOAT')
        ensure_column('registrations', 'transaction_id', 'VARCHAR(100)')
        ensure_column('registrations', 'payment_screenshot', 'VARCHAR(255)')
        ensure_column('registrations', 'payment_status', 'VARCHAR(20)', "'Pending'")
        ensure_column('registrations', 'reg_status', 'VARCHAR(20)', "'Pending'")
        ensure_column('registrations', 'notified', 'BOOLEAN', 'FALSE')
        ensure_column('registrations', 'district', 'VARCHAR(100)')
        
        ensure_column('registrations', 'reminder_7d_sent', 'BOOLEAN', 'FALSE')
        ensure_column('registrations', 'reminder_3d_sent', 'BOOLEAN', 'FALSE')
        ensure_column('registrations', 'reminder_1d_sent', 'BOOLEAN', 'FALSE')
        ensure_column('registrations', 'registered_by_id', 'INTEGER REFERENCES admins(id)')
        ensure_column('registrations', 'registered_by_name', 'VARCHAR(100)')
        
        # List of columns to ensure exist in donations table:
        ensure_column('donations', 'transaction_id', 'VARCHAR(100)')
        ensure_column('donations', 'payment_screenshot', 'VARCHAR(255)')
        ensure_column('donations', 'payment_status', 'VARCHAR(20)', "'Pending'")
        ensure_column('donations', 'notified', 'BOOLEAN', 'FALSE')

        # List of columns to ensure exist in admins table:
        ensure_column('admins', 'last_active', 'TIMESTAMP')
        
        # List of columns to ensure exist in room_allotments table:
        ensure_column('room_allotments', 'notified_room_number', 'VARCHAR(100)')
        ensure_column('room_allotments', 'notified_room_whatsapp', 'VARCHAR(100)')

        # Drop unique constraint on admins.email if it exists (PostgreSQL) on both primary and secondary
        try:
            db.session.execute(db.text("ALTER TABLE admins DROP CONSTRAINT IF EXISTS admins_email_key"))
            db.session.commit()
        except Exception:
            db.session.rollback()
            try:
                # Fallback: attempt without IF EXISTS (older Postgres)
                db.session.execute(db.text("ALTER TABLE admins DROP CONSTRAINT admins_email_key"))
                db.session.commit()
            except Exception as e:
                db.session.rollback()
                print(f"admins_email_key constraint already removed or not found on primary: {e}")

        if secondary_engine:
            with secondary_engine.begin() as conn:
                try:
                    conn.execute(db.text("ALTER TABLE admins DROP CONSTRAINT IF EXISTS admins_email_key"))
                except Exception:
                    try:
                        conn.execute(db.text("ALTER TABLE admins DROP CONSTRAINT admins_email_key"))
                    except Exception as e:
                        print(f"admins_email_key constraint drop skipped on secondary: {e}")
                    
        # Detect legacy registrations added by admin (Cash payment or no screenshot) and mark them on both primary and secondary
        try:
            db.session.execute(db.text("UPDATE registrations SET registered_by_name = 'Done by Admin' WHERE registered_by_name IS NULL AND (payment_mode = 'Cash' OR payment_screenshot IS NULL OR payment_screenshot = '')"))
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            print(f"Could not update old admin registrations on primary: {e}")

        if secondary_engine:
            with secondary_engine.begin() as conn:
                try:
                    conn.execute(db.text("UPDATE registrations SET registered_by_name = 'Done by Admin' WHERE registered_by_name IS NULL AND (payment_mode = 'Cash' OR payment_screenshot IS NULL OR payment_screenshot = '')"))
                except Exception as e:
                    print(f"Could not update old admin registrations on secondary: {e}")
                
        # Seed WhatsApp templates if empty
        templates_to_seed = [
            {
                'key': 'reg_success',
                'description': 'Sent automatically to devotees when their registration is approved by the admin.',
                'variables': 'name,reg_id,phone,email,city,accommodation',
                'template_text': (
                    "Dear {name},\n\n"
                    "With divine blessings and heartfelt joy, we are happy to confirm your successful registration for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
                    "Your Registration Details:\n\n"
                    "Name: {name}\n"
                    "Phone Number: {phone}\n"
                    "Email: {email}\n"
                    "City: {city}\n"
                    "Accommodation: {accommodation}\n\n"
                    "Program Details:\n\n"
                    "Event: 3-Day Spiritual Program Anantapur\n"
                    "Venue: Revenue Kalyana Mandapam (Revenue Bhavan), Beside Krishna Kalamandir, Near Clock Tower, Anantapur, Andhra Pradesh\n"
                    "Dates: July 24-26, 2026\n"
                    "Venue Location: https://www.google.com/maps/place/MHJW%2BQGV+Krishna+Kala+Mandir,+near+Clock+Tower,+Kamalanagar,+Anantapur,+Andhra+Pradesh+515001/\n\n"
                    "May this sacred gathering fill your heart with peace, devotion, positivity, and spiritual upliftment. We sincerely thank you for choosing to be part of this divine journey.\n\n"
                    "Please carry your registration confirmation during your visit. Further updates and instructions will be shared soon.\n\n"
                    "We look forward to welcoming you with love and prayers.\n\n"
                    "Jai Guru"
                )
            },
            {
                'key': 'room_allot',
                'description': 'Sent to devotees when room numbers are assigned or changed.',
                'variables': 'name,reg_id,room_number,arrival_date,departure_date',
                'template_text': (
                    "Dear {name},\n\n"
                    "With divine blessings, we are happy to inform you that your accommodation has been successfully allotted for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
                    "Accommodation Details:\n\n"
                    "Name: {name}\n"
                    "Room Number: {room_number}\n"
                    "Check-In Date: {arrival_date}\n"
                    "Check-Out Date: {departure_date}\n\n"
                    "We kindly request you to carry your registration confirmation during your visit and maintain the peaceful and spiritual atmosphere throughout the program.\n\n"
                    "May this sacred gathering bring peace, devotion, joy, and spiritual upliftment into your life.\n\n"
                    "We look forward to welcoming you with love and prayers.\n\n"
                    "Jai Guru"
                )
            },
            {
                'key': 'reminder_7d',
                'description': 'Sent 7 days before the event (July 17, 2026).',
                'variables': 'name,reg_id',
                'template_text': (
                    "Dear {name},\n\n"
                    "Jai Guru!\n\n"
                    "This is a loving reminder that the 3-Day Spiritual Program in Anantapur starts in exactly 7 days, on July 24, 2026! We are eagerly looking forward to meditating and serving together.\n\n"
                    "Venue Details:\n"
                    "Revenue Kalyana Mandapam (Revenue Bhavan), Beside Krishna Kalamandir, Near Clock Tower, Anantapur, Andhra Pradesh\n"
                    "Location: https://www.google.com/maps/place/MHJW%2BQGV+Krishna+Kala+Mandir,+near+Clock+Tower,+Kamalanagar,+Anantapur,+Andhra+Pradesh+515001/\n\n"
                    "Please complete your travel arrangements. If you need any assistance, feel free to reply to this message.\n\n"
                    "In divine friendship,\n"
                    "YSS Anantapur Team"
                )
            },
            {
                'key': 'reminder_3d',
                'description': 'Sent 3 days before the event (July 21, 2026).',
                'variables': 'name,reg_id',
                'template_text': (
                    "Dear {name},\n\n"
                    "Jai Guru!\n\n"
                    "Only 3 days left until our sacred 3-Day Spiritual Program begins on July 24, 2026.\n\n"
                    "Important Checklist for Your Visit:\n"
                    "1. Carrying registration ID: {reg_id}\n"
                    "2. Bring your personal meditation shawl or cushion if preferred.\n"
                    "3. Keep loose, comfortable clothing (preferably white or light colors).\n"
                    "4. Accommodation check-in begins on July 23rd afternoon.\n\n"
                    "For any urgent queries, contact us at 9441665181 or 8019682209.\n\n"
                    "Warm regards,\n"
                    "YSS Anantapur Team"
                )
            },
            {
                'key': 'reminder_1d',
                'description': 'Sent 1 day before the event (July 23, 2026).',
                'variables': 'name,reg_id',
                'template_text': (
                    "Dear {name},\n\n"
                    "Jai Guru!\n\n"
                    "The 3-Day Spiritual Program starts TOMORROW at 9:00 AM!\n\n"
                    "Please ensure you arrive at the venue by 8:00 AM for check-in and seating.\n"
                    "Venue: Revenue Kalyana Mandapam, Anantapur.\n\n"
                    "Please show this message or your Registration ID: {reg_id} at the reception desk to collect your entry badge.\n\n"
                    "Safe travels! We pray for a deeply uplifting spiritual experience for you.\n\n"
                    "In Master's Service,\n"
                    "YSS Anantapur Team"
                )
            },
            {
                'key': 'non_kriyaban_info',
                'description': 'Sent to non-Kriyaban members upon registration approval.',
                'variables': '',
                'template_text': (
                    "*సాధనా సంగం 2026 అనంతపురం*\n\n"
                    "*క్రియాయోగ దీక్ష*\n"
                    "*తీసుకోదలచిన వై.ఎస్.ఎస్*\n"
                    "*సభ్యులకు సూచనలు*\n\n"
                    "1.  సాధనాసంగం చివరిరోజు అనగా జూలై 26వ తేదీన క్రియాయోగ దీక్షా కార్యక్రమం నిర్వహించబడుతుంది.\n\n"
                    "2.  దీక్ష తీసుకోదలచిన సభ్యులు వై.ఎస్.ఎస్ రాంచీ ద్వారా లభ్యమయ్యే 18 పాఠాలు పొంది ఉండాలి.\n\n"
                    "3.  ఈ పాఠాల ద్వారా పొందిన ప్రశ్నావళిని(Step-I & Step-II forms) పూర్తిచేసి రాంచీకి పంపి ఉండాలి. ప్రశ్నావళినీ ఇంకా పంపనివారు, పూర్తిచేసి అనంతపురం ధ్యానకేంద్రం ఆఫీసులో కూడా అందచేయవచ్చు.\n\n"
                    "4.  సాధనా సంఘం మొదలయ్యే రోజుకు పాఠాలు అందడం చివరి దశలో ఉన్నవారు, క్రియాదీక్ష తీసుకొనదలచినచో, పూర్తి చేసిన ప్రశ్నావళితో స్వామీజీని కలిసి ప్రత్యేక  అనుమతి తీసుకోవలసి ఉంటుంది.\n\n"
                    "5.  అనంతపురంలో జరగబోయే సాధనాసంగంలో క్రియాయోగదీక్ష తీసుకోదలచిన సభ్యులు ముందుగానే అనంతపురం ధ్యానకేంద్రం ఆఫీసునందు లేదా శ్రీ A. నరసింహులు (సెల్ నం. 9441665181) గారికి గాని తెలియపరచవలసినదిగా ప్రార్థన.\n\n\n"
                    "  దివ్య స్నేహంలో,\n\n"
                    "            మేనేజింగ్ కమిటీ\n"
                    "            యోగదా సత్సంగ\n"
                    "            ధ్యాన కేంద్రం,\n"
                    "            అనంతపురం.\n\n\n"
                    "Sadhana Sangam 2026 – Anantapur\n\n"
                    "Instructions for YSS Members Wishing to Receive Kriya Yoga Initiation\n"
                    "The Kriya Yoga Initiation Ceremony will be conducted on the last day of Sadhana Sangam, July 26, 2026.\n"
                    "Members wishing to receive initiation should have obtained and studied the 18 Lessons provided by Yogoda Satsanga Society of India (YSS) Ranchi.\n"
                    "The questionnaires received through these lessons (Step-I and Step-II Forms) should be completed and submitted to Ranchi. Those who have not yet submitted the questionnaires may complete them and submit them at the Anantapur Meditation Centre Office.\n"
                    "Those who are in the final stages of receiving or completing the lessons by the commencement of Sadhana Sangam and wish to receive Kriya Initiation should meet Swamiji personally and obtain special permission, along with their completed questionnaires.\n"
                    "Members intending to receive Kriya Yoga Initiation during the Sadhana Sangam at Anantapur are kindly requested to inform the Anantapur Meditation Centre Office in advance or contact A. Narasimhulu (Mobile: +91 9441665181).\n\n"
                    "In Divine Friendship,\n\n"
                    "Managing Committee\n"
                    "Yogoda Satsanga Society of India Meditation Centre\n"
                    "Anantapur"
                )
            }
        ]
        for t in templates_to_seed:
            existing = WhatsAppTemplate.query.filter_by(key=t['key']).first()
            if not existing:
                new_t = WhatsAppTemplate(
                    key=t['key'],
                    description=t['description'],
                    variables=t['variables'],
                    template_text=t['template_text']
                )
                db.session.add(new_t)
        db.session.commit()

        print("Database tables verified/created successfully.")
    except Exception as e:
        print("!!! DATABASE INITIALIZATION FAILED !!!")
        print(f"Error details: {e}")
        # On Render, if we can't connect to DB, the app should probably fail fast
        # but let's just log it for now so we can see it in the dashboard.

@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@login_manager.user_loader
def load_user(user_id):
    return Admin.query.get(int(user_id))

# ─── SEED DATABASE ────────────────────────────────────────────────────────────
def seed_data():
    # Create admin
    if not Admin.query.first():
        admin = Admin(
            lesson_no='00000',
            name='YSS Admin',
            email=app.config['ADMIN_EMAIL'],
            mobile='9441665181',
            is_main_admin=True
        )
        admin.set_password(app.config['ADMIN_PASSWORD'])
        db.session.add(admin)
        db.session.commit()
    else:
        # Force update existing main admin details in database on launch
        main_admin = Admin.query.filter_by(is_main_admin=True).first()
        if main_admin:
            main_admin.mobile = '9441665181'
            main_admin.set_password(app.config['ADMIN_PASSWORD'])
            db.session.commit()

    # Seed schedule
    if not EventSchedule.query.first():
        schedule_data = [
            # Day 1
            (1, 'Day 1 – 24 July 2026', '09:30 AM', '10:30 AM', 'Opening satsanga', 'talk', 1),
            (1, 'Day 1 – 24 July 2026', '10:30 AM', '12:00 PM', 'Review of Energization Exercises (YSS Lessons students only)', 'meditation', 2),
            (1, 'Day 1 – 24 July 2026', '02:30 PM', '03:30 PM', 'Review of Hong-Sau technique (YSS Lessons students only)', 'meditation', 3),
            (1, 'Day 1 – 24 July 2026', '04:30 PM', '06:15 PM', 'Energization Exercises and meditation', 'meditation', 4),
            (1, 'Day 1 – 24 July 2026', '06:30 PM', '07:30 PM', 'Spiritual discourse', 'talk', 5),
            (1, 'Day 1 – 24 July 2026', '07:30 PM', '08:30 PM', 'Video show on Guruji', 'talk', 6),
            
            # Day 2
            (2, 'Day 2 – 25 July 2026', '06:00 AM', '09:00 AM', 'Energization Exercises and meditation', 'meditation', 1),
            (2, 'Day 2 – 25 July 2026', '09:45 AM', '10:45 AM', 'Spiritual discourse', 'talk', 2),
            (2, 'Day 2 – 25 July 2026', '11:00 AM', '12:00 PM', 'Review of Aum technique (YSS Lessons students only)', 'meditation', 3),
            (2, 'Day 2 – 25 July 2026', '02:30 PM', '03:30 PM', 'Question-Answer session', 'talk', 4),
            (2, 'Day 2 – 25 July 2026', '04:30 PM', '06:15 PM', 'Energization Exercises and meditation', 'meditation', 5),
            (2, 'Day 2 – 25 July 2026', '06:30 PM', '07:30 PM', 'Spiritual discourse', 'talk', 6),
            
            # Day 3
            (3, 'Day 3 – 26 July 2026', '08:00 AM', '12:00 PM', 'Kriya Yoga diksha (for eligible YSS devotees only)', 'meditation', 1),
            (3, 'Day 3 – 26 July 2026', '02:00 PM', '03:30 PM', 'Closing satsanga and prasad', 'food', 2),
            (3, 'Day 3 – 26 July 2026', '04:00 PM', '06:00 PM', 'Kriya Yoga review and check-up (YSS Kriyabans only)', 'meditation', 3),
        ]
        for s in schedule_data:
            item = EventSchedule(
                day_number=s[0], day_label=s[1], start_time=s[2],
                end_time=s[3], activity=s[4], category=s[5], sort_order=s[6]
            )
            db.session.add(item)
    db.session.commit()

    # Seed Rooms
    if not Room.query.first():
        # 10 rooms with 4 beds
        for i in range(1, 11):
            db.session.add(Room(room_number=f"Room {i}", capacity=4))
        # 25 rooms with 3 beds
        for i in range(11, 36):
            db.session.add(Room(room_number=f"Room {i}", capacity=3))
        db.session.commit()

    # Seed WhatsApp Templates
    if not WhatsAppTemplate.query.first():
        templates = [
            {
                "key": "reg_success",
                "description": "Sent to devotees upon registration approval and payment validation.",
                "variables": "name, reg_id, phone, email, city, accommodation",
                "template_text": (
                    "Dear {name},\n\n"
                    "With divine blessings and heartfelt joy, we are happy to confirm your successful registration for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
                    "Your Registration Details:\n\n"
                    "Name: {name}\n"
                    "Registration ID: {reg_id}\n"
                    "Phone Number: {phone}\n"
                    "Email: {email}\n"
                    "City: {city}\n"
                    "Accommodation: {accommodation}\n\n"
                    "Program Details:\n\n"
                    "Event: 3-Day Spiritual Program Anantapur\n"
                    "Venue: Revenue Kalyana Mandapam (Revenue Bhavan), Beside Krishna Kalamandir, Near Clock Tower, Anantapur, Andhra Pradesh\n"
                    "Dates: July 24-26, 2026\n"
                    "Venue Location: https://www.google.com/maps/place/MHJW%2BQGV+Krishna+Kala+Mandir,+near+Clock+Tower,+Kamalanagar,+Anantapur,+Andhra+Pradesh+515001/\n\n"
                    "May this sacred gathering fill your heart with peace, devotion, positivity, and spiritual upliftment.\n\n"
                    "Please carry your registration confirmation during your visit.\n\n"
                    "Jai Guru"
                )
            },
            {
                "key": "room_allot",
                "description": "Sent to devotees when accommodation rooms are assigned.",
                "variables": "name, reg_id, room_number, arrival_date, departure_date",
                "template_text": (
                    "Dear {name},\n\n"
                    "With divine blessings, we are happy to inform you that your accommodation has been successfully allotted for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
                    "Accommodation Details:\n\n"
                    "Name: {name}\n"
                    "Registration ID: {reg_id}\n"
                    "Room Number: {room_number}\n"
                    "Check-In Date: {arrival_date}\n"
                    "Check-Out Date: {departure_date}\n\n"
                    "We kindly request you to carry your registration confirmation during your visit and maintain the peaceful and spiritual atmosphere throughout the program.\n\n"
                    "Jai Guru"
                )
            },
            {
                "key": "reminder_7d",
                "description": "Reminder sent 7 days before the event starts.",
                "variables": "name, reg_id",
                "template_text": (
                    "Dear {name},\n\n"
                    "Jai Guru! Just 7 days left for our sacred 3-Day Spiritual Program at Anantapur (starting 24 July 2026).\n\n"
                    "We request you to prepare for your travel. Please ensure to check your assigned accommodation details (if registered for lodging) and keep your registration ID {reg_id} handy.\n\n"
                    "Looking forward to welcoming you soon.\n\n"
                    "In Divine Friendship,\nYSS Anantapur Committee"
                )
            },
            {
                "key": "reminder_3d",
                "description": "Reminder sent 3 days before the event starts.",
                "variables": "name, reg_id",
                "template_text": (
                    "Dear {name},\n\n"
                    "Jai Guru! Only 3 days left until the start of the 3-Day Spiritual Program at Anantapur.\n\n"
                    "Venue: Revenue Kalyana Mandapam (Revenue Bhavan), Beside Krishna Kalamandir, Near Clock Tower, Anantapur\n\n"
                    "Please pack light spiritual attire and carry your registration details. May Guruji's blessings guide your journey!\n\n"
                    "Jai Guru"
                )
            },
            {
                "key": "reminder_1d",
                "description": "Final reminder sent 1 day before the event starts.",
                "variables": "name, reg_id",
                "template_text": (
                    "Dear {name},\n\n"
                    "Jai Guru! The YSS Anantapur 3-Day Spiritual Program begins tomorrow, 24 July 2026, at 9:30 AM!\n\n"
                    "Please plan to arrive by 8:30 AM to complete check-in and room assignment validations smoothly.\n\n"
                    "Let us gather in love and devotion.\n\n"
                    "Jai Guru"
                )
            }
        ]
        for t in templates:
            db.session.add(WhatsAppTemplate(**t))
        db.session.commit()

# ─── EXCEL HELPERS ────────────────────────────────────────────────────────────
def update_registrations_excel():
    path = os.path.join(app.config['EXPORTS_DIR'], 'registrations.xlsx')
    wb = openpyxl.Workbook()
    regs = Registration.query.order_by(Registration.id).all()
    
    # Stats
    total = len(regs)
    kriyabans = len([r for r in regs if r.is_kriyaban])
    non_kriyabans = total - kriyabans
    acco_yes = len([r for r in regs if r.accommodation])
    acco_no = total - acco_yes

    def create_sheet(name, data_list):
        ws = wb.create_sheet(title=name)
        # Summary table only on 'All Registrations'
        if name == 'All Registrations':
            ws.append(['SUMMARY STATISTICS'])
            ws.append(['Category', 'Count'])
            ws.append(['Total Registrations', total])
            ws.append(['Kriyabans', kriyabans])
            ws.append(['Non-Kriyabans', non_kriyabans])
            ws.append(['Accommodation Needed', acco_yes])
            ws.append(['Accommodation Not Needed', acco_no])
            ws.append([]) # Spacer row
        
        headers = ['S.No','Reg ID','Lesson No','Full Name','Gender','Age','WhatsApp','Email','City/Town','District','State',
                   'Kriya','Accommodation','Volunteer','Arrival','Departure',
                   'Amount','Payment Mode','Transaction ID','Status','Date']
        ws.append(headers)
        for i, r in enumerate(data_list, 1):
            # For admin entries, some fields should be hyphenated
            is_admin = r.lesson_no == 'ADMIN' or r.gender == '-'
            ws.append([i, r.reg_id or '-', r.lesson_no or '-', r.full_name or '-', 
                       ('-' if is_admin else r.gender) or '-', 
                       ('-' if is_admin or r.age==0 else r.age) or '-', 
                       r.whatsapp or '-', 
                       ('-' if is_admin else r.email) or '-', 
                       r.place or '-', 
                       r.district or '-',
                       ('-' if is_admin else r.state) or '-',
                       ('-' if is_admin else ('Kriyaban' if r.is_kriyaban else 'Non-Kriyaban')),
                       ('Yes' if r.accommodation else 'No'), 
                       ('-' if is_admin else ('Yes' if r.volunteer else 'No')),
                       ('-' if is_admin else (r.arrival_date or '-')), 
                       ('-' if is_admin else (r.departure_date or '-')),
                       r.amount or 0, r.payment_mode or '-', r.transaction_id or '-', 
                       r.reg_status or '-',
                       r.created_at.strftime('%d-%m-%Y %I:%M %p') if r.created_at else '-'])
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column].width = min(max_length + 2, 40)

    # Remove default sheet
    del wb['Sheet']
    
    create_sheet('All Registrations', regs)
    create_sheet('Kriyabans', [r for r in regs if r.is_kriyaban])
    create_sheet('Non-Kriyabans', [r for r in regs if not r.is_kriyaban])
    create_sheet('Accommodation', [r for r in regs if r.accommodation])

    try:
        wb.save(path)
    except Exception as e:
        app.logger.error(f"Error saving registrations excel: {e}")
    finally:
        wb.close()

def update_donations_excel():
    path = os.path.join(app.config['EXPORTS_DIR'], 'donations.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Donations'
    headers = ['S.No','Donation ID','Lesson No','Name','Age','Place','WhatsApp',
               'Amount (₹)','Payment Mode','Transaction ID','Payment Screenshot','Payment Status','Date']
    ws.append(headers)
    dons = Donation.query.order_by(Donation.id).all()
    for i, d in enumerate(dons, 1):
        ws.append([i, d.donation_id or '-', d.lesson_no or '-', d.name or '-', d.age or 0, d.place or '-',
                   d.whatsapp or '-', d.amount or 0, d.payment_mode or '-', d.transaction_id or '-',
                   d.payment_screenshot or '-', d.payment_status or '-',
                   d.created_at.strftime('%d-%m-%Y') if d.created_at else '-'])
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = max_length + 2

    try:
        wb.save(path)
    except Exception as e:
        app.logger.error(f"Error saving donations excel: {e}")
    finally:
        wb.close()

def send_registration_email(reg):
    try:
        # Build the exact message body requested by the user
        body_text = (
            f"Dear {reg.full_name},\n\n"
            f"With divine blessings and heartfelt joy, we are happy to confirm your successful registration for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
            f"Your Registration Details:\n\n"
            f"Name: {reg.full_name}\n"
            f"Phone Number: {reg.whatsapp}\n"
            f"Email: {reg.email}\n"
            f"City: {reg.place}\n"
            f"Accommodation: {'Yes' if reg.accommodation else 'No'}\n\n"
            f"Program Details:\n"
            f"Event: {app.config['EVENT_NAME']}\n"
            f"Venue: {app.config['EVENT_VENUE']}\n"
            f"Dates: {app.config['EVENT_DATES']}\n\n"
            f"Venue Location:\n"
            f"https://www.google.com/maps/place/MHJW%2BQGV+Krishna+Kala+Mandir,+near+Clock+Tower,+Kamalanagar,+Anantapur,+Andhra+Pradesh+515001/\n\n"
            f"May this sacred gathering fill your heart with peace, devotion, positivity, and spiritual upliftment. We sincerely thank you for choosing to be part of this divine journey.\n\n"
            f"Please carry your registration confirmation during your visit. Further updates and instructions will be shared soon.\n\n"
            f"We look forward to welcoming you with love and prayers.\n\n"
            f"Jai Guru"
        )
        
        msg = Message(
            subject='Successful Registration Confirmation – YSS Anantapur',
            recipients=[reg.email],
            body=body_text
        )
        
        # Generate the individual PDF ID card and attach it
        from fpdf import FPDF
        pdf = FPDF(orientation='P', unit='mm', format=(85, 120))
        pdf.add_page()
        pdf.set_margin(0)
        
        # Draw borders / Header background
        pdf.set_fill_color(181, 51, 10) # Terracotta #B5330A
        pdf.rect(0, 0, 85, 25, 'F')
        
        # Gold stripe
        pdf.set_fill_color(212, 175, 55) # Gold #D4AF37
        pdf.rect(0, 25, 85, 2, 'F')
        
        # Header Text
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('helvetica', 'B', 11)
        pdf.cell(85, 12, 'Y.S.D.K., ANANTAPUR', 0, 1, 'C')
        pdf.set_font('helvetica', 'B', 9)
        pdf.cell(85, 0, '3-DAY SADHANA SANGAM', 0, 1, 'C')
        
        # Registration ID Block
        pdf.ln(18)
        pdf.set_text_color(181, 51, 10)
        pdf.set_font('helvetica', 'B', 13)
        pdf.cell(85, 6, f"ID: {reg.reg_id}", 0, 1, 'C')
        
        pdf.set_draw_color(181, 51, 10)
        pdf.line(10, 41, 75, 41)
        
        # Details
        pdf.ln(8)
        pdf.set_text_color(51, 51, 51)
        
        def add_row(label, val):
            pdf.set_font('helvetica', 'B', 9)
            pdf.cell(30, 8, f"  {label}", 0, 0, 'L')
            pdf.set_font('helvetica', '', 9)
            pdf.cell(55, 8, str(val), 0, 1, 'L')
            
        add_row('Name:', reg.full_name)
        add_row('Lesson No:', reg.lesson_no)
        add_row('Mobile No:', reg.whatsapp)
        add_row('City/Town:', reg.place)
        add_row('Acco Needed:', 'Yes' if reg.accommodation else 'No')
        
        # Footer
        pdf.ln(8)
        pdf.set_font('helvetica', 'I', 7)
        pdf.set_text_color(153, 153, 153)
        pdf.cell(85, 4, 'Scan QR at check-in counter upon arrival', 0, 1, 'C')
        pdf.set_font('helvetica', 'B', 9)
        pdf.set_text_color(181, 51, 10)
        pdf.cell(85, 5, 'Jai Guru', 0, 1, 'C')
        
        # Attach individual PDF ID card
        msg.attach(f"YSS_ID_Card_{reg.reg_id}.pdf", "application/pdf", pdf.output())
        
        mail.send(msg)
        print(f"REGISTRATION EMAIL WITH ID CARD SENT TO {reg.email}")
    except Exception as e:
        app.logger.warning(f'Email send failed: {e}')

def send_submission_email(reg):
    if not reg.email:
        return
    try:
        body_text = (
            f"Dear {reg.full_name},\n\n"
            f"Jai Guru! We have received your registration submission for the YSS 3-Day Spiritual Program in Anantapur.\n\n"
            f"Your Submission Details:\n"
            f"• Name: {reg.full_name}\n"
            f"• Phone Number: {reg.whatsapp}\n"
            f"• Email: {reg.email}\n"
            f"• City/Town: {reg.place}\n"
            f"• Accommodation Needed: {'Yes' if reg.accommodation else 'No'}\n\n"
            f"Our admin team is currently verifying your payment details. Once verified and approved, you will receive your official printable entry ID Card and program confirmation email.\n\n"
            f"Thank you for your patience.\n\n"
            f"Jai Guru"
        )
        msg = Message(
            subject='Registration Submission Received – YSS Anantapur',
            sender=app.config.get('MAIL_DEFAULT_SENDER'),
            recipients=[reg.email],
            body=body_text
        )
        mail.send(msg)
        print(f"REGISTRATION SUBMISSION RECEIVED EMAIL SENT TO {reg.email}")
    except Exception as e:
        app.logger.warning(f'Submission email send failed: {e}')

def send_admin_email_alert(reg):
    """
    Sends an email notification to the Admin.
    """
    admin_email = app.config.get('ADMIN_EMAIL', app.config.get('MAIL_USERNAME'))
    try:
        msg = Message(
            subject=f'New Registration Alert - {reg.full_name}',
            sender=app.config.get('MAIL_DEFAULT_SENDER'),
            recipients=[admin_email],
            body=f"Jai Guru!\n\nA new registration has been submitted by {reg.full_name}.\n\nPlease check the admin panel for approval.\n\nRegards,\nYSS Spiritual Program System"
        )
        mail.send(msg)
        print(f"EMAIL SENT TO ADMIN ({admin_email}) for Reg ID: {reg.reg_id}")
    except Exception as e:
        app.logger.warning(f"Failed to send admin email: {e}")

def update_registrations_excel_async():
    import threading
    def job():
        with app.app_context():
            update_registrations_excel()
    threading.Thread(target=job, daemon=True).start()

def send_submission_email_async(reg_id):
    import threading
    def job():
        with app.app_context():
            reg = Registration.query.get(reg_id)
            if reg:
                send_submission_email(reg)
    threading.Thread(target=job, daemon=True).start()

def send_admin_email_alert_async(reg_id):
    import threading
    def job():
        with app.app_context():
            reg = Registration.query.get(reg_id)
            if reg:
                send_admin_email_alert(reg)
    threading.Thread(target=job, daemon=True).start()

def send_member_whatsapp_async(reg_id):
    import threading
    def job():
        with app.app_context():
            reg = Registration.query.get(reg_id)
            if reg:
                try:
                    send_member_whatsapp(reg)
                except Exception as e:
                    print(f"Async WhatsApp send failed for reg {reg_id}: {e}")
                    reg.notified = False
                    db.session.commit()
    threading.Thread(target=job, daemon=True).start()

def send_room_allotment_whatsapp_async(reg_id):
    import threading
    def job():
        with app.app_context():
            allotment = RoomAllotment.query.filter_by(registration_id=reg_id).first()
            if allotment and allotment.room:
                reg = allotment.registration
                room = allotment.room
                gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
                if not gateway_url:
                    print(f"Async room allotment WhatsApp send failed for reg {reg_id}: WhatsApp Gateway is not configured.")
                    return
                try:
                    body_text = format_whatsapp_template(
                        'room_allot',
                        name=reg.full_name,
                        reg_id=reg.reg_id,
                        room_number=room.room_number,
                        arrival_date=reg.arrival_date if reg.arrival_date else '24-07-2026',
                        departure_date=reg.departure_date if reg.departure_date else '26-07-2026'
                    )
                    if not body_text:
                        body_text = (
                            f"Dear {reg.full_name},\n\n"
                            f"With divine blessings, we are happy to inform you that your accommodation has been successfully allotted for the 3-Day Spiritual Program at Anantapur.\n\n"
                            f"Accommodation Details:\n\n"
                            f"Name: {reg.full_name}\n"
                            f"Room Number: {room.room_number}\n"
                            f"Check-In Date: {reg.arrival_date if reg.arrival_date else '24-07-2026'}\n"
                            f"Check-Out Date: {reg.departure_date if reg.departure_date else '26-07-2026'}\n\n"
                            f"Jai Guru"
                        )
                    import requests
                    r = requests.post(
                        f"{gateway_url}/send",
                        json={
                            'to': reg.whatsapp,
                            'message': body_text
                        },
                        timeout=10
                    )
                    if r.status_code == 200:
                        allotment.notified_room_whatsapp = room.room_number
                        allotment.notified_room_number = room.room_number
                        db.session.commit()
                        print(f"Async room allotment WhatsApp send succeeded for reg {reg_id}")
                    else:
                        print(f"Async room allotment WhatsApp send returned status {r.status_code}: {r.text}")
                except Exception as e:
                    print(f"Async room allotment WhatsApp send failed for reg {reg_id}: {e}")
    threading.Thread(target=job, daemon=True).start()

def send_registration_email_async(reg_id):
    import threading
    def job():
        with app.app_context():
            reg = Registration.query.get(reg_id)
            if reg:
                send_registration_email(reg)
    threading.Thread(target=job, daemon=True).start()

def log_action(action_desc):
    """
    Saves an entry into ActivityLog with details of the current logged-in admin.
    """
    try:
        admin_id = None
        admin_name = 'System'
        if current_user and current_user.is_authenticated:
            admin_id = current_user.id
            admin_name = current_user.name
            
            # Update last_active timestamp
            try:
                current_user.last_active = datetime.utcnow()
                db.session.commit()
            except Exception as ex:
                db.session.rollback()
                print(f"Error updating admin last active: {ex}")
                
        ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if ip and ',' in ip:
            ip = ip.split(',')[0].strip()
            
        log_entry = ActivityLog(
            admin_id=admin_id,
            admin_name=admin_name,
            action=action_desc,
            ip_address=ip
        )
        db.session.add(log_entry)
        db.session.commit()
    except Exception as e:
        print(f"Failed to write ActivityLog: {e}")

@app.before_request
def update_last_active():
    """Heartbeat: update last_active on every authenticated page request."""
    if current_user and current_user.is_authenticated:
        try:
            current_user.last_active = datetime.utcnow()
            db.session.commit()
        except Exception:
            db.session.rollback()

def format_whatsapp_template(key, **kwargs):
    """
    Retrieves a template by key, performs variable substitution, and returns the formatted text.
    If the template is not found, falls back to empty string.
    """
    t = WhatsAppTemplate.query.filter_by(key=key).first()
    if not t:
        return ""
    text = t.template_text
    import re
    def replace_var(match):
        var_name = match.group(1)
        return str(kwargs.get(var_name, match.group(0)))
    return re.sub(r'\{([a-zA-Z0-9_]+)\}', replace_var, text)

def send_member_whatsapp(reg):
    """
    Sends a WhatsApp message to the member after Admin approval using db template.
    """
    message = format_whatsapp_template(
        'reg_success',
        name=reg.full_name,
        reg_id=reg.reg_id,
        phone=reg.whatsapp,
        email=reg.email or 'N/A',
        city=reg.place,
        accommodation='Yes' if reg.accommodation else 'No'
    )
    if not message:
        print(f"WHATSAPP ERROR: Template reg_success not found in database.")
        raise Exception("WhatsApp template 'reg_success' not found in database.")
        
    print(f"WHATSAPP LOG: {message}")
    
    # Try sending via self-hosted gateway
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    if not gateway_url:
        raise Exception("WhatsApp Gateway is not configured.")
        
    r = requests.post(
        f"{gateway_url}/send",
        json={
            'to': reg.whatsapp,
            'message': message
        },
        timeout=10
    )
    print(f"AUTOMATED WHATSAPP STATUS: {r.status_code} - {r.text}")
    if r.status_code != 200:
        raise Exception(f"WhatsApp gateway returned status {r.status_code}: {r.text}")

    # If the user is a non-Kriyaban, send the initiation instructions message
    if not reg.is_kriyaban:
        non_kri_message = format_whatsapp_template('non_kriyaban_info')
        if non_kri_message:
            import time
            time.sleep(1.5)  # brief delay to avoid out-of-order delivery
            print("WHATSAPP LOG (Non-Kriyaban): Sending initiation instructions...")
            r_non = requests.post(
                f"{gateway_url}/send",
                json={
                    'to': reg.whatsapp,
                    'message': non_kri_message
                },
                timeout=10
            )
            print(f"AUTOMATED NON-KRIYABAN WHATSAPP STATUS: {r_non.status_code} - {r_non.text}")
            if r_non.status_code != 200:
                raise Exception(f"WhatsApp gateway returned status {r_non.status_code} for non-kriyaban message: {r_non.text}")


# ─── PUBLIC ROUTES ────────────────────────────────────────────────────────────
@app.route('/sw.js')
def serve_sw():
    response = make_response(send_file(os.path.join(app.root_path, 'static', 'sw.js')))
    response.headers['Content-Type'] = 'application/javascript'
    response.headers['Service-Worker-Allowed'] = '/'
    return response

@app.route('/manifest.json')
def serve_manifest():
    response = make_response(send_file(os.path.join(app.root_path, 'static', 'manifest.json')))
    response.headers['Content-Type'] = 'application/json'
    return response

@app.route('/offline')
def offline():
    return render_template('offline.html')

@app.route('/debug-db')
def debug_db():
    from sqlalchemy import inspect
    res = {}
    
    # Run migrations explicitly
    def ensure_column_explicit_on_engine(engine, insp, table_name, col_name, col_type_sql, default_sql=None):
        if not insp or not insp.has_table(table_name):
            return f"Table {table_name} not found on {engine.name if hasattr(engine, 'name') else 'engine'}"
        columns = [c['name'] for c in insp.get_columns(table_name)]
        if col_name not in columns:
            with engine.begin() as conn:
                try:
                    sql = f"ALTER TABLE {table_name} ADD COLUMN IF NOT EXISTS {col_name} {col_type_sql}"
                    if default_sql is not None:
                        sql += f" DEFAULT {default_sql}"
                    conn.execute(db.text(sql))
                    return f"Added column {col_name} to {table_name} on {engine.name if hasattr(engine, 'name') else 'engine'} (Postgres)"
                except Exception as e_pg:
                    try:
                        sql = f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type_sql}"
                        if default_sql is not None:
                            sql += f" DEFAULT {default_sql}"
                        conn.execute(db.text(sql))
                        return f"Added column {col_name} to {table_name} on {engine.name if hasattr(engine, 'name') else 'engine'} (SQLite/Fallback)"
                    except Exception as e_sql:
                        return f"Failed to add column {col_name} to {table_name} on {engine.name if hasattr(engine, 'name') else 'engine'}: PG={e_pg}, SQL={e_sql}"
        return f"Column {col_name} already exists on {table_name} on {engine.name if hasattr(engine, 'name') else 'engine'}"

    inspector = inspect(db.engine)
    migration_results = []
    
    cols_to_add = [
        ('registrations', 'state', 'VARCHAR(100)', None),
        ('registrations', 'email', 'VARCHAR(120)', None),
        ('registrations', 'country_code', 'VARCHAR(10)', "'+91'"),
        ('registrations', 'amount', 'FLOAT', None),
        ('registrations', 'transaction_id', 'VARCHAR(100)', None),
        ('registrations', 'payment_screenshot', 'VARCHAR(255)', None),
        ('registrations', 'payment_status', 'VARCHAR(20)', "'Pending'"),
        ('registrations', 'reg_status', 'VARCHAR(20)', "'Pending'"),
        ('registrations', 'notified', 'BOOLEAN', 'FALSE'),
        ('registrations', 'district', 'VARCHAR(100)', None),
        ('registrations', 'reminder_7d_sent', 'BOOLEAN', 'FALSE'),
        ('registrations', 'reminder_3d_sent', 'BOOLEAN', 'FALSE'),
        ('registrations', 'reminder_1d_sent', 'BOOLEAN', 'FALSE'),
        ('registrations', 'registered_by_id', 'INTEGER REFERENCES admins(id)', None),
        ('registrations', 'registered_by_name', 'VARCHAR(100)', None),
        ('donations', 'transaction_id', 'VARCHAR(100)', None),
        ('donations', 'payment_screenshot', 'VARCHAR(255)', None),
        ('donations', 'payment_status', 'VARCHAR(20)', "'Pending'"),
        ('donations', 'notified', 'BOOLEAN', 'FALSE')
    ]
    
    # Run on primary engine
    for table, col, col_type, default in cols_to_add:
        migration_results.append(ensure_column_explicit_on_engine(db.engine, inspector, table, col, col_type, default))
        
    # Run on secondary engine if configured
    if secondary_engine:
        secondary_inspector = inspect(secondary_engine)
        for table, col, col_type, default in cols_to_add:
            migration_results.append(ensure_column_explicit_on_engine(secondary_engine, secondary_inspector, table, col, col_type, default))
            
    # Refresh inspectors and get columns
    inspector = inspect(db.engine)
    primary_columns = {}
    for table in ['registrations', 'donations', 'admins', 'room_allotments']:
        if inspector.has_table(table):
            primary_columns[table] = [c['name'] for c in inspector.get_columns(table)]
        else:
            primary_columns[table] = "Not found"
            
    secondary_columns = {}
    if secondary_engine:
        sec_inspector = inspect(secondary_engine)
        for table in ['registrations', 'donations', 'admins', 'room_allotments']:
            if sec_inspector.has_table(table):
                secondary_columns[table] = [c['name'] for c in sec_inspector.get_columns(table)]
            else:
                secondary_columns[table] = "Not found"
                
    return jsonify({
        "migration_results": migration_results,
        "primary_columns": primary_columns,
        "secondary_columns": secondary_columns if secondary_engine else "No secondary database configured"
    })

@app.route('/')
def index():
    return render_template('index.html', config=app.config)

@app.route('/about')
def about():
    return render_template('about.html', config=app.config)

@app.route('/gallery')
def gallery():
    images = GalleryImage.query.order_by(GalleryImage.created_at.desc()).all()
    return render_template('gallery.html', images=images, config=app.config)

@app.route('/schedule')
def schedule():
    days = {}
    schedules = EventSchedule.query.order_by(EventSchedule.day_number, EventSchedule.sort_order).all()
    for s in schedules:
        days.setdefault(s.day_number, {'label': s.day_label, 'items': []})['items'].append(s)
    return render_template('schedule.html', days=days, config=app.config)



@app.route('/contact', methods=['GET', 'POST'])
def contact():
    if request.method == 'POST':
        flash('Your message has been sent. We will get back to you soon!', 'success')
        return redirect(url_for('contact'))
    return render_template('contact.html', config=app.config)

def normalize_lesson_no(val):
    val = (val or '').strip()
    if not val:
        return ''
    if val.upper().startswith('NEW MEMBER'):
        return val
    # Remove any existing L- or L - prefix (case-insensitive)
    import re
    cleaned = re.sub(r'^L\s*-\s*', '', val, flags=re.IGNORECASE)
    cleaned = re.sub(r'^L-\s*', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'^L\s*', '', cleaned, flags=re.IGNORECASE)
    return f"L - {cleaned}"

# ─── REGISTRATION ─────────────────────────────────────────────────────────────
@app.route('/registration', methods=['GET', 'POST'])
def registration():
    if request.method == 'POST':
        print("DEBUG: [registration] POST request received")
        errors = []
        is_new_member = request.form.get('is_new_member') == 'yes'
        lesson_no = request.form.get('lesson_no', '').strip()
        full_name = request.form.get('full_name', '').strip()
        gender = request.form.get('gender', '').strip()
        age = request.form.get('age', '').strip()
        place = request.form.get('place', '').strip()
        district = request.form.get('district', '').strip()
        state = request.form.get('state', '').strip()
        email = request.form.get('email', '').strip()
        country_code = request.form.get('country_code', '+91')
        whatsapp = request.form.get('whatsapp', '').strip()
        is_kriyaban = request.form.get('is_kriyaban') == 'yes'
        accommodation = request.form.get('accommodation') == 'yes'
        volunteer = request.form.get('volunteer') == 'yes'
        arrival_date = request.form.get('arrival_date', '').strip()
        departure_date = request.form.get('departure_date', '').strip()
        payment_mode = request.form.get('payment_mode', '').strip()
        transaction_id = request.form.get('transaction_id', '').strip()
        
        print(f"DEBUG: [registration] Form data: full_name={full_name}, whatsapp={whatsapp}, lesson_no={lesson_no}")
        screenshot_filename = None
        file = request.files.get('payment_screenshot')
        if file and file.filename != '':
            print("DEBUG: [registration] Processing payment screenshot upload")
            import werkzeug.utils, uuid
            filename = werkzeug.utils.secure_filename(file.filename)
            file_ext = os.path.splitext(filename)[1]
            new_filename = f"screenshot_{uuid.uuid4().hex[:8]}{file_ext}"
            uploads_dir = os.path.join(app.root_path, 'static', 'uploads')
            os.makedirs(uploads_dir, exist_ok=True)
            file.save(os.path.join(uploads_dir, new_filename))
            screenshot_filename = new_filename
            print(f"DEBUG: [registration] Screenshot saved as {screenshot_filename}")

        # Set/Normalize Lesson Number
        if is_new_member:
            new_members = Registration.query.filter(Registration.lesson_no.like('NEW MEMBER %')).all()
            import re
            max_i = 0
            for r in new_members:
                match = re.match(r'^NEW MEMBER\s+(\d+)$', r.lesson_no, re.IGNORECASE)
                if match:
                    num = int(match.group(1))
                    if num > max_i:
                        max_i = num
            lesson_no = f"NEW MEMBER {max_i + 1}"
            print(f"DEBUG: [registration] Generated new member lesson no: {lesson_no}")
        else:
            lesson_no = normalize_lesson_no(lesson_no)
            print(f"DEBUG: [registration] Normalized lesson no: {lesson_no}")

        if not is_new_member and not lesson_no: errors.append('Lesson Number is required.')
        if not full_name: errors.append('Full Name is required.')
        if not gender: errors.append('Gender is required.')
        if not age or not age.isdigit(): errors.append('Valid Age is required.')
        if not place: errors.append('City/Village/Town is required.')
        if not state: errors.append('State is required.')
        if email and '@' not in email: errors.append('Valid Email Address is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not arrival_date: errors.append('Date of Arrival is required.')
        if not departure_date: errors.append('Date of Departure is required.')
        if not payment_mode: errors.append('Payment Mode is required.')
        
        print(f"DEBUG: [registration] Initial validation errors count: {len(errors)}")

        if not errors:
            existing_reg = Registration.query.filter(Registration.full_name.ilike(full_name), Registration.whatsapp == whatsapp).first()
            if existing_reg:
                errors.append('A registration with this Name and Mobile number already exists.')
                
            if payment_mode == 'UPI' and transaction_id:
                existing_txn = Registration.query.filter(Registration.transaction_id.ilike(transaction_id)).first()
                if existing_txn:
                    errors.append('This Transaction ID has already been submitted.')
                    
            if not is_new_member and lesson_no and lesson_no.upper() not in ['0', '00', '000', '0000', '00000', 'NA', 'N/A', '-', 'ADMIN', 'NONE', '']:
                existing_lesson = Registration.query.filter(Registration.lesson_no.ilike(lesson_no)).first()
                if existing_lesson:
                    errors.append(f'Lesson Number {lesson_no} is already registered.')
            print(f"DEBUG: [registration] DB check validation errors count: {len(errors)}")

        if errors:
            print("DEBUG: [registration] Validation failed, flashing errors")
            for e in errors:
                flash(e, 'error')
            form_data = dict(request.form)
            form_data['lesson_no'] = lesson_no
            return render_template('registration.html', config=app.config, form=form_data)

        # Calculate amount
        base_fee = 1800
        acc_fee = 1000 if accommodation else 0
        total_amount = base_fee + acc_fee

        print(f"DEBUG: [registration] Creating Registration object, amount={total_amount}")
        reg = Registration(
            lesson_no=lesson_no, full_name=full_name, gender=gender,
            age=int(age), place=place, district=district, state=state, email=email, country_code=country_code, whatsapp=whatsapp,
            is_kriyaban=is_kriyaban, accommodation=accommodation,
            volunteer=volunteer, arrival_date=arrival_date,
            departure_date=departure_date, payment_mode=payment_mode,
            amount=total_amount,
            transaction_id=transaction_id or None, payment_screenshot=screenshot_filename
        )
        print("DEBUG: [registration] Adding Registration to session")
        db.session.add(reg)
        print("DEBUG: [registration] Committing transaction")
        db.session.commit()
        print(f"DEBUG: [registration] Committed successfully! id={reg.id}, reg_id={reg.reg_id}")
        
        print("DEBUG: [registration] Triggering async Excel update")
        update_registrations_excel_async()
        print("DEBUG: [registration] Triggering async submission email")
        send_submission_email_async(reg.id)
        print("DEBUG: [registration] Triggering async admin alert")
        send_admin_email_alert_async(reg.id)
        
        print(f"DEBUG: [registration] Redirecting to success page for {reg.reg_id}")
        return redirect(url_for('reg_success', reg_id=reg.reg_id))

    return render_template('registration.html', config=app.config, form={})

@app.route('/registration/success/<reg_id>')
def reg_success(reg_id):
    reg = Registration.query.filter_by(reg_id=reg_id).first_or_404()
    return render_template('reg_success.html', reg=reg, config=app.config)

@app.route('/registration/card/<reg_id>')
def reg_card_only(reg_id):
    reg = Registration.query.filter_by(reg_id=reg_id).first_or_404()
    plain = request.args.get('plain') == '1'
    return render_template('id_card_only.html', reg=reg, config=app.config, plain=plain)

# ─── DONATION ─────────────────────────────────────────────────────────────────
@app.route('/donation', methods=['GET', 'POST'])
def donation():
    if request.method == 'POST':
        errors = []
        import time
        from werkzeug.utils import secure_filename
        lesson_no = request.form.get('lesson_no', '').strip()
        name = request.form.get('name', '').strip()
        age = request.form.get('age', '').strip()
        place = request.form.get('place', '').strip()
        whatsapp = request.form.get('whatsapp', '').strip()
        amount = request.form.get('amount', '').strip()
        payment_mode = request.form.get('payment_mode', '').strip()
        transaction_id = request.form.get('transaction_id', '').strip()

        if not lesson_no: errors.append('Lesson Number is required.')
        if not name: errors.append('Name is required.')
        if not age or not age.isdigit(): errors.append('Valid Age is required.')
        if not place: errors.append('Place is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not amount: errors.append('Donation Amount is required.')
        try:
            float(amount)
        except:
            errors.append('Valid Amount is required.')
        if not payment_mode: errors.append('Payment Mode is required.')

        screenshot_filename = None
        if 'payment_screenshot' in request.files:
            file = request.files['payment_screenshot']
            if file and file.filename:
                screenshot_filename = secure_filename(f"donation_{int(time.time())}_{file.filename}")
                os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], screenshot_filename))

        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template('donation.html', config=app.config, form=request.form)

        don = Donation(
            lesson_no=lesson_no, name=name, age=int(age), place=place,
            whatsapp=whatsapp, amount=float(amount), payment_mode=payment_mode,
            transaction_id=transaction_id or None, payment_screenshot=screenshot_filename
        )
        db.session.add(don)
        db.session.commit()
        update_donations_excel()
        return redirect(url_for('donation_success', don_id=don.donation_id))

    return render_template('donation.html', config=app.config, form={})

@app.route('/donation/success/<don_id>')
def donation_success(don_id):
    don = Donation.query.filter_by(donation_id=don_id).first_or_404()
    return render_template('don_success.html', don=don, config=app.config)

@app.route('/admin/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        admin = Admin.query.filter_by(email=email).first()
        if admin:
            token = uuid.uuid4().hex
            admin.reset_token = token
            admin.reset_token_expiry = datetime.utcnow() + timedelta(hours=1)
            db.session.commit()
            
            reset_url = url_for('reset_password', token=token, _external=True)
            msg = Message('Password Reset Request – YSS Anantapur',
                         sender=app.config['MAIL_USERNAME'],
                         recipients=[admin.email])
            msg.body = f"Jai Guru!\n\nTo reset your admin password, please click the link below:\n{reset_url}\n\nIf you did not request this, please ignore this email.\n\nRegards,\nYSS Spiritual Program Team"
            try:
                mail.send(msg)
                flash('A password reset link has been sent to your email.', 'success')
            except Exception as e:
                app.logger.warning(f'Forgot password email failed: {e}')
                flash('Failed to send password reset email. Please contact the main admin.', 'error')
        else:
            flash('Email address not found.', 'error')
    return render_template('admin/forgot_password.html', config=app.config)

@app.route('/admin/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    admin = Admin.query.filter_by(reset_token=token).first()
    if not admin or admin.reset_token_expiry < datetime.utcnow():
        flash('Invalid or expired reset token.', 'error')
        return redirect(url_for('admin_login'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        confirm = request.form.get('confirm_password')
        if password == confirm:
            admin.set_password(password)
            admin.reset_token = None
            admin.reset_token_expiry = None
            db.session.commit()
            flash('Password updated successfully. Please login.', 'success')
            return redirect(url_for('admin_login'))
        flash('Passwords do not match.', 'error')
    return render_template('admin/reset_password.html', config=app.config)

@app.route('/admin/manage', methods=['GET', 'POST'])
@login_required
def admin_manage():
    if not current_user.is_main_admin:
        flash('Only Main Admin can access this page.', 'error')
        return redirect(url_for('admin_registrations'))
    
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            lesson_no = request.form.get('lesson_no')
            name = request.form.get('name')
            email = request.form.get('email')
            mobile = request.form.get('mobile')
            password = request.form.get('password')
            
            if not name or not email or not mobile or not password:
                flash('Name, Email, Mobile, and Password are required.', 'error')
            else:
                try:
                    new_admin = Admin(lesson_no=lesson_no, name=name, email=email, mobile=mobile)
                    new_admin.set_password(password)
                    db.session.add(new_admin)
                    db.session.commit()
                    flash('New admin added successfully.', 'success')
                except Exception as e:
                    db.session.rollback()
                    flash(f'Error adding admin: {e}', 'error')
        elif action == 'delete':
            admin_id = request.form.get('admin_id')
            admin_to_del = Admin.query.get(int(admin_id))
            if admin_to_del and not admin_to_del.is_main_admin:
                db.session.delete(admin_to_del)
                db.session.commit()
                flash('Admin deleted.', 'success')
            else:
                flash('Cannot delete main admin.', 'error')
                
    admins = Admin.query.all()
    five_mins_ago = datetime.utcnow() - timedelta(minutes=5)
    
    # Per-admin stats
    from sqlalchemy import func
    reg_counts = dict(
        db.session.query(Registration.registered_by_id, func.count(Registration.id))
        .group_by(Registration.registered_by_id).all()
    )
    activity_counts = dict(
        db.session.query(ActivityLog.admin_id, func.count(ActivityLog.id))
        .group_by(ActivityLog.admin_id).all()
    )
    
    admin_stats = {}
    for a in admins:
        admin_stats[a.id] = {
            'reg_count': reg_counts.get(a.id, 0),
            'activity_count': activity_counts.get(a.id, 0),
            'is_online': bool(a.last_active and (datetime.utcnow() - a.last_active).total_seconds() < 300)
        }
    
    return render_template('admin/manage_admins.html', admins=admins, admin_stats=admin_stats, config=app.config)

# ─── ADMIN AUTH ───────────────────────────────────────────────────────────────
@app.route('/admin')
def admin_redirect():
    return redirect(url_for('admin_login'))

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        login_identifier = request.form.get('login_identifier', '').strip()
        password = request.form.get('password', '').strip()
        
        admin = Admin.query.filter(db.or_(Admin.name == login_identifier, Admin.email == login_identifier)).first()
        
        if admin and admin.check_password(password):
            login_user(admin)
            log_action("Admin logged in successfully")
            return redirect(url_for('admin_registrations'))
        flash('Invalid Admin Name or password.', 'error')
    return render_template('admin/login.html', config=app.config)

@app.route('/admin/logout')
@login_required
def admin_logout():
    log_action("Admin logged out")
    logout_user()
    flash('Logged out successfully.', 'success')
    return redirect(url_for('admin_login'))



# ─── ADMIN REGISTRATIONS ──────────────────────────────────────────────────────
@app.route('/admin/add-registration', methods=['GET', 'POST'])
@login_required
def admin_add_registration():
    if request.method == 'POST':
        errors = []
        lesson_no = request.form.get('lesson_no', '').strip()
        full_name = request.form.get('full_name', '').strip()
        gender = request.form.get('gender', '').strip()
        age = request.form.get('age', '').strip()
        place = request.form.get('place', '').strip()
        district = request.form.get('district', '').strip()
        state = request.form.get('state', '').strip()
        email = request.form.get('email', '').strip()
        country_code = request.form.get('country_code', '+91')
        whatsapp = request.form.get('whatsapp', '').strip()
        is_kriyaban = request.form.get('is_kriyaban') == 'yes'
        accommodation = request.form.get('accommodation') == 'yes'
        volunteer = request.form.get('volunteer') == 'yes'
        arrival_date = request.form.get('arrival_date', '').strip()
        departure_date = request.form.get('departure_date', '').strip()
        payment_mode = request.form.get('payment_mode', '').strip()
        transaction_id = request.form.get('transaction_id', '').strip()
        
        screenshot_filename = None
        file = request.files.get('payment_screenshot')
        if file and file.filename != '':
            import werkzeug.utils, uuid
            filename = werkzeug.utils.secure_filename(file.filename)
            file_ext = os.path.splitext(filename)[1]
            new_filename = f"screenshot_{uuid.uuid4().hex[:8]}{file_ext}"
            uploads_dir = os.path.join(app.root_path, 'static', 'uploads')
            os.makedirs(uploads_dir, exist_ok=True)
            file.save(os.path.join(uploads_dir, new_filename))
            screenshot_filename = new_filename

        if not lesson_no: errors.append('Lesson Number is required.')
        if not full_name: errors.append('Full Name is required.')
        if not gender: errors.append('Gender is required.')
        if not age or not age.isdigit(): errors.append('Valid Age is required.')
        if not place: errors.append('City/Village/Town is required.')
        if not state: errors.append('State is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not arrival_date: errors.append('Date of Arrival is required.')
        if not departure_date: errors.append('Date of Departure is required.')
        if not payment_mode: errors.append('Payment Mode is required.')

        if not errors:
            existing_reg = Registration.query.filter(Registration.full_name.ilike(full_name), Registration.whatsapp == whatsapp).first()
            if existing_reg:
                errors.append('A registration with this Name and Mobile number already exists.')
                
            if payment_mode == 'UPI' and transaction_id:
                existing_txn = Registration.query.filter(Registration.transaction_id.ilike(transaction_id)).first()
                if existing_txn:
                    errors.append('This Transaction ID has already been submitted.')
                    
            if lesson_no and lesson_no.upper() not in ['0', '00', '000', '0000', '00000', 'NA', 'N/A', '-', 'ADMIN', 'NONE', '']:
                existing_lesson = Registration.query.filter(Registration.lesson_no.ilike(lesson_no)).first()
                if existing_lesson:
                    errors.append(f'Lesson Number {lesson_no} is already registered.')

        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template('admin/add_registration.html', config=app.config, form=request.form)

        base_fee = 1800
        acc_fee = 1000 if accommodation else 0
        total_amount = base_fee + acc_fee

        reg = Registration(
            lesson_no=lesson_no, full_name=full_name, gender=gender,
            age=int(age), place=place, district=district, state=state, email=email, country_code=country_code, whatsapp=whatsapp,
            is_kriyaban=is_kriyaban, accommodation=accommodation,
            volunteer=volunteer, arrival_date=arrival_date,
            departure_date=departure_date, payment_mode=payment_mode,
            amount=total_amount,
            transaction_id=transaction_id, payment_screenshot=screenshot_filename,
            payment_status='Paid' if payment_mode == 'Cash' or transaction_id else 'Pending',
            reg_status='Approved' if payment_mode == 'Cash' or transaction_id else 'Pending',
            notified=False,
            registered_by_id=current_user.id,
            registered_by_name=current_user.name
        )
        db.session.add(reg)
        db.session.commit()
        update_registrations_excel_async()
        log_action(f"Admin added registration for {reg.full_name} (Lesson: {reg.lesson_no}, Mobile: {reg.whatsapp}, Mode: {reg.payment_mode}, Amount: {reg.amount})")
        if email and reg.payment_status == 'Paid':
            send_registration_email_async(reg.id)
        flash('Participant added successfully.', 'success')
        return redirect(url_for('admin_registrations'))

    return render_template('admin/add_registration.html', config=app.config, form={})

@app.route('/admin/edit-registration/<int:reg_id>', methods=['GET', 'POST'])
@login_required
def admin_edit_registration(reg_id):
    reg = Registration.query.get_or_404(reg_id)
    
    if request.method == 'POST':
        errors = []
        lesson_no = request.form.get('lesson_no', '').strip()
        full_name = request.form.get('full_name', '').strip()
        gender = request.form.get('gender', '').strip()
        age = request.form.get('age', '').strip()
        place = request.form.get('place', '').strip()
        district = request.form.get('district', '').strip()
        state = request.form.get('state', '').strip()
        email = request.form.get('email', '').strip()
        country_code = request.form.get('country_code', '+91')
        whatsapp = request.form.get('whatsapp', '').strip()
        is_kriyaban = request.form.get('is_kriyaban') == 'yes'
        accommodation = request.form.get('accommodation') == 'yes'
        volunteer = request.form.get('volunteer') == 'yes'
        arrival_date = request.form.get('arrival_date', '').strip()
        departure_date = request.form.get('departure_date', '').strip()
        payment_mode = request.form.get('payment_mode', '').strip()
        amount = request.form.get('amount', '').strip()
        transaction_id = request.form.get('transaction_id', '').strip()
        
        file = request.files.get('payment_screenshot')
        if file and file.filename != '':
            import werkzeug.utils, uuid, os
            filename = werkzeug.utils.secure_filename(file.filename)
            file_ext = os.path.splitext(filename)[1]
            new_filename = f"screenshot_{uuid.uuid4().hex[:8]}{file_ext}"
            uploads_dir = os.path.join(app.root_path, 'static', 'uploads')
            os.makedirs(uploads_dir, exist_ok=True)
            file.save(os.path.join(uploads_dir, new_filename))
            
            if reg.payment_screenshot:
                try:
                    old_path = os.path.join(uploads_dir, reg.payment_screenshot)
                    if os.path.exists(old_path):
                        os.remove(old_path)
                except: pass
                
            reg.payment_screenshot = new_filename

        if not lesson_no: errors.append('Lesson Number is required.')
        if not full_name: errors.append('Full Name is required.')
        if not gender: errors.append('Gender is required.')
        if not age or not age.isdigit(): errors.append('Valid Age is required.')
        if not place: errors.append('City/Village/Town is required.')
        if not state: errors.append('State is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not arrival_date: errors.append('Date of Arrival is required.')
        if not departure_date: errors.append('Date of Departure is required.')
        if not payment_mode: errors.append('Payment Mode is required.')
        if not amount or not amount.isdigit(): errors.append('Valid Amount is required.')

        if not errors:
            if lesson_no != reg.lesson_no and lesson_no.upper() not in ['0', '00', '000', '0000', '00000', 'NA', 'N/A', '-', 'ADMIN', 'NONE', '']:
                existing_lesson = Registration.query.filter(Registration.lesson_no.ilike(lesson_no)).first()
                if existing_lesson:
                    errors.append(f'Lesson Number {lesson_no} is already registered.')
            
            if transaction_id and transaction_id != reg.transaction_id and payment_mode == 'UPI':
                existing_txn = Registration.query.filter(Registration.transaction_id.ilike(transaction_id)).first()
                if existing_txn:
                    errors.append('This Transaction ID has already been submitted.')

        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template('admin/edit_registration.html', config=app.config, form=request.form, reg=reg)

        reg.lesson_no = lesson_no
        reg.full_name = full_name
        reg.gender = gender
        reg.age = int(age)
        reg.place = place
        reg.district = district
        reg.state = state
        reg.email = email
        reg.country_code = country_code
        reg.whatsapp = whatsapp
        reg.is_kriyaban = is_kriyaban
        reg.accommodation = accommodation
        reg.volunteer = volunteer
        reg.arrival_date = arrival_date
        reg.departure_date = departure_date
        reg.payment_mode = payment_mode
        reg.amount = int(amount)
        if transaction_id:
            reg.transaction_id = transaction_id
            
        db.session.commit()
        update_registrations_excel_async()
        log_action(f"Edited registration {reg.reg_id} for {reg.full_name} (Lesson: {reg.lesson_no}, Mobile: {reg.whatsapp}, Amount: {reg.amount})")
        flash('Participant updated successfully.', 'success')
        return redirect(url_for('admin_registrations'))
        
    return render_template('admin/edit_registration.html', config=app.config, form={}, reg=reg)

# ─── ADMIN REGISTRATIONS ──────────────────────────────────────────────────────
@app.route('/admin/registrations')
@login_required
def admin_registrations():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    reg_status = request.args.get('reg_status', '')
    is_kriyaban = request.args.get('is_kriyaban', '')
    accommodation = request.args.get('accommodation', '')
    notified = request.args.get('notified', '')
    payment_mode = request.args.get('payment_mode', '')

    # Show both Approved and Rejected registrations regardless of notification status
    q = Registration.query.filter(Registration.reg_status.in_(['Approved', 'Rejected']))
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%'),
                             Registration.reg_id.ilike(f'%{search}%')))
    if reg_status:
        q = q.filter_by(reg_status=reg_status)
    if is_kriyaban == 'true':
        q = q.filter_by(is_kriyaban=True)
    elif is_kriyaban == 'false':
        q = q.filter_by(is_kriyaban=False)
    if accommodation == 'true':
        q = q.filter_by(accommodation=True)
    elif accommodation == 'false':
        q = q.filter_by(accommodation=False)
    if notified == 'true':
        q = q.filter_by(notified=True)
    elif notified == 'false':
        q = q.filter_by(notified=False)
    if payment_mode:
        q = q.filter_by(payment_mode=payment_mode)

    registrations = q.order_by(Registration.id.desc()).all()
    
    total_registered = Registration.query.filter(Registration.reg_status != 'Rejected').count()
    approval_pending = Registration.query.filter_by(reg_status='Pending').count()
    approved_devotees = Registration.query.filter_by(reg_status='Approved').count()
    collected_amount = db.session.query(db.func.sum(Registration.amount)).filter(Registration.payment_status == 'Paid').scalar() or 0
    total_reg_fee = Registration.query.filter(Registration.payment_status == 'Paid').count() * 1800
    total_acco_fee = Registration.query.filter(Registration.payment_status == 'Paid', Registration.accommodation == True).count() * 1000
    
    stats = {
        'total_registered': total_registered,
        'approval_pending': approval_pending,
        'approved_devotees': approved_devotees,
        'collected_amount': collected_amount,
        'total_reg_fee': total_reg_fee,
        'total_acco_fee': total_acco_fee
    }
    
    return render_template('admin/registrations.html', registrations=registrations,
                           search=search, reg_status=reg_status,
                           is_kriyaban=is_kriyaban, accommodation=accommodation,
                           notified=notified, payment_mode=payment_mode,
                           stats=stats, config=app.config)

@app.route('/admin/requests')
@login_required
def admin_requests():
    search = request.args.get('search', '')
    q = Registration.query.filter(Registration.payment_status == 'Pending', Registration.reg_status != 'Rejected')
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%'),
                             Registration.reg_id.ilike(f'%{search}%')))
    pending_count = q.count()
    registrations = q.order_by(Registration.id.desc()).all()
    return render_template('admin/requests.html', registrations=registrations,
                           search=search, pending_count=pending_count, config=app.config)

@app.route('/api/registrations/<int:rid>/approve', methods=['POST'])
@login_required
def approve_registration(rid):
    reg = Registration.query.get_or_404(rid)
    reg.payment_status = 'Paid'
    reg.reg_status = 'Approved'
    reg.notified = True
    db.session.commit()
    update_registrations_excel_async()
    log_action(f"Approved registration {reg.reg_id} for {reg.full_name}")
    send_member_whatsapp_async(reg.id)
    send_registration_email_async(reg.id)
    return jsonify({'success': True, 'message': 'Registration approved successfully'})

@app.route('/api/registrations/<int:rid>/decline', methods=['POST'])
@login_required
def decline_registration(rid):
    reg = Registration.query.get_or_404(rid)
    reg.payment_status = 'Pending'
    reg.reg_status = 'Rejected'
    db.session.commit()
    update_registrations_excel_async()
    log_action(f"Declined/Rejected registration {reg.reg_id} for {reg.full_name}")
    return jsonify({'success': True, 'message': 'Registration declined'})

@app.route('/api/registrations/<int:rid>/notified', methods=['POST'])
@login_required
def mark_notified(rid):
    reg = Registration.query.get_or_404(rid)
    reg.notified = True
    db.session.commit()
    return jsonify({'success': True, 'message': 'Member marked as notified'})

@app.route('/api/registrations/<int:rid>/notify-whatsapp', methods=['POST'])
@login_required
def notify_whatsapp_single(rid):
    reg = Registration.query.get_or_404(rid)
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    if not gateway_url:
        return jsonify({'success': False, 'message': 'WhatsApp Gateway is not configured. Please set it up first.'})
        
    try:
        send_member_whatsapp(reg)
        reg.notified = True
        db.session.commit()
        return jsonify({'success': True, 'message': f'WhatsApp message sent to {reg.full_name}'})
    except Exception as e:
        print(f"Failed to send WhatsApp for registration {reg.id}: {e}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/registrations/notify-all', methods=['POST'])
@login_required
def registrations_notify_all():
    pending = Registration.query.filter_by(reg_status='Approved', notified=False).all()
    if not pending:
        return jsonify({'success': False, 'message': 'No pending unnotified registrations found'})
        
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    if not gateway_url:
        return jsonify({'success': False, 'message': 'WhatsApp Gateway is not configured. Please set it up first.'})
        
    success_count = 0
    fail_count = 0
    
    import time
    for reg in pending:
        try:
            send_member_whatsapp(reg)
            reg.notified = True
            db.session.commit()
            success_count += 1
            time.sleep(1.5)
        except Exception as e:
            print(f"Failed to send WhatsApp for registration {reg.id}: {e}")
            db.session.rollback()
            fail_count += 1
            
    return jsonify({
        'success': True,
        'message': f"Sent WhatsApp confirmations to {success_count} devotees. {fail_count} failed."
    })

@app.route('/api/registrations/reset-notifications', methods=['POST'])
@login_required
def registrations_reset_notifications():
    try:
        approved = Registration.query.filter_by(reg_status='Approved').all()
        for reg in approved:
            reg.notified = False
        db.session.commit()
        return jsonify({
            'success': True,
            'message': f"Successfully refreshed notification status for all {len(approved)} approved devotees."
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'message': f"Failed to reset notification status: {str(e)}"
        })

@app.route('/admin/registrations/export')
@login_required
def export_registrations():
    update_registrations_excel()
    path = os.path.join(app.config['EXPORTS_DIR'], 'registrations.xlsx')
    return send_file(path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='registrations.xlsx')

@app.route('/admin/registrations/export-pdf')
@login_required
def export_registrations_pdf():
    from fpdf import FPDF
    regs = Registration.query.order_by(Registration.id).all()
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    
    # Stats
    total = len(regs)
    kriyabans = len([r for r in regs if r.is_kriyaban])
    non_kriyabans = total - kriyabans
    acco_yes = len([r for r in regs if r.accommodation])
    acco_no = total - acco_yes

    def add_section(title, data_list):
        pdf.add_page()
        pdf.set_font('helvetica', 'B', 16)
        pdf.cell(0, 10, 'YSS Anantapur - ' + title, 0, 1, 'C')
        pdf.ln(5)
        
        # Summary Table at Top of First Page
        if title == 'Overall Summary & All Records':
            pdf.set_font('helvetica', 'B', 10)
            pdf.cell(60, 8, 'Category', 1, 0, 'C')
            pdf.cell(30, 8, 'Count', 1, 1, 'C')
            pdf.set_font('helvetica', '', 10)
            stats = [
                ('Total Registrations', total),
                ('Kriyabans', kriyabans),
                ('Non-Kriyabans', non_kriyabans),
                ('Accommodation Yes', acco_yes),
                ('Accommodation No', acco_no)
            ]
            for cat, val in stats:
                pdf.cell(60, 8, cat, 1, 0, 'L')
                pdf.cell(30, 8, str(val), 1, 1, 'C')
            pdf.ln(10)

        pdf.set_font('helvetica', 'B', 8)
        cols = [('S.No', 10), ('Reg ID', 25), ('Name', 45), ('Mobile', 30), ('Place', 35), ('Amount', 20), ('Mode', 20), ('Kriya', 25), ('Acco', 15), ('Status', 25)]
        for txt, w in cols:
            pdf.cell(w, 8, txt, 1, 0, 'C')
        pdf.ln()
        
        pdf.set_font('helvetica', '', 8)
        for i, r in enumerate(data_list, 1):
            pdf.cell(10, 8, str(i), 1, 0, 'C')
            pdf.cell(25, 8, r.reg_id or '-', 1, 0, 'C')
            pdf.cell(45, 8, (r.full_name[:25] if r.full_name else '-'), 1, 0, 'L')
            pdf.cell(30, 8, r.whatsapp or '-', 1, 0, 'C')
            pdf.cell(35, 8, (r.place[:20] if r.place else '-'), 1, 0, 'L')
            pdf.cell(20, 8, str(int(r.amount)) if r.amount else '0', 1, 0, 'C')
            pdf.cell(20, 8, r.payment_mode or '-', 1, 0, 'C')
            pdf.cell(25, 8, 'Kriyaban' if r.is_kriyaban else 'Non-Kri', 1, 0, 'C')
            pdf.cell(15, 8, 'Yes' if r.accommodation else 'No', 1, 0, 'C')
            pdf.cell(25, 8, r.reg_status or '-', 1, 0, 'C')
            pdf.ln()

    add_section('Overall Summary & All Records', regs)
        
    import io
    pdf_out = io.BytesIO(pdf.output())
    
    # If view=1 is passed, open inline so the user can print it. Otherwise download it.
    as_attachment = request.args.get('view') != '1'
    return send_file(pdf_out, mimetype='application/pdf', as_attachment=as_attachment, download_name='registrations_report.pdf')

@app.route('/api/registrations/manual', methods=['POST'])
@login_required
def manual_registration():
    data = request.json
    amount_val = float(data.get('amount', 0))
    accommodation_val = True if amount_val >= 2800 else False
    reg = Registration(
        lesson_no=data.get('lesson_no', 'ADMIN'),
        full_name=data.get('full_name'),
        gender='-',
        age=0,
        place=data.get('place', 'Admin Entry'),
        country_code=data.get('country_code', '+91'),
        whatsapp=data.get('whatsapp'),
        arrival_date='2026-07-24',
        departure_date='2026-07-26',
        payment_mode=data.get('payment_mode', 'Cash'),
        amount=amount_val,
        accommodation=accommodation_val,
        payment_status='Paid',
        reg_status='Approved',
        registered_by_id=current_user.id,
        registered_by_name=current_user.name
    )
    db.session.add(reg)
    db.session.commit()
    update_registrations_excel()
    log_action(f"Manual quick-entry registration for {reg.full_name} (Mobile: {reg.whatsapp}, Mode: {reg.payment_mode}, Amount: {reg.amount})")
    return jsonify({'success': True})


@app.route('/api/registrations/<int:rid>', methods=['PUT', 'DELETE'])
@login_required
def api_registration(rid):
    reg = Registration.query.get_or_404(rid)
    if request.method == 'DELETE':
        if reg.payment_screenshot:
            try:
                filepath = os.path.join(app.root_path, 'static', 'uploads', reg.payment_screenshot)
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception as e:
                app.logger.error(f"Failed to delete screenshot {reg.payment_screenshot}: {e}")
                
        # Delete associated room allotment to prevent foreign key constraint errors
        allotment = RoomAllotment.query.filter_by(registration_id=reg.id).first()
        if allotment:
            db.session.delete(allotment)
            
        db.session.delete(reg)
        db.session.commit()
        update_registrations_excel()
        return jsonify({'success': True})
    data = request.json
    reg.payment_status = data.get('payment_status', reg.payment_status)
    reg.reg_status = data.get('reg_status', reg.reg_status)
    db.session.commit()
    update_registrations_excel()
    return jsonify({'success': True})

# ─── ADMIN DONATIONS ──────────────────────────────────────────────────────────
@app.route('/admin/donations')
@login_required
def admin_donations():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    # Processed donations: Received and Notified, or Failed
    q = Donation.query.filter(
        db.or_(
            db.and_(Donation.payment_status == 'Received', Donation.notified == True),
            Donation.payment_status == 'Failed'
        )
    )
    if search:
        q = q.filter(db.or_(Donation.name.ilike(f'%{search}%'),
                             Donation.whatsapp.ilike(f'%{search}%'),
                             Donation.donation_id.ilike(f'%{search}%')))
    donations = q.order_by(Donation.id.asc()).all()
    
    total_donations = Donation.query.filter_by(payment_status='Received').count()
    donated_amount = db.session.query(db.func.sum(Donation.amount)).filter(Donation.payment_status == 'Received').scalar() or 0
    
    stats = {
        'total_donations': total_donations,
        'donated_amount': donated_amount
    }
    
    return render_template('admin/donations.html', donations=donations,
                           search=search, stats=stats, config=app.config)

@app.route('/admin/donation-requests')
@login_required
def admin_donation_requests():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    # Pending requests: Pending status OR Received but not yet notified
    q = Donation.query.filter(
        db.or_(
            Donation.payment_status == 'Pending',
            db.and_(Donation.payment_status == 'Received', Donation.notified == False)
        )
    )
    if search:
        q = q.filter(db.or_(Donation.name.ilike(f'%{search}%'),
                             Donation.whatsapp.ilike(f'%{search}%'),
                             Donation.donation_id.ilike(f'%{search}%')))
    pending_count = q.count()
    donations = q.order_by(Donation.id.asc()).all()
    return render_template('admin/donation_requests.html', donations=donations,
                           search=search, pending_count=pending_count, config=app.config)

@app.route('/admin/donations/export')
@login_required
def export_donations():
    update_donations_excel()
    path = os.path.join(app.config['EXPORTS_DIR'], 'donations.xlsx')
    return send_file(path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='donations.xlsx')

@app.route('/api/donations/<int:did>/notified', methods=['POST'])
@login_required
def mark_don_notified(did):
    don = Donation.query.get_or_404(did)
    don.notified = True
    db.session.commit()
    return jsonify({'success': True, 'message': 'Donor marked as notified'})

@app.route('/api/donations/manual', methods=['POST'])
@login_required
def manual_donation():
    data = request.json
    don = Donation(
        lesson_no=data.get('lesson_no', 'ADMIN'),
        name=data.get('name'),
        age=0,
        place=data.get('place', 'Admin Entry'),
        whatsapp=data.get('whatsapp'),
        amount=float(data.get('amount', 0)),
        payment_mode=data.get('payment_mode', 'Cash'),
        payment_status='Received',
        notified=True
    )
    db.session.add(don)
    db.session.commit()
    update_donations_excel()
    return jsonify({'success': True})

@app.route('/api/donations/<int:did>', methods=['PUT', 'DELETE'])
@login_required
def api_donation(did):
    don = Donation.query.get_or_404(did)
    if request.method == 'DELETE':
        if don.payment_screenshot:
            try:
                filepath = os.path.join(app.root_path, 'static', 'uploads', don.payment_screenshot)
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception as e:
                app.logger.error(f"Failed to delete screenshot {don.payment_screenshot}: {e}")
                
        db.session.delete(don)
        db.session.commit()
        update_donations_excel()
        return jsonify({'success': True})
    data = request.json
    don.payment_status = data.get('payment_status', don.payment_status)
    db.session.commit()
    update_donations_excel()
    return jsonify({'success': True})

# ─── ADMIN ID CARDS ───────────────────────────────────────────────────────────
@app.route('/admin/id-cards')
@login_required
def admin_id_cards():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    is_kriyaban = request.args.get('is_kriyaban', '')
    
    # Only show approved participants in ID cards section
    q = Registration.query.filter_by(reg_status='Approved')
    
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%')))
                             
    if is_kriyaban == 'true':
        q = q.filter_by(is_kriyaban=True)
    elif is_kriyaban == 'false':
        q = q.filter_by(is_kriyaban=False)
                             
    registrations = q.order_by(Registration.id).all()
    return render_template('admin/id_cards.html', registrations=registrations,
                           search=search, is_kriyaban=is_kriyaban, config=app.config)

# ─── ADMIN ROOM ALLOTMENT ────────────────────────────────────────────────────
@app.route('/admin/room-allotment')
@login_required
def admin_room_allotment():
    rooms = Room.query.order_by(Room.id).all()
    
    # Get all people who want accommodation and are approved/pending
    accommodating_users = Registration.query.filter_by(accommodation=True).all()
    
    # Find users with allotments
    allotted_users_map = {a.registration_id: a.room_id for a in RoomAllotment.query.all()} if RoomAllotment.query.first() else {}
    
    # Actually, simpler query for allotments:
    allotments = RoomAllotment.query.all()
    allotted_reg_ids = [a.registration_id for a in allotments]
    
    unallocated = Registration.query.filter(
        Registration.accommodation == True,
        Registration.reg_status == 'Approved',
        ~Registration.id.in_(allotted_reg_ids) if allotted_reg_ids else True
    ).all()
    
    def sort_by_arrival(reg):
        d = reg.arrival_date
        if not d: return '9999-99-99'
        parts = d.split('-')
        if len(parts) == 3 and len(parts[0]) == 2:
            return f"{parts[2]}-{parts[1]}-{parts[0]}"
        return d
        
    unallocated.sort(key=sort_by_arrival)
    
    allocated = Registration.query.filter(
        Registration.reg_status == 'Approved',
        Registration.id.in_(allotted_reg_ids) if allotted_reg_ids else False
    ).all()
    
    room_occupancy = {r.id: [] for r in rooms}
    for a in allotments:
        reg = Registration.query.get(a.registration_id)
        if reg and a.room_id in room_occupancy:
            room_occupancy[a.room_id].append(reg)

    return render_template('admin/room_allotment.html', rooms=rooms, 
                           unallocated=unallocated, allocated=allocated, 
                           room_occupancy=room_occupancy, config=app.config)

@app.route('/admin/room-allotment/add', methods=['POST'])
@login_required
def admin_room_add():
    room_number = request.form.get('room_number')
    capacity = request.form.get('capacity')
    if room_number and capacity and capacity.isdigit():
        if Room.query.filter_by(room_number=room_number).first():
            flash('Room name already exists.', 'error')
        else:
            db.session.add(Room(room_number=room_number, capacity=int(capacity)))
            db.session.commit()
            flash('Room added successfully.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/rename', methods=['POST'])
@login_required
def admin_room_rename():
    room_id = request.form.get('room_id')
    new_name = request.form.get('new_name')
    new_capacity = request.form.get('new_capacity')
    room = Room.query.get(room_id)
    if room:
        if new_name and new_name != room.room_number:
            existing = Room.query.filter_by(room_number=new_name).first()
            if existing:
                flash('Another room with this name already exists.', 'error')
                return redirect(url_for('admin_room_allotment'))
            room.room_number = new_name
        if new_capacity and new_capacity.isdigit():
            room.capacity = int(new_capacity)
        db.session.commit()
        flash('Room updated successfully.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/allot', methods=['POST'])
@login_required
def admin_room_allot():
    registration_id = request.form.get('registration_id')
    room_id = request.form.get('room_id')
    
    room = Room.query.get(room_id)
    reg = Registration.query.get(registration_id)
    
    if not room or not reg:
        flash('Invalid room or registration.', 'error')
        return redirect(url_for('admin_room_allotment'))
        
    current_occupants = RoomAllotment.query.filter_by(room_id=room_id).count()
    if current_occupants >= room.capacity:
        flash(f'Room {room.room_number} is already full.', 'error')
        return redirect(url_for('admin_room_allotment'))
        
    # Check if already allotted
    existing = RoomAllotment.query.filter_by(registration_id=reg.id).first()
    if existing:
        old_room = existing.room.room_number if existing.room else "None"
        existing.room_id = room.id
        log_action(f"Re-allotted room from {old_room} to {room.room_number} for devotee {reg.full_name} ({reg.reg_id})")
    else:
        new_allotment = RoomAllotment(registration_id=reg.id, room_id=room.id)
        db.session.add(new_allotment)
        log_action(f"Allotted room {room.room_number} to devotee {reg.full_name} ({reg.reg_id})")
        
    db.session.commit()
    send_room_allotment_whatsapp_async(reg.id)
    flash(f'{reg.full_name} allotted to {room.room_number}.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/unallot', methods=['POST'])
@login_required
def admin_room_unallot():
    registration_id = request.form.get('registration_id')
    allotment = RoomAllotment.query.filter_by(registration_id=registration_id).first()
    if allotment:
        reg = allotment.registration
        room_name = allotment.room.room_number if allotment.room else "None"
        db.session.delete(allotment)
        db.session.commit()
        log_action(f"Removed room allotment ({room_name}) for devotee {reg.full_name} ({reg.reg_id})")
        flash('Allotment removed.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/notify', methods=['POST'])
@login_required
def admin_room_notify():
    allotments = RoomAllotment.query.all()
    pending = []
    for a in allotments:
        # If never notified via WhatsApp OR room changed since last WhatsApp notification
        if not a.notified_room_whatsapp or a.notified_room_whatsapp != a.room.room_number:
            pending.append(a)
            
    if not pending:
        flash('All attendees are already notified of their current room assignments via WhatsApp.', 'info')
        return redirect(url_for('admin_room_allotment'))
        
    success_count = 0
    fail_count = 0
    
    # Try sending via self-hosted gateway
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    
    if not gateway_url:
        flash('WhatsApp Gateway URL is not configured. Please set it up first.', 'error')
        return redirect(url_for('admin_room_allotment'))
        
    for a in pending:
        reg = a.registration
        room = a.room
        try:
            body_text = format_whatsapp_template(
                'room_allot',
                name=reg.full_name,
                reg_id=reg.reg_id,
                room_number=room.room_number,
                arrival_date=reg.arrival_date if reg.arrival_date else '24-07-2026',
                departure_date=reg.departure_date if reg.departure_date else '26-07-2026'
            )
            if not body_text:
                body_text = (
                    f"Dear {reg.full_name},\n\n"
                    f"With divine blessings, we are happy to inform you that your accommodation has been successfully allotted for the 3-Day Spiritual Program at Anantapur.\n\n"
                    f"Accommodation Details:\n\n"
                    f"Name: {reg.full_name}\n"
                    f"Room Number: {room.room_number}\n"
                    f"Check-In Date: {reg.arrival_date if reg.arrival_date else '24-07-2026'}\n"
                    f"Check-Out Date: {reg.departure_date if reg.departure_date else '26-07-2026'}\n\n"
                    f"Jai Guru"
                )
            r = requests.post(
                f"{gateway_url}/send",
                json={
                    'to': reg.whatsapp,
                    'message': body_text
                },
                timeout=5
            )
            if r.status_code == 200:
                a.notified_room_whatsapp = room.room_number
                a.notified_room_number = room.room_number # keep notified_room_number updated too
                success_count += 1
            else:
                print(f"Failed to send WhatsApp: {r.status_code} - {r.text}")
                fail_count += 1
        except Exception as e:
            print(f"Failed to notify {reg.whatsapp} via WhatsApp: {e}")
            fail_count += 1
            
    db.session.commit()
    
    if fail_count > 0:
        flash(f"Successfully notified {success_count} attendees via WhatsApp. {fail_count} failed to deliver.", 'warning')
    else:
        flash(f"Successfully notified {success_count} attendees via WhatsApp!", 'success')
        
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/reset-notifications', methods=['POST'])
@login_required
def admin_room_reset_notifications():
    allotments = RoomAllotment.query.all()
    for a in allotments:
        a.notified_room_number = None
        a.notified_room_whatsapp = None
    db.session.commit()
    flash('WhatsApp notification flags reset successfully.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/mark-notified/<int:reg_id>', methods=['POST'])
@login_required
def admin_room_mark_notified(reg_id):
    allotment = RoomAllotment.query.filter_by(registration_id=reg_id).first()
    if allotment:
        allotment.notified_room_whatsapp = allotment.room.room_number
        db.session.commit()
        return jsonify({'success': True})
    return jsonify({'success': False, 'message': 'Allotment not found'}), 404

# ─── ADMIN SCHEDULE MGMT ─────────────────────────────────────────────────────
@app.route('/admin/schedule', methods=['GET', 'POST'])
@login_required
def admin_schedule():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            item = EventSchedule(
                day_number=int(request.form.get('day_number')),
                day_label=request.form.get('day_label'),
                start_time=request.form.get('start_time'),
                end_time=request.form.get('end_time'),
                activity=request.form.get('activity'),
                category=request.form.get('category'),
                sort_order=int(request.form.get('sort_order', 99))
            )
            db.session.add(item)
            db.session.commit()
            flash('Schedule item added.', 'success')
        elif action == 'delete':
            item = EventSchedule.query.get(int(request.form.get('item_id')))
            if item:
                db.session.delete(item)
                db.session.commit()
            flash('Schedule item deleted.', 'success')
        elif action == 'edit':
            item = EventSchedule.query.get(int(request.form.get('item_id')))
            if item:
                item.day_number = int(request.form.get('day_number'))
                item.day_label = request.form.get('day_label')
                item.start_time = request.form.get('start_time')
                item.end_time = request.form.get('end_time')
                item.activity = request.form.get('activity')
                item.category = request.form.get('category')
                item.sort_order = int(request.form.get('sort_order', 99))
                db.session.commit()
            flash('Schedule item updated.', 'success')
        return redirect(url_for('admin_schedule'))

    schedules = EventSchedule.query.order_by(EventSchedule.day_number, EventSchedule.sort_order).all()
    return render_template('admin/schedule_mgmt.html', schedules=schedules, config=app.config)

# ─── ERROR HANDLERS ───────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('errors/404.html', config=app.config), 404

@app.errorhandler(500)
def server_error(e):
    return render_template('errors/500.html', config=app.config), 500


# ─── ADMIN CREDENTIALS ────────────────────────────────────────────────────────
@app.route('/admin/credentials', methods=['GET', 'POST'])
@login_required
def admin_credentials():
    if not current_user.is_main_admin:
        flash('Unauthorized access.', 'error')
        return redirect(url_for('admin_registrations'))
        
    if request.method == 'POST':
        new_email = request.form.get('new_email')
        new_password = request.form.get('new_password')
        
        updated = False
        if new_email and new_email != current_user.email:
            # Check if email exists
            if Admin.query.filter_by(email=new_email).first():
                flash('Email already in use.', 'error')
            else:
                current_user.email = new_email
                updated = True
        
        if new_password:
            current_user.set_password(new_password)
            updated = True
            
        if updated:
            db.session.commit()
            flash('Admin credentials updated successfully.', 'success')
            
        return redirect(url_for('admin_credentials'))
        
    return render_template('admin/manage.html', config=app.config)

# ─── ADMIN GALLERY MANAGEMENT ─────────────────────────────────────────────────
@app.route('/admin/gallery')
@login_required
def admin_gallery():
    images = GalleryImage.query.order_by(GalleryImage.created_at.desc()).all()
    return render_template('admin/gallery_mgmt.html', images=images, config=app.config)

@app.route('/admin/gallery/upload', methods=['POST'])
@login_required
def admin_gallery_upload():
    if 'image' not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for('admin_gallery'))
        
    file = request.files['image']
    caption = request.form.get('caption', '').strip()
    
    if file.filename == '':
        flash("No file selected.", "error")
        return redirect(url_for('admin_gallery'))
        
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.jpg', '.jpeg', '.png', '.webp', '.gif']:
        flash("Unsupported image format. Please upload JPG, PNG, WEBP or GIF.", "error")
        return redirect(url_for('admin_gallery'))
        
    try:
        filename = f"{uuid.uuid4().hex}{ext}"
        filepath = os.path.join(app.root_path, 'static', 'uploads', 'gallery', filename)
        file.save(filepath)
        
        new_photo = GalleryImage(filename=filename, caption=caption)
        db.session.add(new_photo)
        db.session.commit()
        
        log_action(f"Uploaded gallery photo '{filename}' (caption: '{caption}')")
        flash("Photo uploaded and published successfully!", "success")
    except Exception as e:
        print(f"Gallery upload error: {e}")
        flash(f"Failed to upload photo: {e}", "error")
        
    return redirect(url_for('admin_gallery'))

@app.route('/admin/gallery/delete/<int:photo_id>', methods=['POST'])
@login_required
def admin_gallery_delete(photo_id):
    photo = GalleryImage.query.get_or_404(photo_id)
    try:
        filepath = os.path.join(app.root_path, 'static', 'uploads', 'gallery', photo.filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            
        filename = photo.filename
        caption = photo.caption
        db.session.delete(photo)
        db.session.commit()
        
        log_action(f"Deleted gallery photo '{filename}' (caption: '{caption}')")
        flash("Photo deleted successfully.", "success")
    except Exception as e:
        print(f"Gallery delete error: {e}")
        flash(f"Failed to delete photo: {e}", "error")
        
    return redirect(url_for('admin_gallery'))

# ─── ADMIN ACTIVITY LOG & ACTIVE TRACKER ──────────────────────────────────────
@app.route('/admin/activity-log')
@app.route('/admin/activity_log')
@login_required
def admin_activity_log():
    page = request.args.get('page', 1, type=int)
    per_page = 25
    
    five_mins_ago = datetime.utcnow() - timedelta(minutes=5)
    online_admins = Admin.query.filter(Admin.last_active >= five_mins_ago).all()
    all_admins = Admin.query.order_by(Admin.name.asc()).all()
    
    logs_paginated = ActivityLog.query.order_by(ActivityLog.timestamp.desc()).paginate(page=page, per_page=per_page, error_out=False)
    now = datetime.utcnow()
    
    # Per-admin stats for the summary panel
    from sqlalchemy import func
    reg_counts = dict(
        db.session.query(Registration.registered_by_id, func.count(Registration.id))
        .group_by(Registration.registered_by_id).all()
    )
    activity_counts = dict(
        db.session.query(ActivityLog.admin_id, func.count(ActivityLog.id))
        .group_by(ActivityLog.admin_id).all()
    )
    admin_stats = {}
    for a in all_admins:
        admin_stats[a.id] = {
            'reg_count': reg_counts.get(a.id, 0),
            'activity_count': activity_counts.get(a.id, 0),
            'is_online': bool(a.last_active and (now - a.last_active).total_seconds() < 300)
        }
    
    return render_template(
        'admin/activity_log.html',
        logs=logs_paginated.items,
        pagination=logs_paginated,
        online_admins=online_admins,
        all_admins=all_admins,
        admin_stats=admin_stats,
        now=now,
        config=app.config
    )

@app.route('/admin/activity-log/clear', methods=['POST'])
@login_required
def admin_clear_activity_log():
    if not current_user.is_main_admin:
        flash("Only the Main Admin can clear activity logs.", "error")
        return redirect(url_for('admin_activity_log'))
        
    try:
        ActivityLog.query.delete()
        db.session.commit()
        log_action("Cleared all activity logs")
        flash("Activity logs cleared successfully.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Failed to clear activity logs: {e}", "error")
        
    return redirect(url_for('admin_activity_log'))

# ─── ADMIN WHATSAPP SETUP ───────────────────────────────────────────────────
@app.route('/admin/whatsapp-setup')
@app.route('/admin/whatsapp_setup')
@login_required
def admin_whatsapp_setup():
    templates = WhatsAppTemplate.query.order_by(WhatsAppTemplate.key.asc()).all()
    count_7d_pending = Registration.query.filter_by(reg_status='Approved', reminder_7d_sent=False).count()
    count_3d_pending = Registration.query.filter_by(reg_status='Approved', reminder_3d_sent=False).count()
    count_1d_pending = Registration.query.filter_by(reg_status='Approved', reminder_1d_sent=False).count()
    return render_template(
        'admin/whatsapp_setup.html',
        config=app.config,
        templates=templates,
        count_7d_pending=count_7d_pending,
        count_3d_pending=count_3d_pending,
        count_1d_pending=count_1d_pending
    )

@app.route('/admin/whatsapp-templates/update', methods=['POST'])
@login_required
def admin_whatsapp_template_update():
    key = request.form.get('key')
    template_text = request.form.get('template_text')
    t = WhatsAppTemplate.query.filter_by(key=key).first()
    if t:
        t.template_text = template_text
        db.session.commit()
        log_action(f"Updated WhatsApp template '{key}'")
        flash(f"WhatsApp template '{key}' updated successfully.", "success")
    else:
        flash("Template not found.", "error")
    return redirect(url_for('admin_whatsapp_setup'))

@app.route('/admin/whatsapp-templates/add', methods=['POST'])
@login_required
def admin_whatsapp_template_add():
    key = request.form.get('key', '').strip().lower().replace(' ', '_')
    description = request.form.get('description', '').strip()
    template_text = request.form.get('template_text', '').strip()
    variables = request.form.get('variables', '').strip()
    
    if not key:
        flash("Template key is required.", "error")
        return redirect(url_for('admin_whatsapp_setup'))
        
    existing = WhatsAppTemplate.query.filter_by(key=key).first()
    if existing:
        flash(f"Template with key '{key}' already exists.", "error")
        return redirect(url_for('admin_whatsapp_setup'))
        
    try:
        new_t = WhatsAppTemplate(
            key=key,
            description=description,
            variables=variables,
            template_text=template_text
        )
        db.session.add(new_t)
        db.session.commit()
        log_action(f"Added WhatsApp template '{key}'")
        flash(f"WhatsApp template '{key}' added successfully.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Failed to add template: {e}", "error")
        
    return redirect(url_for('admin_whatsapp_setup'))

@app.route('/admin/whatsapp-send-all-due-reminders', methods=['POST'])
@login_required
def admin_whatsapp_send_all_due_reminders():
    from datetime import date
    today = date.today()
    event_date = date(2026, 7, 24)
    days_left = (event_date - today).days
    
    due_reminders = []
    if days_left <= 7:
        due_reminders.append((7, 'reminder_7d'))
    if days_left <= 3:
        due_reminders.append((3, 'reminder_3d'))
    if days_left <= 1:
        due_reminders.append((1, 'reminder_1d'))
        
    if not due_reminders:
        flash(f"No reminders are currently due/past due. Days remaining: {days_left}.", "info")
        return redirect(url_for('admin_whatsapp_setup'))
        
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    if not gateway_url:
        flash("WhatsApp Gateway URL is not configured. Please set it up first.", "error")
        return redirect(url_for('admin_whatsapp_setup'))
        
    import requests
    success_count = 0
    fail_count = 0
    
    for days, key in due_reminders:
        if days == 7:
            pending = Registration.query.filter_by(reg_status='Approved', reminder_7d_sent=False).all()
        elif days == 3:
            pending = Registration.query.filter_by(reg_status='Approved', reminder_3d_sent=False).all()
        else:
            pending = Registration.query.filter_by(reg_status='Approved', reminder_1d_sent=False).all()
            
        for reg in pending:
            try:
                body = format_whatsapp_template(key, name=reg.full_name, reg_id=reg.reg_id)
                if not body:
                    fail_count += 1
                    continue
                    
                r = requests.post(
                    f"{gateway_url}/send",
                    json={
                        'to': reg.whatsapp,
                        'message': body
                    },
                    timeout=5
                )
                if r.status_code == 200:
                    if days == 7:
                        reg.reminder_7d_sent = True
                    elif days == 3:
                        reg.reminder_3d_sent = True
                    else:
                        reg.reminder_1d_sent = True
                    success_count += 1
                else:
                    fail_count += 1
            except Exception as e:
                print(f"Error sending {days}-day reminder to {reg.whatsapp}: {e}")
                fail_count += 1
                
    db.session.commit()
    log_action(f"Manually sent all due reminders (Days remaining: {days_left}): {success_count} messages sent, {fail_count} failed.")
    
    if fail_count > 0:
        flash(f"Dispatched {success_count} due reminders. {fail_count} messages failed to send.", "warning")
    else:
        flash(f"Successfully sent {success_count} due reminders!", "success")
        
    return redirect(url_for('admin_whatsapp_setup'))

@app.route('/admin/whatsapp-send-reminders/<int:days>', methods=['POST'])
@login_required
def admin_whatsapp_send_reminders(days):
    if days not in [7, 3, 1]:
        flash("Invalid reminder interval.", "error")
        return redirect(url_for('admin_whatsapp_setup'))
        
    if days == 7:
        pending = Registration.query.filter_by(reg_status='Approved', reminder_7d_sent=False).all()
        key = 'reminder_7d'
    elif days == 3:
        pending = Registration.query.filter_by(reg_status='Approved', reminder_3d_sent=False).all()
        key = 'reminder_3d'
    else:
        pending = Registration.query.filter_by(reg_status='Approved', reminder_1d_sent=False).all()
        key = 'reminder_1d'
        
    if not pending:
        flash(f"No pending approved devotees require the {days}-day reminder.", "info")
        return redirect(url_for('admin_whatsapp_setup'))
        
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
    if not gateway_url:
        flash("WhatsApp Gateway URL is not configured. Please set it up first.", "error")
        return redirect(url_for('admin_whatsapp_setup'))
        
    import requests
    success_count = 0
    fail_count = 0
    
    for reg in pending:
        try:
            body = format_whatsapp_template(key, name=reg.full_name, reg_id=reg.reg_id)
            if not body:
                fail_count += 1
                continue
                
            r = requests.post(
                f"{gateway_url}/send",
                json={
                    'to': reg.whatsapp,
                    'message': body
                },
                timeout=5
            )
            if r.status_code == 200:
                if days == 7:
                    reg.reminder_7d_sent = True
                elif days == 3:
                    reg.reminder_3d_sent = True
                else:
                    reg.reminder_1d_sent = True
                success_count += 1
            else:
                fail_count += 1
        except Exception as e:
            print(f"Error sending {days}-day reminder to {reg.whatsapp}: {e}")
            fail_count += 1
            
    db.session.commit()
    log_action(f"Manually sent {days}-day reminder WhatsApp messages to {success_count} devotees (failed: {fail_count})")
    
    if fail_count > 0:
        flash(f"Sent {success_count} reminders. {fail_count} messages failed to dispatch.", "warning")
    else:
        flash(f"Successfully sent {success_count} reminders!", "success")
        
    return redirect(url_for('admin_whatsapp_setup'))

@app.route('/admin/whatsapp-status')
@login_required
def admin_whatsapp_status():
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL', 'http://localhost:3000')
    status_data = {
        'status': 'Offline',
        'qr': None,
        'phone': None,
        'gateway_url': gateway_url
    }
    try:
        # Increased timeout to 2.0s to prevent transient network timeout false-alarms
        r = requests.get(f"{gateway_url}/status", timeout=2.0)
        if r.status_code == 200:
            res = r.json()
            status_data['status'] = res.get('status', 'Disconnected')
            status_data['qr'] = res.get('qr')
            status_data['phone'] = res.get('phone')
    except Exception as e:
        print(f"Error fetching WhatsApp gateway status: {e}")
    return jsonify(status_data)

@app.route('/admin/whatsapp-messages')
@login_required
def admin_whatsapp_messages():
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL', 'http://localhost:3000')
    try:
        r = requests.get(f"{gateway_url}/recent-messages", timeout=1)
        if r.status_code == 200:
            return jsonify(r.json())
    except Exception as e:
        print(f"Error fetching recent messages: {e}")
    return jsonify({'messages': []})

@app.route('/admin/whatsapp-reset', methods=['POST'])
@login_required
def admin_whatsapp_reset():
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL', 'http://localhost:3000')
    try:
        r = requests.post(f"{gateway_url}/reset", timeout=5.0)
        if r.status_code == 200:
            flash("WhatsApp Gateway connection reset successfully. Generating new QR...", "success")
        else:
            flash(f"Failed to reset WhatsApp gateway: {r.text}", "error")
    except Exception as e:
        flash(f"Error communicating with WhatsApp gateway: {e}", "error")
    return redirect(url_for('admin_whatsapp_setup'))

def run_automatic_reminders_scheduler():
    """
    Background worker thread that runs every hour to check if the current date matches the reminder dates,
    and automatically triggers reminders.
    """
    import time
    from datetime import date
    while True:
        try:
            with app.app_context():
                today = date.today()
                # Target dates in UTC or IST. Event is 2026-07-24
                event_date = date(2026, 7, 24)
                days_left = (event_date - today).days
                
                if days_left in [7, 3, 1]:
                    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
                    if gateway_url:
                        import requests
                        
                        if days_left == 7:
                            pending = Registration.query.filter_by(reg_status='Approved', reminder_7d_sent=False).all()
                            key = 'reminder_7d'
                        elif days_left == 3:
                            pending = Registration.query.filter_by(reg_status='Approved', reminder_3d_sent=False).all()
                            key = 'reminder_3d'
                        else:
                            pending = Registration.query.filter_by(reg_status='Approved', reminder_1d_sent=False).all()
                            key = 'reminder_1d'
                            
                        sent_count = 0
                        for reg in pending:
                            body = format_whatsapp_template(key, name=reg.full_name, reg_id=reg.reg_id)
                            if body:
                                try:
                                    r = requests.post(
                                        f"{gateway_url}/send",
                                        json={'to': reg.whatsapp, 'message': body},
                                        timeout=5
                                    )
                                    if r.status_code == 200:
                                        if days_left == 7:
                                            reg.reminder_7d_sent = True
                                        elif days_left == 3:
                                            reg.reminder_3d_sent = True
                                        else:
                                            reg.reminder_1d_sent = True
                                        sent_count += 1
                                except Exception as e:
                                    print(f"Auto-scheduler sending error: {e}")
                        if sent_count > 0:
                            db.session.commit()
                            log_action(f"Auto-scheduled daemon successfully sent {days_left}-day reminders to {sent_count} devotees")
        except Exception as e:
            print(f"Auto-reminder background scheduler error: {e}")
        
        # Sleep for 1 hour
        time.sleep(3600)

# Start background scheduler thread
import threading
scheduler_thread = threading.Thread(target=run_automatic_reminders_scheduler, daemon=True)
scheduler_thread.start()

# ─── MAIN ─────────────────────────────────────────────────────────────────────
with app.app_context():
    seed_data()

if __name__ == '__main__':
    app.run(host='127.0.0.1', debug=False, port=5002)

