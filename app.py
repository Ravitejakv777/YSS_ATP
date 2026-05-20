from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file, make_response
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_mail import Mail, Message
from models import db, Admin, Registration, Donation, EventSchedule, Room, RoomAllotment
from config import Config
from datetime import datetime, timedelta
import os, openpyxl, uuid

app = Flask(__name__)
app.config.from_object(Config)
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db.init_app(app)
mail = Mail(app)

login_manager = LoginManager(app)
login_manager.login_view = 'admin_login'
login_manager.login_message = 'Please login to access admin panel.'

# Initialize Database on Startup (Required for Render)
with app.app_context():
    try:
        # Ensure required directories exist
        os.makedirs(app.config.get('EXPORTS_DIR', os.path.join(app.root_path, 'exports')), exist_ok=True)
        os.makedirs(os.path.join(app.root_path, 'instance'), exist_ok=True)
        
        print("Checking database connection...")
        db.create_all()
        # Safe migration for PostgreSQL / SQLite
        # Column 1: notified_room_number
        try:
            db.session.execute(db.text("ALTER TABLE room_allotments ADD COLUMN IF NOT EXISTS notified_room_number VARCHAR(100)"))
            db.session.commit()
        except Exception:
            db.session.rollback()
            try:
                db.session.execute(db.text("ALTER TABLE room_allotments ADD COLUMN notified_room_number VARCHAR(100)"))
                db.session.commit()
            except Exception:
                db.session.rollback()
                
        # Column 2: notified_room_whatsapp
        try:
            db.session.execute(db.text("ALTER TABLE room_allotments ADD COLUMN IF NOT EXISTS notified_room_whatsapp VARCHAR(100)"))
            db.session.commit()
            print("RoomAllotment table migrations verified successfully.")
        except Exception as migration_error:
            db.session.rollback()
            try:
                db.session.execute(db.text("ALTER TABLE room_allotments ADD COLUMN notified_room_whatsapp VARCHAR(100)"))
                db.session.commit()
                print("RoomAllotment whatsapp column added successfully.")
            except Exception as sqlite_err:
                db.session.rollback()
                print(f"Skipping migration (column might already exist): {sqlite_err}")
        # Column 3: district
        try:
            db.session.execute(db.text("ALTER TABLE registrations ADD COLUMN IF NOT EXISTS district VARCHAR(100)"))
            db.session.commit()
            print("Registration table migrations verified successfully.")
        except Exception as migration_error:
            db.session.rollback()
            try:
                db.session.execute(db.text("ALTER TABLE registrations ADD COLUMN district VARCHAR(100)"))
                db.session.commit()
                print("Registration district column added successfully.")
            except Exception as sqlite_err:
                db.session.rollback()
                print(f"Skipping migration (column might already exist): {sqlite_err}")
                
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

# ─── EXCEL HELPERS ────────────────────────────────────────────────────────────
def update_registrations_excel():
    path = os.path.join(app.config['EXPORTS_DIR'], 'registrations.xlsx')
    wb = openpyxl.Workbook()
    regs = Registration.query.order_by(Registration.id).all()
    
    # Stats
    total = len(regs)
    locals_count = len([r for r in regs if r.place and 'anantapur' in r.place.lower()])
    non_locals = total - locals_count
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
            ws.append(['Locals (Anantapur)', locals_count])
            ws.append(['Non-Locals', non_locals])
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
    create_sheet('Locals', [r for r in regs if r.place and 'anantapur' in r.place.lower()])
    create_sheet('Non-Locals', [r for r in regs if not r.place or 'anantapur' not in r.place.lower()])
    create_sheet('Kriyabans', [r for r in regs if r.is_kriyaban])
    create_sheet('Non-Kriyabans', [r for r in regs if not r.is_kriyaban])
    create_sheet('Accommodation', [r for r in regs if r.accommodation])

    wb.save(path)
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

    wb.save(path)
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

def send_member_whatsapp(reg):
    """
    Sends a WhatsApp message to the member after Admin approval.
    """
    message = (
        f"Dear {reg.full_name},\n\n"
        f"With divine blessings and heartfelt joy, we are happy to confirm your successful registration for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
        f"Your Registration Details:\n\n"
        f"Name: {reg.full_name}\n"
        f"Phone Number: {reg.whatsapp}\n"
        f"Email: {reg.email}\n"
        f"City: {reg.place}\n"
        f"Accommodation: {'Yes' if reg.accommodation else 'No'}\n\n"
        f"Program Details:\n\n"
        f"Event: 3-Day Spiritual Program Anantapur\n"
        f"Venue: Revenue Kalyana Mandapam (Revenue Bhavan), Beside Krishna Kalamandir, Near Clock Tower, Anantapur, Andhra Pradesh\n"
        f"Dates: June 24-26, 2026\n"
        f"Venue Location: https://www.google.com/maps/place/MHJW%2BQGV+Krishna+Kala+Mandir,+near+Clock+Tower,+Kamalanagar,+Anantapur,+Andhra+Pradesh+515001/\n\n"
        f"May this sacred gathering fill your heart with peace, devotion, positivity, and spiritual upliftment. We sincerely thank you for choosing to be part of this divine journey.\n\n"
        f"Please carry your registration confirmation during your visit. Further updates and instructions will be shared soon.\n\n"
        f"We look forward to welcoming you with love and prayers.\n\n"
        f"Jai Guru"
    )
    print(f"WHATSAPP LOG: {message}")
    
    # Try sending via self-hosted gateway
    import requests
    try:
        gateway_url = app.config.get('WHATSAPP_GATEWAY_URL')
        if gateway_url:
            r = requests.post(
                f"{gateway_url}/send",
                json={
                    'to': reg.whatsapp,
                    'message': message
                },
                timeout=5
            )
            print(f"AUTOMATED WHATSAPP STATUS: {r.status_code} - {r.text}")
    except Exception as e:
        print(f"Failed to send automated WhatsApp: {e}")


# ─── PUBLIC ROUTES ────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html', config=app.config)

@app.route('/about')
def about():
    return render_template('about.html', config=app.config)

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

# ─── REGISTRATION ─────────────────────────────────────────────────────────────
@app.route('/registration', methods=['GET', 'POST'])
def registration():
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
        whatsapp = request.form.get('whatsapp')
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
        if not email or '@' not in email: errors.append('Valid Email Address is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not arrival_date: errors.append('Date of Arrival is required.')
        if not departure_date: errors.append('Date of Departure is required.')
        if not payment_mode: errors.append('Payment Mode is required.')
        
        if not transaction_id: errors.append('Transaction ID is required.')
        if not screenshot_filename: errors.append('Payment Screenshot is required.')

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
            return render_template('registration.html', config=app.config, form=request.form)

        # Calculate amount
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
            transaction_id=transaction_id, payment_screenshot=screenshot_filename
        )
        db.session.add(reg)
        db.session.commit()
        update_registrations_excel()
        send_submission_email(reg)
        send_admin_email_alert(reg) # Send Email to Admin
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
        if not transaction_id: errors.append('Transaction ID is required.')

        screenshot_filename = None
        if 'payment_screenshot' in request.files:
            file = request.files['payment_screenshot']
            if file and file.filename:
                screenshot_filename = secure_filename(f"donation_{int(time.time())}_{file.filename}")
                os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], screenshot_filename))
        
        if not screenshot_filename and payment_mode == 'UPI':
            errors.append('Payment Screenshot is required for UPI.')

        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template('donation.html', config=app.config, form=request.form)

        don = Donation(
            lesson_no=lesson_no, name=name, age=int(age), place=place,
            whatsapp=whatsapp, amount=float(amount), payment_mode=payment_mode,
            transaction_id=transaction_id, payment_screenshot=screenshot_filename
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
            mail.send(msg)
            flash('A password reset link has been sent to your email.', 'success')
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
            
            if Admin.query.filter_by(email=email).first():
                flash('Admin with this email already exists.', 'error')
            else:
                new_admin = Admin(lesson_no=lesson_no, name=name, email=email, mobile=mobile)
                new_admin.set_password(password)
                db.session.add(new_admin)
                db.session.commit()
                flash('New admin added successfully.', 'success')
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
    return render_template('admin/manage_admins.html', admins=admins, config=app.config)

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
            return redirect(url_for('admin_registrations'))
        flash('Invalid Admin Name or password.', 'error')
    return render_template('admin/login.html', config=app.config)

@app.route('/admin/logout')
@login_required
def admin_logout():
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
            notified=True if email and (payment_mode == 'Cash' or transaction_id) else False
        )
        db.session.add(reg)
        db.session.commit()
        update_registrations_excel()
        if email and reg.payment_status == 'Paid':
            send_registration_email(reg)
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
        update_registrations_excel()
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
    # Show both Paid (Approved and Notified) and Rejected (Declined) registrations
    q = Registration.query.filter(
        db.or_(
            db.and_(Registration.payment_status == 'Paid', Registration.notified == True),
            Registration.reg_status == 'Rejected'
        )
    )
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%'),
                             Registration.reg_id.ilike(f'%{search}%')))
    if reg_status:
        q = q.filter_by(reg_status=reg_status)
    registrations = q.order_by(Registration.id.asc()).all()
    
    total_registered = Registration.query.count()
    approval_pending = Registration.query.filter(
        db.or_(
            Registration.payment_status == 'Pending',
            db.and_(Registration.reg_status == 'Approved', Registration.notified == False)
        )
    ).count()
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
                           search=search, reg_status=reg_status, stats=stats, config=app.config)

# ─── ADMIN REQUESTS (Pending payment transactions) ────────────────────────────
@app.route('/admin/requests')
@login_required
def admin_requests():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    q = Registration.query.filter(Registration.payment_status == 'Pending')
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%'),
                             Registration.reg_id.ilike(f'%{search}%')))
    pending_count = q.count()
    registrations = q.order_by(Registration.id.asc()).all()
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
    update_registrations_excel()
    send_member_whatsapp(reg) # Send WhatsApp to Member
    send_registration_email(reg) # Send Email with printable ID Card
    return jsonify({'success': True, 'message': 'Registration approved successfully'})

@app.route('/api/registrations/<int:rid>/decline', methods=['POST'])
@login_required
def decline_registration(rid):
    reg = Registration.query.get_or_404(rid)
    reg.payment_status = 'Pending'
    reg.reg_status = 'Rejected'
    db.session.commit()
    update_registrations_excel()
    return jsonify({'success': True, 'message': 'Registration declined'})

@app.route('/api/registrations/<int:rid>/notified', methods=['POST'])
@login_required
def mark_notified(rid):
    reg = Registration.query.get_or_404(rid)
    reg.notified = True
    db.session.commit()
    return jsonify({'success': True, 'message': 'Member marked as notified'})

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
    locals_count = len([r for r in regs if r.place and 'anantapur' in r.place.lower()])
    non_locals = total - locals_count
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
                ('Locals (Anantapur)', locals_count),
                ('Non-Locals', non_locals),
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
    add_section('Locals (Anantapur)', [r for r in regs if r.place and 'anantapur' in r.place.lower()])
    add_section('Non-Locals', [r for r in regs if not r.place or 'anantapur' not in r.place.lower()])
    add_section('Kriyabans', [r for r in regs if r.is_kriyaban])
    add_section('Non-Kriyabans', [r for r in regs if not r.is_kriyaban])
    add_section('Accommodation Needed', [r for r in regs if r.accommodation])
        
    import io
    pdf_out = io.BytesIO(pdf.output())
    return send_file(pdf_out, mimetype='application/pdf', as_attachment=True, download_name='registrations_report.pdf')

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
        reg_status='Approved'
    )
    db.session.add(reg)
    db.session.commit()
    update_registrations_excel()
    return jsonify({'success': True})


@app.route('/api/registrations/<int:rid>', methods=['PUT', 'DELETE'])
@login_required
def api_registration(rid):
    reg = Registration.query.get_or_404(rid)
    if request.method == 'DELETE':
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
    
    # Only show approved participants in ID cards section
    q = Registration.query.filter_by(reg_status='Approved')
    
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%')))
                             
    registrations = q.order_by(Registration.id).all()
    return render_template('admin/id_cards.html', registrations=registrations,
                           search=search, config=app.config)

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
        if new_name:
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
        existing.room_id = room.id
    else:
        new_allotment = RoomAllotment(registration_id=reg.id, room_id=room.id)
        db.session.add(new_allotment)
        
    db.session.commit()
    flash(f'{reg.full_name} allotted to {room.room_number}.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/unallot', methods=['POST'])
@login_required
def admin_room_unallot():
    registration_id = request.form.get('registration_id')
    allotment = RoomAllotment.query.filter_by(registration_id=registration_id).first()
    if allotment:
        db.session.delete(allotment)
        db.session.commit()
        flash('Allotment removed.', 'success')
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/notify', methods=['POST'])
@login_required
def admin_room_notify():
    allotments = RoomAllotment.query.all()
    pending = []
    for a in allotments:
        # If never notified OR room changed since last notification
        if not a.notified_room_number or a.notified_room_number != a.room.room_number:
            pending.append(a)
            
    if not pending:
        flash('All attendees are already notified of their current room assignments.', 'info')
        return redirect(url_for('admin_room_allotment'))
        
    success_count = 0
    fail_count = 0
    
    for a in pending:
        reg = a.registration
        room = a.room
        try:
            body_text = (
                f"Dear {reg.full_name},\n\n"
                f"With divine blessings, we are happy to inform you that your accommodation has been successfully allotted for the 3-Day Spiritual Program at Anantapur inspired by the teachings of Paramahansa Yogananda.\n\n"
                f"Accommodation Details:\n\n"
                f"Name: {reg.full_name}\n"
                f"Room Number: {room.room_number}\n"
                f"Number of Members: {room.capacity}-Bed Room\n"
                f"Check-In Date: {reg.arrival_date if reg.arrival_date else '24-07-2026'}\n"
                f"Check-Out Date: {reg.departure_date if reg.departure_date else '26-07-2026'}\n\n"
                f"We kindly request you to carry your registration confirmation during your visit and maintain the peaceful and spiritual atmosphere throughout the program.\n\n"
                f"May this sacred gathering bring peace, devotion, joy, and spiritual upliftment into your life.\n\n"
                f"We look forward to welcoming you with love and prayers.\n\n"
                f"Jai Guru"
            )
            msg = Message(
                subject=f"Accommodation Allotment Confirmation – YSS Anantapur",
                recipients=[reg.email],
                body=body_text
            )
            mail.send(msg)
            a.notified_room_number = room.room_number
            success_count += 1
        except Exception as e:
            print(f"Failed to notify {reg.email}: {e}")
            fail_count += 1
            
    db.session.commit()
    
    if fail_count > 0:
        flash(f"Successfully notified {success_count} attendees via Email. {fail_count} failed due to mail server issues.", 'warning')
    else:
        flash(f"Successfully notified {success_count} attendees via Email!", 'success')
        
    return redirect(url_for('admin_room_allotment'))

@app.route('/admin/room-allotment/reset-notifications', methods=['POST'])
@login_required
def admin_room_reset_notifications():
    allotments = RoomAllotment.query.all()
    for a in allotments:
        a.notified_room_number = None
    db.session.commit()
    flash('Email notification flags reset successfully.', 'success')
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

# ─── ADMIN WHATSAPP SETUP ───────────────────────────────────────────────────
@app.route('/admin/whatsapp-setup')
@login_required
def admin_whatsapp_setup():
    return render_template('admin/whatsapp_setup.html', config=app.config)

@app.route('/admin/whatsapp-status')
@login_required
def admin_whatsapp_status():
    import requests
    gateway_url = app.config.get('WHATSAPP_GATEWAY_URL', 'http://localhost:3000')
    status_data = {
        'status': 'Offline',
        'qr': None,
        'gateway_url': gateway_url
    }
    try:
        r = requests.get(f"{gateway_url}/status", timeout=0.8)
        if r.status_code == 200:
            res = r.json()
            status_data['status'] = res.get('status', 'Disconnected')
            status_data['qr'] = res.get('qr')
    except Exception as e:
        print(f"Error fetching WhatsApp gateway status: {e}")
    return jsonify(status_data)

# ─── MAIN ─────────────────────────────────────────────────────────────────────
with app.app_context():
    seed_data()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
