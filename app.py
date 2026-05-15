from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file, make_response
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_mail import Mail, Message
from models import db, Admin, Registration, Donation, EventSchedule
from config import Config
from datetime import datetime
import os, openpyxl

app = Flask(__name__)
app.config.from_object(Config)
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db.init_app(app)
mail = Mail(app)

login_manager = LoginManager(app)
login_manager.login_view = 'admin_login'
login_manager.login_message = 'Please login to access admin panel.'

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
            mobile='9494457607'
        )
        admin.set_password(app.config['ADMIN_PASSWORD'])
        db.session.add(admin)

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
        
        headers = ['S.No','Reg ID','Lesson No','Full Name','Gender','Age','WhatsApp','Email','City/Town','State',
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
        msg = Message(
            subject='Registration Confirmed – YSS 3-Day Spiritual Program, Anantapur',
            recipients=[app.config['MAIL_USERNAME']],
            html=render_template('email_reg.html', reg=reg, config=app.config)
        )
        mail.send(msg)
    except Exception as e:
        app.logger.warning(f'Email send failed: {e}')

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
        state = request.form.get('state', '').strip()
        email = request.form.get('email', '').strip()
        country_code = request.form.get('country_code', '+91').strip()
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
        if not email or '@' not in email: errors.append('Valid Email Address is required.')
        if not whatsapp or not whatsapp.isdigit() or len(whatsapp) != 10: errors.append('WhatsApp Number must be exactly 10 digits.')
        if not arrival_date: errors.append('Date of Arrival is required.')
        if not departure_date: errors.append('Date of Departure is required.')
        if not payment_mode: errors.append('Payment Mode is required.')
        if not transaction_id: errors.append('Transaction ID is required.')


        if errors:
            for e in errors:
                flash(e, 'error')
            return render_template('registration.html', config=app.config, form=request.form)

        reg = Registration(
            lesson_no=lesson_no, full_name=full_name, gender=gender,
            age=int(age), place=place, state=state, email=email, country_code=country_code, whatsapp=whatsapp,
            is_kriyaban=is_kriyaban, accommodation=accommodation,
            volunteer=volunteer, arrival_date=arrival_date,
            departure_date=departure_date, payment_mode=payment_mode,
            transaction_id=transaction_id, payment_screenshot=screenshot_filename
        )
        db.session.add(reg)
        db.session.commit()
        update_registrations_excel()
        send_registration_email(reg)
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
    return render_template('donation_success.html', don=don, config=app.config)

# ─── ADMIN AUTH ───────────────────────────────────────────────────────────────
@app.route('/admin')
def admin_redirect():
    return redirect(url_for('admin_login'))

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '').strip()
        admin = Admin.query.filter_by(email=email).first()
        if admin and admin.check_password(password):
            login_user(admin)
            return redirect(url_for('admin_registrations'))
        flash('Invalid email or password.', 'error')
    return render_template('admin/login.html', config=app.config)

@app.route('/admin/logout')
@login_required
def admin_logout():
    logout_user()
    flash('Logged out successfully.', 'success')
    return redirect(url_for('admin_login'))



# ─── ADMIN REGISTRATIONS ──────────────────────────────────────────────────────
@app.route('/admin/registrations')
@login_required
def admin_registrations():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    pay_status = request.args.get('pay_status', '')
    reg_status = request.args.get('reg_status', '')
    q = Registration.query
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%'),
                             Registration.reg_id.ilike(f'%{search}%')))
    if pay_status:
        q = q.filter_by(payment_status=pay_status)
    if reg_status:
        q = q.filter_by(reg_status=reg_status)
    pagination = q.order_by(Registration.id.desc()).paginate(page=page, per_page=10)
    return render_template('admin/registrations.html', pagination=pagination,
                           search=search, pay_status=pay_status, reg_status=reg_status, config=app.config)

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
    pay_status = request.args.get('pay_status', '')
    q = Donation.query
    if search:
        q = q.filter(db.or_(Donation.name.ilike(f'%{search}%'),
                             Donation.whatsapp.ilike(f'%{search}%'),
                             Donation.donation_id.ilike(f'%{search}%')))
    if pay_status:
        q = q.filter_by(payment_status=pay_status)
    pagination = q.order_by(Donation.id.desc()).paginate(page=page, per_page=10)
    return render_template('admin/donations.html', pagination=pagination,
                           search=search, pay_status=pay_status, config=app.config)

@app.route('/admin/donations/export')
@login_required
def export_donations():
    update_donations_excel()
    path = os.path.join(app.config['EXPORTS_DIR'], 'donations.xlsx')
    return send_file(path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='donations.xlsx')

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
    reg_status = request.args.get('reg_status', '')
    q = Registration.query
    if search:
        q = q.filter(db.or_(Registration.full_name.ilike(f'%{search}%'),
                             Registration.whatsapp.ilike(f'%{search}%')))
    if reg_status:
        q = q.filter_by(reg_status=reg_status)
    pagination = q.order_by(Registration.id).paginate(page=page, per_page=6)
    return render_template('admin/id_cards.html', pagination=pagination,
                           search=search, reg_status=reg_status, config=app.config)

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

# ─── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        seed_data()
    app.run(debug=True, port=5000)
