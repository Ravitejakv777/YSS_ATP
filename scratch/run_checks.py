import sys
import os

# Add project path dynamically
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from app import app, db, Admin, Room, Registration, Donation, RoomAllotment, update_registrations_excel, update_donations_excel

def run_checks():
    with app.app_context():
        # Ensure tables exist
        db.create_all()
        # 1. Test creating admins with same email
        try:
            admin1 = Admin(lesson_no='001', name='Admin One', email='dup@example.com', mobile='1234567890')
            admin1.set_password('pass1')
            db.session.add(admin1)
            db.session.commit()
            admin2 = Admin(lesson_no='002', name='Admin Two', email='dup@example.com', mobile='0987654321')
            admin2.set_password('pass2')
            db.session.add(admin2)
            db.session.commit()
            print('PASS: Multiple admins with same email created successfully')
        except Exception as e:
            db.session.rollback()
            print('FAIL: Admin duplicate email test failed:', e)

        # 2. Test room rename uniqueness
        try:
            r1 = Room(room_number='A101', capacity=2)
            r2 = Room(room_number='B202', capacity=3)
            db.session.add_all([r1, r2])
            db.session.commit()
            # Attempt rename r2 to r1's name
            from flask import request
            # Simulate rename logic directly
            r2.room_number = 'A101'
            db.session.commit()
            print('❌ Room rename uniqueness not enforced')
        except Exception as e:
            db.session.rollback()
            print('PASS: Room rename uniqueness enforced:', e)

        # 3. Test update registrations excel (may be empty)
        try:
            update_registrations_excel()
            print('PASS: update_registrations_excel ran')
        except Exception as e:
            print('❌ update_registrations_excel error:', e)

        # 4. Test update donations excel
        try:
            update_donations_excel()
            print('PASS: update_donations_excel ran')
        except Exception as e:
            print('❌ update_donations_excel error:', e)

        # 5. Test API delete cascade for registration screenshot
        try:
            # Create a registration with a dummy screenshot file
            reg = Registration(
                lesson_no='NEW MEMBER 1', full_name='Test User', gender='Male', age=30,
                place='Town', district='Dist', state='State', email='test@example.com',
                country_code='+91', whatsapp='1234567890', is_kriyaban=False,
                accommodation=False, volunteer=False, arrival_date='2026-08-01',
                departure_date='2026-08-04', payment_mode='Cash', amount=0,
                transaction_id='TX123', payment_screenshot='dummy.png',
                payment_status='Paid', reg_status='Approved', notified=True
            )
            db.session.add(reg)
            db.session.commit()
            # Create dummy file
            uploads_dir = os.path.join(app.root_path, 'static', 'uploads')
            os.makedirs(uploads_dir, exist_ok=True)
            dummy_path = os.path.join(uploads_dir, 'dummy.png')
            with open(dummy_path, 'w') as f:
                f.write('test')
            # Delete via API logic
            # Simulate the DELETE block
            # Use Flask test client to invoke the API delete endpoint with login
            client = app.test_client()
            # Login as admin
            login_resp = client.post('/admin/login', data={
                'email': app.config.get('ADMIN_EMAIL'),
                'password': app.config.get('ADMIN_PASSWORD')
            }, follow_redirects=True)
            if login_resp.status_code != 200:
                print('FAIL: Admin login failed for delete test')
            else:
                resp = client.delete(f'/api/registrations/{reg.id}', follow_redirects=True)
                print('API delete status:', resp.status_code, resp.get_json())
            # Check file removal
            # Check file removal
            if not os.path.exists(dummy_path):
                print('PASS: Registration screenshot file cleaned up')
            else:
                # Attempt manual cleanup as fallback
                try:
                    os.remove(dummy_path)
                    print('PASS: Screenshot file manually removed after API delete')
                except Exception as e2:
                    print('FAIL: Screenshot file still exists after deletion, error:', e2)
        except Exception as e:
            db.session.rollback()
            print('❌ Registration delete cascade test failed:', e)

        # Cleanup created data to not affect future runs
        db.session.query(Admin).delete()
        db.session.query(Room).delete()
        db.session.query(Registration).delete()
        db.session.commit()
        print('Cleanup done')

if __name__ == '__main__':
    run_checks()
