from app import app, db
from sqlalchemy import text

def migrate():
    with app.app_context():
        try:
            # Add is_main_admin column
            db.session.execute(text("ALTER TABLE admins ADD COLUMN IF NOT EXISTS is_main_admin BOOLEAN DEFAULT FALSE"))
            # Add reset_token column
            db.session.execute(text("ALTER TABLE admins ADD COLUMN IF NOT EXISTS reset_token VARCHAR(100)"))
            # Add reset_token_expiry column
            db.session.execute(text("ALTER TABLE admins ADD COLUMN IF NOT EXISTS reset_token_expiry TIMESTAMP"))
            
            db.session.commit()
            print("Migration successful: Added is_main_admin, reset_token, and reset_token_expiry to admins table.")
        except Exception as e:
            db.session.rollback()
            print(f"Migration failed: {e}")

if __name__ == "__main__":
    migrate()
