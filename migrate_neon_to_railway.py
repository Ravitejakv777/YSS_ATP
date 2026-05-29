import sys
import os
import argparse
from sqlalchemy import create_engine, MetaData
from sqlalchemy.orm import sessionmaker

# Add parent directory to path to import app and models
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from app import app
from models import db, Admin, Registration, Donation, EventSchedule, Room, RoomAllotment, ActivityLog, WhatsAppTemplate, GalleryImage

def migrate(source_url, target_url):
    print("Initializing migration...")
    print(f"Source (Neon): {source_url[:40]}...")
    print(f"Target (Railway): {target_url[:40]}...")

    # Set up source engine and session
    source_engine = create_engine(source_url)
    SourceSession = sessionmaker(bind=source_engine)
    source_session = SourceSession()

    # Set up target engine and session
    target_engine = create_engine(target_url)
    TargetSession = sessionmaker(bind=target_engine)
    target_session = TargetSession()

    # Verify connection to both
    try:
        source_session.execute(db.text("SELECT 1"))
        print("Connected to Source Database successfully!")
    except Exception as e:
        print(f"Error connecting to Source Database: {e}")
        return False

    try:
        target_session.execute(db.text("SELECT 1"))
        print("Connected to Target Database successfully!")
    except Exception as e:
        print(f"Error connecting to Target Database: {e}")
        return False

    # Ensure tables exist in target
    print("Ensuring target database tables exist...")
    # Bind SQLAlchemy metadata to target engine and create tables
    db.metadata.create_all(bind=target_engine)
    print("Target tables verified/created successfully.")

    # Define the models in dependency order
    models_to_migrate = [
        Admin,
        Room,
        Registration,
        RoomAllotment,
        Donation,
        EventSchedule,
        ActivityLog,
        WhatsAppTemplate,
        GalleryImage
    ]

    # Map tables for sequence reset in PostgreSQL
    table_names = {
        Admin: 'admins',
        Room: 'rooms',
        Registration: 'registrations',
        RoomAllotment: 'room_allotments',
        Donation: 'donations',
        EventSchedule: 'event_schedules',
        ActivityLog: 'activity_logs',
        WhatsAppTemplate: 'whatsapp_templates',
        GalleryImage: 'gallery_images'
    }

    try:
        # We migrate table by table
        for model in models_to_migrate:
            table_name = table_names[model]
            print(f"\nMigrating table '{table_name}'...")

            # Clear target table to prevent duplicates or unique constraint failures if re-run
            try:
                target_session.query(model).delete()
                target_session.commit()
            except Exception as e:
                target_session.rollback()
                print(f"Warning: Could not clear target table {table_name}: {e}")

            # Fetch all records from source
            source_records = source_session.query(model).all()
            total_records = len(source_records)
            print(f"Found {total_records} records in source '{table_name}'.")

            if total_records == 0:
                continue

            # Copy records
            copied_count = 0
            for record in source_records:
                # Create a new instance of the model with the same attributes
                # By using the mapper inspect, we get all column values
                attrs = {c.key: getattr(record, c.key) for c in model.__table__.columns}
                new_record = model(**attrs)
                target_session.add(new_record)
                copied_count += 1
                
                if copied_count % 100 == 0:
                    target_session.commit()
                    print(f"Copied {copied_count}/{total_records} records...")

            target_session.commit()
            print(f"Successfully migrated {copied_count} records to target '{table_name}'.")

            # Reset sequence if PostgreSQL (Railway/Neon are PostgreSQL)
            try:
                print(f"Resetting PostgreSQL identity sequence for table '{table_name}'...")
                # Since id is serial, we reset sequence value to max id
                seq_query = f"SELECT setval(pg_get_serial_sequence('{table_name}', 'id'), COALESCE((SELECT MAX(id) FROM {table_name}), 1))"
                target_session.execute(db.text(seq_query))
                target_session.commit()
                print(f"Sequence reset for '{table_name}' complete.")
            except Exception as seq_err:
                target_session.rollback()
                print(f"Note: Could not reset sequence for '{table_name}' (might not be PostgreSQL or sequence doesn't exist): {seq_err}")

        print("\n=========================================")
        print("MIGRATION COMPLETED SUCCESSFULLY!")
        print("=========================================")
        return True

    except Exception as main_err:
        target_session.rollback()
        print(f"\nMigration failed: {main_err}")
        return False
    finally:
        source_session.close()
        target_session.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Migrate data from Neon PostgreSQL to Railway PostgreSQL.")
    parser.add_argument("--source", help="Source Neon database URL (overrides .env or command line prompt)")
    parser.add_argument("--target", help="Target Railway database URL (overrides env/prompt)")

    args = parser.parse_args()

    # Detect Neon URL
    source = args.source or os.environ.get("NEON_DATABASE_URL")
    if not source:
        # Fallback to the one provided by user
        source = "postgresql://neondb_owner:npg_p8NMRz0KxkEU@ep-broad-cloud-aoefntjz-pooler.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require"

    target = args.target or os.environ.get("DATABASE_URL")

    if not target:
        print("ERROR: Target database URL (Railway) not specified.")
        print("Please provide it using '--target <url>' or set the 'DATABASE_URL' environment variable.")
        sys.exit(1)

    # Standardize 'postgres://' to 'postgresql://' for SQLAlchemy
    if source.startswith("postgres://"):
        source = source.replace("postgres://", "postgresql://", 1)
    if target.startswith("postgres://"):
        target = target.replace("postgres://", "postgresql://", 1)

    success = migrate(source, target)
    if not success:
        sys.exit(1)
