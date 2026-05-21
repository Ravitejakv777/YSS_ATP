import sqlite3

def alter_table():
    conn = sqlite3.connect('instance/database.db')
    c = conn.cursor()
    # Create new table without UNIQUE on email
    c.execute("""
    CREATE TABLE admins_new (
        id INTEGER NOT NULL, 
        lesson_no VARCHAR(20) NOT NULL, 
        name VARCHAR(100) NOT NULL, 
        email VARCHAR(120) NOT NULL, 
        mobile VARCHAR(15) NOT NULL, 
        password_hash VARCHAR(256) NOT NULL, 
        is_main_admin BOOLEAN, 
        reset_token VARCHAR(100), 
        reset_token_expiry DATETIME, 
        last_active DATETIME, 
        created_at DATETIME, 
        PRIMARY KEY (id)
    );
    """)
    # Copy data
    c.execute("INSERT INTO admins_new SELECT id, lesson_no, name, email, mobile, password_hash, is_main_admin, reset_token, reset_token_expiry, last_active, created_at FROM admins;")
    # Drop old table
    c.execute("DROP TABLE admins;")
    # Rename new table
    c.execute("ALTER TABLE admins_new RENAME TO admins;")
    conn.commit()
    print("Table 'admins' altered successfully!")

alter_table()
