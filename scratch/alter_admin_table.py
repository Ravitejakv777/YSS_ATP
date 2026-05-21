import sqlite3

def alter_table():
    conn = sqlite3.connect('instance/database.db')
    c = conn.cursor()
    # Create new table without UNIQUE on email
    c.execute("""
    CREATE TABLE admin_new (
        id INTEGER NOT NULL, 
        lesson_no VARCHAR(50), 
        name VARCHAR(100) NOT NULL, 
        email VARCHAR(120) NOT NULL, 
        mobile VARCHAR(20) NOT NULL, 
        password_hash VARCHAR(256), 
        is_main_admin BOOLEAN, 
        PRIMARY KEY (id)
    );
    """)
    # Copy data
    c.execute("INSERT INTO admin_new SELECT * FROM admin;")
    # Drop old table
    c.execute("DROP TABLE admin;")
    # Rename new table
    c.execute("ALTER TABLE admin_new RENAME TO admin;")
    conn.commit()
    print("Table altered successfully!")

alter_table()
