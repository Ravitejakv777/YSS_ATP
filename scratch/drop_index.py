import sqlite3

def drop_unique_email():
    conn = sqlite3.connect('instance/database.db')
    c = conn.cursor()
    # Find the unique index for the email column in admin table
    indexes = c.execute("PRAGMA index_list('admin');").fetchall()
    for idx in indexes:
        if idx[2] == 1: # unique index
            idx_name = idx[1]
            cols = c.execute(f"PRAGMA index_info('{idx_name}');").fetchall()
            for col in cols:
                if col[2] == 'email':
                    print(f"Dropping unique index {idx_name}")
                    c.execute(f"DROP INDEX IF EXISTS {idx_name}")
                    conn.commit()
                    return
    print("No unique index found on admin email. It might be defined inline.")

drop_unique_email()
