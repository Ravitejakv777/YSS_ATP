import sqlite3

def check_indexes():
    conn = sqlite3.connect('instance/database.db')
    c = conn.cursor()
    indexes = c.execute("PRAGMA index_list('admins');").fetchall()
    print("Indexes on admins table:", indexes)
    for idx in indexes:
        if idx[2] == 1: # unique index
            idx_name = idx[1]
            cols = c.execute(f"PRAGMA index_info('{idx_name}');").fetchall()
            print(f"Index {idx_name} columns: {cols}")
            for col in cols:
                if col[2] == 'email':
                    print(f"Dropping unique index {idx_name}")
                    c.execute(f"DROP INDEX IF EXISTS {idx_name}")
                    conn.commit()
                    return
    print("No unique index found on admins email.")

check_indexes()
