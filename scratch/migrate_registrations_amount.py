import sqlite3
db_path = 'c:/Users/megha/.gemini/antigravity/scratch/yss-anantapur/instance/yss.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
try:
    cursor.execute('ALTER TABLE registrations ADD COLUMN amount FLOAT')
    print("Added amount to registrations table")
except Exception as e:
    print(f"Error: {e}")
conn.commit()
conn.close()
