import sqlite3
import os

db_path = 'c:/Users/megha/.gemini/antigravity/scratch/yss-anantapur/instance/yss.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    cursor.execute('ALTER TABLE donations ADD COLUMN transaction_id VARCHAR(100)')
    print("Added transaction_id to donations table")
except Exception as e:
    print(f"Error adding transaction_id: {e}")

try:
    cursor.execute('ALTER TABLE donations ADD COLUMN payment_screenshot VARCHAR(255)')
    print("Added payment_screenshot to donations table")
except Exception as e:
    print(f"Error adding payment_screenshot: {e}")

conn.commit()
conn.close()
