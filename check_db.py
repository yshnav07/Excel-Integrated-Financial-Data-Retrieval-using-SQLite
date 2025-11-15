import sqlite3
DB = r"C:\Users\vaish\Downloads\prices.db"
connection=sqlite3.connect(DB)
connection.row_factory = sqlite3.Row
cur = connection.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
print("Tables:", [r["name"] for r in cur.fetchall()])

cur.execute("SELECT * FROM prices LIMIT 5;")
for row in cur.fetchall():
    print(dict(row))

connection.close()