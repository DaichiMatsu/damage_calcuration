import sqlite3

conn = sqlite3.connect("pokemon.db")
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS pokemon(
    No INTEGER,
    name TEXT,
    hp INTEGER,
    attack INTEGER,
    defense INTEGER,
    sp_attack INTEGER,
    sp_defense INTEGER,
    speed INTEGER,
    type1 TEXT,
    type2 TEXT
)
""")

conn.commit()
conn.close()