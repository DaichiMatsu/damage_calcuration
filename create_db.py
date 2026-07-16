import sqlite3

# pokemon.db
pokemon_conn = sqlite3.connect("pokemon.db")
pokemon_cur = pokemon_conn.cursor()

pokemon_cur.execute("""
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

pokemon_conn.commit()
pokemon_conn.close()

# move.db
move_conn = sqlite3.connect("move.db")
move_cur = move_conn.cursor()

move_cur.execute("""
CREATE TABLE IF NOT EXISTS move(
    name TEXT,
    type TEXT,
    class TEXT,
    power INTEGER
    
)
""")

move_conn.commit()
move_conn.close()
