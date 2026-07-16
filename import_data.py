import sqlite3
import csv

# pokemon.db
pokemon_conn = sqlite3.connect("pokemon.db")
pokemon_cur = pokemon_conn.cursor()

with open("data/pokemon.csv", encoding="utf-8") as f:
    pokemon_reader = csv.reader(f)
    next(pokemon_reader)  # ヘッダーを飛ばす

    # 既存データ削除
    pokemon_cur.execute("DELETE FROM pokemon")

    # csvの全行をpokemon.dbに追加
    for row in pokemon_reader:  
        pokemon_cur.execute(
            "INSERT INTO pokemon VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            row
        )

pokemon_conn.commit()
pokemon_conn.close()


# move.db
move_conn = sqlite3.connect("move.db")
move_cur = move_conn.cursor()

with open("data/move.csv", encoding="utf-8") as f:
    reader = csv.reader(f)
    next(reader)  # ヘッダーを飛ばす

    # 既存データ削除
    move_cur.execute("DELETE FROM move")

    # csvの全行をmove.dbに追加
    for row in reader:  
        move_cur.execute(
            "INSERT INTO move VALUES (?, ?, ?, ?)",
            row
        )

move_conn.commit()
move_conn.close()