import sqlite3
import csv

conn = sqlite3.connect("pokemon.db")
cur = conn.cursor()

with open("data/pokemon.csv", encoding="utf-8") as f:
    reader = csv.reader(f)
    next(reader)  # ヘッダーを飛ばす

    # 既存データ削除
    cur.execute("DELETE FROM pokemon")

    # csvの全行をpokemon.dbに追加
    for row in reader:  
        cur.execute(
            "INSERT INTO pokemon VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            row
        )

conn.commit()
conn.close()