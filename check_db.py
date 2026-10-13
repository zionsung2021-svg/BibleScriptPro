# -*- coding: utf-8 -*-
import sqlite3, os

paths = ["C:/BibleScript/backend/bible.db", "C:/BibleScript/bible.db"]
for p in paths:
    if os.path.exists(p):
        size = os.path.getsize(p)
        print("PATH:", p, "SIZE:", size)
        conn = sqlite3.connect(p)
        tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        print("TABLES:", tables)
        for t in tables:
            t_name = t[0]
            if t_name != 'sqlite_sequence':
                cols = conn.execute(f"PRAGMA table_info({t_name});").fetchall()
                print("COLUMNS in", t_name, ":", [c[1] for c in cols])
                sample = conn.execute(f"SELECT * FROM {t_name} LIMIT 1;").fetchone()
                print("SAMPLE:", sample)
