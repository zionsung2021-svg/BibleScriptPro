import sqlite3

conn = sqlite3.connect("bible.db")
cur = conn.cursor()

# 1. 시편 관련 책 이름 확인
cur.execute("SELECT DISTINCT book FROM bible WHERE book LIKE '%시%'")
books = cur.fetchall()
print("▶ DB에 등록된 책 이름:", books)

# 2. 23편 구절 샘플 조회
cur.execute("SELECT book, chapter, verse, content FROM bible WHERE chapter = 23 LIMIT 3")
sample = cur.fetchall()
print("▶ 23장/편 샘플 데이터:", sample)

conn.close()