import os
import re
import sqlite3

DB_PATH = os.path.join("backend", "bible.db")

ENG_TO_KOR_BOOK = {
    "Genesis": "창세기", "Exodus": "출애굽기", "Leviticus": "레위기", "Numbers": "민수기",
    "Deuteronomy": "신명기", "Joshua": "여호수아", "Judges": "사사기", "Ruth": "룻기",
    "1 Samuel": "사무엘상", "2 Samuel": "사무엘하", "1 Kings": "열왕기상", "2 Kings": "열왕기하",
    "1 Chronicles": "역대상", "2 Chronicles": "역대하", "Ezra": "에스라", "Nehemiah": "느헤미야",
    "Esther": "에스더", "Job": "욥기", "Psalms": "시편", "Psalm": "시편", "Proverbs": "잠언",
    "Ecclesiastes": "전도서", "Song of Songs": "아가", "Song of Solomon": "아가", "Isaiah": "이사야",
    "Jeremiah": "예레미야", "Lamentations": "예레미야애가", "Ezekiel": "에스겔", "Daniel": "다니엘",
    "Hosea": "호세아", "Joel": "요엘", "Amos": "아모스", "Obadiah": "오바댜", "Jonah": "요나",
    "Micah": "미가", "Nahum": "나훔", "Habakkuk": "하박국", "Zephaniah": "스바냐", "Haggai": "학개",
    "Zechariah": "스가랴", "Malachi": "말라기", "Matthew": "마태복음", "Mark": "마가복음",
    "Luke": "누가복음", "John": "요한복음", "Acts": "사도행전", "Romans": "로마서",
    "1 Corinthians": "고린도전서", "2 Corinthians": "고린도후서", "Galatians": "갈라디아서",
    "Ephesians": "에베소서", "Philippians": "빌립보서", "Colossians": "골로새서",
    "1 Thessalonians": "데살로니가전서", "2 Thessalonians": "데살로니가후서",
    "1 Timothy": "디모데전서", "2 Timothy": "디모데후서", "Titus": "디도서", "Philemon": "빌레몬서",
    "Hebrews": "히브리서", "James": "야고보서", "1 Peter": "베드로전서", "2 Peter": "베드로후서",
    "1 John": "요한일서", "2 John": "요한이서", "3 John": "요한삼서", "Jude": "유다서", "Revelation": "요한계시록"
}

KOR_ABBR_MAP = {
    "창": "창세기", "출": "출애굽기", "레": "레위기", "민": "민수기", "신": "신명기",
    "수": "여호수아", "삿": "사사기", "룻": "룻기", "삼상": "사무엘상", "삼하": "사무엘하",
    "왕상": "열왕기상", "왕하": "열왕기하", "대상": "역대상", "대하": "역대하", "스": "에스라",
    "느": "느헤미야", "에": "에스더", "욥": "욥기", "시": "시편", "잠": "잠언", "전": "전도서",
    "아": "아가", "사": "이사야", "렘": "예레미야", "애": "예레미야애가", "겔": "에스겔",
    "단": "다니엘", "호": "호세아", "욜": "요엘", "암": "아모스", "옵": "오바댜", "욘": "요나",
    "미": "미가", "나": "나훔", "합": "하박국", "습": "스바냐", "학": "학개", "슥": "스가랴",
    "말": "말라기", "마": "마태복음", "막": "마가복음", "눅": "누가복음", "요": "요한복음",
    "행": "사도행전", "롬": "로마서", "고전": "고린도전서", "고후": "고린도후서", "갈": "갈라디아서",
    "엡": "에베소서", "빌": "빌립보서", "골": "골로새서", "살전": "데살로니가전서", "살후": "데살로니가후서",
    "딤전": "디모데전서", "딤후": "디모데후서", "딛": "디도서", "몬": "빌레몬서", "히": "히브리서",
    "약": "야고보서", "벧전": "베드로전서", "벧후": "베드로후서", "요일": "요한일서", "요이": "요한이서",
    "요삼": "요한삼서", "유": "유다서", "계": "요한계시록"
}

def read_text_file(filename):
    for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
        if os.path.exists(filename):
            try:
                with open(filename, "r", encoding=enc) as f:
                    return f.readlines()
            except UnicodeDecodeError:
                continue
    return []

def main():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('bibles', 'bible')")
    table_name = cur.fetchone()[0]

    cur.execute(f"PRAGMA table_info({table_name})")
    cols = [c[1] for c in cur.fetchall()]
    verse_col = "verse_num" if "verse_num" in cols else "verse"
    text_col = "text" if "text" in cols else "content"

    # version 컬럼 없을 시 생성 및 기본값 채우기
    if "version" not in cols:
        print("[INFO] 'version' 컬럼을 생성합니다.")
        cur.execute(f"ALTER TABLE {table_name} ADD COLUMN version TEXT DEFAULT '개역개정'")
        cur.execute(f"UPDATE {table_name} SET version = '개역개정' WHERE version IS NULL")
        conn.commit()

    # 중복 제거 후 적재 준비
    cur.execute(f"DELETE FROM {table_name} WHERE version IN ('NIV', '표준새번역')")
    conn.commit()

    # NIV 파싱
    niv_lines = read_text_file("NIV성경.txt")
    niv_records = []
    current_book, current_ch = "", 0
    h_re = re.compile(r"^\[([0-9A-Za-z\s]+)\s+(\d+)\]")
    v_re = re.compile(r"^(\d+)\.(.*)$")

    for line in niv_lines:
        line = line.strip()
        if not line:
            continue
        hm = h_re.match(line)
        if hm:
            eng_b = hm.group(1).strip()
            current_ch = int(hm.group(2))
            current_book = ENG_TO_KOR_BOOK.get(eng_b, eng_b)
            continue
        vm = v_re.match(line)
        if vm and current_book:
            niv_records.append(("NIV", current_book, current_ch, int(vm.group(1)), vm.group(2).strip()))

    if niv_records:
        cur.executemany(f"INSERT INTO {table_name} (version, book, chapter, {verse_col}, {text_col}) VALUES (?, ?, ?, ?, ?)", niv_records)
        print(f"[OK] NIV: {len(niv_records)}구절 적재 완료.")

    # 표준새번역 파싱
    sae_lines = read_text_file("표준새번역성경.txt")
    sae_records = []
    sae_re = re.compile(r"^([가-힣0-9]+)(\d+):(\d+)\s*(.*)$")

    for line in sae_lines:
        line = line.strip()
        if not line:
            continue
        sm = sae_re.match(line)
        if sm:
            raw_abbr = sm.group(1)
            ch = int(sm.group(2))
            v_num = int(sm.group(3))
            raw_text = sm.group(4)
            # <소제목>, (각주) 제거
            cleaned = re.sub(r"<[^>]+>", "", raw_text)
            cleaned = re.sub(r"\([a-zA-Z0-9\.\s>\"'~,-]+\)", "", cleaned)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            book_name = KOR_ABBR_MAP.get(raw_abbr, raw_abbr)
            sae_records.append(("표준새번역", book_name, ch, v_num, cleaned))

    if sae_records:
        cur.executemany(f"INSERT INTO {table_name} (version, book, chapter, {verse_col}, {text_col}) VALUES (?, ?, ?, ?, ?)", sae_records)
        print(f"[OK] 표준새번역: {len(sae_records)}구절 적재 완료.")

    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_bible_ver_search ON {table_name} (version, book, chapter, {verse_col})")
    conn.commit()
    conn.close()
    print("[SUCCESS] DB 업데이트가 성공적으로 완료되었습니다!")

if __name__ == "__main__":
    main()