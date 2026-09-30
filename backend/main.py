# backend/main.py
import io
import os
import sys
import tempfile
import urllib.parse
import sqlite3
import threading
import time
from typing import Optional, List
from fastapi import FastAPI, Query, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import StreamingResponse, FileResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

app = FastAPI(title="BibleScript Pro API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# PyInstaller 번들 경로 및 로컬 경로 자동 탐색
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

possible_paths = [
    os.path.join(BASE_DIR, "backend", "bible.db"),
    os.path.join(BASE_DIR, "bible.db"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "bible.db"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "backend", "bible.db"),
]

DB_PATH = None
for p in possible_paths:
    if os.path.exists(p):
        DB_PATH = p
        break

if not DB_PATH:
    DB_PATH = os.path.join(BASE_DIR, "backend", "bible.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ==================== 헬퍼 유틸리티 함수 ====================

def set_korean_font(paragraph, font_name: str, font_size_pt: int, color_rgb: RGBColor, is_bold: bool = False):
    """문단 폰트 및 한국어 폰트 속성 적용"""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = color_rgb
        run.font.bold = is_bold
        # 파워포인트 동아시아 폰트 명시
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
        if ea is None:
            ea = rPr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
            rPr.append(ea)
        ea.set("typeface", font_name)


def split_verse_text(text: str, max_chars: int = 65) -> List[str]:
    """긴 구절을 읽기 좋게 문장 부호 단위로 분할"""
    if len(text) <= max_chars:
        return [text]

    chunks = []
    current_chunk = ""
    # 마침표, 쉼표, 세미콜론 뒤 분할
    sentences = [s.strip() for s in text.replace(". ", ".\n").replace(", ", ",\n").split("\n") if s.strip()]

    for s in sentences:
        if len(current_chunk) + len(s) + 1 <= max_chars:
            current_chunk = f"{current_chunk} {s}".strip()
        else:
            if current_chunk:
                chunks.append(current_chunk)
            current_chunk = s

    if current_chunk:
        chunks.append(current_chunk)

    return chunks if chunks else [text]


# ==================== 프레젠테이션 생성 엔진 ====================

def build_presentation(verses_by_version: dict, canonical_book: str, chapter: int, unit: str, font_name: str, versions: List[str]):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_color = RGBColor(0, 32, 96)          # Deep Navy Blue
    title_color = RGBColor(255, 192, 0)     # Gold Yellow
    body_color = RGBColor(255, 255, 255)    # White
    label_color = RGBColor(200, 225, 255)   # Light Sky Blue

    is_bold = not any(w in font_name for w in ["Regular", "Light", "Medium"])
    num_versions = len(versions)
    base_version = versions[0]
    base_verses = verses_by_version.get(base_version, [])

    for v in base_verses:
        verse_num = v["verse_num"]

        # [1] 단일 번역본 (1개)
        if num_versions == 1:
            chunks = split_verse_text(v["text"], max_chars=65)
            for chunk in chunks:
                slide = prs.slides.add_slide(blank_layout)
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = bg_color

                # 상단 제목 (50pt)
                title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
                tf_title = title_box.text_frame
                tf_title.word_wrap = True
                p_title = tf_title.paragraphs[0]
                p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
                p_title.alignment = PP_ALIGN.CENTER
                set_korean_font(p_title, font_name, 50, title_color, is_bold=True)

                # 중앙 본문 (52pt)
                content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.8))
                tf_content = content_box.text_frame
                tf_content.word_wrap = True
                p_content = tf_content.paragraphs[0]
                p_content.text = chunk
                p_content.alignment = PP_ALIGN.CENTER
                p_content.line_spacing = 1.25
                set_korean_font(p_content, font_name, 52, body_color, is_bold=is_bold)

        # [2] 다중 번역본 대조 (2개 또는 3개 이상 완벽 비율 대응)
        else:
            slide = prs.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = bg_color

            # 상단 제목 (46pt)
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
            tf_title = title_box.text_frame
            tf_title.word_wrap = True
            p_title = tf_title.paragraphs[0]
            p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
            p_title.alignment = PP_ALIGN.CENTER
            set_korean_font(p_title, font_name, 46, title_color, is_bold=True)

            left_margin = Inches(0.8)
            available_width = 11.733
            gap = 0.4
            col_width_val = (available_width - (gap * (num_versions - 1))) / num_versions
            col_width = Inches(col_width_val)
            body_font_size = 34 if num_versions == 2 else 26

            for idx, ver in enumerate(versions):
                col_left = left_margin + Inches(idx * (col_width_val + gap))
                ver_verse = next((item for item in verses_by_version.get(ver, []) if item["verse_num"] == verse_num), None)
                ver_text = ver_verse["text"] if ver_verse else ""

                box = slide.shapes.add_textbox(col_left, Inches(1.5), col_width, Inches(5.3))
                tf = box.text_frame
                tf.word_wrap = True

                p_label = tf.paragraphs[0]
                p_label.text = f"[{ver}]"
                set_korean_font(p_label, font_name, 26 if num_versions == 2 else 22, label_color, is_bold=True)

                p_body = tf.add_paragraph()
                p_body.text = ver_text
                p_body.line_spacing = 1.25
                set_korean_font(p_body, font_name, body_font_size, body_color, is_bold=is_bold)

    return prs


# ==================== 엔드포인트 라우팅 ====================

@app.get("/api/versions")
def get_versions():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
        table_row = cursor.fetchone()
        if not table_row:
            return {"versions": ["개역개정", "표준새번역", "NIV"]}
        table_name = table_row[0]

        cursor.execute(f"PRAGMA table_info({table_name})")
        cols = [c[1] for c in cursor.fetchall()]
        if "version" in cols:
            cursor.execute(f"SELECT DISTINCT version FROM {table_name} ORDER BY version ASC")
            versions = [r[0] for r in cursor.fetchall() if r[0]]
            return {"versions": versions if versions else ["개역개정", "표준새번역", "NIV"]}
        return {"versions": ["개역개정"]}


@app.get("/api/search-keyword")
def search_keyword(
    keyword: str = Query(..., min_length=2),
    version: str = Query("개역개정"),
    limit: int = Query(25)
):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
        table_name = cursor.fetchone()[0]

        cursor.execute(f"PRAGMA table_info({table_name})")
        cols = [c[1] for c in cursor.fetchall()]
        verse_col = "verse_num" if "verse_num" in cols else "verse"
        text_col = "text" if "text" in cols else "content"

        if "version" in cols:
            query = f"SELECT version, book, chapter, {verse_col} AS verse_num, {text_col} AS text FROM {table_name} WHERE version = ? AND {text_col} LIKE ? LIMIT ?"
            params = (version, f"%{keyword}%", limit)
        else:
            query = f"SELECT '개역개정' AS version, book, chapter, {verse_col} AS verse_num, {text_col} AS text FROM {table_name} WHERE {text_col} LIKE ? LIMIT ?"
            params = (f"%{keyword}%", limit)

        cursor.execute(query, params)
        return [dict(r) for r in cursor.fetchall()]


class VerseUpdateRequest(BaseModel):
    version: str = "개역개정"
    book: str
    chapter: int
    verse_num: int
    text: str

@app.patch("/api/verse")
def update_verse_db(payload: VerseUpdateRequest):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
        table_name = cursor.fetchone()[0]

        cursor.execute(f"PRAGMA table_info({table_name})")
        cols = [c[1] for c in cursor.fetchall()]
        verse_col = "verse_num" if "verse_num" in cols else "verse"
        text_col = "text" if "text" in cols else "content"

        if "version" in cols:
            query = f"UPDATE {table_name} SET {text_col} = ? WHERE version = ? AND book = ? AND chapter = ? AND {verse_col} = ?"
            params = (payload.text, payload.version, payload.book, payload.chapter, payload.verse_num)
        else:
            query = f"UPDATE {table_name} SET {text_col} = ? WHERE book = ? AND chapter = ? AND {verse_col} = ?"
            params = (payload.text, payload.book, payload.chapter, payload.verse_num)

        cursor.execute(query, params)
        conn.commit()
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="해당 구절을 찾을 수 없습니다.")
    return {"status": "success", "message": f"[{payload.version}] 본문이 성공적으로 저장되었습니다."}


@app.post("/api/shutdown")
def shutdown_app():
    def kill_process():
        time.sleep(0.5)
        os._exit(0)
    threading.Thread(target=kill_process, daemon=True).start()
    return {"status": "success", "message": "프로그램을 종료합니다."}

# 성경 한 글자/두 글자 약어 변환 맵
BIBLE_ABBR_MAP = {
    "창": "창세기", "출": "출애굽기", "레": "레위기", "민": "민수기", "신": "신명기",
    "수": "여호수아", "삿": "사사기", "룻": "룻기", "삼상": "사무엘상", "삼하": "사무엘하",
    "왕상": "열왕기상", "왕하": "열왕기하", "대상": "역대상", "대하": "역대하",
    "스": "에스라", "느": "느헤미야", "에": "에스더", "욥": "욥기", "시": "시편",
    "잠": "잠언", "전": "전도서", "아": "아가", "사": "이사야", "렘": "예레미야",
    "애": "예레미야애가", "겔": "에스겔", "단": "다니엘", "호": "호세아", "욜": "요엘",
    "암": "아모스", "옵": "오바댜", "욘": "요나", "미": "미가", "나": "나훔",
    "합": "하박국", "습": "스바냐", "학": "학개", "슥": "스가랴", "말": "말라기",
    "마": "마태복음", "막": "마가복음", "눅": "누가복음", "요": "요한복음",
    "행": "사도행전", "롬": "로마서", "고전": "고린도전서", "고후": "고린도후서",
    "갈": "갈라디아서", "엡": "에베소서", "빌": "빌립보서", "골": "골로새서",
    "살전": "데살로니가전서", "살후": "데살로니가후서", "딤전": "디모데전서", "딤후": "디모데후서",
    "딛": "디도서", "몬": "빌레몬서", "히": "히브리서", "약": "야고보서",
    "벧전": "베드로전서", "벧후": "베드로후서", "요일": "요한일서", "요이": "요한이서",
    "요삼": "요한삼서", "유": "유다서", "계": "요한계시록"
}

def normalize_book_name(name: str) -> str:
    cleaned = name.strip()
    return BIBLE_ABBR_MAP.get(cleaned, cleaned)


class MultiPptRequest(BaseModel):
    lines: Optional[str] = None
    query_text: Optional[str] = None
    versions: List[str] = ["개역개정"]
    font_name: str = "경기천년제목 Bold"

# 다중 본문 합본 PPT 생성 엔드포인트
@app.api_route("/api/open-multi-ppt", methods=["GET", "POST"])
def create_multi_ppt_stream_or_open(payload: Optional[MultiPptRequest] = None, lines: Optional[str] = Query(None), versions: List[str] = Query(["개역개정"]), font_name: str = Query("경기천년제목 Bold")):
    import re
    # POST 바디 또는 GET 쿼리 파라미터 호환 처리
    input_text = ""
    if payload:
        input_text = payload.query_text or payload.lines or ""
    elif lines:
        input_text = lines
    use_versions = payload.versions if payload and payload.versions else versions
    use_font = payload.font_name if payload and payload.font_name else font_name

    if not input_text.strip():
        raise HTTPException(status_code=400, detail="본문 목록이 비어 있습니다.")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
        table_name = cursor.fetchone()[0]

        cursor.execute(f"PRAGMA table_info({table_name})")
        cols = [c[1] for c in cursor.fetchall()]
        has_ver_col = "version" in cols
        verse_col = "verse_num" if "verse_num" in cols else "verse"
        text_col = "text" if "text" in cols else "content"

        for line in input_text.strip().split("\n"):
            line = line.strip()
            if not line:
                continue

            # 구절 파싱 (예: "창 1:1-3", "창세기 1:1", "시 23:1")
            m = re.match(r"^([가-힣]+)\s*(\d+)[:장편]\s*(\d+)(?:[-~](\d+))?", line)
            if not m:
                continue

            b_name, ch, sv, ev = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)) if m.group(4) else int(m.group(3))
            canonical = normalize_book_name(b_name)
            unit = "편" if canonical == "시편" else "장"

            v_by_ver = {}
            for ver in use_versions:
                q = f"SELECT {verse_col} AS verse_num, {text_col} AS text FROM {table_name} WHERE book = ? AND chapter = ? AND {verse_col} BETWEEN ? AND ?"
                p = [canonical, ch, sv, ev]
                if has_ver_col:
                    q += " AND version = ?"
                    p.append(ver)
                q += f" ORDER BY {verse_col} ASC"
                cursor.execute(q, p)
                v_by_ver[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in cursor.fetchall()]

            if v_by_ver.get(use_versions[0]):
                sub_prs = build_presentation(v_by_ver, canonical, ch, unit, use_font, use_versions)
                # 생성된 슬라이드를 메인 프레젠테이션으로 병합 (blank slide 생성 로직 활용)
                for slide in sub_prs.slides:
                    # build_presentation이 생성한 슬라이드들을 순차 포함
                    pass

    # 전체 다중 프레젠테이션 직접 빌드 (안정적인 단일 파일 생성)
    # 실제로는 파싱된 모든 구절 세트를 순회하여 프레젠테이션에 슬라이드를 직접 추가
    full_prs = Presentation()
    full_prs.slide_width = Inches(13.333)
    full_prs.slide_height = Inches(7.5)
    blank_layout = full_prs.slide_layouts[6]
    bg_color = RGBColor(0, 32, 96)
    title_color = RGBColor(255, 192, 0)
    body_color = RGBColor(255, 255, 255)
    label_color = RGBColor(200, 225, 255)
    is_bold = not any(w in use_font for w in ["Regular", "Light", "Medium"])
    num_versions = len(use_versions)

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
        table_name = cursor.fetchone()[0]
        cursor.execute(f"PRAGMA table_info({table_name})")
        cols = [c[1] for c in cursor.fetchall()]
        has_ver_col = "version" in cols
        verse_col = "verse_num" if "verse_num" in cols else "verse"
        text_col = "text" if "text" in cols else "content"

        has_any_slide = False
        for line in input_text.strip().split("\n"):
            line = line.strip()
            if not line:
                continue
            m = re.match(r"^([가-힣]+)\s*(\d+)[:장편]\s*(\d+)(?:[-~](\d+))?", line)
            if not m:
                continue
            b_name, ch, sv, ev = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)) if m.group(4) else int(m.group(3))
            canonical = normalize_book_name(b_name)
            unit = "편" if canonical == "시편" else "장"

            v_by_ver = {}
            for ver in use_versions:
                q = f"SELECT {verse_col} AS verse_num, {text_col} AS text FROM {table_name} WHERE book = ? AND chapter = ? AND {verse_col} BETWEEN ? AND ?"
                p = [canonical, ch, sv, ev]
                if has_ver_col:
                    q += " AND version = ?"
                    p.append(ver)
                q += f" ORDER BY {verse_col} ASC"
                cursor.execute(q, p)
                v_by_ver[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in cursor.fetchall()]

            base_v = v_by_ver.get(use_versions[0], [])
            for v in base_v:
                has_any_slide = True
                v_num = v["verse_num"]
                if num_versions == 1:
                    chunks = split_verse_text(v["text"], max_chars=65)
                    for chunk in chunks:
                        slide = full_prs.slides.add_slide(blank_layout)
                        slide.background.fill.solid()
                        slide.background.fill.fore_color.rgb = bg_color
                        tb = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
                        p = tb.text_frame.paragraphs[0]
                        p.text = f"{canonical} {ch}{unit} {v_num}절"
                        p.alignment = PP_ALIGN.CENTER
                        set_korean_font(p, use_font, 50, title_color, is_bold=True)

                        cb = slide.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.8))
                        pc = cb.text_frame.paragraphs[0]
                        pc.text = chunk
                        pc.alignment = PP_ALIGN.CENTER
                        pc.line_spacing = 1.25
                        set_korean_font(pc, use_font, 52, body_color, is_bold=is_bold)
                else:
                    slide = full_prs.slides.add_slide(blank_layout)
                    slide.background.fill.solid()
                    slide.background.fill.fore_color.rgb = bg_color
                    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
                    p = tb.text_frame.paragraphs[0]
                    p.text = f"{canonical} {ch}{unit} {v_num}절"
                    p.alignment = PP_ALIGN.CENTER
                    set_korean_font(p, use_font, 46, title_color, is_bold=True)

                    left_margin = Inches(0.8)
                    available_width = 11.733
                    gap = 0.4
                    col_width_val = (available_width - (gap * (num_versions - 1))) / num_versions
                    col_width = Inches(col_width_val)
                    body_font_size = 34 if num_versions == 2 else 26

                    for idx, ver in enumerate(use_versions):
                        col_left = left_margin + Inches(idx * (col_width_val + gap))
                        ver_verse = next((item for item in v_by_ver.get(ver, []) if item["verse_num"] == v_num), None)
                        ver_text = ver_verse["text"] if ver_verse else ""

                        box = slide.shapes.add_textbox(col_left, Inches(1.5), col_width, Inches(5.3))
                        tf = box.text_frame
                        tf.word_wrap = True
                        p_label = tf.paragraphs[0]
                        p_label.text = f"[{ver}]"
                        set_korean_font(p_label, use_font, 26 if num_versions == 2 else 22, label_color, is_bold=True)

                        p_body = tf.add_paragraph()
                        p_body.text = ver_text
                        p_body.line_spacing = 1.25
                        set_korean_font(p_body, use_font, body_font_size, body_color, is_bold=is_bold)

    if not has_any_slide:
        raise HTTPException(status_code=404, detail="입력된 구절에서 유효한 성경 본문을 찾지 못했습니다.")

    filename = f"성경본문_합본_{int(time.time())}.pptx"
    encoded_filename = urllib.parse.quote(filename)

    if os.name == "nt":
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, filename)
        full_prs.save(file_path)
        try:
            os.startfile(file_path)
        except Exception:
            pass
        return {"status": "success", "mode": "local", "file_path": file_path}

    file_stream = io.BytesIO()
    full_prs.save(file_stream)
    file_stream.seek(0)
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
    )

# 기존 @app.post("/api/open-ppt") 대신 아래처럼 교체:
@app.api_route("/api/open-ppt", methods=["GET", "POST"])
def create_ppt_stream_or_open(
    book: str = Query(...),
    chapter: int = Query(..., ge=1),
    start_verse: Optional[int] = Query(None),
    end_verse: Optional[int] = Query(None),
    versions: List[str] = Query(["개역개정"]),
    font_name: str = Query("경기천년제목 Bold")
):
    try:
        s_verse = int(start_verse) if (start_verse and str(start_verse).strip().isdigit()) else None
        e_verse = int(end_verse) if (end_verse and str(end_verse).strip().isdigit()) else None

        unit = "편" if book == "시편" else "장"
        verses_by_version = {}

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('verses', 'bibles', 'bible')")
            table_name = cursor.fetchone()[0]

            cursor.execute(f"PRAGMA table_info({table_name})")
            cols = [c[1] for c in cursor.fetchall()]
            has_version_col = "version" in cols
            verse_col = "verse_num" if "verse_num" in cols else "verse"
            text_col = "text" if "text" in cols else "content"

            for ver in versions:
                query = f"SELECT {verse_col} AS verse_num, {text_col} AS text FROM {table_name} WHERE book = ? AND chapter = ?"
                params = [book, chapter]

                if has_version_col:
                    query += " AND version = ?"
                    params.append(ver)

                if s_verse and e_verse:
                    query += f" AND {verse_col} BETWEEN ? AND ?"
                    params.extend([s_verse, e_verse])
                elif s_verse:
                    query += f" AND {verse_col} = ?"
                    params.append(s_verse)

                query += f" ORDER BY {verse_col} ASC"
                cursor.execute(query, params)
                rows = cursor.fetchall()
                verses_by_version[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in rows]

        base_ver = versions[0]
        if not verses_by_version.get(base_ver):
            raise HTTPException(status_code=404, detail="성경 본문 데이터를 찾을 수 없습니다.")

        start_v = s_verse if s_verse else verses_by_version[base_ver][0]["verse_num"]
        end_v = e_verse if e_verse else verses_by_version[base_ver][-1]["verse_num"]

        prs = build_presentation(verses_by_version, book, chapter, unit, font_name, versions)

        ver_tag = "_".join(versions)
        filename = f"{book}_{chapter}{unit}_{start_v}_{end_v}절_{ver_tag}.pptx"
        encoded_filename = urllib.parse.quote(filename)

        # [분기 1] Windows 로컬 데스크톱 (.exe 환경) -> PPT 즉시 열기
        if os.name == "nt":
            temp_dir = tempfile.gettempdir()
            file_path = os.path.join(temp_dir, filename)
            prs.save(file_path)
            try:
                os.startfile(file_path)
            except Exception:
                pass
            return {"status": "success", "mode": "local", "file_path": file_path}

        # [분기 2] Render 웹 서버 (Linux 환경) -> 브라우저 스트림 다운로드
        file_stream = io.BytesIO()
        prs.save(file_stream)
        file_stream.seek(0)
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
        )

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"{str(e)}\n{traceback.format_exc()}")


# 정적 파일 마운트 (단일 exe 및 웹 배포 공용 방어 로직)
FRONTEND_DIST = os.path.join(BASE_DIR, "frontend", "dist")
ASSETS_DIR = os.path.join(FRONTEND_DIST, "assets")

if os.path.exists(ASSETS_DIR):
    app.mount("/assets", StaticFiles(directory=ASSETS_DIR), name="assets")

if os.path.exists(FRONTEND_DIST):
    @app.get("/{full_path:path}")
    def serve_frontend(full_path: str):
        file_path = os.path.join(FRONTEND_DIST, full_path)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return FileResponse(file_path)
        index_file = os.path.join(FRONTEND_DIST, "index.html")
        if os.path.exists(index_file):
            return FileResponse(index_file)
        return {"status": "running", "message": "BibleScript Pro Backend API is Active"}