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
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

        # ---------------- 1개 단일 번역본 (54pt 대형 뷰) ----------------
        if num_versions == 1:
            chunks = split_verse_text(v["text"], max_chars=65)
            for chunk in chunks:
                slide = prs.slides.add_slide(blank_layout)
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = bg_color

                # 상단 제목 (54pt)
                title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
                tf_title = title_box.text_frame
                tf_title.word_wrap = True
                p_title = tf_title.paragraphs[0]
                p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
                p_title.alignment = PP_ALIGN.CENTER
                set_korean_font(p_title, font_name, 54, title_color, is_bold=True)

                # 중앙 본문 (54pt, 1.2배 줄간격)
                content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.8))
                tf_content = content_box.text_frame
                tf_content.word_wrap = True
                p_content = tf_content.paragraphs[0]
                p_content.text = chunk
                p_content.alignment = PP_ALIGN.CENTER
                p_content.line_spacing = 1.2
                set_korean_font(p_content, font_name, 54, body_color, is_bold=is_bold)

        # ---------------- 2개 번역본 비교 (좌우 2단 분할 뷰) ----------------
        elif num_versions == 2:
            slide = prs.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = bg_color

            # 상단 제목 (48pt)
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(0.9))
            tf_title = title_box.text_frame
            tf_title.word_wrap = True
            p_title = tf_title.paragraphs[0]
            p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
            p_title.alignment = PP_ALIGN.CENTER
            set_korean_font(p_title, font_name, 48, title_color, is_bold=True)

            col_width = Inches(5.6)
            gap = Inches(0.533)
            left_margin = Inches(0.8)
            top_content = Inches(1.5)
            content_height = Inches(5.3)

            for idx, ver in enumerate(versions):
                col_left = left_margin + idx * (col_width + gap)
                ver_verse = next((item for item in verses_by_version.get(ver, []) if item["verse_num"] == verse_num), None)
                ver_text = ver_verse["text"] if ver_verse else ""

                box = slide.shapes.add_textbox(col_left, top_content, col_width, content_height)
                tf = box.text_frame
                tf.word_wrap = True

                # 번역본 라벨
                p_label = tf.paragraphs[0]
                p_label.text = f"[{ver}]"
                set_korean_font(p_label, font_name, 28, label_color, is_bold=True)

                # 본문 내용 (36pt, 1.25줄간격)
                p_body = tf.add_paragraph()
                p_body.text = ver_text
                p_body.line_spacing = 1.25
                set_korean_font(p_body, font_name, 36, body_color, is_bold=is_bold)

        # ---------------- 3개 이상 번역본 (다단 균등 분할 뷰) ----------------
        else:
            slide = prs.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = bg_color

            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(0.9))
            tf_title = title_box.text_frame
            tf_title.word_wrap = True
            p_title = tf_title.paragraphs[0]
            p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
            p_title.alignment = PP_ALIGN.CENTER
            set_korean_font(p_title, font_name, 44, title_color, is_bold=True)

            total_width = Inches(11.733)
            gap = Inches(0.4)
            col_width = (total_width - (gap * (num_versions - 1))) / num_versions
            left_margin = Inches(0.8)

            for idx, ver in enumerate(versions):
                col_left = left_margin + idx * (col_width + gap)
                ver_verse = next((item for item in verses_by_version.get(ver, []) if item["verse_num"] == verse_num), None)
                ver_text = ver_verse["text"] if ver_verse else ""

                box = slide.shapes.add_textbox(col_left, Inches(1.5), col_width, Inches(5.3))
                tf = box.text_frame
                tf.word_wrap = True

                p_label = tf.paragraphs[0]
                p_label.text = f"[{ver}]"
                set_korean_font(p_label, font_name, 24, label_color, is_bold=True)

                p_body = tf.add_paragraph()
                p_body.text = ver_text
                p_body.line_spacing = 1.2
                body_font_size = 28 if num_versions == 3 else 22
                set_korean_font(p_body, font_name, body_font_size, body_color, is_bold=is_bold)

    return prs


# ==================== 엔드포인트 라우팅 ====================

@app.get("/api/versions")
def get_versions():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT version FROM verses ORDER BY version ASC")
        rows = cursor.fetchall()
        versions = [r["version"] for r in rows]
        return {"versions": versions if versions else ["개역개정"]}


@app.get("/api/search-keyword")
def search_keyword(
    keyword: str = Query(..., min_length=2),
    version: str = Query("개역개정"),
    limit: int = Query(25)
):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT version, book, chapter, verse_num, text 
            FROM verses 
            WHERE version = ? AND text LIKE ? 
            LIMIT ?
            """,
            (version, f"%{keyword}%", limit)
        )
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
        cursor.execute(
            """
            UPDATE verses 
            SET text = ? 
            WHERE version = ? AND book = ? AND chapter = ? AND verse_num = ?
            """,
            (payload.text, payload.version, payload.book, payload.chapter, payload.verse_num)
        )
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


# 로컬 / 웹 자동 분기 파워포인트 생성 API
@app.post("/api/open-ppt")
def create_ppt_stream_or_open(
    book: str = Query(...),
    chapter: int = Query(..., ge=1),
    start_verse: Optional[int] = Query(None),
    end_verse: Optional[int] = Query(None),
    versions: List[str] = Query(["개역개정"]),
    font_name: str = Query("경기천년제목 Bold")
):
    try:
        # 빈 문자열("") 또는 None 방어 처리
        s_verse = int(start_verse) if (start_verse and str(start_verse).strip().isdigit()) else None
        e_verse = int(end_verse) if (end_verse and str(end_verse).strip().isdigit()) else None

        unit = "편" if book == "시편" else "장"
        verses_by_version = {}

    with get_db() as conn:
            cursor = conn.cursor()
            
            # version 컬럼 유무 확인 (구버전/신버전 DB 호환)
            cursor.execute("PRAGMA table_info(verses)")
            columns = [col[1] for col in cursor.fetchall()]
            has_version_col = "version" in columns

            for ver in versions:
                if has_version_col:
                    query = "SELECT verse_num, text FROM verses WHERE version = ? AND book = ? AND chapter = ?"
                    params = [ver, book, chapter]
                else:
                    query = "SELECT verse_num, text FROM verses WHERE book = ? AND chapter = ?"
                    params = [book, chapter]

                if s_verse and e_verse:
                    query += " AND verse_num BETWEEN ? AND ? ORDER BY verse_num"
                    params.extend([s_verse, e_verse])
                elif s_verse:
                    query += " AND verse_num = ?"
                    params.append(s_verse)
                else:
                    query += " ORDER BY verse_num"

                cursor.execute(query, params)
                rows = cursor.fetchall()
                verses_by_version[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in rows]

    base_ver = versions[0]
    if not verses_by_version.get(base_ver):
        raise HTTPException(status_code=404, detail="성경 본문 데이터를 찾을 수 없습니다.")

    start_v = start_verse if start_verse else verses_by_version[base_ver][0]["verse_num"]
    end_v = end_verse if end_verse else verses_by_version[base_ver][-1]["verse_num"]

    prs = build_presentation(verses_by_version, book, chapter, unit, font_name, versions)
    
    ver_tag = "_".join(versions)
    filename = f"{book}_{chapter}{unit}_{start_v}_{end_v}절_{ver_tag}.pptx"
    encoded_filename = urllib.parse.quote(filename)

    # [분기 1] Windows 로컬 데스크톱 (.exe 환경) ➔ 파워포인트 즉시 실행
    if os.name == "nt":
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, filename)
        prs.save(file_path)
        os.startfile(file_path)
        return {"status": "success", "mode": "local", "file_path": file_path}

    # [분기 2] Render 웹 서버 (Linux 환경) ➔ 브라우저 파일 다운로드
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

# 정적 파일 마운트 (단일 exe 및 웹 배포 공용)
FRONTEND_DIST = os.path.join(BASE_DIR, "frontend", "dist")
if os.path.exists(FRONTEND_DIST):
    app.mount("/assets", StaticFiles(directory=os.path.join(FRONTEND_DIST, "assets")), name="assets")

    @app.get("/{full_path:path}")
    def serve_frontend(full_path: str):
        file_path = os.path.join(FRONTEND_DIST, full_path)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse(os.path.join(FRONTEND_DIST, "index.html"))

