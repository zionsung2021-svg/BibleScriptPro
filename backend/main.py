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