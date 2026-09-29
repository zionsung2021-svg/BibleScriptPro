# run_local.py
import io
import os
import sys
import tempfile
import urllib.parse
import sqlite3
import threading
import time
import webbrowser
import re
from typing import Optional, List
import traceback

# 1. --noconsole 빌드 환경에서 uvicorn 로거 충돌(isatty) 방지
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

import uvicorn
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

# =====================================================================
# 핵심 1: 실행 환경별 (로컬/PyInstaller) 기준 경로 정의
# =====================================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =====================================================================
# 핵심 2: 진짜 성경 DB를 찾아내는 스마트 DB 연결 함수
# =====================================================================
def get_db():
    candidates = []

    # (1) PyInstaller 임시 풀림 경로
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
        candidates.extend([os.path.join(base, "backend", "bible.db"), os.path.join(base, "bible.db")])

    # (2) 현재 실행 파일(.exe) 기준 경로
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    candidates.extend([os.path.join(exe_dir, "backend", "bible.db"), os.path.join(exe_dir, "bible.db")])

    # (3) 윈도우 절대 경로 fallback (오타 방지 적용)
    desktop_onedrive = os.path.join(os.path.expanduser("~"), "OneDrive", "바탕 화면", "BibleScript")
    candidates.extend([
        os.path.join(desktop_onedrive, "backend", "bible.db"),
        os.path.join(desktop_onedrive, "bible.db"),
        r"C:\Users\admin\OneDrive\바탕 화면\BibleScript\backend\bible.db",
        r"C:\Users\admin\OneDrive\바탕 화면\BibleScript\bible.db"
    ])

    last_checked_paths = []
    for path in candidates:
        last_checked_paths.append(path)
        # 파일이 실제로 존재하고, 100KB 이상인 정상 DB만 연결 시도
        if os.path.exists(path) and os.path.getsize(path) > 100000:
            try:
                conn = sqlite3.connect(path)
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                # 테이블 존재 여부 검증
                cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
                if cursor.fetchone()[0] > 0:
                    return conn
                conn.close()
            except Exception:
                continue

    searched = "\n".join(list(dict.fromkeys(last_checked_paths))) # 중복 경로 제거 후 출력
    raise HTTPException(status_code=500, detail=f"정상적인 성경 DB를 찾을 수 없습니다.\n검색 경로:\n{searched}")

# =====================================================================
# 핵심 3: 테이블 이름을 자동 감지하는 헬퍼 함수 (verses가 아닐 경우 대비)
# =====================================================================
def get_target_table(cursor):
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    tables = [r[0] for r in cursor.fetchall()]
    if not tables:
        raise HTTPException(status_code=500, detail="DB 안에 사용할 수 있는 테이블이 없습니다.")
    return "verses" if "verses" in tables else tables[0]

# =====================================================================
# 슬라이드 생성 유틸리티 함수
# =====================================================================
def set_korean_font(paragraph, font_name: str, size_pt: int, color_rgb: RGBColor, is_bold: bool = False):
    for r in paragraph.runs:
        r.font.name = font_name
        r.font.size = Pt(size_pt)
        r.font.color.rgb = color_rgb
        r.font.bold = is_bold

# =====================================================================
# 성경 66권 약어 매핑 사전 및 정규화 함수
# =====================================================================
BIBLE_ABBR_MAP = {
    # 구약 (39권)
    "창": "창세기", "출": "출애굽기", "레": "레위기", "민": "민수기", "신": "신명기",
    "수": "여호수아", "삿": "사사기", "룻": "룻기", "삼상": "사무엘상", "삼하": "사무엘하",
    "왕상": "열왕기상", "왕하": "열왕기하", "대상": "역대상", "대하": "역대하", "스": "에스라",
    "느": "느헤미야", "에": "에스더", "욥": "욥기", "시": "시편", "잠": "잠언",
    "전": "전도서", "아": "아가", "사": "이사야", "렘": "예레미야", "애": "예레미야애가",
    "겔": "에스겔", "단": "다니엘", "호": "호세아", "욜": "요엘", "암": "아모스",
    "옵": "오바댜", "욘": "요나", "미": "미가", "나": "나훔", "합": "하박국",
    "습": "스바냐", "학": "학개", "슥": "스가랴", "말": "말라기",
    # 신약 (27권)
    "마": "마태복음", "막": "마가복음", "눅": "누가복음", "요": "요한복음", "행": "사도행전",
    "롬": "로마서", "고전": "고린도전서", "고후": "고린도후서", "갈": "갈라디아서", "엡": "에베소서",
    "빌": "빌립보서", "골": "골로새서", "살전": "데살로니가전서", "살후": "데살로니가후서",
    "딤전": "디모데전서", "딤후": "디모데후서", "딛": "디도서", "몬": "빌레몬서", "히": "히브리서",
    "약": "야고보서", "벧전": "베드로전서", "벧후": "베드로후서", "요일": "요한일서",
    "요이": "요한이서", "요삼": "요한삼서", "유": "유다서", "계": "요한계시록"
}

def normalize_book_name(input_book: str) -> str:
    cleaned = input_book.strip()
    return BIBLE_ABBR_MAP.get(cleaned, cleaned)


def parse_bible_queries(raw_text: str) -> List[dict]:
    items = re.split(r"[,;\n]+", raw_text)
    parsed = []
    pattern = re.compile(
        r"^\s*([가-힣]+)\s*(\d+)(?:[장:\s]+(\d+)(?:\s*[-~]\s*(\d+))?)?"
    )

    for item in items:
        item = item.strip()
        if not item:
            continue
        m = pattern.search(item)
        if m:
            b_name = normalize_book_name(m.group(1))
            ch = int(m.group(2))
            s_v = int(m.group(3)) if m.group(3) else None
            e_v = int(m.group(4)) if m.group(4) else (s_v if s_v else None)
            parsed.append(
                {
                    "book": b_name,
                    "chapter": ch,
                    "start_verse": s_v,
                    "end_verse": e_v,
                }
            )
    return parsed

def split_verse_text(text: str) -> List[str]:
    # 사용자가 오탈자 수정 등에서 수동으로 '//'를 넣은 경우에만 슬라이드 분할
    if "//" in text:
        parts = [part.strip() for part in text.split("//") if part.strip()]
        return parts if parts else [text]
    # 그 외에는 한 절 전체를 무조건 한 슬라이드에 표시
    return [text]


def calculate_dynamic_font_size(text: str, base_size: int = 48) -> int:
    """글자 수에 따라 한 슬라이드에 다 들어갈 수 있도록 글자 크기를 자동 조절 (기본 48pt)"""
    length = len(text)
    if length <= 60:
        return 48
    elif length <= 85:
        return 44
    elif length <= 115:
        return 40
    elif length <= 150:
        return 36
    else:
        return 32

def build_presentation(verses_by_version, canonical_book, chapter, unit, font_name, versions):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_color = RGBColor(0, 32, 96)
    title_color = RGBColor(255, 192, 0)
    body_color = RGBColor(255, 255, 255)
    label_color = RGBColor(200, 225, 255)
    is_bold = not any(w in font_name for w in ["Regular", "Light", "Medium"])
    num_versions = len(versions)
    base_ver = versions[0]
    base_verses = verses_by_version.get(base_ver, [])

    for v in base_verses:
        verse_num = v["verse_num"]

        # [1] 단일 번역본 (1개)
        if num_versions == 1:
            chunks = split_verse_text(v["text"])
            for chunk in chunks:
                slide = prs.slides.add_slide(blank_layout)
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = bg_color

                title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
                tf_title = title_box.text_frame
                p_title = tf_title.paragraphs[0]
                p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
                p_title.alignment = PP_ALIGN.CENTER
                set_korean_font(p_title, font_name, 50, title_color, is_bold=True)

                dynamic_size = calculate_dynamic_font_size(chunk, base_size=48)
                content_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(4.8))
                tf_content = content_box.text_frame
                tf_content.word_wrap = True
                p_content = tf_content.paragraphs[0]
                p_content.text = chunk
                p_content.alignment = PP_ALIGN.CENTER
                p_content.line_spacing = 1.3
                set_korean_font(p_content, font_name, dynamic_size, body_color, is_bold=is_bold)

        # [2] 다중 번역본 대조 (2개 또는 3개 이상 완벽 대응)
        else:
            slide = prs.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = bg_color

            # 상단 제목
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
            tf_title = title_box.text_frame
            p_title = tf_title.paragraphs[0]
            p_title.text = f"{canonical_book} {chapter}{unit} {verse_num}절"
            p_title.alignment = PP_ALIGN.CENTER
            set_korean_font(p_title, font_name, 46, title_color, is_bold=True)

            # 슬라이드 전체 가로 13.333 인치 기준 유동 계산
            left_margin = Inches(0.8)
            available_width = 11.733  # 13.333 - (0.8 * 2)
            gap = 0.4
            col_width_val = (available_width - (gap * (num_versions - 1))) / num_versions
            col_width = Inches(col_width_val)
            font_body_size = 34 if num_versions == 2 else 26  # 3개 대조 시 가독성을 위해 26pt로 자동 조절

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
                set_korean_font(p_body, font_name, font_body_size, body_color, is_bold=is_bold)

    return prs
# =====================================================================
# API 엔드포인트
# =====================================================================
@app.get("/api/versions")
def get_versions():
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            target_table = get_target_table(cursor)
            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [col[1] for col in cursor.fetchall()]
            
            if "version" in columns:
                cursor.execute(f"SELECT DISTINCT version FROM {target_table} ORDER BY version ASC")
                versions = [r["version"] for r in cursor.fetchall()]
                if versions:
                    return {"versions": versions}
        return {"versions": ["개역개정"]}
    except Exception:
        return {"versions": ["개역개정"]}

@app.get("/api/search-keyword")
def search_keyword(
    keyword: str = Query(..., min_length=2),
    version: str = Query("개역개정"),
    limit: int = Query(25)
):
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            target_table = get_target_table(cursor)

            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [col[1] for col in cursor.fetchall()]

            v_col = "verse_num" if "verse_num" in columns else ("verse" if "verse" in columns else columns[3])
            t_col = "text" if "text" in columns else ("content" if "content" in columns else columns[4])

            if "version" in columns:
                cursor.execute(
                    f"SELECT version, book, chapter, {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE version = ? AND {t_col} LIKE ? LIMIT ?",
                    (version, f"%{keyword}%", limit)
                )
            else:
                cursor.execute(
                    f"SELECT '{version}' as version, book, chapter, {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE {t_col} LIKE ? LIMIT ?",
                    (f"%{keyword}%", limit)
                )
            return [dict(r) for r in cursor.fetchall()]
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"검색 처리 오류: {str(e)}")

class VerseUpdateRequest(BaseModel):
    version: str = "개역개정"
    book: str
    chapter: int
    verse_num: int
    text: str

@app.patch("/api/verse")
def update_verse_db(payload: VerseUpdateRequest):
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            target_table = get_target_table(cursor)

            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [col[1] for col in cursor.fetchall()]

            if "version" in columns:
                cursor.execute(
                    f"UPDATE {target_table} SET text = ? WHERE version = ? AND book = ? AND chapter = ? AND verse_num = ?",
                    (payload.text, payload.version, payload.book, payload.chapter, payload.verse_num)
                )
            else:
                cursor.execute(
                    f"UPDATE {target_table} SET text = ? WHERE book = ? AND chapter = ? AND verse_num = ?",
                    (payload.text, payload.book, payload.chapter, payload.verse_num)
                )
            conn.commit()
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"저장 오류: {str(e)}")


@app.post("/api/shutdown")
def shutdown_app():
    def kill_process():
        time.sleep(0.5)
        os._exit(0)
    threading.Thread(target=kill_process, daemon=True).start()
    return {"status": "success"}

@app.post("/api/open-ppt")
def create_ppt_stream_or_open(
    book: str = Query(...),
    chapter: int = Query(..., ge=1),
    start_verse: Optional[str] = Query(None),
    end_verse: Optional[str] = Query(None),
    versions: List[str] = Query(["개역개정"]),
    font_name: str = Query("경기천년제목 Medium")  # 1) 기본 글꼴 변경
):
    try:
        # 2) 약어를 정식 명칭으로 자동 변환 (예: '계' -> '요한계시록')
        canonical_book = normalize_book_name(book)

        s_verse = int(start_verse) if (start_verse and str(start_verse).strip().isdigit()) else None
        e_verse = int(end_verse) if (end_verse and str(end_verse).strip().isdigit()) else None

        unit = "편" if canonical_book == "시편" else "장"
        verses_by_version = {}

        with get_db() as conn:
            cursor = conn.cursor()
            target_table = get_target_table(cursor)

            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [col[1] for col in cursor.fetchall()]

            v_col = "verse_num" if "verse_num" in columns else ("verse" if "verse" in columns else columns[3])
            t_col = "text" if "text" in columns else ("content" if "content" in columns else columns[4])
            has_version_col = "version" in columns

            for ver in versions:
                # 3) book 파라미터 대신 변환된 canonical_book 사용
                if has_version_col:
                    query = f"SELECT {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE version = ? AND book = ? AND chapter = ?"
                    params = [ver, canonical_book, chapter]
                else:
                    query = f"SELECT {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE book = ? AND chapter = ?"
                    params = [canonical_book, chapter]

                if s_verse and e_verse:
                    query += f" AND {v_col} BETWEEN ? AND ? ORDER BY {v_col}"
                    params.extend([s_verse, e_verse])
                elif s_verse:
                    query += f" AND {v_col} = ?"
                    params.append(s_verse)
                else:
                    query += f" ORDER BY {v_col}"

                cursor.execute(query, params)
                verses_by_version[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in cursor.fetchall()]

        base_ver = versions[0]
        if not verses_by_version.get(base_ver):
            raise HTTPException(status_code=404, detail=f"[{canonical_book} {chapter}{unit}] 해당 구절 데이터를 찾을 수 없습니다.")

        start_v = s_verse if s_verse else verses_by_version[base_ver][0]["verse_num"]
        end_v = e_verse if e_verse else verses_by_version[base_ver][-1]["verse_num"]

        prs = build_presentation(verses_by_version, canonical_book, chapter, unit, font_name, versions)
        ver_tag = "_".join(versions)

        # 중복 열림 방지를 위해 타임스탬프 추가
        time_tag = time.strftime("%H%M%S")
        filename = f"{canonical_book}_{chapter}{unit}_{start_v}_{end_v}절_{ver_tag}.pptx"
        encoded_filename = urllib.parse.quote(filename)

        if os.name == "nt":
            temp_dir = tempfile.gettempdir()
            file_path = os.path.join(temp_dir, filename)
            prs.save(file_path)
            os.startfile(file_path)
            return {"status": "success", "mode": "local", "file_path": file_path}

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
        raise HTTPException(status_code=500, detail=f"PPT 생성 중 서버 오류: {str(e)}\n{traceback.format_exc()}")

from pydantic import BaseModel

class MultiPptRequest(BaseModel):
    query_text: str
    versions: List[str] = ["개역개정"]
    font_name: str = "경기천년제목 Medium"

@app.post("/api/open-multi-ppt")
def create_multi_ppt(payload: MultiPptRequest):
    try:
        targets = parse_bible_queries(payload.query_text)
        if not targets:
            raise HTTPException(status_code=400, detail="유효한 성경 구절 형식이 아닙니다. (예: 창 1:1-3, 요 3:16)")

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        bg_color = RGBColor(0, 32, 96)
        title_color = RGBColor(255, 192, 0)
        body_color = RGBColor(255, 255, 255)
        label_color = RGBColor(200, 225, 255)
        is_bold = not any(w in payload.font_name for w in ["Regular", "Light", "Medium"])
        num_versions = len(payload.versions)

        with get_db() as conn:
            cursor = conn.cursor()
            target_table = get_target_table(cursor)
            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [col[1] for col in cursor.fetchall()]

            v_col = "verse_num" if "verse_num" in columns else ("verse" if "verse" in columns else columns[3])
            t_col = "text" if "text" in columns else ("content" if "content" in columns else columns[4])
            has_version_col = "version" in columns

            for item in targets:
                book = item["book"]
                chapter = item["chapter"]
                s_verse = item["start_verse"]
                e_verse = item["end_verse"]
                unit = "편" if book == "시편" else "장"
                verses_by_version = {}

                for ver in payload.versions:
                    if has_version_col:
                        query = f"SELECT {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE version = ? AND book = ? AND chapter = ?"
                        params = [ver, book, chapter]
                    else:
                        query = f"SELECT {v_col} AS verse_num, {t_col} AS text FROM {target_table} WHERE book = ? AND chapter = ?"
                        params = [book, chapter]

                    if s_verse and e_verse:
                        query += f" AND {v_col} BETWEEN ? AND ? ORDER BY {v_col}"
                        params.extend([s_verse, e_verse])
                    elif s_verse:
                        query += f" AND {v_col} = ?"
                        params.append(s_verse)
                    else:
                        query += f" ORDER BY {v_col}"

                    cursor.execute(query, params)
                    verses_by_version[ver] = [{"verse_num": r["verse_num"], "text": r["text"]} for r in cursor.fetchall()]

                base_ver = payload.versions[0]
                base_verses = verses_by_version.get(base_ver, [])
                if not base_verses:
                    continue

                for v in base_verses:
                    verse_num = v["verse_num"]
                    if num_versions == 1:
                        chunks = split_verse_text(v["text"])
                        for chunk in chunks:
                            slide = prs.slides.add_slide(blank_layout)
                            slide.background.fill.solid()
                            slide.background.fill.fore_color.rgb = bg_color

                            # 상단 제목
                            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
                            tf_title = title_box.text_frame
                            p_title = tf_title.paragraphs[0]
                            p_title.text = f"{book} {chapter}{unit} {verse_num}절"
                            p_title.alignment = PP_ALIGN.CENTER
                            set_korean_font(p_title, payload.font_name, 50, title_color, is_bold=True)

                            # 글자 수에 따른 48pt 자동 축소 적용
                            dynamic_size = calculate_dynamic_font_size(chunk, base_size=48)

                            content_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(4.8))
                            tf_content = content_box.text_frame
                            tf_content.word_wrap = True
                            p_content = tf_content.paragraphs[0]
                            p_content.text = chunk
                            p_content.alignment = PP_ALIGN.CENTER
                            p_content.line_spacing = 1.3
                            set_korean_font(p_content, payload.font_name, dynamic_size, body_color, is_bold=is_bold)
                    else:
                        slide = prs.slides.add_slide(blank_layout)
                        slide.background.fill.solid()
                        slide.background.fill.fore_color.rgb = bg_color

                        title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(11.333), Inches(0.9))
                        tf_title = title_box.text_frame
                        p_title = tf_title.paragraphs[0]
                        p_title.text = f"{book} {chapter}{unit} {verse_num}절"
                        p_title.alignment = PP_ALIGN.CENTER
                        set_korean_font(p_title, payload.font_name, 48, title_color, is_bold=True)

                        col_width = Inches(5.6)
                        gap = Inches(0.533)
                        left_margin = Inches(0.8)
                        for idx, ver in enumerate(payload.versions):
                            col_left = left_margin + idx * (col_width + gap)
                            ver_verse = next((item for item in verses_by_version.get(ver, []) if item["verse_num"] == verse_num), None)
                            ver_text = ver_verse["text"] if ver_verse else ""

                            box = slide.shapes.add_textbox(col_left, Inches(1.5), col_width, Inches(5.3))
                            tf = box.text_frame
                            tf.word_wrap = True

                            p_label = tf.paragraphs[0]
                            p_label.text = f"[{ver}]"
                            set_korean_font(p_label, payload.font_name, 28, label_color, is_bold=True)

                            p_body = tf.add_paragraph()
                            p_body.text = ver_text
                            p_body.line_spacing = 1.25
                            set_korean_font(p_body, payload.font_name, 36, body_color, is_bold=is_bold)

        if len(prs.slides) == 0:
            raise HTTPException(status_code=404, detail="지정한 구절 데이터를 DB에서 찾을 수 없습니다.")

        temp_dir = tempfile.gettempdir()
        time_tag = time.strftime("%H%M%S")
        file_path = os.path.join(temp_dir, f"연속성경구절_{time_tag}.pptx")
        prs.save(file_path)
        if os.name == "nt":
            os.startfile(file_path)
        return {"status": "success", "file_path": file_path}

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"다중 PPT 생성 오류: {str(e)}")

# =====================================================================
# 정적 파일 마운트 (이 블록은 반드시 파일의 맨 마지막에 위치해야 합니다)
# =====================================================================
FRONTEND_DIST = os.path.join(BASE_DIR, "frontend", "dist")
if os.path.exists(FRONTEND_DIST):
    app.mount("/assets", StaticFiles(directory=os.path.join(FRONTEND_DIST, "assets")), name="assets")

    @app.get("/{full_path:path}")
    def serve_frontend(full_path: str):
        file_path = os.path.join(FRONTEND_DIST, full_path)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse(os.path.join(FRONTEND_DIST, "index.html"))

# =====================================================================
# 네이티브 윈도우 창(WebView) 초고속 구동 및 즉각 종료 핸들러
# =====================================================================
import os
import socket
import threading
import time
import webview

# 전역 main_window 변수 선언 (백엔드가 pywebview 창을 직접 닫기 위함)
main_window = None

@app.post("/api/shutdown")
def shutdown_app():
    def force_exit():
        time.sleep(0.1)
        # 1. 화면의 pywebview 창 즉시 파괴
        if main_window:
            try:
                main_window.destroy()
            except Exception:
                pass
        # 2. 파이썬 백그라운드 프로세스 강제 종료
        os._exit(0)

    threading.Thread(target=force_exit, daemon=True).start()
    return {"status": "shutting_down"}


def start_server():
    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8000,
        log_config=None,
        access_log=False,
        log_level="critical",
    )


def wait_for_server(port=8000, timeout=5.0):
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.05):
                return True
        except (OSError, ConnectionRefusedError):
            time.sleep(0.03)
    return False


if __name__ == "__main__":
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    wait_for_server(8000)

    # 전역 변수에 창 객체 등록
    main_window = webview.create_window(
        title="BibleScript Pro",
        url="http://127.0.0.1:8000",
        width=1280,
        height=850,
        min_size=(960, 640),
        resizable=True,
        confirm_close=False,
    )

    def on_closed():
        os._exit(0)

    main_window.events.closed += on_closed
    webview.start()
    os._exit(0)


