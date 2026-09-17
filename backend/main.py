import io
import urllib.parse
import sqlite3
import os
from typing import Optional
from fastapi import FastAPI, Query, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

app = FastAPI(title="BibleScript API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = "bible.db"

# 66권 전체 약어 및 별칭 -> 표준 공식 명칭 매핑 테이블
BIBLE_NAME_MAP = {
    # 구약 (39권)
    "창": "창세기", "창세": "창세기", "창세기": "창세기",
    "출": "출애굽기", "출애": "출애굽기", "출애굽기": "출애굽기",
    "레": "레위기", "레위": "레위기", "레위기": "레위기",
    "민": "민수기", "민수": "민수기", "민수기": "민수기",
    "신": "신명기", "신명": "신명기", "신명기": "신명기",
    "수": "여호수아", "여호": "여호수아", "여호수아": "여호수아",
    "삿": "사사기", "사사": "사사기", "사사기": "사사기",
    "룻": "룻기", "룻기": "룻기",
    "삼상": "사무엘상", "사무엘상": "사무엘상",
    "삼하": "사무엘하", "사무엘하": "사무엘하",
    "왕상": "열왕기상", "열왕기상": "열왕기상",
    "왕하": "열왕기하", "열왕기하": "열왕기하",
    "대상": "역대상", "역대상": "역대상",
    "대하": "역대하", "역대하": "역대하",
    "스": "에스라", "에스라": "에스라",
    "느": "느헤미야", "느헤": "느헤미야", "느헤미야": "느헤미야",
    "에": "에스더", "에스더": "에스더",
    "욥": "욥기", "욥기": "욥기",
    "시": "시편", "시편": "시편",
    "잠": "잠언", "잠언": "잠언",
    "전": "전도서", "전도": "전도서", "전도서": "전도서",
    "아": "아가", "아가": "아가", "아가서": "아가",
    "사": "이사야", "이사": "이사야", "이사야": "이사야",
    "렘": "예레미야", "예레": "예레미야", "예레미야": "예레미야",
    "애": "예레미야애가", "애가": "예레미야애가", "예레미야애가": "예레미야애가",
    "겔": "에스겔", "에스겔": "에스겔",
    "단": "다니엘", "다니": "다니엘", "다니엘": "다니엘",
    "호": "호세아", "호세": "호세아", "호세아": "호세아",
    "욜": "요엘", "요엘": "요엘",
    "암": "아모스", "아모": "아모스", "아모스": "아모스",
    "옵": "오바댜", "오바": "오바댜", "오바댜": "오바댜",
    "욘": "요나", "요나": "요나",
    "미": "미가", "미가": "미가",
    "나": "나훔", "나훔": "나훔",
    "합": "하박국", "하박": "하박국", "하박국": "하박국",
    "습": "스바냐", "스바": "스바냐", "스바냐": "스바냐",
    "학": "학개", "학개": "학개",
    "슥": "스가랴", "스가": "스가랴", "스가랴": "스가랴",
    "말": "말라기", "말라": "말라기", "말라기": "말라기",

    # 신약 (27권)
    "마": "마태복음", "마태": "마태복음", "마태복음": "마태복음",
    "막": "마가복음", "마가": "마가복음", "마가복음": "마가복음",
    "눅": "누가복음", "누가": "누가복음", "누가복음": "누가복음",
    "요": "요한복음", "요한": "요한복음", "요한복음": "요한복음",
    "행": "사도행전", "사도": "사도행전", "사도행전": "사도행전",
    "롬": "로마서", "로마": "로마서", "로마서": "로마서",
    "고전": "고린도전서", "고린도전서": "고린도전서",
    "고후": "고린도후서", "고린도후서": "고린도후서",
    "갈": "갈라디아서", "갈라": "갈라디아서", "갈라디아서": "갈라디아서",
    "엡": "에베소서", "에베": "에베소서", "에베소서": "에베소서",
    "빌": "빌립보서", "빌립": "빌립보서", "빌립보서": "빌립보서",
    "골": "골로새서", "골로": "골로새서", "골로새서": "골로새서",
    "살전": "데살로니가전서", "데살로니가전서": "데살로니가전서",
    "살후": "데살로니가후서", "데살로니가후서": "데살로니가후서",
    "딤전": "디모데전서", "디모데전서": "디모데전서",
    "딤후": "디모데후서", "디모데후서": "디모데후서",
    "딛": "디도서", "디도": "디도서", "디도서": "디도서",
    "몬": "빌레몬서", "빌레": "빌레몬서", "빌레몬서": "빌레몬서",
    "히": "히브리서", "히브": "히브리서", "히브리서": "히브리서",
    "약": "야고보서", "야고": "야고보서", "야고보서": "야고보서",
    "벧전": "베드로전서", "베드로전서": "베드로전서",
    "벧후": "베드로후서", "베드로후서": "베드로후서",
    "요일": "요한일서", "요한일서": "요한일서", "요1": "요한일서",
    "요이": "요한이서", "요한이서": "요한이서", "요2": "요한이서",
    "요삼": "요한삼서", "요한삼서": "요한삼서", "요3": "요한삼서",
    "유": "유다서", "유다": "유다서", "유다서": "유다서",
    "계": "요한계시록", "계시": "요한계시록", "계시록": "요한계시록", "요한계시록": "요한계시록"
}

def resolve_book_name(input_name: str) -> str:
    cleaned = input_name.strip()
    return BIBLE_NAME_MAP.get(cleaned, cleaned)

def get_unit_name(book: str) -> str:
    return "편" if book in ["시편", "시"] else "장"

def set_korean_font(paragraph, font_name: str, size_pt: int, color_rgb: RGBColor, is_bold: bool = True):
    """한글(East Asian) 폰트가 파워포인트에서 깨지지 않도록 XML 레벨에서 설정"""
    paragraph.font.size = Pt(size_pt)
    paragraph.font.color.rgb = color_rgb
    paragraph.font.bold = is_bold
    paragraph.font.name = font_name

    # 동아시아 글꼴(ea) XML 속성 강제 설정
    for run in paragraph.runs:
        rPr = run._r.get_or_add_rPr()
        for ea in rPr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}ea'):
            rPr.remove(ea)
        ea_xml = parse_xml(f'<a:ea {nsdecls("a")} typeface="{font_name}"/>')
        rPr.append(ea_xml)

def get_bible_verses_from_db(book: str, chapter: int, start_v: Optional[int], end_v: Optional[int]):
    if not os.path.exists(DB_PATH):
        raise HTTPException(
            status_code=500,
            detail="데이터베이스(bible.db)가 존재하지 않습니다."
        )

    canonical_book = resolve_book_name(book)
    unit = get_unit_name(canonical_book)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    actual_start = start_v if (start_v is not None and start_v > 0) else 1

    if end_v is None or end_v <= 0:
        cur.execute(
            "SELECT MAX(verse) FROM bible WHERE book = ? AND chapter = ?",
            (canonical_book, chapter)
        )
        max_row = cur.fetchone()
        if not max_row or max_row[0] is None:
            conn.close()
            raise HTTPException(
                status_code=404,
                detail=f"'{book}'({canonical_book}) {chapter}{unit} 본문을 DB에서 찾을 수 없습니다."
            )
        actual_end = max_row[0]
    else:
        actual_end = end_v

    cur.execute(
        """
        SELECT verse, content FROM bible 
        WHERE book = ? AND chapter = ? AND verse BETWEEN ? AND ? 
        ORDER BY verse ASC
        """,
        (canonical_book, chapter, actual_start, actual_end)
    )
    rows = cur.fetchall()
    conn.close()

    if not rows:
        raise HTTPException(
            status_code=404,
            detail=f"{canonical_book} {chapter}{unit} {actual_start}~{actual_end}절 본문을 찾을 수 없습니다."
        )

    return [{"verse_num": r[0], "text": r[1]} for r in rows], canonical_book, unit, actual_start, actual_end

@app.get("/api/search")
def search_bible(
    book: str = Query(..., description="성경 이름 또는 약어 (예: 창, 시, 롬)"),
    chapter: int = Query(..., ge=1),
    start_verse: Optional[int] = Query(None),
    end_verse: Optional[int] = Query(None),
    version: str = Query("개역개정")
):
    verses, canonical_book, unit, start_v, end_v = get_bible_verses_from_db(book, chapter, start_verse, end_verse)
    formatted = "\n\n".join([f"[{canonical_book} {chapter}{unit} {v['verse_num']}절]\n{v['text']}" for v in verses])
    return {
        "book": canonical_book,
        "input_book": book,
        "chapter": chapter,
        "unit": unit,
        "start_verse": start_v,
        "end_verse": end_v,
        "total_verses": len(verses),
        "reference": f"{canonical_book} {chapter}{unit} {start_v}~{end_v}절",
        "verses": verses,
        "text": formatted
    }

@app.get("/api/generate-ppt")
def generate_ppt(
    book: str = Query(...),
    chapter: int = Query(..., ge=1),
    start_verse: Optional[int] = Query(None),
    end_verse: Optional[int] = Query(None),
    version: str = Query("개역개정"),
    font_name: str = Query("경기천년제목 Bold")
):
    verses, canonical_book, unit, start_v, end_v = get_bible_verses_from_db(book, chapter, start_verse, end_verse)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    bg_color = RGBColor(0, 32, 96)       # Navy Blue
    title_color = RGBColor(255, 192, 0)  # Gold Yellow
    body_color = RGBColor(255, 255, 255) # Pure White

    # 폰트명 및 가중치 매핑
    is_bold = True
    actual_font_family = font_name

    if font_name == "경기천년바탕 Bold":
        actual_font_family = "경기천년바탕 Bold"
        is_bold = True
    elif font_name == "경기천년바탕 Regular":
        actual_font_family = "경기천년바탕 Regular"
        is_bold = False
    elif font_name == "경기천년제목 Bold":
        actual_font_family = "경기천년제목 Bold"
        is_bold = True
    elif font_name == "경기천년제목 Light":
        actual_font_family = "경기천년제목 Light"
        is_bold = False
    elif font_name == "경기천년제목 Medium":
        actual_font_family = "경기천년제목 Medium"
        is_bold = False
    elif font_name == "경기천년제목V Bold":
        actual_font_family = "경기천년제목V Bold"
        is_bold = True

    for v in verses:
        slide = prs.slides.add_slide(blank_layout)
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        text_length = len(v["text"])

        # 1. 상단 중앙 장/절 제목
        title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(1.1))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = f"{canonical_book} {chapter}{unit} {v['verse_num']}절"
        p_title.alignment = PP_ALIGN.CENTER
        set_korean_font(p_title, actual_font_family, 46, title_color, is_bold=True)

        # 2. 중앙 본문
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.85), Inches(11.733), Inches(4.8))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True
        p_content = tf_content.paragraphs[0]
        p_content.text = v["text"]
        p_content.alignment = PP_ALIGN.CENTER

        if text_length > 100:
            font_size = 36
        elif text_length > 55:
            font_size = 42
        else:
            font_size = 48

        set_korean_font(p_content, actual_font_family, font_size, body_color, is_bold=is_bold)

    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)

    filename = f"{canonical_book}_{chapter}{unit}_{start_v}_{end_v}절.pptx"
    encoded_filename = urllib.parse.quote(filename)

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )

# --- [정적 파일 서빙] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 프로젝트 루트 디렉토리 (backend의 부모 폴더 또는 현재 폴더)
PARENT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))

search_dirs = [
    os.path.join(PARENT_DIR, "frontend", "dist"),
    os.path.join(BASE_DIR, "frontend", "dist"),
    os.path.join(os.getcwd(), "frontend", "dist"),
    os.path.join(os.getcwd(), "dist"),
    os.path.join(BASE_DIR, "dist")
]

dist_dir = None
for path in search_dirs:
    if os.path.exists(os.path.join(path, "index.html")):
        dist_dir = path
        break

if dist_dir:
    assets_dir = os.path.join(dist_dir, "assets")
    if os.path.exists(assets_dir):
        app.mount("/assets", StaticFiles(directory=assets_dir), name="assets")

    @app.get("/{full_path:path}")
    def serve_frontend(full_path: str):
        if full_path.startswith("api/") or full_path in ["docs", "openapi.json"]:
            raise HTTPException(status_code=404, detail="Not Found")
        
        target = os.path.join(dist_dir, full_path)
        if full_path and os.path.exists(target) and os.path.isfile(target):
            return FileResponse(target)
        return FileResponse(os.path.join(dist_dir, "index.html"))
else:
    @app.get("/")
    def fallback():
        return {
            "status": "online",
            "message": "Frontend build not found.",
            "searched": search_dirs
        }