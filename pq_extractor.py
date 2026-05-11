# -*- coding: utf-8 -*-
"""
배전감리 PQ 심사 데이터 추출 시스템
한국전력공사 배전건설부 - 감리용역 적격심사 PQ 자동 추출

사용법: python pq_extractor.py <업체PDF경로> [입찰공고일] [총예정전기공사비구간]
예시: python pq_extractor.py "증빙/(주)성문기술단.pdf" 2026-03-09 10억미만
"""

import sys
import os
import re
import json
from datetime import datetime, timedelta
from pathlib import Path

# Windows 한글 출력 설정
sys.stdout.reconfigure(encoding='utf-8')

# PDF/OCR
import fitz  # PyMuPDF
import easyocr

# Excel
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

###############################################################################
# 설정
###############################################################################
BIDDING_DATE = "2026-03-09"  # 입찰공고일 (기본값)
COST_TIER = "10억미만"  # 총 예정전기공사비 구간

# 점수 기준표 (10억원 미만 구간 기준)
# 책임감리원 전기분야 경력 점수표
CHIEF_ELEC_CAREER_SCORE = {
    # (등급, 경력연수): 점수
    # 10억원 미만 구간
    "특급": [(1, 4.0), (2, 5.5), (3, 7.0), (4, 8.5), (0, 10.0)],  # 1미만→4, 2이상→5.5, ...
    "고급": [(2, 4.0), (3, 5.5), (4, 7.0), (5, 8.5), (0, 10.0)],
    "중급": [(3, 4.0), (4, 5.5), (5, 7.0), (6, 8.5), (0, 10.0)],
}

# 참여분야 경력 점수표 (나-1: 설계,시공,감리 등 전체)
CHIEF_FIELD_CAREER_SCORE_1 = {
    "특급": [(1, 4.0), (2, 5.5), (3, 7.0), (4, 8.5), (0, 10.0)],
    "고급": [(2, 4.0), (3, 5.5), (4, 7.0), (5, 8.5), (0, 10.0)],
    "중급": [(3, 4.0), (4, 5.5), (5, 7.0), (6, 8.5), (0, 10.0)],
}

# 참여분야 경력 점수표 (나-2: 감독, 감리만)
CHIEF_FIELD_CAREER_SCORE_2 = {
    "특급": [(1, 1.0), (2, 2.0), (3, 3.0), (4, 4.0), (0, 5.0)],
    "고급": [(2, 1.0), (3, 2.0), (4, 3.0), (5, 4.0), (0, 5.0)],
    "중급": [(3, 1.0), (4, 2.0), (5, 3.0), (6, 4.0), (0, 5.0)],
}

# 자본비율 점수표
CAPITAL_RATIO_SCORE = [
    (100, 1.0),   # 100% 이상
    (75, 0.85),   # 75% 이상
    (50, 0.7),    # 50% 이상
    (25, 0.55),   # 25% 이상
    (0, 0.4),     # 25% 미만
]

# 유동비율 점수표
CURRENT_RATIO_SCORE = [
    (100, 2.0),   # 100% 이상
    (75, 1.7),    # 75% 이상
    (50, 1.4),    # 50% 이상
    (25, 1.1),    # 25% 이상
    (0, 0.8),     # 25% 미만
]

# 비상주 중첩도 점수표
NONRESIDENT_OVERLAP_SCORE = {
    0: 4.0, 1: 4.0, 2: 4.0, 3: 4.0, 4: 4.0, 5: 4.0,
    6: 3.2, 7: 2.4, 8: 1.6, 9: 0.8,
}

# 가점 - 책임감리원 자격보유
CERT_BONUS = {
    "기술사": 1.0,
    "기능장": 0.8,
    "기사": 0.7,
    "산업기사": 0.5,
}


###############################################################################
# OCR 엔진 초기화 + 파일 캐시 (JSON 영구 저장)
###############################################################################
import gc
import hashlib

_reader = None
_ocr_memory_cache = {}  # 메모리 캐시 (세션 내 중복 방지)
_current_pdf_path = None
_cache_dir = None


def get_reader():
    global _reader
    if _reader is None:
        print("[INFO] OCR 엔진 초기화 중 (최초 1회)...")
        try:
            _reader = easyocr.Reader(['ko', 'en'], gpu=False)
            print("[INFO] OCR 엔진 준비 완료")
        except Exception as e:
            print(f"[WARN] OCR 엔진 초기화 실패 ({type(e).__name__}): {e}")
            _reader = False  # 실패 표시 (None과 구분)
    if _reader is False:
        return None
    return _reader


def set_current_pdf(pdf_path):
    """현재 처리중인 PDF 설정 → 파일 캐시 디렉토리 결정"""
    global _current_pdf_path, _cache_dir
    _current_pdf_path = pdf_path
    _cache_dir = None


def _get_cache_dir():
    """PDF별 캐시 디렉토리 (증빙/.ocr_cache/<파일명>_<크기>/)"""
    global _cache_dir
    if _cache_dir is not None:
        return _cache_dir
    if not _current_pdf_path:
        return None
    base_dir = os.path.join(os.path.dirname(_current_pdf_path), ".ocr_cache")
    pdf_name = os.path.splitext(os.path.basename(_current_pdf_path))[0]
    pdf_size = os.path.getsize(_current_pdf_path)
    _cache_dir = os.path.join(base_dir, f"{pdf_name}_{pdf_size}")
    os.makedirs(_cache_dir, exist_ok=True)
    return _cache_dir


def _load_page_cache(page_num):
    """파일 캐시에서 페이지 OCR 결과 로드"""
    cache_dir = _get_cache_dir()
    if not cache_dir:
        return None
    cache_file = os.path.join(cache_dir, f"page_{page_num:04d}.json")
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return [(item["text"], item["conf"], item["bbox"]) for item in data]
        except (json.JSONDecodeError, KeyError, TypeError):
            return None
    return None


def _to_native(obj):
    """numpy 타입을 Python 네이티브 타입으로 변환"""
    if isinstance(obj, (list, tuple)):
        return [_to_native(x) for x in obj]
    try:
        return obj.item()  # numpy scalar → Python int/float
    except (AttributeError, ValueError):
        return obj


def _save_page_cache(page_num, parsed):
    """OCR 결과를 JSON 파일로 영구 저장"""
    cache_dir = _get_cache_dir()
    if not cache_dir:
        return
    cache_file = os.path.join(cache_dir, f"page_{page_num:04d}.json")
    data = [{"text": t, "conf": float(c), "bbox": _to_native(b)} for (t, c, b) in parsed]
    try:
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f"  [WARN] 캐시 저장 실패 (p{page_num+1}): {e}")


def ocr_page(doc, page_num, zoom=1.0):
    """PDF 페이지를 OCR하여 텍스트 리스트 반환 (메모리+파일 캐시)"""
    # 1) 메모리 캐시
    cache_key = (id(doc), page_num)
    if cache_key in _ocr_memory_cache:
        return _ocr_memory_cache[cache_key]

    # 2) 파일(JSON) 캐시 → 있으면 OCR 건너뛰기
    cached = _load_page_cache(page_num)
    if cached is not None:
        _ocr_memory_cache[cache_key] = cached
        return cached

    # 3) OCR 실행
    reader = get_reader()
    page = doc[page_num]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes('png')
    del pix
    gc.collect()

    try:
        results = reader.readtext(img_data, batch_size=1)
    except Exception as e:
        # numpy ArrayMemoryError, MemoryError, RuntimeError 등 모두 캐치
        print(f"  [WARN] 페이지 {page_num+1} OCR 실패({type(e).__name__}), 저해상도 재시도...")
        del img_data
        gc.collect()
        # 0.8배로 축소하여 재시도
        mat = fitz.Matrix(0.8, 0.8)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes('png')
        del pix
        gc.collect()
        try:
            results = reader.readtext(img_data, batch_size=1)
        except Exception as e2:
            print(f"  [ERROR] 페이지 {page_num+1} OCR 완전 실패: {type(e2).__name__}")
            del img_data
            gc.collect()
            return []

    parsed = [(text, conf, bbox) for (bbox, text, conf) in results]

    # 캐시 저장 (메모리 + 파일)
    _ocr_memory_cache[cache_key] = parsed
    _save_page_cache(page_num, parsed)

    del img_data, results
    gc.collect()
    return parsed


def ocr_page_text(doc, page_num, zoom=1.0):
    """페이지의 전체 텍스트를 하나의 문자열로 반환"""
    results = ocr_page(doc, page_num, zoom)
    return ' '.join([text for (text, conf, bbox) in results])


###############################################################################
# 텍스트 추출: PyMuPDF 우선 → OCR 폴백
###############################################################################

def get_page_text(doc, page_num):
    """페이지 텍스트 추출 (PyMuPDF 내장 텍스트 우선, 부족하면 OCR 폴백)"""
    if page_num >= doc.page_count:
        return ""
    page = doc[page_num]
    text = page.get_text()
    if text and len(text.strip()) > 20:
        return text
    # 텍스트가 없거나 부족하면 OCR 폴백
    return ocr_page_text(doc, page_num)


def get_page_lines(doc, page_num):
    """페이지의 줄 단위 텍스트 리스트 반환"""
    text = get_page_text(doc, page_num)
    return [l.strip() for l in text.split('\n') if l.strip()]


###############################################################################
# 데이터 추출 함수 (PyMuPDF 텍스트 우선 방식)
###############################################################################

def smart_find_page(doc, keywords, search_range, label=""):
    """핵심 페이지를 빠르게 찾기 - PyMuPDF 내장 텍스트만 사용 (OCR 없음, 크래시 방지)"""
    for page_num in search_range:
        if page_num >= doc.page_count:
            break
        text = doc[page_num].get_text()  # OCR 폴백 없이 내장 텍스트만
        if any(kw in text for kw in keywords):
            if label:
                print(f"    [{label}] 페이지 {page_num+1}에서 발견")
            return page_num
    return -1


def extract_company_name(doc):
    """업체명 추출 (PyMuPDF 텍스트 우선)"""
    for page_num in range(min(5, doc.page_count)):
        lines = get_page_lines(doc, page_num)
        for line in lines:
            # "(주)XXX" 또는 "㈜XXX" 패턴
            m = re.search(r'[(\(]주[)\)]?\s*[가-힣]+|㈜\s*[가-힣]+|주식회사\s*[가-힣]+', line)
            if m:
                return m.group(0).strip()
    return "미확인"


def extract_summary_table(doc):
    """종합득점표(양식2-3) 추출 - 대괄호 배점 순서 기반"""
    result = {
        "참여감리원": 0,    # [50]
        "유사용역": 0,      # 1st [10]
        "신용도": 0,        # 2nd [10]
        "기술개발": 0,      # 3rd [10]
        "업무중첩도": 0,    # 4th [10]
        "교체빈도": 0,      # 1st [5]
        "작업계획": 0,      # 2nd [5]
        "가감점": 0,        # [  ]
        "총점": 0,
        "page": -1,
    }

    page_num = smart_find_page(doc, ['종합득점표', '양식2-3'],
                                range(10, min(25, doc.page_count)), "종합득점표")
    if page_num < 0:
        page_num = smart_find_page(doc, ['종합득점표'],
                                    range(0, min(30, doc.page_count)), "종합득점표(확장)")

    if page_num >= 0:
        result["page"] = page_num + 1
        lines = get_page_lines(doc, page_num)

        # 대괄호 [XX] 패턴으로 배점 순서 추출
        # 종합득점표 구조: [50] → [10] → [10] → [10] → [10] → [5] → [5] → [ ] → 100(총배점)
        # 각 대괄호 다음 줄의 숫자가 업체제출 점수
        bracket_scores = []  # (배점, 업체제출점수) 쌍
        for i, line in enumerate(lines):
            m = re.match(r'^\s*\[(\d+|[\s]*)\]\s*$', line.strip())
            if m:
                bracket_val = m.group(1).strip()
                # 다음 줄에서 업체제출 점수 찾기
                for j in range(i+1, min(i+3, len(lines))):
                    score_m = re.match(r'^\s*(\d+\.?\d*)\s*$', lines[j].strip())
                    if score_m:
                        score = float(score_m.group(1))
                        bracket_scores.append((bracket_val, score))
                        break

        # 순서대로 매핑: [50], [10]x4, [5]x2, [ ]
        # 항목 이름 순서 (종합득점표 표준)
        category_order = ["참여감리원", "유사용역", "신용도", "기술개발",
                          "업무중첩도", "교체빈도", "작업계획", "가감점"]
        for idx, (bv, score) in enumerate(bracket_scores):
            if idx < len(category_order):
                result[category_order[idx]] = score

        # 총점: "100" 다음 줄
        for i, line in enumerate(lines):
            if line.strip() == '100':
                for j in range(i+1, min(i+3, len(lines))):
                    m = re.match(r'^\s*(\d+\.?\d+)\s*$', lines[j].strip())
                    if m:
                        val = float(m.group(1))
                        if 50 <= val <= 120:
                            result["총점"] = val
                            break

    return result


def extract_personnel(doc):
    """참여감리원 자격사항(양식2-4) 추출 - PyMuPDF 텍스트 사용"""
    personnel = {
        "책임감리원": {"성명": "", "등급": "", "자격증": "", "생년월일": ""},
        "보조감리원": [],
        "비상주감리원": {"성명": "", "등급": "", "자격증": "", "생년월일": ""},
        "page": -1,
    }

    # 양식2-4 페이지 찾기 (종합득점표 양식2-3과 구분 필요)
    page_num = smart_find_page(doc, ['양식2-4'],
                                range(12, min(25, doc.page_count)), "참여감리원")
    if page_num < 0:
        page_num = smart_find_page(doc, ['참여감리원 자격사항', '자격등급'],
                                    range(12, min(30, doc.page_count)), "참여감리원(확장)")

    if page_num >= 0:
        personnel["page"] = page_num + 1
        lines = get_page_lines(doc, page_num)

        grades = ['특급', '고급', '중급', '초급']
        certs = ['기술사', '기능장', '기사', '산업기사']
        # 이름 필터 목록
        name_filter = {'책임감리원', '보조감리원', '비상주', '감리원', '전기', '토목',
                       '특급', '고급', '중급', '초급', '소계', '합계', '배점', '평점',
                       '자격사항', '자격내용', '자격등급', '참여감리원', '학력내용',
                       '소속회사', '생년월일', '비상주감리원', '전기분야', '참여분야',
                       '교육수료', '종료일자', '최종학력', '취득일', '자격명',
                       '성문기술단', '졸업일'}

        # 1단계: 모든 이름 후보 수집 (순서대로)
        all_names = []
        for i, line in enumerate(lines):
            name_match = re.search(r'(?<![가-힣])([가-힣]{2,4})(?![가-힣])', line)
            if name_match:
                name = name_match.group(1)
                if name not in name_filter and not any(kw in name for kw in
                    ['감리', '기술', '분야', '비고', '양식', '교육', '재직', '증명',
                     '사본', '추가', '확인', '발급', '등본', '경우', '첨부', '상주',
                     '법인', '협회', '시행', '건설', '경력', '학위', '자격']):
                    all_names.append((i, name))

        # 2단계: 섹션 마커 위치 파악
        chief_line = -1
        asst_line = -1
        nonres_line = -1
        for i, line in enumerate(lines):
            if '책임감리원' in line and chief_line < 0:
                chief_line = i
            if ('보조감리원' in line or '보조 감리원' in line) and asst_line < 0:
                asst_line = i
            if '비상주' in line and nonres_line < 0:
                nonres_line = i

        # 3단계: 이름 할당 (테이블 구조 - 책임/비상주가 같은 행에 나란히)
        if len(all_names) >= 1:
            personnel["책임감리원"]["성명"] = all_names[0][1]
        if len(all_names) >= 2:
            if asst_line < 0 or all_names[1][0] < asst_line:
                personnel["비상주감리원"]["성명"] = all_names[1][1]
            else:
                personnel["보조감리원"].append({"성명": all_names[1][1], "등급": "", "경력일수": ""})
        for idx, (line_num, name) in enumerate(all_names[2:], 2):
            if asst_line >= 0 and line_num >= asst_line:
                personnel["보조감리원"].append({"성명": name, "등급": "", "경력일수": ""})

        # 4단계: 등급 - 순서 기반 (테이블 열 순서: 1=책임, 2=비상주, 3=보조)
        all_grades = []
        for i, line in enumerate(lines):
            for grade in grades:
                if grade == line.strip():  # 등급만 있는 줄
                    all_grades.append((i, grade))
        if len(all_grades) >= 1:
            personnel["책임감리원"]["등급"] = all_grades[0][1]
        if len(all_grades) >= 2:
            personnel["비상주감리원"]["등급"] = all_grades[1][1]
        if len(all_grades) >= 3 and personnel["보조감리원"]:
            personnel["보조감리원"][0]["등급"] = all_grades[2][1]

        # 5단계: 자격증/생년월일 추출
        for i, line in enumerate(lines):
            for cert in certs:
                if cert in line:
                    cert_match = re.search(r'(전기[가-힣]*기[사술장]|[가-힣]*기[사술장])', line)
                    if cert_match and not personnel["책임감리원"]["자격증"]:
                        personnel["책임감리원"]["자격증"] = cert_match.group(1)

            birth_match = re.search(r'(\d{4}-\d{2}-\d{2})', line)
            if birth_match:
                bdate = birth_match.group(1)
                year = int(bdate[:4])
                if 1940 < year < 2000:
                    if not personnel["책임감리원"]["생년월일"]:
                        personnel["책임감리원"]["생년월일"] = bdate
                    elif not personnel["비상주감리원"]["생년월일"]:
                        personnel["비상주감리원"]["생년월일"] = bdate

    return personnel


def extract_career_summary(doc):
    """책임감리원 경력실적사항(양식2-5) 추출 - PyMuPDF 텍스트 사용"""
    career = {
        "책임_전기분야_개월": 0,
        "책임_전기분야_년": 0,
        "책임_참여분야1_개월": 0,
        "책임_참여분야1_년": 0,
        "책임_참여분야2_개월": 0,
        "책임_참여분야2_년": 0,
        "책임_전기분야_점수": 0,
        "책임_참여분야1_점수": 0,
        "책임_참여분야2_점수": 0,
        "page": -1,
    }

    page_num = smart_find_page(doc, ['양식2-5', '경력실적사항'],
                                range(17, min(30, doc.page_count)), "경력실적")
    if page_num < 0:
        page_num = smart_find_page(doc, ['책임감리원', '전기분야'],
                                    range(15, min(35, doc.page_count)), "경력실적(확장)")

    if page_num >= 0:
        career["page"] = page_num + 1
        lines = get_page_lines(doc, page_num)

        section = None  # "전기", "참여1", "참여2"
        months_count = 0
        score_count = 0

        for line in lines:
            # 섹션 판별
            if '가.' in line and '전기분야' in line:
                section = "전기"
            elif '나-1' in line or ('나.' in line and '참여분야' in line and '감독' not in line and '감리' not in line):
                section = "참여1"
            elif '나-2' in line or ('참여분야' in line and ('감독' in line or '감리' in line)):
                section = "참여2"

            # 개월수 추출 (예: 409.33개월)
            m = re.search(r'(\d+\.?\d*)\s*개월', line)
            if m:
                months = float(m.group(1))
                if section == "전기" and career["책임_전기분야_개월"] == 0:
                    career["책임_전기분야_개월"] = months
                    career["책임_전기분야_년"] = round(months / 12, 2)
                elif section == "참여1" and career["책임_참여분야1_개월"] == 0:
                    career["책임_참여분야1_개월"] = months
                    career["책임_참여분야1_년"] = round(months / 12, 2)
                elif section == "참여2" and career["책임_참여분야2_개월"] == 0:
                    career["책임_참여분야2_개월"] = months
                    career["책임_참여분야2_년"] = round(months / 12, 2)
                elif career["책임_전기분야_개월"] == 0:
                    career["책임_전기분야_개월"] = months
                    career["책임_전기분야_년"] = round(months / 12, 2)

            # 점수 추출 (예: "10 점", "라. 전기분야 10 점")
            m = re.search(r'(\d+\.?\d*)\s*점', line)
            if m:
                score = float(m.group(1))
                if '라' in line or '전기분야' in line:
                    if score <= 10:
                        career["책임_전기분야_점수"] = score
                elif '마-1' in line or ('참여분야' in line and '감독' not in line):
                    if score <= 10:
                        career["책임_참여분야1_점수"] = score
                elif '마-2' in line or ('참여분야' in line and ('감독' in line or '감리' in line)):
                    if score <= 5:
                        career["책임_참여분야2_점수"] = score
                else:
                    # 순서대로 할당
                    if career["책임_전기분야_점수"] == 0 and score <= 10:
                        career["책임_전기분야_점수"] = score
                    elif career["책임_참여분야1_점수"] == 0 and score <= 10:
                        career["책임_참여분야1_점수"] = score
                    elif career["책임_참여분야2_점수"] == 0 and score <= 5:
                        career["책임_참여분야2_점수"] = score

    return career


def extract_similar_project(doc):
    """유사용역 수행실적 추출 - 여러 페이지에 걸쳐 합계 찾기"""
    result = {
        "적용금액": 0,
        "사정금액": 0,
        "비율": 0,
        "점수": 0,
        "page": -1,
    }

    # 유사용역 시작 페이지 찾기
    start_page = smart_find_page(doc, ['유사용역실적', '유사용역', '환산금액'],
                                range(25, min(40, doc.page_count)), "유사용역")

    if start_page >= 0:
        result["page"] = start_page + 1

        # 시작 페이지부터 최대 10페이지 범위에서 합계/누계 찾기
        for page_num in range(start_page, min(start_page + 10, doc.page_count)):
            text = doc[page_num].get_text()
            if not text.strip():
                break  # 빈 페이지 도달 시 중단
            lines = [l.strip() for l in text.split('\n') if l.strip()]

            for line in lines:
                # 합계/누계 금액 추출 (천원 단위)
                if '누계' in line or '합계' in line:
                    nums = re.findall(r'[\d,]+', line)
                    for n in nums:
                        try:
                            val = int(n.replace(',', ''))
                            if val > 10000:
                                if result["적용금액"] == 0 or val > result["적용금액"]:
                                    result["적용금액"] = val
                        except ValueError:
                            pass

                # 사정금액
                if '사정금액' in line:
                    nums = re.findall(r'[\d,]+', line)
                    for n in nums:
                        try:
                            val = int(n.replace(',', ''))
                            if val > 10000:
                                result["사정금액"] = val
                        except ValueError:
                            pass

                # 비율 (예: 912%)
                m = re.search(r'(\d+)\s*%', line)
                if m:
                    val = int(m.group(1))
                    if val > 10:
                        result["비율"] = val

                # 평점/점수
                if '평점' in line or '평가점수' in line or '평가' in line:
                    m = re.search(r'(\d+\.?\d*)', line)
                    if m:
                        score = float(m.group(1))
                        if 0 < score <= 10:
                            result["점수"] = score

    return result


def extract_financial(doc):
    """재정상태 건실도 추출 - PyMuPDF 텍스트 + OCR 결합"""
    result = {
        "자기자본비율": 0,
        "평균자기자본비율": 0,
        "자본비율_대비": 0,
        "자본비율_점수": 0,
        "유동비율": 0,
        "평균유동비율": 0,
        "유동비율_대비": 0,
        "유동비율_점수": 0,
        "page": -1,
    }

    # 재정상태 페이지 찾기 (공사감리용역수행현황확인서)
    page_num = smart_find_page(doc, ['자기자본비율', '유동비율', '수행현황확인서'],
                                range(30, min(50, doc.page_count)), "재정상태")
    if page_num < 0:
        page_num = smart_find_page(doc, ['자기자본', '유동비율'],
                                    range(25, min(60, doc.page_count)), "재정상태(확장)")

    if page_num >= 0:
        result["page"] = page_num + 1

        # PyMuPDF 텍스트 시도
        text = get_page_text(doc, page_num)
        lines = [l.strip() for l in text.split('\n') if l.strip()]

        # 자기자본비율 추출
        for i, line in enumerate(lines):
            if '자기자본비율' in line:
                # 같은 줄 또는 다음 몇 줄에서 숫자 찾기
                search_text = ' '.join(lines[i:min(i+5, len(lines))])
                nums = re.findall(r'(\d+\.?\d+)', search_text)
                nums_float = [float(n) for n in nums if 0 < float(n) < 200]
                if len(nums_float) >= 2:
                    result["자기자본비율"] = nums_float[0]
                    result["평균자기자본비율"] = nums_float[1]
                elif len(nums_float) == 1:
                    result["자기자본비율"] = nums_float[0]

            if '유동비율' in line and '자기자본' not in line:
                search_text = ' '.join(lines[i:min(i+5, len(lines))])
                nums = re.findall(r'(\d+\.?\d+)', search_text)
                nums_float = [float(n) for n in nums if 0 < float(n) < 500]
                if len(nums_float) >= 2:
                    result["유동비율"] = nums_float[0]
                    result["평균유동비율"] = nums_float[1]
                elif len(nums_float) == 1:
                    result["유동비율"] = nums_float[0]

        # PyMuPDF로 못 찾았으면 OCR 시도 (캐시된 결과만 사용, 새 OCR 실행 안함)
        if result["자기자본비율"] == 0:
            try:
                items = ocr_page(doc, page_num)
                if items:
                    texts = [t for (t, c, b) in items]
                    for i, t in enumerate(texts):
                        if '자기자본' in t:
                            found_nums = []
                            for j in range(i+1, min(i+8, len(texts))):
                                m = re.search(r'(\d[\d,]*\.?\d*)', texts[j])
                                if m:
                                    val = float(m.group(1).replace(',', ''))
                                    if 0 < val < 200:
                                        found_nums.append(val)
                            if len(found_nums) >= 2:
                                result["자기자본비율"] = found_nums[0]
                                result["평균자기자본비율"] = found_nums[1]
            except Exception as e:
                print(f"  [WARN] 재정상태 OCR 폴백 실패: {type(e).__name__}")

        # 대비 계산
        if result["평균자기자본비율"] > 0:
            result["자본비율_대비"] = round(result["자기자본비율"] / result["평균자기자본비율"] * 100, 2)
        if result["평균유동비율"] > 0:
            result["유동비율_대비"] = round(result["유동비율"] / result["평균유동비율"] * 100, 2)

        # 점수 계산
        for threshold, score in CAPITAL_RATIO_SCORE:
            if result["자본비율_대비"] >= threshold:
                result["자본비율_점수"] = score
                break
        for threshold, score in CURRENT_RATIO_SCORE:
            if result["유동비율_대비"] >= threshold:
                result["유동비율_점수"] = score
                break

    return result


def extract_sanctions(doc):
    """제재내역 (영업정지) 추출 - PyMuPDF 텍스트 사용"""
    result = {
        "업체_영업정지": "없음",
        "감리원_자격정지": "없음",
        "감점": 0,
        "page": -1,
    }

    # 제재내역은 보통 재정상태와 같은 페이지 또는 근처
    page_num = smart_find_page(doc, ['제재내역', '영업정지', '이하여백'],
                                range(33, min(50, doc.page_count)), "제재내역")
    if page_num < 0:
        page_num = smart_find_page(doc, ['제재', '영업정지', '벌점'],
                                    range(25, min(60, doc.page_count)), "제재내역(확장)")

    if page_num >= 0:
        result["page"] = page_num + 1
        text = get_page_text(doc, page_num)

        if '이하여백' in text or '해당없음' in text or '해당 없음' in text:
            pass  # 제재내역 없음 확인
        else:
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            for line in lines:
                # 영업정지 기간 추출
                m = re.search(r'영업정지\s*[:：]?\s*(\d+)\s*[개월일]', line)
                if m:
                    result["업체_영업정지"] = f"{m.group(1)}개월"
                    months = int(m.group(1))
                    result["감점"] += months * 0.2  # 1월당 -0.2점

                # 자격정지 추출
                m = re.search(r'자격정지\s*[:：]?\s*(\d+)\s*[개월일]', line)
                if m:
                    result["감리원_자격정지"] = f"{m.group(1)}개월"
                    months = int(m.group(1))
                    result["감점"] += (months // 3) * 0.2  # 3월당 -0.2점

                # 부실벌점 추출
                m = re.search(r'벌점\s*[:：]?\s*(\d+\.?\d*)', line)
                if m:
                    result["감점"] += float(m.group(1))

    return result


def check_disqualification(personnel, career):
    """실격 여부 확인"""
    reasons = []

    # 1. 등급 미달 확인 (입찰공고에서 요구하는 등급)
    chief_grade = personnel["책임감리원"]["등급"]
    if chief_grade and chief_grade not in ["특급", "고급", "중급"]:
        reasons.append(f"책임감리원 등급 미달({chief_grade})")

    # 비상주감리원 고급 미만 시 실격
    nonres_grade = personnel["비상주감리원"]["등급"]
    if nonres_grade and nonres_grade not in ["특급", "고급"]:
        reasons.append(f"비상주감리원 등급 미달({nonres_grade})")

    return reasons


def get_cert_bonus(cert_name):
    """자격증 가점 계산"""
    if not cert_name:
        return 0

    if '기술사' in cert_name:
        return 1.0
    elif '기능장' in cert_name:
        return 0.8
    elif '기사' in cert_name and '산업' not in cert_name:
        return 0.7
    elif '산업기사' in cert_name:
        return 0.5
    return 0


###############################################################################
# 메인 분석 함수
###############################################################################

def analyze_company(pdf_path, bidding_date=BIDDING_DATE):
    """업체 PQ 제출 PDF 전체 분석"""
    print(f"\n{'='*60}")
    print(f"  배전감리 PQ 심사 데이터 추출")
    print(f"  대상 파일: {os.path.basename(pdf_path)}")
    print(f"  입찰공고일: {bidding_date}")
    print(f"{'='*60}\n")

    set_current_pdf(pdf_path)  # 파일 캐시 경로 설정
    doc = fitz.open(pdf_path)
    print(f"[INFO] PDF 열림: 총 {doc.page_count} 페이지")
    cache_dir = _get_cache_dir()
    if cache_dir and os.path.exists(cache_dir):
        cached_count = len([f for f in os.listdir(cache_dir) if f.endswith('.json')])
        if cached_count > 0:
            print(f"[INFO] 캐시 발견: {cached_count}페이지 OCR 결과 재사용")
    print()

    # 1. 업체명 추출
    print("[1/5] 업체명 추출 중...")
    company_name = extract_company_name(doc)
    print(f"  → 업체명: {company_name}")

    # 2. 종합득점표 추출
    print("[2/5] 종합득점표 추출 중...")
    summary = extract_summary_table(doc)
    print(f"  → 종합득점표 위치: p{summary['page']}")
    print(f"  → 참여감리원: {summary['참여감리원']}점, 유사용역: {summary['유사용역']}점")
    print(f"  → 기술개발: {summary['기술개발']}점, 업무중첩도: {summary['업무중첩도']}점")
    print(f"  → 교체빈도: {summary['교체빈도']}점, 가감점: {summary['가감점']}점")
    print(f"  → 업체제출 총점: {summary['총점']}")

    # 3. 참여감리원 자격사항 추출
    print("[3/5] 참여감리원 자격사항 추출 중...")
    personnel = extract_personnel(doc)
    print(f"  → 책임감리원: {personnel['책임감리원']['성명']} / {personnel['책임감리원']['등급']}")
    if personnel["보조감리원"]:
        print(f"  → 보조감리원: {personnel['보조감리원'][0]['성명']}")
    print(f"  → 비상주감리원: {personnel['비상주감리원']['성명']} / {personnel['비상주감리원']['등급']}")

    # 4. 책임감리원 경력 추출
    print("[4/5] 책임감리원 경력 추출 중...")
    career = extract_career_summary(doc)
    print(f"  → 전기분야: {career['책임_전기분야_년']}년 → {career['책임_전기분야_점수']}점")
    print(f"  → 참여분야(전체): {career['책임_참여분야1_년']}년 → {career['책임_참여분야1_점수']}점")
    print(f"  → 참여분야(감독감리): {career['책임_참여분야2_년']}년 → {career['책임_참여분야2_점수']}점")

    # 5. 유사용역 실적 추출
    print("[5/5] 유사용역 수행실적 추출 중...")
    similar = extract_similar_project(doc)
    print(f"  → 적용금액: {similar['적용금액']:,}원")
    if similar["점수"] > 0:
        print(f"  → 평점: {similar['점수']}점")

    # 제재내역 추출 (간략)
    sanctions = extract_sanctions(doc)

    # 실격 여부 판정
    disq_reasons = check_disqualification(personnel, career)

    # 가점 계산
    cert_bonus = get_cert_bonus(personnel["책임감리원"]["자격증"])

    doc.close()

    # 결과 종합
    result = {
        "업체명": company_name,
        "실격여부": "실격" if disq_reasons else "적격",
        "실격사유": "; ".join(disq_reasons) if disq_reasons else "없음",
        "책임_성명": personnel["책임감리원"]["성명"],
        "책임_등급": personnel["책임감리원"]["등급"],
        "책임_자격증": personnel["책임감리원"]["자격증"],
        "책임_전기경력_개월": career["책임_전기분야_개월"],
        "책임_전기경력_년": career["책임_전기분야_년"],
        "책임_전기경력_점수": career["책임_전기분야_점수"],
        "책임_배전경력1_년": career["책임_참여분야1_년"],
        "책임_배전경력1_점수": career["책임_참여분야1_점수"],
        "책임_배전경력2_년": career["책임_참여분야2_년"],
        "책임_배전경력2_점수": career["책임_참여분야2_점수"],
        "보조1_성명": personnel["보조감리원"][0]["성명"] if personnel["보조감리원"] else "",
        "비상주_성명": personnel["비상주감리원"]["성명"],
        "비상주_등급": personnel["비상주감리원"]["등급"],
        "유사용역_적용금액": similar["적용금액"],
        "유사용역_점수": summary["유사용역"],
        "기술개발_점수": summary["기술개발"],
        "업무중첩_점수": summary["업무중첩도"],
        "교체빈도_점수": summary["교체빈도"],
        "참여감리원_소계": summary["참여감리원"],
        "영업정지": sanctions["업체_영업정지"],
        "가점_자격증": cert_bonus,
        "업체제출_총점": summary["총점"],
        # 근거 페이지
        "근거페이지": {
            "종합득점표": summary["page"],
            "참여감리원": personnel["page"],
            "경력실적": career["page"],
            "유사용역": similar["page"],
            "제재내역": sanctions["page"],
        }
    }

    return result


###############################################################################
# Excel 출력
###############################################################################

def export_to_excel(result, output_path):
    """분석 결과를 Excel로 출력"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PQ심사결과"

    # 스타일 설정
    header_font = Font(name='맑은 고딕', bold=True, size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(name='맑은 고딕', bold=True, size=11, color="FFFFFF")
    data_font = Font(name='맑은 고딕', size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 제목
    ws.merge_cells('A1:H1')
    ws['A1'] = '배전감리 PQ 심사 데이터 추출 결과'
    ws['A1'].font = Font(name='맑은 고딕', bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:H2')
    ws['A2'] = f'입찰공고일: {BIDDING_DATE} | 추출일시: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
    ws['A2'].font = Font(name='맑은 고딕', size=9, color="666666")
    ws['A2'].alignment = Alignment(horizontal='center')

    # 헤더 (3행)
    headers = [
        '업체명', '실격여부\n(사유)', '책임감리원\n성명/등급',
        '책임_전기경력\n(년/점수)', '책임_배전경력\n(년/점수)',
        '보조/비상주\n성명', '유사용역\n최종실적액',
        '근거페이지\n(PDF)'
    ]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border

    # 데이터 (4행)
    row = 5

    # 1. 업체명
    ws.cell(row=row, column=1, value=result["업체명"]).font = data_font

    # 2. 실격여부
    disq_text = result["실격여부"]
    if result["실격사유"] != "없음":
        disq_text += f"\n({result['실격사유']})"
    cell = ws.cell(row=row, column=2, value=disq_text)
    cell.font = data_font
    if result["실격여부"] == "실격":
        cell.font = Font(name='맑은 고딕', size=10, color="FF0000", bold=True)

    # 3. 책임감리원 성명/등급
    ws.cell(row=row, column=3,
            value=f"{result['책임_성명']}\n{result['책임_등급']}").font = data_font

    # 4. 전기경력
    ws.cell(row=row, column=4,
            value=f"{result['책임_전기경력_년']}년\n→ {result['책임_전기경력_점수']}점").font = data_font

    # 5. 배전경력
    ws.cell(row=row, column=5,
            value=f"전체: {result['책임_배전경력1_년']}년 → {result['책임_배전경력1_점수']}점\n"
                  f"감독감리: {result['책임_배전경력2_년']}년 → {result['책임_배전경력2_점수']}점").font = data_font

    # 6. 보조/비상주
    ws.cell(row=row, column=6,
            value=f"보조: {result['보조1_성명']}\n비상주: {result['비상주_성명']}({result['비상주_등급']})").font = data_font

    # 7. 유사용역 실적액
    ws.cell(row=row, column=7,
            value=f"{result['유사용역_적용금액']:,}원").font = data_font

    # 8. 근거페이지
    pages = result["근거페이지"]
    ws.cell(row=row, column=8,
            value=f"종합득점표: p{pages['종합득점표']}\n"
                  f"참여감리원: p{pages['참여감리원']}\n"
                  f"경력실적: p{pages['경력실적']}").font = data_font

    # 스타일 적용
    for col in range(1, 9):
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(vertical='top', wrap_text=True)

    # 열 너비 조정
    widths = [18, 15, 15, 18, 22, 18, 18, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64+i)].width = w

    # 행 높이
    ws.row_height = 60
    ws.row_dimensions[4].height = 40
    ws.row_dimensions[5].height = 100

    # 상세 데이터 시트
    ws2 = wb.create_sheet("상세데이터")
    ws2['A1'] = '항목'
    ws2['B1'] = '값'
    ws2['C1'] = '비고'
    ws2['A1'].font = header_font
    ws2['B1'].font = header_font
    ws2['C1'].font = header_font

    detail_rows = [
        ('업체명', result['업체명'], ''),
        ('실격여부', result['실격여부'], result['실격사유']),
        ('---참여감리원---', '', ''),
        ('책임감리원 성명', result['책임_성명'], ''),
        ('책임감리원 등급', result['책임_등급'], ''),
        ('책임감리원 자격증', result['책임_자격증'], f"가점: {result['가점_자격증']}점"),
        ('책임_전기경력(개월)', result['책임_전기경력_개월'], ''),
        ('책임_전기경력(년)', result['책임_전기경력_년'], f"절사후 → {result['책임_전기경력_점수']}점"),
        ('책임_배전경력1(년)', result['책임_배전경력1_년'], f"전체 → {result['책임_배전경력1_점수']}점"),
        ('책임_배전경력2(년)', result['책임_배전경력2_년'], f"감독감리 → {result['책임_배전경력2_점수']}점"),
        ('보조감리원 성명', result['보조1_성명'], ''),
        ('비상주감리원 성명', result['비상주_성명'], ''),
        ('비상주감리원 등급', result['비상주_등급'], ''),
        ('---유사용역---', '', ''),
        ('유사용역 적용금액', f"{result['유사용역_적용금액']:,}원", ''),
        ('영업정지', result['영업정지'], ''),
        ('---합계---', '', ''),
        ('업체제출 총점', result['업체제출_총점'], '100점 만점 기준'),
    ]

    for i, (item, val, note) in enumerate(detail_rows, 2):
        ws2.cell(row=i, column=1, value=item).font = data_font
        ws2.cell(row=i, column=2, value=str(val)).font = data_font
        ws2.cell(row=i, column=3, value=note).font = data_font

    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 25
    ws2.column_dimensions['C'].width = 30

    wb.save(output_path)
    print(f"\n[OUTPUT] Excel 저장 완료: {output_path}")


###############################################################################
# CSV 출력
###############################################################################

def print_csv_result(result):
    """CSV 형태로 콘솔 출력"""
    print(f"\n{'='*60}")
    print("  CSV 출력 결과")
    print(f"{'='*60}")
    print(f"1. [업체명]: {result['업체명']}")
    print(f"2. [실격여부]: {result['실격여부']} ({result['실격사유']})")
    print(f"3. [책임_성명/등급]: {result['책임_성명']} / {result['책임_등급']}")
    print(f"4. [책임_전기경력_일수]: {result['책임_전기경력_개월']}개월 ({result['책임_전기경력_년']}년) → {result['책임_전기경력_점수']}점")
    print(f"   [책임_배전경력_일수]: 전체 {result['책임_배전경력1_년']}년 → {result['책임_배전경력1_점수']}점 / 감독감리 {result['책임_배전경력2_년']}년 → {result['책임_배전경력2_점수']}점")
    print(f"5. [보조1_성명]: {result['보조1_성명']}")
    print(f"   [비상주_성명/등급]: {result['비상주_성명']} / {result['비상주_등급']}")
    print(f"6. [유사용역_최종실적액]: {result['유사용역_적용금액']:,}원")
    print(f"7. [영업정지]: {result['영업정지']}")
    print(f"8. [가점_자격증]: {result['가점_자격증']}점 ({result['책임_자격증']})")

    pages = result["근거페이지"]
    print(f"9. [근거페이지]: 종합득점표 p{pages['종합득점표']}, 참여감리원 p{pages['참여감리원']}, "
          f"경력실적 p{pages['경력실적']}")

    print(f"\n[업체제출 총점]: {result['업체제출_총점']}점")


###############################################################################
# 메인 실행
###############################################################################

if __name__ == "__main__":
    # 기본 경로
    pdf_path = r"C:\Users\Admin\Desktop\project\pq\증빙\(주)성문기술단(505-81-31517)-복사.pdf"

    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    if len(sys.argv) > 2:
        BIDDING_DATE = sys.argv[2]

    if not os.path.exists(pdf_path):
        print(f"[ERROR] 파일을 찾을 수 없습니다: {pdf_path}")
        sys.exit(1)

    # 분석 실행
    result = analyze_company(pdf_path, BIDDING_DATE)

    # CSV 형태 콘솔 출력
    print_csv_result(result)

    # Excel 출력
    output_dir = os.path.dirname(pdf_path) or "."
    output_path = os.path.join(os.path.dirname(os.path.dirname(pdf_path)),
                               "PQ심사결과_" + result["업체명"].replace("(", "").replace(")", "") + ".xlsx")
    export_to_excel(result, output_path)

    print(f"\n{'='*60}")
    print("  분석 완료!")
    print(f"{'='*60}")
