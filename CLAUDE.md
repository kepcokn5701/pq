# 배전감리 PQ 심사 교차검증 시스템 — 프로젝트 메모리

## 프로젝트 개요

한국전력공사 배전건설부 감리용역 적격심사(PQ) 자동 교차검증 시스템.
업체가 제출한 PDF(사업수행능력 평가서)와 Excel(심사표)을 읽어 점수를 독립 재계산하고,
업체 기재값과 기준 계산값의 불일치를 자동 감지한다.

- **서버 실행**: `python app.py` (포트 5002)
- **접속 URL**: `http://localhost:5002`
- **입찰공고일**: `BIDDING_DATE = "2026-03-09"` (pq_extractor.py 상단 상수)

---

## 핵심 파일 구조

```
pq/
├── app.py                  Flask 웹 서버, API 라우트, cross_verify()
├── pq_extractor.py         PDF 파싱 + 점수 계산 엔진
├── templates/index.html    단일 페이지 UI
├── 증빙/                   업체 PDF 파일 위치 (gitignore)
├── 심사기준/
│   └── pq심사기준.pdf      감리용역 PQ 심사 기준 원본
│   └── pq심사기준_OCR.txt  OCR 추출본 (EasyOCR, 36페이지)
└── uploads/                Flask 업로드 임시폴더
```

---

## 점수 기준표 (pq_extractor.py 상단 상수)

| 상수 | 내용 |
|---|---|
| `CHIEF_ELEC_CAREER_SCORE` | 책임감리원 전기분야 경력 점수 (등급×년수) |
| `CHIEF_FIELD_CAREER_SCORE_1` | 참여분야 경력1 (설계·시공·감리 전체) |
| `CHIEF_FIELD_CAREER_SCORE_2` | 참여분야 경력2 (감독·감리만) |
| `SIMILAR_PROJECT_SCORE` | 유사용역 실적 (사정금액 대비 %, 0~100 범위) |
| `NONRESIDENT_GRADE_SCORE` | 비상주감리원 등급 (특급=3, 고급=2) |
| `TECH_INVEST_SCORE` | 기술개발 투자실적 (A÷B%, 3.0/2.5/2.0/1.5/0 → 4/3.5/3/2.5/2점) |
| `TECH_DEV_SCORE` | 개발실적 (특허/실용신안/전력신기술 × 기간비율 100/80/60%) |
| `REPLACEMENT_RATE_SCORE` | 교체빈도 감리업체 점수 (높은 임계값 순 정렬 주의) |

**주의**: `lookup_score(value, table)` 함수는 내림차순 임계값 테이블 가정.
`TECH_INVEST_SCORE`의 임계값은 plain % (3.0, 2.5, ...), ×100 아님.

---

## 주요 함수

### pq_extractor.py

| 함수 | 역할 |
|---|---|
| `analyze_company(pdf_path, bidding_date)` | PDF 전체 분석 → result dict 반환 |
| `extract_tech_development(doc, bidding_date)` | 양식2-9(기술개발) 세부증빙 파싱 |
| `calculate_scores_by_criteria(pdf_data)` | 심사기준표 적용 독립 계산 → criteria_scores |
| `lookup_score(value, table)` | 임계값 테이블 조회 |
| `lookup_career_score(grade, years, table)` | 등급+경력년수 조회 |
| `_dev_score_from_ratio(item_type, ratio_pct)` | 개발실적 타입×비율 → 점수 |
| `smart_find_page(doc, keywords, range)` | 내장 텍스트만으로 페이지 탐색 (OCR 없음) |
| `get_page_text(doc, page_num)` | 내장 텍스트 우선, 없으면 OCR 폴백 |

### app.py

| 함수 | 역할 |
|---|---|
| `cross_verify(pdf_data, excel_data)` | 3방향 비교 (기준계산/PDF추출/Excel제출) |
| `_compare_numeric(category, label, pdf_val, excel_val, ...)` | 수치 비교 항목 생성 |
| `_add_criteria(item, criteria_value, criteria_basis, pdf_val)` | 기준계산값 필드 추가 |
| `parse_company_excel(path)` | 업체 제출 Excel 파싱 |

---

## extract_tech_development() 핵심 로직

**양식2-9 페이지 탐색**: `smart_find_page(['양식2-9', '기술개발투자실적'], range(70,100))`

**가. 개발실적 파싱**:
- 날짜쌍(지정일 + 유효기간)을 1세트로 인식 (다음 라인도 날짜면 유효기간)
- 기재비율% + 기재평점을 읽어 타입 역산: `base = 기재평점 / (기재비율/100)`
  - base ≈ 2.0 → 전력신기술, ≈ 1.0 → 특허, ≈ 0.5 → 실용신안
- 지정일 ~ 입찰공고일 경과년수 계산 → 구간(5년이하/5~10년/10~20년) 판정
- `_dev_score_from_ratio(타입, 계산구간)`으로 독립 재계산
- 건별 합산, max 4.0점

**나. 기술투자실적**: `(A)` 레이블 직전 숫자 = 투자액 합계, `(B)` 직전 = 매출액 합계

**다. 교육실적**: 1.5~4.0 범위 합계 점수 탐색 (개인별 1점 마크와 구분)

---

## UI 컬럼 구조 (index.html)

8열 테이블:
1. 심사항목 (category)
2. 배점 (alloc)
3. 세부항목 (sub)
4. **기준 계산값** `criteria_scores` 기반 독립 계산 (노란 배경)
5. **PDF 추출값** `pdf_data` 직접 추출
6. **Excel 제출값** `excel_data` 업체 제출
7. 검증결과 (✓/⚠/✗)
8. 비고

`critFmt(calcScore, pdfScore, basis)` → 불일치 시 `⚠` + 붉은색 표시

---

## criteria_scores 현재 구현 항목

```
책임_전기경력_계산점수 / 근거
책임_배전경력1_계산점수 / 근거
책임_배전경력2_계산점수 / 근거
비상주_등급_계산점수 / 근거
유사용역_계산점수 / 근거
가점_계산점수 / 근거
기술투자_계산점수 / 근거   ← 세부증빙 A÷B% 기반
개발실적_계산점수 / 근거   ← 세부증빙 건별 합산 기반
```

---

## 주요 구현 이슈 / 해결 내역

### 서버 재시작
`debug=True, use_reloader=False` 설정.
- reloader OFF: 분석 도중 파일 변경으로 서버가 재시작되어 "Failed to fetch" 오류 방지
- **서버 코드 변경** (app.py, pq_extractor.py): 서버 수동 재시작 필요 (`python app.py`)
- **프론트엔드 변경** (index.html): 브라우저 F5, 캐시 문제 시 Ctrl+Shift+R (강력 새로고침)
- **서버 재시작**: `netstat -ano | findstr :5002`로 잔여 프로세스 확인 후 `taskkill`로 종료, `python app.py` 재실행

### REPLACEMENT_RATE_SCORE 정렬 방향
높은 임계값(50%)이 먼저 와야 함 — `lookup_score`가 내림차순 가정.
낮은 값부터 정렬 시 항상 첫 번째 항목만 매칭되는 버그 발생.

### TECH_INVEST_SCORE 임계값 단위
plain % 사용 (3.0, 2.5, 2.0, 1.5). 이전에 ×100 형식(300, 250 ...)으로 잘못 작성됐다가 수정됨.

### 개발실적 이중계산 버그
날짜쌍(지정일+유효기간)을 각각 별도 항목으로 처리해 20건이 되던 문제.
`date_pat.match(line) AND date_pat.match(all_lines[i+1])` 조건으로 쌍 감지 후 advance_to로 건너뜀.

### 양식2-9 섹션 헤더 타입 불일치
`(1)전력신기술` 헤더 아래에 특허/실용신안 항목이 섞여 있음.
섹션 헤더로 타입을 결정하지 않고, 기재평점/기재비율로 타입을 역산하는 방식으로 해결.

---

## 테스트 파일

`증빙/(주)성문기술단(505-81-31517)-복사.pdf`
- 134페이지, 텍스트 추출 가능
- 기술개발 페이지: p81 (0-indexed p80)
- 기대 결과:
  - 개발실적: 10건 (특허 6건 × 60% + 실용신안 4건 × 80%) → 5.2 → max 4.0점
  - 투자비율: 4.88% → 4.0점
  - 교육실적 기재점수: 2.0점

### OCR 속도 최적화 (2026-05-21)
EasyOCR CPU 기준 속도 벤치마크:
- 헤더 OCR (상단25%, 1.5x): ~3초/페이지
- 전체 OCR (1.5x): ~12초/페이지
- 전체 OCR (2.0x): ~15초/페이지

PaddleOCR 평가 결과: 이 Windows CPU 환경에서 EasyOCR보다 2~3배 느림 → **EasyOCR 유지**

적용된 최적화:
- `ocr_page()` 기본 zoom: 2.0x → 1.5x (정확도 유지, ~20% 속도 향상)
- `extract_personnel_scan()` zoom: 2.0x → 1.5x
- `_find_personnel_page_ocr()` 검색범위: 80→25페이지(1차), 20→10페이지(2차)
- `classify_pages_ocr()` 최대 OCR 30회 + 앞쪽 60%만 탐색
- 스캔 PDF 조기 감지: `is_scan_pdf` 플래그로 불필요한 전체 OCR 분류 스킵
- `page_map` 활용: `extract_personnel_scan()`에서 이미 찾은 양식2-4 위치 재사용
- `gc.collect()` 정상 경로에서 제거 (에러 시에만 유지)

결과: 스캔 PDF(205페이지) 332초 → 108초 (3배 개선)

---

## 미구현 / TODO

- [x] 스캔 PDF 대응: `smart_find_page_with_ocr()` 구현 완료
- [ ] 스캔 PDF OCR 정확도: 양식2-4 헤더 인식률 개선 (스캔 이미지 품질 의존)
- [ ] 업무중첩도 세부 계산 (상주/비상주)
- [ ] 교체빈도 세부 계산
- [ ] 참여감리원 항목 외부 검증 (전기공사협회 인증서 연동)
