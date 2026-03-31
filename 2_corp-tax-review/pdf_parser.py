"""
법인세 세무조정계산서 PDF 파서
PDF에서 각 서식별 데이터를 추출하는 모듈
"""
import re
import os
import fitz  # PyMuPDF


def extract_all_text(pdf_path):
    """PDF 전체 페이지의 텍스트를 추출"""
    doc = fitz.open(pdf_path)
    pages = {}
    for i in range(doc.page_count):
        pages[i + 1] = doc[i].get_text()
    doc.close()
    return pages


def extract_all_numbers(text):
    """텍스트에서 숫자(콤마 포함)를 순서대로 추출"""
    nums = []
    for m in re.finditer(r'(?<!\d)([\d,]{1,20})(?!\d)', text):
        raw = m.group(1).replace(",", "")
        if raw.isdigit() and len(raw) >= 1:
            nums.append(int(raw))
    return nums


def extract_large_numbers(text, min_val=10000):
    return [n for n in extract_all_numbers(text) if n >= min_val]


def find_page_by_keyword(pages, keyword, skip_toc=True, min_page=4):
    """키워드가 포함된 첫 번째 페이지 번호 반환 (목차 건너뛰기)"""
    for pg, text in pages.items():
        if skip_toc and pg < min_page:
            continue
        if keyword in text.replace("\n", "").replace(" ", ""):
            return pg
    return None


def find_all_pages_by_keyword(pages, keyword, min_page=4):
    """키워드가 포함된 모든 페이지 번호 반환"""
    result = []
    for pg, text in pages.items():
        if pg < min_page:
            continue
        if keyword in text.replace("\n", "").replace(" ", ""):
            result.append(pg)
    return result


def join_text(text):
    return " ".join(text.split())


def extract_line_amounts(text, min_val=1000):
    """텍스트에서 줄 단위로 금액만 추출"""
    amounts = []
    for line in text.strip().split("\n"):
        m = re.match(r'^([\d,]+)$', line.strip())
        if m:
            raw = m.group(1).replace(",", "")
            if raw.isdigit() and int(raw) >= min_val:
                amounts.append(int(raw))
    return amounts


def extract_code_amounts(text):
    """코드번호(01, 02...) 다음의 금액을 매핑"""
    lines = text.strip().split("\n")
    code_amounts = {}
    for i, line in enumerate(lines):
        if re.match(r'^\d{1,2}$', line.strip()):
            code = int(line.strip())
            for j in range(i + 1, min(i + 3, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    val = int(m.group(1).replace(",", ""))
                    if val >= 1000:
                        code_amounts[code] = val
                        break
                elif lines[j].strip() and not re.match(r'^\d{1,2}$', lines[j].strip()):
                    break
    return code_amounts


# ============================================================
# 회사 기본정보
# ============================================================
def parse_company_info(pages):
    info = {}
    for pg in range(1, min(8, len(pages) + 1)):
        text = pages.get(pg, "")
        flat = join_text(text)

        m = re.search(r'(\d{3}-\d{2}-\d{5})', flat)
        if m and "사업자등록번호" not in info:
            info["사업자등록번호"] = m.group(1)

        if "법인명" not in info:
            # 1순위: "법인명" 키워드 뒤 (주)xxx 또는 주식회사xxx
            m = re.search(r'(?:법\s*인\s*명|법인명)\s*(\(주\)\S+)', flat)
            if m:
                info["법인명"] = m.group(1)
            else:
                m = re.search(r'(?:법\s*인\s*명|법인명)\s*(주식회사\S+)', flat)
                if m and '외부감사' not in m.group(1) and '법률' not in m.group(1):
                    info["법인명"] = m.group(1)
        # 2순위: 페이지 텍스트에서 (주)xxx 단독 패턴
        if "법인명" not in info:
            lines = text.strip().split("\n")
            for line in lines[:20]:
                m2 = re.match(r'^\s*\(주\)(\S+)\s*$', line.strip())
                if m2:
                    info["법인명"] = f"(주){m2.group(1)}"
                    break

        if "대표자" not in info:
            m = re.search(r'(?:대표자\s*성명|대\s*표\s*자\s*성\s*명)\s*(\S+)', flat)
            if m and m.group(1) not in ["⑥", "("]:
                info["대표자"] = m.group(1)

        m = re.search(r'(\d{4}\.\d{2}\.\d{2})\s*~\s*(\d{4}\.\d{2}\.\d{2})', flat)
        if m and "사업연도_시작" not in info:
            info["사업연도_시작"] = m.group(1)
            info["사업연도_종료"] = m.group(2)

        m = re.search(r'업\s*태\s*(.+?)(?:⑨|종목)', flat)
        if m and "업태" not in info:
            info["업태"] = m.group(1).strip()

        m = re.search(r'종목\s*(.+?)(?:⑩|주업종)', flat)
        if m and "종목" not in info:
            info["종목"] = m.group(1).strip()

    # 본점소재지 파싱
    for pg in range(1, min(10, len(pages) + 1)):
        text = pages.get(pg, "")
        lines = text.strip().split("\n")
        for line in lines:
            stripped = line.strip()
            # "지서울..." "지경기..." 등 주소 패턴
            m = re.match(r'^지?(서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충북|충남|전북|전남|경북|경남|제주)(.+)', stripped)
            if m and len(stripped) > 8:
                info["본점소재지"] = stripped if not stripped.startswith("지") else stripped[1:]
                시도 = m.group(1)
                info["수도권"] = 시도 in ("서울", "경기", "인천")
                break
        if "본점소재지" in info:
            break

    for pg, text in pages.items():
        flat_check = text.replace("\n", "").replace(" ", "")
        if "중소기업기준검토표" in flat_check or "중소기업등기준검토표" in flat_check:
            info["중소기업"] = True
            lines = text.strip().split("\n")
            # 소기업/중기업 구분: "소기업" 또는 "중기업" 키워드가 서식 제목 근처에 있음
            for i, line in enumerate(lines):
                stripped = line.strip()
                if stripped == "소" and i + 2 < len(lines) and lines[i+1].strip() == "기" and lines[i+2].strip() == "업":
                    info["소기업"] = True
                    break
                if stripped == "중" and i + 2 < len(lines) and lines[i+1].strip() == "기" and lines[i+2].strip() == "업":
                    info["소기업"] = False
                    break
                if "소기업" in stripped and "중소기업" not in stripped:
                    info["소기업"] = True
                    break
            if "소기업" in info:
                break

    if "법인명" not in info:
        for pg, text in pages.items():
            m = re.search(r'주식회사\S+', text)
            if m:
                info["법인명"] = m.group(0)
                break

    return info


# ============================================================
# 법인세과세표준및세액조정계산서 (별지제3호서식)
# ============================================================
def parse_tax_adjustment(pages):
    data = {}
    target_pg = None
    for pg, text in pages.items():
        flat = join_text(text)
        if "법인세과세표준및세액조정계산서" in flat and "결산서상당기순손익" in flat:
            target_pg = pg
            break
    if not target_pg:
        return data

    text = pages[target_pg]
    lines = text.strip().split("\n")

    # 키워드 기반 추출
    def find_amount_after(keyword, min_val=1000):
        found = False
        for i, line in enumerate(lines):
            if keyword in line.replace(" ", ""):
                found = True
                continue
            if found:
                m = re.match(r'^([\d,]+)$', line.strip())
                if m:
                    val = int(m.group(1).replace(",", ""))
                    if val >= min_val:
                        return val
                if re.match(r'^\d{1,2}$', line.strip()):
                    continue
                if len(line.strip()) > 5 and not line.strip().replace(",", "").isdigit():
                    m2 = re.match(r'^([\d,]+)', line.strip())
                    if m2:
                        val = int(m2.group(1).replace(",", ""))
                        if val >= min_val:
                            return val
                    found = False
        return None

    data["결산서상당기순손익"] = find_amount_after("결산서상당기순손익", 1000)
    data["익금산입"] = find_amount_after("익금산입", 1000)
    val = find_amount_after("손금산입", 1000)
    data["손금산입"] = val if val and val != data.get("익금산입") else 0
    data["각사업연도소득금액"] = find_amount_after("각사업연도소득금액", 1000)

    # 코드번호 기반
    code_amounts = extract_code_amounts(text)
    if 10 in code_amounts:
        data["과세표준"] = code_amounts[10]
    if 12 in code_amounts:
        data["산출세액"] = code_amounts[12]
    if 17 in code_amounts:
        data["최저한세적용대상_공제감면세액"] = code_amounts[17]
    if 19 in code_amounts:
        data["최저한세적용제외_공제감면세액"] = code_amounts[19]
    if 18 in code_amounts:
        data["차감세액"] = code_amounts[18]
    if 22 in code_amounts:
        data["중간예납세액"] = code_amounts[22]
    if 24 in code_amounts:
        data["원천납부세액"] = code_amounts[24]
    if 26 in code_amounts:
        data["기납부세액"] = code_amounts[26]
    if 30 in code_amounts:
        data["차감납부할세액"] = code_amounts[30]
    if 46 in code_amounts:
        data["차감납부할세액계"] = code_amounts[46]
    if 48 in code_amounts:
        data["분납세액"] = code_amounts[48]
    if 49 in code_amounts:
        data["차감납부세액"] = code_amounts[49]

    # 세율
    if 11 in code_amounts and code_amounts[11] in [9, 19, 21, 24]:
        data["세율"] = code_amounts[11]
    if "세율" not in data:
        for i, line in enumerate(lines):
            if "세" in line and i + 1 < len(lines) and "율" in lines[i + 1]:
                for j in range(i + 2, min(i + 5, len(lines))):
                    if lines[j].strip() in ["9", "19", "21", "24"]:
                        data["세율"] = int(lines[j].strip())
                        break

    return data


# ============================================================
# 최저한세조정계산서 (별지제4호서식)
# ============================================================
def parse_minimum_tax(pages, pdf_path=None):
    """최저한세조정계산서 - 테이블 추출 방식"""
    data = {}

    if pdf_path:
        try:
            import fitz as _fitz
            _doc = _fitz.open(pdf_path)
            for pg_num in range(len(_doc)):
                page_obj = _doc[pg_num]
                tabs = page_obj.find_tables()
                for tab in tabs.tables:
                    rows = tab.extract()
                    if len(rows) < 20 or len(rows[0]) < 5:
                        continue
                    found = False
                    for row in rows:
                        cell0 = (row[0] or "").replace(" ", "")
                        if "(122)" in cell0:
                            found = True
                            break
                    if not found:
                        continue

                    def safe_int(val):
                        if not val or not isinstance(val, str):
                            return None
                        v = val.strip().replace(",", "")
                        if not v.isdigit():
                            return None
                        return int(v)

                    for row in rows:
                        cell0 = (row[0] or "").replace(" ", "")
                        val_감면후 = safe_int(row[3] if len(row) > 3 else None)
                        val_최저한세 = safe_int(row[4] if len(row) > 4 else None)
                        if "(122)" in cell0:  # 산출세액
                            if val_감면후:
                                data["산출세액_감면후"] = val_감면후
                            if val_최저한세:
                                data["산출세액"] = val_최저한세
                        elif "(123)" in cell0:  # 감면세액
                            if val_감면후:
                                data["감면세액"] = val_감면후
                        elif "(124)" in cell0:  # 세액공제
                            if val_감면후:
                                data["세액공제합계"] = val_감면후
                        elif "(125)" in cell0:  # 차감세액
                            if val_감면후:
                                data["차감세액"] = val_감면후
                    if data:
                        break
                if data:
                    break
            _doc.close()
        except Exception:
            pass

    return data


# ============================================================
# 공제감면세액합계표 (별지제8호서식 갑/을)
# ============================================================
def parse_deduction_credits(pages):
    data = {
        "세액감면": {},
        "세액공제": {},
        "감면소계_적용제외": 0,
        "감면소계_적용대상": 0,
        "공제소계_적용대상": 0,
        "감면합계": 0,
        "공제합계": 0,
    }

    for pg, text in pages.items():
        if pg < 4:
            continue
        flat = join_text(text)
        lines = text.strip().split("\n")

        # 최저한세 적용대상 감면 소계 (130번) 및 공제감면 합계 (150, 151번)
        flat_nospace = flat.replace(" ", "")
        if "최저한세적용대상공제감면세액" in flat_nospace or "공제감면세액총계" in flat_nospace or "추가납부세액합계표" in flat_nospace:
            for i, line in enumerate(lines):
                stripped = line.strip()
                # 130번 = 감면 소계
                if stripped == "130":
                    for j in range(i + 1, min(i + 5, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            raw = m.group(1).replace(",", "")
                            if raw.isdigit() and int(raw) >= 1000:
                                data["감면소계_적용대상"] = int(raw)
                                break
                # 150번 = 최저한세적용대상 공제감면세액 합계
                if stripped == "150":
                    for j in range(i + 1, min(i + 5, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            raw = m.group(1).replace(",", "")
                            if raw.isdigit() and int(raw) >= 1000:
                                data["공제감면합계_적용대상"] = int(raw)
                                break
                # 151번 = 최저한세적용제외 공제감면세액 합계
                if stripped == "151":
                    for j in range(i + 1, min(i + 5, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            raw = m.group(1).replace(",", "")
                            if raw.isdigit() and int(raw) >= 1000:
                                data["공제감면합계_적용제외"] = int(raw)
                                break

        # 중소기업특별세액감면 (조특법 제7조)
        if "중소기업에" in flat and "법제7조" in flat:
            # 라인 기반 파싱: 법제7조 다음 줄들에서 숫자 추출
            sme_nums = []  # 큰 숫자 (금액)
            sme_small = []  # 작은 숫자 (공제율 등)
            found_법제7조 = False
            for i, line in enumerate(lines):
                if "법제7조" in line.strip():
                    found_법제7조 = True
                    continue
                if found_법제7조:
                    stripped = line.strip().split("×")[0].strip()
                    m = re.match(r'^([\d,]+)$', stripped)
                    if m:
                        val = int(m.group(1).replace(",", ""))
                        if val >= 1000:
                            sme_nums.append(val)
                        elif 1 <= val <= 100:
                            sme_small.append(val)
                    if len(sme_nums) >= 4 or (found_법제7조 and "합" in line and "계" in line):
                        break

            sme_data = {}
            if len(sme_nums) >= 4:
                # 순서: 감면대상소득, 산출세액(×포함줄), 총소득, 감면세액
                sme_data = {
                    "감면대상소득_계산서": sme_nums[0],
                    "산출세액_계산서": sme_nums[1],
                    "총소득_계산서": sme_nums[2],
                    "대상세액": sme_nums[3],
                    "감면세액": sme_nums[3],
                }
            elif len(sme_nums) >= 2:
                sme_data = {
                    "대상세액": sme_nums[0],
                    "감면세액": sme_nums[1] if len(sme_nums) > 1 else sme_nums[0],
                }
            # 공제율: 작은 숫자 중 첫 번째 (20, 30 등)
            if sme_small:
                sme_data["공제율"] = sme_small[0]
            if sme_data:
                data["세액감면"]["중소기업특별세액감면"] = sme_data

        # 통합고용세액공제 (합계표 내, p.11)
        if "통합고용세액공제" in flat and ("전기이월" in flat or "18S" in flat):
            idx = flat.find("통합고용세액공제")
            after = flat[idx:idx + 300]
            nums = [n for n in extract_all_numbers(after) if n >= 100000]
            if nums and len(nums) >= 3:
                data["세액공제"]["통합고용세액공제_합계표"] = {
                    "전기이월": nums[0],
                    "당기발생": nums[1],
                    "공제세액": nums[2],
                }

    data["감면합계"] = data["감면소계_적용제외"] + data["감면소계_적용대상"]

    return data


# ============================================================
# 세액공제조정명세서(3) (별지제8호서식부표3)
# ============================================================
def parse_tax_credit_adjustment(pages, pdf_path=None, fitz_doc=None):
    """세액공제조정명세서(3) - 2. 당기공제세액 및 이월액계산
    테이블 추출 방식 (한글 깨짐 대비)"""
    data = {}

    # 1차: 키워드 기반 (한글이 안 깨진 PDF)
    for pg, text in pages.items():
        if pg < 4:
            continue
        flat = join_text(text)
        if "당기공제세액" in flat.replace(" ", "") and "세액공제조정명세서" in flat.replace(" ", ""):
            data["_found_page"] = pg
            break

    # 2차: 테이블 추출 방식 (한글 깨짐 대비)
    _doc = fitz_doc
    if _doc is None and pdf_path:
        try:
            import fitz as _fitz
            _doc = _fitz.open(pdf_path)
        except Exception:
            pass
    if _doc:
        try:
            # 후보 페이지 찾기: 텍스트에서 16,480,000 같은 큰 금액+16컬럼 테이블이 있는 페이지
            # 또는 "(107)", "(120)" 등이 있는 페이지
            candidate_pages = []
            for pg, text in pages.items():
                flat = text.replace(" ", "")
                # 세액공제조정명세서(3) 관련 키워드
                if ("부표3" in flat or "부표" in flat or "(107)" in flat or "(120)" in flat
                    or "당기공제세액" in flat or "이월액계산" in flat):
                    candidate_pages.append(pg - 1)  # pages는 1-indexed, fitz는 0-indexed
            if not candidate_pages:
                candidate_pages = list(range(min(30, len(_doc))))

            for pg_num in candidate_pages:
                if pg_num < 0 or pg_num >= len(_doc):
                    continue
                page_obj = _doc[pg_num]
                tabs = page_obj.find_tables()
                for tab in tabs.tables:
                    rows = tab.extract()
                    if len(rows) < 3 or len(rows[0]) < 14:
                        continue
                    # 세액공제조정명세서(3) 식별: 16컬럼 테이블, 헤더에 (105), (107), (120), (123) 등
                    header_str = str(rows[0]).replace(" ", "")
                    if "(107)" not in header_str and "(120)" not in header_str:
                        if len(rows) > 1:
                            header_str = str(rows[1]).replace(" ", "")
                        if "(107)" not in header_str and "(120)" not in header_str:
                            continue

                    # 컬럼 인덱스 파악: (107)당기분, (120)계, (121)최저한세, (123)공제세액, (125)이월액
                    # 16컬럼 구조: [0]구분, [1]사업연도, [2](107)당기분, [3](108)이월분, [4](109)당기분,
                    #   [5-9](110~114), [10](120)계, [11](121)최저한세, [12](122)그밖, [13](123)공제세액, [14](124)소멸, [15](125)이월액
                    col_당기분 = 2   # (107)
                    col_계 = 10      # (120)
                    col_최저한세 = 11 # (121)
                    col_공제세액 = 13 # (123)
                    col_이월액 = 15   # (125)

                    def safe_int(val):
                        if not val or not isinstance(val, str):
                            return None
                        # ☞ 등 특수기호 제거
                        import re as _re
                        v = _re.sub(r'[^\d,\-]', '', val)
                        v = v.replace(",", "")
                        if not v:
                            return None
                        neg = v.startswith("-")
                        if neg:
                            v = v[1:]
                        if not v.isdigit():
                            return None
                        return -int(v) if neg else int(v)

                    # 데이터 행 파싱 (헤더 이후)
                    당기_rows = []
                    소계_row = None
                    합계_row = None
                    last_data_row = None
                    for row in rows[2:]:
                        if not row:
                            continue
                        first_cell = (row[0] or "").strip().replace(" ", "")
                        # 소계/합계 식별 (한글 깨짐 대비: 마지막 데이터 행도 추적)
                        if "소계" in first_cell:
                            소계_row = row
                        elif "합계" in first_cell:
                            합계_row = row
                        else:
                            # 당기분 데이터 행 (금액이 있는 행)
                            val = safe_int(row[col_당기분] if len(row) > col_당기분 else None)
                            if val is not None and val > 0:
                                당기_rows.append(row)
                            # (107)에 금액 있는 마지막 행 추적 → 합계 후보
                            v_계 = safe_int(row[col_계] if len(row) > col_계 else None)
                            if v_계 is not None and v_계 > 0:
                                last_data_row = row

                    # 합계 행이 없으면 마지막 데이터 행을 합계로 사용
                    if not 합계_row and last_data_row and last_data_row not in 당기_rows[:1]:
                        합계_row = last_data_row

                    # 당기분 행 (첫 번째 또는 유일한 데이터 행)
                    if 당기_rows:
                        row = 당기_rows[0]
                        v = safe_int(row[col_당기분] if len(row) > col_당기분 else None)
                        if v is not None:
                            data["당기분_당기분"] = v
                        v = safe_int(row[col_공제세액] if len(row) > col_공제세액 else None)
                        if v is not None:
                            data["당기분_공제세액"] = v

                    # 소계 행
                    if 소계_row:
                        for key, col in [("소계_당기분", col_당기분), ("소계_계", col_계),
                                         ("소계_최저한세", col_최저한세), ("소계_공제세액", col_공제세액),
                                         ("소계_이월액", col_이월액)]:
                            v = safe_int(소계_row[col] if len(소계_row) > col else None)
                            if v is not None:
                                data[key] = v

                    # 합계 행
                    if 합계_row:
                        for key, col in [("합계_당기분", col_당기분), ("합계_계", col_계),
                                         ("합계_최저한세", col_최저한세), ("합계_공제세액", col_공제세액),
                                         ("합계_이월액", col_이월액)]:
                            v = safe_int(합계_row[col] if len(합계_row) > col else None)
                            if v is not None:
                                data[key] = v

                    if data.get("합계_당기분") or data.get("당기분_당기분"):
                        break
                if data.get("합계_당기분") or data.get("당기분_당기분"):
                    break
        except Exception:
            pass

    # _found_page는 내부용이므로 제거
    data.pop("_found_page", None)
    return data


# ============================================================
# 세액공제신청서 (별지제1호서식 별지)
# ============================================================
def parse_tax_credit_application(pages, pdf_path=None, fitz_doc=None):
    """세액공제신청서 - 구분/코드/대상세액/공제세액
    테이블 추출 방식 (한글 깨짐 대비)"""
    data = {"항목": [], "대상세액_합계": 0, "공제세액_합계": 0}

    # 1차: 키워드 기반 (한글이 안 깨진 PDF)
    start_pg = None
    for pg, text in pages.items():
        flat = join_text(text)
        if "세액공제신청서" in flat.replace(" ", ""):
            start_pg = pg
            break

    if start_pg is not None:
        scan_pages = [p for p in sorted(pages.keys()) if start_pg <= p <= start_pg + 4]
        for pg in scan_pages:
            text = pages[pg]
            lines = text.strip().split("\n")
            i = 0
            while i < len(lines):
                stripped = lines[i].strip()
                if re.match(r'^[0-9A-Za-z]{2,3}$', stripped):
                    code = stripped
                    nums = []
                    for j in range(i + 1, min(i + 4, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            nums.append(int(m.group(1).replace(",", "")))
                        else:
                            break
                    if nums:
                        item = {"코드": code}
                        if len(nums) >= 2:
                            item["대상세액"] = nums[0]
                            item["공제세액"] = nums[1]
                        elif len(nums) == 1:
                            item["대상세액"] = nums[0]
                        data["항목"].append(item)
                if stripped == "1A3" or ("세액공제합계" in stripped.replace(" ", "")):
                    nums = []
                    for j in range(i + 1, min(i + 4, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            nums.append(int(m.group(1).replace(",", "")))
                        else:
                            break
                    if len(nums) >= 2:
                        data["대상세액_합계"] = nums[0]
                        data["공제세액_합계"] = nums[1]
                    elif len(nums) == 1:
                        data["대상세액_합계"] = nums[0]
                i += 1

    # 2차: 테이블 추출 (한글 깨짐 대비)
    _doc = fitz_doc
    if _doc is None and pdf_path:
        try:
            import fitz as _fitz
            _doc = _fitz.open(pdf_path)
        except Exception:
            pass
    if not data.get("대상세액_합계") and _doc:
        try:
            # 후보 페이지: 세액공제신청서 관련 키워드
            candidate_pages = []
            for pg, text in pages.items():
                flat = text.replace(" ", "")
                if ("신청서" in flat or "대상세액" in flat or "공제세액" in flat
                    or "18S" in text or "14N" in text or "1A3" in text):
                    candidate_pages.append(pg - 1)
            if not candidate_pages:
                candidate_pages = list(range(min(30, len(_doc))))

            for pg_num in candidate_pages:
                if pg_num < 0 or pg_num >= len(_doc):
                    continue
                page_obj = _doc[pg_num]
                tabs = page_obj.find_tables()
                for tab in tabs.tables:
                    rows = tab.extract()
                    if len(rows) < 10 or len(rows[0]) < 5:
                        continue
                    # 세액공제신청서 식별: 6컬럼, 코드 컬럼에 14B/18S/14N 등
                    codes_found = set()
                    for row in rows:
                        code = (row[2] if len(row) > 2 and row[2] else "").strip()
                        if re.match(r'^[0-9A-Za-z]{2,3}$', code):
                            codes_found.add(code)
                    if not codes_found.intersection({"18S", "14N", "14J", "18A", "14H", "18G", "1B4", "1B5", "14E", "14B"}):
                        continue

                    def safe_int(val):
                        if not val or not isinstance(val, str):
                            return None
                        v = val.strip().replace(",", "")
                        if not v.isdigit():
                            return None
                        return int(v)

                    # 대상세액=col[4], 공제세액=col[5]
                    for row in rows:
                        code = (row[2] if len(row) > 2 and row[2] else "").strip()
                        대상 = safe_int(row[4] if len(row) > 4 else None)
                        공제 = safe_int(row[5] if len(row) > 5 else None)
                        if code == "1A3" or "합계" in str(row[0] or ""):
                            if 대상:
                                data["대상세액_합계"] = 대상
                            if 공제:
                                data["공제세액_합계"] = 공제
                        elif 대상 and re.match(r'^[0-9A-Za-z]{2,3}$', code):
                            item = {"코드": code, "대상세액": 대상}
                            if 공제:
                                item["공제세액"] = 공제
                            data["항목"].append(item)
                    if data.get("대상세액_합계"):
                        break
                if data.get("대상세액_합계"):
                    break
        except Exception:
            pass

    return data


# ============================================================
# 소득금액조정합계표
# ============================================================
def parse_income_adjustment(pages):
    data = {"익금산입_항목": [], "손금산입_항목": [], "익금산입_합계": 0, "손금산입_합계": 0}
    target_pg = find_page_by_keyword(pages, "소득금액조정합계표")
    if not target_pg:
        return data

    text = pages[target_pg]
    lines = text.strip().split("\n")

    # 항목 파싱: "과목명" → "금액 처분" → "코드" 패턴
    items = []
    합계_values = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # 합계 행
        if line == "합계":
            # 다음 줄에서 금액 추출
            for j in range(i + 1, min(i + 3, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    합계_values.append(int(m.group(1).replace(",", "")))
                    break
            i += 1
            continue

        # "금액 처분" 패턴 매칭 (예: "107,172 유보" 또는 "5,790,197 유보")
        m = re.match(r'^([\d,]+)\s*(유보|사외유출|기타사외유출|기타)', line)
        if m and items:
            val = int(m.group(1).replace(",", ""))
            처분 = m.group(2)
            items[-1]["금액"] = val
            items[-1]["처분"] = 처분
            # 다음 줄이 코드번호
            if i + 1 < len(lines) and re.match(r'^\d{3}$', lines[i + 1].strip()):
                items[-1]["코드"] = lines[i + 1].strip()
                i += 1
            i += 1
            continue

        # 과목명 (한글 포함, 최소 3자)
        if len(line) >= 3 and re.search(r'[가-힣]', line) and "법인세법" not in line and "서식" not in line and "사업" not in line and "법인명" not in line and "사업자" not in line and "소득금액" not in line and "익금산입" not in line and "손금산입" not in line and "처분" not in line and "코드" not in line and "과" != line and "목" != line:
            items.append({"과목": line.replace(" ", ""), "금액": 0, "처분": "", "코드": ""})

        i += 1

    # 합계값으로 익금/손금 구분
    # items 금액 합산해서 첫 번째 합계와 매칭되는 그룹이 익금산입
    if 합계_values:
        data["익금산입_합계"] = 합계_values[0] if 합계_values else 0
        data["손금산입_합계"] = 합계_values[1] if len(합계_values) > 1 else 0

    # 항목을 익금/손금으로 분류 (합계값으로 역추적)
    valid_items = [item for item in items if item["금액"] > 0]
    if data["익금산입_합계"] > 0:
        # 부분합이 익금산입_합계가 되는 조합 찾기
        running = 0
        for item in valid_items:
            running += item["금액"]
            data["익금산입_항목"].append(item)
            if running >= data["익금산입_합계"]:
                break
        # 나머지는 손금산입
        remaining = valid_items[len(data["익금산입_항목"]):]
        data["손금산입_항목"] = remaining
    else:
        data["익금산입_항목"] = valid_items

    return data


# ============================================================
# 감가상각비조정명세서합계표
# ============================================================
def parse_depreciation(pages):
    data = {}
    target_pg = find_page_by_keyword(pages, "감가상각비조정명세서합계표")
    if not target_pg:
        return data

    text = pages[target_pg]
    code_amounts = {}
    lines = text.strip().split("\n")
    for i, line in enumerate(lines):
        if re.match(r'^0[1-9]$', line.strip()):
            code = int(line.strip())
            for j in range(i + 1, min(i + 5, len(lines))):
                # 한 줄에 여러 금액이 있을 수 있음 (합계 건축물 기계장치...)
                # 첫 번째 금액이 합계
                nums = re.findall(r'[\d,]{4,}', lines[j])
                if nums:
                    val = int(nums[0].replace(",", ""))
                    if val >= 1000:
                        code_amounts[code] = val
                        break
                elif lines[j].strip() and not re.match(r'^\d{1,2}$', lines[j].strip()):
                    # 텍스트 줄이면 건너뛰되 다음 줄 계속 확인
                    continue

    data["기말현재액"] = code_amounts.get(1)
    data["감가상각누계액"] = code_amounts.get(2)
    data["미상각잔액"] = code_amounts.get(3)
    data["상각범위액"] = code_amounts.get(4)
    data["회사손금계상액"] = code_amounts.get(5)

    # 기말현재액 행에서 건축물/기계장치/기타자산 소계 파싱
    # 패턴: 01 다음에 "합계 건축물" (한 줄), "기계장치" (다음 줄), "기타자산" (다음 줄)
    for i, line in enumerate(lines):
        if line.strip() == "01":
            all_vals = []
            for j in range(i + 1, min(i + 5, len(lines))):
                nums = re.findall(r'[\d,]{4,}', lines[j])
                for n in nums:
                    v = int(n.replace(",", ""))
                    if v >= 1000:
                        all_vals.append(v)
                # 다음 코드(02)가 나오면 중단
                if re.match(r'^0[2-9]$', lines[j].strip()):
                    break
            # all_vals: [합계, 건축물, 기계장치, 기타자산]
            if len(all_vals) >= 2:
                data["기말_건축물"] = all_vals[1]
            if len(all_vals) >= 3:
                data["기말_기계장치"] = all_vals[2]
            if len(all_vals) >= 4:
                data["기말_기타자산"] = all_vals[3]
            break

    # 양도자산 기말현재액 합산: 미상각분감가상각조정명세서 중 [양도자산] 페이지
    양도_기말 = 0
    for pg, text in pages.items():
        if "양도자산" in text:
            # "총계" 또는 첫 번째 (5)기말현재액 다음 금액
            page_lines = text.strip().split("\n")
            found_total = False
            for k, pl in enumerate(page_lines):
                if pl.strip() == "총계":
                    found_total = True
                    continue
                if found_total:
                    nums = re.findall(r'[\d,]{4,}', pl)
                    if nums:
                        양도_기말 = int(nums[0].replace(",", ""))
                        break
    if 양도_기말 > 0:
        data["양도자산_기말현재액"] = 양도_기말

    return data


# ============================================================
# 업무용승용차관련비용명세서
# ============================================================
def parse_vehicle_expenses(pages):
    data = {}
    target_pg = find_page_by_keyword(pages, "업무용승용차관련비용명세서")
    if not target_pg:
        return data

    text = pages[target_pg]
    flat = join_text(text)

    m = re.search(r'(\d{2,3}[가-힣]\d{4})', flat)
    if m:
        data["차량번호"] = m.group(1)

    m = re.search(r'(BMW|벤츠|아우디|포르쉐|테슬라|기아|현대|제네시스)\S*', flat)
    if m:
        data["차종"] = m.group(0)

    nums = extract_large_numbers(flat, 10_000_000)
    if nums:
        data["취득가액"] = max(nums)

    amounts = extract_large_numbers(text, 100000)
    for a in amounts:
        if 5000000 < a < 10000000:
            if "감가상각비" not in data:
                data["감가상각비"] = a
            elif a != data["감가상각비"]:
                data["한도초과금액"] = a
        if 1000000 < a < 5000000:
            data["한도내금액"] = a

    return data


# ============================================================
# 기업업무추진비(접대비) 조정명세서
# ============================================================
def parse_entertainment_expense(pages):
    data = {}
    for pg, text in pages.items():
        flat = join_text(text)
        if "기업업무추진비조정명세서(갑)" in flat or ("기업업무추진비조정명세서" in flat and "갑" in flat):
            m = re.search(r'기업업무추진비해당금액\s*([\d,]+)', flat)
            if m:
                data["기업업무추진비해당금액"] = int(m.group(1).replace(",", ""))
            m = re.search(r'기업업무추진비한도액합계.*?([\d,]{5,})', flat)
            if m:
                data["한도액합계"] = int(m.group(1).replace(",", ""))
            break

    for pg, text in pages.items():
        flat = join_text(text)
        if "기업업무추진비조정명세서(을)" in flat or ("기업업무추진비조정명세서" in flat and "을" in flat and "수입금액명세" in flat):
            m = re.search(r'일반수입금액\s*([\d,]+)', flat)
            if not m:
                m = re.search(r'금\s*액\s*([\d,]{7,})', flat)
            if m:
                data["수입금액"] = int(m.group(1).replace(",", ""))
            m = re.search(r'접대비.*?(\d[\d,]+)', flat)
            if m:
                data["접대비계상액"] = int(m.group(1).replace(",", ""))
            break

    return data


# ============================================================
# 표준재무상태표 (별지제3호의2서식)
# ============================================================
def parse_colon_amount(lines, start_idx, search_range=5):
    """콜론 구분 금액 추출 (표준재무상태표/손익계산서 공통)
    예: ':', ':', '3:736:854:578' → 3736854578
    """
    for j in range(start_idx, min(start_idx + search_range, len(lines))):
        line = lines[j].strip()
        # 콜론 구분 숫자 패턴 (1자리 이상 숫자:숫자... 또는 공백+숫자)
        m = re.search(r'(\d[\d:]+)', line)
        if m:
            raw = m.group(1).replace(":", "")
            if raw.isdigit() and int(raw) >= 1000:
                return int(raw)
    return None


def parse_standard_balance_sheet(pages, pdf_path=None, fitz_doc=None):
    """표준재무상태표 - 코드:금액 형식 파싱"""
    data = {}
    target_pages = find_all_pages_by_keyword(pages, "표준재무상태표")
    if not target_pages:
        return data

    keyword_map = {
        "자산총계": "자산총계",
        "부채총계": "부채총계",
        "자본총계": "자본총계",
        "부채와자본총계": "부채와자본총계",
        "자본금": "자본금",
        "이익잉여금": "이익잉여금",
        "미처분이익잉여금": "미처분이익잉여금",
        "유형자산": "유형자산_합계",
        "무형자산": "무형자산_합계",
    }

    # 유형자산 개별항목 (차변합계 산출용)
    # 토지, 건설중인유형자산은 감가상각 대상 아니므로 제외
    tangible_items = ["건물", "구축물", "기계장치", "선박", "차량운반구",
                      "항공기", "건설용장비", "공구", "비품", "시설장치", "기타유형자산"]
    # 차감항목 (제외)
    exclude_keywords = ["감가상각", "손상차손", "정부보조금", "대손충당금"]

    tangible_debit_total = 0  # 유형자산 차변합계
    intangible_debit_total = 0  # 무형자산 차변합계
    in_tangible = False
    in_intangible = False

    all_lines = []
    for pg in sorted(target_pages):
        page_lines = pages[pg].strip().split("\n")
        all_lines.extend(page_lines)

    for i, line in enumerate(all_lines):
        flat = line.strip().replace(" ", "")

        # 섹션 추적
        if "(2)유형자산" in flat.replace(" ", ""):
            in_tangible = True
            in_intangible = False
        elif "(3)무형자산" in flat.replace(" ", ""):
            in_tangible = False
            in_intangible = True
        elif "자산총계" in flat:
            in_tangible = False
            in_intangible = False

        # 유형자산 개별항목 차변금액 수집
        # "비품외" 등 다른 항목에 부분매칭 방지
        tangible_exclude_items = ["비품외", "토지", "건설중"]
        if in_tangible and not any(ex in flat for ex in exclude_keywords) and not any(ex in flat for ex in tangible_exclude_items):
            for item_name in tangible_items:
                if item_name in flat:
                    val = parse_colon_amount(all_lines, i + 1, 8)
                    if val:
                        tangible_debit_total += val
                    break

        # 무형자산 개별항목 차변금액 수집 (회원권은 상각 대상 아니므로 제외)
        intangible_items = ["영업권", "산업재산권", "개발비", "기타무형자산"]
        intangible_exclude = ["회원권"]
        if in_intangible and not any(ex in flat for ex in exclude_keywords) and not any(ex in flat for ex in intangible_exclude):
            for item_name in intangible_items:
                if item_name in flat:
                    val = parse_colon_amount(all_lines, i + 1, 8)
                    if val:
                        intangible_debit_total += val
                    break

        # 기존 키워드 매칭
        for keyword, field in keyword_map.items():
            if keyword in flat and field not in data:
                if keyword == "유형자산" and "감가상각" in flat:
                    continue
                if keyword == "유형자산" and any(k in flat for k in ["건설중", "투자", "무형"]):
                    continue
                if keyword == "자본금" and any(k in flat for k in ["보통주", "우선주"]):
                    continue
                if keyword == "이익잉여금" and "미처분" in flat:
                    continue
                val = parse_colon_amount(all_lines, i + 1, 8)
                if val:
                    data[field] = val

    if tangible_debit_total > 0:
        data["유형자산_차변합계"] = tangible_debit_total
    if intangible_debit_total > 0:
        data["무형자산_차변합계"] = intangible_debit_total

    # 좌표 기반 합계표준재무상태표 차변합계 파싱
    _doc = fitz_doc
    if _doc is None and pdf_path:
        try:
            import fitz as _fitz
            _doc = _fitz.open(pdf_path)
        except Exception:
            pass
    if _doc:
        try:
            debit_data = _parse_bs_debit_by_coords(_doc, pages)
            if debit_data:
                data.update(debit_data)
        except Exception:
            pass

    return data


def _parse_bs_debit_by_coords(doc, pages):
    """합계표준재무상태표에서 좌표 기반으로 차변합계 추출"""
    data = {}

    # 합계표준재무상태표 페이지 찾기
    bs_pages = []
    for pg, text in pages.items():
        flat = text.replace(" ", "").replace("\n", "")
        if "합계표준재무상태표" in flat or ("표준재무상태표" in flat and "차변합계" in flat):
            bs_pages.append(pg - 1)  # pages dict is 1-indexed, fitz is 0-indexed

    if not bs_pages:
        return data

    # 모든 span 수집
    all_spans = []
    for pg_idx in bs_pages:
        if pg_idx < 0 or pg_idx >= len(doc):
            continue
        page = doc[pg_idx]
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    text = s["text"].strip()
                    if text:
                        all_spans.append({
                            "text": text,
                            "x": s["bbox"][0],
                            "y": round(s["bbox"][1]),
                            "page": pg_idx
                        })

    # 계정과목별 y좌표와 item_x 찾기
    key_items = {}
    for s in all_spans:
        flat = s["text"].replace(" ", "")
        if "(2)유형자산" in flat and "건설중" not in flat and "유형자산" not in key_items:
            key_items["유형자산"] = (s["y"], s["page"], s["x"])
        if re.match(r'^1\.\s*토지$', s["text"].strip()) and "토지" not in key_items:
            key_items["토지"] = (s["y"], s["page"], s["x"])
        if "(3)무형자산" in flat and "무형자산" not in key_items:
            key_items["무형자산"] = (s["y"], s["page"], s["x"])
        if "회원권" in flat and re.match(r'^\d+\.', flat) and "회원권" not in key_items:
            key_items["회원권"] = (s["y"], s["page"], s["x"])

    def find_debit(target_y, target_page, item_x, tolerance=3):
        """계정과목 왼쪽의 span들을 합쳐서 차변합계 추출"""
        row_spans = [s for s in all_spans
                     if s["page"] == target_page
                     and abs(s["y"] - target_y) <= tolerance
                     and s["x"] < item_x]
        row_spans.sort(key=lambda s: s["x"])
        combined = "".join(s["text"] for s in row_spans).lstrip(":").strip()
        m = re.search(r"(\d[\d:]*)", combined)
        if m:
            raw = m.group(1).replace(":", "")
            if raw.isdigit() and int(raw) >= 1000:
                return int(raw)
        return 0

    if "유형자산" in key_items:
        data["유형자산_차변_좌표"] = find_debit(*key_items["유형자산"])
    if "토지" in key_items:
        data["토지_차변_좌표"] = find_debit(*key_items["토지"])
    if "무형자산" in key_items:
        data["무형자산_차변_좌표"] = find_debit(*key_items["무형자산"])
    if "회원권" in key_items:
        data["회원권_차변_좌표"] = find_debit(*key_items["회원권"])

    return data


# ============================================================
# 표준손익계산서
# ============================================================
def parse_standard_income_statement(pages):
    """표준손익계산서 - 감가상각비, 법인세비용 등 추출"""
    data = {}
    target_pages = find_all_pages_by_keyword(pages, "표준손익계산서")

    keyword_map = {
        "유형자산감가상각비": "유형자산감가상각비",
        "무형자산상각비": "무형자산상각비",
        "당기순이익": "당기순이익",
    }

    for pg in target_pages:
        lines = pages[pg].strip().split("\n")
        for i, line in enumerate(lines):
            flat_line = line.replace(" ", "")

            for keyword, field in keyword_map.items():
                if keyword in flat_line and field not in data:
                    val = parse_colon_amount(lines, i + 1, 8)
                    if val:
                        data[field] = val

            # 법인세비용 (법인세차감전 제외)
            if "법인세비용" in flat_line and "차감전" not in flat_line and "법인세비용" not in data:
                val = parse_colon_amount(lines, i + 1, 10)
                if val:
                    data["법인세비용"] = val

            # 이자수익
            if "이자수익" in flat_line and "이자수익" not in data:
                val = parse_colon_amount(lines, i + 1, 8)
                if val:
                    data["이자수익"] = val

    return data


# ============================================================
# 결산보고서 재무상태표 + 손익계산서
# ============================================================
def parse_financial_statements(pages):
    data = {"재무상태표": {}, "손익계산서": {}}

    # 결산보고서 재무상태표 (서식명 기반 탐색)
    for pg, text in pages.items():
        if pg < 4:
            continue
        flat = join_text(text)
        if "재무상태표" in flat and "회사명" in flat and "유동자산" in flat.replace(" ", ""):
            lines = text.strip().split("\n")

            # 연속된 줄을 합쳐서 키워드 찾기 (자/산/총/계 같이 분리된 경우)
            # 금액 라인과 그 이전 4줄을 합쳐서 키워드 검색
            for i, line in enumerate(lines):
                m = re.match(r'^([\d,]+)$', line.strip())
                if m:
                    val = int(m.group(1).replace(",", ""))
                    if val >= 100000:
                        # 앞 10줄을 모두 합쳐서 키워드 검색
                        prev_text = "".join(lines[k].strip() for k in range(max(0, i - 10), i))
                        # 가장 가까운 키워드 매칭 (뒤에서부터)
                        if "자본총계" in prev_text and "자본총계" not in data["재무상태표"]:
                            data["재무상태표"]["자본총계"] = val
                        elif "부채총계" in prev_text and "자산총계" in data["재무상태표"] and "부채총계" not in data["재무상태표"]:
                            data["재무상태표"]["부채총계"] = val
                        elif "자산총계" in prev_text and "자산총계" not in data["재무상태표"]:
                            data["재무상태표"]["자산총계"] = val

            # 섹션별 금액 추출 (자본금, 자본잉여금, 이익잉여금)
            # 당기/전기 2열 구조: 금액이 연속 2개 나오면 첫 번째가 당기
            section_map = {
                "Ⅰ.자본금": "자본금",
                "Ⅱ.자본잉여금": "자본잉여금",
                "미처분이익잉여금": "이익잉여금",
            }
            current_section = None
            for i, line in enumerate(lines):
                flat_line = "".join(lines[k].strip() for k in range(max(0, i - 5), i + 1))
                for keyword, field in section_map.items():
                    if keyword in flat_line and field not in data["재무상태표"]:
                        current_section = field
                m = re.match(r'^([\d,]+)$', line.strip())
                if m and current_section:
                    val = int(m.group(1).replace(",", ""))
                    if val >= 10000 and current_section not in data["재무상태표"]:
                        data["재무상태표"][current_section] = val
                        current_section = None

            # 유형자산
            for i, line in enumerate(lines):
                flat_line = "".join(lines[k].strip() for k in range(max(0, i - 3), i + 1))
                if "유형자산" in flat_line and "감가상각" not in flat_line:
                    for j in range(i + 1, min(i + 5, len(lines))):
                        m = re.match(r'^([\d,]+)$', lines[j].strip())
                        if m:
                            val = int(m.group(1).replace(",", ""))
                            if val >= 1000:
                                data["재무상태표"]["유형자산"] = val
                                break
                    break
            break

    # 결산보고서 손익계산서 (서식명 기반 탐색, 당기/전기 2열)
    data["손익계산서_전기"] = {}
    for pg, text in pages.items():
        if pg < 4:
            continue
        flat = join_text(text)
        # 결산보고서 손익계산서: "손익계산서" + "회사명" + ("매출" 또는 "판매비")
        is_income_stmt = ("손익계산서" in flat and "회사명" in flat and
                         "매출" in flat.replace(" ", ""))
        if not is_income_stmt:
            continue
        # 표준손익계산서 제외 (별지제3호 서식)
        if "별지제3호" in flat or "표준손익계산서" in flat:
            continue

        lines = text.strip().split("\n")
        # 주요 계정과목 매핑 (분리된 글자 합치기)
        key_map = {
            "매출액": "매출액", "매출원가": "매출원가", "매출총이익": "매출총이익",
            "판매비와관리비": "판관비", "영업이익": "영업이익",
            "영업외수익": "영업외수익", "영업외비용": "영업외비용",
            "법인세차감전이익": "법인세차감전이익",
            "법인세등": "법인세등", "당기순이익": "당기순이익",
        }
        # 2열 구조: 키워드 매칭 후 금액 2개 연속 (당기, 전기)
        i = 0
        while i < len(lines):
            # 현재 줄부터 뒤로 5줄까지 합쳐서 키워드 매칭
            combined = ""
            for k in range(i, min(i + 6, len(lines))):
                combined += lines[k].strip().replace(" ", "")
                for keyword, field in key_map.items():
                    if keyword in combined and field not in data["손익계산서"]:
                        # 금액 2개 찾기 (당기, 전기)
                        amounts = []
                        for j in range(k + 1, min(k + 10, len(lines))):
                            m = re.match(r'^-?([\d,]+)$', lines[j].strip())
                            if m:
                                amounts.append(int(m.group(0).replace(",", "")))
                                if len(amounts) == 2:
                                    break
                            # 다른 키워드 나오면 중단
                            ln_flat = lines[j].strip().replace(" ", "")
                            if any(kw in ln_flat for kw in key_map) and len(ln_flat) > 2:
                                break
                        if amounts:
                            data["손익계산서"][field] = amounts[0]
                            if len(amounts) >= 2:
                                data["손익계산서_전기"][field] = amounts[1]
                        i = k
                        break
                else:
                    continue
                break
            i += 1

        # 판관비 세부항목 (급여, 퇴직급여, 감가상각비 등)
        detail_map = {
            "급여": "급여", "퇴직급여": "퇴직급여", "복리후생비": "복리후생비",
            "감가상각비": "감가상각비", "지급임차료": "지급임차료",
            "보험료": "보험료", "지급수수료": "지급수수료",
            "광고선전비": "광고선전비", "외주용역비": "외주용역비",
            "접대비": "접대비",
        }
        if "판관비_세부" not in data:
            data["판관비_세부"] = {}
            data["판관비_세부_전기"] = {}
        for idx, line in enumerate(lines):
            flat_line = "".join(lines[k2].strip() for k2 in range(max(0, idx - 4), idx + 1)).replace(" ", "")
            for keyword, field in detail_map.items():
                if keyword in flat_line and field not in data["판관비_세부"]:
                    amounts = []
                    for j in range(idx + 1, min(idx + 5, len(lines))):
                        m = re.match(r'^-?([\d,]+)$', lines[j].strip())
                        if m:
                            amounts.append(int(m.group(0).replace(",", "")))
                            if len(amounts) == 2:
                                break
                        elif lines[j].strip() and not re.match(r'^[\d,]+$', lines[j].strip()):
                            break
                    if amounts:
                        data["판관비_세부"][field] = amounts[0]
                        if len(amounts) >= 2:
                            data["판관비_세부_전기"][field] = amounts[1]
        break

    return data


# ============================================================
# 소득구분계산서 (별지제48호서식)
# ============================================================
def parse_income_classification(pages):
    """소득구분계산서 - 감면대상소득, 기타분, 매출비율 등"""
    data = {}
    target_pg = find_page_by_keyword(pages, "소득구분계산서")
    if not target_pg:
        return data

    text = pages[target_pg]
    flat = join_text(text)
    lines = text.strip().split("\n")

    # 코드번호 기반 추출 (01=매출액, 03=매출총이익, 07=영업이익, 21=소득, 25=과세표준)
    code_amounts = extract_code_amounts(text)

    data["매출액"] = code_amounts.get(1)
    data["매출총이익"] = code_amounts.get(3)
    data["영업이익"] = code_amounts.get(7)
    data["각사업연도소득"] = code_amounts.get(21)
    data["과세표준"] = code_amounts.get(25)

    # 감면분/기타분 금액 추출
    # "감면분또는" 키워드 이후 업종명과 금액
    amounts = extract_line_amounts(text, 1000)

    # 감면대상소득 = (12)과세표준의 감면분 금액
    # PDF에서 과세표준 행(코드25)의 감면분 금액을 찾기
    # 패턴: 247,607,266 다음에 247,290,024 그리고 317,242
    if data.get("과세표준"):
        과세표준 = data["과세표준"]
        # 과세표준 뒤에 나오는 금액들이 감면분/기타분
        found_과세표준 = False
        sub_amounts = []
        for line in lines:
            m = re.match(r'^([\d,]+)$', line.strip())
            if m:
                val = int(m.group(1).replace(",", ""))
                if val == 과세표준:
                    found_과세표준 = True
                    continue
                if found_과세표준 and val < 과세표준 and val >= 1000:
                    sub_amounts.append(val)
                    if len(sub_amounts) >= 2:
                        break

        if sub_amounts:
            data["감면대상소득"] = sub_amounts[0]
            if len(sub_amounts) >= 2:
                data["기타분소득"] = sub_amounts[1]
            # 검증: 감면 + 기타 = 과세표준
            합계 = sum(sub_amounts[:2]) if len(sub_amounts) >= 2 else sub_amounts[0]
            data["소득합계_일치"] = (합계 == 과세표준)

    # 업종명 추출
    m = re.search(r'(광고|제조|도소매|건설|운수|서비스|컨설팅)\S*', flat)
    if m:
        data["감면업종"] = m.group(0)

    return data


# ============================================================
# 자본금과적립금조정명세서 (갑/을)
# ============================================================
def parse_capital_reserves(pages, pdf_path=None, fitz_doc=None):
    data = {"이월결손금": 0, "유보소득_기말": 0, "갑7_을병계": 0}

    target_pg = find_page_by_keyword(pages, "자본금과적립금조정명세서(갑)")
    if not target_pg:
        return data

    text = pages[target_pg]
    flat = join_text(text)
    lines = text.strip().split("\n")

    # 결손금 확인
    in_deficit_section = False
    for i, line in enumerate(lines):
        flat_line = line.strip().replace(" ", "")
        if "결손금발생" in flat_line or "기공제액" in flat_line:
            in_deficit_section = True
        if in_deficit_section and flat_line == "계":
            for j in range(i + 1, min(i + 5, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    raw = m.group(1).replace(",", "")
                    if raw.isdigit() and int(raw) >= 100000:
                        data["이월결손금"] = int(raw)
                        break
            break

    # 갑 서식 각 항목 기말잔액 추출 (금액 순서 기반)
    # 순서: 자본금(기초/기말), 자본잉여금(기초/기말), 이익잉여금(기초/증가/기말), 6.계, 7.을병계, ...
    all_amounts = extract_line_amounts(text, 1000)

    # (을) 합계를 먼저 파싱해서 갑에서 찾기 (아래에서 을 파싱 후 매칭)
    data["갑_all_amounts"] = all_amounts

    # 갑 서식 기말잔액 amounts
    amounts = extract_line_amounts(text, 10000)
    data["갑_기말잔액_amounts"] = amounts

    # 갑 서식에서 자본금/이익잉여금/자본총계 - 테이블 추출 방식
    # 한글 깨짐 대비: 키워드 대신 테이블 구조(8컬럼, 코드 01/14/20)로 갑 페이지 식별
    _doc = fitz_doc
    if _doc is None and pdf_path:
        try:
            import fitz as _fitz
            _doc = _fitz.open(pdf_path)
        except Exception:
            pass
    if _doc:
        try:
            # 후보 페이지: 자본금과적립금조정명세서(갑) 관련
            candidate_pages = []
            for pg, text in pages.items():
                flat = text.replace(" ", "")
                if "자본금과적립금" in flat or "적립금조정명세서" in flat or "자본금계산서" in flat:
                    candidate_pages.append(pg - 1)
            # 키워드 못 찾으면 갑 서식은 보통 20~30번 페이지
            if not candidate_pages:
                candidate_pages = list(range(15, min(40, len(_doc))))

            for pg_num in candidate_pages:
                if pg_num < 0 or pg_num >= len(_doc):
                    continue
                page_obj = _doc[pg_num]
                tabs = page_obj.find_tables()
                for tab in tabs.tables:
                    rows = tab.extract()
                    if len(rows) < 10 or len(rows[0]) < 7:
                        continue
                    # 갑 서식 식별: 코드 01, 02, 15, 18, 14, 20, 21이 모두 있는 테이블
                    codes_found = set()
                    for row in rows:
                        code = row[2] if len(row) > 2 and row[2] else ""
                        code = code.strip() if code else ""
                        if code in ("01", "02", "14", "15", "18", "20", "21"):
                            codes_found.add(code)
                    if not codes_found.issuperset({"01", "14", "20"}):
                        continue
                    # 갑 서식 확인됨 - 기말잔액 추출
                    for row in rows:
                        code_col = row[2] if len(row) > 2 and row[2] else ""
                        code_col = code_col.strip() if code_col else ""
                        기말_val = row[6] if len(row) > 6 and row[6] else ""
                        if not 기말_val or not isinstance(기말_val, str):
                            continue
                        기말_num = 기말_val.strip().replace(",", "")
                        is_negative = 기말_num.startswith("-")
                        if is_negative:
                            기말_num = 기말_num[1:]
                        if not 기말_num.isdigit():
                            continue
                        기말_int = -int(기말_num) if is_negative else int(기말_num)
                        if code_col == "01":
                            data["갑_자본금"] = 기말_int
                        elif code_col == "14":
                            data["갑_이익잉여금"] = 기말_int
                        elif code_col == "20":
                            data["갑_자본총계"] = 기말_int
                    break
                if "갑_자본금" in data or "갑_자본총계" in data:
                    break
        except Exception:
            pass

    # 을 서식 - 유보소득 합계 + 개별 항목
    target_pg2 = find_page_by_keyword(pages, "자본금과적립금조정명세서(을)")
    if target_pg2:
        text2 = pages[target_pg2]
        lines2 = text2.strip().split("\n")
        for i, line in enumerate(lines2):
            if "합" in line.strip():
                for j in range(i + 1, min(i + 10, len(lines2))):
                    if "계" in lines2[j].strip():
                        합계_amounts = []
                        for k in range(j + 1, min(j + 5, len(lines2))):
                            m2 = re.match(r'^([\d,]+)$', lines2[k].strip())
                            if m2:
                                raw = m2.group(1).replace(",", "")
                                if raw.isdigit() and int(raw) >= 1000:
                                    합계_amounts.append(int(raw))
                        if 합계_amounts:
                            data["유보소득_기말"] = 합계_amounts[-1]
                        break

    # (갑) 7번 = (을) 합계 매칭: 갑 금액 목록에서 을 합계와 같은 값 찾기
    을합계 = data.get("유보소득_기말", 0)
    if 을합계 and 을합계 in data.get("갑_all_amounts", []):
        data["갑7_을병계"] = 을합계

    return data


# ============================================================
# 이익잉여금처분계산서
# ============================================================
def parse_profit_disposition(pages):
    data = {}
    target_pg = find_page_by_keyword(pages, "이익잉여금처분")
    if not target_pg:
        return data

    text = pages[target_pg]
    lines = text.strip().split("\n")
    for i, line in enumerate(lines):
        flat_line = line.strip().replace(" ", "")
        if "미처분이익잉여금" in flat_line and "미처분이익잉여금" not in data:
            for j in range(i + 1, min(i + 5, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    data["미처분이익잉여금"] = int(m.group(1).replace(",", ""))
                    break
        elif "당기순이익" in flat_line and "당기순이익" not in data:
            for j in range(i + 1, min(i + 5, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    data["당기순이익"] = int(m.group(1).replace(",", ""))
                    break
        elif "현금배당" in flat_line and "배당금" not in data:
            for j in range(i + 1, min(i + 5, len(lines))):
                m = re.match(r'^([\d,]+)$', lines[j].strip())
                if m:
                    data["배당금"] = int(m.group(1).replace(",", ""))
                    break

    return data


# ============================================================
# 농어촌특별세
# ============================================================
def parse_rural_special_tax(pages):
    data = {}
    target_pg = find_page_by_keyword(pages, "농어촌특별세과세표준및세액신고서")
    if not target_pg:
        # 조정계산서에서도 확인
        target_pg = find_page_by_keyword(pages, "농어촌특별세과세표준및세액조정계산서")
    if not target_pg:
        return data

    text = pages[target_pg]
    amounts = []
    for line in text.strip().split("\n"):
        m = re.match(r'^([\d,]+)$', line.strip())
        if m:
            raw = m.group(1).replace(",", "")
            if raw and raw.isdigit() and int(raw) >= 10000:
                amounts.append(int(raw))

    if amounts:
        unique_amounts = sorted(set(amounts), reverse=True)
        if len(unique_amounts) >= 2:
            data["과세표준"] = unique_amounts[0]
            data["산출세액"] = unique_amounts[1]
        elif len(unique_amounts) == 1:
            data["산출세액"] = unique_amounts[0]

    return data


# ============================================================
# 중소기업기준검토표
# ============================================================
def parse_sme_review(pages):
    data = {"중소기업해당": False}
    target_pg = find_page_by_keyword(pages, "중소기업기준검토표")
    if target_pg:
        data["중소기업해당"] = True
    return data


# ============================================================
# 통합고용세액공제
# ============================================================
def parse_employment_credit(pages):
    data = {}
    for pg, text in pages.items():
        flat = join_text(text)
        if "통합고용세액공제" in flat and "공제세액계산서" in flat:
            nums = extract_large_numbers(text, 100000)
            if nums:
                data["공제세액"] = nums[-1]
            break
    return data


# ============================================================
# 통합 파싱
# ============================================================
def parse_all(pdf_path):
    pages = extract_all_text(pdf_path)

    # fitz Document를 한 번만 열어서 테이블 추출 함수들에 공유
    import fitz as _fitz
    _doc = _fitz.open(pdf_path)

    result = {
        "파일경로": pdf_path,
        "총페이지수": len(pages),
        "회사정보": parse_company_info(pages),
        "세액조정": parse_tax_adjustment(pages),
        "최저한세": parse_minimum_tax(pages, pdf_path),
        "공제감면": parse_deduction_credits(pages),
        "세액공제조정": parse_tax_credit_adjustment(pages, pdf_path, _doc),
        "소득금액조정": parse_income_adjustment(pages),
        "감가상각": parse_depreciation(pages),
        "표준재무상태표": parse_standard_balance_sheet(pages, pdf_path, _doc),
        "표준손익계산서": parse_standard_income_statement(pages),
        "업무용승용차": parse_vehicle_expenses(pages),
        "접대비": parse_entertainment_expense(pages),
        "재무제표": parse_financial_statements(pages),
        "소득구분": parse_income_classification(pages),
        "자본금적립금": parse_capital_reserves(pages, pdf_path, _doc),
        "이익잉여금처분": parse_profit_disposition(pages),
        "농어촌특별세": parse_rural_special_tax(pages),
        "중소기업": parse_sme_review(pages),
        "고용세액공제": parse_employment_credit(pages),
        "세액공제신청서": parse_tax_credit_application(pages, pdf_path, _doc),
    }

    _doc.close()
    return result


if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("Usage: python pdf_parser.py <pdf_path>")
        sys.exit(1)

    data = parse_all(sys.argv[1])
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
