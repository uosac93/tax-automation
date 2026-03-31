"""
건보공단 지급통보서 PDF 파서
- 요양급여비용 지급통보서 (건강보험)
- 의료급여비용 지급통보서 (의료보호)

라인 구조 기반 파싱 — PyMuPDF get_text() 출력 순서에 의존
"""
import re
import fitz  # PyMuPDF


def extract_all_text(pdf_path, password=None):
    """PDF 전체 페이지의 텍스트를 추출"""
    doc = fitz.open(pdf_path)
    if doc.is_encrypted:
        if not password:
            doc.close()
            raise ValueError("PDF_PASSWORD_REQUIRED")
        if not doc.authenticate(password):
            doc.close()
            raise ValueError("PDF_PASSWORD_WRONG")
    pages = {}
    for i in range(doc.page_count):
        pages[i + 1] = doc[i].get_text()
    total_pages = doc.page_count
    doc.close()
    return pages, total_pages


def clean_number(text):
    """텍스트에서 숫자 추출 (콤마, 공백 제거)"""
    if not text:
        return 0
    cleaned = text.replace(",", "").replace(" ", "").replace("원", "").strip()
    negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = cleaned[1:-1]
        negative = True
    if cleaned.startswith("-") or cleaned.startswith("△"):
        cleaned = cleaned[1:]
        negative = True
    cleaned = re.sub(r'[^\d]', '', cleaned)
    if cleaned.isdigit():
        val = int(cleaned)
        return -val if negative else val
    return 0


def _get_number_at(lines, idx):
    """특정 라인의 숫자값 추출"""
    if 0 <= idx < len(lines):
        return clean_number(lines[idx].strip())
    return 0


def _find_line_index(lines, keyword):
    """키워드가 포함된 라인 인덱스 찾기"""
    for i, line in enumerate(lines):
        if keyword in line:
            return i
    return -1


def _find_date_in_range(lines, start, end):
    """라인 범위에서 YYYY-MM-DD 날짜 찾기"""
    for i in range(start, min(end, len(lines))):
        m = re.match(r'^\s*(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})\s*$', lines[i].strip())
        if m:
            return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    return ''


def _find_last_payment(lines):
    """
    페이지 끝에서 실지급액 찾기.
    footer 영역(사업자번호, 공단, ※ 등)을 제외하고
    마지막 유효한 큰 숫자를 찾음.
    """
    for i in range(len(lines) - 1, max(len(lines) - 20, 0), -1):
        line = lines[i].strip()
        # footer 제외
        if any(kw in line for kw in ['105-82', '사업자', '공단사업자']):
            continue
        if line.startswith('(') or line.startswith('※'):
            continue
        if '원천징수' in line or '목적' in line or '144조' in line:
            continue
        # 순수 숫자 라인만
        cleaned = line.replace(',', '')
        if re.match(r'^\d+$', cleaned):
            val = int(cleaned)
            if val > 1000:
                return val
    return 0


def detect_pdf_type(page_text):
    """첫 페이지 텍스트로 PDF 유형 감지"""
    first_line = page_text.strip().split('\n')[0] if page_text.strip() else ''
    if '의료급여' in first_line:
        return 'medical_aid'
    elif '요양급여' in first_line:
        return 'nhis'
    if '의료급여' in page_text[:500]:
        return 'medical_aid'
    return 'nhis'


def parse_medical_aid_page(lines):
    """
    의료급여비용 지급통보서 1페이지 파싱

    라인 구조:
    L0:  "의료급여비용 지급통보서"
    L6:  "31590101 (OO치과의원)"
    L17: "진료년월" (라벨)
    L18: "2025.02" (값)
    L26: "2025.03.17" (수납일자)
    L30: "2025.03.20" (지급일자)
    L45: "소득세" (라벨)
    L46: "주민세" (라벨)
    L47: "세액계" (라벨)
    L48: 38,260 (소득세 값)
    L49: 3,820 (주민세 값)
    L50: 42,080 (세액계 값)
    L51: 1,276,900 (기관부담금)
    L73: 1,234,820 (실지급액)
    L114: 20250359 (지급차수 - 심사결정내역)
    L115: 1,293,800 (총진료비)
    L116: 16,900 (본인부담금)
    L117: 1,276,900 (기관부담금)
    """
    record = {
        'institution': '',
        'month': '',
        'total_charge': 0,
        'patient_amount': 0,
        'insurer_amount': 0,
        'payment_amount': 0,
        'income_tax': 0,
        'resident_tax': 0,
        'payment_date': '',
        'claim_type': '의료급여'
    }

    # 기관명: "(기관명)" 패턴
    for line in lines[:10]:
        m = re.search(r'\((.+?)\)', line)
        if m:
            name = m.group(1)
            if any(kw in name for kw in ['의원', '병원', '치과', '한의원', '약국', '클리닉']):
                record['institution'] = name
                break

    # 진료년월: 10~25 라인에서 "YYYY.MM" 패턴
    for line in lines[10:25]:
        stripped = line.strip()
        m = re.match(r'^(20\d{2})[.\-/]?\s*(\d{2})$', stripped)
        if m:
            record['month'] = f"{m.group(1)}-{m.group(2)}"
            break

    # 지급일자: 25~35 라인
    record['payment_date'] = _find_date_in_range(lines, 25, 35)

    # 소득세/주민세/세액계: "소득세" 라벨 찾기
    tax_idx = _find_line_index(lines, '소득세')
    if tax_idx >= 0:
        # 라벨 3줄 (소득세, 주민세, 세액계) 다음에 값 3줄
        # 소득세 값 = tax_idx + 3 (라벨 3개 건너뜀)
        val1 = _get_number_at(lines, tax_idx + 3)  # 소득세
        val2 = _get_number_at(lines, tax_idx + 4)  # 주민세
        val3 = _get_number_at(lines, tax_idx + 5)  # 세액계

        # 검증: 소득세 + 주민세 = 세액계
        if val1 + val2 == val3:
            record['income_tax'] = val1
            record['resident_tax'] = val2
        elif val1 > 0:
            record['income_tax'] = val1
            record['resident_tax'] = val2

        # 기관부담금: 세액계 값 다음 줄 (tax_idx + 6)
        insurer_val = _get_number_at(lines, tax_idx + 6)
        if insurer_val > 10000:
            record['insurer_amount'] = insurer_val

    # 실지급액: footer 제외한 마지막 큰 숫자
    record['payment_amount'] = _find_last_payment(lines[:90])
    if record['payment_amount'] == 0:
        # 대안: 기관부담금 이후 0들 지나서 첫 큰 숫자
        if tax_idx >= 0:
            zero_run = 0
            for i in range(tax_idx + 7, min(tax_idx + 30, len(lines))):
                val = _get_number_at(lines, i)
                if val == 0:
                    zero_run += 1
                elif val > 10000 and zero_run >= 3:
                    record['payment_amount'] = val
                    break

    # 심사결정내역: 지급차수(20XXXXXX) 다음 총진료비/본인부담금
    for i in range(90, min(len(lines), 140)):
        stripped = lines[i].strip()
        if re.match(r'^20\d{5,6}$', stripped):
            # 다음 줄: 총진료비, 본인부담금, 기관부담금
            t = _get_number_at(lines, i + 1)
            p = _get_number_at(lines, i + 2)
            ins = _get_number_at(lines, i + 3)
            if t > 0:
                record['total_charge'] = t
                record['patient_amount'] = p
                if record['insurer_amount'] == 0 and ins > 0:
                    record['insurer_amount'] = ins
            break

    # 보정
    if record['total_charge'] == 0 and record['insurer_amount'] > 0:
        record['total_charge'] = record['insurer_amount'] + record['patient_amount']

    if record['payment_amount'] == 0 and record['insurer_amount'] > 0:
        record['payment_amount'] = (record['insurer_amount']
                                     - record['income_tax']
                                     - record['resident_tax'])

    return record


def parse_nhis_page(lines):
    """
    요양급여비용 지급통보서 1페이지 파싱

    라인 구조:
    L2:  "31590101 (OO치과의원)"
    L4:  지급차수일자
    L12: 진료년월 (000000=가지급, YYYYMM=본지급)
    L45~52: 청구 및 심 데이터 (건수, 총진료비, 본인부담금, 공단부담금, ...)
    L88-90: 소득세/주민세/세액계 라벨
    L91: 지급일자
    L92~102: 원천징수 데이터
    L121 or L125: 실지급액
    """
    record = {
        'institution': '',
        'month': '',
        'total_charge': 0,
        'patient_amount': 0,
        'insurer_amount': 0,
        'payment_amount': 0,
        'income_tax': 0,
        'resident_tax': 0,
        'payment_date': '',
        'claim_type': '요양급여',
        '_is_advance': False,  # 가지급 여부
    }

    # 기관명
    for line in lines[:5]:
        m = re.search(r'\((.+?)\)', line)
        if m:
            name = m.group(1)
            if any(kw in name for kw in ['의원', '병원', '치과', '한의원', '약국', '클리닉']):
                record['institution'] = name
                break

    # 진료년월 (L12)
    if len(lines) > 12:
        stripped = lines[12].strip()
        if re.match(r'^20\d{4}$', stripped) and stripped != '000000':
            record['month'] = f"{stripped[:4]}-{stripped[4:6]}"
        elif stripped == '000000':
            record['_is_advance'] = True
            # 가지급: 지급차수일자(L4)에서 월 추출
            if len(lines) > 4:
                date_m = re.match(r'^(20\d{2})-(\d{2})-\d{2}$', lines[4].strip())
                if date_m:
                    record['month'] = f"{date_m.group(1)}-{date_m.group(2)}"

    # 진료비 데이터: 40~70 라인에서 첫 번째 큰 금액
    first_big_idx = -1
    for i in range(40, min(70, len(lines))):
        val = _get_number_at(lines, i)
        if val >= 100000:
            first_big_idx = i
            break

    if first_big_idx > 0:
        total = _get_number_at(lines, first_big_idx)
        patient = _get_number_at(lines, first_big_idx + 1)
        insurer = _get_number_at(lines, first_big_idx + 2)

        record['total_charge'] = total
        record['patient_amount'] = patient
        record['insurer_amount'] = insurer

        # 본지급 페이지는 심사결정 금액이 있을 수 있음 (더 정확)
        # 심사건수 다음에 심사총진료비가 나옴
        if not record['_is_advance']:
            # first_big_idx+3 이후에 심사 건수(작은 수) → 심사 총진료비
            for j in range(first_big_idx + 3, min(first_big_idx + 12, len(lines))):
                candidate = _get_number_at(lines, j)
                if candidate >= 100000:
                    if candidate == total:
                        # 심사결정 == 청구 → 동일하므로 청구값 유지
                        break
                    # 심사결정 금액이 다르면 심사결정 사용
                    stotal = candidate
                    spatient = _get_number_at(lines, j + 1)
                    sinsurer = _get_number_at(lines, j + 2)
                    if stotal > 0 and spatient >= 0:
                        record['total_charge'] = stotal
                        record['patient_amount'] = spatient
                        record['insurer_amount'] = sinsurer
                    break

    # 지급일자 & 원천징수
    tax_label_idx = _find_line_index(lines, '소득세')
    if tax_label_idx >= 0:
        # 지급일자: 소득세 라벨 이후 첫 날짜
        record['payment_date'] = _find_date_in_range(lines, tax_label_idx, tax_label_idx + 10)

        # 지급일자 라인 찾기
        date_line = -1
        for i in range(tax_label_idx, min(tax_label_idx + 10, len(lines))):
            if re.match(r'^\s*20\d{2}-\d{2}-\d{2}\s*$', lines[i].strip()):
                date_line = i
                break

        if date_line >= 0:
            # 지급일 다음부터 숫자들: 공단부담금정산, ?, 소득세, 주민세, 세액계, ...
            post_nums = []
            for i in range(date_line + 1, min(date_line + 12, len(lines))):
                stripped = lines[i].strip()
                if stripped and re.match(r'^[\d,]+$', stripped.replace(',', '').replace(' ', '')):
                    post_nums.append((i, _get_number_at(lines, i)))
                elif any(kw in stripped for kw in ['환수', '정산', '대불']):
                    break

            # 소득세+주민세=세액계 패턴 찾기
            for j in range(len(post_nums) - 2):
                a = post_nums[j][1]
                b = post_nums[j + 1][1]
                c = post_nums[j + 2][1]
                if a > 0 and b > 0 and a + b == c:
                    record['income_tax'] = a
                    record['resident_tax'] = b
                    break

    # 실지급액: footer 제외한 마지막 큰 숫자
    record['payment_amount'] = _find_last_payment(lines)

    # 지급일자 없으면 상단에서
    if not record['payment_date']:
        record['payment_date'] = _find_date_in_range(lines, 2, 10)

    # 보정
    if record['insurer_amount'] == 0 and record['total_charge'] > 0:
        record['insurer_amount'] = record['total_charge'] - record['patient_amount']

    return record


def parse_pdf_auto(pdf_path, password=None):
    """
    PDF 유형을 자동 감지하여 적절한 파서 호출
    """
    pages, total_pages = extract_all_text(pdf_path, password)

    first_text = pages.get(1, '')
    pdf_type = detect_pdf_type(first_text)

    result = {
        'institution': '',
        'period': '',
        'records': [],
        'summary': {
            'total_charge': 0,
            'insurer_total': 0,
            'patient_total': 0,
            'payment_total': 0
        },
        'total_pages': total_pages,
        'raw_text': "\n".join(pages.values())
    }

    for pg_num, pg_text in pages.items():
        lines = pg_text.strip().split('\n')

        if pdf_type == 'medical_aid':
            record = parse_medical_aid_page(lines)
        else:
            record = parse_nhis_page(lines)

        # _is_advance 플래그 유지 (UI에서 별도 표시용)
        is_advance = record.get('_is_advance', False)
        record.pop('_is_advance', None)
        if is_advance:
            record['is_advance'] = True

        # 유효한 레코드만 추가
        if record['total_charge'] > 0 or record['insurer_amount'] > 0:
            if record['institution'] and not result['institution']:
                result['institution'] = record['institution']

            # 중복 체크
            is_dup = any(
                r['month'] == record['month'] and
                r['total_charge'] == record['total_charge'] and
                r['claim_type'] == record['claim_type']
                for r in result['records']
            )
            if not is_dup:
                result['records'].append(record)

    result['records'].sort(key=lambda r: r.get('month', ''))

    if result['records']:
        months = [r['month'] for r in result['records'] if r['month']]
        if months:
            result['period'] = f"{min(months)} ~ {max(months)}"

    for rec in result['records']:
        result['summary']['total_charge'] += rec.get('total_charge', 0)
        result['summary']['insurer_total'] += rec.get('insurer_amount', 0)
        result['summary']['patient_total'] += rec.get('patient_amount', 0)
        result['summary']['payment_total'] += rec.get('payment_amount', 0)

    return result, pdf_type
