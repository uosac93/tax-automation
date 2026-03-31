"""
법인세 신고 검토 보고서 PDF 생성
"""
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable, KeepTogether,
                                 PageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# ── 폰트 등록 (CID 내장 폰트 - Windows 폰트 파일 접근 안 함) ──
try:
    pdfmetrics.registerFont(UnicodeCIDFont('HYSMyeongJo-Medium'))  # 명조체 계열
    FONT = 'HYSMyeongJo-Medium'
    FONT_B = 'HYSMyeongJo-Medium'  # CID는 볼드 별도 없음
except Exception:
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('HYGothic-Medium'))
        FONT = 'HYGothic-Medium'
        FONT_B = 'HYGothic-Medium'
    except Exception:
        FONT = 'Helvetica'
        FONT_B = 'Helvetica-Bold'

# ── 색상 ──
NAVY = colors.Color(0.10, 0.16, 0.28)       # 진한 남색 (헤더)
NAVY_LIGHT = colors.Color(0.15, 0.22, 0.36)  # 중간 남색
BLUE = colors.Color(0.22, 0.40, 0.65)        # 파란 강조
BLUE_LIGHT = colors.Color(0.90, 0.93, 0.97)  # 연한 파란 배경
GRAY_BG = colors.Color(0.96, 0.96, 0.97)     # 회색 배경
GRAY_LINE = colors.Color(0.78, 0.78, 0.80)   # 테두리선
GRAY_TEXT = colors.Color(0.45, 0.45, 0.45)    # 보조 텍스트
BLACK = colors.Color(0.10, 0.10, 0.10)        # 본문
WHITE = colors.white
RED = colors.Color(0.85, 0.20, 0.20)
GREEN = colors.Color(0.18, 0.58, 0.30)
ORANGE = colors.Color(0.85, 0.55, 0.15)

PAGE_W = A4[0]
CONTENT_W = PAGE_W - 36*mm  # 좌우 마진 18mm


def _fmt(val):
    if val is None:
        return "-"
    if isinstance(val, (int, float)):
        v = int(val)
        if v < 0:
            return f"△{abs(v):,}"
        return f"{v:,}"
    return str(val)


def _pct(val, total):
    if not total or not val:
        return "-"
    return f"{val / total * 100:.1f}%"


def _delta(cur, prev):
    if cur is None or prev is None:
        return "-"
    d = cur - prev
    if d > 0:
        return f"+{d:,}"
    elif d < 0:
        return f"-{abs(d):,}"
    return "0"


def _delta_pct(cur, prev):
    if cur is None or prev is None or prev == 0:
        return "-"
    d = (cur - prev) / abs(prev) * 100
    return f"{d:+.1f}%"


# ── 스타일 헬퍼 ──
def _p(text, font=FONT, size=8, color=BLACK, align=TA_LEFT, bold=False):
    """간단한 Paragraph 생성"""
    fn = FONT_B if bold else font
    style = ParagraphStyle('_tmp', fontName=fn, fontSize=size,
                           textColor=color, alignment=align,
                           leading=size * 1.4)
    return Paragraph(str(text), style)


def _header_table(corp_name, period_start, period_end, biz_no):
    """보고서 상단 헤더"""
    # 타이틀
    title = _p(f"{corp_name} 법인세 신고 사항", FONT_B, 18, NAVY, TA_CENTER, True)
    # 회사 정보
    info_line = f"{corp_name}"
    if biz_no:
        info_line += f"  |  {biz_no}"
    info_line += f"  |  {period_start} ~ {period_end}"
    info = _p(info_line, FONT, 9, GRAY_TEXT, TA_CENTER)
    date_line = _p(f"작성일: {datetime.now().strftime('%Y년 %m월 %d일')}", FONT, 8, GRAY_TEXT, TA_CENTER)

    data = [[title], [info], [date_line]]
    t = Table(data, colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ('TOPPADDING', (0, 0), (0, 0), 16),
        ('BOTTOMPADDING', (0, 0), (0, 0), 8),
        ('TOPPADDING', (0, 1), (0, 1), 4),
        ('BOTTOMPADDING', (0, 1), (0, 1), 4),
        ('TOPPADDING', (0, 2), (0, 2), 2),
        ('BOTTOMPADDING', (0, 2), (0, 2), 12),
        ('LINEBELOW', (0, -1), (-1, -1), 2, NAVY),
    ]))
    return t


def _section(num, title):
    """섹션 번호 + 제목"""
    data = [[_p(f"  {num}", FONT_B, 9, WHITE, TA_CENTER, True),
             _p(f"  {title}", FONT_B, 11, NAVY, TA_LEFT, True)]]
    t = Table(data, colWidths=[12*mm, CONTENT_W - 12*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), NAVY),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEBELOW', (0, 0), (-1, 0), 1, NAVY),
    ]))
    return t


def generate_report_pdf(data, results, output_path):
    """법인세 검토 보고서 PDF 생성"""
    info = data.get("회사정보", {})
    tax = data.get("세액조정", {})
    fs = data.get("재무제표", {})
    is_data = fs.get("손익계산서", {})
    is_prev = fs.get("손익계산서_전기", {})
    if not is_prev:
        is_prev = data.get("손익계산서_전기", {})

    raw_name = info.get("법인명", "")
    import re as _re
    _m = _re.match(r'^주식회사\s*(.+)$', raw_name)
    if _m:
        corp_name = f"(주){_m.group(1)}"
    else:
        _m = _re.match(r'^(.+?)\s*주식회사$', raw_name)
        corp_name = f"{_m.group(1)}(주)" if _m else raw_name
    period_start = info.get("사업연도_시작", "")
    period_end = info.get("사업연도_종료", "")
    biz_no = info.get("사업자등록번호", "")
    cur_year = period_end[:4] if period_end else "당기"
    prev_year = str(int(cur_year) - 1) if cur_year.isdigit() else "전기"

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=18*mm, bottomMargin=15*mm,
        leftMargin=18*mm, rightMargin=18*mm
    )

    story = []

    # ═══════════════════════════════════════
    #  헤더
    # ═══════════════════════════════════════
    story.append(_header_table(corp_name, period_start, period_end, biz_no))
    story.append(Spacer(1, 8*mm))

    # ═══════════════════════════════════════
    #  01. 손익계산서 주요 항목 비교
    # ═══════════════════════════════════════
    story.append(_section("01", "손익계산서 주요 항목 비교"))
    story.append(Spacer(1, 3*mm))

    sales_cur = is_data.get("매출액")
    sales_prev = is_prev.get("매출액")

    is_items = [
        ("매출액", True), ("매출원가", False), ("매출총이익", True),
        ("판관비", False), ("영업이익", True),
        ("영업외수익", False), ("영업외비용", False),
        ("법인세차감전이익", True), ("법인세등", False), ("당기순이익", True),
    ]

    # 헤더 행
    hdr = [
        _p("계정과목", FONT_B, 8, WHITE, TA_CENTER, True),
        _p(f"{cur_year}년(당기)", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("비율", FONT_B, 8, WHITE, TA_CENTER, True),
        _p(f"{prev_year}년(전기)", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("비율", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("증감액", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("증감률", FONT_B, 8, WHITE, TA_CENTER, True),
    ]
    tbl_rows = [hdr]

    for key, is_bold in is_items:
        cur = is_data.get(key)
        prev = is_prev.get(key)
        fn = FONT_B if is_bold else FONT
        name_color = NAVY if is_bold else BLACK
        delta_val = _delta(cur, prev)
        delta_color = BLACK
        if delta_val.startswith("+"):
            delta_color = GREEN
        elif delta_val.startswith("-"):
            delta_color = RED

        row = [
            _p(key, fn, 8, name_color, TA_LEFT, is_bold),
            _p(_fmt(cur), fn, 8, BLACK, TA_RIGHT, is_bold),
            _p(_pct(cur, sales_cur), FONT, 7.5, GRAY_TEXT, TA_RIGHT),
            _p(_fmt(prev), fn, 8, BLACK, TA_RIGHT, is_bold),
            _p(_pct(prev, sales_prev), FONT, 7.5, GRAY_TEXT, TA_RIGHT),
            _p(delta_val, fn, 8, delta_color, TA_RIGHT, is_bold),
            _p(_delta_pct(cur, prev), FONT, 7.5, delta_color, TA_RIGHT),
        ]
        tbl_rows.append(row)

    # 전체 너비 = CONTENT_W (174mm)에 맞춤
    col_w = [36*mm, 28*mm, 16*mm, 28*mm, 16*mm, 28*mm, 22*mm]
    tbl = Table(tbl_rows, colWidths=col_w, repeatRows=1)
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, GRAY_LINE),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]
    # 주요 항목 행 배경색
    for i, (_, is_bold) in enumerate(is_items):
        row_idx = i + 1
        if is_bold:
            style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), BLUE_LIGHT))
        elif row_idx % 2 == 0:
            style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), GRAY_BG))
    tbl.setStyle(TableStyle(style))
    story.append(tbl)
    story.append(Spacer(1, 8*mm))

    # ═══════════════════════════════════════
    #  02. 주요 세무조정 내역
    # ═══════════════════════════════════════
    story.append(_section("02", "주요 세무조정 내역"))
    story.append(Spacer(1, 3*mm))

    adj = data.get("소득금액조정", {})
    adj_ik = adj.get("익금산입_항목", [])
    adj_sk = adj.get("손금산입_항목", [])

    adj_hdr = [
        _p("구분", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("조정항목", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("금액", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("소득처분", FONT_B, 8, WHITE, TA_CENTER, True),
    ]
    adj_rows = [adj_hdr]

    for item in adj_ik:
        adj_rows.append([
            _p("익금산입", FONT, 8, BLUE, TA_CENTER),
            _p(item.get("항목명", "") or item.get("과목", ""), FONT, 8, BLACK, TA_LEFT),
            _p(_fmt(item.get("금액")), FONT, 8, BLACK, TA_RIGHT),
            _p(item.get("처분", ""), FONT, 8, GRAY_TEXT, TA_CENTER),
        ])
    for item in adj_sk:
        adj_rows.append([
            _p("손금산입", FONT, 8, BLUE, TA_CENTER),
            _p(item.get("항목명", "") or item.get("과목", ""), FONT, 8, BLACK, TA_LEFT),
            _p(_fmt(item.get("금액")), FONT, 8, BLACK, TA_RIGHT),
            _p(item.get("처분", ""), FONT, 8, GRAY_TEXT, TA_CENTER),
        ])

    # 합계
    ik_total = adj.get("익금산입_합계")
    sk_total = adj.get("손금산입_합계")
    if ik_total:
        adj_rows.append([
            _p("", FONT_B, 8, BLACK, TA_CENTER),
            _p("익금산입 합계", FONT_B, 8, NAVY, TA_LEFT, True),
            _p(_fmt(ik_total), FONT_B, 8, NAVY, TA_RIGHT, True),
            _p("", FONT, 8, BLACK, TA_CENTER),
        ])
    if sk_total:
        adj_rows.append([
            _p("", FONT_B, 8, BLACK, TA_CENTER),
            _p("손금산입 합계", FONT_B, 8, NAVY, TA_LEFT, True),
            _p(_fmt(sk_total), FONT_B, 8, NAVY, TA_RIGHT, True),
            _p("", FONT, 8, BLACK, TA_CENTER),
        ])

    if len(adj_rows) > 1:
        col_w2 = [30*mm, 64*mm, 40*mm, 40*mm]  # = 174mm
        tbl2 = Table(adj_rows, colWidths=col_w2, repeatRows=1)
        style2 = [
            ('BACKGROUND', (0, 0), (-1, 0), NAVY),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, GRAY_LINE),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, GRAY_BG]),
        ]
        # 합계행 강조
        total_start = len(adj_rows) - (1 if ik_total else 0) - (1 if sk_total else 0)
        for r in range(total_start, len(adj_rows)):
            style2.append(('BACKGROUND', (0, r), (-1, r), BLUE_LIGHT))
            style2.append(('LINEABOVE', (0, r), (-1, r), 1, NAVY))
        tbl2.setStyle(TableStyle(style2))
        story.append(tbl2)
    else:
        story.append(_p("세무조정 내역 없음", FONT, 9, GRAY_TEXT))
    story.append(Spacer(1, 8*mm))

    # ═══════════════════════════════════════
    #  03. 법인세 산출 내역
    # ═══════════════════════════════════════
    story.append(_section("03", "법인세 산출 내역"))
    story.append(Spacer(1, 3*mm))

    납부 = tax.get("차감납부할세액")
    지방세 = int(납부 * 0.1) if 납부 else None
    합계납부 = (납부 + 지방세) if (납부 and 지방세) else None

    tax_items = [
        ("결산서상 당기순손익", tax.get("결산서상당기순손익"), False, False),
        ("(+) 익금산입", tax.get("익금산입"), False, False),
        ("(-) 손금산입", tax.get("손금산입"), False, False),
        ("각사업연도 소득금액", tax.get("각사업연도소득금액"), True, True),
        ("(-) 이월결손금 공제", tax.get("이월결손금공제"), False, False),
        ("(-) 비과세소득", tax.get("비과세소득"), False, False),
        ("(-) 소득공제", tax.get("소득공제"), False, False),
        ("과세표준", tax.get("과세표준"), True, True),
        ("산출세액", tax.get("산출세액"), False, False),
        ("(-) 공제감면세액", tax.get("최저한세적용대상_공제감면세액"), False, False),
        ("(-) 기납부세액", tax.get("기납부세액"), False, False),
        ("법인세 차감납부세액", 납부, True, True),
        ("법인지방소득세 (10%)", 지방세, False, False),
        ("합계 납부세액", 합계납부, True, True),
    ]

    tax_hdr = [
        _p("구분", FONT_B, 8, WHITE, TA_CENTER, True),
        _p("금액 (원)", FONT_B, 8, WHITE, TA_CENTER, True),
    ]
    tax_rows = [tax_hdr]
    for label, val, is_bold, is_highlight in tax_items:
        if val is None and not is_bold:
            continue
        tax_rows.append([
            _p(label, FONT_B if is_bold else FONT, 8.5 if is_bold else 8,
               NAVY if is_bold else BLACK, TA_LEFT, is_bold),
            _p(_fmt(val), FONT_B if is_bold else FONT, 8.5 if is_bold else 8,
               NAVY if is_bold else BLACK, TA_RIGHT, is_bold),
        ])

    col_w3 = [104*mm, 70*mm]  # = 174mm
    tbl3 = Table(tax_rows, colWidths=col_w3, repeatRows=1)
    style3 = [
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, GRAY_LINE),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]
    row_i = 1
    for label, val, is_bold, is_highlight in tax_items:
        if val is None and not is_bold:
            continue
        if is_highlight:
            style3.append(('BACKGROUND', (0, row_i), (-1, row_i), BLUE_LIGHT))
            style3.append(('LINEABOVE', (0, row_i), (-1, row_i), 1, NAVY))
        elif row_i % 2 == 0:
            style3.append(('BACKGROUND', (0, row_i), (-1, row_i), GRAY_BG))
        row_i += 1
    tbl3.setStyle(TableStyle(style3))
    story.append(tbl3)
    story.append(Spacer(1, 8*mm))

    # ═══════════════════════════════════════
    #  푸터
    # ═══════════════════════════════════════
    story.append(Spacer(1, 12*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=GRAY_LINE))
    story.append(Spacer(1, 3*mm))
    story.append(_p(f"Corp Tax_AI  |  {datetime.now().strftime('%Y.%m.%d %H:%M')}  |  박양훈 세무사",
                    FONT, 7.5, GRAY_TEXT, TA_CENTER))

    doc.build(story)
    return output_path
