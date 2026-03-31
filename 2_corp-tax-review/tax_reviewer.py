"""
법인세 검토 엔진
파싱된 데이터를 기반으로 세무조정 항목별 검증 수행
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ReviewItem:
    카테고리: str
    항목명: str
    상태: str  # "정상", "이슈", "확인필요", "해당없음"
    신고서금액: Optional[int] = None
    검증금액: Optional[int] = None
    차이금액: Optional[int] = None
    비고: str = ""
    리스크수준: str = "낮음"


def format_won(amount):
    if amount is None:
        return "N/A"
    if amount < 0:
        return f"({abs(amount):,})"
    return f"{amount:,}"


def calculate_corporate_tax(taxable_income):
    """과세표준 구간별 법인세 산출세액 (2025년 귀속)"""
    if taxable_income <= 0:
        return 0
    if taxable_income <= 200_000_000:
        return int(taxable_income * 0.09)
    elif taxable_income <= 20_000_000_000:
        return int(200_000_000 * 0.09 + (taxable_income - 200_000_000) * 0.19)
    elif taxable_income <= 300_000_000_000:
        return int(200_000_000 * 0.09 + 19_800_000_000 * 0.19 + (taxable_income - 20_000_000_000) * 0.21)
    else:
        return int(200_000_000 * 0.09 + 19_800_000_000 * 0.19 + 280_000_000_000 * 0.21 + (taxable_income - 300_000_000_000) * 0.24)


def get_tax_rate(taxable_income):
    """최고 적용 세율 (참고용)"""
    if taxable_income <= 200_000_000:
        return 9
    elif taxable_income <= 20_000_000_000:
        return 19
    elif taxable_income <= 300_000_000_000:
        return 21
    return 24


def get_business_months(data):
    """사업연도 개월수 계산"""
    info = data.get("회사정보", {})
    시작 = info.get("사업연도_시작", "")
    종료 = info.get("사업연도_종료", "")
    if 시작 and 종료:
        try:
            from datetime import datetime
            s = datetime.strptime(시작, "%Y.%m.%d")
            e = datetime.strptime(종료, "%Y.%m.%d")
            months = (e.year - s.year) * 12 + (e.month - s.month) + 1
            if 1 <= months <= 12:
                return months
        except:
            pass
    return 12


def calculate_corporate_tax_annualized(taxable_income, months):
    """사업연도가 12개월 미만일 때 연환산 과세표준으로 산출세액 계산
    연환산 과세표준 = 과세표준 × 12 / 사업월수
    산출세액 = 연환산 산출세액 × 사업월수 / 12
    """
    if months >= 12:
        return calculate_corporate_tax(taxable_income)
    연환산_과세표준 = int(taxable_income * 12 / months)
    연환산_산출세액 = calculate_corporate_tax(연환산_과세표준)
    return int(연환산_산출세액 * months / 12)


# ============================================================
# 1. 이월결손금 확인
# ============================================================
def review_carried_forward_loss(data):
    results = []
    cap = data.get("자본금적립금", {})
    tax = data.get("세액조정", {})

    이월결손금 = cap.get("이월결손금", 0)
    과세표준 = tax.get("과세표준")
    각사업연도소득 = tax.get("각사업연도소득금액")

    if 이월결손금 and 이월결손금 > 0:
        if 과세표준 and 각사업연도소득:
            차감액 = 각사업연도소득 - 과세표준
            if 차감액 >= 이월결손금:
                results.append(ReviewItem(
                    카테고리="이월결손금",
                    항목명="이월결손금 및 결손금 검토",
                    상태="정상",
                    신고서금액=이월결손금,
                    검증금액=차감액,
                    비고=f"[자본금과적립금조정명세서(갑)] 이월결손금 {format_won(이월결손금)} → [법인세과세표준및세액조정계산서] 과세표준에서 차감됨"
                ))
            else:
                results.append(ReviewItem(
                    카테고리="이월결손금",
                    항목명="이월결손금 및 결손금 검토",
                    상태="이슈",
                    신고서금액=이월결손금,
                    검증금액=차감액,
                    비고=f"[자본금과적립금조정명세서(갑)] 이월결손금 {format_won(이월결손금)} 중 {format_won(차감액)}만 차감 - 확인 필요",
                    리스크수준="높음"
                ))
    else:
        results.append(ReviewItem(
            카테고리="이월결손금",
            항목명="이월결손금 및 결손금 검토",
            상태="정상",
            비고="[자본금과적립금조정명세서(갑)] 이월결손금 없음"
        ))

    # 전기/당기 당기순손실 → 자본금과적립금조정명세서(갑) 결손금 반영 확인
    fs = data.get("재무제표", {})
    전기손익 = fs.get("손익계산서_전기", {})
    전기순이익 = 전기손익.get("당기순이익")
    당기순손익 = tax.get("결산서상당기순손익") or fs.get("손익계산서", {}).get("당기순이익")

    if 전기순이익 is not None and 전기순이익 < 0:
        갑_이익잉여금 = cap.get("갑_이익잉여금")
        # 전기 결손이 당기 이익으로 상쇄되어 이익잉여금이 양수가 된 경우도 정상
        결손_반영 = (이월결손금 and 이월결손금 > 0) or (갑_이익잉여금 is not None and 갑_이익잉여금 < 0)
        당기이익_상쇄 = 갑_이익잉여금 is not None and 갑_이익잉여금 > 0  # 당기 흑자전환
        if 결손_반영:
            반영내역 = f"결손금 {format_won(이월결손금)}" if 이월결손금 else f"이익잉여금 {format_won(갑_이익잉여금)}"
            results.append(ReviewItem(
                카테고리="이월결손금",
                항목명="전기 당기순손실 → 결손금 반영",
                상태="정상",
                신고서금액=abs(전기순이익),
                비고=f"[결산서 손익계산서] 전기 당기순손실 {format_won(abs(전기순이익))} → [자본금과적립금조정명세서(갑)] {반영내역} 반영됨"
            ))
        elif 당기이익_상쇄:
            results.append(ReviewItem(
                카테고리="이월결손금",
                항목명="전기 당기순손실 → 결손금 반영",
                상태="정상",
                신고서금액=abs(전기순이익),
                비고=f"[결산서 손익계산서] 전기 당기순손실 {format_won(abs(전기순이익))} → [자본금과적립금조정명세서(갑)] 이익잉여금 {format_won(갑_이익잉여금)} (당기 이익으로 상쇄)"
            ))
        else:
            results.append(ReviewItem(
                카테고리="이월결손금",
                항목명="전기 당기순손실 → 결손금 반영",
                상태="확인필요",
                신고서금액=abs(전기순이익),
                비고=f"[결산서 손익계산서] 전기 당기순손실 {format_won(abs(전기순이익))} → [자본금과적립금조정명세서(갑)] 결손금 미반영 확인 필요",
                리스크수준="보통"
            ))

    if 당기순손익 is not None and 당기순손익 < 0:
        결손금_당기발생 = cap.get("결손금_당기발생", 0)
        if 결손금_당기발생 and 결손금_당기발생 > 0:
            results.append(ReviewItem(
                카테고리="이월결손금",
                항목명="당기 당기순손실 → 결손금 반영",
                상태="정상",
                신고서금액=abs(당기순손익),
                검증금액=결손금_당기발생,
                비고=f"[결산서 손익계산서] 당기순손실 {format_won(abs(당기순손익))} → [자본금과적립금조정명세서(갑)] 결손금 당기발생 {format_won(결손금_당기발생)} 반영됨"
            ))
        else:
            results.append(ReviewItem(
                카테고리="이월결손금",
                항목명="당기 당기순손실 → 결손금 반영",
                상태="확인필요",
                신고서금액=abs(당기순손익),
                비고=f"[결산서 손익계산서] 당기순손실 {format_won(abs(당기순손익))} → [자본금과적립금조정명세서(갑)] 결손금 당기발생 반영 여부 확인 필요",
                리스크수준="보통"
            ))

    return results


# ============================================================
# 2. 서식 간 세액공제/감면 크로스체크
# ============================================================
def review_credit_crosscheck(data):
    results = []
    최저한세 = data.get("최저한세", {})
    공제감면 = data.get("공제감면", {})
    세액공제조정 = data.get("세액공제조정", {})
    tax = data.get("세액조정", {})

    # ── 세액공제: 최저한세 vs 세액공제조정명세서 ──
    최저한세_공제 = 최저한세.get("세액공제")
    조정명세서_공제 = 세액공제조정.get("공제세액_합계")

    if 최저한세_공제 and 조정명세서_공제:
        차이 = abs(최저한세_공제 - 조정명세서_공제)
        status = "정상" if 차이 == 0 else "이슈"
        results.append(ReviewItem(
            카테고리="세액공제",
            항목명="세액공제 서식 간 일치 여부",
            상태=status,
            신고서금액=최저한세_공제,
            검증금액=조정명세서_공제,
            차이금액=차이,
            비고=f"[최저한세조정계산서] [세액공제조정명세서(3)] (124)세액공제 {format_won(최저한세_공제)} vs 공제세액합계 {format_won(조정명세서_공제)}",
            리스크수준="높음" if status == "이슈" else "낮음"
        ))

    return results


# ============================================================
# 3-4. 감가상각 검증 (장부가액 + 감가상각비) — 통합 카드
# ============================================================
def review_asset_vs_depreciation(data):
    results = []
    dep = data.get("감가상각", {})
    std_bs = data.get("표준재무상태표", {})
    std_pl = data.get("표준손익계산서", {})

    비고_lines = []
    전체_일치 = True
    idx = 1

    # 감가상각비: 감가상각비조정명세서합계표 회사손금계상액 vs 표준손익계산서 감가상각비
    회사계상액 = dep.get("회사손금계상액")
    PL유형감가 = std_pl.get("유형자산감가상각비", 0) or 0
    PL무형감가 = std_pl.get("무형자산상각비", 0) or 0
    PL감가합계 = PL유형감가 + PL무형감가

    if 회사계상액 and PL감가합계:
        일치 = 회사계상액 == PL감가합계
        if not 일치:
            전체_일치 = False
        비고_lines.append(f"{idx}. 감가상각비: [감가상각비조정명세서합계표] 회사손금계상액 {format_won(회사계상액)} ↔ [표준손익계산서] 감가상각비 {format_won(PL감가합계)} {'✓' if 일치 else '✗'}")
        idx += 1
    elif 회사계상액:
        전체_일치 = False
        비고_lines.append(f"{idx}. 감가상각비: [감가상각비조정명세서합계표] 회사손금계상액 {format_won(회사계상액)} ↔ [표준손익계산서] 감가상각비 파싱 불가 ✗")
        idx += 1

    if 비고_lines:
        비고_lines.append(f"{'✓ 일치' if 전체_일치 else '✗ 불일치'}")
        results.append(ReviewItem(
            카테고리="감가상각검증",
            항목명="감가상각비 검토",
            상태="정상" if 전체_일치 else "확인필요",
            비고="\n".join(비고_lines)
        ))

    return results


def review_depreciation_vs_income_statement(data):
    return []


# ============================================================
# 5-6. 소득구분계산서 (해당 서식이 있는 경우)
# ============================================================
def review_income_classification(data):
    results = []
    공제감면 = data.get("공제감면", {})
    sme = 공제감면.get("세액감면", {}).get("중소기업특별세액감면", {})
    tax = data.get("세액조정", {})

    소득구분 = data.get("소득구분", {})
    감면대상소득 = 소득구분.get("감면대상소득")
    기타분소득 = 소득구분.get("기타분소득", 0)
    소득구분_과세표준 = 소득구분.get("과세표준")
    소득구분_매출액 = 소득구분.get("매출액")

    if sme and 감면대상소득:
        감면세액 = sme.get("감면세액", 0)
        산출세액 = tax.get("산출세액", 0)
        각사업연도소득 = tax.get("각사업연도소득금액", 0)

        # 1. 감면대상소득 크로스체크: 소득구분계산서 ↔ 공제감면세액계산서(2)
        계산서_감면소득 = sme.get("감면대상소득_계산서")
        계산서_총소득 = sme.get("총소득_계산서")
        총소득_소득구분 = 소득구분.get("각사업연도소득") or 소득구분.get("과세표준") or 각사업연도소득

        # 통합: 분자·분모·공제율 → 공제감면세액 적정성
        회사정보 = data.get("회사정보", {})
        종목 = 회사정보.get("종목", "") or 소득구분.get("감면업종", "")
        수도권 = 회사정보.get("수도권")
        소기업 = 회사정보.get("소기업")
        신고_공제율 = sme.get("공제율")

        if 계산서_감면소득 is not None:
            차이_분자 = abs(감면대상소득 - 계산서_감면소득)
            차이_분모 = abs(소득구분_과세표준 - 계산서_총소득) if 소득구분_과세표준 and 계산서_총소득 else 0

            # 업종 기반 적정 감면율
            적정_감면율 = None
            구분_설명 = ""
            공제율_일치 = True
            if 종목 and 신고_공제율 is not None:
                지식기반_키워드 = ["광고", "엔지니어링", "연구개발", "전기통신", "컴퓨터", "프로그래밍",
                              "시스템", "영상", "오디오", "디자인", "소프트웨어", "방송",
                              "정보서비스", "출판", "창작", "예술"]
                도소매의료_키워드 = ["도매", "소매", "의료"]
                is_도소매의료 = any(k in 종목 for k in 도소매의료_키워드)
                is_지식기반 = any(k in 종목 for k in 지식기반_키워드)

                if is_도소매의료:
                    if 소기업 is True:
                        적정_감면율 = 10; 구분_설명 = "도소매·의료 소기업"
                    elif 소기업 is False and 수도권 is False:
                        적정_감면율 = 5; 구분_설명 = "도소매·의료 중기업 수도권외"
                    elif 소기업 is False and 수도권 is True:
                        적정_감면율 = 0; 구분_설명 = "도소매·의료 중기업 수도권 (감면 없음)"
                elif is_지식기반:
                    if 소기업 is True:
                        적정_감면율 = 20 if 수도권 else 30
                        구분_설명 = f"지식기반 소기업 {'수도권' if 수도권 else '수도권외'}"
                    else:
                        적정_감면율 = 10 if 수도권 else 15
                        구분_설명 = f"지식기반 중기업 {'수도권' if 수도권 else '수도권외'}"
                else:
                    if 소기업 is True:
                        적정_감면율 = 20 if 수도권 else 30
                        구분_설명 = f"기타업종 소기업 {'수도권' if 수도권 else '수도권외'}"
                    elif 소기업 is False and 수도권 is False:
                        적정_감면율 = 15; 구분_설명 = "기타업종 중기업 수도권외"
                    elif 소기업 is False and 수도권 is True:
                        적정_감면율 = 0; 구분_설명 = "기타업종 중기업 수도권 (감면 없음)"
                공제율_일치 = (적정_감면율 is not None and 신고_공제율 == 적정_감면율)

            # 3개 서식 감면세액 일치: 합계표(갑) ↔ 최저한세(123) ↔ 계산서(2)
            최저한세 = data.get("최저한세", {})
            최저한세_감면 = 최저한세.get("감면세액")
            합계표_감면 = data.get("공제감면", {}).get("감면소계_적용대상")

            서식_금액 = {}
            if 합계표_감면: 서식_금액["공제감면세액합계표(갑)"] = 합계표_감면
            if 최저한세_감면: 서식_금액["최저한세조정계산서(123)"] = 최저한세_감면
            서식_금액["공제감면세액계산서(2)"] = 감면세액

            금액들 = list(서식_금액.values())
            서식_일치 = all(v == 금액들[0] for v in 금액들) if 금액들 else True

            전체_정상 = (차이_분자 == 0 and 차이_분모 == 0 and 공제율_일치 and 서식_일치)
            status = "정상" if 전체_정상 else "확인필요"

            비고_parts = []
            # 분자: 소득구분 ↔ 계산서(2)
            비고_parts.append(f"1. 감면대상소득(분자): [소득구분계산서] {format_won(감면대상소득)} ↔ [공제감면세액계산서(2)] {format_won(계산서_감면소득)} {'✓' if 차이_분자 == 0 else '✗'}")
            # 분모: 소득구분 ↔ 계산서(2)
            비고_parts.append(f"2. 과세표준(분모): [소득구분계산서] {format_won(소득구분_과세표준)} ↔ [공제감면세액계산서(2)] {format_won(계산서_총소득)} {'✓' if 차이_분모 == 0 else '✗'}")
            # 공제율
            idx = 3
            if 적정_감면율 is not None:
                비고_parts.append(f"{idx}. 공제율: {구분_설명} → 적정 {적정_감면율}%, 신고 {신고_공제율}% {'✓' if 공제율_일치 else '✗'}")
                idx += 1
            # 감면세액 서식 간 비교
            금액_strs = [f"[{k}] {format_won(v)}" for k, v in 서식_금액.items()]
            if 금액_strs:
                비고_parts.append(f"{idx}. 감면세액: {' ↔ '.join(금액_strs)} {'✓' if 서식_일치 else '✗'}")
            비고_parts.append(f"{'✓ 일치' if 서식_일치 else '✗ 불일치'}")

            results.append(ReviewItem(
                카테고리="세액감면",
                항목명="감면세액 검토",
                상태=status,
                신고서금액=감면세액,
                검증금액=감면세액 if 전체_정상 else None,
                차이금액=0 if 전체_정상 else max(금액들) - min(금액들) if 금액들 else 0,
                비고="\n".join(비고_parts)
            ))


    elif sme:
        감면세액 = sme.get("감면세액", 0)
        산출세액 = tax.get("산출세액", 0)
        if 산출세액 and 감면세액:
            감면비율 = round(감면세액 / 산출세액 * 100, 2)
            results.append(ReviewItem(
                카테고리="소득구분",
                항목명="소득구분계산서 미포함",
                상태="확인필요",
                신고서금액=감면세액,
                비고=f"[소득구분계산서] PDF에 미포함 → 감면대상소득 확인 불가 (감면비율 {감면비율}%)"
            ))

    return results


# ============================================================
# 7. 자본금과적립금조정명세서(갑) 크로스체크
# ============================================================
def review_capital_reserves_crosscheck(data):
    results = []
    cap = data.get("자본금적립금", {})
    fs = data.get("재무제표", {})
    bs = fs.get("재무상태표", {})

    # (갑) 7번 "(을)+(병)계" vs (을) 합계 크로스체크
    갑7 = cap.get("갑7_을병계", 0)
    을합계 = cap.get("유보소득_기말", 0)
    if 갑7 and 을합계:
        차이 = abs(갑7 - 을합계)
        status = "정상" if 차이 == 0 else "이슈"
        results.append(ReviewItem(
            카테고리="갑을검증",
            항목명="자본금과적립금(갑)(을) 유보잔액 검토",
            상태=status,
            신고서금액=갑7,
            검증금액=을합계,
            차이금액=차이,
            비고=f"[자본금과적립금조정명세서(갑)] 7.(을)+(병)계 {format_won(갑7)} vs [자본금과적립금조정명세서(을)] 합계 기말잔액 {format_won(을합계)}",
            리스크수준="높음" if status == "이슈" else "낮음"
        ))
    elif 갑7:
        results.append(ReviewItem(
            카테고리="갑을검증",
            항목명="자본금과적립금(갑)(을) 유보잔액 검토",
            상태="확인필요",
            신고서금액=갑7,
            비고=f"[자본금과적립금조정명세서(갑)] 7.(을)+(병)계 {format_won(갑7)} → [자본금과적립금조정명세서(을)] 합계 파싱 불가"
        ))

    # 갑 서식 기말잔액 vs 재무상태표 (테이블 추출 방식)
    std_bs = data.get("표준재무상태표", {})
    갑_자본금 = cap.get("갑_자본금")
    갑_이익잉여금 = cap.get("갑_이익잉여금")
    갑_자본총계 = cap.get("갑_자본총계")

    # 갑 파싱 성공 시 비교, 실패 시에도 BS 값으로 표시
    갑_파싱됨 = 갑_자본금 is not None or 갑_이익잉여금 is not None or 갑_자본총계 is not None
    비고_lines = []
    전체_일치 = True
    idx = 1

    if 갑_파싱됨:
        for 항목명, 갑_val in [("자본금", 갑_자본금), ("이익잉여금", 갑_이익잉여금), ("자본총계", 갑_자본총계)]:
            bs_val = std_bs.get(항목명) or bs.get(항목명)
            if not bs_val and not 갑_val:
                continue
            # 이익잉여금은 BS에서 음수 파싱이 안 될 수 있으므로 절대값 비교
            일치 = bs_val is not None and 갑_val is not None and (bs_val == 갑_val or abs(bs_val) == abs(갑_val))
            if not 일치:
                전체_일치 = False
            비고_lines.append(f"{idx}. {항목명}: [자본금과적립금조정명세서(갑)] {format_won(갑_val or 0)} ↔ [재무상태표] {format_won(bs_val or 0)} {'✓' if 일치 else '✗'}")
            idx += 1
    else:
        # 갑 파싱 실패 → BS 값만 표시
        전체_일치 = False
        for 항목명 in ["자본금", "이익잉여금", "자본총계"]:
            bs_val = std_bs.get(항목명) or bs.get(항목명)
            if bs_val:
                비고_lines.append(f"{idx}. {항목명}: [자본금과적립금조정명세서(갑)] 파싱 불가 ↔ [재무상태표] {format_won(bs_val)}")
                idx += 1

    if 비고_lines:
        비고_lines.append(f"{'✓ 일치' if 전체_일치 else '✗ 불일치'}" if 갑_파싱됨 else "→ 갑 서식 파싱 불가 → 수동 확인 필요")
        results.append(ReviewItem(
            카테고리="기말잔액검증",
            항목명="자본금과적립금(갑) 기말잔액",
            상태="정상" if 전체_일치 else "확인필요",
            비고="\n".join(비고_lines)
        ))

    return results


# ============================================================
# 8. 세액공제조정명세서 → 농어촌특별세 과세표준
# ============================================================
def review_credit_to_rural_tax(data):
    results = []
    세액공제조정 = data.get("세액공제조정", {})
    농어촌 = data.get("농어촌특별세", {})

    공제세액 = 세액공제조정.get("공제세액_합계")
    농어촌과세표준 = 농어촌.get("과세표준")

    if 공제세액 and 농어촌과세표준:
        차이 = abs(농어촌과세표준 - 공제세액)
        if 차이 == 0:
            status = "정상"
            비고 = f"[세액공제조정명세서(3)] 공제세액합계 {format_won(공제세액)} = [농어촌특별세과세표준및세액신고서] 과세표준 {format_won(농어촌과세표준)}"
        else:
            status = "확인필요"
            비고 = f"[세액공제조정명세서(3)] 공제세액합계 {format_won(공제세액)} ≠ [농어촌특별세과세표준및세액신고서] 과세표준 {format_won(농어촌과세표준)} (차이 {format_won(차이)}원 → 전자신고공제 등 비과세 공제항목 확인)"

        results.append(ReviewItem(
            카테고리="농어촌특별세",
            항목명="세액공제 → 농어촌특별세 과세표준",
            상태=status,
            신고서금액=공제세액,
            검증금액=농어촌과세표준,
            차이금액=차이,
            비고=비고
        ))

    return results


# ============================================================
# 기존 검토 항목들
# ============================================================
def review_income_calculation(data):
    """삭제됨"""
    return []
    # 소득금액조정합계표 익금산입 합계 vs 세무조정계산서 익금산입
    조정 = data.get("소득금액조정", {})
    조정합계_익금 = 조정.get("익금산입_합계", 0)
    조정합계_손금 = 조정.get("손금산입_합계", 0)

    if 조정합계_익금 and 익금산입:
        차이 = abs(익금산입 - 조정합계_익금)
        if 차이 == 0:
            status = "정상"
            비고 = f"[소득금액조정합계표] 익금산입 합계 {format_won(조정합계_익금)} = [법인세과세표준및세액조정계산서] 익금산입 {format_won(익금산입)}"
        else:
            status = "확인필요"
            비고 = f"[소득금액조정합계표] 익금산입 합계 {format_won(조정합계_익금)} vs [법인세과세표준및세액조정계산서] 익금산입 {format_won(익금산입)} → 차이 {format_won(차이)} (법인세비용 손금불산입 등 확인)"
        results.append(ReviewItem(
            카테고리="소득금액",
            항목명="익금산입 합계 일치 여부",
            상태=status,
            신고서금액=조정합계_익금,
            검증금액=익금산입,
            차이금액=차이,
            비고=비고
        ))

    if 조정합계_손금 and 손금산입:
        차이 = abs(손금산입 - 조정합계_손금)
        if 차이 == 0:
            status = "정상"
            비고 = f"[소득금액조정합계표] 손금산입 합계 {format_won(조정합계_손금)} = [법인세과세표준및세액조정계산서] 손금산입 {format_won(손금산입)}"
        else:
            status = "확인필요"
            비고 = f"[소득금액조정합계표] 손금산입 합계 {format_won(조정합계_손금)} vs [법인세과세표준및세액조정계산서] 손금산입 {format_won(손금산입)} → 차이 {format_won(차이)} 확인"
        results.append(ReviewItem(
            카테고리="소득금액",
            항목명="손금산입 합계 일치 여부",
            상태=status,
            신고서금액=조정합계_손금,
            검증금액=손금산입,
            차이금액=차이,
            비고=비고
        ))

    return results


def review_tax_calculation(data):
    """세액 계산 검증 - 삭제된 항목: 세율/산출세액/차감세액/분납"""
    return []


def review_vehicle_expenses(data):
    results = []
    veh = data.get("업무용승용차", {})
    if not veh:
        return results

    취득가액 = veh.get("취득가액")
    차량번호 = veh.get("차량번호", "")
    차종 = veh.get("차종", "")
    한도초과 = veh.get("한도초과금액") or veh.get("손금불산입")

    if 한도초과 and 한도초과 > 0:
        results.append(ReviewItem(
            카테고리="세무조정",
            항목명=f"업무용승용차 한도초과 ({차량번호} {차종})",
            상태="정상",
            신고서금액=한도초과,
            비고=f"[업무용승용차관련비용명세서] 취득가액 {format_won(취득가액)}, 감가상각비 연 800만원 한도"
        ))

    return results


def review_financial_consistency(data):
    """재무상태표 등식 - 사용하지 않음"""
    return []


def _get_sme_reduction_rate_range(종목):
    """업종별 가능한 감면율 범위 반환 (조특법 §7)
    도소매업/의료업: 수도권중기업 0%, 비수도권중기업 5%, 소기업 10%
    기타업종: 수도권중기업 10%, 비수도권중기업 15%, 수도권소기업 20%, 비수도권소기업 30%
    """
    도소매의료 = ["도매", "소매", "도소매", "의료", "병원", "의원", "치과", "한의원",
                "약국", "안과", "이비인후과", "피부과", "정형외과", "산부인과"]
    for kw in 도소매의료:
        if kw in 종목:
            return [5, 10], "도소매/의료업"
    return [10, 15, 20, 30], "기타업종"


def review_sme_special_reduction(data):
    """삭제됨 - 감면세액 검토 카드에서 통합 처리"""
    return []


# ============================================================
# 신설법인 자본금 계상 여부
# ============================================================
def review_new_corp_capital(data):
    """신설법인(1기)인 경우 자본금 계상 확인"""
    results = []
    info = data.get("회사정보", {})
    std_bs = data.get("표준재무상태표", {})
    fs = data.get("재무제표", {})
    bs = fs.get("재무상태표", {})

    자본금 = std_bs.get("자본금") or bs.get("자본금")

    # 사업연도 시작이 법인 설립일과 유사하면 1기로 추정
    시작 = info.get("사업연도_시작", "")
    if 시작:
        # 제N기 정보가 있으면 확인 (1기이면 신설법인)
        # 자본금이 있으면 정상 확인
        if 자본금 and 자본금 > 0:
            results.append(ReviewItem(
                카테고리="자본금",
                항목명="자본금 계상 확인",
                상태="정상",
                신고서금액=자본금,
                비고=f"[표준재무상태표] 자본금 {format_won(자본금)} 계상됨"
            ))
        elif 자본금 == 0 or 자본금 is None:
            results.append(ReviewItem(
                카테고리="자본금",
                항목명="자본금 계상 확인",
                상태="이슈",
                비고="[표준재무상태표] 자본금 0원 또는 미계상 → 신설법인 자본금 계상 여부 확인",
                리스크수준="높음"
            ))

    return results


# ============================================================
# 소득금액조정 처분 내용 공란 확인
# ============================================================
def review_adjustment_disposition(data):
    """삭제됨"""
    return []


# ============================================================
# 인정이자 미수수익 결산서 반영여부
# ============================================================
def review_deemed_interest(data):
    """특수관계자 인정이자가 있으면 미수수익이 결산서에 반영됐는지 확인"""
    results = []
    adj = data.get("소득금액조정", {})
    std_bs = data.get("표준재무상태표", {})
    std_pl = data.get("표준손익계산서", {})

    익금항목 = adj.get("익금산입_항목", [])

    # 인정이자 관련 항목 찾기
    인정이자_항목 = [item for item in 익금항목 if "인정이자" in item.get("과목", "") or "미수수익" in item.get("과목", "")]

    if 인정이자_항목:
        이자수익 = std_pl.get("이자수익", 0)
        if 이자수익 and 이자수익 > 0:
            results.append(ReviewItem(
                카테고리="인정이자",
                항목명="인정이자 미수수익 결산서 반영 확인",
                상태="확인필요",
                신고서금액=sum(i.get("금액", 0) for i in 인정이자_항목),
                검증금액=이자수익,
                비고=f"[소득금액조정합계표] 인정이자 세무조정 있음 → [표준손익계산서] 이자수익 {format_won(이자수익)} 반영 여부 확인"
            ))
        else:
            results.append(ReviewItem(
                카테고리="인정이자",
                항목명="인정이자 미수수익 결산서 반영 확인",
                상태="확인필요",
                신고서금액=sum(i.get("금액", 0) for i in 인정이자_항목),
                비고=f"[소득금액조정합계표] 인정이자 세무조정 있음 → [표준손익계산서] 이자수익 미계상, 결산서 반영 여부 확인",
                리스크수준="높음"
            ))

    return results


# ============================================================
# 합계표 ↔ 세액조정계산서 크로스체크 (기타)
# ============================================================
def review_summary_vs_tax_adjustment(data):
    """공제감면세액합계표(갑/을) ↔ 법인세과세표준및세액조정계산서 크로스체크"""
    results = []
    공제감면 = data.get("공제감면", {})
    tax = data.get("세액조정", {})

    # 1) 최저한세 적용대상: 합계표(150) vs 세액조정(17)
    합계표_적용대상 = 공제감면.get("공제감면합계_적용대상")
    세액조정_17 = tax.get("최저한세적용대상_공제감면세액")

    if 합계표_적용대상 is not None and 세액조정_17 is not None:
        차이 = abs(합계표_적용대상 - 세액조정_17)
        status = "정상" if 차이 == 0 else "이슈"
        results.append(ReviewItem(
            카테고리="공제감면",
            항목명="공제감면세액 합계 검토",
            상태=status,
            신고서금액=합계표_적용대상,
            검증금액=세액조정_17,
            차이금액=차이,
            비고=f"[공제감면세액합계표(갑/을)] (150) {format_won(합계표_적용대상)} vs [법인세과세표준및세액조정계산서] (17) {format_won(세액조정_17)}",
            리스크수준="높음" if status == "이슈" else "낮음"
        ))

    # 2) 최저한세 적용제외: 합계표(151) vs 세액조정(19)
    합계표_적용제외 = 공제감면.get("공제감면합계_적용제외")
    세액조정_19 = tax.get("최저한세적용제외_공제감면세액")

    if 합계표_적용제외 is not None and 세액조정_19 is not None:
        차이 = abs(합계표_적용제외 - 세액조정_19)
        status = "정상" if 차이 == 0 else "이슈"
        results.append(ReviewItem(
            카테고리="합계표검증",
            항목명="최저한세 적용제외 공제감면세액",
            상태=status,
            신고서금액=합계표_적용제외,
            검증금액=세액조정_19,
            차이금액=차이,
            비고=f"[공제감면세액합계표(갑/을)] (151) {format_won(합계표_적용제외)} vs [법인세과세표준및세액조정계산서] (19) {format_won(세액조정_19)}",
            리스크수준="높음" if status == "이슈" else "낮음"
        ))

    return results


# ============================================================
# 공제세액 크로스체크: 세액공제조정명세서(3) ↔ 세액공제신청서 ↔ 최저한세
# ============================================================
def review_tax_credit_crosscheck(data):
    """공제세액 3개 서식 크로스체크"""
    results = []
    조정 = data.get("세액공제조정", {})
    신청 = data.get("세액공제신청서", {})
    최저한세 = data.get("최저한세", {})

    비고_lines = []
    전체_일치 = True
    idx = 1

    # 1) 세액공제조정명세서(3) 당기분(107) = 세액공제신청서 대상세액 합계
    조정_당기분 = 조정.get("당기분_당기분") or 조정.get("소계_당기분") or 조정.get("합계_당기분")
    신청_대상세액 = 신청.get("대상세액_합계")

    if 조정_당기분 and 신청_대상세액:
        일치 = 조정_당기분 == 신청_대상세액
        if not 일치:
            전체_일치 = False
        비고_lines.append(f"{idx}. [세액공제조정명세서(3)] 당기분(107): {format_won(조정_당기분)} ↔ [세액공제신청서] 대상세액: {format_won(신청_대상세액)} {'✓' if 일치 else '✗'}")
        idx += 1

    # 2) 세액공제신청서 공제세액 = 세액공제조정명세서(3) 당기분 행의 (123)공제세액
    신청_공제세액 = 신청.get("공제세액_합계")
    조정_당기_공제세액 = 조정.get("당기분_공제세액")

    if 신청_공제세액 and 조정_당기_공제세액:
        일치 = 신청_공제세액 == 조정_당기_공제세액
        if not 일치:
            전체_일치 = False
        비고_lines.append(f"{idx}. [세액공제신청서] 공제세액: {format_won(신청_공제세액)} ↔ [세액공제조정명세서(3)] 당기분 (123)공제세액: {format_won(조정_당기_공제세액)} {'✓' if 일치 else '✗'}")
        idx += 1

    # 3) 세액공제조정명세서(3) (123)합계 = 최저한세조정계산서 세액공제합계
    조정_공제세액_합계 = 조정.get("소계_공제세액") or 조정.get("합계_공제세액")
    최저한세_세액공제 = 최저한세.get("세액공제합계")

    if 조정_공제세액_합계 and 최저한세_세액공제:
        일치 = 조정_공제세액_합계 == 최저한세_세액공제
        if not 일치:
            전체_일치 = False
        비고_lines.append(f"{idx}. [세액공제조정명세서(3)] (123)합계: {format_won(조정_공제세액_합계)} ↔ [최저한세조정계산서] 세액공제: {format_won(최저한세_세액공제)} {'✓' if 일치 else '✗'}")
        idx += 1

    if 비고_lines:
        비고_lines.append(f"{'✓ 일치' if 전체_일치 else '✗ 불일치'}")
        results.append(ReviewItem(
            카테고리="세액공제",
            항목명="공제세액 검토",
            상태="정상" if 전체_일치 else "확인필요",
            신고서금액=조정_당기분,
            검증금액=신청_대상세액,
            비고="\n".join(비고_lines)
        ))

    return results


# ============================================================
# 표준손익계산서 이자수익 → 소득구분계산서 기타분 반영 여부
# ============================================================
def review_interest_income_classification(data):
    """표준손익계산서에 이자수익이 있으면 소득구분계산서 기타분소득에 반영됐는지 확인"""
    results = []
    std_pl = data.get("표준손익계산서", {})
    소득구분 = data.get("소득구분", {})

    # 소득구분계산서가 없거나 감면대상소득이 없으면 검토 불필요 (감면 없는 법인)
    if not 소득구분:
        return results
    감면대상소득 = 소득구분.get("감면대상소득")
    과세표준_소득구분 = 소득구분.get("과세표준")
    # 감면대상소득과 과세표준 모두 없으면 소득구분계산서가 실제로 없는 것
    if not 감면대상소득 and not 과세표준_소득구분:
        return results

    이자수익 = std_pl.get("이자수익", 0) or 0
    기타분소득 = 소득구분.get("기타분소득", 0) or 0

    if 이자수익 and 이자수익 > 0:
        if 기타분소득 and 기타분소득 > 0:
            status = "정상"
            비고 = (f"1. [표준손익계산서] 이자수익: {format_won(이자수익)}원\n"
                   f"2. [소득구분계산서] 기타분소득: {format_won(기타분소득)}원\n"
                   f"→ 이자수익이 기타분으로 구분됨")
        else:
            status = "확인필요"
            비고 = (f"1. [표준손익계산서] 이자수익: {format_won(이자수익)}원\n"
                   f"2. [소득구분계산서] 기타분소득: 0원\n"
                   f"→ 이자수익이 기타분에 미반영 → 감면대상소득 과대계상 가능성")
        results.append(ReviewItem(
            카테고리="소득구분검증",
            항목명="이자수익 기타분 반영 여부",
            상태=status,
            신고서금액=이자수익,
            검증금액=기타분소득,
            비고=비고,
            리스크수준="보통" if status == "확인필요" else "낮음"
        ))

    return results


# ============================================================
# 전체 검토 실행
# ============================================================
def run_all_reviews(data):
    all_results = []

    # 1. 이월결손금
    all_results.extend(review_carried_forward_loss(data))
    # 2. 서식 간 세액공제/감면 크로스체크
    all_results.extend(review_credit_crosscheck(data))
    # 3. 유형자산 vs 감가상각
    all_results.extend(review_asset_vs_depreciation(data))
    # 4. 감가상각비 비교
    all_results.extend(review_depreciation_vs_income_statement(data))
    # 5-6. 소득구분계산서
    all_results.extend(review_income_classification(data))
    # 7. 자본금과적립금 크로스체크
    all_results.extend(review_capital_reserves_crosscheck(data))
    # 8. 세액공제 → 농어촌특별세
    all_results.extend(review_credit_to_rural_tax(data))

    # 소득금액/세액 계산
    all_results.extend(review_income_calculation(data))
    all_results.extend(review_tax_calculation(data))
    all_results.extend(review_vehicle_expenses(data))
    all_results.extend(review_financial_consistency(data))
    all_results.extend(review_sme_special_reduction(data))

    # 합계표 ↔ 세액조정 크로스체크 (기타)
    all_results.extend(review_summary_vs_tax_adjustment(data))
    # 공제세액 크로스체크: 조정명세서(3) ↔ 신청서 ↔ 최저한세
    all_results.extend(review_tax_credit_crosscheck(data))
    # 이자수익 → 소득구분 기타분 (기타)
    all_results.extend(review_interest_income_classification(data))

    # 추가 검증항목
    all_results.extend(review_new_corp_capital(data))
    all_results.extend(review_adjustment_disposition(data))
    all_results.extend(review_deemed_interest(data))

    return all_results
