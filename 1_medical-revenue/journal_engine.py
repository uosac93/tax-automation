"""
병의원 매출 분개 생성 엔진

회계처리 방식:
[매출인식 - 진료월 말일자]
  요양급여(건강보험):
    CR 요양급여수입(413): 총진료비 (심사결정)
    DR 미수금(120): 공단부담금
    DR 일반진료수입(411): -본인부담금 (차감)

  의료급여(의료보호):
    CR 의료급여수입(414): 총진료비
    DR 미수금(120): 기관부담금
    DR 일반진료수입(411): -본인부담금 (차감)

[입금 - 지급일자]
  DR 보통예금(103): 실수령액
  DR 선납세금(136): 소득세
  DR 선납주민세(127): 주민세
  CR 미수금(120): 미수금 회수
"""
from dataclasses import dataclass
import calendar


# ============================================================
# 기본 계정과목 설정
# ============================================================
DEFAULT_ACCOUNTS = {
    # 자산 (차변)
    'ar': {'code': '120', 'name': '미수금'},
    'cash': {'code': '101', 'name': '현금'},
    'bank': {'code': '103', 'name': '보통예금'},
    'prepaid_tax': {'code': '136', 'name': '선납세금'},
    'prepaid_resident': {'code': '127', 'name': '선납주민세'},

    # 수익 (대변)
    'rev_general': {'code': '411', 'name': '일반진료수입'},
    'rev_insurance': {'code': '413', 'name': '요양급여수입'},
    'rev_medical_aid': {'code': '414', 'name': '의료급여수입'},
}


@dataclass
class JournalEntry:
    """분개 항목"""
    date: str
    description: str
    debit_account: str
    debit_name: str
    debit_amount: int
    credit_account: str
    credit_name: str
    credit_amount: int
    revenue_type: str = ''
    month: str = ''
    partner_code: str = ''
    partner_name: str = ''


class JournalGenerator:
    def __init__(self, accounts=None):
        self.accounts = accounts or DEFAULT_ACCOUNTS

    def get_account(self, key):
        acct = self.accounts.get(key, {})
        return acct.get('code', ''), acct.get('name', '')

    def generate_from_records(self, records, payment_method='card'):
        entries = []

        for rec in records:
            month = rec.get('month', '')
            claim_type = rec.get('claim_type', '요양급여')
            entry_date = self._month_to_last_date(month)

            total_charge = rec.get('total_charge', 0)
            patient_amount = rec.get('patient_amount', 0)
            insurer_amount = rec.get('insurer_amount', 0)
            payment_amount = rec.get('payment_amount', 0)
            income_tax = rec.get('income_tax', 0)
            resident_tax = rec.get('resident_tax', 0)
            payment_date = rec.get('payment_date', '')

            partner_code = '00357'
            partner_name = '건강보험'

            # 보험 유형별 수익 계정
            if claim_type == '의료급여':
                rev_key = 'rev_medical_aid'
                type_label = '의료급여'
            else:
                rev_key = 'rev_insurance'
                type_label = '요양급여'

            rev_code, rev_name = self.get_account(rev_key)
            gen_code, gen_name = self.get_account('rev_general')
            ar_code, ar_name = self.get_account('ar')

            # ── 매출인식 (진료월 말일) ──
            if total_charge > 0:
                # 1) CR 보험수입: 총진료비
                entries.append(JournalEntry(
                    date=entry_date,
                    description=f"{month} {type_label}수입",
                    debit_account='',
                    debit_name='',
                    debit_amount=0,
                    credit_account=rev_code,
                    credit_name=rev_name,
                    credit_amount=total_charge,
                    revenue_type=f'{type_label}-매출',
                    month=month,
                    partner_code=partner_code,
                    partner_name=partner_name
                ))

                # 2) DR 미수금: 공단부담금(기관부담금)
                ar_amount = insurer_amount if insurer_amount > 0 else (total_charge - patient_amount)
                if ar_amount > 0:
                    entries.append(JournalEntry(
                        date=entry_date,
                        description=f"{month} {type_label} 미수금",
                        debit_account=ar_code,
                        debit_name=ar_name,
                        debit_amount=ar_amount,
                        credit_account='',
                        credit_name='',
                        credit_amount=0,
                        revenue_type=f'{type_label}-미수금',
                        month=month,
                        partner_code=partner_code,
                        partner_name=partner_name
                    ))

                # 3) CR 일반진료수입: -본인부담금 (대변 마이너스 = 차감)
                if patient_amount > 0:
                    entries.append(JournalEntry(
                        date=entry_date,
                        description=f"{month} {type_label} 본인부담금 차감",
                        debit_account='',
                        debit_name='',
                        debit_amount=0,
                        credit_account=gen_code,
                        credit_name=gen_name,
                        credit_amount=-patient_amount,
                        revenue_type=f'{type_label}-본인부담차감',
                        month=month,
                        partner_code=partner_code,
                        partner_name=partner_name
                    ))

            # ── 입금 (지급일) ──
            if payment_date and payment_amount > 0:
                bank_code, bank_name = self.get_account('bank')
                tax_code, tax_name = self.get_account('prepaid_tax')
                res_code, res_name = self.get_account('prepaid_resident')

                # DR 보통예금
                entries.append(JournalEntry(
                    date=payment_date,
                    description=f"{month} {type_label} 공단입금",
                    debit_account=bank_code,
                    debit_name=bank_name,
                    debit_amount=payment_amount,
                    credit_account='',
                    credit_name='',
                    credit_amount=0,
                    revenue_type=f'{type_label}-입금',
                    month=month,
                    partner_code=partner_code,
                    partner_name=partner_name
                ))

                # DR 선납세금
                if income_tax > 0:
                    entries.append(JournalEntry(
                        date=payment_date,
                        description=f"{month} {type_label} 소득세 원천징수",
                        debit_account=tax_code,
                        debit_name=tax_name,
                        debit_amount=income_tax,
                        credit_account='',
                        credit_name='',
                        credit_amount=0,
                        revenue_type=f'{type_label}-원천징수',
                        month=month,
                        partner_code=partner_code,
                        partner_name=partner_name
                    ))

                # DR 선납주민세
                if resident_tax > 0:
                    entries.append(JournalEntry(
                        date=payment_date,
                        description=f"{month} {type_label} 주민세 원천징수",
                        debit_account=res_code,
                        debit_name=res_name,
                        debit_amount=resident_tax,
                        credit_account='',
                        credit_name='',
                        credit_amount=0,
                        revenue_type=f'{type_label}-원천징수',
                        month=month,
                        partner_code=partner_code,
                        partner_name=partner_name
                    ))

                # CR 미수금 회수
                ar_recover = payment_amount + income_tax + resident_tax
                entries.append(JournalEntry(
                    date=payment_date,
                    description=f"{month} {type_label} 미수금 회수",
                    debit_account='',
                    debit_name='',
                    debit_amount=0,
                    credit_account=ar_code,
                    credit_name=ar_name,
                    credit_amount=ar_recover,
                    revenue_type=f'{type_label}-미수금회수',
                    month=month,
                    partner_code=partner_code,
                    partner_name=partner_name
                ))

        entries.sort(key=lambda e: e.date)
        return entries

    def generate_cash_entries(self, cash_records):
        """카드/현금영수증/현금 매출 분개 생성"""
        entries = []
        gen_code, gen_name = self.get_account('rev_general')
        cash_code, cash_name = self.get_account('cash')
        bank_code, bank_name = self.get_account('bank')

        for rec in cash_records:
            month = rec.get('month', '')
            entry_date = self._month_to_last_date(month)

            # 카드매출 → DR 미수금(카드) / CR 일반진료수입
            card_amt = rec.get('card_amount', 0)
            if card_amt > 0:
                entries.append(JournalEntry(
                    date=entry_date,
                    description=f"{month} 카드매출",
                    debit_account='108',
                    debit_name='카드미수금',
                    debit_amount=card_amt,
                    credit_account=gen_code,
                    credit_name=gen_name,
                    credit_amount=card_amt,
                    revenue_type='카드매출',
                    month=month,
                ))

            # 현금영수증 → DR 현금 / CR 일반진료수입
            receipt_amt = rec.get('receipt_amount', 0)
            if receipt_amt > 0:
                entries.append(JournalEntry(
                    date=entry_date,
                    description=f"{month} 현금영수증매출",
                    debit_account=cash_code,
                    debit_name=cash_name,
                    debit_amount=receipt_amt,
                    credit_account=gen_code,
                    credit_name=gen_name,
                    credit_amount=receipt_amt,
                    revenue_type='현금영수증매출',
                    month=month,
                ))

            # 현금매출 → DR 현금 / CR 일반진료수입
            cash_amt = rec.get('cash_amount', 0)
            if cash_amt > 0:
                entries.append(JournalEntry(
                    date=entry_date,
                    description=f"{month} 현금매출",
                    debit_account=cash_code,
                    debit_name=cash_name,
                    debit_amount=cash_amt,
                    credit_account=gen_code,
                    credit_name=gen_name,
                    credit_amount=cash_amt,
                    revenue_type='현금매출',
                    month=month,
                ))

        entries.sort(key=lambda e: e.date)
        return entries

    def get_monthly_summary(self, entries):
        summary = {}
        for entry in entries:
            month = entry.month
            if not month:
                continue
            if month not in summary:
                summary[month] = {
                    'insurance': 0,
                    'medical_aid': 0,
                    'general_deduct': 0,
                    'ar_amount': 0,
                    'deposit': 0,
                    'tax_total': 0,
                    'card': 0,
                    'receipt': 0,
                    'cash_sale': 0,
                    'total': 0,
                }

            rt = entry.revenue_type
            if rt == '요양급여-매출':
                summary[month]['insurance'] += entry.credit_amount
            elif rt == '의료급여-매출':
                summary[month]['medical_aid'] += entry.credit_amount
            elif rt.endswith('-본인부담차감'):
                summary[month]['general_deduct'] += entry.credit_amount
            elif rt.endswith('-미수금') and entry.debit_amount > 0:
                summary[month]['ar_amount'] += entry.debit_amount
            elif rt.endswith('-입금'):
                summary[month]['deposit'] += entry.debit_amount
            elif rt.endswith('-원천징수'):
                summary[month]['tax_total'] += entry.debit_amount
            elif rt == '카드매출':
                summary[month]['card'] += entry.credit_amount
            elif rt == '현금영수증매출':
                summary[month]['receipt'] += entry.credit_amount
            elif rt == '현금매출':
                summary[month]['cash_sale'] += entry.credit_amount

        for month, data in summary.items():
            data['total'] = (data['insurance'] + data['medical_aid']
                            + data['card'] + data['receipt'] + data['cash_sale'])

        return dict(sorted(summary.items()))

    def _month_to_last_date(self, month_str):
        if not month_str or len(month_str) < 7:
            return month_str
        try:
            year = int(month_str[:4])
            month = int(month_str[5:7])
            last_day = calendar.monthrange(year, month)[1]
            return f"{year}-{month:02d}-{last_day:02d}"
        except (ValueError, IndexError):
            return month_str

    def validate_entries(self, entries):
        total_debit = sum(e.debit_amount for e in entries)
        total_credit = sum(e.credit_amount for e in entries)
        return {
            'balanced': total_debit == total_credit,
            'total_debit': total_debit,
            'total_credit': total_credit,
            'difference': total_debit - total_credit,
            'entry_count': len(entries)
        }


def format_won(amount):
    if amount is None:
        return "N/A"
    if amount < 0:
        return f"({abs(amount):,})"
    return f"{amount:,}"
