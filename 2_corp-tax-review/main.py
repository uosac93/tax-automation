"""
법인세 신고서 자동 검토 프로그램
사용법: python main.py <PDF파일경로>
"""
import sys
import os
from datetime import datetime

# 모듈 임포트
from pdf_parser import parse_all
from tax_reviewer import run_all_reviews
from report_generator import print_console_report, generate_excel_report


def main():
    # PDF 경로 확인
    if len(sys.argv) < 2:
        # 바탕화면 '법인세 검토' 폴더 + 현재 폴더에서 PDF 자동 탐색
        search_dirs = ["."]
        # 바탕화면 Corp Tax_review 폴더 우선 탐색
        corp_tax_dir = os.path.expanduser(r"~\Desktop\Corp Tax_review")
        if os.path.isdir(corp_tax_dir):
            search_dirs.insert(0, corp_tax_dir)
        # 기존 법인세 검토 폴더도 탐색
        desktop_dir = os.path.expanduser(r"~\Desktop\Claude Code\법인세 검토")
        if os.path.isdir(desktop_dir):
            search_dirs.insert(0, desktop_dir)
        pdf_files = []
        for d in search_dirs:
            for f in os.listdir(d):
                if f.endswith(".pdf"):
                    pdf_files.append(os.path.join(d, f))
        if len(pdf_files) == 1:
            pdf_path = pdf_files[0]
            print(f"  자동 탐색된 PDF: {pdf_path}")
        elif len(pdf_files) > 1:
            print("  여러 PDF 파일이 있습니다:")
            for i, f in enumerate(pdf_files, 1):
                print(f"    {i}. {f}")
            choice = input("  번호를 선택하세요: ").strip()
            try:
                pdf_path = pdf_files[int(choice) - 1]
            except (ValueError, IndexError):
                print("  잘못된 선택입니다.")
                sys.exit(1)
        else:
            print("사용법: python main.py <PDF파일경로>")
            print("  또는 현재 폴더에 PDF 파일을 넣어주세요.")
            sys.exit(1)
    else:
        pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"오류: 파일을 찾을 수 없습니다 - {pdf_path}")
        sys.exit(1)

    print("\n" + "=" * 80)
    print("  법인세 신고서 자동 검토 시스템")
    print("=" * 80)
    print(f"\n  대상 파일: {os.path.basename(pdf_path)}")
    print(f"  검토 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Step 1: PDF 파싱
    print("\n  [1/3] PDF 데이터 추출 중...")
    try:
        data = parse_all(pdf_path)
        print(f"    → {data['총페이지수']}페이지 처리 완료")
        print(f"    → 법인명: {data['회사정보'].get('법인명', 'N/A')}")
    except Exception as e:
        print(f"  오류: PDF 파싱 실패 - {e}")
        sys.exit(1)

    # Step 2: 검토 실행
    print("\n  [2/3] 세무 검토 수행 중...")
    results = run_all_reviews(data)
    이슈수 = sum(1 for r in results if r.상태 == "이슈")
    정상수 = sum(1 for r in results if r.상태 == "정상")
    확인필요수 = sum(1 for r in results if r.상태 == "확인필요")
    print(f"    → {len(results)}개 항목 검토 완료")
    print(f"    → 정상: {정상수} / 확인필요: {확인필요수} / 이슈: {이슈수}")

    # Step 3: 리포트 생성
    print("\n  [3/3] 검토 보고서 생성 중...")

    # 콘솔 출력
    print_console_report(data, results)

    # Excel 보고서
    법인명 = data["회사정보"].get("법인명", "검토대상")
    날짜 = datetime.now().strftime("%Y%m%d")
    excel_filename = f"법인세검토결과_{법인명}_{날짜}.xlsx"
    excel_path = os.path.join(os.path.dirname(pdf_path) or ".", excel_filename)

    try:
        generate_excel_report(data, results, excel_path)
    except Exception as e:
        print(f"  Excel 생성 오류: {e}")
        # 현재 폴더에 시도
        excel_path = excel_filename
        try:
            generate_excel_report(data, results, excel_path)
        except Exception as e2:
            print(f"  Excel 생성 실패: {e2}")

    print(f"\n  검토 완료: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    # GUI 모드 (기본) vs CLI 모드 (--cli 인수)
    if "--cli" in sys.argv:
        sys.argv.remove("--cli")
        main()
    else:
        from gui_app import App
        app = App()
        app.mainloop()
