# 세무 업무 자동화 프로그램

세무사 박양훈이 세무 실무에서 반복되는 작업을 자동화하기 위해 직접 개발한 프로그램 모음입니다.

## 실행 방법

### 사전 준비
- Python 3.10 이상 설치 ([python.org](https://www.python.org/downloads/))
- Git 설치 후 레포 클론:
```bash
git clone https://github.com/uosac93/tax-automation.git
cd tax-automation
```

### 각 프로그램 실행

**1. 병의원 매출 자동 회계처리**
```bash
cd 1_medical-revenue
pip install -r requirements.txt
python main.py
```

**2. 법인세 신고서 자동 검토**
```bash
cd 2_corp-tax-review
pip install -r requirements.txt
python main.py
```

**3. 기준시가 계산기**
```bash
cd 3_standard-price-calculator
pip install -r requirements.txt
python app.pyw
```
> 기준시가 계산기는 공공데이터포털 API 키가 필요합니다.
> `index.html`과 `building_land_registry.py`의 `YOUR_API_KEY` 부분을 본인의 API 키로 교체하세요.

## 프로그램 목록

### 1. 병의원 매출 자동 회계처리 (`1_medical-revenue`)
- **기술**: Python (PDF 파싱, 분개 엔진, Excel/CSV 출력)
- 요양기관정보마당 지급통보서 PDF → 자동 분개 생성
- 더존 Wehago/Douzen ERP import용 CSV 출력

### 2. 법인세 신고서 자동 검토 (`2_corp-tax-review`)
- **기술**: Python (PDF 파싱, docx/Excel/PDF 보고서 생성)
- 법인세 신고서 PDF 분석 → 세액공제·감면 항목 자동 검증
- 세액공제·세액감면 적용 여부 대조, 다중 포맷 보고서 생성

### 3. 기준시가 계산기 + 건축물대장 조회 (`3_standard-price-calculator`)
- **기술**: Python + JavaScript (하이브리드 데스크톱 앱)
- 오피스텔·상업용 건물 기준시가 산출
- 공공데이터포털 건축물대장 API 연동, 국세청 XML 기준데이터 기반 실시간 계산

## 기술 스택
- **언어**: Python 3, JavaScript, HTML/CSS
- **라이브러리**: openpyxl, python-docx, PyInstaller, Tkinter, PyMuPDF, reportlab
- **API**: 공공데이터포털, Claude API
- **기타**: PyWebView, customtkinter

## 개발자
- 박양훈 (세무사)
- uosac93@gmail.com
