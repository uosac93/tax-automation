"""
건축물대장 / 토지대장 자동 조회 스크립트
- 주소 입력 → API 조회 → 결과 출력
- 공공데이터포털 API 키 필요 (data.go.kr)

사용법:
    python building_land_registry.py "서울특별시 강남구 테헤란로 123"
    python building_land_registry.py --pnu 1168010100100010000
"""

import requests
import json
import sys
import os
from urllib.parse import quote

# ============================================================
# API 키 설정 - 여기에 발급받은 키를 입력하세요
# ============================================================
API_KEY = os.environ.get("DATA_GO_KR_API_KEY", "YOUR_DATA_GO_KR_API_KEY")

# API 엔드포인트
BUILDING_API = "http://apis.data.go.kr/1613000/BldRgstHubService"
LAND_API = "http://apis.data.go.kr/1160100/service/GetLandInfoService"
JUSO_API = "https://business.juso.go.kr/addrlink/addrLinkApi.do"

# 주소 API 키 (도로명주소 → PNU 변환용, 별도 신청 필요 시 아래 키 사용)
JUSO_API_KEY = os.environ.get("JUSO_API_KEY", "")



def search_building_by_address(sigungu_cd, bjdong_cd, bun="", ji=""):
    """건축물대장 기본개요 조회"""
    url = f"{BUILDING_API}/getBrBasisOulnInfo"
    params = {
        "serviceKey": API_KEY,
        "sigunguCd": sigungu_cd,
        "bjdongCd": bjdong_cd,
        "numOfRows": "10",
        "pageNo": "1",
        "type": "json",
    }
    if bun:
        params["platGbCd"] = "0"  # 0:대지, 1:산
        params["bun"] = bun.zfill(4)
    if ji:
        params["ji"] = ji.zfill(4)

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        items = data.get("response", {}).get("body", {}).get("items", {}).get("item", [])
        if isinstance(items, dict):
            items = [items]
        return items
    except Exception as e:
        print(f"[건축물대장 조회 오류] {e}")
        return []


def search_building_title(sigungu_cd, bjdong_cd, bun="", ji=""):
    """건축물대장 표제부 조회"""
    url = f"{BUILDING_API}/getBrTitleInfo"
    params = {
        "serviceKey": API_KEY,
        "sigunguCd": sigungu_cd,
        "bjdongCd": bjdong_cd,
        "numOfRows": "10",
        "pageNo": "1",
        "type": "json",
    }
    if bun:
        params["platGbCd"] = "0"
        params["bun"] = bun.zfill(4)
    if ji:
        params["ji"] = ji.zfill(4)

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        items = data.get("response", {}).get("body", {}).get("items", {}).get("item", [])
        if isinstance(items, dict):
            items = [items]
        return items
    except Exception as e:
        print(f"[건축물대장 표제부 조회 오류] {e}")
        return []


def search_building_floor(sigungu_cd, bjdong_cd, bun="", ji=""):
    """건축물대장 층별개요 조회"""
    url = f"{BUILDING_API}/getBrFlrOulnInfo"
    params = {
        "serviceKey": API_KEY,
        "sigunguCd": sigungu_cd,
        "bjdongCd": bjdong_cd,
        "numOfRows": "50",
        "pageNo": "1",
        "type": "json",
    }
    if bun:
        params["platGbCd"] = "0"
        params["bun"] = bun.zfill(4)
    if ji:
        params["ji"] = ji.zfill(4)

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        items = data.get("response", {}).get("body", {}).get("items", {}).get("item", [])
        if isinstance(items, dict):
            items = [items]
        return items
    except Exception as e:
        print(f"[건축물대장 층별 조회 오류] {e}")
        return []


def search_land_info(pnu):
    """토지대장 조회 (토지e음 연계)"""
    url = f"{LAND_API}/getLandInfoItem"
    params = {
        "serviceKey": API_KEY,
        "pnu": pnu,
        "numOfRows": "10",
        "pageNo": "1",
        "type": "json",
        "stdrYear": "2025",
    }

    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        items = data.get("response", {}).get("body", {}).get("items", {}).get("item", [])
        if isinstance(items, dict):
            items = [items]
        return items
    except Exception as e:
        print(f"[토지대장 조회 오류] {e}")
        return []


def format_building_result(items):
    """건축물대장 결과 포맷팅"""
    if not items:
        return "건축물대장 조회 결과 없음"

    results = []
    for item in items:
        info = {
            "대장종류": item.get("regstrKindCdNm", ""),
            "대장구분": item.get("regstrGbCdNm", ""),
            "대지위치": item.get("platPlc", ""),
            "도로명주소": item.get("newPlatPlc", ""),
            "건물명": item.get("bldNm", ""),
            "주용도": item.get("mainPurpsCdNm", ""),
            "구조": item.get("strctCdNm", ""),
            "지붕": item.get("roofCdNm", ""),
            "대지면적(㎡)": item.get("platArea", ""),
            "건축면적(㎡)": item.get("archArea", ""),
            "건폐율(%)": item.get("bcRat", ""),
            "연면적(㎡)": item.get("totArea", ""),
            "용적률(%)": item.get("vlRat", ""),
            "층수(지상)": item.get("grndFlrCnt", ""),
            "층수(지하)": item.get("ugrndFlrCnt", ""),
            "사용승인일": item.get("useAprDay", ""),
            "허가일": item.get("pmsDay", ""),
        }
        # 빈 값 제거
        info = {k: v for k, v in info.items() if v}
        results.append(info)

    return results


def format_land_result(items):
    """토지대장 결과 포맷팅"""
    if not items:
        return "토지대장 조회 결과 없음"

    jimok_map = {
        "01": "전", "02": "답", "03": "과수원", "04": "목장용지",
        "05": "임야", "06": "광천지", "07": "염전", "08": "대",
        "09": "공장용지", "10": "학교용지", "11": "주차장", "12": "주유소용지",
        "13": "창고용지", "14": "도로", "15": "철도용지", "16": "하천",
        "17": "제방", "18": "구거", "19": "유지", "20": "양어장",
        "21": "수도용지", "22": "공원", "23": "체육용지", "24": "유원지",
        "25": "종교용지", "26": "사적지", "27": "묘지", "28": "잡종지",
    }

    results = []
    for item in items:
        jimok_cd = str(item.get("lndcgrCd", ""))
        info = {
            "PNU": item.get("pnu", ""),
            "소재지": item.get("ldCodeNm", ""),
            "지목코드": jimok_cd,
            "지목": jimok_map.get(jimok_cd.zfill(2), jimok_cd),
            "면적(㎡)": item.get("lndpclAr", ""),
            "공시지가(원/㎡)": item.get("pblntfPclnd", ""),
            "기준연도": item.get("stdrYear", ""),
        }
        info = {k: v for k, v in info.items() if v}
        results.append(info)

    return results


def query_all(sigungu_cd, bjdong_cd, bun="", ji="", pnu=""):
    """건축물대장 + 토지대장 통합 조회"""
    print(f"\n{'='*60}")
    print(f"  건축물대장/토지대장 조회")
    print(f"  시군구코드: {sigungu_cd} | 법정동코드: {bjdong_cd}")
    if bun:
        print(f"  본번: {bun} | 부번: {ji or '0'}")
    print(f"{'='*60}\n")

    # 1. 건축물대장 기본개요
    print("[1] 건축물대장 기본개요 조회 중...")
    basis = search_building_by_address(sigungu_cd, bjdong_cd, bun, ji)
    basis_result = format_building_result(basis)
    if isinstance(basis_result, list):
        for i, item in enumerate(basis_result, 1):
            print(f"\n  --- 건물 {i} ---")
            for k, v in item.items():
                print(f"  {k}: {v}")
    else:
        print(f"  {basis_result}")

    # 2. 건축물대장 표제부
    print("\n[2] 건축물대장 표제부 조회 중...")
    title = search_building_title(sigungu_cd, bjdong_cd, bun, ji)
    title_result = format_building_result(title)
    if isinstance(title_result, list):
        for i, item in enumerate(title_result, 1):
            print(f"\n  --- 표제부 {i} ---")
            for k, v in item.items():
                print(f"  {k}: {v}")
    else:
        print(f"  {title_result}")

    land_result = []

    # 3. 토지대장 (PNU 필요)
    if pnu:
        print("\n[3] 토지대장 조회 중...")
        land = search_land_info(pnu)
        land_result = format_land_result(land)
        if isinstance(land_result, list):
            for i, item in enumerate(land_result, 1):
                print(f"\n  --- 토지 {i} ---")
                for k, v in item.items():
                    print(f"  {k}: {v}")
        else:
            print(f"  {land_result}")

    print(f"\n{'='*60}")
    print("  조회 완료")
    print(f"{'='*60}\n")

    return {
        "building_basis": basis_result if isinstance(basis_result, list) else [],
        "building_title": title_result if isinstance(title_result, list) else [],
        "land": land_result if pnu and isinstance(land_result, list) else [],
    }


# ============================================================
# 주요 지역 법정동코드 (자주 쓰는 지역)
# 전체 목록: https://www.code.go.kr
# ============================================================
SAMPLE_CODES = {
    "서울 강남구 역삼동": {"sigungu": "11680", "bjdong": "10300"},
    "서울 강남구 삼성동": {"sigungu": "11680", "bjdong": "10500"},
    "서울 강남구 대치동": {"sigungu": "11680", "bjdong": "10800"},
    "서울 서초구 서초동": {"sigungu": "11650", "bjdong": "10300"},
    "서울 송파구 잠실동": {"sigungu": "11710", "bjdong": "10600"},
    "서울 중구 명동": {"sigungu": "11140", "bjdong": "10200"},
    "서울 종로구 종로": {"sigungu": "11110", "bjdong": "15400"},
    "서울 마포구 서교동": {"sigungu": "11440", "bjdong": "10100"},
}


if __name__ == "__main__":
    if API_KEY == "YOUR_DATA_GO_KR_API_KEY":
        print("=" * 60)
        print("  API 키가 설정되지 않았습니다!")
        print()
        print("  1. data.go.kr 에서 회원가입")
        print("  2. '건축물대장정보 서비스' API 활용신청")
        print("  3. 발급받은 키를 이 파일의 API_KEY에 입력")
        print("     또는 환경변수 설정:")
        print("     set DATA_GO_KR_API_KEY=발급받은키")
        print("=" * 60)
        sys.exit(1)

    if len(sys.argv) > 1 and sys.argv[1] == "--pnu":
        # PNU 직접 입력
        pnu = sys.argv[2] if len(sys.argv) > 2 else ""
        sigungu = pnu[:5]
        bjdong = pnu[5:10]
        bun = pnu[11:15] if len(pnu) > 14 else ""
        ji = pnu[15:19] if len(pnu) > 18 else ""
        query_all(sigungu, bjdong, bun, ji, pnu)
    elif len(sys.argv) > 1:
        # 시군구코드 법정동코드 본번 부번 형태
        print("사용법:")
        print('  python building_land_registry.py --pnu 1168010100100010000')
        print()
        print("또는 Claude에게 주소를 알려주시면 코드를 찾아서 조회합니다.")
        print()
        print("등록된 샘플 지역:")
        for name, codes in SAMPLE_CODES.items():
            print(f"  {name}: 시군구={codes['sigungu']}, 법정동={codes['bjdong']}")
    else:
        # 대화형 모드
        print("건축물대장/토지대장 조회 시스템")
        print("-" * 40)
        sigungu = input("시군구코드 (5자리): ").strip()
        bjdong = input("법정동코드 (5자리): ").strip()
        bun = input("본번 (없으면 Enter): ").strip()
        ji = input("부번 (없으면 Enter): ").strip()
        pnu = input("PNU 19자리 (토지대장용, 없으면 Enter): ").strip()
        query_all(sigungu, bjdong, bun, ji, pnu)
