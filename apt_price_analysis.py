#!/usr/bin/env python3
"""
국토부 실거래가 공개시스템 CSV 직접 다운로드 기반
6개 도시 아파트 평균 매매가 분석 (API 키 불필요)

다운로드 출처: https://rt.molit.go.kr/pt/xls/xls.do
"""

import requests
import urllib3
import csv
import io
import json
import time
from datetime import datetime, timedelta

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://rt.molit.go.kr"
DOWNLOAD_URL = BASE_URL + "/pt/xls/ptXlsCSVDown.do"

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

HEADERS_GET = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-Dest": "document",
    "Upgrade-Insecure-Requests": "1",
}

HEADERS_POST = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin": BASE_URL,
    "Referer": BASE_URL + "/pt/xls/xls.do",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Dest": "document",
    "Upgrade-Insecure-Requests": "1",
}

# 조회 기간: 전월 말일 기준 최근 6개월 (실행 시점에 자동 계산)
def _compute_date_range():
    today = datetime.now()
    to  = today.replace(day=1) - timedelta(days=1)   # 전월 말일
    m, y = to.month - 5, to.year
    if m <= 0:
        m, y = m + 12, y - 1
    frm = to.replace(year=y, month=m, day=1)
    return frm.strftime("%Y-%m-%d"), to.strftime("%Y-%m-%d")

FROM_DT, TO_DT = _compute_date_range()

# 6개 도시 설정: (sido코드, 필터 키워드 or None)
# 창원은 경남(48)을 받아 '창원시'로 필터
CITIES = {
    "부산": {"sido": "26", "filter": None},
    "대구": {"sido": "27", "filter": None},
    "광주": {"sido": "29", "filter": None},
    "대전": {"sido": "30", "filter": None},
    "울산": {"sido": "31", "filter": None},
    "창원": {"sido": "48", "filter": "창원시"},
}

# 전월세 수집 대상: 청약 단지별 구/시군구
# srhDelngSecd="3" = 전월세(임대). 국토부 사이트 변경 시 "2" 로 조정 필요.
RENT_TARGETS = {
    "서울 강남구":   {"sido": "11", "gu": "강남구"},
    "서울 서초구":   {"sido": "11", "gu": "서초구"},
    "서울 용산구":   {"sido": "11", "gu": "용산구"},
    "서울 성동구":   {"sido": "11", "gu": "성동구"},
    "서울 동대문구": {"sido": "11", "gu": "동대문구"},
    "서울 양천구":   {"sido": "11", "gu": "양천구"},
    "서울 노원구":   {"sido": "11", "gu": "노원구"},
    "서울 마포구":   {"sido": "11", "gu": "마포구"},
    "부산 해운대구": {"sido": "26", "gu": "해운대구"},
    "대구 수성구":   {"sido": "27", "gu": "수성구"},
    "대전 유성구":   {"sido": "30", "gu": "유성구"},
    "울산 남구":     {"sido": "31", "gu": "남구"},
    "창원 성산구":   {"sido": "48", "gu": "성산구"},
    "광주 광산구":   {"sido": "29", "gu": "광산구"},
}

# 면적 구간: (하한 이상, 상한 미만) ㎡
AREA_BRACKETS = {"59": (50.0, 70.0), "84": (75.0, 95.0)}


def make_session():
    """세션 초기화 — 브라우저 헤더로 WMONID/JSESSIONID 쿠키 획득"""
    session = requests.Session()
    session.get(BASE_URL + "/pt/xls/xls.do", headers=HEADERS_GET, timeout=15, verify=False)
    return session


def download_csv(session, sido_cd, from_dt, to_dt):
    """CSV 파일 다운로드 → EUC-KR 디코딩 후 텍스트 반환"""
    params = {
        "srhThingNo":    "A",   # 아파트
        "srhDelngSecd":  "1",   # 매매
        "srhAddrGbn":    "1",   # 지번주소
        "srhLfstsSecd":  "1",
        "srhNewRonSecd": "",
        "srhSidoCd":     sido_cd,
        "srhSggCd":      "",
        "srhEmdCd":      "",
        "srhRoadNm":     "",
        "srhLoadCd":     "",
        "srhHsmpCd":     "",
        "srhArea":       "",
        "srhLrArea":     "",
        "srhFromAmount": "",
        "srhToAmount":   "",
        "srhFromDt":     from_dt,
        "srhToDt":       to_dt,
        "mobileAt":      "",
        "sidoNm":        "",
        "sggNm":         "",
        "emdNm":         "",
        "loadNm":        "",
        "areaNm":        "",
        "hsmpNm":        "",
    }
    resp = session.post(DOWNLOAD_URL, data=params, headers=HEADERS_POST, timeout=120, verify=False)
    resp.raise_for_status()
    return resp.content.decode("euc-kr", errors="replace")


def parse_csv(text, city_filter=None):
    """
    CSV 파싱 → (거래금액 리스트, 레코드 리스트) 반환
    - 앞부분 공지/검색조건 행 건너뜀
    - 컬럼 헤더 행(NO,시군구,...) 기준으로 데이터 읽기
    - city_filter가 있으면 '시군구' 컬럼에 해당 문자열 포함 행만 선택
    """
    lines = text.splitlines()

    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('"NO"'):
            header_idx = i
            break

    if header_idx is None:
        return [], []

    data_text = "\n".join(lines[header_idx:])
    reader = csv.DictReader(io.StringIO(data_text))

    amounts = []
    records = []
    for row in reader:
        sigungu = row.get("시군구", "")
        if city_filter and city_filter not in sigungu:
            continue

        # 해제 거래 제외 (해제사유발생일이 날짜 값이면 취소된 거래)
        if row.get("해제사유발생일", "-").strip() not in ("-", ""):
            continue

        raw = row.get("거래금액(만원)", "").replace(",", "").strip()
        if not raw or not raw.lstrip("-").isdigit():
            continue

        amount = float(raw)
        amounts.append(amount)

        area_raw = row.get("전용면적(㎡)", "").strip()
        records.append({
            "시군구":       sigungu,
            "단지명":       row.get("단지명", "").strip(),
            "전용면적_㎡":  float(area_raw) if area_raw else None,
            "계약년월":     row.get("계약년월", "").strip(),
            "계약일":       row.get("계약일", "").strip(),
            "거래금액_만원": int(amount),
            "층":           row.get("층", "").strip(),
            "건축년도":     row.get("건축년도", "").strip(),
            "거래유형":     row.get("거래유형", "").strip(),
            "도로명":       row.get("도로명", "").strip(),
        })

    return amounts, records


def download_rent_csv(session, sido_cd, from_dt, to_dt):
    """아파트 전월세 CSV 다운로드 (월세 건 포함)"""
    params = {
        "srhThingNo":    "A",   # 아파트
        "srhDelngSecd":  "3",   # 전월세(임대) — 매매는 "1"
        "srhAddrGbn":    "1",
        "srhLfstsSecd":  "",
        "srhNewRonSecd": "",
        "srhSidoCd":     sido_cd,
        "srhSggCd": "", "srhEmdCd": "", "srhRoadNm": "",
        "srhLoadCd": "", "srhHsmpCd": "",
        "srhArea": "", "srhLrArea": "",
        "srhFromAmount": "", "srhToAmount": "",
        "srhFromDt":     from_dt,
        "srhToDt":       to_dt,
        "mobileAt": "", "sidoNm": "", "sggNm": "",
        "emdNm": "", "loadNm": "", "areaNm": "", "hsmpNm": "",
    }
    resp = session.post(DOWNLOAD_URL, data=params, headers=HEADERS_POST, timeout=120, verify=False)
    resp.raise_for_status()
    return resp.content.decode("euc-kr", errors="replace")


def parse_rent_csv(text, gu_filter):
    """전월세 CSV 파싱 → 월세(>0) 레코드만 반환"""
    lines = text.splitlines()
    header_idx = next((i for i, l in enumerate(lines) if l.strip().startswith('"NO"')), None)
    if header_idx is None:
        return []
    reader = csv.DictReader(io.StringIO("\n".join(lines[header_idx:])))
    records = []
    for row in reader:
        if gu_filter and gu_filter not in row.get("시군구", ""):
            continue
        try:
            area = float(row.get("전용면적(㎡)", "").strip())
            rent = float((row.get("월세(만원)", "") or "0").replace(",", "").strip())
            dep  = float((row.get("보증금(만원)", "") or "0").replace(",", "").strip())
        except (ValueError, AttributeError):
            continue
        if rent <= 0:
            continue  # 전세(월세=0) 제외
        records.append({"area": area, "rent": rent, "deposit": dep,
                        "ym": row.get("계약년월", "").strip()})
    return records


def aggregate_rent(records):
    """면적 구간(59형/84형)별 월세 통계 계산"""
    buckets = {k: [] for k in AREA_BRACKETS}
    for r in records:
        for key, (lo, hi) in AREA_BRACKETS.items():
            if lo <= r["area"] < hi:
                buckets[key].append(r["rent"])
    result = {}
    for key, rents in buckets.items():
        if not rents:
            continue
        n = len(rents)
        result[key] = {
            "count":  n,
            "avg":    round(sum(rents) / n),
            "median": round(sorted(rents)[n // 2]),
            "min":    int(min(rents)),
            "max":    int(max(rents)),
        }
    return result


def collect_rent_data(session):
    """청약 단지 구/시군구별 최근 6개월 월세 실거래 수집"""
    rent = {}
    sido_cache = {}  # 같은 sido는 한 번만 다운로드

    for city, cfg in RENT_TARGETS.items():
        sido, gu = cfg["sido"], cfg["gu"]
        print(f"  [전월세] {city} ...", end=" ", flush=True)
        try:
            if sido not in sido_cache:
                sido_cache[sido] = download_rent_csv(session, sido, FROM_DT, TO_DT)
                time.sleep(1)
            records = parse_rent_csv(sido_cache[sido], gu)
            stats   = aggregate_rent(records)
            if stats:
                rent[city] = stats
                print({k: f"{v['avg']}만원({v['count']}건)" for k, v in stats.items()})
            else:
                print("월세 데이터 없음")
        except Exception as e:
            print(f"오류: {e}")
    return rent


def calc_stats(amounts):
    if not amounts:
        return None
    n = len(amounts)
    avg = sum(amounts) / n
    mid = sorted(amounts)[n // 2]
    return {"count": n, "avg": avg, "median": mid,
            "min": min(amounts), "max": max(amounts)}


def save_data_json(results, meta, filepath, rent_data=None):
    """GitHub Pages용 경량 JSON (transactions 제외, HTML이 기대하는 영문 키)"""
    summary = {
        city: {
            "count":  s["count"],
            "avg":    round(s["avg"], 1),
            "median": round(s["median"], 1),
            "min":    int(s["min"]),
            "max":    int(s["max"]),
        }
        for city, s in results.items()
    }
    output = {"meta": meta, "summary": summary}
    if rent_data:
        output["rent"] = rent_data
        output["rent_updated"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  data.json 저장 완료: {filepath}")


def save_json(results, all_records, filepath):
    """분석 결과와 전체 실거래 레코드를 JSON으로 저장"""
    output = {
        "meta": {
            "source":        "국토교통부 실거래가 공개시스템",
            "url":           "https://rt.molit.go.kr",
            "period_from":   FROM_DT,
            "period_to":     TO_DT,
            "downloaded_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "unit":          "만원",
        },
        "summary": {
            city: {
                "거래건수":     s["count"],
                "평균가_만원":  round(s["avg"], 1),
                "중위가_만원":  round(s["median"], 1),
                "최저가_만원":  int(s["min"]),
                "최고가_만원":  int(s["max"]),
            }
            for city, s in sorted(results.items(), key=lambda x: x[1]["avg"], reverse=True)
        },
        "transactions": all_records,
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    total = sum(len(v) for v in all_records.values())
    print(f"\n  JSON 저장 완료: {filepath}")
    print(f"  총 {total:,}건 레코드 저장됨")


def fmt(amount):
    """만원 → 억/만원 문자열"""
    if amount >= 10000:
        eok = int(amount // 10000)
        man = int(amount % 10000)
        return f"{eok}억 {man:,}만원" if man else f"{eok}억원"
    return f"{int(amount):,}만원"


def main():
    print("=" * 70)
    print("  국토부 실거래가 공개시스템 — 6개 도시 아파트 매매 분석")
    print(f"  기간: {FROM_DT} ~ {TO_DT}")
    print("=" * 70)

    session = make_session()
    results = {}
    all_records = {}

    for city, cfg in CITIES.items():
        print(f"\n[{city}] 다운로드 중...", end=" ", flush=True)
        try:
            text = download_csv(session, cfg["sido"], FROM_DT, TO_DT)
            amounts, records = parse_csv(text, city_filter=cfg["filter"])
            stats = calc_stats(amounts)
            if stats:
                results[city] = stats
                all_records[city] = records
                print(f"{stats['count']:,}건 수집")
            else:
                print("데이터 없음")
        except Exception as e:
            print(f"오류: {e}")
        time.sleep(1)  # 서버 부하 방지

    if not results:
        print("\n수집된 데이터가 없습니다.")
        return

    # 결과 출력
    print()
    print("=" * 70)
    print(f"  6개 도시 아파트 매매 평균가 (단위: 만원)")
    print(f"  기준: {FROM_DT[:7]} ~ {TO_DT[:7]}")
    print("=" * 70)
    print(f"{'순위':<4} {'도시':<5} {'거래건수':>8}  {'평균 매매가':>17}  {'중위 매매가':>17}")
    print("-" * 70)

    sorted_r = sorted(results.items(), key=lambda x: x[1]["avg"], reverse=True)
    for rank, (city, s) in enumerate(sorted_r, 1):
        print(f"{rank:<4} {city:<5} {s['count']:>7,}건  "
              f"{fmt(s['avg']):>17}  "
              f"{fmt(s['median']):>17}")

    print("=" * 70)
    print()

    # 최고/최저 도시
    top = sorted_r[0]
    bot = sorted_r[-1]
    print(f"  ▶ 최고 평균가: {top[0]} ({fmt(top[1]['avg'])})")
    print(f"  ▶ 최저 평균가: {bot[0]} ({fmt(bot[1]['avg'])})")
    ratio = top[1]['avg'] / bot[1]['avg']
    print(f"  ▶ 최고/최저 비율: {ratio:.2f}배")
    print()
    print("  * 출처: 국토교통부 실거래가 공개시스템 (rt.molit.go.kr)")
    print("  * 계약일 기준 데이터 / 해제 거래 포함될 수 있음")

    # 전월세 실거래 수집
    print("\n[전월세 실거래 데이터 수집]")
    print(f"  대상: {len(RENT_TARGETS)}개 구/시군구 · 기간: {FROM_DT} ~ {TO_DT}")
    rent_data = collect_rent_data(session)
    if rent_data:
        print(f"  전월세 수집 완료: {len(rent_data)}개 지역")
    else:
        print("  전월세 데이터 없음 (파라미터 확인 필요 — srhDelngSecd 값 조정)")

    # JSON 저장
    meta = {
        "source":        "국토교통부 실거래가 공개시스템",
        "url":           "https://rt.molit.go.kr",
        "period_from":   FROM_DT,
        "period_to":     TO_DT,
        "downloaded_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "unit":          "만원",
    }
    save_json(results, all_records, "apt_transactions.json")
    save_data_json(results, meta, "data.json", rent_data)


if __name__ == "__main__":
    main()
