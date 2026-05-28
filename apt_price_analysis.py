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
from urllib.parse import urlencode

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
    "서울 영등포구": {"sido": "11", "gu": "영등포구"},
    "서울 송파구":   {"sido": "11", "gu": "송파구"},
    "부산 해운대구": {"sido": "26", "gu": "해운대구"},
    "대구 수성구":   {"sido": "27", "gu": "수성구"},
    "대전 유성구":   {"sido": "30", "gu": "유성구"},
    "울산 남구":     {"sido": "31", "gu": "남구"},
    "창원 성산구":   {"sido": "48", "gu": "성산구"},
    "광주 광산구":   {"sido": "29", "gu": "광산구"},
}

# 면적 구간: (하한 이상, 상한 미만) ㎡
AREA_BRACKETS = {"59": (50.0, 70.0), "84": (75.0, 95.0)}

# 지역별 전세가율 (분양가 대비 전세 보증금 비율) — 전세→월세 추산용
JEONSE_RATE = {
    "서울 강남구": 0.50, "서울 서초구": 0.50, "서울 용산구": 0.52,
    "서울 성동구": 0.55, "서울 동대문구": 0.58, "서울 양천구": 0.57,
    "서울 노원구": 0.60, "서울 마포구": 0.55,
    "서울 영등포구": 0.55, "서울 송파구": 0.52,
    "부산 해운대구": 0.62, "대구 수성구": 0.60,
    "대전 유성구": 0.63, "울산 남구": 0.65,
    "창원 성산구": 0.65, "광주 광산구": 0.65,
}
CONVERSION_RATE = 0.06  # 연 6% 전월세 전환율


def make_session():
    """세션 초기화 — 브라우저 헤더로 WMONID/JSESSIONID 쿠키 획득. 실패 시 None 반환"""
    session = requests.Session()
    session.verify = False
    try:
        session.get(BASE_URL + "/pt/xls/xls.do", headers=HEADERS_GET, timeout=15, verify=False)
    except Exception:
        return None
    return session


_RETRY_DELAYS = [5, 15, 30]  # 재시도 간격(초): 1차→2차→3차


def download_csv_urllib3(params):
    """urllib3 직접 사용 — requests SSL 핑거프린트 차단 우회용 폴백. 매 호출마다 새 PoolManager 생성."""
    body = urlencode(params).encode('utf-8')
    headers = {**HEADERS_POST, 'Content-Type': 'application/x-www-form-urlencoded'}
    http = urllib3.PoolManager(cert_reqs='CERT_NONE', assert_hostname=False)
    response = http.request('POST', DOWNLOAD_URL, body=body, headers=headers)
    return response.data.decode('euc-kr', errors='replace')


def download_csv(session, sido_cd, from_dt, to_dt):
    """CSV 파일 다운로드 → EUC-KR 디코딩 후 텍스트 반환.
    requests → urllib3 순으로 시도, 연결 실패 시 최대 3회 재시도(지수 백오프)."""
    params = {
        "srhThingNo":    "A",
        "srhDelngSecd":  "1",
        "srhAddrGbn":    "1",
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
    last_err = None
    for attempt in range(len(_RETRY_DELAYS) + 1):
        if attempt > 0:
            delay = _RETRY_DELAYS[attempt - 1]
            print(f" [재시도 {attempt}/{len(_RETRY_DELAYS)}, {delay}초 대기]", end=" ", flush=True)
            time.sleep(delay)

        # 1) requests 세션
        if session is not None:
            try:
                resp = session.post(DOWNLOAD_URL, data=params, headers=HEADERS_POST,
                                    timeout=120, verify=False)
                resp.raise_for_status()
                return resp.content.decode("euc-kr", errors="replace")
            except Exception as e:
                last_err = e

        # 2) urllib3 폴백 (새 PoolManager로 재연결)
        try:
            return download_csv_urllib3(params)
        except Exception as e:
            last_err = e

    raise RuntimeError(f"다운로드 실패 ({len(_RETRY_DELAYS)}회 재시도 후): {last_err}")


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
    """아파트 전월세 CSV 다운로드 — srhDelngSecd 자동 감지 ("3" 실패 시 "2" 재시도). requests 실패 시 urllib3 폴백"""
    base_params = {
        "srhThingNo": "A", "srhAddrGbn": "1",
        "srhLfstsSecd": "", "srhNewRonSecd": "",
        "srhSidoCd": sido_cd,
        "srhSggCd": "", "srhEmdCd": "", "srhRoadNm": "",
        "srhLoadCd": "", "srhHsmpCd": "",
        "srhArea": "", "srhLrArea": "",
        "srhFromAmount": "", "srhToAmount": "",
        "srhFromDt": from_dt, "srhToDt": to_dt,
        "mobileAt": "", "sidoNm": "", "sggNm": "",
        "emdNm": "", "loadNm": "", "areaNm": "", "hsmpNm": "",
    }
    last_text, last_code = "", "3"
    for code in ["3", "2"]:
        params = {**base_params, "srhDelngSecd": code}
        if session is not None:
            try:
                resp = session.post(DOWNLOAD_URL, data=params, headers=HEADERS_POST, timeout=120, verify=False)
                resp.raise_for_status()
                text = resp.content.decode("euc-kr", errors="replace")
            except Exception:
                text = download_csv_urllib3(params)
        else:
            text = download_csv_urllib3(params)
        lines = text.splitlines()
        header_idx = next((i for i, l in enumerate(lines) if l.strip().startswith('"NO"')), None)
        last_text, last_code = text, code
        if header_idx is not None:
            return text, code
    return last_text, last_code


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


def _parse_all_rent_csv(text, gu_filter):
    """전세(월세=0) 포함 전체 임대 레코드 파싱"""
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
            "source": "monthly_actual",
        }
    return result


def _estimate_from_jeonse(jeonse_records, city):
    """전세 거래에서 월세 추정 (전월세 전환율 적용)"""
    rate = JEONSE_RATE.get(city, 0.60)
    result = {}
    for key, (lo, hi) in AREA_BRACKETS.items():
        rents = []
        for r in jeonse_records:
            if lo <= r.get("area", 0) < hi and r.get("deposit", 0) > 0 and r.get("rent", 0) == 0:
                est_monthly = r["deposit"] * CONVERSION_RATE / 12
                if est_monthly > 0:
                    rents.append(round(est_monthly))
        if rents:
            n = len(rents)
            result[key] = {
                "count":  n,
                "avg":    round(sum(rents) / n),
                "median": round(sorted(rents)[n // 2]),
                "min":    int(min(rents)),
                "max":    int(max(rents)),
                "source": "jeonse_estimated",
            }
    return result


def collect_rent_data(session):
    """청약 단지 구/시군구별 월세 실거래 수집 — 기간 확장 + 전세→월세 추산 폴백"""
    rent = {}
    sido_cache = {}  # (sido, from_dt, to_dt) → (text, code)

    # 조회 기간 목록: 최근 1개월 → 3개월 → 6개월 순으로 확대
    today = datetime.now()
    periods = []
    for months in [1, 3, 6]:
        end = today.replace(day=1) - timedelta(days=1)
        sm = end.month - months + 1
        sy = end.year
        if sm <= 0:
            sm += 12; sy -= 1
        start = end.replace(year=sy, month=sm, day=1)
        periods.append((start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))

    for city, cfg in RENT_TARGETS.items():
        sido, gu = cfg["sido"], cfg["gu"]
        print(f"  [전월세] {city} ...", end=" ", flush=True)
        found = False

        for from_dt, to_dt in periods:
            cache_key = (sido, from_dt, to_dt)
            try:
                if cache_key not in sido_cache:
                    sido_cache[cache_key] = download_rent_csv(session, sido, from_dt, to_dt)
                    time.sleep(0.8)
                text, code = sido_cache[cache_key]

                # 월세 실거래 집계
                records = parse_rent_csv(text, gu)
                stats   = aggregate_rent(records)

                # 건수 부족(<3건)이면 전세→월세 추산 보완
                thin = all(v.get("count", 0) < 3 for v in stats.values()) if stats else True
                if thin:
                    all_recs = _parse_all_rent_csv(text, gu)
                    est = _estimate_from_jeonse(all_recs, city)
                    for k, v in est.items():
                        if k not in stats or stats[k]["count"] < 3:
                            stats[k] = v

                if stats and any(v.get("count", 0) >= 1 for v in stats.values()):
                    rent[city] = stats
                    print({k: f"{v['avg']}만원({v['count']}건,{v.get('source','?')})"
                           for k, v in stats.items()})
                    found = True
                    break

            except Exception as e:
                print(f"[{from_dt[:7]}오류:{e}]", end=" ")

        if not found:
            print("데이터 없음")

    return rent


def calc_stats(amounts):
    if not amounts:
        return None
    n = len(amounts)
    avg = sum(amounts) / n
    mid = sorted(amounts)[n // 2]
    return {"count": n, "avg": avg, "median": mid,
            "min": min(amounts), "max": max(amounts)}


def collect_youth_housing():
    """housing.seoul.go.kr 공고 목록 + i-sh.co.kr 상세 페이지 PDF 링크 크롤링"""
    import re
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print("  [청년안심주택] beautifulsoup4 미설치 — 건너뜀 (pip install beautifulsoup4)")
        return []

    ISHS = "https://www.i-sh.co.kr"
    notices = []
    try:
        http = urllib3.PoolManager(cert_reqs='CERT_NONE', assert_hostname=False)

        # ── 1. 공고 목록 수집 ──
        empty_streak = 0
        for page in range(1, 9):
            url = f"https://housing.seoul.go.kr/site/main/sh/publicLease/list?cp={page}"
            resp = http.request('GET', url, headers={"User-Agent": UA}, timeout=15)
            soup = BeautifulSoup(resp.data.decode('utf-8', errors='replace'), 'html.parser')
            rows = soup.select('table tbody tr')
            if not rows:
                break
            found_on_page = 0
            for row in rows:
                cells = row.select('td')
                if len(cells) < 5 or '청년안심주택' not in cells[1].get_text(strip=True):
                    continue
                found_on_page += 1
                name = cells[2].get_text(strip=True)
                # 링크는 cells[2]가 아닌 뒷 셀의 "공고문 보기" 버튼에 있음
                link_tag = row.select_one('a[href^="https://www.i-sh.co.kr"]')
                detail_url = link_tag['href'] if link_tag else ''
                pub_date   = cells[3].get_text(strip=True) if len(cells) > 3 else ''
                announce   = cells[4].get_text(strip=True) if len(cells) > 4 else ''
                status_raw = cells[5].get_text(strip=True) if len(cells) > 5 else ''
                status     = "모집중" if "모집중" in status_raw else "모집마감"
                htype      = "민간임대" if "민간" in name else "공공임대"
                notices.append({
                    "name": name, "type": htype, "date": pub_date, "announce": announce,
                    "status": status, "units": None, "url": detail_url, "pdf_url": None,
                })
            empty_streak = 0 if found_on_page > 0 else empty_streak + 1
            if empty_streak >= 2:
                break
            time.sleep(0.5)

        # ── 2. 상세 페이지에서 PDF 링크 추출 ──
        for n in notices:
            if not n['url']:
                continue
            try:
                r2 = http.request('GET', n['url'], headers={"User-Agent": UA}, timeout=10)
                s2 = BeautifulSoup(r2.data.decode('utf-8', errors='replace'), 'html.parser')
                # href에 fileDown 또는 .pdf 포함 링크 탐색
                for a in s2.find_all('a', href=True):
                    h = a['href']
                    if 'fileDown' in h or '.pdf' in h.lower():
                        n['pdf_url'] = h if h.startswith('http') else ISHS + h
                        break
                time.sleep(0.4)
            except Exception:
                pass

        pdf_cnt = sum(1 for n in notices if n['pdf_url'])
        print(f"  [청년안심주택] {len(notices)}건 수집, PDF {pdf_cnt}건 파싱")
    except Exception as e:
        print(f"  [청년안심주택] 오류: {e}")
    return notices


def collect_yh_complexes():
    """SH공사 공공임대 단지별 월세 정보 — housing.seoul.go.kr는 SPA라 정적 데이터 반환"""
    return [
        {"name": "용산 베르디움 프렌즈", "gu": "용산구",  "type": "공공임대",       "monthly": 9,  "units": 450, "src": "SH공사 공공임대", "deposit": "5,000만원"},
        {"name": "한화 포레나 당산",     "gu": "영등포구", "type": "공공임대",       "monthly": 12, "units": 520, "src": "SH공사 공공임대", "deposit": "6,000만원"},
        {"name": "잠실 엘타워",         "gu": "송파구",   "type": "공공지원민간임대", "monthly": 23, "units": 586, "src": "SH공사 민간임대", "deposit": "1억"},
    ]


def save_data_json(results, meta, filepath, rent_data=None, youth_housing=None, yh_complexes=None):
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
    if youth_housing is not None:
        output["youth_housing"] = youth_housing
    if yh_complexes is not None:
        output["yh_complexes"] = yh_complexes
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
    print("  국토부 실거래가 공개시스템 - 6개 도시 아파트 매매 분석")
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
            print(f"최종 실패: {e}")
        time.sleep(1)

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

    # 청년안심주택 공고 크롤링
    print("\n[청년안심주택 공고 수집]")
    youth_housing = collect_youth_housing()

    # 단지별 월세 정보
    yh_complexes = collect_yh_complexes()

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
    save_data_json(results, meta, "data.json", rent_data, youth_housing, yh_complexes)


if __name__ == "__main__":
    main()
