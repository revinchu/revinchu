#!/usr/bin/env python3
"""
네이버 파워링크 광고 순위 체크 프로그램
- 키워드 리스트 파일을 읽어서 각 키워드별로 네이버 광고 검색
- 특정 업체명(검색어)이 PC/모바일 각각 몇 순위에 있는지 확인
"""

import argparse
import csv
import sys
import time
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup


# --- 설정 ---
PC_URL_TEMPLATE = (
    "https://ad.search.naver.com/search.naver?where=ad&query={query}"
)
MOBILE_URL_TEMPLATE = (
    "https://m.ad.search.naver.com/search.naver?where=m_expd&query={query}"
)

PC_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)
MOBILE_USER_AGENT = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) "
    "Version/17.0 Mobile/15E148 Safari/604.1"
)

REQUEST_TIMEOUT = 15
DEFAULT_DELAY = 1.0  # 요청 사이 기본 딜레이 (초)


@dataclass
class AdResult:
    """하나의 광고 결과"""
    rank: int
    title: str
    description: str = ""
    url: str = ""


@dataclass
class RankCheckResult:
    """특정 키워드에 대한 순위 확인 결과"""
    keyword: str
    target: str
    pc_rank: int = 0          # 0 = 미발견
    mobile_rank: int = 0      # 0 = 미발견
    pc_total: int = 0
    mobile_total: int = 0
    pc_matched_title: str = ""
    mobile_matched_title: str = ""
    pc_ads: list = field(default_factory=list)
    mobile_ads: list = field(default_factory=list)


def fetch_page(url: str, user_agent: str) -> str:
    """URL에서 HTML 페이지를 가져옴"""
    headers = {
        "User-Agent": user_agent,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Referer": "https://www.naver.com/",
    }
    resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_pc_ads(html: str) -> list[AdResult]:
    """PC 광고 검색 결과 파싱"""
    soup = BeautifulSoup(html, "lxml")
    ads: list[AdResult] = []

    # 전략 1: ol/ul > li 구조 (가장 일반적)
    for selector in [
        "ul.lst_type li",
        "ol.lst_type li",
        "ul#content li",
        "div.lst_type > div",
        "ul.srch_lst li",
        "div.wrap_ad_area > div",
    ]:
        items = soup.select(selector)
        if items:
            for i, item in enumerate(items, 1):
                title_el = (
                    item.select_one("a.lnk_tit")
                    or item.select_one(".tit_area a")
                    or item.select_one("a.tit")
                    or item.select_one(".area_text_title a")
                    or item.select_one("h3 a")
                    or item.select_one("a[class*='tit']")
                    or item.select_one(".title a")
                )
                if title_el:
                    title = title_el.get_text(strip=True)
                    desc_el = (
                        item.select_one(".ad_dsc")
                        or item.select_one(".dsc_area")
                        or item.select_one(".desc")
                        or item.select_one("p")
                    )
                    desc = desc_el.get_text(strip=True) if desc_el else ""
                    link = title_el.get("href", "")
                    ads.append(AdResult(rank=i, title=title, description=desc, url=link))
            if ads:
                return ads

    # 전략 2: data-cr-area 또는 data-area 기반
    items = soup.select("[data-cr-area] > div, [data-area] > div")
    if items:
        for i, item in enumerate(items, 1):
            title_el = item.select_one("a")
            if title_el and title_el.get_text(strip=True):
                ads.append(AdResult(
                    rank=i,
                    title=title_el.get_text(strip=True),
                ))
        if ads:
            return ads

    # 전략 3: 모든 주요 제목 링크 추출 (폴백)
    # 광고 제목은 보통 큰 링크 텍스트로 나타남
    seen_titles = set()
    rank = 0
    for a_tag in soup.find_all("a"):
        text = a_tag.get_text(strip=True)
        # 제목처럼 보이는 것만 필터 (2글자 이상, 네비게이션/메뉴 제외)
        if (
            len(text) >= 2
            and text not in seen_titles
            and not any(skip in text for skip in [
                "검색", "로그인", "네이버", "더보기", "이전", "다음",
                "통합검색", "내 업체", "등록하기", "광고문의",
                "이용약관", "개인정보", "고객센터",
            ])
        ):
            parent = a_tag.parent
            # 부모 요소가 제목 태그이거나, 제목 관련 클래스가 있는 경우
            is_title_like = (
                parent and parent.name in ("h2", "h3", "h4", "strong", "em")
            ) or any(
                cls and "tit" in cls.lower()
                for cls in (a_tag.get("class") or [])
            )
            if is_title_like:
                rank += 1
                seen_titles.add(text)
                ads.append(AdResult(rank=rank, title=text))

    # 전략 3이 결과가 없으면, 좀 더 넓은 범위로 시도
    if not ads:
        rank = 0
        for heading in soup.find_all(["h2", "h3", "h4"]):
            text = heading.get_text(strip=True)
            if len(text) >= 2 and text not in seen_titles:
                rank += 1
                seen_titles.add(text)
                ads.append(AdResult(rank=rank, title=text))

    return ads


def parse_mobile_ads(html: str) -> list[AdResult]:
    """모바일 광고 검색 결과 파싱"""
    soup = BeautifulSoup(html, "lxml")
    ads: list[AdResult] = []

    # 전략 1: 모바일 전용 선택자
    for selector in [
        "ul.lst_type li",
        "ol.lst_type li",
        "div.lst_type > div",
        "ul._list li",
        "div._list > div",
        "ul.srch_lst li",
        "div.wrap_ad_area > div",
    ]:
        items = soup.select(selector)
        if items:
            for i, item in enumerate(items, 1):
                title_el = (
                    item.select_one("a.lnk_tit")
                    or item.select_one(".tit_area a")
                    or item.select_one("a.tit")
                    or item.select_one(".area_text_title a")
                    or item.select_one("a[class*='tit']")
                    or item.select_one("h3 a")
                    or item.select_one(".title a")
                )
                if title_el:
                    title = title_el.get_text(strip=True)
                    desc_el = (
                        item.select_one(".ad_dsc")
                        or item.select_one(".dsc_area")
                        or item.select_one(".desc")
                        or item.select_one("p")
                    )
                    desc = desc_el.get_text(strip=True) if desc_el else ""
                    link = title_el.get("href", "")
                    ads.append(AdResult(rank=i, title=title, description=desc, url=link))
            if ads:
                return ads

    # 전략 2: data 속성 기반
    items = soup.select("[data-cr-area] > div, [data-area] > div")
    if items:
        for i, item in enumerate(items, 1):
            title_el = item.select_one("a")
            if title_el and title_el.get_text(strip=True):
                ads.append(AdResult(
                    rank=i,
                    title=title_el.get_text(strip=True),
                ))
        if ads:
            return ads

    # 전략 3: 제목 태그 기반 폴백
    seen_titles = set()
    rank = 0
    for heading in soup.find_all(["h2", "h3", "h4"]):
        a_tag = heading.find("a")
        text = (a_tag or heading).get_text(strip=True)
        if (
            len(text) >= 2
            and text not in seen_titles
            and not any(skip in text for skip in [
                "검색", "로그인", "네이버", "더보기", "파워링크",
                "통합검색", "이용약관", "개인정보", "고객센터",
            ])
        ):
            rank += 1
            seen_titles.add(text)
            ads.append(AdResult(rank=rank, title=text))

    return ads


def find_target_rank(ads: list[AdResult], target: str) -> tuple[int, str]:
    """
    광고 목록에서 타겟 업체 찾기.
    부분 문자열 매칭 사용 (예: '하늘마음' → '하늘마음한의원 덕천' 매칭)
    반환: (순위, 매칭된_제목) / 미발견 시 (0, "")
    """
    target_lower = target.lower().strip()
    for ad in ads:
        title_lower = ad.title.lower().strip()
        if target_lower in title_lower or title_lower in target_lower:
            return ad.rank, ad.title
    return 0, ""


def check_keyword(keyword: str, target: str, delay: float = DEFAULT_DELAY) -> RankCheckResult:
    """하나의 키워드에 대해 PC/모바일 순위 확인"""
    result = RankCheckResult(keyword=keyword, target=target)
    encoded = urllib.parse.quote(keyword)

    # PC 검색
    try:
        pc_url = PC_URL_TEMPLATE.format(query=encoded)
        pc_html = fetch_page(pc_url, PC_USER_AGENT)
        result.pc_ads = parse_pc_ads(pc_html)
        result.pc_total = len(result.pc_ads)
        result.pc_rank, result.pc_matched_title = find_target_rank(result.pc_ads, target)
    except requests.RequestException as e:
        print(f"  [PC 오류] {keyword}: {e}", file=sys.stderr)

    time.sleep(delay)

    # 모바일 검색
    try:
        mobile_url = MOBILE_URL_TEMPLATE.format(query=encoded)
        mobile_html = fetch_page(mobile_url, MOBILE_USER_AGENT)
        result.mobile_ads = parse_mobile_ads(mobile_html)
        result.mobile_total = len(result.mobile_ads)
        result.mobile_rank, result.mobile_matched_title = find_target_rank(result.mobile_ads, target)
    except requests.RequestException as e:
        print(f"  [MO 오류] {keyword}: {e}", file=sys.stderr)

    return result


def load_keywords(filepath: str) -> list[str]:
    """키워드 파일 읽기 (한 줄에 하나)"""
    path = Path(filepath)
    if not path.exists():
        print(f"오류: 키워드 파일을 찾을 수 없습니다: {filepath}", file=sys.stderr)
        sys.exit(1)
    keywords = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            keywords.append(line)
    if not keywords:
        print(f"오류: 키워드 파일이 비어있습니다: {filepath}", file=sys.stderr)
        sys.exit(1)
    return keywords


def print_results(results: list[RankCheckResult], target: str) -> None:
    """결과를 테이블 형식으로 출력"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print()
    print("=" * 80)
    print(f"  네이버 파워링크 광고 순위 체크 결과")
    print(f"  검색 대상: {target}")
    print(f"  조회 시각: {now}")
    print("=" * 80)
    print()

    # 헤더
    print(f"{'키워드':<20} {'PC순위':>8} {'PC전체':>8} {'MO순위':>8} {'MO전체':>8}")
    print("-" * 80)

    for r in results:
        pc_rank_str = f"{r.pc_rank}위" if r.pc_rank > 0 else "미노출"
        mo_rank_str = f"{r.mobile_rank}위" if r.mobile_rank > 0 else "미노출"
        pc_total_str = f"{r.pc_total}건"
        mo_total_str = f"{r.mobile_total}건"

        print(f"{r.keyword:<20} {pc_rank_str:>8} {pc_total_str:>8} {mo_rank_str:>8} {mo_total_str:>8}")

    print("-" * 80)
    print()

    # 상세 정보
    for r in results:
        if r.pc_ads or r.mobile_ads:
            print(f"--- [{r.keyword}] 상세 광고 목록 ---")
            if r.pc_ads:
                print(f"  [PC] 총 {len(r.pc_ads)}건")
                for ad in r.pc_ads:
                    marker = " ★" if r.pc_rank == ad.rank and r.pc_rank > 0 else ""
                    print(f"    {ad.rank}위: {ad.title}{marker}")
            if r.mobile_ads:
                print(f"  [MO] 총 {len(r.mobile_ads)}건")
                for ad in r.mobile_ads:
                    marker = " ★" if r.mobile_rank == ad.rank and r.mobile_rank > 0 else ""
                    print(f"    {ad.rank}위: {ad.title}{marker}")
            print()


def export_csv(results: list[RankCheckResult], target: str, filepath: str) -> None:
    """결과를 CSV 파일로 내보내기"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["조회시각", "검색대상", "키워드", "PC순위", "PC전체", "PC매칭제목",
                         "MO순위", "MO전체", "MO매칭제목"])
        for r in results:
            writer.writerow([
                now,
                target,
                r.keyword,
                r.pc_rank if r.pc_rank > 0 else "미노출",
                r.pc_total,
                r.pc_matched_title,
                r.mobile_rank if r.mobile_rank > 0 else "미노출",
                r.mobile_total,
                r.mobile_matched_title,
            ])
    print(f"CSV 저장 완료: {filepath}")


def save_debug_html(keyword: str, html: str, device: str) -> None:
    """디버깅용 HTML 저장"""
    debug_dir = Path("debug_html")
    debug_dir.mkdir(exist_ok=True)
    safe_keyword = keyword.replace(" ", "_").replace("/", "_")
    path = debug_dir / f"{safe_keyword}_{device}.html"
    path.write_text(html, encoding="utf-8")
    print(f"  디버그 HTML 저장: {path}")


def main():
    parser = argparse.ArgumentParser(
        description="네이버 파워링크 광고 순위 체크 프로그램",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  # 키워드 파일로 순위 확인
  python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음

  # 단일 키워드 확인
  python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음

  # CSV로 결과 저장
  python naver_ad_rank_checker.py -k keywords.txt -t 하늘마음 -o result.csv

  # 디버그 모드 (HTML 파일 저장)
  python naver_ad_rank_checker.py -q 부산가려움증 -t 하늘마음 --debug
        """,
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-k", "--keywords-file", help="키워드 목록 파일 경로 (한 줄에 하나)")
    group.add_argument("-q", "--query", help="단일 키워드 직접 입력")

    parser.add_argument("-t", "--target", required=True,
                        help="찾을 업체명 (부분 매칭, 예: 하늘마음)")
    parser.add_argument("-o", "--output", help="결과를 CSV 파일로 저장")
    parser.add_argument("--debug", action="store_true",
                        help="디버그 모드: 가져온 HTML을 debug_html/ 폴더에 저장")
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY,
                        help=f"요청 사이 딜레이 초 (기본: {DEFAULT_DELAY})")

    args = parser.parse_args()

    # 키워드 로드
    if args.keywords_file:
        keywords = load_keywords(args.keywords_file)
    else:
        keywords = [args.query]

    target = args.target
    delay = args.delay

    print(f"검색 대상: {target}")
    print(f"키워드 수: {len(keywords)}")
    print(f"요청 딜레이: {delay}초")
    print()

    results: list[RankCheckResult] = []

    for i, keyword in enumerate(keywords, 1):
        print(f"[{i}/{len(keywords)}] '{keyword}' 검색 중...")

        if args.debug:
            # 디버그 모드: HTML 저장
            encoded = urllib.parse.quote(keyword)
            try:
                pc_html = fetch_page(PC_URL_TEMPLATE.format(query=encoded), PC_USER_AGENT)
                save_debug_html(keyword, pc_html, "pc")
            except requests.RequestException as e:
                print(f"  [PC 오류] {e}", file=sys.stderr)

            try:
                mobile_html = fetch_page(MOBILE_URL_TEMPLATE.format(query=encoded), MOBILE_USER_AGENT)
                save_debug_html(keyword, mobile_html, "mobile")
            except requests.RequestException as e:
                print(f"  [MO 오류] {e}", file=sys.stderr)

        result = check_keyword(keyword, target, delay)
        results.append(result)

        # 순위 미리보기
        pc_str = f"PC {result.pc_rank}위" if result.pc_rank > 0 else "PC 미노출"
        mo_str = f"MO {result.mobile_rank}위" if result.mobile_rank > 0 else "MO 미노출"
        print(f"  → {pc_str} / {mo_str}")

        if i < len(keywords):
            time.sleep(delay)

    # 결과 출력
    print_results(results, target)

    # CSV 저장
    if args.output:
        export_csv(results, target, args.output)


if __name__ == "__main__":
    main()
