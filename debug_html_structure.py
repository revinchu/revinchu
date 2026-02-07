#!/usr/bin/env python3
"""
네이버 광고 페이지 HTML 구조 분석 도구
- 실제 페이지를 가져와서 HTML 구조를 분석하고 출력
- 파서가 제대로 동작하지 않을 때 HTML 구조를 파악하는 데 사용
"""

import sys
import urllib.parse

import requests
from bs4 import BeautifulSoup

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


def analyze_html(html: str, label: str) -> None:
    """HTML 구조 분석"""
    soup = BeautifulSoup(html, "lxml")

    print(f"\n{'='*60}")
    print(f"  {label} HTML 구조 분석")
    print(f"{'='*60}\n")

    # 1. 전체 구조 요약
    print("[1] body 직계 자식 태그:")
    body = soup.find("body")
    if body:
        for child in body.children:
            if hasattr(child, "name") and child.name:
                classes = child.get("class", [])
                id_attr = child.get("id", "")
                print(f"  <{child.name}> class={classes} id='{id_attr}'")

    # 2. 리스트 구조 찾기
    print("\n[2] 리스트 (ul/ol) 구조:")
    for tag in soup.find_all(["ul", "ol"]):
        classes = tag.get("class", [])
        li_count = len(tag.find_all("li", recursive=False))
        if li_count > 0:
            print(f"  <{tag.name}> class={classes} → {li_count}개 li")
            for i, li in enumerate(tag.find_all("li", recursive=False)[:3], 1):
                li_classes = li.get("class", [])
                first_text = li.get_text(strip=True)[:80]
                print(f"    li[{i}] class={li_classes}: {first_text}...")

    # 3. 제목 태그 찾기
    print("\n[3] 제목 태그 (h1~h4):")
    for tag in soup.find_all(["h1", "h2", "h3", "h4"]):
        classes = tag.get("class", [])
        text = tag.get_text(strip=True)[:80]
        a_tag = tag.find("a")
        print(f"  <{tag.name}> class={classes}: {text}")
        if a_tag:
            print(f"    └ <a> href={a_tag.get('href', '')[:60]}")

    # 4. 'tit' 포함 클래스 찾기
    print("\n[4] 'tit' 클래스 포함 요소:")
    for tag in soup.find_all(True, class_=True):
        classes = tag.get("class", [])
        if any("tit" in c.lower() for c in classes):
            text = tag.get_text(strip=True)[:80]
            print(f"  <{tag.name}> class={classes}: {text}")

    # 5. data 속성 가진 주요 요소
    print("\n[5] data-* 속성 가진 주요 요소:")
    for tag in soup.find_all(True):
        data_attrs = {k: v for k, v in tag.attrs.items() if k.startswith("data-")}
        if data_attrs and tag.name not in ("script", "style", "link"):
            text = tag.get_text(strip=True)[:50]
            if text:
                print(f"  <{tag.name}> {data_attrs}: {text}...")

    # 6. 전체 텍스트에서 비즈니스명처럼 보이는 것 추출
    print("\n[6] 주요 링크 텍스트 (2자 이상):")
    seen = set()
    for a in soup.find_all("a"):
        text = a.get_text(strip=True)
        parent_tag = a.parent.name if a.parent else ""
        parent_class = a.parent.get("class", []) if a.parent else []
        a_class = a.get("class", [])
        if len(text) >= 2 and text not in seen and len(text) < 50:
            seen.add(text)
            print(f"  [{parent_tag} class={parent_class}] <a class={a_class}> {text}")

    # 7. Raw HTML 처음 부분 (참고용)
    print(f"\n[7] HTML 앞부분 (2000자):")
    print(html[:2000])
    print("\n... (생략)")


def main():
    keyword = sys.argv[1] if len(sys.argv) > 1 else "부산가려움증"
    encoded = urllib.parse.quote(keyword)

    print(f"키워드: {keyword}")

    headers_pc = {"User-Agent": PC_USER_AGENT, "Accept-Language": "ko-KR,ko;q=0.9"}
    headers_mo = {"User-Agent": MOBILE_USER_AGENT, "Accept-Language": "ko-KR,ko;q=0.9"}

    # PC
    try:
        pc_url = f"https://ad.search.naver.com/search.naver?where=ad&query={encoded}"
        r = requests.get(pc_url, headers=headers_pc, timeout=15)
        r.raise_for_status()
        r.encoding = r.apparent_encoding or "utf-8"
        analyze_html(r.text, f"PC ({keyword})")

        # HTML 파일 저장
        with open(f"debug_{keyword}_pc.html", "w", encoding="utf-8") as f:
            f.write(r.text)
        print(f"\n→ PC HTML 저장: debug_{keyword}_pc.html")
    except Exception as e:
        print(f"PC 요청 실패: {e}")

    # 모바일
    try:
        mo_url = f"https://m.ad.search.naver.com/search.naver?where=m_expd&query={encoded}"
        r = requests.get(mo_url, headers=headers_mo, timeout=15)
        r.raise_for_status()
        r.encoding = r.apparent_encoding or "utf-8"
        analyze_html(r.text, f"Mobile ({keyword})")

        with open(f"debug_{keyword}_mobile.html", "w", encoding="utf-8") as f:
            f.write(r.text)
        print(f"\n→ Mobile HTML 저장: debug_{keyword}_mobile.html")
    except Exception as e:
        print(f"Mobile 요청 실패: {e}")


if __name__ == "__main__":
    main()
