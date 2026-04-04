"""
네이버 키워드 검색량 대시보드
Naver keyword search volume dashboard (Flask web app)

실행 방법:
  # 1. 환경변수 설정 (.env 파일 또는 직접 export)
  export NAVER_API_KEY=your_api_key
  export NAVER_SECRET_KEY=your_secret_key
  export NAVER_CUSTOMER_ID=your_customer_id

  # 2. 서버 실행
  python dashboard.py

  # 3. 브라우저에서 열기
  http://localhost:5000
"""

import os
import json
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify
from dotenv import load_dotenv

from naver_api import NaverAdClient, NaverAdAPIError

load_dotenv()

app = Flask(__name__)

KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"


# ── 키워드 파일 헬퍼 ──────────────────────────────────────────────────

def load_keywords() -> list[str]:
    if not KEYWORDS_FILE.exists():
        return []
    lines = KEYWORDS_FILE.read_text(encoding="utf-8").splitlines()
    return [ln.strip() for ln in lines if ln.strip() and not ln.startswith("#")]


def save_keywords(keywords: list[str]) -> None:
    KEYWORDS_FILE.write_text(
        "\n".join(keywords) + "\n", encoding="utf-8"
    )


# ── 라우트 ────────────────────────────────────────────────────────────

@app.route("/")
def index():
    keywords = load_keywords()
    credentials_ok = all([
        os.environ.get("NAVER_API_KEY"),
        os.environ.get("NAVER_SECRET_KEY"),
        os.environ.get("NAVER_CUSTOMER_ID"),
    ])
    return render_template(
        "index.html",
        keywords=keywords,
        credentials_ok=credentials_ok,
    )


@app.route("/api/search-volume")
def api_search_volume():
    """키워드 검색량 일괄 조회"""
    keywords = load_keywords()
    if not keywords:
        return jsonify({"error": "키워드가 없습니다.", "data": []})

    try:
        client = NaverAdClient()
        results = client.get_keyword_stats(keywords)
    except NaverAdAPIError as e:
        return jsonify({"error": str(e), "data": []}), 502

    # 응답 정규화: 숫자 "< 10" 처리 (API가 문자열로 반환할 수 있음)
    def parse_count(val):
        if isinstance(val, int):
            return val
        if isinstance(val, str) and val.strip() == "< 10":
            return 5  # 10 미만은 5로 표시
        try:
            return int(val)
        except (ValueError, TypeError):
            return 0

    normalized = []
    for r in results:
        pc = parse_count(r.get("monthlyPcQcCnt", 0))
        mobile = parse_count(r.get("monthlyMobileQcCnt", 0))
        normalized.append({
            "keyword": r["relKeyword"],
            "pc": pc,
            "mobile": mobile,
            "total": pc + mobile,
            "competition": r.get("compIdx", "-"),
        })

    return jsonify({
        "data": normalized,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


@app.route("/api/keywords", methods=["GET"])
def api_get_keywords():
    return jsonify({"keywords": load_keywords()})


@app.route("/api/keywords", methods=["POST"])
def api_add_keyword():
    body = request.get_json(silent=True) or {}
    kw = (body.get("keyword") or "").strip()
    if not kw:
        return jsonify({"error": "키워드를 입력하세요."}), 400

    keywords = load_keywords()
    if kw in keywords:
        return jsonify({"error": "이미 등록된 키워드입니다."}), 409

    keywords.append(kw)
    save_keywords(keywords)
    return jsonify({"keywords": keywords}), 201


@app.route("/api/keywords/<path:keyword>", methods=["DELETE"])
def api_delete_keyword(keyword: str):
    keywords = load_keywords()
    if keyword not in keywords:
        return jsonify({"error": "키워드를 찾을 수 없습니다."}), 404

    keywords = [k for k in keywords if k != keyword]
    save_keywords(keywords)
    return jsonify({"keywords": keywords})


@app.route("/api/keywords/bulk", methods=["POST"])
def api_bulk_keywords():
    """여러 키워드를 한 번에 등록 (줄바꿈 구분)"""
    body = request.get_json(silent=True) or {}
    raw = body.get("text", "")
    new_keywords = [ln.strip() for ln in raw.splitlines() if ln.strip() and not ln.startswith("#")]

    if not new_keywords:
        return jsonify({"error": "키워드를 입력하세요."}), 400

    existing = load_keywords()
    added = []
    for kw in new_keywords:
        if kw not in existing:
            existing.append(kw)
            added.append(kw)

    save_keywords(existing)
    return jsonify({"keywords": existing, "added": added}), 201


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    print(f"대시보드 시작: http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
