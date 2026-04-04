"""
네이버 검색광고 API 클라이언트
Naver Search Ad API client for keyword search volume data.

필요한 환경변수 (Required environment variables):
  NAVER_API_KEY      - 검색광고 API 액세스 키
  NAVER_SECRET_KEY   - 검색광고 API 시크릿 키
  NAVER_CUSTOMER_ID  - 검색광고 고객 ID

API 문서: https://naver.github.io/searchad-apidoc/
"""

import os
import time
import hmac
import hashlib
import base64
import requests
from typing import Optional


BASE_URL = "https://api.naver.com"


class NaverAdAPIError(Exception):
    pass


class NaverAdClient:
    def __init__(
        self,
        api_key: Optional[str] = None,
        secret_key: Optional[str] = None,
        customer_id: Optional[str] = None,
    ):
        self.api_key = api_key or os.environ.get("NAVER_API_KEY", "")
        self.secret_key = secret_key or os.environ.get("NAVER_SECRET_KEY", "")
        self.customer_id = customer_id or os.environ.get("NAVER_CUSTOMER_ID", "")

        if not all([self.api_key, self.secret_key, self.customer_id]):
            raise NaverAdAPIError(
                "API 인증 정보가 없습니다. "
                "NAVER_API_KEY, NAVER_SECRET_KEY, NAVER_CUSTOMER_ID 환경변수를 설정하세요."
            )

    def _sign(self, timestamp: str, method: str, path: str) -> str:
        """HMAC-SHA256 서명 생성"""
        message = f"{timestamp}.{method}.{path}"
        hashed = hmac.new(
            self.secret_key.encode("utf-8"),
            message.encode("utf-8"),
            hashlib.sha256,
        )
        return base64.b64encode(hashed.digest()).decode("utf-8")

    def _headers(self, method: str, path: str) -> dict:
        timestamp = str(int(time.time() * 1000))
        return {
            "Content-Type": "application/json; charset=UTF-8",
            "X-Timestamp": timestamp,
            "X-API-KEY": self.api_key,
            "X-Customer": self.customer_id,
            "X-Signature": self._sign(timestamp, method, path),
        }

    def get_keyword_stats(self, keywords: list[str]) -> list[dict]:
        """
        키워드 검색량 조회 (최대 5개씩 배치)

        Returns list of dicts with keys:
          relKeyword, monthlyPcQcCnt, monthlyMobileQcCnt,
          monthlyAvePcClkCnt, monthlyAveMobileClkCnt, compIdx
        """
        path = "/keywordstool"
        all_results: list[dict] = []

        # API는 한 번에 최대 5개 키워드 허용
        for i in range(0, len(keywords), 5):
            batch = keywords[i : i + 5]
            params = {
                "hintKeywords": ",".join(batch),
                "showDetail": "1",
            }
            headers = self._headers("GET", path)
            resp = requests.get(
                BASE_URL + path,
                headers=headers,
                params=params,
                timeout=10,
            )
            if resp.status_code != 200:
                raise NaverAdAPIError(
                    f"API 오류 {resp.status_code}: {resp.text}"
                )
            data = resp.json()
            all_results.extend(data.get("keywordList", []))

            if i + 5 < len(keywords):
                time.sleep(0.3)  # 배치 간 짧은 딜레이

        # 요청한 키워드와 매칭하여 반환 (API는 연관 키워드도 같이 반환함)
        requested = {kw.strip() for kw in keywords}
        matched = [r for r in all_results if r.get("relKeyword") in requested]

        # 매칭 안 된 키워드는 0으로 채워서 포함
        found = {r["relKeyword"] for r in matched}
        for kw in keywords:
            if kw not in found:
                matched.append(
                    {
                        "relKeyword": kw,
                        "monthlyPcQcCnt": 0,
                        "monthlyMobileQcCnt": 0,
                        "monthlyAvePcClkCnt": 0,
                        "monthlyAveMobileClkCnt": 0,
                        "compIdx": "-",
                    }
                )

        # 요청 순서 유지
        order = {kw: i for i, kw in enumerate(keywords)}
        matched.sort(key=lambda r: order.get(r["relKeyword"], 9999))
        return matched
