"""data.go.kr 일반 인증키(Decoding)를 클립보드에서 읽어 이 PC에만 저장한다 (화면·저장소에 남지 않음).

    1) data.go.kr 승인된 API 상세 화면에서 '일반 인증키(Decoding)' 값만 복사(Ctrl+C)
    2) python cloudflare/set_data_key.py
"""
import pathlib
import re
import subprocess
import urllib.error
import urllib.parse
import urllib.request

KEY_FILE = pathlib.Path.home() / ".wizx" / "data_go_kr_key"


def clipboard():
    out = subprocess.run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"], capture_output=True)
    return out.stdout.decode("utf-8", "replace").strip()


def main():
    key = clipboard()
    if not re.fullmatch(r"[A-Za-z0-9+/=]{32,128}", key):
        print(f"클립보드 값이 인증키 형식이 아닙니다 (길이 {len(key)}). '일반 인증키(Decoding)' 값만 복사한 뒤 다시 실행하세요.")
        return 1
    url = ("https://apis.data.go.kr/B551182/hospInfoServicev2/getHospBasisList?numOfRows=1&_type=json&serviceKey="
           + urllib.parse.quote(key, safe=""))
    try:
        urllib.request.urlopen(url, timeout=20).read()
        print("확인: 심평원 API 호출 성공")
    except urllib.error.HTTPError as e:
        print(f"경고: 심평원 API 가 거절했습니다 (HTTP {e.code}). 승인 직후라면 1시간쯤 뒤 다시 시도하세요. 키는 저장합니다.")
    KEY_FILE.parent.mkdir(exist_ok=True)
    KEY_FILE.write_text(key, encoding="ascii")
    subprocess.run(["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value ' '"])
    print(f"저장했습니다: {KEY_FILE} (보안을 위해 클립보드를 비웠습니다)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
