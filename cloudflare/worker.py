"""Cloudflare Python Workers 진입점.

Flask 앱(medical_platform)을 Cloudflare의 WSGI 어댑터로 실행하고, DB는 D1 바인딩(DB)을 쓴다.
설정값은 wrangler vars / secrets 에서 읽어 os.environ 으로 옮긴 뒤 앱을 한 번만 만든다.
cron 트리거는 30분 넘게 미응대인 문의를 이메일로 알린다.

무료 플랜은 요청당 CPU 가 빠듯해서, 누구에게나 같은 공개 페이지(GET)는 인스턴스 메모리에 10분간 보관해
다시 요청되면 Flask 를 거치지 않고 바로 돌려준다. 관리자·상담 신청·광고 유입(utm) 요청은 보관하지 않는다.
"""
import os
import re
import time

import js
from pyodide.ffi import to_js
from workers import WorkerEntrypoint, wsgi

# 최상단에서 import 해야 배포 시 메모리 스냅샷에 포함되어 요청 CPU 시간에서 빠진다 (설정값은 호출 시점에 읽음)
from medical_platform import notify
from medical_platform.app import create_app
from medical_platform.db import init_db

CONFIG_KEYS = ("SECRET_KEY", "ADMIN_INITIAL_PASSWORD", "BRAND_NAME", "SITE_URL",
               "NOTIFY_EMAIL", "MAIL_BACKEND", "MAIL_FROM", "RESEND_API_KEY")

CACHEABLE = re.compile(r"^/(en|ko|zh|ja|vi)/(hospitals(/[\w-]+)?|magazine(/[\w-]+)?|filming-locations|concierge|"
                       r"how-it-works|privacy)?$|^/(robots\.txt|sitemap\.xml|llms\.txt|llms-full\.txt)$")
CACHE_TTL = 600
CACHE_MAX = 400
_cache = {}  # URL 경로+쿼리 → (만료 시각, content-type, 본문)
_ready = False

# 앱 생성과 템플릿 컴파일을 import 시점(= 배포 시 메모리 스냅샷)에 끝내 둔다.
# 새 인스턴스의 첫 요청은 비밀값·DB 바인딩만 채우고 스키마 버전만 확인한다 (무료 플랜 CPU 한도 대응).
os.environ.setdefault("SECRET_KEY", "set-on-first-request")
os.environ["DB_BACKEND"] = "d1"
_app = create_app({"SKIP_INIT_DB": True, "SESSION_COOKIE_SECURE": True})
for _name in _app.jinja_env.list_templates():
    if _name.startswith("public/"):
        _app.jinja_env.get_template(_name)


def get_app(env):
    global _ready
    if not _ready:
        for key in CONFIG_KEYS:
            try:
                value = getattr(env, key)
            except AttributeError:
                continue
            if isinstance(value, str) and value:
                os.environ[key] = value
        if os.environ.get("SECRET_KEY") in (None, "", "set-on-first-request"):
            raise RuntimeError("SECRET_KEY 시크릿이 없습니다: npx wrangler secret put SECRET_KEY")
        _app.config.update(SECRET_KEY=os.environ["SECRET_KEY"], D1_BINDING=env.DB,
                           BRAND_NAME=os.environ.get("BRAND_NAME", _app.config["BRAND_NAME"]),
                           ADMIN_INITIAL_PASSWORD=os.environ.get("ADMIN_INITIAL_PASSWORD") or None)
        with _app.app_context():
            init_db()
        _ready = True
    return _app


def cache_key(request):
    if request.method != "GET":
        return None
    url = js.URL.new(request.url)
    if not CACHEABLE.match(url.pathname) or "utm_" in url.search:
        return None
    return url.pathname + url.search


def cached_response(content_type, body, state):
    init = to_js({"status": 200, "headers": {"content-type": content_type, "x-page-cache": state}},
                 dict_converter=js.Object.fromEntries)
    return js.Response.new(body, init)


class Default(WorkerEntrypoint):
    async def fetch(self, request):
        key = cache_key(request)
        if key:
            hit = _cache.get(key)
            if hit and hit[0] > time.time():
                return cached_response(hit[1], hit[2], "HIT")
        resp = await wsgi.fetch(get_app(self.env), request, self.env)
        if not key or resp.status != 200:
            return resp
        body = await resp.text()
        content_type = resp.headers.get("content-type") or "text/html; charset=utf-8"
        if len(_cache) >= CACHE_MAX:
            _cache.pop(next(iter(_cache)))
        _cache[key] = (time.time() + CACHE_TTL, content_type, body)
        # 처음 방문자에게는 세션 쿠키 등 원래 헤더를 그대로 돌려준다 (보관본에는 쿠키를 넣지 않는다)
        return js.Response.new(body, resp)

    async def scheduled(self, controller, env, ctx):
        app = get_app(self.env)
        with app.app_context():
            notify.send_sla_alerts()
