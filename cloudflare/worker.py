"""Cloudflare Python Workers 진입점.

Flask 앱(medical_platform)을 Cloudflare의 WSGI 어댑터로 실행하고, DB는 D1 바인딩(DB)을 쓴다.
설정값은 wrangler vars / secrets 에서 읽어 os.environ 으로 옮긴 뒤 앱을 한 번만 만든다.
cron 트리거는 30분 넘게 미응대인 문의를 이메일로 알린다.
"""
import os

from workers import WorkerEntrypoint, wsgi

# 최상단에서 import 해야 배포 시 메모리 스냅샷에 포함되어 요청 CPU 시간에서 빠진다 (설정값은 호출 시점에 읽음)
from medical_platform import notify
from medical_platform.app import create_app

CONFIG_KEYS = ("SECRET_KEY", "ADMIN_INITIAL_PASSWORD", "BRAND_NAME", "SITE_URL",
               "NOTIFY_EMAIL", "MAIL_BACKEND", "MAIL_FROM", "RESEND_API_KEY")

_app = None


def get_app(env):
    global _app
    if _app is None:
        for key in CONFIG_KEYS:
            try:
                value = getattr(env, key)
            except AttributeError:
                continue
            if isinstance(value, str) and value:
                os.environ[key] = value
        if not os.environ.get("SECRET_KEY"):
            raise RuntimeError("SECRET_KEY 시크릿이 없습니다: npx wrangler secret put SECRET_KEY")
        os.environ["DB_BACKEND"] = "d1"
        _app = create_app({"D1_BINDING": env.DB, "SESSION_COOKIE_SECURE": True})
    return _app


class Default(WorkerEntrypoint):
    async def fetch(self, request):
        return await wsgi.fetch(get_app(self.env), request, self.env)

    async def scheduled(self, controller, env, ctx):
        app = get_app(self.env)
        with app.app_context():
            notify.send_sla_alerts()
