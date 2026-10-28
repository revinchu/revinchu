"""WIZ X 해외환자 유치 플랫폼 (가칭 MEDIPRISM) - Flask 애플리케이션 팩토리.

실행:
    python -m medical_platform.app            # http://127.0.0.1:5000
환경 변수:
    SECRET_KEY        세션 서명 키 (운영 시 반드시 설정)
    DATABASE_PATH     SQLite 파일 경로 (기본: medical_platform/instance/platform.db)
    BRAND_NAME        사이트 브랜드명 (기본: MEDIPRISM)
    ADMIN_INITIAL_PASSWORD  첫 실행 시 관리자 비밀번호 (설정하면 데모 직원 계정을 만들지 않음 — 운영 배포용)
    DB_BACKEND        sqlite(기본) | d1 (Cloudflare Workers, cloudflare/worker.py 에서 설정)
"""
import os
import secrets

from flask import Flask, abort, request, session
from jinja2 import DictLoader
from markupsafe import Markup

from . import db as dbmod
from . import i18n

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _persistent_secret(instance_path):
    """SECRET_KEY 미설정 시 instance/secret_key 에 한 번 생성해 재사용한다 (재시작·다중 워커에도 세션 유지)."""
    path = os.path.join(instance_path, "secret_key")
    os.makedirs(instance_path, exist_ok=True)
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
        os.chmod(path, 0o600)
    with open(path) as f:
        return f.read().strip()


def create_app(test_config=None):
    app = Flask(__name__, instance_path=os.path.join(BASE_DIR, "instance"))
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SECRET_KEY") or _persistent_secret(app.instance_path),
        DATABASE=os.environ.get("DATABASE_PATH") or os.path.join(app.instance_path, "platform.db"),
        BRAND_NAME=os.environ.get("BRAND_NAME", "MEDIPRISM"),
        COMPANY_NAME="WIZ X",
        PRIVACY_POLICY_VERSION="2026-10-v1",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        DB_BACKEND=os.environ.get("DB_BACKEND", "sqlite"),
        ADMIN_INITIAL_PASSWORD=os.environ.get("ADMIN_INITIAL_PASSWORD") or None,
    )
    if test_config:
        app.config.update(test_config)
    if app.config["DB_BACKEND"] == "sqlite":
        os.makedirs(os.path.dirname(app.config["DATABASE"]) or ".", exist_ok=True)

    # Cloudflare Workers 빌드는 템플릿·아이콘을 _embedded.py 로 묶어 파일시스템 없이 동작한다 (cloudflare/build.py)
    try:
        from . import _embedded
    except ImportError:
        _embedded = None
    if _embedded is not None:
        app.jinja_loader = DictLoader(_embedded.TEMPLATES)

    app.teardown_appcontext(dbmod.close_db)

    @app.before_request
    def csrf_protect():
        if request.method == "POST":
            token = session.get("_csrf")
            if not token or token != request.form.get("_csrf"):
                abort(400, description="CSRF token missing or invalid")

    def csrf_token():
        if "_csrf" not in session:
            session["_csrf"] = secrets.token_urlsafe(32)
        return session["_csrf"]

    app.jinja_env.globals.update(
        csrf_token=csrf_token,
        LANGUAGES=i18n.LANGUAGES,
        brand=lambda: app.config["BRAND_NAME"],
        company=lambda: app.config["COMPANY_NAME"],
    )

    icon_dir = os.path.join(BASE_DIR, "static", "vendor", "lucide", "icons")
    icon_cache = {}

    def icon(name, size=18):
        """Lucide(ISC) SVG 아이콘을 인라인으로 삽입한다."""
        if name not in icon_cache:
            if _embedded is not None:
                svg = _embedded.ICONS[name]
            else:
                with open(os.path.join(icon_dir, f"{name}.svg")) as f:
                    svg = f.read()
            svg = svg.split("-->", 1)[-1].strip()
            icon_cache[name] = svg.replace('width="24"', 'width="{s}"').replace('height="24"', 'height="{s}"').replace(
                "<svg", '<svg aria-hidden="true" focusable="false"', 1).replace('stroke-width="2"', 'stroke-width="1.75"')
        return Markup(icon_cache[name].replace("{s}", str(int(size))))

    app.jinja_env.globals["icon"] = icon

    @app.template_filter("krw")
    def krw(value):
        return "-" if value is None else f"₩{value:,}"

    from .public import bp as public_bp
    from .admin import bp as admin_bp
    app.register_blueprint(public_bp)
    app.register_blueprint(admin_bp)

    with app.app_context():
        dbmod.init_db()

    return app


if __name__ == "__main__":
    create_app().run(debug=os.environ.get("FLASK_DEBUG") == "1")
