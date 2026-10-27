"""WIZ X 해외환자 유치 플랫폼 (가칭 MEDIPRISM) - Flask 애플리케이션 팩토리.

실행:
    python -m medical_platform.app            # http://127.0.0.1:5000
환경 변수:
    SECRET_KEY        세션 서명 키 (운영 시 반드시 설정)
    DATABASE_PATH     SQLite 파일 경로 (기본: medical_platform/instance/platform.db)
    BRAND_NAME        사이트 브랜드명 (기본: MEDIPRISM)
"""
import os
import secrets

from flask import Flask, abort, request, session

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
    )
    if test_config:
        app.config.update(test_config)
    os.makedirs(os.path.dirname(app.config["DATABASE"]) or ".", exist_ok=True)

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
