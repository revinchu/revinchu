"""운영 서버(WSGI) 진입점. 예: gunicorn medical_platform.wsgi:application"""
from .app import create_app

application = create_app()
