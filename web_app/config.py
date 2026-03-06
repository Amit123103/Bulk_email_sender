"""Flask + SocketIO Web App Configuration"""
import os
from pathlib import Path

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'bulk-email-pro-secret-key-change-in-production')
    BASE_DIR = Path(__file__).parent
    DATA_DIR = BASE_DIR / 'data'
    USERS_DIR = DATA_DIR / 'users'
    UPLOAD_FOLDER = DATA_DIR / 'uploads'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

    @staticmethod
    def init_dirs():
        for d in [Config.DATA_DIR, Config.USERS_DIR, Config.UPLOAD_FOLDER]:
            d.mkdir(parents=True, exist_ok=True)
