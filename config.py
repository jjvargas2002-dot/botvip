import os
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

class Config:
    db_url = os.getenv("DATABASE_URL", "postgresql://postgres:postgres@localhost:5432/botbitel")
    if db_url and db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
        
    SQLALCHEMY_DATABASE_URI = db_url
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": 280,
        "pool_timeout": 30
    }
    SECRET_KEY = os.getenv("SECRET_KEY", "botvip_secret_key_12345")

    WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
    PHONE_NUMBER_ID = os.getenv("PHONE_NUMBER_ID", "1098169976722957")

    STATUS_CAJAS_WEBHOOK_URL = os.getenv("STATUS_CAJAS_WEBHOOK_URL")
    STATUS_CAJAS_WEBHOOK_SECRET = os.getenv("STATUS_CAJAS_WEBHOOK_SECRET")