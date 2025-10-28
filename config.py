# app/config.py
import os
from dotenv import load_dotenv

load_dotenv()

APP_TITLE = "SPS Copilot – PPSPS + PDP"
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_BASE_URL = os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
OPENROUTER_DEFAULT_MODEL = "openai/gpt-5"   # ou "openai/gpt-5-chat"
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "")
SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret")

# Stripe
STRIPE_SECRET_KEY = os.getenv("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "")