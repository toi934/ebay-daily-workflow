"""CPaSS ログイン情報 (GitHub Actions版: 環境変数から読み込む)"""
import os

CPASS_EMAIL = os.environ.get("CPASS_EMAIL", "")
CPASS_PASSWORD = os.environ.get("CPASS_PASSWORD", "")
CPASS_LOGIN_URL = "https://cpass.ebay.com/login"
CPASS_ORDER_URL = "https://cpass.ebay.com/order/paid"

def is_configured():
    return bool(CPASS_EMAIL.strip()) and bool(CPASS_PASSWORD.strip())
