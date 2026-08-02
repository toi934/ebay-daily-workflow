"""CPaSS ログイン情報 (GitHub Actions版: 環境変数から読み込む)"""
import os

CPASS_EMAIL = os.environ.get("CPASS_EMAIL", "")
CPASS_PASSWORD = os.environ.get("CPASS_PASSWORD", "")
# 2026/08/02: CPaSSは発送元地域ごとに専用サイトを使う仕様に変更されており、地域指定なしの
# URLでログインすると日本発送用のShipping Labelを購入できない状態になる(戸井さん実機確認・
# 指示)。必ず日本用URL(https://cpass.ebay.com/JP)を経由してログインすること。
CPASS_LOGIN_URL = "https://cpass.ebay.com/JP"
CPASS_ORDER_URL = "https://cpass.ebay.com/jp/order/paid"

def is_configured():
    return bool(CPASS_EMAIL.strip()) and bool(CPASS_PASSWORD.strip())
