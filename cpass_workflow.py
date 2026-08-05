"""CPaSS 編集→保存→DHL価格取得 ワークフロー

メイン関数: process_all_orders_for_dhl()
  1. ログイン
  2. 発送手続き待ち の全件を 発送手続き へ移動
  3. 各注文について:
     - 編集ダイアログを開く
     - 重量・寸法・HSコード を自動入力
     - 保存
     - 配送を割り当て → DHL の上限価格を取得
  4. 結果を {order_no: dhl_jpy_price} の dict で返す

使い方:
    from cpass_workflow import process_all_orders_for_dhl

    results = process_all_orders_for_dhl(target_order_nos=["20-14650-92130"])
    # → {"20-14650-92130": {"package_no": "2877", "dhl_price": 4849, ...}, ...}

注意:
- CPaSS の DOM 構造は変更される可能性があるので、セレクタは適宜調整必要
- 1件あたり 5〜15秒 程度かかる
"""

import sys
import os
import re
import time
import json
import requests
import xml.etree.ElementTree as ET

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import cpass_config
import hs_code_lookup
import dimension_weight_lookup

# ============================================================
# ★2026/08/04確定: DHL見積もり用の重量は、タイトルからのキーワード推定
#   (dimension_weight_lookup)ではなく、まずeBay実データを最優先で使う。
#   戸井さん指摘「送料が全部0.5kgで計算されている、ちゃんと調べろ」に対応。
#   ローカル版(ﾀｽｸ2_売上管理表)と同一ロジック。詳細コメントはローカル版参照。
#
#   ★GA版特有の注意: このリポジトリ(toi934/ebay-daily-workflow)には従来、
#   eBay Trading API用の認証情報(APP_ID/DEV_ID/CERT_ID/TSUJOU_TOKEN/SENMON_TOKEN)が
#   Secretsとして登録されていない（Google Drive/CPaSS/Gmail用のSecretsのみ）。
#   この重量取得を有効にするには、タスク1(ebay-restock)・タスク5(ebay-auto-reply)と
#   同じ値で以下5つのSecretsをこのリポジトリにも追加する必要がある:
#     APP_ID, DEV_ID, CERT_ID, TSUJOU_TOKEN, SENMON_TOKEN
#   未設定の場合はGetItemが失敗し、従来通りタイトルキーワード推定にフォールバックする
#   （エラーにはならず処理は継続するが、重量精度の改善効果が得られない）。
# ============================================================

TRADING_API_URL = "https://api.ebay.com/ws/api.dll"
TRADING_NS = {"e": "urn:ebay:apis:eBLBaseComponents"}
_SHIP_PROFILE_WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:kg|㎏)", re.IGNORECASE)
# ★2026/08/05確定: 実際のShipping Policy名は半角"kg"ではなく全角キログラム記号
#   "㎏"(U+338F)1文字が使われているケースがほとんどだった（診断スクリプトdiag_weight_check.pyで
#   直近48時間の実注文15件全件を検証し確認）。旧正規表現は半角"kg"のみ一致する作りだったため
#   15件全件で一致せず、8/4の修正は一度も実際には効いておらず旧来のタイトルキーワード推定に
#   毎回フォールバックしていた（戸井さん指摘「昨日と同じ」の直接の原因）。

_GA_ACCOUNTS = {
    "tsujou": {
        "TOKEN": os.environ.get("TSUJOU_TOKEN", ""),
        "APP_ID": os.environ.get("APP_ID", ""),
        "DEV_ID": os.environ.get("DEV_ID", ""),
        "CERT_ID": os.environ.get("CERT_ID", ""),
    },
    "senmon": {
        "TOKEN": os.environ.get("SENMON_TOKEN", ""),
        "APP_ID": os.environ.get("APP_ID", ""),
        "DEV_ID": os.environ.get("DEV_ID", ""),
        "CERT_ID": os.environ.get("CERT_ID", ""),
    },
}


def resolve_account_by_item_id(item_id):
    """US Item IDの先頭数字でアカウントを判定する
    （タスク1除外リスト判定・タスク6と同じ確定ルールを流用: 2026/07/27）

    3から始まる → 専門(senmon) / それ以外（1から始まる等）→ 通常(tsujou、デフォルト)
    """
    s = str(item_id or "").strip()
    if s.startswith("3"):
        return "senmon"
    return "tsujou"


def get_ebay_shipping_weight_kg(item_id, account_name=None):
    """Trading API GetItemで、このitem_idに割り当てられているShipping Policy名
    から実重量(kg)を取得する（ローカル版と同一ロジック。詳細はそちらのコメント参照）。

    Returns:
        (weight_kg: float|None, source: str, error: str|None)
    """
    if not item_id:
        return None, "none", "item_idが空のため取得不可"

    if account_name is None:
        account_name = resolve_account_by_item_id(item_id)

    account = _GA_ACCOUNTS.get(account_name, {})
    if not account.get("TOKEN") or not account.get("APP_ID"):
        return None, "none", "eBay API認証情報が未設定（Secrets未登録の可能性）"

    xml_body = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">\n'
        '  <RequesterCredentials>\n'
        '    <eBayAuthToken>' + account["TOKEN"] + '</eBayAuthToken>\n'
        '  </RequesterCredentials>\n'
        '  <ItemID>' + str(item_id) + '</ItemID>\n'
        '  <DetailLevel>ReturnAll</DetailLevel>\n'
        '</GetItemRequest>\n'
    )
    headers = {
        "X-EBAY-API-CALL-NAME": "GetItem",
        "X-EBAY-API-SITEID": "0",
        "X-EBAY-API-COMPATIBILITY-LEVEL": "967",
        "X-EBAY-API-APP-NAME": account["APP_ID"],
        "X-EBAY-API-DEV-NAME": account["DEV_ID"],
        "X-EBAY-API-CERT-NAME": account["CERT_ID"],
        "Content-Type": "text/xml",
    }

    try:
        resp = requests.post(TRADING_API_URL, headers=headers,
                              data=xml_body.encode("utf-8"), timeout=20)
    except Exception as e:
        return None, "none", "GetItemリクエスト失敗: " + str(e)[:150]

    if resp.status_code != 200:
        return None, "none", "GetItem HTTP " + str(resp.status_code)

    try:
        # ★2026/08/05確定: resp.text ではなく resp.content(生バイト)を渡すこと。
        #   eBay GetItemのレスポンスヘッダはcharsetを明示しないことがあり、その場合
        #   requestsはRFC2616のデフォルトに従いISO-8859-1と誤って推測してresp.textを
        #   デコードしてしまう。実際のXML本文はUTF-8（"㎏"等の全角文字を含む）のため、
        #   resp.textを使うと文字化け（例: "㎏"→"ã\x8e\x8f"）した文字列がXMLパース後の
        #   値にそのまま残ってしまい、後段の正規表現(_SHIP_PROFILE_WEIGHT_RE)が絶対に
        #   一致しなくなっていた（診断スクリプトdiag_weight_check.pyで実注文15件全件が
        #   この文字化けで重量取得失敗することを確認・再現済み）。resp.content(bytes)を
        #   渡せばElementTreeがXML宣言の encoding="utf-8" を正しく使ってパースするため、
        #   文字化けせず正しい文字列が得られる。
        root = ET.fromstring(resp.content)
    except Exception as e:
        return None, "none", "GetItemレスポンス解析エラー: " + str(e)[:150]

    ack = root.findtext("e:Ack", default="", namespaces=TRADING_NS)
    if ack not in ("Success", "Warning"):
        errors = root.findall(".//e:Errors", TRADING_NS)
        msgs = [(e.findtext("e:ShortMessage", default="", namespaces=TRADING_NS) or "")
                for e in errors]
        return None, "none", "GetItem Ack=" + str(ack) + " " + "; ".join(m for m in msgs if m)[:150]

    item = root.find(".//e:Item", TRADING_NS)
    if item is None:
        return None, "none", "GetItem: Item要素が見つかりません"

    # ① 最優先: Shipping Policy名から実重量を抽出（例: "1.5kg_202605" → 1.5）
    profile_name = item.findtext(
        "e:SellerProfiles/e:SellerShippingProfile/e:ShippingProfileName",
        default="", namespaces=TRADING_NS)
    if profile_name:
        m = _SHIP_PROFILE_WEIGHT_RE.search(profile_name)
        if m:
            try:
                w = float(m.group(1))
                if w > 0:
                    return w, "shipping_policy(" + profile_name + ")", None
            except ValueError:
                pass

    # ② 次点: ShippingPackageDetailsの実測重量（設定されている場合。lb/oz想定）
    weight_major = item.findtext("e:ShippingPackageDetails/e:WeightMajor",
                                  default="", namespaces=TRADING_NS)
    weight_minor = item.findtext("e:ShippingPackageDetails/e:WeightMinor",
                                  default="", namespaces=TRADING_NS)
    try:
        lb = float(weight_major) if weight_major else 0.0
        oz = float(weight_minor) if weight_minor else 0.0
        total_lb = lb + oz / 16.0
        if total_lb > 0:
            return round(total_lb * 0.45359237, 3), "package_weight_lb_oz", None
    except ValueError:
        pass

    return (None, "none",
            "Shipping Policy名(" + repr(profile_name) +
            ")からもPackage重量からも取得できず")


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# CPaSS タブ別URL（HTMLダンプで確認済み）
# 2026/08/02: CPaSSは発送元地域ごとに専用サイトを使う仕様に変更されており、地域指定なし
# (/jp/を含まない)URLでアクセスすると日本発送用のShipping Labelを購入できない状態になる
# (CPaSS「最新のお知らせ」モーダル・戸井さん実機確認・指示により判明)。全URLに/jp/を追加。
CPASS_ENTRY_URL = "https://cpass.ebay.com/jp/order/paid"
CPASS_TAB_URLS = {
    "発送手続き待ち": "https://cpass.ebay.com/jp/order/paid",
    "発送手続き":     "https://cpass.ebay.com/jp/order/readytoship",
    "キャンセル":     "https://cpass.ebay.com/jp/order/cancelled",
    "出荷待ち":       "https://cpass.ebay.com/jp/order/labelprinted",
    "出荷":           "https://cpass.ebay.com/jp/order/intransit",
}


def _dismiss_shipping_origin_modal(page):
    """CPaSSログインページに新しく現れるようになった配送元国選択モーダルを閉じる（2026/08/02追加）"""
    try:
        japan_btn = page.locator('text="Japan"').first
        if japan_btn.is_visible(timeout=2000):
            japan_btn.click(timeout=2000)
            time.sleep(0.5)
            continue_btn = page.locator('button:has-text("Continue")').first
            if continue_btn.is_visible(timeout=2000):
                continue_btn.click(timeout=2000)
                time.sleep(1.5)
                print("    [OK] 配送元国選択モーダルを閉じた(Japan/Continue)")
                return True
    except Exception as e:
        print("    [DEBUG] Japan/Continue経路失敗: " + str(e)[:80])

    try:
        for sel in ['[role="dialog"] button:has-text("x")', 'button[aria-label="Close"]',
                    '.ant-modal-close', '[role="dialog"] svg']:
            close_btn = page.locator(sel).first
            if close_btn.is_visible(timeout=1500):
                close_btn.click(timeout=1500)
                time.sleep(1)
                print("    [OK] 配送元国選択モーダルを閉じた(closeボタン, " + sel + ")")
                return True
    except Exception as e:
        print("    [DEBUG] closeボタン経路失敗: " + str(e)[:80])

    print("    [DEBUG] 配送元国選択モーダルは検出されず")
    return False


def _dismiss_announcement_modal(page):
    """CPaSS ログイン後の注文一覧ページに不定期に出る「最新のお知らせ」モーダルを閉じる
    （2026/08/02 六報で戸井さんが実機確認・「×を閉じて、進め」との指示を受けて追加）。

    ログインページの配送元国選択モーダル（_dismiss_shipping_origin_modal）とは別物。
    こちらはログイン後の注文一覧ページ（CPASS_ENTRY_URL等）で表示される、お知らせ一覧
    （例: 1/50ページ）のモーダルで、右上の×ボタンで閉じる。閉じないまま後続のチェック
    ボックス操作・確認ダイアログクリックを行うと、モーダルが画面をブロックして
    「発送手続き待ち→発送手続き」一括移動の確認ダイアログが閉じられない等の不具合の
    原因になる可能性があるため、ページ遷移直後に必ず呼び出す（存在しなければ何もしない）。
    """
    try:
        found = page.locator('text=最新のお知らせ').first
        if not found.is_visible(timeout=1500):
            return False
    except Exception:
        return False

    print("    [DEBUG] 「最新のお知らせ」モーダルを検出")

    for sel in ['[role="dialog"] button[aria-label="Close"]', '[role="dialog"] .ant-modal-close',
                '.ant-modal-close', 'button[aria-label="Close"]',
                '[role="dialog"] button:has-text("×")', '[role="dialog"] svg']:
        try:
            close_btn = page.locator(sel).first
            if close_btn.is_visible(timeout=1500):
                close_btn.click(timeout=2000)
                time.sleep(1)
                print("    [OK] 「最新のお知らせ」モーダルを閉じた（" + sel + "）")
                return True
        except Exception:
            continue

    print("    [WARN] 「最新のお知らせ」モーダルを検出したが閉じるボタンが見つからず")
    return False


def _login(page):
    """CPaSS にログイン"""
    print("CPaSS ログイン中...")
    page.goto(cpass_config.CPASS_LOGIN_URL, wait_until="domcontentloaded", timeout=45000)
    time.sleep(3)
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass

    _dismiss_shipping_origin_modal(page)
    time.sleep(1)

    # ユーザー名
    for sel in ['input[type="text"]', 'input[type="email"]', '#userid']:
        try:
            if page.locator(sel).first.is_visible(timeout=2000):
                page.locator(sel).first.fill(cpass_config.CPASS_EMAIL)
                break
        except Exception:
            pass

    # パスワード
    for sel in ['input[type="password"]', '#pass']:
        try:
            if page.locator(sel).first.is_visible(timeout=2000):
                page.locator(sel).first.fill(cpass_config.CPASS_PASSWORD)
                break
        except Exception:
            pass

    # サインインボタン
    for sel in ['button:has-text("サインイン")', 'button:has-text("Sign in")',
                'button[type="submit"]', 'input[type="submit"]']:
        try:
            if page.locator(sel).first.is_visible(timeout=2000):
                page.locator(sel).first.click()
                break
        except Exception:
            pass

    time.sleep(4)
    print("ログイン後URL: " + page.url)

    if "/login" in page.url:
        print("    [WARN] ログイン後もURLが/loginのまま -> ログイン失敗の疑い。モーダル閉じ->再試行します")
        _dismiss_shipping_origin_modal(page)
        time.sleep(1)
        for sel in ['input[type="text"]', 'input[type="email"]', '#userid']:
            try:
                if page.locator(sel).first.is_visible(timeout=2000):
                    page.locator(sel).first.fill(cpass_config.CPASS_EMAIL)
                    break
            except Exception:
                pass
        for sel in ['input[type="password"]', '#pass']:
            try:
                if page.locator(sel).first.is_visible(timeout=2000):
                    page.locator(sel).first.fill(cpass_config.CPASS_PASSWORD)
                    break
            except Exception:
                pass
        for sel in ['button:has-text("サインイン")', 'button:has-text("Sign in")',
                    'button[type="submit"]', 'input[type="submit"]']:
            try:
                if page.locator(sel).first.is_visible(timeout=2000):
                    page.locator(sel).first.click()
                    break
            except Exception:
                pass
        time.sleep(4)
        print("    再試行後URL: " + page.url)
        if "/login" in page.url:
            print("    [ERROR] 再試行後もログインに失敗しています。CPaSS側のUI変更やパスワード変更の可能性があります")


def _scrape_orders_from_page(page):
    """現在のページ（発送手続き待ち or 発送手続き）から注文情報を抽出

    ★ HTMLダンプ確認済み DOM 構造（2026/05/25）:
      .pkg_wrapper
        .title_pkgnumber .value <a>2904</a>   ← パッケージ番号
        .order_num .value  "16-14672-37993"   ← 注文番号
        .txn_item .item_title a "商品タイトル" ← タイトル
    """
    extracted = page.evaluate(
        """() => {
            const results = [];
            const wrappers = document.querySelectorAll('.pkg_wrapper');
            for (const wrapper of wrappers) {
                // 注文番号
                const orderEl = wrapper.querySelector('.order_num .value');
                if (!orderEl) continue;
                const orderNo = (orderEl.textContent || '').trim();
                if (!orderNo) continue;

                // パッケージ番号
                const pkgEl = wrapper.querySelector('.title_pkgnumber .value');
                const pkgNo = pkgEl ? (pkgEl.textContent || '').trim().replace(/\\s+/g, '') : '';

                // 商品タイトル（リンクテキスト）
                let title = '';
                const titleLink = wrapper.querySelector('.item_title a')
                               || wrapper.querySelector('.txn_item_info a');
                if (titleLink) {
                    title = (titleLink.textContent || '').trim();
                }

                // アイテムID（タイトルリンクのhrefから）
                let itemId = '';
                if (titleLink) {
                    const href = titleLink.getAttribute('href') || '';
                    const m = href.match(/\\/itm\\/(\\d+)/);
                    if (m) itemId = m[1];
                }

                results.push({
                    package_no: pkgNo,
                    order_no: orderNo,
                    item_id: itemId,
                    title: title,
                });
            }
            return results;
        }"""
    )
    return extracted


def _scrape_all_orders_with_pagination(page, max_pages=20):
    """「発送手続き」タブの全ページを巡回して .pkg_wrapper を収集する

    ★2026/07/09追加・重大バグ修正:
    旧実装は _scrape_orders_from_page を1回呼ぶだけで、Ant Design の
    ページネーション（.ant-pagination、1ページ100件・2ページ目以降あり）を
    一切考慮していなかった。「発送手続き」タブに233件溜まっている状態で
    1ページ目の100件しか読まず、新規注文の大半（12件中11件）が
    「対象注文のうちCPaSS発送手続きタブで見つからないもの」として
    毎回スキップされ、DHL価格取得・BL/BR記入がほぼ0件になっていた
    （2026/07/09 戸井さん報告で発覚）。
    """
    all_orders = {}
    for page_idx in range(1, max_pages + 1):
        page_orders = _scrape_orders_from_page(page)
        for o in page_orders:
            all_orders[o["order_no"]] = o
        total_text = None
        has_next = False
        try:
            total_text = page.evaluate(
                """() => { const el = document.querySelector('.ant-pagination-total-text'); return el ? el.textContent : null; }"""
            )
            has_next = page.evaluate(
                """() => {
                    const next = document.querySelector('.ant-pagination-next');
                    if (!next) return false;
                    return next.getAttribute('aria-disabled') !== 'true'
                        && !next.classList.contains('ant-pagination-disabled');
                }"""
            )
        except Exception as e:
            print("    [WARN] ページネーション検出失敗: " + str(e)[:60])
        print("    [ページ" + str(page_idx) + "] " + str(len(page_orders)) + "件 (累計"
              + str(len(all_orders)) + "件" + (" / 全" + str(total_text) if total_text else "") + ")")
        if not has_next:
            break
        try:
            page.evaluate("""() => { document.querySelector('.ant-pagination-next').click(); }""")
        except Exception as e:
            print("    [WARN] 次ページクリック失敗: " + str(e)[:60])
            break
        time.sleep(2)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        time.sleep(1)
    return list(all_orders.values())


def _navigate_to_sidebar_tab(page, tab_label):
    """指定タブへ移動（直接URL goto を使用 → 確実）

    tab_label: "発送手続き待ち" / "発送手続き" / "キャンセル" 等
    """
    url = CPASS_TAB_URLS.get(tab_label)
    if url:
        print("  タブ移動「" + tab_label + "」→ " + url)
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
    else:
        print("  [警告] URL不明「" + tab_label + "」→ スキップ")
        return
    time.sleep(3)
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass
    print("    現在URL: " + page.url)


def _move_all_to_processing(page):
    """発送手続き待ち の全件を 発送手続き へ移動"""
    print("発送手続き待ち → 発送手続き へ移動...")
    # まず発送手続き待ちタブへ
    _navigate_to_sidebar_tab(page, "発送手続き待ち")
    time.sleep(2)
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass

    # 「すべて」のチェックボックスをクリック
    print("  全選択チェックボックスをクリック...")
    selected = False
    for sel in [
        'input[type="checkbox"]:near(:text("すべて"))',
        'label:has-text("すべて") input[type="checkbox"]',
        'span:has-text("すべて") >> xpath=.. >> input[type="checkbox"]',
    ]:
        try:
            page.locator(sel).first.check(timeout=2000)
            selected = True
            print("    [OK] " + sel)
            break
        except Exception:
            pass
    if not selected:
        # 「すべて」というテキスト要素を見つけてその近くのチェックボックスをクリック
        try:
            page.evaluate(
                """() => {
                    const labels = Array.from(document.querySelectorAll('*'))
                        .filter(el => (el.textContent || '').trim().startsWith('すべて'));
                    for (const lbl of labels) {
                        let elem = lbl;
                        for (let i = 0; i < 5; i++) {
                            elem = elem.parentElement;
                            if (!elem) break;
                            const cb = elem.querySelector('input[type="checkbox"]');
                            if (cb) { cb.click(); return true; }
                        }
                    }
                    return false;
                }"""
            )
            selected = True
            print("    [OK] JS-based すべてチェックボックス")
        except Exception as e:
            print("    [失敗] " + str(e)[:80])

    if not selected:
        print("  警告: 全選択チェックボックスが見つかりません。手動で実行してください")
        return False

    time.sleep(1)

    # 「発送手続き」ボタンをクリック（一括処理）
    print("  「発送手続き」一括ボタンをクリック...")
    clicked = False
    for sel in [
        'button:has-text("発送手続き")',
        'a:has-text("発送手続き")',
        '[role="button"]:has-text("発送手続き")',
    ]:
        try:
            page.locator(sel).first.click(timeout=3000)
            clicked = True
            print("    [OK] " + sel)
            break
        except Exception:
            pass

    if not clicked:
        print("  警告: 「発送手続き」ボタンが見つかりません")
        return False

    time.sleep(2)

    # 確認ダイアログが出た場合「確認」ボタンをクリック（最大12秒待機）
    # ★2026/07/14修正（重大バグ）: 実際のダイアログはAntD(.ant-modal)ではなく新UI
    #   (.sp-modal-dialog > .prompt-modal-body)で、確認ボタンのテキストは「確 認」
    #   （半角スペース入り）・クラスは ant-btn-primary ではなく ant-btn-default btn blue
    #   だった。旧セレクタが一切マッチせず、毎回「確認ダイアログなし or 閉じ済み」と誤判定
    #   してスキップしていたため、実際には対象パッケージが「発送手続き」タブへ一切移動され
    #   ていなかった（ログ・GitHub Actionsは共に「Success」と表示されるため誰も気づけない
    #   状態だった）。クリック後にダイアログが実際に消えたことまで検証してから次に進む。
    print("  確認ダイアログを待機...")
    dialog_closed = False
    for attempt in range(12):
        try:
            clicked = page.evaluate(
                """() => {
                    // ボタンテキストの空白を全て除去して「確認」「OK」と一致するものを探す
                    const btns = Array.from(document.querySelectorAll('button'));
                    for (const b of btns) {
                        const t = (b.textContent || '').replace(/\\s+/g, '');
                        if (t === '確認' || t === 'OK') {
                            b.click();
                            return true;
                        }
                    }
                    // フォールバック: 新UIの確認ダイアログ内で「閉じる」以外のボタン
                    // （＝確認・続行ボタン）をクリック
                    const modal = document.querySelector(
                        '.sp-modal-dialog .prompt-modal-body, .prompt-modal-body'
                    );
                    if (modal) {
                        const cand = Array.from(modal.querySelectorAll('button')).find(
                            b => (b.textContent || '').replace(/\\s+/g, '') !== '閉じる'
                        );
                        if (cand) { cand.click(); return true; }
                    }
                    return false;
                }"""
            )
            if clicked:
                time.sleep(1)
                still_open = page.evaluate(
                    "() => !!document.querySelector("
                    "'.sp-modal-dialog .prompt-modal-body, .prompt-modal-body')"
                )
                if not still_open:
                    print("    [OK] 確認ダイアログ → クリックで閉じたことを確認済み")
                    dialog_closed = True
                    break
                else:
                    print("    [WARN] クリックしたがダイアログがまだ残っている → 再試行")
        except Exception as e:
            print("    [WARN] 確認ダイアログ処理中に例外: " + str(e)[:80])
        time.sleep(1)
    if not dialog_closed:
        print("    [ERROR] 確認ダイアログを閉じられませんでした（対象パッケージが移動されて"
              "いない可能性が高い。CPaSS側を要確認）")

    time.sleep(3)
    return True


def _os_click(page, viewport_x, viewport_y):
    """CDP でウィンドウ実座標を取得 → ctypes で本物クリック"""
    import ctypes

    page.bring_to_front()
    time.sleep(0.5)

    # DPI対応（論理ピクセルで SetCursorPos を使う）
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    # Chromiumウィンドウを OS レベルで最前面に
    try:
        hwnd = ctypes.windll.user32.FindWindowW("Chrome_WidgetWin_1", None)
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            print("    Chrome窓フォーカス OK (hwnd=" + str(hwnd) + ")")
        else:
            print("    Chrome窓が見つからず")
    except Exception as fe:
        print("    Chrome窓フォーカス失敗: " + str(fe)[:60])

    # CDP でウィンドウの実スクリーン座標を取得
    try:
        cdp = page.context.new_cdp_session(page)
        wft   = cdp.send("Browser.getWindowForTarget")
        bnds  = cdp.send("Browser.getWindowBounds", {"windowId": wft["windowId"]})["bounds"]
        cdp.detach()

        inner_h = page.evaluate("() => window.innerHeight")
        chrome_h = max(0, bnds["height"] - inner_h)  # 負にならないよう補正

        screen_x = int(bnds["left"] + viewport_x)
        screen_y = int(bnds["top"]  + chrome_h + viewport_y)
        print("    OS クリック: bounds=(" + str(bnds["left"]) + "," + str(bnds["top"]) +
              " " + str(bnds["width"]) + "x" + str(bnds["height"]) + ")"
              " chrome_h=" + str(chrome_h) +
              " → screen(" + str(screen_x) + "," + str(screen_y) + ")")
    except Exception as e:
        # CDP 失敗時フォールバック: screenX/Y + viewport
        js = page.evaluate("() => ({sx: window.screenX, sy: window.screenY})")
        screen_x = int(js["sx"] + viewport_x)
        screen_y = int(js["sy"] + viewport_y)
        print("    OS クリック(fallback): screen(" + str(screen_x) + "," + str(screen_y) + ") err=" + str(e)[:60])

    ctypes.windll.user32.SetCursorPos(screen_x, screen_y)
    time.sleep(0.3)
    ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)   # LEFTDOWN
    time.sleep(0.1)
    ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)   # LEFTUP


def _dismiss_copyright_dialog(page):
    """CPaSS 著作権ダイアログ（合同会社Skillful Sailor Inc.）を自動で「はい」クリック

    ★2026/06/20 確認: 処理中に「著作権に関して」ダイアログが出ることがある。
    「はい(Y)」をクリックしないと後続のボタンクリックがタイムアウトする。
    """
    try:
        # HTMLモーダルとして「はい」ボタンを探す
        for sel in [
            'button:has-text("はい")',
            'button:has-text("はい(Y)")',
            'button:has-text("Yes")',
        ]:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=800):
                loc.click(timeout=2000)
                print("    [著作権ダイアログ] はい クリック OK")
                time.sleep(0.5)
                return True
    except Exception:
        pass
    return False


def _find_order_page(page, order_no, max_pages=20):
    """「発送手続き」タブでorder_noが載っているページまでめくって移動する。

    ★2026/07/09追加・重大バグ修正:
    _scrape_all_orders_with_pagination で全ページ分の注文一覧は正しく取得できる
    ようになったが、_open_detail_dialog 側は「今ブラウザに表示されているページ」
    の .pkg_wrapper しか見ておらず、全ページ巡回が終わった時点でブラウザは
    最後のページ（例:3ページ目）に取り残されたままだった。そのため対象注文が
    1・2ページ目にある場合は毎回「この注文番号自体が発送手続きタブに存在しません」
    となり、DHL価格取得が0件になっていた（2026/07/09 戸井さん報告・run#46で確認）。
    ここで改めて先頭ページから探し直し、見つかったページに留まってから
    編集ボタンを探す。
    """
    _navigate_to_sidebar_tab(page, "発送手続き")
    for page_idx in range(1, max_pages + 1):
        found = page.evaluate(
            """(orderNo) => {
                const wrappers = document.querySelectorAll('.pkg_wrapper');
                for (const w of wrappers) {
                    const val = w.querySelector('.order_num .value');
                    if (val && (val.textContent || '').trim() === orderNo) return true;
                }
                return false;
            }""",
            order_no,
        )
        if found:
            print("    [OK] order=" + order_no + " はページ" + str(page_idx) + "で発見")
            return True
        has_next = False
        try:
            has_next = page.evaluate(
                """() => {
                    const next = document.querySelector('.ant-pagination-next');
                    if (!next) return false;
                    return next.getAttribute('aria-disabled') !== 'true'
                        && !next.classList.contains('ant-pagination-disabled');
                }"""
            )
        except Exception:
            pass
        if not has_next:
            break
        try:
            page.evaluate("""() => { document.querySelector('.ant-pagination-next').click(); }""")
        except Exception:
            break
        time.sleep(2)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass
        time.sleep(1)
    print("    [DEBUG] 全ページを探索してもorder=" + order_no + " が見つかりません")
    return False


def _open_detail_dialog(page, order_no):
    """指定の order_no の「編集」ボタンをクリックしてダイアログを開く

    ★2026/06/20 確認: CPaSS「発送手続き」タブのボタンは「編集」。
    """
    # 著作権ダイアログを先に閉じる
    _dismiss_copyright_dialog(page)

    # ★2026/07/09追加: 編集ボタンを探す前に、対象注文が載っているページまで移動する
    if not _find_order_page(page, order_no):
        _save_screenshot(page, "cpass_action_ss.png")
        return False

    # 残ダイアログを閉じる
    for sel in ['button:has-text("閉じる")', 'button:has-text("Close")', '[aria-label="Close"]']:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=500):
                btn.click(timeout=1500)
                time.sleep(0.5)
                break
        except Exception:
            pass
    try:
        page.keyboard.press("Escape")
        time.sleep(0.3)
    except Exception:
        pass

    # order_no に対応する「編集」ボタンのインデックスを取得
    # ★2026/06/20 確認: CPaSS「発送手続き」タブのボタンは「編集」（「詳細を見る」ではない）
    # 各 .pkg_wrapper > .title_action に [配送を割り当て][編集][アクション▼] の3ボタンあり
    # 「アクション」ドロップダウンは ant-dropdown-trigger クラスがつくため除外
    btn_index = page.evaluate(
        """(orderNo) => {
            const allEditBtns = Array.from(
                document.querySelectorAll('.pkg_wrapper button')
            ).filter(b => {
                const txt = (b.textContent || '').trim();
                return txt.includes('編集') && !b.classList.contains('ant-dropdown-trigger');
            });
            const wrappers = Array.from(document.querySelectorAll('.pkg_wrapper'));
            for (let i = 0; i < wrappers.length; i++) {
                const val = wrappers[i].querySelector('.order_num .value');
                if (val && (val.textContent || '').trim() === orderNo) {
                    const btn = Array.from(wrappers[i].querySelectorAll('button'))
                        .find(b => {
                            const txt = (b.textContent || '').trim();
                            return txt.includes('編集') && !b.classList.contains('ant-dropdown-trigger');
                        });
                    if (!btn) return -1;
                    return allEditBtns.indexOf(btn);
                }
            }
            return -1;
        }""",
        order_no,
    )
    print("    編集ボタン index: " + str(btn_index))

    if btn_index < 0:
        print("    ボタンが見つかりません → スキップ")
        # ★2026/07/06追加: なぜ見つからないか原因調査用に該当行のボタン一覧をダンプ
        try:
            dbg_row = page.evaluate(
                """(orderNo) => {
                    const wrappers = Array.from(document.querySelectorAll('.pkg_wrapper'));
                    for (const w of wrappers) {
                        const val = w.querySelector('.order_num .value');
                        if (val && (val.textContent || '').trim() === orderNo) {
                            return Array.from(w.querySelectorAll('button'))
                                .map(b => (b.textContent || '').trim().slice(0, 20));
                        }
                    }
                    return null;
                }""",
                order_no,
            )
            if dbg_row is None:
                print("    [DEBUG] この注文番号自体が発送手続きタブに存在しません（.pkg_wrapper未検出）")
            else:
                print("    [DEBUG] 該当行のボタン一覧: " + json.dumps(dbg_row, ensure_ascii=False))
        except Exception as e:
            print("    [DEBUG] 行ボタン一覧取得失敗: " + str(e)[:60])
        _save_screenshot(page, "cpass_action_ss.png")
        return False

    detail_btns = page.locator('.pkg_wrapper button:not(.ant-dropdown-trigger):has-text("編集")')
    detail_btn = detail_btns.nth(btn_index)
    detail_btn.scroll_into_view_if_needed(timeout=5000)
    time.sleep(0.8)
    page.bring_to_front()

    # 著作権ダイアログが出ていれば閉じてからクリック
    _dismiss_copyright_dialog(page)

    # まず JS クリック（Playwright locator クリックがタイムアウトする場合の代替）
    clicked = False
    try:
        page.evaluate(
            """(idx) => {
                const btns = Array.from(document.querySelectorAll('.pkg_wrapper button'))
                    .filter(b => b.textContent.includes('編集') && !b.classList.contains('ant-dropdown-trigger'));
                if (btns[idx]) { btns[idx].scrollIntoView({block:'center'}); btns[idx].click(); return true; }
                return false;
            }""",
            btn_index,
        )
        clicked = True
        print("    編集ボタン JS クリック OK")
    except Exception as e:
        print("    JS クリック失敗: " + str(e)[:60])

    if not clicked:
        try:
            detail_btn.click(timeout=5000)
            clicked = True
            print("    編集ボタン Locator クリック OK")
        except Exception as e:
            print("    編集ボタン クリック失敗: " + str(e)[:80])
            _save_screenshot(page, "cpass_action_ss.png")
            return False

    # ダイアログが開くのを待つ（「閉じる」ボタンが出れば開いた判定）
    try:
        page.wait_for_selector(
            'button:has-text("閉じる"), button:has-text("Close"), '
            '.ant-modal-content, [role="dialog"], .ant-drawer-content',
            state="visible",
            timeout=10000,
        )
        print("    ダイアログ展開 OK")
        time.sleep(2)
        _save_screenshot(page, "cpass_action_ss.png")
        return True
    except Exception as e:
        print("    ダイアログ展開タイムアウト: " + str(e)[:80])
        _save_screenshot(page, "cpass_action_ss.png")
        return False


# 旧関数名の互換エイリアス
def _open_edit_dialog(page, order_no):
    return _open_detail_dialog(page, order_no)


def _save_screenshot(page, filename):
    """スクリーンショットをスクリプトと同じフォルダに保存"""
    try:
        ss_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        page.screenshot(path=ss_path)
        print("    スクリーンショット保存: " + ss_path)
    except Exception as e:
        print("    スクリーンショット失敗: " + str(e)[:60])


def _close_all_dialogs(page, max_attempts=4):
    """開いている全てのダイアログ/モーダル/ドロワーを強制的に閉じる

    ★2026/07/06追加: 前の注文の外側編集ダイアログが「保存する」後も閉じきらずに
    残ってしまうと、次の注文の編集ボタンをクリックした際に
    「もう開いている（＝前注文のダイアログ）」を誤って「開けた」と判定し、
    2件目以降のDHL取得が軒並み失敗する不具合が起きうる（2026/07/06 9件中2件のみ成功で発覚）。
    保存直後 と 次注文の編集ボタンを押す直前 の両方で呼び出し、
    ダイアログ系コンテナが画面上に一切残っていないことを確認してから次に進む。
    """
    for attempt in range(max_attempts):
        clicked = page.evaluate(
            """() => {
                let clicked = false;
                const closeTexts = ['閉じる', 'Close', 'キャンセル', 'Cancel'];
                const btns = Array.from(document.querySelectorAll(
                    'button, [role="button"], .ant-drawer-close, .ant-modal-close'
                ));
                for (const b of btns) {
                    if (b.offsetParent === null) continue;
                    const txt = (b.textContent || '').trim();
                    const isCloseBtn = closeTexts.includes(txt)
                        || b.classList.contains('ant-drawer-close')
                        || b.classList.contains('ant-modal-close')
                        || b.getAttribute('aria-label') === 'Close';
                    if (isCloseBtn) { b.click(); clicked = true; }
                }
                return clicked;
            }"""
        )
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        time.sleep(0.8)
        still_open = page.evaluate(
            """() => !!(document.querySelector('.ant-drawer-content')
                || document.querySelector('.ant-modal-content')
                || document.querySelector('[role="dialog"]'))"""
        )
        if not still_open:
            if attempt > 0:
                print("    [OK] 残存ダイアログを" + str(attempt + 1) + "回目のクローズ試行で解消")
            return True
        if not clicked:
            break
    print("    [WARN] 残存ダイアログが閉じきれていません（次の注文で誤検知の恐れあり）")
    return False


def _fill_edit_form_and_save(page, weight_kg, length_cm, width_cm, height_cm, hs_code):
    """フォーム入力 → 配送を割り当て(DHL価格取得) → 保存する

    ★正しい順序（CLAUDE.md 2026/05/24 確認）:
    1. フォーム入力（重量・寸法・HSコード）
    2. 「配送を割り当て」クリック → DHL金額表示
    3. 「保存する」クリック

    Args:
        weight_kg: 単位重量(kg)に入力する推定重量（★2026/07/17確定: 「梱包」欄には入力しない）
        length_cm, width_cm, height_cm: サイズ
        hs_code: HSコード（10桁）

    Returns:
        tuple: (saved_ok: bool, dhl_price_jpy: int|None)
    """
    # 2026/08/02: dimension_weight_lookupのデフォルト値は0.5kgのため通常0にはならないが、
    # 万が一None/0が渡ってきた場合に備え、無効な重量でDHL見積もりが失敗するのを防ぐため
    # 最小値0.1kgにクランプする。
    try:
        if not weight_kg or float(weight_kg) <= 0:
            print("    [WARN] weight_kgが0以下のため0.1kgにフォールバック（元の値: " + str(weight_kg) + "）")
            weight_kg = 0.1
    except (TypeError, ValueError):
        print("    [WARN] weight_kgが不正な値のため0.1kgにフォールバック（元の値: " + str(weight_kg) + "）")
        weight_kg = 0.1

    print("  フォーム入力: 重量=" + str(weight_kg) +
          "kg, " + str(length_cm) + "x" + str(width_cm) + "x" + str(height_cm) +
          "cm, HS=" + hs_code)

    time.sleep(1)

    # 1. JS で input フィールドにラベルベースで入力
    filled = page.evaluate(
        """(args) => {
            // ヘルパー: label/前のテキストから input を見つける
            // ★2026/07/10確定バグ修正: input[type="text"/"number"]に限定すると、
            //   AntDのInputNumber(単位重量/梱包/長さ/幅/高さ)はtype属性が無く、
            //   AntDのSelect検索欄(HSコード/原産国)はtype="search"のため、
            //   どちらもここで拾えず5階層上まで探索が続き、結果として
            //   「告知のための商品名(英語)」等の無関係な最初のinput[type=text]を誤って
            //   掴んでいた（単位重量とHSコードが同一要素を指し、書き込みが競合していた）。
            //   型を限定せず「最初に見つかったinput/textarea」を返す方式に変更し、
            //   各ラベルに最も近い正しい入力欄だけを拾うようにした。
            function findInputNear(labelText) {
                const labels = Array.from(document.querySelectorAll('*'))
                    .filter(el => {
                        const t = (el.textContent || '').trim();
                        return t.startsWith(labelText) && el.children.length === 0;
                    });
                for (const lbl of labels) {
                    let parent = lbl;
                    for (let i = 0; i < 6; i++) {
                        parent = parent.parentElement;
                        if (!parent) break;
                        const inp = parent.querySelector('input, textarea');
                        if (inp) return inp;
                    }
                    // 次の兄弟要素から探す
                    let next = lbl.nextElementSibling;
                    while (next) {
                        const inp = next.querySelector('input');
                        if (inp) return inp;
                        next = next.nextElementSibling;
                    }
                }
                return null;
            }
            function setValue(input, value) {
                if (!input) return false;
                const proto = input.tagName === 'TEXTAREA'
                    ? window.HTMLTextAreaElement.prototype
                    : window.HTMLInputElement.prototype;
                const nativeSetter = Object.getOwnPropertyDescriptor(proto, 'value').set;
                nativeSetter.call(input, String(value));
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                // 2026/08/02: AntD InputNumber等はonBlurで内部stateを確定コミットする
                // ものがあり、input/changeだけでは「配送を割り当て」クリック時にまだ古い
                // (0の)値をバックエンドへ送ってしまうタイミング競合が起きうる。
                // 明示的にfocus→blurを発火させ、コミットを確実に発生させる。
                try { input.focus(); } catch (e) {}
                input.dispatchEvent(new Event('blur', { bubbles: true }));
                try { input.blur(); } catch (e) {}
                return true;
            }
            const results = {};

            // ★2026/07/17確定: 「梱包」欄への書き込みは廃止。
            //   単位重量と梱包の両方に同じ推定重量を入れると、CPaSS側が自動的に合算してしまい
            //   （例: 各2kg→合計4kg扱い）、実際より重い扱いになってDHL送料が過大に計算される
            //   バグが判明（戸井さん指摘）。以後は「単位重量」欄だけに入力し、「梱包」欄は
            //   触らない（0のまま）。過去にCPaSS上へ書き込み済みの「梱包」値は遡って修正しない
            //   （戸井さん指示）。
            results.weight = null;

            // 単位重量（アイテムリストの「単位重量(kg)」入力欄）
            // ★2026/07/10確定: eBay継承値が0のままだとDHL価格パネルが無限ロードして
            //   「選択」ボタンが出てこないバグが判明（run#45〜#51で対象11件が該当・戸井さん指示）
            //   → Google検索ベースの推定重量(weight_kg)で明示的に上書きする。重量(Net)欄は触らない
            const uInput = findInputNear('単位重量');
            results.unit_weight = setValue(uInput, args.weight);

            // 長さ・幅・高さ
            const lInput = findInputNear('長さ');
            results.length = setValue(lInput, args.length);
            const wdInput = findInputNear('幅');
            results.width = setValue(wdInput, args.width);
            const hInput = findInputNear('高さ');
            results.height = setValue(hInput, args.height);

            // HSコード
            const hsInput = findInputNear('HSコード');
            results.hs = setValue(hsInput, args.hs);

            return results;
        }""",
        {"weight": weight_kg, "length": length_cm, "width": width_cm,
         "height": height_cm, "hs": hs_code},
    )

    print("    入力結果: " + json.dumps(filled, ensure_ascii=False))
    time.sleep(1)

    # ★2026/07/05v3: 入力後の実値を読み戻して検証（Reactが値を戻していないか確認）
    try:
        verify = page.evaluate(
            """() => {
                function findInputNear(labelText) {
                    const labels = Array.from(document.querySelectorAll('*'))
                        .filter(el => {
                            const t = (el.textContent || '').trim();
                            return t.startsWith(labelText) && el.children.length === 0;
                        });
                    for (const lbl of labels) {
                        let parent = lbl;
                        for (let i = 0; i < 6; i++) {
                            parent = parent.parentElement;
                            if (!parent) break;
                            const inp = parent.querySelector('input, textarea');
                            if (inp) return inp;
                        }
                        let next = lbl.nextElementSibling;
                        while (next) {
                            const inp = next.querySelector('input');
                            if (inp) return inp;
                            next = next.nextElementSibling;
                        }
                    }
                    return null;
                }
                const out = {};
                for (const lbl of ['梱包', '単位重量', '長さ', '幅', '高さ', 'HSコード']) {
                    const inp = findInputNear(lbl);
                    out[lbl] = inp ? inp.value : null;
                }
                return out;
            }"""
        )
        print("    [DEBUG] 入力後の実値: " + json.dumps(verify, ensure_ascii=False))

        # 2026/08/02: 単位重量の読み戻しが0/空のままなら、まだReact側のstateコミットが
        # 間に合っていない可能性が高いため、もう一度同じ値を入力し直し、コミットのための
        # 待機時間を追加で取る（「無効な重量」でDHL見積もりが失敗する既知バグへの対策）。
        uw = verify.get("単位重量") if verify else None
        try:
            uw_val = float(uw) if uw not in (None, "") else 0.0
        except (TypeError, ValueError):
            uw_val = 0.0
        if uw_val <= 0:
            print("    [WARN] 単位重量の読み戻しが0のため再入力を実施します")
            page.evaluate(
                """(args) => {
                    function findInputNear(labelText) {
                        const labels = Array.from(document.querySelectorAll('*'))
                            .filter(el => {
                                const t = (el.textContent || '').trim();
                                return t.startsWith(labelText) && el.children.length === 0;
                            });
                        for (const lbl of labels) {
                            let parent = lbl;
                            for (let i = 0; i < 6; i++) {
                                parent = parent.parentElement;
                                if (!parent) break;
                                const inp = parent.querySelector('input, textarea');
                                if (inp) return inp;
                            }
                        }
                        return null;
                    }
                    const uInput = findInputNear('単位重量');
                    if (!uInput) return false;
                    const nativeSetter = Object.getOwnPropertyDescriptor(
                        window.HTMLInputElement.prototype, 'value').set;
                    nativeSetter.call(uInput, String(args.weight));
                    uInput.dispatchEvent(new Event('input', { bubbles: true }));
                    uInput.dispatchEvent(new Event('change', { bubbles: true }));
                    try { uInput.focus(); } catch (e) {}
                    uInput.dispatchEvent(new Event('blur', { bubbles: true }));
                    try { uInput.blur(); } catch (e) {}
                    return true;
                }""",
                {"weight": weight_kg},
            )
            time.sleep(2)
    except Exception as e:
        print("    [DEBUG] 実値検証失敗: " + str(e)[:60])

    # 2026/08/02: 入力直後にすぐ「配送を割り当て」を押すと、CPaSS側のstate反映
    # (デバウンス等)に間に合わず古い重量(0)でバックエンドの価格計算APIが呼ばれてしまい、
    # 結果として配送業者一覧が正しく表示されず「DHL「選択」ボタンが見つかりません」
    # (バグ①)につながっていた疑いがある。人間が手動操作する場合は自然に数秒の間が
    # 空くため再現しなかった。→ 明示的なバッファ待機を追加する。
    time.sleep(2)

    # 2. 「配送を割り当て」クリック → 内側モーダルが開く → DHL「選択」→価格取得
    print("  「配送を割り当て」クリック...")
    dhl_price = None
    clicked_assign = False
    # ★2026/07/05修正: 保存ボタンと同じJS座標クリック方式に変更
    #   - button以外（a / [role=button] / span系）も対象
    #   - テキストは「配送を割り当て」→「割り当て」→「assign」の順で緩めて検索
    #   - ツールバーの #btnAssignShipping は除外（誤クリック防止）
    rect_assign = page.evaluate(
        """() => {
            const texts = ['配送を割り当て', '割り当て', 'assign'];
            const containers = [
                document.querySelector('.ant-drawer-body'),
                document.querySelector('.ant-modal-content'),
                document.querySelector('[role="dialog"]'),
                document.body
            ].filter(Boolean);
            for (const t of texts) {
                for (const container of containers) {
                    const cands = Array.from(container.querySelectorAll(
                        'button, a, [role="button"], span[class*="btn"], div[class*="btn"]'
                    )).filter(el => {
                        if (el.id === 'btnAssignShipping') return false;
                        if (el.closest('#btnAssignShipping')) return false;
                        if (el.offsetParent === null) return false;
                        const txt = (el.textContent || '').trim();
                        if (!txt || txt.length > 30) return false;
                        return txt.toLowerCase().includes(t.toLowerCase());
                    });
                    for (const el of cands) {
                        el.scrollIntoView({block: 'center'});
                        const r = el.getBoundingClientRect();
                        if (!r.width || !r.height) continue;
                        const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
                        const hit = document.elementFromPoint(cx, cy);
                        // ★前面に別要素（ダイアログのオーバーレイ等）が被っている候補は除外
                        if (!hit || !(el.contains(hit) || hit.contains(el))) continue;
                        return {x: cx, y: cy,
                                txt: (el.textContent || '').trim(),
                                tag: el.tagName, matched: t};
                    }
                }
            }
            return null;
        }"""
    )
    if rect_assign and rect_assign.get('x'):
        page.mouse.click(rect_assign['x'], rect_assign['y'])
        clicked_assign = True
        print("    [OK] JS座標クリック: <" + str(rect_assign.get('tag')) + "> '"
              + str(rect_assign.get('txt')) + "' (match=" + str(rect_assign.get('matched')) + ")")
    else:
        # fallback: 従来のlocator方式
        for sel in [
            '.ant-drawer-body button:has-text("配送を割り当て")',
            '.ant-modal-content button:has-text("配送を割り当て")',
            '[role="dialog"] button:has-text("配送を割り当て")',
            'button:has-text("配送を割り当て"):not(#btnAssignShipping)',
        ]:
            try:
                page.locator(sel).first.click(timeout=3000)
                clicked_assign = True
                print("    [OK] " + sel)
                break
            except Exception:
                pass

    if clicked_assign:
        # ★2026/07/13確定バグ修正: CPaSSの「配送を割り当て」パネルが.ant-modalベースの
        #   旧UIから、独自クラス(.sp-modal-dialog/.sp-modal-content)ベースの新UIに変わっていた。
        #   実機調査（27-14854-47644で確認）の結果、DHL等の各配送業者は
        #   .shipping_method > .shipping_service（サービス名）+ .action（「選 択」ボタン）
        #   という構造になっており、旧セレクタ('.ant-modal'/.ant-card等)に一切マッチしないため
        #   毎回「割り当てパネルが15秒以内に出現せず」「DHL「選択」ボタンが見つかりません」に
        #   なっていた（DHL自体は選択可能な注文でも常に失敗していた）。
        #   さらに「選択」ボタンの実テキストは「選 択」（間に半角スペース）で、
        #   旧正規表現 /^(選択|select)$/i の完全一致にも失敗していた。
        #   → 新UIのセレクタ(.sp-modal-content/.sp-modal-dialog/.shipping_method)を追加し、
        #     ボタンテキストは空白除去してから比較するよう修正。
        SP_MODAL_SEL = '.sp-modal-content, .sp-modal-dialog, .assign_shipping'
        panel_js = (
            """() => {
                if (document.querySelector('""" + SP_MODAL_SEL + """')) return true;
                if (document.querySelector('.ant-modal')) return true;
                const cands = Array.from(document.querySelectorAll('button, a, [role="button"], .action'))
                    .filter(el => {
                        const t = (el.textContent || '').replace(/\\s+/g, '');
                        return /^(選択|select)$/i.test(t) && el.offsetParent !== null;
                    });
                return cands.length > 0
                    || (document.body.innerText || '').includes('Friendly Reminder');
            }"""
        )
        try:
            page.wait_for_function(panel_js, timeout=15000)
            print("    割り当てパネル出現 OK")
        except Exception:
            print("    [WARN] 割り当てパネルが15秒以内に出現せず")
        time.sleep(2)
        _save_screenshot(page, "cpass_after_assign_ss.png")

        # ★★★ 内側モーダル内のDHL「個別価格」を取得（2026/07/03 修正・2026/07/13新UI対応）★★★
        # 旧実装はレンジ表示「X - Y JPY」の上限(Y)を返していたため全注文が同額(8995等)になっていた。
        # 新実装: DHL行の「選択」をクリック→実計算価格（単一の「X,XXX JPY」）を優先取得。
        # ★2026/07/13再修正: SP_MODAL_SEL(.sp-modal-content/.sp-modal-dialog)はCPaSS内の
        #   他の無関係なダイアログ（例:「差出人の住所編集」）とクラス名が衝突しており、
        #   document.querySelector()がDOM順で先に出てくる無関係な方を掴んでしまうことが
        #   実機調査（27-14854-47644）で確認された。.assign_shippingは実機調査で他要素と
        #   衝突しない一意なクラスと確認済みのため、これを最優先で使う。
        modal_text = page.evaluate(
            """() => {
                const modal = document.querySelector('.assign_shipping')
                    || document.querySelector('.sp-modal-content')
                    || document.querySelector('.sp-modal-dialog')
                    || document.querySelector('.ant-modal');
                return modal ? (modal.innerText || '').slice(0, 1500) : null;
            }"""
        )
        if modal_text:
            print("    [DEBUG] 内側モーダル: " + modal_text.replace("\n", " | ")[:800])
        else:
            # ★2026/07/05: モーダルが無い場合は画面全体のテキストをダンプ（原因調査用）
            page_text = page.evaluate(
                """() => {
                    const t = document.body.innerText || '';
                    const i = t.indexOf('配送概要');
                    return t.slice(Math.max(0, i), Math.max(0, i) + 1800);
                }"""
            )
            print("    [DEBUG] モーダルなし。画面テキスト: "
                  + (page_text or "").replace("\n", " | ")[:1000])

        # ★2026/07/05v2: DHL「選択」をページ全体から検索（.ant-modal限定を廃止）＋最大18秒ポーリング
        # ★2026/07/13追加: 新UIの .shipping_method / .shipping_service を行候補に追加。
        #   ボタンテキストの空白（「選 択」）を除去してから比較するよう修正。
        pick_js = (
            """() => {
                const roots = [document.querySelector('.assign_shipping'),
                               document.querySelector('.sp-modal-content'),
                               document.querySelector('.sp-modal-dialog'),
                               document.querySelector('.ant-modal'), document].filter(Boolean);
                for (const root of roots) {
                    const rows = Array.from(root.querySelectorAll(
                        'tr, .ant-list-item, .ant-card, .shipping_method, .shipping_service, '
                        + '[class*="item"], [class*="row"], [class*="card"], [class*="shipping_method"], li'));
                    const dhlRows = rows.filter(r => (r.textContent || '').toLowerCase().includes('dhl')
                        && (r.textContent || '').length < 1200);
                    for (const row of dhlRows) {
                        const btn = Array.from(row.querySelectorAll('button, a, [role="button"], .action'))
                            .find(b => {
                                const t = (b.textContent || '').replace(/\\s+/g, '');
                                return /^(選択|select)$/i.test(t) && b.offsetParent !== null;
                            });
                        if (btn) { btn.scrollIntoView({block: 'center'}); btn.click(); return true; }
                    }
                    for (const row of dhlRows) {
                        const radio = row.querySelector('input[type="radio"]');
                        if (radio && !radio.checked && radio.offsetParent !== null) { radio.click(); return true; }
                    }
                }
                return false;
            }"""
        )
        picked = False
        # 2026/08/02: 18秒→25秒に延長。価格計算が非同期のため行(.shipping_method等)
        # 自体がまだ描画されていないタイミングでポーリングが打ち切られていた可能性が
        # あるための保険的措置。
        _deadline = time.time() + 25
        while time.time() < _deadline and not picked:
            picked = page.evaluate(pick_js)
            if not picked:
                time.sleep(1.5)
        if picked:
            print("    [OK] DHL「選択」クリック")
            time.sleep(2)
            # ★2026/07/13 四次修正（実機調査で判明）: 「選択」をクリックすると内側の
            #   .assign_shippingパネルは閉じてしまい、外側の詳細ダイアログの「配送概要」に
            #   「見積もり」ボタンが表示される。このボタンを押さないと推定配送料は「-」の
            #   ままで、確定した個別価格（.quote_line内）が表示されないことを
            #   27-14854-47644で実機確認した（旧実装はこのステップが存在しなかった）。
            _quote_rect = page.evaluate(
                """() => {
                    const btn = Array.from(document.querySelectorAll('button, a, [role="button"]'))
                        .find(b => (b.textContent || '').trim() === '見積もり' && b.offsetParent !== null);
                    if (!btn) return null;
                    const r = btn.getBoundingClientRect();
                    return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                }"""
            )
            if _quote_rect and _quote_rect.get('x'):
                page.mouse.click(_quote_rect['x'], _quote_rect['y'])
                print("    [OK] 「見積もり」クリック")
                time.sleep(2)
            else:
                print("    [WARN] 「見積もり」ボタンが見つかりません（価格が空欄になる可能性）")
        else:
            print("    [WARN] DHL「選択」ボタンが見つかりません")
            # 2026/08/02: 次回以降の原因切り分けのため、配送業者一覧
            # (.shipping_method等)がそもそも描画されていたか、描画されていたなら
            # 各行の中身(DHLの有無・「選択」ボタンの有無)を実機再現せずログだけで
            # 確認できるようダンプする。
            try:
                rows_dbg = page.evaluate(
                    """() => {
                        const roots = [document.querySelector('.assign_shipping'),
                                       document.querySelector('.sp-modal-content'),
                                       document.querySelector('.sp-modal-dialog'),
                                       document.querySelector('.ant-modal'), document].filter(Boolean);
                        for (const root of roots) {
                            const rows = Array.from(root.querySelectorAll(
                                '.shipping_method, .shipping_service, tr, .ant-list-item, .ant-card, li'));
                            if (rows.length) {
                                return {
                                    root_matched: root.className || root.tagName,
                                    row_count: rows.length,
                                    rows: rows.slice(0, 15).map(r => (r.textContent || '').replace(/\\s+/g,' ').trim().slice(0, 80))
                                };
                            }
                        }
                        return {root_matched: null, row_count: 0, rows: []};
                    }"""
                )
                print("    [DEBUG] 配送業者行一覧(" + str(rows_dbg.get("row_count")) + "件, root="
                      + str(rows_dbg.get("root_matched")) + "): "
                      + json.dumps(rows_dbg.get("rows"), ensure_ascii=False)[:1200])
            except Exception as e:
                print("    [DEBUG] 配送業者行一覧の取得失敗: " + str(e)[:60])

        # ★2026/07/09 重大バグ修正:
        # 旧実装は「選択」ボタンが見つからない/割り当てパネルが出ない場合でも
        # document全体からdhlを含む要素を検索するフォールバックを実行していた。
        # このフォールバックは要素の可視性（表示中かどうか）を一切見ておらず、
        # 前の注文で開いたDHL価格パネルが閉じきらずDOM上に残っている（display:none等で
        # 非表示だがquerySelectorAllには引っかかる）場合、そのまま前注文の価格を
        # 拾ってしまい、複数注文が同一金額（例:全件10078円）になるバグが発生していた
        # （2026/07/09 戸井さん報告・通常7/9・専門7/9ファイルで確認）。
        # → 「選択」クリックに成功した場合、または本物の割り当てパネル(.ant-modal等)が
        #   実際に出現した場合のみ価格検索を行う。それ以外は価格なし（None）のまま
        #   保存し、後で戸井さんが手動確認できるよう空欄で残す方が安全。
        _price_result = None
        if picked or modal_text:
            _price_js = (
                """() => {
                    const rangeSrc = '([\\\\d,]+)\\\\s*[-〜~–]\\\\s*([\\\\d,]+)\\\\s*JPY';
                    function isVisible(el) {
                        return !!(el && el.offsetParent !== null);
                    }
                    function findPrice(root) {
                        if (!root) return null;
                        const all = Array.from(root.querySelectorAll('*'));
                        const dhlNodes = all.filter(el => {
                            const t = (el.textContent || '');
                            return t.toLowerCase().includes('dhl') && el.children.length === 0
                                && isVisible(el);
                        });
                        for (const node of dhlNodes) {
                            let card = node;
                            for (let i = 0; i < 8; i++) {
                                card = card.parentElement;
                                if (!card) break;
                                if (!isVisible(card)) continue;
                                const txt = card.textContent || '';
                                if (!txt.toLowerCase().includes('dhl')) continue;
                                // レンジ表記を除去→残った単一価格＝実計算価格
                                const cleaned = txt.replace(new RegExp(rangeSrc, 'g'), ' ');
                                const singles = Array.from(cleaned.matchAll(/([\\d,]{3,})\\s*JPY/g))
                                    .map(m => parseInt(m[1].replace(/,/g, ''), 10)).filter(n => n >= 100);
                                if (singles.length) return { price: singles[0], src: 'single' };
                                const m = txt.match(new RegExp(rangeSrc));
                                if (m) return { price: parseInt(m[2].replace(/,/g, ''), 10), src: 'range-max' };
                            }
                        }
                        return null;
                    }
                    // ★2026/07/13 四次修正（実機調査で判明）:
                    //   DHL「選択」クリック後は.assign_shippingパネルが閉じてしまい、
                    //   代わりに外側ダイアログの「配送概要」セクション内の.quote_line要素に
                    //   確定した個別価格（「推定配送料 XX,XXX JPY」）が表示される。
                    //   「見積もり」ボタンをクリックした後はここが最も確実な取得元。
                    //   DHL以外の別配送業者の.quote_lineを誤って拾わないよう、
                    //   コンテナに「推定配送料」と「dhl」の両方が含まれる場合のみ採用する。
                    function findQuoteLine() {
                        const lines = Array.from(document.querySelectorAll('.quote_line')).filter(isVisible);
                        for (const line of lines) {
                            const container = line.closest('.left_carrier') || line.parentElement;
                            const ctext = (container ? container.textContent : line.textContent) || '';
                            if (!ctext.includes('推定配送料')) continue;
                            if (!ctext.toLowerCase().includes('dhl')) continue;
                            const cleaned = ctext.replace(new RegExp(rangeSrc, 'g'), ' ');
                            const singles = Array.from(cleaned.matchAll(/([\\d,]{3,})\\s*JPY/g))
                                .map(m => parseInt(m[1].replace(/,/g, ''), 10)).filter(n => n >= 100);
                            if (singles.length) return { price: singles[0], src: 'quote_line' };
                        }
                        return null;
                    }
                    // ★2026/07/13 二次修正（根本原因判明）:
                    //   SP_MODAL_SEL('.sp-modal-content, .sp-modal-dialog, .assign_shipping')は
                    //   CPaSS内の別の無関係なダイアログ（例:「差出人の住所編集」）と同じクラス名を
                    //   共有しており、document.querySelector()はDOM順で最初に出てくる要素を返すため、
                    //   実際のDHLパネルが開いていても無関係な方（DHLテキストなし）を掴んでしまい、
                    //   findPriceが毎回nullを返してdocument全体フォールバックに落ちていた。
                    //   document全体フォールバックはページ上の無関係な要素（別注文の情報等）を
                    //   拾ってしまう危険があり、これが「全注文が同一の誤った価格(4687円)になる」
                    //   バグの直接原因だった（run#74で確認）。
                    //   → .assign_shippingは実機調査で他要素と衝突しない一意なクラスと確認済み
                    //     なのでこれを最優先root候補にし、document全体へのフォールバックは廃止。
                    //     .assign_shipping等の本物のパネルが見つからない場合は価格なし(null)の
                    //     まま返す（誤った価格を書き込むより空欄の方が安全）。
                    const panel = document.querySelector('.assign_shipping')
                        || document.querySelector('.sp-modal-content')
                        || document.querySelector('.sp-modal-dialog')
                        || document.querySelector('.ant-modal')
                        || document.querySelector('.ant-drawer-body');
                    return findQuoteLine() || findPrice(panel);
                }"""
            )
            # ★2026/07/05v2: 価格が非同期計算されるため最大12秒ポーリング
            _deadline2 = time.time() + 12
            while time.time() < _deadline2:
                _price_result = page.evaluate(_price_js)
                if _price_result and _price_result.get("price"):
                    break
                time.sleep(1.5)
        else:
            print("    [WARN] 割り当てパネル未検出のため価格検索をスキップ（誤った価格の書き込み防止）")
        dhl_price = None
        if _price_result and _price_result.get("price"):
            dhl_price = _price_result["price"]
            if _price_result.get("src") in ("single", "quote_line"):
                print("    DHL個別価格: " + str(dhl_price) + " JPY (src=" + str(_price_result.get("src")) + ")")
            else:
                print("    [WARN] 個別価格が見つからずレンジ上限を使用: " + str(dhl_price) + " JPY")
        else:
            print("    DHL価格取得失敗（価格なしで保存を続行）")

        # 内側モーダル（.ant-modal または新UIの.assign_shipping）の「閉じる」ボタンを閉じる
        # ★外側ダイアログの「閉じる」と区別するため内側モーダル内のみを対象にする
        # ★2026/07/13 四次修正（実機調査で判明・重大）: 「見積もり」導線が判明したことで、
        #   DHL「選択」クリック後は.assign_shippingパネルが自動的に閉じる（DOM上から消える）
        #   ことを確認した。ところがこのステップは従来.sp-modal-content/.sp-modal-dialogにも
        #   フォールバックしていたため、.assign_shippingが既に消えている状態だと**外側の
        #   詳細/編集ダイアログ自身**（これも同じクラスを共有）を掴んでその「閉じる」ボタンを
        #   押してしまい、保存前に編集内容ごと破棄する重大なリスクがあった
        #   （さらに閉じるボタンが見つからない場合はEscapeキーを押していたが、これも外側
        #   ダイアログを閉じてしまう可能性があった）。
        #   → .assign_shipping（新UI）または.ant-modal（旧UI）が実際に存在する場合のみ閉じる
        #     操作を行い、どちらも存在しない（＝選択時に自動で閉じていた）場合は何もしない。
        #     Escapeキーのフォールバックも廃止（外側ダイアログを誤って閉じるリスクがあるため）。
        print("  内側モーダルを閉じる...")
        closed = False
        rect2 = page.evaluate(
            """() => {
                const modal = document.querySelector('.assign_shipping')
                    || document.querySelector('.ant-modal');
                if (modal) {
                    const btns = Array.from(modal.querySelectorAll('button, [role="button"], a'));
                    const btn = [...btns].reverse().find(b => (b.textContent||'').trim() === '閉じる' && b.offsetParent !== null);
                    if (btn) {
                        const r = btn.getBoundingClientRect();
                        return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                    }
                    // ×ボタン（旧UI: .ant-modal-close / 新UI: sp-modal内の閉じるアイコン）
                    const closeBtn = modal.querySelector('.ant-modal-close')
                        || modal.querySelector('[class*="close"]');
                    if (closeBtn) {
                        const r = closeBtn.getBoundingClientRect();
                        return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                    }
                }
                return null;
            }"""
        )
        if rect2 and rect2.get('x'):
            page.mouse.click(rect2['x'], rect2['y'])
            time.sleep(1.5)
            closed = True
            print("    閉じる [OK]")
        else:
            print("    内側パネルは既に閉じていました（想定通り）")
    else:
        print("    警告: 「配送を割り当て」ボタンが見つかりません")
        # ★2026/07/05追加: 原因調査用にダイアログ内の全クリック要素をログ出力
        try:
            dbg = page.evaluate(
                """() => {
                    const dlg = document.querySelector('.ant-drawer-body')
                        || document.querySelector('.ant-modal-content')
                        || document.querySelector('[role="dialog"]')
                        || document.body;
                    return Array.from(dlg.querySelectorAll('button, a, [role="button"]'))
                        .filter(el => el.offsetParent !== null)
                        .map(el => el.tagName + ':' + (el.textContent || '').trim().slice(0, 20))
                        .filter(t => t.split(':')[1])
                        .slice(0, 60);
                }"""
            )
            print("    [DEBUG] ダイアログ内クリック要素一覧: "
                  + json.dumps(dbg, ensure_ascii=False)[:1500])
        except Exception as e:
            print("    [DEBUG] 要素一覧取得失敗: " + str(e)[:60])
        _save_screenshot(page, "cpass_action_ss.png")

    # 3. 「保存する」ボタンを座標クリック（ダイアログ内を優先）
    time.sleep(1)
    saved = False
    rect_save = page.evaluate(
        """() => {
            // ダイアログ内を優先して探す（ツールバーボタンとの混同を防ぐ）
            const containers = [
                document.querySelector('.ant-drawer-body'),
                document.querySelector('.ant-modal-content'),
                document.querySelector('[role="dialog"]'),
                document.body
            ].filter(Boolean);
            for (const container of containers) {
                const btns = Array.from(container.querySelectorAll('button'));
                const btn = btns.find(b => (b.textContent||'').trim() === '保存する' && b.offsetParent !== null);
                if (btn) {
                    const r = btn.getBoundingClientRect();
                    return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                }
            }
            return null;
        }"""
    )
    if rect_save and rect_save.get('x'):
        page.mouse.click(rect_save['x'], rect_save['y'])
        saved = True
        print("    保存する [OK]")
        time.sleep(2)
    else:
        # fallback: ダイアログ内に限定したlocator
        for sel in [
            '.ant-drawer-body button:has-text("保存する")',
            '.ant-modal-content button:has-text("保存する")',
            '[role="dialog"] button:has-text("保存する")',
            'button:has-text("保存する")',
        ]:
            try:
                page.locator(sel).first.click(timeout=3000)
                saved = True
                print("    保存OK (" + sel + ")")
                break
            except Exception:
                pass
    if not saved:
        print("    警告: 保存ボタンが見つかりません")
        _save_screenshot(page, "cpass_action_ss.png")
    time.sleep(2)

    # ★2026/07/06追加: 保存後に外側の編集ダイアログが残っていないか確認・強制クローズ
    # → 残ったままだと次の注文で「もう開いている」と誤認識され以降の注文が軒並み失敗するため
    _close_all_dialogs(page)

    return saved, dhl_price


def _move_single_order_to_processing(page, order_no):
    """指定注文を「発送手続き」へ移動（アクション → 発送手続き）"""
    for sel in [
        '.ant-dropdown:not(.ant-dropdown-hidden) .ant-dropdown-menu-item:has-text("発送手続き")',
        '.ant-dropdown-menu-item:has-text("発送手続き")',
        'li[role="menuitem"]:has-text("発送手続き")',
    ]:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=3000)
            loc.click(timeout=3000)
            print("  → 発送手続きへ移動クリック OK")
            time.sleep(2)
            # 確認ダイアログ（★2026/07/14修正: ボタンテキストは「確 認」で空白入り・
            # クラスもant-btn-primaryではないため、空白除去テキスト一致のJSクリックを使う）
            try:
                clicked = page.evaluate(
                    """() => {
                        const btns = Array.from(document.querySelectorAll('button'));
                        for (const b of btns) {
                            const t = (b.textContent || '').replace(/\\s+/g, '');
                            if (t === '確認' || t === 'OK') { b.click(); return true; }
                        }
                        const modal = document.querySelector(
                            '.sp-modal-dialog .prompt-modal-body, .prompt-modal-body'
                        );
                        if (modal) {
                            const cand = Array.from(modal.querySelectorAll('button')).find(
                                b => (b.textContent || '').replace(/\\s+/g, '') !== '閉じる'
                            );
                            if (cand) { cand.click(); return true; }
                        }
                        return false;
                    }"""
                )
                if clicked:
                    print("  → 確認ダイアログ OK")
                    time.sleep(2)
            except Exception:
                pass
            return True
        except Exception:
            continue
    print("  → 発送手続きメニュー見つからず")
    return False


def process_all_orders_for_dhl(target_order_nos=None, headless=False, move_waiting=True):
    """「発送手続き待ち」で各注文を編集→DHL取得→発送手続きへ移動

    修正版: 編集ダイアログは「発送手続き待ち」タブにのみ存在する。
    「発送手続き」タブには「編集」メニューがないため、
    1注文ずつ「待ち」で編集してからその注文を「発送手続き」へ移動する。

    Args:
        target_order_nos: 処理対象の注文番号リスト（None なら全件）
        headless: ブラウザ非表示モード
        move_waiting: True なら編集後に発送手続きへ移動する（False=編集のみ）

    Returns:
        dict: {order_no: {package_no, dhl_price_jpy, title, item_id}}
    """
    from playwright.sync_api import sync_playwright

    results = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=[
                "--window-position=0,0",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        try:
            context = browser.new_context(
                user_agent=USER_AGENT,
                viewport={"width": 1400, "height": 900},
                locale="ja-JP",
            )
            context.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            page = context.new_page()

            _login(page)

            print("エントリURLへ: " + CPASS_ENTRY_URL)
            page.goto(CPASS_ENTRY_URL, wait_until="domcontentloaded", timeout=60000)
            time.sleep(3)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass

            _dismiss_announcement_modal(page)

            # ★正しいフロー（2026/05/24 確認）:
            # Step A: 発送手続き待ち を全件 発送手続き へ一括移動
            # Step B: 発送手続き タブで「詳細を見る」→入力→配送を割り当て→DHL取得→保存
            if move_waiting:
                _move_all_to_processing(page)

            # 発送手続きタブへ移動して注文一覧を取得
            print("発送手続き 注文一覧取得...")
            _navigate_to_sidebar_tab(page, "発送手続き")
            time.sleep(3)
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass

            _dismiss_announcement_modal(page)

            # デバッグ: ページHTMLを保存
            try:
                debug_html = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                          "cpass_processing_dump.html")
                with open(debug_html, "w", encoding="utf-8") as f:
                    f.write(page.content())
                print("  HTML保存: " + debug_html)
            except Exception:
                pass

            # ★2026/07/09修正: 1ページ目だけでなく全ページを巡回して収集する
            orders = _scrape_all_orders_with_pagination(page)
            print("  発送手続き 件数(全ページ合計): " + str(len(orders)))

            # 対象フィルタ
            if target_order_nos is not None:
                target_set = set(target_order_nos)
                # ★2026/07/06追加: 対象注文のうちCPaSS「発送手続き」タブに
                # そもそも存在しないものを明示的にログ出力（原因切り分け用）
                found_set = set(o["order_no"] for o in orders)
                missing = target_set - found_set
                if missing:
                    print("  [WARN] 対象注文のうちCPaSS発送手続きタブで見つからないもの("
                          + str(len(missing)) + "件): " + ", ".join(sorted(missing)))
                orders = [o for o in orders if o["order_no"] in target_set]
                print("  対象絞り込み後: " + str(len(orders)) + " 件")

            # 各注文を処理（待ちタブで編集 → 発送手続きへ移動）
            for idx, order in enumerate(orders):
                print()
                print("--- [" + str(idx + 1) + "/" + str(len(orders)) +
                      "] order=" + order["order_no"] + " pkg=" + order["package_no"] + " ---")
                print("  Title: " + order.get("title", "")[:60])

                # 寸法・重量・HS推定
                # ★2026/08/04確定: 重量はeBay実データ(Shipping Policy名)を最優先で使う。
                #   タイトルキーワード推定は寸法(長さ幅高さ)にのみ使い、重量はフォールバック時のみ使用する。
                dims = dimension_weight_lookup.lookup_dimensions_weight(order.get("title", ""))
                hs_code = hs_code_lookup.lookup_hs_code(order.get("title", ""))

                ebay_weight_kg, weight_source, weight_err = get_ebay_shipping_weight_kg(
                    order.get("item_id", ""))
                if ebay_weight_kg:
                    dims["weight_kg"] = ebay_weight_kg
                    print("  推定: " + dims["category"] + " / HS=" + hs_code +
                          " / 重量=" + str(ebay_weight_kg) + "kg [eBay実データ: " + weight_source + "]")
                else:
                    weight_source = "keyword_fallback"
                    print("  推定: " + dims["category"] + " / HS=" + hs_code +
                          " / 重量=" + str(dims["weight_kg"]) + "kg " +
                          "[★eBay実データ取得失敗→タイトル推定にフォールバック: " + str(weight_err) + "]")

                # ★ 前の注文の残存ダイアログ（外側編集ダイアログ含む）を確実に閉じる
                # ★2026/07/06: ここが不十分だと次注文の編集ボタンクリック後に
                #   前注文のダイアログをそのまま「開けた」と誤判定してしまう
                _close_all_dialogs(page)

                # ★「発送手続き」タブで「詳細を見る」からダイアログを開く
                if not _open_edit_dialog(page, order["order_no"]):
                    print("  編集ダイアログ開けず → スキップ")
                    continue

                # フォーム入力 → 配送を割り当て → DHL価格取得 → 保存する
                # ★正しい順序: 配送割り当てでDHL価格確認後に保存
                _saved, dhl_price = _fill_edit_form_and_save(
                    page,
                    weight_kg=dims["weight_kg"],
                    length_cm=dims["length_cm"],
                    width_cm=dims["width_cm"],
                    height_cm=dims["height_cm"],
                    hs_code=hs_code,
                )

                results[order["order_no"]] = {
                    "package_no": order["package_no"],
                    "dhl_price_jpy": dhl_price,
                    "title": order.get("title", ""),
                    "item_id": order.get("item_id", ""),
                    "weight_kg": dims["weight_kg"],
                    "weight_source": weight_source,
                    "dims": [dims["length_cm"], dims["width_cm"], dims["height_cm"]],
                    "hs_code": hs_code,
                }

                # 発送手続きタブで処理済みのため移動不要

                time.sleep(2)

        finally:
            browser.close()

    return results


if __name__ == "__main__":
    # テスト: 最初の1件だけ処理
    target = None
    test_one = False
    for arg in sys.argv[1:]:
        if arg == "--one":
            test_one = True
        elif arg.startswith("--order="):
            target = [arg.split("=", 1)[1]]
        elif arg == "--no-move":
            pass

    print("=" * 60)
    print("CPaSS DHL価格取得ワークフロー テスト")
    print("=" * 60)
    print()

    move_waiting = "--no-move" not in sys.argv

    results = process_all_orders_for_dhl(
        target_order_nos=target,
        headless=False,
        move_waiting=move_waiting,
    )

    print()
    print("=" * 60)
    print("結果: " + str(len(results)) + " 件")
    print("=" * 60)
    for order_no, info in results.items():
        print("  " + order_no + " → " +
              ("¥" + str(info["dhl_price_jpy"]) if info["dhl_price_jpy"] else "取得失敗") +
              " (" + info.get("title", "")[:40] + ")")

    # JSON保存
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "cpass_dhl_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print()
    print("保存: " + out_path)
