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

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import cpass_config
import hs_code_lookup
import dimension_weight_lookup

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# CPaSS タブ別URL（HTMLダンプで確認済み）
CPASS_ENTRY_URL = "https://cpass.ebay.com/order/paid"
CPASS_TAB_URLS = {
    "発送手続き待ち": "https://cpass.ebay.com/order/paid",
    "発送手続き":     "https://cpass.ebay.com/order/readytoship",
    "キャンセル":     "https://cpass.ebay.com/order/cancelled",
    "出荷待ち":       "https://cpass.ebay.com/order/labelprinted",
    "出荷":           "https://cpass.ebay.com/order/intransit",
}


def _login(page):
    """CPaSS にログイン"""
    print("CPaSS ログイン中...")
    page.goto(cpass_config.CPASS_LOGIN_URL, wait_until="domcontentloaded", timeout=45000)
    time.sleep(3)
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass

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
    print("  確認ダイアログを待機...")
    dialog_closed = False
    for attempt in range(12):
        try:
            # Ant Design Modal の確認ボタン（プライマリボタン）
            for sel in [
                'button.ant-btn-primary',
                '.ant-modal-confirm-btns button.ant-btn-primary',
                '.ant-modal-footer button.ant-btn-primary',
                'button:has-text("確認")',
                'button:has-text("OK")',
            ]:
                try:
                    btn = page.locator(sel).first
                    if btn.is_visible(timeout=500):
                        btn.click(timeout=2000)
                        print("    [OK] 確認ダイアログ → クリック (" + sel + ")")
                        dialog_closed = True
                        break
                except Exception:
                    pass
            if dialog_closed:
                break
        except Exception:
            pass
        time.sleep(1)
    if not dialog_closed:
        print("    確認ダイアログなし or 閉じ済み")

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


def _open_detail_dialog(page, order_no):
    """指定の order_no の「編集」ボタンをクリックしてダイアログを開く

    ★2026/06/20 確認: CPaSS「発送手続き」タブのボタンは「編集」。
    """
    # 著作権ダイアログを先に閉じる
    _dismiss_copyright_dialog(page)

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


def _fill_edit_form_and_save(page, weight_kg, length_cm, width_cm, height_cm, hs_code):
    """フォーム入力 → 配送を割り当て(DHL価格取得) → 保存する

    ★正しい順序（CLAUDE.md 2026/05/24 確認）:
    1. フォーム入力（重量・寸法・HSコード）
    2. 「配送を割り当て」クリック → DHL金額表示
    3. 「保存する」クリック

    Args:
        weight_kg: 梱包重量(kg)
        length_cm, width_cm, height_cm: サイズ
        hs_code: HSコード（10桁）

    Returns:
        tuple: (saved_ok: bool, dhl_price_jpy: int|None)
    """
    print("  フォーム入力: 重量=" + str(weight_kg) +
          "kg, " + str(length_cm) + "x" + str(width_cm) + "x" + str(height_cm) +
          "cm, HS=" + hs_code)

    time.sleep(1)

    # 1. JS で input フィールドにラベルベースで入力
    filled = page.evaluate(
        """(args) => {
            // ヘルパー: label/前のテキストから input を見つける
            function findInputNear(labelText) {
                const labels = Array.from(document.querySelectorAll('*'))
                    .filter(el => {
                        const t = (el.textContent || '').trim();
                        return t.startsWith(labelText) && el.children.length === 0;
                    });
                for (const lbl of labels) {
                    let parent = lbl;
                    for (let i = 0; i < 5; i++) {
                        parent = parent.parentElement;
                        if (!parent) break;
                        const inp = parent.querySelector('input[type="text"], input[type="number"]');
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
                const nativeSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                nativeSetter.call(input, String(value));
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            }
            const results = {};

            // 梱包重量（配送概要セクションの「梱包」入力欄）
            // ※「単位重量」はeBayから引き継がれるので触らない
            const wInput = findInputNear('梱包');
            results.weight = setValue(wInput, args.weight);

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

    # ★ 残存モーダルを閉じる（前の注文の内側モーダルが残っている場合）
    page.evaluate(
        """() => {
            const btns = Array.from(document.querySelectorAll('button'));
            // 表示中の「閉じる」ボタンをすべてクリック
            btns.filter(b => (b.textContent||'').trim() === '閉じる' && b.offsetParent !== null)
                .forEach(b => b.click());
        }"""
    )
    time.sleep(1)

    # 2. 「配送を割り当て」クリック → 内側モーダルが開く → DHL「選択」→価格取得
    print("  「配送を割り当て」クリック...")
    dhl_price = None
    clicked_assign = False
    for sel in [
        'button:has-text("配送を割り当て")',
        'a:has-text("配送を割り当て")',
        'button:has-text("Assign Shipping")',
    ]:
        try:
            page.locator(sel).first.click(timeout=3000)
            clicked_assign = True
            print("    [OK] " + sel)
            break
        except Exception:
            pass

    if clicked_assign:
        time.sleep(3)

        # 内側モーダル内のDHL価格を取得
        dhl_price = page.evaluate(
            """() => {
                const all = Array.from(document.querySelectorAll('*'));
                const dhlNodes = all.filter(el => {
                    const t = (el.textContent || '');
                    return t.toLowerCase().includes('dhl') && el.children.length === 0;
                });
                for (const node of dhlNodes) {
                    let card = node;
                    for (let i = 0; i < 8; i++) {
                        card = card.parentElement;
                        if (!card) break;
                        const txt = card.textContent || '';
                        const m = txt.match(/([\\d,]+)\\s*[-〜~–]\\s*([\\d,]+)\\s*JPY/);
                        if (m) return parseInt(m[2].replace(/,/g, ''), 10);
                        const m2 = txt.match(/([\\d,]+)\\s*JPY/);
                        if (m2 && txt.toLowerCase().includes('dhl')) {
                            return parseInt(m2[1].replace(/,/g, ''), 10);
                        }
                    }
                }
                return null;
            }"""
        )
        if dhl_price:
            print("    DHL上限: " + str(dhl_price) + " JPY")
        else:
            print("    DHL価格取得失敗（価格なしで保存を続行）")

        # 内側モーダル（.ant-modal内）の「閉じる」ボタンを閉じる
        # ★外側ダイアログの「閉じる」と区別するため .ant-modal 内のみを対象にする
        print("  内側モーダルを閉じる...")
        closed = False
        rect2 = page.evaluate(
            """() => {
                // ant-modal 内の閉じるボタンを探す
                const modal = document.querySelector('.ant-modal');
                if (modal) {
                    const btns = Array.from(modal.querySelectorAll('button'));
                    const btn = [...btns].reverse().find(b => (b.textContent||'').trim() === '閉じる' && b.offsetParent !== null);
                    if (btn) {
                        const r = btn.getBoundingClientRect();
                        return {x: r.left + r.width / 2, y: r.top + r.height / 2};
                    }
                    // ×ボタン
                    const closeBtn = modal.querySelector('.ant-modal-close');
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
        if not closed:
            page.keyboard.press("Escape")
            time.sleep(1.5)
    else:
        print("    警告: 「配送を割り当て」ボタンが見つかりません")
        _save_screenshot(page, "cpass_action_ss.png")

    # 3. 「保存する」ボタンを座標クリック
    time.sleep(1)
    saved = False
    rect_save = page.evaluate(
        """() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const btn = btns.find(b => (b.textContent||'').trim() === '保存する' && b.offsetParent !== null);
            if (!btn) return null;
            const r = btn.getBoundingClientRect();
            return {x: r.left + r.width / 2, y: r.top + r.height / 2};
        }"""
    )
    if rect_save and rect_save.get('x'):
        page.mouse.click(rect_save['x'], rect_save['y'])
        saved = True
        print("    保存する [OK]")
        time.sleep(2)
    else:
        # fallback: locator
        for sel in ['button:has-text("保存する")', 'button:has-text("保存")']:
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
            # 確認ダイアログ
            for dlg_sel in ['button.ant-btn-primary', 'button:has-text("確認")', 'button:has-text("OK")']:
                try:
                    btn = page.locator(dlg_sel).first
                    if btn.is_visible(timeout=1500):
                        btn.click(timeout=2000)
                        print("  → 確認ダイアログ OK")
                        time.sleep(2)
                        break
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

            # デバッグ: ページHTMLを保存
            try:
                debug_html = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                          "cpass_processing_dump.html")
                with open(debug_html, "w", encoding="utf-8") as f:
                    f.write(page.content())
                print("  HTML保存: " + debug_html)
            except Exception:
                pass

            orders = _scrape_orders_from_page(page)
            print("  発送手続き 件数: " + str(len(orders)))

            # 対象フィルタ
            if target_order_nos is not None:
                target_set = set(target_order_nos)
                orders = [o for o in orders if o["order_no"] in target_set]
                print("  対象絞り込み後: " + str(len(orders)) + " 件")

            # 各注文を処理（待ちタブで編集 → 発送手続きへ移動）
            for idx, order in enumerate(orders):
                print()
                print("--- [" + str(idx + 1) + "/" + str(len(orders)) +
                      "] order=" + order["order_no"] + " pkg=" + order["package_no"] + " ---")
                print("  Title: " + order.get("title", "")[:60])

                # 寸法・重量・HS推定
                dims = dimension_weight_lookup.lookup_dimensions_weight(order.get("title", ""))
                hs_code = hs_code_lookup.lookup_hs_code(order.get("title", ""))
                print("  推定: " + dims["category"] + " / HS=" + hs_code)

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
