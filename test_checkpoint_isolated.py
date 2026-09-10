"""★2026/09/10追加: チェックポイント機構の隔離検証専用スクリプト。

【安全設計】
- アクセスするGoogle Driveフォルダ・ファイルはテスト専用のみ:
    フォルダ = "売上管理表_チェックポイントテスト"（本番の「売上管理表」フォルダとは別物・別ID）
    ファイル = "通常_チェックポイントテスト_20260910.xlsm"（本番ファイルの複製から作成したテスト専用コピー）
  本番フォルダ名 "売上管理表" は本スクリプト内のどこにも登場しない。
- CPaSS・Playwrightは一切呼び出さない。cpass_workflowモジュールをダミー(空)としてsys.modulesに
  登録した上で daily_workflow_ga.py をimportすることで、cpass_workflow の実体（Playwright起動処理・
  CPaSSログイン処理）が一切ロードされない状態にしている。
- 送料は実際にCPaSSから取得せず、明確にテストとわかるダミー値（99991〜99993円）を使う。
- 使う関数は daily_workflow_ga.py の実物（process_xlsm / get_target_orders / _download_xlsm /
  _upload_xlsm / _get_drive_service / _find_folder_id）をそのままimportして使う。
  → 本番コードと全く同じロジックで「1件ごとのExcel書き込み→Drive途中保存→次回起動時の
    スキップ/再開」を検証できる。

TEST_ORDER_LIMIT環境変数で、今回処理する対象注文の件数を制限できる（例: "1"を指定すると
最初の1件だけ処理して終了する。これにより「1件保存後に途中終了→次回起動で残りから再開」の
シナリオをテストできる）。
"""
import sys
import os
import time
import types
import tempfile

# cpass_workflowをダミーモジュールとして登録（Playwright/CPaSSを一切ロードしない）
sys.modules['cpass_workflow'] = types.ModuleType('cpass_workflow')

import daily_workflow_ga as ga  # 本番と全く同じ関数を利用

TEST_FOLDER_NAME = "売上管理表_チェックポイントテスト"
TEST_FILE_NAME = "通常_チェックポイントテスト_20260910.xlsm"

_order_limit_env = os.environ.get("TEST_ORDER_LIMIT", "").strip()
ORDER_LIMIT = int(_order_limit_env) if _order_limit_env else None

# ダミー送料・ダミー梱包番号（本番のDHL実料金と混同しないよう、明確に区別できる値にしている）
FAKE_PRICES = {
    "07-15147-67281": 99991,
    "01-15163-10197": 99992,
    "05-15156-05008": 99993,
}
FAKE_PACKAGE_NOS = {
    "07-15147-67281": 3811,
    "01-15163-10197": 3815,
    "05-15156-05008": 3817,
}


def main():
    print("=" * 60)
    print(f"[TEST] チェックポイント隔離検証開始  {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[TEST] 対象Driveフォルダ: 「{TEST_FOLDER_NAME}」（本番「売上管理表」フォルダとは別物）")
    print(f"[TEST] 対象ファイル: {TEST_FILE_NAME}")
    print(f"[TEST] CPaSS/Playwrightは一切呼び出しません（cpass_workflowはダミーモジュール）")
    print(f"[TEST] TEST_ORDER_LIMIT={ORDER_LIMIT}")
    print("=" * 60)

    # ★診断用: サービスアカウントのメールアドレスを表示（秘密鍵は表示しない）。
    # テスト用フォルダがサービスアカウントに共有されていない可能性の切り分けのため。
    try:
        _sa_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
        print(f"[TEST][DEBUG] サービスアカウントのメールアドレス: {_sa_info.get('client_email')}")
    except Exception as _e:
        print(f"[TEST][DEBUG] サービスアカウント情報の読み取りに失敗: {_e}")

    service = ga._get_drive_service()

    # ★診断用: サービスアカウントが「見える」フォルダを一覧表示（本番「売上管理表」フォルダが
    # そもそも見えているか、テスト用フォルダとの違いを確認するため）。
    try:
        _folders = service.files().list(
            q="mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id,name,owners(emailAddress))",
            pageSize=50,
        ).execute().get("files", [])
        print(f"[TEST][DEBUG] サービスアカウントから見えるフォルダ一覧({len(_folders)}件):")
        for _f in _folders:
            _owners = ",".join(o.get("emailAddress", "?") for o in _f.get("owners", []))
            print(f"    - {_f['name']} (id={_f['id']}, owner={_owners})")
    except Exception as _e:
        print(f"[TEST][DEBUG] フォルダ一覧取得に失敗: {_e}")

    folder_id = ga._find_folder_id(service, TEST_FOLDER_NAME)
    print(f"[TEST] テストフォルダID: {folder_id}")

    resp = service.files().list(
        q=f"'{folder_id}' in parents and name='{TEST_FILE_NAME}' and trashed=false",
        fields="files(id,name,modifiedTime)",
    ).execute()
    files = resp.get("files", [])
    if not files:
        print(f"[TEST][ERROR] テストファイルが見つかりません: {TEST_FILE_NAME}")
        sys.exit(1)
    file_id = files[0]["id"]
    print(f"[TEST] テストファイルID: {file_id} (modifiedTime={files[0]['modifiedTime']})")

    with tempfile.TemporaryDirectory() as workdir:
        local_path = os.path.join(workdir, TEST_FILE_NAME)
        ga._download_xlsm(service, file_id, local_path)
        print(f"[TEST] DL完了: {local_path}")

        target_order_nos = ga.get_target_orders(local_path)
        ordered_targets = sorted(target_order_nos)
        print(f"[TEST] 対象注文(送料空白, get_target_ordersの実結果): {ordered_targets}")

        if ORDER_LIMIT is not None:
            skipped_this_run = ordered_targets[ORDER_LIMIT:]
            ordered_targets = ordered_targets[:ORDER_LIMIT]
            print(f"[TEST] TEST_ORDER_LIMIT={ORDER_LIMIT} のため今回処理: {ordered_targets}")
            print(f"[TEST] 今回は処理しない(次回起動で再開されるはず): {skipped_this_run}")

        cpass_results = {}
        checkpoint_count = 0

        for idx, order_no in enumerate(ordered_targets, start=1):
            print()
            print(f"[TEST] {idx}/{len(ordered_targets)}件目 処理開始（ダミー送料取得、CPaSS未呼び出し）: {order_no}")
            time.sleep(2)  # 実際の1件あたり処理時間を模した待機
            info = {
                "package_no": FAKE_PACKAGE_NOS.get(order_no),
                "dhl_price_jpy": FAKE_PRICES.get(order_no, 99999),
                "title": "TEST-DUMMY",
                "item_id": "TEST-DUMMY",
            }
            cpass_results[order_no] = info
            print(f"  [CHECKPOINT] {idx}/{len(ordered_targets)}件目まで処理済み: 対象注文={order_no} "
                  f"送料(ダミー)={info['dhl_price_jpy']}")

            ok, num_writes = ga.process_xlsm(local_path, cpass_results, dry_run=False)
            if ok and num_writes > 0:
                ga._upload_xlsm(service, file_id, local_path)
                checkpoint_count += 1
                print(f"  [CHECKPOINT] {idx}/{len(ordered_targets)}件目時点: {num_writes}セル保存 "
                      f"対象注文={order_no} → Google Drive(テスト専用フォルダ)アップロード成功")
            elif ok and num_writes == 0:
                print(f"  [CHECKPOINT] {idx}/{len(ordered_targets)}件目時点: 新規書き込みなし、ULスキップ")
            else:
                print(f"  [CHECKPOINT][WARN] {idx}/{len(ordered_targets)}件目: process_xlsm失敗")

        print()
        print("=" * 60)
        print(f"[TEST] 完了  {time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"[TEST] 今回チェックポイント保存回数: {checkpoint_count}")
        print(f"[TEST] 今回処理した注文: {ordered_targets}")
        print("=" * 60)


if __name__ == "__main__":
    main()
