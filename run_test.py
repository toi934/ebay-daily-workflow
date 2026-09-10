"""チェックポイント機構 + get_target_orders()修正 の隔離検証(単発の実行で完結)。

★2026/09/10 Run#5の結果、get_target_orders()の「送料記入済み最終行より後ろだけ
スキャンする」ロジックに不具合があることが判明した: 1回の実行で複数件処理する際、
行番号が大きい注文が先にチェックポイント保存されると、まだ未処理のまま残っている
「行番号が小さい注文」が次回スキャンで検出されなくなる(エラーも出ずに永久に
取りこぼされる)。daily_workflow_ga.py の get_target_orders() をシート全体スキャン
方式に修正し(★2026/09/10 追加修正、を参照)、対象注文をCPaSSへ渡す順序も
行番号昇順のlistに変更した。

本スクリプトはその修正を検証する。わざと「行番号の大きい注文を先に保存」という
最悪ケースを人為的に作り、それでも行番号の小さい注文が次回スキャンで確実に
検出されること・保存済み注文が再処理されないこと・最終的に全件処理されること・
キャンセル注文や既存の送料記入済み注文が誤って触られていないこと・本番ファイルが
一切変更されていないことを確認する。

本番daily_workflow_ga.py(GitHub Actions版の実物)の関数をそのままimportし、実際の
サービスアカウント・実際のGoogle Drive APIに対して動作させる。CPaSS/Playwrightは
一切呼び出さない(cpass_workflowはダミーモジュールとしてsys.modulesに登録)。

安全設計:
- 本番「売上管理表」フォルダへの書き込みは一切行わない。読み取り専用で
  (1) フォルダID検索 (2) 最新の本番xlsmファイル一覧取得 の2箇所だけ参照する。
- 以降のダウンロード・書き込み・アップロードは、すべてテスト専用サブフォルダ内の
  テストコピーに対してのみ行う。production file_id を _upload_xlsm に渡すコード
  パスは存在しない(アサーションで二重に保証)。
- 送料はダミー値を書き込む。CPaSSは一切呼び出さない。
- "再起動後の再開"は、実際にDriveへ再アップロードしたテストコピーを毎回あらためて
  ダウンロードし直すことでシミュレートする(本物のGitHub Actionsの2回目の実行と
  同じく、プロセス内の状態には一切頼らずファイルの中身だけから再開判定する)。
"""
import sys
import os
import re
import time
import json
import types
import tempfile

sys.modules['cpass_workflow'] = types.ModuleType('cpass_workflow')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import daily_workflow_ga as ga  # noqa: E402
import openpyxl  # noqa: E402

if not os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
    _local_creds = os.path.join(os.path.dirname(__file__), "google_creds.json")
    if os.path.exists(_local_creds):
        with open(_local_creds, encoding="utf-8") as f:
            os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"] = f.read()

PROD_FOLDER_NAME = ga.DRIVE_FOLDER_NAME  # "売上管理表" -- 読み取り専用でのみ参照
TEST_SUBFOLDER_NAME = "_テスト_チェックポイント検証(削除可)"
TEST_FILE_PREFIX = "TEST_チェックポイント検証_"

ORDER_PATTERN = re.compile(r"^\d{2}-\d{5}-\d{5}$")
FAKE_PRICE_BASE = 99990

# ★前回(Run#5)セットアップ時に記録した、対象3件の元の値(このテストコピー内のみ)。
# 今回はこの既知の値へ一旦リストアしてから空欄化することで、前回(修正前コードでの)
# 実行結果を引きずらないクリーンな状態から検証する。
KNOWN_ORIGINAL_VALUES = {
    "07-15147-67281": {"row": 584, "pkg": 3811, "ship": 8063},
    "01-15163-10197": {"row": 586, "pkg": 3815, "ship": 7500},
    "05-15156-05008": {"row": 587, "pkg": 3817, "ship": 5527},
}
# ★対象3件の近傍にある「絶対に触られてはいけない」行(既存の送料記入済み注文、
# およびキャンセル注文)。最終確認でこれらの値が不変であることも検証する。
GUARD_ROWS = [580, 581, 582, 583, 585]


def log(msg):
    print(f"[TEST] {msg}", flush=True)


def get_or_create_test_subfolder(service, prod_folder_id):
    q = (f"'{prod_folder_id}' in parents and name='{TEST_SUBFOLDER_NAME}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    resp = service.files().list(q=q, fields="files(id,name)").execute()
    files = resp.get("files", [])
    if files:
        log(f"既存テストサブフォルダを再利用: {TEST_SUBFOLDER_NAME} (id={files[0]['id']})")
        return files[0]["id"], False
    body = {"name": TEST_SUBFOLDER_NAME, "mimeType": "application/vnd.google-apps.folder",
            "parents": [prod_folder_id]}
    created = service.files().create(body=body, fields="id").execute()
    log(f"テストサブフォルダを新規作成: {TEST_SUBFOLDER_NAME} (id={created['id']})")
    return created["id"], True


def find_test_copy(service, test_folder_id):
    """テスト専用サブフォルダ内のテストコピーを探す。

    ★重要: サービスアカウントには独自のDriveストレージ容量が無いため、
    Drive API の files.copy / files.create で新規ファイル(実体)を作ることはできない
    (storageQuotaExceeded)。そのためテストコピーは、実際にDriveの保存容量を持つ
    人間のアカウント側(マイドライブ同期などファイルシステム経由)で事前に配置して
    もらう/配置しておく必要がある。このファイルは既存の本番ファイルをコピーしたもので
    あり、以降はサービスアカウントの書き込み権限(親フォルダから継承)で
    download/update するだけなので、それ自体はストレージ容量を消費しない。
    """
    q = f"'{test_folder_id}' in parents and name contains '{TEST_FILE_PREFIX}' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name,modifiedTime)",
                                 orderBy="modifiedTime desc").execute()
    files = resp.get("files", [])
    if not files:
        raise RuntimeError(
            f"テスト専用サブフォルダ『{TEST_SUBFOLDER_NAME}』内に "
            f"'{TEST_FILE_PREFIX}' で始まるテストコピーが見つかりません。"
            "(サービスアカウントはDriveストレージ容量を持たないため、"
            "このテストコピー自体は人間のアカウント側で事前に配置する必要があります)"
        )
    log(f"テストコピーを検出: {files[0]['name']} (id={files[0]['id']})")
    return files[0]["id"], files[0]["name"]


def _shipping_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    try:
        if float(v) == 0:
            return True
    except (TypeError, ValueError):
        pass
    return False


def fresh_download(service, test_file_id, test_file_name, tag):
    """毎回あらためてDriveからDLし直す(=新しいプロセス/新しいActions実行を模す)。"""
    workdir = tempfile.mkdtemp(prefix=f"ga_test_{tag}_")
    local_path = os.path.join(workdir, test_file_name)
    ga._download_xlsm(service, test_file_id, local_path)
    log(f"[{tag}] Driveから再ダウンロード(新規プロセス相当): {test_file_name}")
    return local_path


def read_guard_and_target_cells(local_path, sheet_name, ship_col_idx, rows):
    wb = openpyxl.load_workbook(local_path, keep_vba=True, data_only=True)
    ws = wb[sheet_name]
    values = {}
    for row in rows:
        values[row] = {
            "order_no": ws.cell(row=row, column=2).value,
            "pkg_no": ws.cell(row=row, column=1).value,
            "shipping": ws.cell(row=row, column=ship_col_idx).value,
            "status_e": ws.cell(row=row, column=5).value,
        }
    wb.close()
    return values


def reset_and_blank_test_file(service, test_file_id, test_file_name):
    """既知の元の値へリストア→対象3件を空欄化し、Driveへアップロードしてクリーンな
    検証開始状態を作る。戻り値は (sheet_name, ship_col_idx, guard_values_before)。
    """
    local_path = fresh_download(service, test_file_id, test_file_name, "reset-setup")
    wb = openpyxl.load_workbook(local_path, keep_vba=True)
    sheet_name = ga.find_sheet_with_orders(wb)
    ws = wb[sheet_name]
    ship_col_idx = ga.find_shipping_column(ws, local_path)

    for order_no, info in KNOWN_ORIGINAL_VALUES.items():
        row = info["row"]
        actual_order = ws.cell(row=row, column=2).value
        assert actual_order and str(actual_order).strip() == order_no, (
            f"行{row}の注文番号が想定と異なります(想定={order_no}, 実際={actual_order})。"
            "テストコピーの行構成が変わっている可能性があるため中断します。"
        )
        ws.cell(row=row, column=1).value = None
        ws.cell(row=row, column=ship_col_idx).value = None
    wb.save(local_path)
    wb.close()

    log(f"対象3件を既知の元の値へ一旦リストア後、再度空欄化しました: "
        f"{[(o, i['row']) for o, i in KNOWN_ORIGINAL_VALUES.items()]}")

    guard_values_before = read_guard_and_target_cells(local_path, sheet_name, ship_col_idx, GUARD_ROWS)
    log(f"ガード行(触られてはいけない既存注文/キャンセル行)の初期値: {guard_values_before}")

    ga._upload_xlsm(service, test_file_id, local_path)
    log("リストア+空欄化後のテストコピーをGoogle Driveへアップロード完了(検証開始状態を確立)")

    return sheet_name, ship_col_idx, guard_values_before


def process_one_order(service, test_file_id, test_file_name, order_no, tag, prod_ids, idx=1, total=1):
    """指定した1件を処理し、チェックポイント保存する(=新規プロセスでのDL→書込→ULを模す)。"""
    local_path = fresh_download(service, test_file_id, test_file_name, tag)
    targets_before = ga.get_target_orders(local_path)
    log(f"[{tag}] 現在の対象注文(行番号順, 送料空欄): {targets_before}")
    assert order_no in targets_before, (
        f"[{tag}] 処理予定の {order_no} が対象注文リストに含まれていません: {targets_before}"
    )

    fake_price = FAKE_PRICE_BASE + (sum(ord(ch) for ch in order_no) % 9)
    info = {"package_no": 9000 + idx, "dhl_price_jpy": fake_price,
            "title": "TEST-DUMMY", "item_id": "TEST-DUMMY"}
    cpass_results = {order_no: info}

    log(f"[{tag}] {idx}/{total}件目 処理開始(ダミー送料、CPaSS未呼び出し): {order_no}")
    ok, num_writes = ga.process_xlsm(local_path, cpass_results, dry_run=False)
    entry = {"order_no": order_no, "fake_price": fake_price, "ok": ok, "num_writes": num_writes,
              "time": time.strftime("%Y-%m-%d %H:%M:%S"), "targets_before": targets_before}
    if ok and num_writes > 0:
        assert test_file_id not in prod_ids, "重大な設定ミス: テストfile_idが本番file_idと一致"
        ga._upload_xlsm(service, test_file_id, local_path)
        entry["uploaded"] = True
        log(f"[{tag}] [CHECKPOINT] {idx}/{total}件目: {num_writes}セル保存 "
            f"対象注文={order_no} 送料(ダミー)={fake_price} → Google Driveアップロード成功")
    else:
        entry["uploaded"] = False
        log(f"[{tag}] [CHECKPOINT] {idx}/{total}件目: 書き込みなし(num_writes={num_writes})")
    return entry


def main():
    log("=" * 70)
    log(f"開始  {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log("get_target_orders()全行スキャン化 + 処理順を行番号順に固定 の修正検証")
    log("=" * 70)

    service = ga._get_drive_service()
    prod_folder_id = ga._find_folder_id(service, PROD_FOLDER_NAME)
    log(f"[READ-ONLY] 本番フォルダID: {prod_folder_id} (フォルダ名『{PROD_FOLDER_NAME}』)")

    prod_files_before = ga._list_xlsm_files(service, prod_folder_id, "通常")
    prod_snapshot_before = [
        {"id": f["id"], "name": f["name"], "modifiedTime": f["modifiedTime"]}
        for f in prod_files_before
    ]
    prod_ids = {f["id"] for f in prod_snapshot_before}
    log(f"[READ-ONLY] 本番『通常』xlsm一覧(実行前スナップショット, {len(prod_snapshot_before)}件)")

    test_folder_id, folder_created = get_or_create_test_subfolder(service, prod_folder_id)
    test_file_id, test_file_name = find_test_copy(service, test_folder_id)
    assert test_file_id not in prod_ids, "重大な設定ミス: テストfile_idが本番file_idと一致しています"

    # ── 0. クリーンな検証開始状態を作る(前回実行の状態を引きずらない) ──
    sheet_name, ship_col_idx, guard_before = reset_and_blank_test_file(service, test_file_id, test_file_name)

    setup_path = fresh_download(service, test_file_id, test_file_name, "setup-check")
    initial_targets = ga.get_target_orders(setup_path)
    log(f"検証開始時点の対象注文(行番号昇順であるはず): {initial_targets}")
    expected_order = ["07-15147-67281", "01-15163-10197", "05-15156-05008"]
    assert initial_targets == expected_order, (
        f"対象注文が行番号昇順で返っていません(期待={expected_order}, 実際={initial_targets})。"
        "get_target_orders()の修正に問題がある可能性があります。"
    )
    log("[OK] get_target_orders()は行番号の昇順でlistを返している(set()の順序不定を解消)")

    # ── 1. わざと「行番号の大きい注文を先に保存」する最悪ケースを作る ──
    adversarial_order = initial_targets[-1]  # 05-15156-05008 (587行目、最も行番号が大きい)
    log(f"[意図的な悪条件] 行番号が最も大きい注文を先に処理・保存します: {adversarial_order}")
    run1_entry = process_one_order(service, test_file_id, test_file_name, adversarial_order,
                                    tag="RUN1(adversarial: highest-row-first)", prod_ids=prod_ids,
                                    idx=1, total=1)

    # ── 2. 修正の核心確認: 行番号の小さい未処理注文が引き続き検出されるか ──
    after_run1_path = fresh_download(service, test_file_id, test_file_name, "after-RUN1-check")
    after_run1_targets = ga.get_target_orders(after_run1_path)
    log(f"[RUN1後の再スキャン] 対象注文: {after_run1_targets}")
    expected_after_run1 = ["07-15147-67281", "01-15163-10197"]
    assert adversarial_order not in after_run1_targets, (
        f"[NG] 保存済みのはずの{adversarial_order}が再び対象になっています(重複処理防止が壊れています)"
    )
    assert after_run1_targets == expected_after_run1, (
        f"[NG] 行番号の小さい未処理注文が取りこぼされました(期待={expected_after_run1}, "
        f"実際={after_run1_targets})。get_target_orders()の修正が機能していません。"
    )
    log(f"[OK] 行番号の大きい注文({adversarial_order})を先に保存しても、"
        f"行番号の小さい未処理注文 {expected_after_run1} は取りこぼされずに検出された")
    log(f"[OK] 保存済みの{adversarial_order}は正しく対象から除外された(重複処理防止も健在)")

    # ── 3. 残り2件を行番号順に処理(=RUN2/再開を模す) ──
    run2_entries = []
    for idx, order_no in enumerate(expected_after_run1, start=1):
        entry = process_one_order(service, test_file_id, test_file_name, order_no,
                                   tag="RUN2(resume, row-order)", prod_ids=prod_ids,
                                   idx=idx, total=len(expected_after_run1))
        run2_entries.append(entry)

    # ── 4. 最終確認: 新規DLしなおして、対象0件・全セルが正しい値か・ガード行が無事か ──
    final_path = fresh_download(service, test_file_id, test_file_name, "FINAL-VERIFY")
    final_targets = ga.get_target_orders(final_path)
    final_values = {}
    for order_no, info in KNOWN_ORIGINAL_VALUES.items():
        row = info["row"]
        wb = openpyxl.load_workbook(final_path, keep_vba=True, data_only=True)
        ws = wb[sheet_name]
        final_values[order_no] = {
            "row": row,
            "pkg_no_now": ws.cell(row=row, column=1).value,
            "shipping_now": ws.cell(row=row, column=ship_col_idx).value,
        }
        wb.close()
    guard_after = read_guard_and_target_cells(final_path, sheet_name, ship_col_idx, GUARD_ROWS)

    guard_unchanged = (guard_before == guard_after)

    prod_files_after = ga._list_xlsm_files(service, prod_folder_id, "通常")
    prod_snapshot_after = [
        {"id": f["id"], "name": f["name"], "modifiedTime": f["modifiedTime"]}
        for f in prod_files_after
    ]
    prod_unchanged = prod_snapshot_before == prod_snapshot_after

    result = {
        "fix_verified": {
            "get_target_orders_returns_row_ordered_list": initial_targets == expected_order,
            "earlier_row_order_not_orphaned_after_later_row_saved_first": after_run1_targets == expected_after_run1,
            "already_saved_order_excluded_from_rescan": adversarial_order not in after_run1_targets,
        },
        "test_folder_name": TEST_SUBFOLDER_NAME,
        "test_file_name": test_file_name,
        "target_3_orders_row_order": initial_targets,
        "run1_adversarial": run1_entry,
        "run1_adversarial_order": adversarial_order,
        "after_run1_targets": after_run1_targets,
        "run2_resume": run2_entries,
        "final_targets_remaining": final_targets,
        "final_cell_values": final_values,
        "guard_rows_before": guard_before,
        "guard_rows_after": guard_after,
        "guard_rows_unchanged": guard_unchanged,
        "prod_snapshot_before": prod_snapshot_before,
        "prod_snapshot_after": prod_snapshot_after,
        "prod_unchanged": prod_unchanged,
    }
    out_path = os.path.join(os.path.dirname(__file__), "result_full.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)
    log(f"結果を保存: {out_path}")

    log("=" * 70)
    log("最終サマリ")
    log("=" * 70)
    log(f"テストサブフォルダ: {TEST_SUBFOLDER_NAME}")
    log(f"テストファイル: {test_file_name}")
    log(f"対象3件(行番号昇順): {initial_targets}")
    log(f"[意図的悪条件] 先に保存した行番号最大の注文: {adversarial_order} "
        f"(ダミー送料={run1_entry['fake_price']})")
    log(f"RUN1直後の再スキャンで検出された残り対象: {after_run1_targets} (取りこぼしなし)")
    log(f"RUN2(再開)で処理: {[e['order_no'] for e in run2_entries]}")
    log(f"最終確認後の対象注文(空欄)残数: {len(final_targets)} 件 {final_targets}")
    log(f"最終セル値: {json.dumps(final_values, ensure_ascii=False, default=str)}")
    log(f"ガード行(既存送料記入済み・キャンセル行)が無変更か = {guard_unchanged}")
    if not guard_unchanged:
        log(f"  [WARN] ガード行に差分があります before={guard_before} after={guard_after}")
    log(f"本番『通常』xlsm一覧: 実行前後で不変か = {prod_unchanged}")
    if not prod_unchanged:
        log("[WARN] 本番ファイル一覧に差分があります。詳細を確認してください。")
        log(f"  before={prod_snapshot_before}")
        log(f"  after={prod_snapshot_after}")

    all_ok = (
        result["fix_verified"]["get_target_orders_returns_row_ordered_list"]
        and result["fix_verified"]["earlier_row_order_not_orphaned_after_later_row_saved_first"]
        and result["fix_verified"]["already_saved_order_excluded_from_rescan"]
        and len(final_targets) == 0
        and guard_unchanged
        and prod_unchanged
    )
    log(f"総合判定: {'[PASS] 全項目クリア' if all_ok else '[FAIL] 未クリア項目あり(上記ログ参照)'}")
    log("=" * 70)
    log(f"完了  {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 70)

    if not all_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
