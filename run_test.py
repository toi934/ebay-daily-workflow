"""チェックポイント機構の隔離検証(単発の実行で "1件目→保存" → "再起動→スキップ→残り処理" まで
一気に検証する版)。

本番daily_workflow_ga.py(GitHub Actions版の実物)の関数をそのままimportし、実際のサービス
アカウント・実際のGoogle Drive APIに対して動作させる。CPaSS/Playwrightは一切呼び出さない
(cpass_workflowはダミーモジュールとしてsys.modulesに登録)。

安全設計:
- 本番「売上管理表」フォルダへの書き込みは一切行わない。読み取り専用で
  (1) フォルダID検索 (2) 最新の本番xlsmファイル一覧取得 の2箇所だけ参照する。
- 本番xlsmファイルへの唯一の操作は Drive API の files.copy (サーバー側コピー)。
  これは新しい別ファイルを作るだけで、元ファイルには一切触れない。
- 以降のダウンロード・書き込み・アップロードは、すべてこのコピー(テスト専用サブ
  フォルダ内)に対してのみ行う。production file_id を _upload_xlsm に渡すコードパスは
  存在しない(アサーションで二重に保証)。
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


def get_or_create_test_copy(service, test_folder_id, prod_folder_id):
    q = f"'{test_folder_id}' in parents and name contains '{TEST_FILE_PREFIX}' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name,modifiedTime)",
                                 orderBy="modifiedTime desc").execute()
    files = resp.get("files", [])
    if files:
        log(f"既存テストコピーを再利用: {files[0]['name']} (id={files[0]['id']})")
        return files[0]["id"], files[0]["name"], False

    prod_files = ga._list_xlsm_files(service, prod_folder_id, "通常")
    if not prod_files:
        raise RuntimeError("本番『通常』xlsmが見つかりません")
    prod = prod_files[0]
    log(f"[READ-ONLY] コピー元本番ファイル: {prod['name']} (id={prod['id']}, "
        f"modifiedTime={prod['modifiedTime']})")

    test_name = f"{TEST_FILE_PREFIX}{time.strftime('%Y%m%d_%H%M%S')}.xlsm"
    copied = service.files().copy(
        fileId=prod["id"], body={"name": test_name, "parents": [test_folder_id]}
    ).execute()
    log(f"Drive側サーバーコピー作成完了(本番ファイルは未変更): {test_name} (id={copied['id']})")
    return copied["id"], test_name, True


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


def select_targets_for_blanking(local_path, n=3):
    wb = openpyxl.load_workbook(local_path, keep_vba=True, data_only=True)
    sheet_name = ga.find_sheet_with_orders(wb)
    ws = wb[sheet_name]
    ship_col_idx = ga.find_shipping_column(ws, local_path)

    candidates = []
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if not val or not isinstance(val, str) or not ORDER_PATTERN.match(val.strip()):
            continue
        e_val = ws.cell(row=row, column=5).value
        if "キャンセル" in str(e_val or ""):
            continue
        ship_val = ws.cell(row=row, column=ship_col_idx).value
        pkg_val = ws.cell(row=row, column=1).value
        if not _shipping_empty(ship_val) and pkg_val not in (None, ""):
            candidates.append({"row": row, "order_no": val.strip(),
                                "orig_pkg": pkg_val, "orig_ship": ship_val})
    wb.close()

    if len(candidates) < n:
        raise RuntimeError(f"送料記入済みの候補行が{len(candidates)}件しかありません(必要={n})")

    return candidates[-n:], sheet_name, ship_col_idx


def blank_targets(local_path, sheet_name, ship_col_idx, chosen):
    wb = openpyxl.load_workbook(local_path, keep_vba=True)
    ws = wb[sheet_name]
    for c in chosen:
        ws.cell(row=c["row"], column=1).value = None
        ws.cell(row=c["row"], column=ship_col_idx).value = None
    wb.save(local_path)
    wb.close()


def fresh_download(service, test_file_id, test_file_name, tag):
    """毎回あらためてDriveからDLし直す(=新しいプロセス/新しいActions実行を模す)。"""
    workdir = tempfile.mkdtemp(prefix=f"ga_test_{tag}_")
    local_path = os.path.join(workdir, test_file_name)
    ga._download_xlsm(service, test_file_id, local_path)
    log(f"[{tag}] Driveから再ダウンロード(新規プロセス相当): {test_file_name}")
    return local_path


def run_pass(service, test_file_id, test_file_name, order_limit, tag, prod_ids):
    """1回分の"起動"を模す: 新規DL → get_target_orders() → 上限件数だけ処理 → 都度チェックポイント保存。"""
    local_path = fresh_download(service, test_file_id, test_file_name, tag)

    targets_before = sorted(ga.get_target_orders(local_path))
    log(f"[{tag}] 現在の対象注文(送料空欄): {targets_before}")

    this_pass_targets = targets_before[:order_limit] if order_limit is not None else targets_before
    log(f"[{tag}] 今回処理する注文: {this_pass_targets}")

    cpass_results = {}
    checkpoint_log = []
    for idx, order_no in enumerate(this_pass_targets, start=1):
        log(f"[{tag}] {idx}/{len(this_pass_targets)}件目 処理開始(ダミー送料、CPaSS未呼び出し): {order_no}")
        time.sleep(1)
        fake_price = FAKE_PRICE_BASE + (sum(ord(ch) for ch in order_no) % 9)
        info = {
            "package_no": 9000 + idx,
            "dhl_price_jpy": fake_price,
            "title": "TEST-DUMMY",
            "item_id": "TEST-DUMMY",
        }
        cpass_results[order_no] = info

        ok, num_writes = ga.process_xlsm(local_path, cpass_results, dry_run=False)
        entry = {"order_no": order_no, "fake_price": fake_price, "ok": ok, "num_writes": num_writes,
                  "time": time.strftime("%Y-%m-%d %H:%M:%S")}
        if ok and num_writes > 0:
            assert test_file_id not in prod_ids, "重大な設定ミス: テストfile_idが本番file_idと一致"
            ga._upload_xlsm(service, test_file_id, local_path)
            entry["uploaded"] = True
            log(f"[{tag}] [CHECKPOINT] {idx}/{len(this_pass_targets)}件目: {num_writes}セル保存 "
                f"対象注文={order_no} 送料(ダミー)={fake_price} → Google Driveアップロード成功")
        else:
            entry["uploaded"] = False
            log(f"[{tag}] [CHECKPOINT] {idx}/{len(this_pass_targets)}件目: 書き込みなし(num_writes={num_writes})")
        checkpoint_log.append(entry)

    targets_after = sorted(ga.get_target_orders(local_path))
    skipped = sorted(set(targets_before) - set(this_pass_targets))
    return {
        "tag": tag,
        "targets_before": targets_before,
        "processed": checkpoint_log,
        "targets_after": targets_after,
        "not_processed_this_pass": skipped,
    }


def main():
    log("=" * 70)
    log(f"開始  {time.strftime('%Y-%m-%d %H:%M:%S')}")
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
    log(f"[READ-ONLY] 本番『通常』xlsm一覧(実行前スナップショット, {len(prod_snapshot_before)}件):")
    for f in prod_snapshot_before:
        log(f"    {f['name']} (id={f['id']}, modifiedTime={f['modifiedTime']})")

    test_folder_id, folder_created = get_or_create_test_subfolder(service, prod_folder_id)
    test_file_id, test_file_name, file_created = get_or_create_test_copy(
        service, test_folder_id, prod_folder_id
    )
    assert test_file_id not in prod_ids, "重大な設定ミス: テストfile_idが本番file_idと一致しています"

    chosen_summary = []
    if file_created:
        local_path = fresh_download(service, test_file_id, test_file_name, "setup")
        chosen, sheet_name, ship_col_idx = select_targets_for_blanking(local_path, n=3)
        log(f"送料空欄化対象として選定した{len(chosen)}件(末尾の送料記入済み行):")
        for c in chosen:
            log(f"    行{c['row']} 注文={c['order_no']} 元梱包番号={c['orig_pkg']} 元送料={c['orig_ship']}")
            chosen_summary.append({k: c[k] for k in ("row", "order_no", "orig_pkg", "orig_ship")})
        blank_targets(local_path, sheet_name, ship_col_idx, chosen)
        ga._upload_xlsm(service, test_file_id, local_path)
        log("空欄化後のテストコピーをGoogle Driveへアップロード完了(セットアップ完了)")
    else:
        log("既存のテストコピーを再利用するため、空欄化セットアップはスキップ")

    # ── 1回目の"起動": 1件だけ処理して正常終了 ──
    result_run1 = run_pass(service, test_file_id, test_file_name, order_limit=1,
                            tag="RUN1(limit=1)", prod_ids=prod_ids)

    # ── 2回目の"起動"(再起動を模す): 新規DLからやり直し、残りを処理 ──
    result_run2 = run_pass(service, test_file_id, test_file_name, order_limit=None,
                            tag="RUN2(resume)", prod_ids=prod_ids)

    # ── 最終確認: もう一度新規DLして全件埋まっていることを確認 ──
    final_path = fresh_download(service, test_file_id, test_file_name, "FINAL-VERIFY")
    final_targets = sorted(ga.get_target_orders(final_path))
    wb = openpyxl.load_workbook(final_path, keep_vba=True, data_only=True)
    sheet_name = ga.find_sheet_with_orders(wb)
    ws = wb[sheet_name]
    ship_col_idx = ga.find_shipping_column(ws, final_path)
    final_values = {}
    for c in chosen_summary:
        row = c["row"]
        final_values[c["order_no"]] = {
            "row": row,
            "pkg_no_now": ws.cell(row=row, column=1).value,
            "shipping_now": ws.cell(row=row, column=ship_col_idx).value,
        }
    wb.close()

    prod_files_after = ga._list_xlsm_files(service, prod_folder_id, "通常")
    prod_snapshot_after = [
        {"id": f["id"], "name": f["name"], "modifiedTime": f["modifiedTime"]}
        for f in prod_files_after
    ]
    prod_unchanged = prod_snapshot_before == prod_snapshot_after

    result = {
        "test_folder_name": TEST_SUBFOLDER_NAME,
        "test_file_name": test_file_name,
        "chosen_for_blanking": chosen_summary,
        "run1": result_run1,
        "run2_resume": result_run2,
        "final_targets_remaining": final_targets,
        "final_cell_values": final_values,
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
    log(f"対象3件: {[c['order_no'] for c in chosen_summary] if chosen_summary else '(前回セットアップ済みを再利用)'}")
    log(f"RUN1で処理: {[p['order_no'] for p in result_run1['processed']]}")
    log(f"RUN2で処理(再開分): {[p['order_no'] for p in result_run2['processed']]}")
    _run1_saved = [p["order_no"] for p in result_run1["processed"] if p.get("uploaded")]
    log(f"RUN2開始時点で対象から除外(=RUN1保存分が正しくスキップされた): "
        f"{[o for o in _run1_saved if o not in result_run2['targets_before']]}")
    log(f"最終確認後の対象注文(空欄)残数: {len(final_targets)} 件 {final_targets}")
    log(f"最終セル値: {json.dumps(final_values, ensure_ascii=False, default=str)}")
    log(f"本番『通常』xlsm一覧: 実行前後で不変か = {prod_unchanged}")
    if not prod_unchanged:
        log("[WARN] 本番ファイル一覧に差分があります。詳細を確認してください。")
        log(f"  before={prod_snapshot_before}")
        log(f"  after={prod_snapshot_after}")
    log("=" * 70)
    log(f"完了  {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 70)


if __name__ == "__main__":
    main()
