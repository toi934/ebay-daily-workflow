"""毎朝11:00 売上管理表ワークフロー (GitHub Actions版)

ローカル版との違い:
- Google Drive APIでxlsmをDL/UL（ローカルマイドライブ不要）
- win32comの代わりにopenpyxlで書き込み（keep_vba=Trueで保護）
- CPaSS/Playwrightはheadlessモードで実行
- 完了後にGmailでメール送信

必要なGitHub Secrets:
  CPASS_EMAIL / CPASS_PASSWORD
  GOOGLE_SERVICE_ACCOUNT_JSON
  GMAIL_APP_PASSWORD

★2026/09/10変更（GitHub Actions 60分タイムアウト対策）:
旧方式は「全注文のCPaSS処理が完了してからまとめてExcel書き込み・Drive UL」だったため、
60分タイムアウトでジョブがKillされるとそのRunの成果が100%失われていた（9/8・9/9・9/10と
3日連続で発生、約160件のバックログが一切進まない状態になっていた）。
→ 1件（またはCHECKPOINT_EVERY_N件）処理するたびにExcel書き込み→Google Driveへ途中保存する
  方式に変更。次回起動時は get_target_orders() の「送料セル空白」スキャンが自動的に
  処理済み注文をスキップするため、追加の状態ファイルなしで再開できる。
  さらに cpass_workflow.process_all_orders_for_dhl() に deadline_ts を渡し、
  60分上限に近づいたら新規注文の処理を開始せず安全に打ち切ってもらう（[SAFE STOP]）。
"""

import sys
import os
import re
import time
import json
import io
import smtplib
import tempfile
from email.mime.text import MIMEText
from datetime import datetime, timezone, timedelta

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# openpyxl互換バグ回避
try:
    from openpyxl.descriptors.base import Typed as _OpxlTyped
    _orig_typed_set = _OpxlTyped.__set__
    def _patched_typed_set(self, instance, value):
        try:
            _orig_typed_set(self, instance, value)
        except (ValueError, TypeError):
            pass
    _OpxlTyped.__set__ = _patched_typed_set
except Exception:
    pass

import openpyxl
import cpass_workflow

# ─── 定数 ───
E_COL_VALUE = "③マーキング番号、リサーチ者記入"
F_COL_VALUE = "仕入未"
EXCHANGE_RATE_KEYWORDS = ["為替", "為"]

# ★2026/09/06追加: シート再構成でBL固定列と実際のヘッダー位置がズレる問題が発覚したため、
# 為替列と同様「国際送料」ヘッダーを動的に探す方式に変更（見つからない場合のみ旧BL/BR固定にフォールバック）。
SHIPPING_HEADER_KEYWORDS = ["国際送料"]
DRIVE_FOLDER_NAME = "売上管理表"
GMAIL_FROM = "gen7m9@gmail.com"
# ★2026/09/14追加: 結果メールの時刻表示をJST固定にするため。
# GitHub Actionsランナーの実行TZ(UTC)に依存しないよう明示的に+9時間する。
JST = timezone(timedelta(hours=9))

# ★2026/09/10追加: 途中保存(チェックポイント)関連のデフォルト設定。
# どちらも環境変数で上書き可能（テスト時などに調整しやすいように）。
DEFAULT_SAFETY_BUDGET_SECONDS = 2700  # 45分（60分上限に対して15分のバッファ）
DEFAULT_CHECKPOINT_EVERY_N = 1        # 1件処理するごとに保存（最も安全）


# ─── Google Drive API ───
def _get_drive_service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(sa_json)
    # GitHub Secretsの改行エスケープ問題を修正（\\n → 実改行）
    if "private_key" in info:
        info["private_key"] = info["private_key"].replace(chr(92) + chr(110), chr(10))
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def _find_folder_id(service, folder_name):
    """マイドライブ内のフォルダIDを取得"""
    q = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name)").execute()
    files = resp.get("files", [])
    if not files:
        raise RuntimeError(f"Google Driveにフォルダ '{folder_name}' が見つかりません")
    return files[0]["id"]


def _list_xlsm_files(service, folder_id, prefix):
    """フォルダ内のxlsmファイル一覧（prefix: '通常' or '専門'）を取得"""
    q = f"'{folder_id}' in parents and name contains '{prefix}' and name contains '.xlsm' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc").execute()
    return resp.get("files", [])


def _download_xlsm(service, file_id, dest_path):
    """Google DriveからxlsmをDL"""
    from googleapiclient.http import MediaIoBaseDownload
    request = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def _upload_xlsm(service, file_id, src_path):
    """Google Driveの既存ファイルを上書きUL

    ★2026/07/09追加・バグ修正:
    通常_7月9日_9時_売上管理表.xlsm のアップロードで
    `EOF occurred in violation of protocol (_ssl.c:2427)` という一過性SSLエラーが
    2回連続で発生し、CPaSSで取得済みのDHL送料8件がGoogle Driveに反映されない
    問題が起きていた（openpyxlでのローカル書き込み・保存自体は成功していた）。
    旧実装は resumable=False で1回のHTTPリクエストにファイル全体を乗せていたため、
    途中で接続が切れると丸ごと失敗し、リトライも一切行っていなかった。
    → resumable=True（チャンク分割アップロード）に変更し、さらに外側で
      最大3回・指数バックオフ付きリトライを行うようにした。
    """
    from googleapiclient.http import MediaFileUpload

    last_err = None
    for attempt in range(1, 4):
        try:
            media = MediaFileUpload(
                src_path,
                mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
                resumable=True,
            )
            request = service.files().update(fileId=file_id, media_body=media)
            response = None
            while response is None:
                _, response = request.next_chunk()
            print(f"  Google Driveにアップロード完了: {os.path.basename(src_path)}"
                  + (f"（試行{attempt}回目で成功）" if attempt > 1 else ""))
            return
        except Exception as e:
            last_err = e
            print(f"  [WARN] Google Driveアップロード失敗(試行{attempt}/3): {str(e)[:150]}")
            if attempt < 3:
                time.sleep(5 * attempt)
    print(f"  [ERROR] Google Driveアップロード 3回とも失敗: {os.path.basename(src_path)}")
    raise last_err


def download_xlsm_files(workdir):
    """通常・専門の最新xlsmをDL。{prefix: (local_path, drive_file_id)} を返す"""
    print("=" * 60)
    print("Step 1: Google DriveからDL")
    print("=" * 60)
    service = _get_drive_service()
    folder_id = _find_folder_id(service, DRIVE_FOLDER_NAME)
    print(f"  フォルダID: {folder_id}")

    result = {}
    for prefix in ["通常", "専門"]:
        files = _list_xlsm_files(service, folder_id, prefix)
        if not files:
            print(f"  {prefix}: ファイルなし")
            continue
        latest = files[0]  # modifiedTime descでソート済み
        local_path = os.path.join(workdir, latest["name"])
        _download_xlsm(service, latest["id"], local_path)
        print(f"  DL: {latest['name']} (id={latest['id']})")
        result[prefix] = (local_path, latest["id"])

    return service, folder_id, result


# ─── Excel処理 ───
def get_shipping_col(xlsm_path):
    if "通常" in os.path.basename(xlsm_path):
        return "BL"
    return "BR"


def find_sheet_with_orders(wb):
    from openpyxl.chartsheet.chartsheet import Chartsheet
    order_pattern = re.compile(r"^\d{2}-\d{5}-\d{5}$")
    best_sheet, best_count = None, 0
    for sname in wb.sheetnames:
        ws = wb[sname]
        if isinstance(ws, Chartsheet):
            continue
        count = 0
        for row in range(2, min(500, ws.max_row) + 1):
            val = ws.cell(row=row, column=2).value
            if val and isinstance(val, str) and order_pattern.match(val.strip()):
                count += 1
        if count > best_count:
            best_count, best_sheet = count, sname
    print(f"  対象シート: {best_sheet} (注文番号{best_count}件)")
    return best_sheet


def find_exchange_cols(ws):
    found = []
    for row in range(1, 4):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if val and isinstance(val, str) and any(kw in val for kw in EXCHANGE_RATE_KEYWORDS):
                found.append({"row": row, "col": col,
                              "letter": openpyxl.utils.get_column_letter(col), "value": val})
    return found


def find_shipping_column(ws, xlsm_path):
    """国際送料の列を探す（ヘッダー行から「国際送料」キーワード、優先）。

    ★2026/09/06追加: シート再構成で「国際送料」ヘッダーの実際の位置が従来の固定列
    (通常=BL/専門=BR)からズレる事例が発覚したため、為替列と同じ方式（ヘッダー文字列を
    動的に探す）に変更した。ヘッダーが見つからない場合のみ従来の固定列にフォールバックする。
    """
    for row in range(1, 4):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if val and isinstance(val, str) and any(kw in val for kw in SHIPPING_HEADER_KEYWORDS):
                return col
    fallback_letter = get_shipping_col(xlsm_path)
    print(f"  [WARN] 「国際送料」ヘッダーが見つからず、従来の固定列({fallback_letter})にフォールバック")
    return openpyxl.utils.column_index_from_string(fallback_letter)


def get_target_orders(xlsm_path):
    """送料セル空白の未処理注文番号を「行番号の昇順リスト」で返す

    ★2026/09/10チェックポイント機構導入時の想定:
      - 実行中にチェックポイント保存された注文 → 次回この関数を呼んだ時点で
        既に送料セルが埋まっている → targetsから自動的に除外される（＝再開・重複防止）。
      追加の状態ファイル(「何件目まで処理済み」等)を持たない設計。

    ★2026/09/10 追加修正（チェックポイント隔離テストで発見した不具合の対応）:
      旧実装は「送料記入済み最終行(last_filled_row)より後ろの行だけ」を対象注文の
      スキャン範囲にしていた。ところが1回の実行で複数件を処理する場合、行番号が
      後ろの注文が先にチェックポイント保存されると last_filled_row がその行まで
      一気に進んでしまい、まだ未処理のまま残っている「それより前の行」の注文が
      次回のスキャンで検出されなくなる（＝送料が空欄のまま、エラーも出さずに
      永久に取りこぼされる）不具合があった。
      → last_filled_row はログ表示用の診断情報としてのみ残し、対象注文の判定
      自体はシート全体（行2〜最終行）を対象にするよう変更。返り値もset()ではなく
      行番号順のlistとし、CPaSSへ渡す処理順が不定にならないようにした。
    """
    order_pattern = re.compile(r"^\d{2}-\d{5}-\d{5}$")

    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True, data_only=True)
    sheet_name = find_sheet_with_orders(wb)
    if not sheet_name:
        wb.close()
        return []
    ws = wb[sheet_name]
    ship_col_idx = find_shipping_column(ws, xlsm_path)
    shipping_col = openpyxl.utils.get_column_letter(ship_col_idx)

    def _shipping_empty(v):
        # ★2026/07/03修正: None/空文字だけでなく 0・空白文字列・数式の空結果も「未記入」扱い
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

    order_rows = []
    last_filled_row = 1
    targets = []
    seen_orders = set()
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if not val or not isinstance(val, str) or not order_pattern.match(val.strip()):
            continue
        order_no = val.strip()
        order_rows.append(row)
        br_val = ws.cell(row=row, column=ship_col_idx).value
        shipping_empty = _shipping_empty(br_val)
        if not shipping_empty:
            last_filled_row = row

        # ★対象判定はシート全体で行う（last_filled_rowによる範囲制限はしない。上記docstring参照）。
        #   「送料記入不要」条件（キャンセル等）は従来どおり維持する。
        e_val = ws.cell(row=row, column=5).value
        if "キャンセル" in str(e_val or ""):
            continue
        if shipping_empty and order_no not in seen_orders:
            targets.append(order_no)
            seen_orders.add(order_no)

    print(f"  送料記入済み最終行: {last_filled_row}（診断用ログのみ。対象注文の判定には使用しない）")

    # ★診断ログ: 末尾15注文行の送料セル生値（誤認バグ調査用）
    for row in order_rows[-15:]:
        _v = ws.cell(row=row, column=ship_col_idx).value
        _o = str(ws.cell(row=row, column=2).value).strip()
        _e = str(ws.cell(row=row, column=5).value or "")[:8]
        print(f"    [DEBUG] 行{row} {_o} {shipping_col}={_v!r} E={_e}")

    wb.close()
    return targets


def process_xlsm(xlsm_path, cpass_results, dry_run=False):
    """xlsmを開いてCPaSSデータで空白セルを埋め、openpyxlで保存

    ★2026/09/10変更: 戻り値を bool から (success: bool, num_writes: int) に変更。
    呼び出し側（途中保存のチェックポイント処理）が「今回新たに書き込みが発生したか」を
    判定できるようにするため。新規書き込みが0件ならGoogle Driveへの再アップロードを
    スキップし、無駄なAPI呼び出し・時間消費を避ける。
    """
    print()
    print("=" * 60)
    print(f"Excel処理: {os.path.basename(xlsm_path)}")
    print("=" * 60)

    order_pattern = re.compile(r"^\d{2}-\d{5}-\d{5}$")

    # 読み取り専用で開いてwritesを収集
    wb = openpyxl.load_workbook(xlsm_path, keep_vba=True, data_only=True)
    sheet_name = find_sheet_with_orders(wb)
    if not sheet_name:
        print("  注文データのシートが見つかりません")
        wb.close()
        return False, 0
    ws = wb[sheet_name]
    ship_col_idx = find_shipping_column(ws, xlsm_path)
    shipping_col = openpyxl.utils.get_column_letter(ship_col_idx)

    ex_cols = find_exchange_cols(ws)
    writes = []
    fill_details = {"A": 0, "E": 0, "F": 0, shipping_col: 0, "exchange": 0}
    # 為替カスケード用: 静的スナップショット参照バグ修正（2026/07/26）
    # 行を舐めながら「直近で確定した為替値」をここに保持し続ける（writesに積んだ値も反映）。
    # これにより同一実行内で新規空欄行が連続しても正しく前の値を引き継げる。
    last_exchange_val = {ec["col"]: None for ec in ex_cols}

    for row in range(2, (ws.max_row or 9999) + 1):
        order_val = ws.cell(row=row, column=2).value
        if not order_val or not isinstance(order_val, str):
            continue
        order_no = order_val.strip()
        if not order_pattern.match(order_no):
            continue

        cpass_info = cpass_results.get(order_no)
        e_val = ws.cell(row=row, column=5).value
        is_cancel = "キャンセル" in str(e_val or "")

        if cpass_info and cpass_info.get("package_no"):
            if not ws.cell(row=row, column=1).value:
                writes.append((row, 1, cpass_info["package_no"]))
                fill_details["A"] += 1

        if not e_val:
            writes.append((row, 5, E_COL_VALUE))
            fill_details["E"] += 1

        f_val = ws.cell(row=row, column=6).value
        if not f_val:
            writes.append((row, 6, F_COL_VALUE))
            fill_details["F"] += 1

        if cpass_info and cpass_info.get("dhl_price_jpy"):
            br_val = ws.cell(row=row, column=ship_col_idx).value
            if not br_val:
                writes.append((row, ship_col_idx, int(cpass_info["dhl_price_jpy"])))
                fill_details[shipping_col] += 1

        # === 為替セル: 直近の確定値を保持しながら上から順にコピー（キャンセル行は書き込みのみスキップ）===
        for ex_col in ex_cols:
            col = ex_col["col"]
            cell_val = ws.cell(row=row, column=col).value
            if cell_val is not None and not (isinstance(cell_val, str) and cell_val.strip() == ""):
                # 実値が入っている行 → 以降の空欄行の引き継ぎ元として更新
                last_exchange_val[col] = cell_val
            elif not is_cancel:
                if last_exchange_val.get(col) is not None:
                    # last_exchange_val はここで書き換えない＝そのまま次の空欄行にも引き継がれる
                    writes.append((row, col, last_exchange_val[col]))
                    fill_details["exchange"] += 1

    wb.close()

    print(f"  書き込み予定: {json.dumps(fill_details, ensure_ascii=False)}")

    if dry_run:
        print("  [DRY RUN] 書き込みなし")
        return True, len(writes)

    if not writes:
        print("  書き込みなし（全セル埋まり済み）")
        return True, 0

    # openpyxlで書き込み（keep_vba=TrueでVBA保持）
    print(f"  openpyxl で書き込み中... ({len(writes)} セル)")
    wb2 = openpyxl.load_workbook(xlsm_path, keep_vba=True)
    ws2 = wb2[sheet_name]
    for (row, col, value) in writes:
        ws2.cell(row=row, column=col).value = value
    wb2.save(xlsm_path)
    wb2.close()
    print(f"  保存OK: {os.path.basename(xlsm_path)}")
    return True, len(writes)


# ─── ★2026/09/25追加: 実行結果の健全性チェック（「処理0件なのにsuccess」の再発防止） ───
UNMATCHED_STATE_FILE = "task2_unmatched_orders.txt"   # 前回実行時点で未処理のまま残った対象注文
ANOMALY_MARKER_FILE = "task2_anomaly.txt"             # 異常時のみ作成 → workflow最終ステップでfailure化


def load_prev_unmatched(path=UNMATCHED_STATE_FILE):
    """前回実行時に未処理で残った注文番号の集合。ファイルが無ければNone（比較不能）。"""
    try:
        with open(path, encoding="utf-8") as f:
            return set(l.strip() for l in f if re.match(r"^\d{2}-\d{5}-\d{5}$", l.strip()))
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"  [WARN] {path} 読み込み失敗: {e}")
        return None


def save_unmatched(orders, path=UNMATCHED_STATE_FILE):
    with open(path, "w", encoding="utf-8") as f:
        for o in sorted(set(orders)):
            f.write(o + "\n")


def evaluate_run_health(target_order_nos, diag, cpass_error, dhl_ok_orders,
                        total_writes, drive_uploads, prev_unmatched):
    """今回の実行が「正常完了」と言えるかを判定する（純粋関数・テスト対象）。

    返り値: dict(level, reasons, notes, new_targets, new_missing)
      level: "ok"      … 正常完了（対象を処理した／本当に新規対象が無い日）
             "warning" … 一部未処理など注意（failureにはしない）
             "anomaly" … 異常。Excelに新規の未処理対象があるのにCPaSS処理0件 等
    方針（戸井さん指示 2026/09/25）:
      - Excel上に未処理対象があるのにCPaSS取得・処理が0件 → 正常終了と判定しない
      - 本当に新規対象注文が0件の日は正常終了のまま
        （毎日CPaSS発送手続きタブに存在しない古い送料空欄注文が約90件あるため、
         「前回実行時にも未処理だった注文」は新規扱いしない）
    """
    diag = diag or {}
    targets = list(target_order_nos or [])
    missing = set(diag.get("missing_order_nos") or [])
    matched = list(diag.get("matched_order_nos") or [])
    move_status = diag.get("move_status")
    reasons, notes = [], []

    if prev_unmatched is None:
        new_targets = list(targets)  # 比較基準が無い初回は全件を新規として扱う
        notes.append("前回の未処理リストが無いため、全対象を新規として判定")
    else:
        new_targets = [o for o in targets if o not in prev_unmatched]
    new_missing = [o for o in new_targets if o in missing]
    processed = len(dhl_ok_orders)

    if cpass_error:
        reasons.append(f"CPaSS処理が例外で中断: {cpass_error}")
    if move_status == "failed":
        reasons.append("発送手続き待ち→発送手続きの一括移動に失敗: " + str(diag.get("move_detail", "")))
    if diag.get("stuck_in_waiting_order_nos"):
        reasons.append("発送手続き待ちの対象注文が発送手続きへ移っていない: "
                       + ", ".join(diag["stuck_in_waiting_order_nos"]))
    if matched and processed == 0 and not diag.get("safe_stopped"):
        reasons.append(f"CPaSS発送手続きタブで対象{len(matched)}件を見つけたが送料取得0件")
    if processed > 0 and total_writes == 0:
        reasons.append(f"送料取得{processed}件なのにExcel書き込み0件")
    if total_writes > 0 and drive_uploads == 0:
        reasons.append(f"Excel書き込み{total_writes}セルなのにGoogle Drive更新0件")
    if processed == 0 and new_missing:
        reasons.append(
            f"Excelに新規の未処理対象{len(new_missing)}件があるのにCPaSS処理0件"
            f"（発送手続きタブに無い: {', '.join(new_missing[:10])}"
            f"{' …' if len(new_missing) > 10 else ''}）")
    elif new_missing:
        notes.append(f"新規対象のうちCPaSS発送手続きタブに無いもの{len(new_missing)}件: "
                     + ", ".join(new_missing[:10]))

    if reasons:
        level = "anomaly"
    elif new_missing or diag.get("safe_stopped"):
        level = "warning"
    else:
        level = "ok"
    return {"level": level, "reasons": reasons, "notes": notes,
            "new_targets": new_targets, "new_missing": new_missing}


def build_health_header(health, processed, total_writes, drive_uploads, n_targets):
    """結果メール冒頭の判定ブロック（正常完了と誤認できない表示）。"""
    lvl = health["level"]
    if lvl == "anomaly":
        head = "■判定: ★異常（正常完了ではありません）★"
    elif lvl == "warning":
        head = "■判定: 注意（一部未処理あり）"
    elif n_targets and not health["new_targets"] and processed == 0:
        head = "■判定: 正常完了（新規の対象注文なし）"
    else:
        head = "■判定: 正常完了"
    lines = [head,
             f"処理対象(CPaSSで送料取得): {processed}件",
             f"送料書込み: {total_writes}セル",
             f"Drive更新: {drive_uploads}回" if drive_uploads else "Drive更新: なし",
             f"新規の対象注文: {len(health['new_targets'])}件"]
    for r in health["reasons"]:
        lines.append("  ×" + r)
    for n in health["notes"]:
        lines.append("  ・" + n)
    return "\n".join(lines) + "\n\n"


# ─── メール送信 ───
def send_result_email(results_text, run_count, error_text="", start_time_jst=None,
                      subject_prefix="", header_text=""):
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    subject = f"{subject_prefix}タスク2（売上管理表）実行結果 ({run_count}回目)"
    if not app_password:
        print(f"  GMAIL_APP_PASSWORD未設定、メールスキップ（件名予定: {subject}）")
        return
    # ★2026/09/14修正: 「実行日時」がワークフローの開始時刻ではなく、この関数の
    # 呼び出し直前（＝処理完了直後）のdatetime.now()になっており、実質的に
    # 完了時刻を「実行日時」と誤表示していた（2026/09/14の調査で判明）。
    # 開始日時（main()冒頭のstart_time_jst）と完了日時（この関数の呼び出し時点）を
    # 分けて表示するよう修正。GitHub Actionsランナーの実行TZ(UTC)に依存しないよう、
    # 明示的にUTC取得→JST(UTC+9)へ変換してから表示する。
    end_time_jst = datetime.now(timezone.utc).astimezone(JST)
    if start_time_jst is not None:
        elapsed = int((end_time_jst - start_time_jst).total_seconds())
        body = (
            f"開始日時: {start_time_jst.strftime('%Y/%m/%d %H:%M:%S')} JST\n"
            f"完了日時: {end_time_jst.strftime('%Y/%m/%d %H:%M:%S')} JST\n"
            f"経過時間: {elapsed}秒\n\n"
        )
    else:
        # start_time_jst未指定時（想定外呼び出し向けフォールバック）
        body = f"完了日時: {end_time_jst.strftime('%Y/%m/%d %H:%M:%S')} JST\n\n"
    body += header_text
    body += results_text
    if error_text:
        body += f"\n\n【エラー】\n{error_text}"
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = GMAIL_FROM
    msg["To"] = GMAIL_FROM
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_FROM, app_password)
            smtp.send_message(msg)
        print(f"  メール送信: {subject}")
    except Exception as e:
        print(f"  メール送信失敗: {e}")


def _get_run_count():
    count_file = "run_count_task2.txt"
    try:
        with open(count_file) as f:
            n = int(f.read().strip()) + 1
    except Exception:
        n = 1
    with open(count_file, "w") as f:
        f.write(str(n))
    return n


# ─── メイン ───
def main():
    dry_run = "--dry-run" in sys.argv
    start_time = datetime.now()
    start_time_jst = datetime.now(timezone.utc).astimezone(JST)
    start_ts = time.time()

    print("=" * 60)
    print(f"売上管理表ワークフロー (GitHub Actions版)  {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    run_count = _get_run_count()
    errors = []
    summary_lines = []

    # ★2026/09/10追加: 途中保存(チェックポイント)関連の設定値。
    # SAFETY_BUDGET_SECONDS: ワークフロー開始からこの秒数を過ぎたら、
    #   cpass_workflow側に新規注文の処理開始を止めてもらう（60分ジョブ上限に対し15分の余裕）。
    # CHECKPOINT_EVERY_N: 何件処理するごとにExcel書き込み・Drive ULを行うか（既定=1件ごと）。
    # どちらもGitHub Actions側の環境変数で上書き可能（テスト時などに調整するため）。
    try:
        SAFETY_BUDGET_SECONDS = int(os.environ.get("SAFETY_BUDGET_SECONDS", str(DEFAULT_SAFETY_BUDGET_SECONDS)))
    except ValueError:
        SAFETY_BUDGET_SECONDS = DEFAULT_SAFETY_BUDGET_SECONDS
    try:
        CHECKPOINT_EVERY_N = max(1, int(os.environ.get("CHECKPOINT_EVERY_N", str(DEFAULT_CHECKPOINT_EVERY_N))))
    except ValueError:
        CHECKPOINT_EVERY_N = DEFAULT_CHECKPOINT_EVERY_N
    deadline_ts = start_ts + SAFETY_BUDGET_SECONDS
    print(f"  安全停止デッドライン: 開始から{SAFETY_BUDGET_SECONDS}秒後 "
          f"/ チェックポイント間隔: {CHECKPOINT_EVERY_N}件ごと")

    with tempfile.TemporaryDirectory() as workdir:
        # Step 1: Google Drive からDL
        try:
            service, folder_id, dl_result = download_xlsm_files(workdir)
        except Exception as e:
            msg = f"Google Drive DL失敗: {e}"
            print(msg)
            errors.append(msg)
            send_result_email("Google Drive DL失敗", run_count, "\n".join(errors), start_time_jst=start_time_jst)
            return

        if not dl_result:
            print("対象ファイルなし、終了")
            send_result_email("対象ファイルなし", run_count, start_time_jst=start_time_jst)
            return

        # Step 2: 対象注文番号収集
        # ★2026/09/10修正: target_order_nosはset()ではなく行番号順のlistとして組み立てる。
        # set()だと反復順序がPythonのハッシュ値依存で不定になり、CPaSSへ渡す処理順も
        # 不定になってしまう（チェックポイント隔離テストで、行順序と異なる順で保存された
        # 場合に未処理注文が取りこぼされる不具合につながることが判明したため）。
        _seen_order_nos = set()
        xlsm_paths = {}
        per_prefix_targets = {}
        for prefix, (local_path, file_id) in dl_result.items():
            try:
                targets = get_target_orders(local_path)
                per_prefix_targets[prefix] = targets
                xlsm_paths[prefix] = (local_path, file_id)
                print(f"  {prefix}: 対象{len(targets)}件")
                summary_lines.append(f"{prefix}: 対象注文{len(targets)}件")
            except Exception as e:
                msg = f"{prefix}の注文抽出失敗: {e}"
                print(msg)
                errors.append(msg)

        # ラウンドロビン方式: 通常・専門から1件ずつ交互に対象へ追加し、
        # 片方が尽きたら残りをもう片方から処理する
        # （一方のアカウントが常に後回しになるのを防ぐ2026/09/13修正）
        target_order_nos = []
        _rr_queues = {p: list(t) for p, t in per_prefix_targets.items()}
        _rr_prefix_order = list(per_prefix_targets.keys())
        while any(_rr_queues[p] for p in _rr_prefix_order):
            for p in _rr_prefix_order:
                if _rr_queues[p]:
                    _order_no = _rr_queues[p].pop(0)
                    if _order_no not in _seen_order_nos:
                        target_order_nos.append(_order_no)
                        _seen_order_nos.add(_order_no)

        print(f"\n送料列が空白の対象注文: {len(target_order_nos)} 件")

        # ★2026/07/13 四報の修正（.assign_shipping自動クローズ対応＋見積もりステップ追加）を
        # 本番投入する前に、まず少数件で実際のPlaywright自動化フローを通して確認するための
        # 安全弁。VERIFY_ORDER_NOS環境変数（カンマ区切りの注文番号）が設定されている場合のみ
        # 対象をその注文に限定する。未設定時（通常の本番運用）は従来通り全件が対象になる。
        _verify_orders_env = os.environ.get("VERIFY_ORDER_NOS", "").strip()
        if _verify_orders_env:
            _verify_set = set(x.strip() for x in _verify_orders_env.split(",") if x.strip())
            _before = len(target_order_nos)
            target_order_nos = [o for o in target_order_nos if o in _verify_set]
            print(f"  [検証モード] VERIFY_ORDER_NOS指定あり → 対象を{_before}件から"
                  f"{len(target_order_nos)}件に制限: {target_order_nos}")

        # ─── 途中保存(チェックポイント)まわりの状態 ───
        # ★2026/09/10追加: cpass_results はCPaSS処理の進行に応じて逐次このdictへ反映され、
        # process_xlsm() へ渡すたびに「その時点までに取得できた送料」でExcelへ書き込む。
        cpass_results = {}
        checkpoint_state = {"since_last_save": 0, "saved_total": 0, "pending_orders": [], "last_error": None,
                            "total_writes": 0, "drive_uploads": 0}
        cpass_diag = {}
        cpass_error = None

        def _do_checkpoint(reason):
            """現時点のcpass_resultsで対象xlsmへ書き込み→Google Driveへ途中保存する。"""
            pending = list(checkpoint_state["pending_orders"])
            any_saved = False
            for _prefix, (_local_path, _file_id) in xlsm_paths.items():
                try:
                    ok, num_writes = process_xlsm(_local_path, cpass_results, dry_run=dry_run)
                    if ok and num_writes > 0 and not dry_run:
                        checkpoint_state["total_writes"] += num_writes
                        _upload_xlsm(service, _file_id, _local_path)
                        checkpoint_state["drive_uploads"] += 1
                        any_saved = True
                        print(f"  [CHECKPOINT] {_prefix}: {num_writes}セル保存 "
                              f"対象注文={pending} → Google Driveアップロード成功 ({reason})")
                    elif ok and num_writes == 0:
                        print(f"  [CHECKPOINT] {_prefix}: 新規書き込みなし、ULスキップ ({reason})")
                except Exception as e:
                    msg = f"{_prefix}のチェックポイント保存失敗: {e}"
                    print(f"  [WARN] {msg}")
                    checkpoint_state["last_error"] = msg
            checkpoint_state["pending_orders"] = []
            return any_saved

        def _on_order_done(order_no, info, idx, total):
            """cpass_workflow側から1注文の処理が終わるたびに呼ばれるコールバック。"""
            cpass_results[order_no] = info
            checkpoint_state["since_last_save"] += 1
            checkpoint_state["saved_total"] += 1
            checkpoint_state["pending_orders"].append(order_no)
            print(f"  [CHECKPOINT] {idx}/{total}件目まで処理済み: 対象注文={order_no} "
                  f"送料={info.get('dhl_price_jpy')}")
            if checkpoint_state["since_last_save"] >= CHECKPOINT_EVERY_N:
                _do_checkpoint(f"{idx}/{total}件目時点")
                checkpoint_state["since_last_save"] = 0

        # Step 3: CPaSS処理（1件ごとに _on_order_done で途中保存）
        # ★2026/07/09修正: 対象0件でも「発送手続き待ち→発送手続き」への移動は必ず実行する。
        # 旧コードは target_order_nos が空だとCPaSS処理自体を丸ごとスキップしており、
        # 物理的な出荷キューが「発送手続き待ち」に滞留し続ける原因になっていた。
        if not dry_run:
            print()
            print("=" * 60)
            print("Step 2: CPaSS ワークフロー実行（1件ごとに途中保存 / 安全停止あり）")
            print("=" * 60)
            try:
                _cpass_final = cpass_workflow.process_all_orders_for_dhl(
                    target_order_nos=target_order_nos,
                    headless=True,
                    move_waiting=True,
                    on_order_done=_on_order_done,
                    deadline_ts=deadline_ts,
                    diagnostics=cpass_diag,
                )
                if _cpass_final:
                    # cpass_workflow側の最終dictで念のため同期（_on_order_doneで既に
                    # ほぼ同内容がcpass_resultsに入っているはずだが、取りこぼし防止の保険）。
                    cpass_results.update(_cpass_final)
                ok_count = sum(1 for v in cpass_results.values() if v.get("dhl_price_jpy"))
                summary_lines.append(f"CPaSS: {ok_count}/{len(target_order_nos)}件 DHL金額取得")
                if len(cpass_results) < len(target_order_nos):
                    remaining = len(target_order_nos) - len(cpass_results)
                    summary_lines.append(f"未処理のまま残: {remaining}件（次回自動継続）")
            except Exception as e:
                msg = f"CPaSS処理エラー: {e}"
                print(msg)
                errors.append(msg)
                cpass_error = str(e)[:200]

        # Step 4: 最終フラッシュ
        # CHECKPOINT_EVERY_N=1（既定）なら基本的にここで新規書き込みは発生しないはずだが、
        # CHECKPOINT_EVERY_N>1設定時の端数や、コールバック失敗時の取りこぼしに対する保険として残す。
        if checkpoint_state["pending_orders"] or checkpoint_state["last_error"]:
            _do_checkpoint("最終フラッシュ前の残チェックポイント")
        for prefix, (local_path, file_id) in xlsm_paths.items():
            try:
                ok, num_writes = process_xlsm(local_path, cpass_results, dry_run=dry_run)
                if ok and num_writes > 0 and not dry_run:
                    checkpoint_state["total_writes"] += num_writes
                    _upload_xlsm(service, file_id, local_path)
                    checkpoint_state["drive_uploads"] += 1
                    summary_lines.append(f"{prefix}: 最終フラッシュ {num_writes}セル 保存・UL完了")
                elif ok:
                    # ★2026/09/25修正: 書き込み0件の日も「保存・UL完了」と出ていて
                    #   正常完了と誤認させていたため、実際の更新有無で表示を分ける。
                    summary_lines.append(f"{prefix}: 最終フラッシュで追加書き込みなし")
            except Exception as e:
                msg = f"{prefix}の処理失敗: {e}"
                print(msg)
                errors.append(msg)

        summary_lines.append(
            f"途中保存(チェックポイント): 累計{checkpoint_state['saved_total']}件処理 "
            f"/ 安全停止デッドライン{SAFETY_BUDGET_SECONDS}秒"
        )

        # ─── ★2026/09/25追加: 実行結果の健全性チェック ───
        health = None
        if not dry_run:
            dhl_ok_orders = [o for o, v in cpass_results.items() if v.get("dhl_price_jpy")]
            prev_unmatched = load_prev_unmatched()
            health = evaluate_run_health(
                target_order_nos, cpass_diag, cpass_error, dhl_ok_orders,
                checkpoint_state["total_writes"], checkpoint_state["drive_uploads"], prev_unmatched)
            print(f"  [健全性チェック] 判定={health['level']} 新規対象={len(health['new_targets'])}件 "
                  f"CPaSS送料取得={len(dhl_ok_orders)}件 書込={checkpoint_state['total_writes']}セル "
                  f"Drive更新={checkpoint_state['drive_uploads']}回 移動={cpass_diag.get('move_status')}")
            for r in health["reasons"]:
                print("    [異常] " + r)
            for n in health["notes"]:
                print("    [注意] " + n)
            if _verify_orders_env:
                print("  [健全性チェック] 検証モードのため未処理リスト・異常マーカーは更新しない")
            else:
                if health["level"] != "anomaly":
                    # 正常/注意の日だけ基準を更新（異常の日は据え置き、翌日も同じ注文を検知し続ける）
                    save_unmatched([o for o in target_order_nos if o not in set(dhl_ok_orders)])
                if health["level"] == "anomaly":
                    with open(ANOMALY_MARKER_FILE, "w", encoding="utf-8") as f:
                        f.write("\n".join(health["reasons"]) + "\n")
            health_header = build_health_header(
                health, len(dhl_ok_orders), checkpoint_state["total_writes"],
                checkpoint_state["drive_uploads"], len(target_order_nos))
            print(health_header.rstrip())

    end_time = datetime.now()
    elapsed = int((end_time - start_time).total_seconds())

    print()
    print("=" * 60)
    print(f"完了  {end_time.strftime('%Y-%m-%d %H:%M:%S')}  ({elapsed}秒)")
    print("=" * 60)

    # メール送信（開始日時・完了日時・経過時間はsend_result_email側でJST表示する）
    results_text = "\n".join(summary_lines) if summary_lines else "処理完了（書き込みなし）"
    _prefix = ""
    _header = ""
    if health is not None:
        _header = health_header
        if health["level"] == "anomaly":
            _prefix = "【要確認・未処理】"
        elif health["level"] == "warning":
            _prefix = "【注意】"
    send_result_email(results_text, run_count, "\n".join(errors), start_time_jst=start_time_jst,
                      subject_prefix=_prefix, header_text=_header)


if __name__ == "__main__":
    main()
