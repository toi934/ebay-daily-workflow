"""一時的なワンオフ修正スクリプト（2026/07/09）
7/9の通常・専門売上管理表に誤って書き込まれたDHL送料(全件同額10078円)を
空欄に戻す。原因はcpass_workflow.pyのDHL価格取得フォールバックのバグ（修正済み）。
このスクリプトは一度限りの修正用。実行後は削除して良い。
"""
import os
import sys
import json
import time

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import openpyxl

SCOPES = ["https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_NAME = "売上管理表"

sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
info = json.loads(sa_json)
if "private_key" in info:
    info["private_key"] = info["private_key"].replace(chr(92) + chr(110), chr(10))
creds = Credentials.from_service_account_info(info, scopes=SCOPES)
service = build("drive", "v3", credentials=creds)


def find_folder_id(name):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name)").execute()
    files = resp.get("files", [])
    if not files:
        raise RuntimeError("folder not found: " + name)
    return files[0]["id"]


def list_xlsm(folder_id, prefix):
    q = f"'{folder_id}' in parents and name contains '{prefix}' and name contains '.xlsm' and trashed=false"
    resp = service.files().list(q=q, fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc").execute()
    return resp.get("files", [])


folder_id = find_folder_id(DRIVE_FOLDER_NAME)
print("folder_id:", folder_id)

targets = {
    "通常": {"col": "BL", "orders": ["02-14885-63234", "06-14878-91869", "06-14879-84362",
                                       "08-14873-31181", "10-14870-71397", "10-14870-90100",
                                       "15-14861-06648", "22-14851-69632"]},
    "専門": {"col": "BR", "orders": ["20-14853-54694", "09-14873-02935", "27-14843-81350"]},
}

for prefix, cfg in targets.items():
    files = list_xlsm(folder_id, prefix)
    if not files:
        print(f"[ERROR] {prefix}: no files found")
        continue
    f = files[0]
    print(f"{prefix}: {f['name']} (id={f['id']}, modified={f['modifiedTime']})")
    local_path = f"/tmp/{prefix}_fix.xlsm"
    request = service.files().get_media(fileId=f["id"])
    with open(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    print(f"  downloaded to {local_path}")

    wb = openpyxl.load_workbook(local_path, keep_vba=True)
    ws = wb.active
    col_letter = cfg["col"]
    fixed = []
    skipped = []
    for row in range(2, ws.max_row + 1):
        order_val = ws.cell(row=row, column=2).value
        if order_val in cfg["orders"]:
            cell = ws[f"{col_letter}{row}"]
            old_val = cell.value
            if old_val == 10078:
                cell.value = None
                fixed.append((order_val, row, old_val))
            else:
                skipped.append((order_val, row, old_val))
    print(f"  fixed rows: {fixed}")
    print(f"  skipped (value was not 10078): {skipped}")
    missing_orders = set(cfg["orders"]) - set(o for o, _, _ in fixed) - set(o for o, _, _ in skipped)
    if missing_orders:
        print(f"  [WARN] orders not found in sheet at all: {missing_orders}")

    wb.save(local_path)
    print("  saved locally")

    last_err = None
    for attempt in range(1, 4):
        try:
            media = MediaFileUpload(local_path, mimetype="application/vnd.ms-excel.sheet.macroEnabled.12", resumable=True)
            req = service.files().update(fileId=f["id"], media_body=media)
            resp = None
            while resp is None:
                _, resp = req.next_chunk()
            print(f"  uploaded successfully (attempt {attempt})")
            last_err = None
            break
        except Exception as e:
            last_err = e
            print(f"  [WARN] upload failed attempt {attempt}: {str(e)[:150]}")
            time.sleep(5 * attempt)
    if last_err:
        print(f"  [ERROR] {prefix} upload failed after 3 attempts: {last_err}")

print("DONE")
