"""Google Drive の直売所売上フォルダから当日分CSVをダウンロードする。

サービスアカウントJSONは環境変数 GOOGLE_SERVICE_ACCOUNT_JSON から読み込む。
対象フォルダIDは環境変数 DRIVE_FOLDER_ID から読み込む。
"""

import io
import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

BASE_DIR = Path(__file__).parent
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def get_drive_service():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        print("[エラー] GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません", file=sys.stderr)
        sys.exit(1)
    info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def download_csvs(folder_id: str):
    service = get_drive_service()
    query = f"'{folder_id}' in parents and mimeType='text/csv' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])

    if not files:
        print("[警告] CSVファイルが見つかりませんでした")
        return

    for f in files:
        dest = BASE_DIR / f["name"]
        request = service.files().get_media(fileId=f["id"])
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        dest.write_bytes(buf.getvalue())
        print(f"[OK] {f['name']} ({len(buf.getvalue()):,} bytes)")


def main():
    folder_id = os.environ.get("DRIVE_FOLDER_ID")
    if not folder_id:
        print("[エラー] DRIVE_FOLDER_ID が設定されていません", file=sys.stderr)
        sys.exit(1)
    download_csvs(folder_id)


if __name__ == "__main__":
    main()
