"""
毎日21:00に実行：
 1. generate_report.py で CSV → HTML と集計サマリーを生成
 2. 本日/週間/月間の数値要約 + 詳細URL を本文にした Gmail を送信

添付なし・プレーンテキスト本文のみ。

使い方:
  1. 下の CONFIG ブロックに送信元Gmail・アプリパスワード・送信先を設定
  2. Windowsタスクスケジューラで毎日21:00にこのスクリプトを実行
"""

import datetime
import glob
import os
import smtplib
import sys
from email.header import Header
from email.mime.text import MIMEText
from pathlib import Path

import generate_report

# ==============================================================
# CONFIG（ここを編集してください）
# ==============================================================

GMAIL_ADDRESS  = "ryo.a.yamada@gmail.com"           # 送信元Gmailアドレス
GMAIL_APP_PASS = os.environ.get("GMAIL_APP_PASS")   # 環境変数から読み込み（Step 3で設定）
TO_ADDRESSES   = [                                   # 送信先（複数可）
    "ryo.yamada@visional.inc",
    "asuka.sarana@gmail.com",
    "maruyama2899m@gmail.com",
    "ryo.a.yamada@gmail.com",
]

BASE_DIR = Path(__file__).parent
CSV_DIR  = BASE_DIR

# ==============================================================


def find_csv() -> Path | None:
    """日別売上 or 商品別売上 or 売上集計 CSVを探す"""
    all_patterns = [
        str(CSV_DIR / "日別売上(年月*).csv"),
        str(CSV_DIR / "売上集計_*.csv"),
        str(CSV_DIR / "商品別売上_*.csv"),
        str(CSV_DIR / "商品別売上(期間*).csv"),
    ]
    for pat in all_patterns:
        files = glob.glob(pat)
        if files:
            latest = Path(sorted(files)[-1])
            print(f"[OK] CSV確認: {latest.name}")
            return latest
    return None


def build_body(summary: dict, report_url: str = "") -> str:
    """数値要約 + 詳細URL のプレーンテキスト本文を組み立てる"""
    rd = summary["report_date"]
    wd = summary["weekday"]
    t, w, m = summary["today"], summary["week"], summary["month"]

    lines = [
        f"【直売所 日次売上レポート】{rd.strftime('%Y/%m/%d')}（{wd}）",
    ]
    if report_url:
        lines.append(f"詳細レポート: {report_url}")
    lines += [
        "",
        "■本日",
        f" 実績 {t['actual']:>10,.0f}円 / 目標 {t['target']:>10,.0f}円 （{t['pct']:.0f}%）",
        "",
        f"■週間累計（{w['start'].month}/{w['start'].day}〜{w['end'].month}/{w['end'].day}）",
        f" 実績 {w['actual']:>10,.0f}円 / 目標 {w['target']:>10,.0f}円 （{w['pct']:.0f}%）",
        "",
        f"■月間累計（4/1〜{rd.month}/{rd.day}）",
        f" 実績 {m['actual']:>10,.0f}円 / 目標 {m['target']:>10,.0f}円 （{m['pct']:.0f}%）",
        f" {m['days_elapsed']}日経過 / 30日",
    ]
    return "\n".join(lines)


def send_gmail(subject: str, body: str) -> bool:
    msg = MIMEText(body, "plain", "utf-8")
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = ", ".join(TO_ADDRESSES)
    msg["Subject"] = Header(subject, "utf-8")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASS)
            smtp.sendmail(GMAIL_ADDRESS, TO_ADDRESSES, msg.as_string())
        print(f"[OK] メール送信成功 → {TO_ADDRESSES}")
        return True
    except Exception as e:
        print(f"[エラー] メール送信失敗: {e}")
        return False


def main():
    today_str = datetime.date.today().strftime("%Y/%m/%d")
    print(f"=== 直売所売上レポート送信 {today_str} ===")

    if not GMAIL_APP_PASS:
        print("[エラー] 環境変数 GMAIL_APP_PASS が設定されていません。")
        print("  PowerShell で以下を実行してから再試行してください：")
        print('  [Environment]::SetEnvironmentVariable("GMAIL_APP_PASS", "xxxx xxxx xxxx xxxx", "User")')
        sys.exit(1)

    if find_csv() is None:
        print("[エラー] 実績CSVが見つかりません。処理を中断します。")
        sys.exit(1)

    try:
        html_path, summary = generate_report.main()
    except SystemExit:
        raise
    except Exception as e:
        print(f"[エラー] レポート生成失敗: {e}")
        sys.exit(1)

    report_url = os.environ.get("REPORT_URL", "")

    rd = summary["report_date"]
    subject = f"【直売所 日次売上レポート】{rd.strftime('%Y/%m/%d')}"
    body = build_body(summary, report_url)
    print("--- 本文プレビュー ---")
    print(body)
    print("----------------------")

    if not send_gmail(subject, body):
        sys.exit(1)

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
