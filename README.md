# kpi-sales-report

直売所本店の日次売上レポートをGmailで自動配信するシステム。

## 構成

| ファイル | 役割 |
|---|---|
| `generate_report.py` | Google Drive上のCSVと目標ファイルから日次/週間/月間の集計を行い、HTMLとサマリー辞書を生成 |
| `send_report.py` | 集計サマリーを本文にしたテキストメールをGmail SMTP経由で送信 |
| `4月日別売上目標.xlsx` | 2026年4月の商品別日別売上目標（静的データ） |

## 実行環境

GitHub Actions の cron スケジュール（毎日 12:00 UTC = 21:00 JST）で自動実行。

## 必要なSecrets（GitHub Settings → Secrets and variables → Actions）

| 名前 | 用途 |
|---|---|
| `GMAIL_APP_PASS` | Gmailアプリパスワード（16文字） |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Google Drive API用サービスアカウントのJSONキー全文 |

## ローカル手動実行

```bash
pip install -r requirements.txt
export GMAIL_APP_PASS="xxxx xxxx xxxx xxxx"
python send_report.py
```
