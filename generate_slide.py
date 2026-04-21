"""
Google Slides に直売所売上レポートスライドを生成・更新する。

環境変数:
  GOOGLE_SERVICE_ACCOUNT_JSON : サービスアカウントJSON文字列
  SLIDES_PRESENTATION_ID      : 更新対象のプレゼンテーションID

使い方:
  send_report.py から generate_slide.main(targets, has_product, report_date) を呼ぶ。
"""

import datetime
import json
import os
import sys
import uuid

from google.oauth2 import service_account
from googleapiclient.discovery import build

import generate_report

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]
FONT = "Noto Sans JP"

# ── スライドサイズ (16:9 = 10" × 5.625" = 720pt × 405pt) ──
W, H = 720, 405
HEADER_H = 44
BODY_Y = HEADER_H + 8
BODY_H = H - BODY_Y - 10
L_X, L_W = 10, 305
R_X, R_W = 330, 380

WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]


# ── ユーティリティ ──

def pt(v):
    """PT → EMU 変換"""
    return int(v * 12700)


def rgb(hex_str):
    """#rrggbb → Slides API RGBColor dict"""
    h = hex_str.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


C_NAVY      = rgb("1c3f60")
C_WHITE     = rgb("ffffff")
C_OVER      = rgb("2a9658")
C_NEAR      = rgb("d4860f")
C_UNDER     = rgb("c94040")
C_BAR_BG    = rgb("f0f0f0")
C_TARGET    = rgb("aaaaaa")
C_GRAY      = rgb("999999")
C_DARK      = rgb("333333")
C_LIGHT     = rgb("7db3d6")
C_HIGHLIGHT = rgb("edf2f7")


def c_pct(p):
    return C_OVER if p >= 100 else C_NEAR if p >= 85 else C_UNDER


def nid():
    return "e" + uuid.uuid4().hex[:12]


def _transform(x, y):
    return {
        "scaleX": 1, "scaleY": 1, "shearX": 0, "shearY": 0,
        "translateX": pt(x), "translateY": pt(y), "unit": "EMU",
    }


def _size(w, h):
    return {
        "width":  {"magnitude": pt(w), "unit": "EMU"},
        "height": {"magnitude": pt(h), "unit": "EMU"},
    }


def mk_rect(pid, x, y, w, h, fill=None):
    oid = nid()
    reqs = [{"createShape": {
        "objectId": oid, "shapeType": "RECTANGLE",
        "elementProperties": {
            "pageObjectId": pid,
            "size": _size(w, h),
            "transform": _transform(x, y),
        },
    }}]
    sp = {"outline": {"propertyState": "NOT_RENDERED"}}
    if fill:
        sp["shapeBackgroundFill"] = {"solidFill": {"color": {"rgbColor": fill}}}
    else:
        sp["shapeBackgroundFill"] = {"propertyState": "NOT_RENDERED"}
    reqs.append({"updateShapeProperties": {
        "objectId": oid,
        "shapeProperties": sp,
        "fields": "shapeBackgroundFill,outline",
    }})
    return reqs


def mk_text(pid, x, y, w, h, text, size, color,
            bold=False, align="LEFT", valign="MIDDLE", bg=None):
    oid = nid()
    reqs = [
        {"createShape": {
            "objectId": oid, "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": pid,
                "size": _size(w, h),
                "transform": _transform(x, y),
            },
        }},
        {"insertText": {"objectId": oid, "insertionIndex": 0, "text": text}},
        {"updateTextStyle": {
            "objectId": oid,
            "style": {
                "bold": bold,
                "fontSize": {"magnitude": size, "unit": "PT"},
                "foregroundColor": {"opaqueColor": {"rgbColor": color}},
                "fontFamily": FONT,
            },
            "fields": "bold,fontSize,foregroundColor,fontFamily",
        }},
        {"updateParagraphStyle": {
            "objectId": oid,
            "style": {"alignment": align, "lineSpacing": 100},
            "fields": "alignment,lineSpacing",
        }},
    ]
    sp = {"contentAlignment": valign, "outline": {"propertyState": "NOT_RENDERED"}}
    if bg:
        sp["shapeBackgroundFill"] = {"solidFill": {"color": {"rgbColor": bg}}}
    else:
        sp["shapeBackgroundFill"] = {"propertyState": "NOT_RENDERED"}
    reqs.append({"updateShapeProperties": {
        "objectId": oid,
        "shapeProperties": sp,
        "fields": "contentAlignment,outline,shapeBackgroundFill",
    }})
    return reqs


def mk_bar(pid, x, y, w, h, pct):
    """進捗バー（背景・塗り・100%縦線）"""
    reqs = mk_rect(pid, x, y, w, h, C_BAR_BG)
    fill_w = min(pct / 120 * w, w)
    if fill_w > 0.5:
        reqs += mk_rect(pid, x, y, fill_w, h, c_pct(pct))
    line_x = x + w * 100 / 120
    reqs += mk_rect(pid, line_x - 0.5, y - 1.5, 1, h + 3, C_TARGET)
    return reqs


# ── サービス取得 ──

def get_service():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        print("[エラー] GOOGLE_SERVICE_ACCOUNT_JSON 未設定", file=sys.stderr)
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json), scopes=SCOPES
    )
    return build("slides", "v1", credentials=creds)


# ── メイン ──

def main(targets, has_product, report_date):
    prs_id = os.environ.get("SLIDES_PRESENTATION_ID")
    if not prs_id:
        print("[スキップ] SLIDES_PRESENTATION_ID 未設定のためスライド生成をスキップ")
        return

    service = get_service()

    # プレゼンテーション取得
    prs = service.presentations().get(presentationId=prs_id).execute()
    slide = prs["slides"][0]
    pid = slide["objectId"]

    # 既存要素を全削除
    elems = [el["objectId"] for el in slide.get("pageElements", [])]
    if elems:
        service.presentations().batchUpdate(
            presentationId=prs_id,
            body={"requests": [{"deleteObject": {"objectId": oid}} for oid in elems]},
        ).execute()

    # データ準備
    week_start = report_date - datetime.timedelta(days=report_date.weekday())
    week_data = targets[
        (targets["日付"] >= week_start)
        & (targets["日付"] <= report_date)
        & targets["実績"].notna()
    ].copy()
    month_data = targets[
        (targets["日付"] >= report_date.replace(day=1))
        & (targets["日付"] <= report_date)
        & targets["実績"].notna()
    ].copy()

    reqs = []

    # ── ヘッダー ──
    reqs += mk_rect(pid, 0, 0, W, HEADER_H, C_NAVY)
    reqs += mk_text(pid, 12, 0, 300, HEADER_H,
                    "直売所本店 売上レポート", 14, C_WHITE, bold=True)
    wd = WEEKDAYS[report_date.weekday()]
    reqs += mk_text(pid, 320, 0, W - 332, HEADER_H,
                    report_date.strftime("%Y/%m/%d") + f"（{wd}）",
                    10, C_LIGHT, align="RIGHT")

    # ── 左列：日次達成率 ──
    reqs += mk_text(pid, L_X, BODY_Y, L_W, 13, "日次達成率（売上全体）", 7, C_GRAY)

    n = len(week_data)
    ROW_H = (BODY_H - 15) / max(n, 1)
    BAR_H = 5

    for i, (_, row) in enumerate(week_data.iterrows()):
        dt = row["日付"]
        wd_r = WEEKDAYS[dt.weekday()]
        actual = row["実績"]
        target_v = row["日次合計"]
        pct = actual / target_v * 100 if target_v > 0 else 0
        is_last = (dt == report_date)
        ry = BODY_Y + 15 + i * ROW_H

        if is_last:
            reqs += mk_rect(pid, 0, ry, L_W + L_X + 5, ROW_H - 1, C_HIGHLIGHT)

        reqs += mk_text(pid, L_X, ry + 1, 68, ROW_H - BAR_H - 3,
                        f"{dt.month}/{dt.day}（{wd_r}）",
                        8 if is_last else 7.5,
                        C_NAVY if is_last else C_DARK, bold=is_last)
        reqs += mk_text(pid, L_X + 70, ry + 1, L_W - 120, ROW_H - BAR_H - 3,
                        f"{actual:,.0f} / {target_v:,.0f}円", 7, C_GRAY)
        reqs += mk_text(pid, L_X + L_W - 44, ry + 1, 44, ROW_H - BAR_H - 3,
                        f"{pct:.0f}%",
                        9 if is_last else 8,
                        c_pct(pct), bold=is_last, align="RIGHT")
        reqs += mk_bar(pid, L_X, ry + ROW_H - BAR_H - 2, L_W - L_X, BAR_H, pct)

    # ── 右列 ──
    ry = float(BODY_Y)

    # 週間累計
    wa = float(week_data["実績"].sum())
    wt = float(week_data["日次合計"].sum())
    wp = wa / wt * 100 if wt > 0 else 0
    reqs += mk_text(pid, R_X, ry, R_W, 13, "週間累計", 7, C_GRAY)
    ry += 14
    reqs += mk_text(pid, R_X, ry, 64, 28, f"{wp:.0f}%", 20, c_pct(wp), bold=True)
    reqs += mk_text(pid, R_X + 66, ry, R_W - 66, 14,
                    f"実績  {wa:,.0f}円", 9, C_DARK, bold=True)
    reqs += mk_text(pid, R_X + 66, ry + 14, R_W - 66, 13,
                    f"目標  {wt:,.0f}円", 8, C_GRAY)
    ry += 30
    reqs += mk_bar(pid, R_X, ry, R_W, 7, wp)
    ry += 14

    # カテゴリー別（週間）
    reqs += mk_text(pid, R_X, ry, R_W, 13, "週間（カテゴリー別）", 7, C_GRAY)
    ry += 14

    cats = list(zip(generate_report.CATEGORIES, generate_report.CATEGORY_DISPLAY))
    cat_area_h = H - ry - 10 - 68   # 月間累計用に68pt確保
    cat_row_h = cat_area_h / len(cats)

    for cat, disp in cats:
        col = f"{cat}_実績"
        ca = float(week_data[col].sum()) if has_product and col in week_data.columns else 0.0
        ct = float(week_data[cat].sum()) if cat in week_data.columns else 0.0
        cp = ca / ct * 100 if ct > 0 else 0

        reqs += mk_text(pid, R_X, ry, 62, cat_row_h - 2, disp, 7, C_DARK)
        reqs += mk_text(pid, R_X + 64, ry, R_W - 108, cat_row_h - 2,
                        f"{ca:,.0f} / {ct:,.0f}円", 6.5, C_GRAY)
        reqs += mk_text(pid, R_X + R_W - 42, ry, 42, cat_row_h - 2,
                        f"{cp:.0f}%", 7, c_pct(cp), bold=True, align="RIGHT")
        bar_h = max(cat_row_h - 11, 3)
        reqs += mk_bar(pid, R_X, ry + cat_row_h - bar_h - 1, R_W, bar_h, cp)
        ry += cat_row_h

    # 月間累計
    ry += 6
    ma = float(month_data["実績"].sum())
    mt = float(month_data["日次合計"].sum())
    mp = ma / mt * 100 if mt > 0 else 0
    de = len(month_data)

    reqs += mk_text(pid, R_X, ry, R_W, 13,
                    f"月間累計（4/1〜{report_date.month}/{report_date.day}）", 7, C_GRAY)
    ry += 14
    reqs += mk_text(pid, R_X, ry, 64, 26, f"{mp:.0f}%", 18, c_pct(mp), bold=True)
    reqs += mk_text(pid, R_X + 66, ry, R_W - 66, 14,
                    f"実績  {ma:,.0f}円", 9, C_DARK, bold=True)
    reqs += mk_text(pid, R_X + 66, ry + 14, R_W - 66, 12,
                    f"目標  {mt:,.0f}円　{de}日経過 / 30日", 8, C_GRAY)
    ry += 28
    reqs += mk_bar(pid, R_X, ry, R_W, 7, mp)

    # batchUpdate 実行
    service.presentations().batchUpdate(
        presentationId=prs_id,
        body={"requests": reqs},
    ).execute()
    print(f"[OK] スライド更新完了 ({len(reqs)} requests)")


if __name__ == "__main__":
    print("generate_slide.py は send_report.py から呼び出してください。")
