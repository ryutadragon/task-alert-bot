"""
サンキャク 案件管理タスクアラートBot
Googleスプレッドシートの案件管理シートを読み取り、
期限が近いタスクをGoogle Chatに通知する。
"""

import json
import os
import sys
from datetime import datetime, date

import gspread
import requests
import google.auth

# ---------- 設定 ----------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

# チェック対象の日付カラム定義 (0-indexed column, ラベル)
DATE_COLUMNS = [
    ("K", "次のタスク期日"),
    ("M", "締切/公開"),
    ("Q", "撮影日①"),
    ("U", "撮影日②"),
    ("Y", "撮影日③"),
]

# 除外ステータス
SKIP_STATUSES = ["6. 完了", "7. 納品/公開待", "x. なし"]

# アラート閾値
ALERT_THRESHOLDS = [
    (0, "🚨 本日"),
    (1, "🔥 明日"),
    (3, "⚠️ 3日後"),
]

# ---------- ユーティリティ ----------


def col_letter_to_index(letter: str) -> int:
    """列文字(A, B, ..., AA, AB...)を0-indexedの数値に変換"""
    result = 0
    for c in letter.upper():
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result - 1


def parse_date(value: str, today: date) -> date | None:
    """複数の日付形式をパースする。年省略時は実行年を補完し、過去なら翌年。"""
    if not value or not value.strip():
        return None

    value = value.strip().replace("-", "/")

    formats_with_year = ["%Y/%m/%d"]
    formats_without_year = ["%m/%d"]

    # まず年付きフォーマットを試行
    for fmt in formats_with_year:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    # 年なしフォーマットを試行
    for fmt in formats_without_year:
        try:
            parsed = datetime.strptime(value, fmt).date()
            parsed = parsed.replace(year=today.year)
            if parsed < today:
                parsed = parsed.replace(year=today.year + 1)
            return parsed
        except ValueError:
            continue

    return None


def get_cell(row: list[str], col_index: int) -> str:
    """行データから安全にセル値を取得"""
    if col_index < len(row):
        return row[col_index].strip()
    return ""


def find_role_columns(header: list[str]) -> dict[str, int]:
    """PM・Dir・Editorのヘッダー名を動的に検索して列インデックスを返す"""
    roles = {}
    for i, h in enumerate(header):
        h_lower = h.strip().lower()
        if "pm" in h_lower:
            roles["PM"] = i
        elif "dir" in h_lower:
            roles["Dir"] = i
        elif "editor" in h_lower or "編集" in h_lower:
            roles["Editor"] = i
    return roles


# ---------- アラート判定 ----------


def classify_alert(target_date: date, today: date) -> tuple[str, int, str] | None:
    """
    日付をアラート分類する。
    Returns: (カテゴリキー, ソート優先度, 表示テキスト) or None
    """
    diff = (target_date - today).days

    if diff < 0:
        return ("overdue", -1, f"💀 {abs(diff)}日超過")
    elif diff == 0:
        return ("today", 0, "🚨 本日期限")
    elif diff == 1:
        return ("tomorrow", 1, "🔥 明日期限")
    elif diff <= 3:
        return ("in3days", 3, "⚠️ 3日後期限")

    return None


# ---------- メイン処理 ----------


def fetch_sheet_data() -> tuple[list[str], list[list[str]]]:
    """スプレッドシートからデータを取得"""
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    sheet_name = os.environ["SHEET_NAME"]

    creds, _ = google.auth.default(scopes=SCOPES)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    all_values = ws.get_all_values()

    if len(all_values) < 2:
        return [], []

    header = all_values[0]
    rows = all_values[1:]
    return header, rows


def build_alerts(header: list[str], rows: list[list[str]], today: date) -> list[dict]:
    """全行・全日付カラムをチェックしてアラートリストを生成"""
    role_cols = find_role_columns(header)
    alerts = []

    col_id = col_letter_to_index("A")
    col_client = col_letter_to_index("B")
    col_project = col_letter_to_index("D")
    col_memo = col_letter_to_index("F")
    col_status = col_letter_to_index("H")

    for row in rows:
        row_id = get_cell(row, col_id)
        if not row_id:
            continue

        status = get_cell(row, col_status)
        if any(skip in status for skip in SKIP_STATUSES):
            continue

        client = get_cell(row, col_client)
        project = get_cell(row, col_project)
        memo = get_cell(row, col_memo) or "（メモなし）"

        # 担当者情報
        role_parts = []
        for role_name, col_idx in role_cols.items():
            person = get_cell(row, col_idx)
            role_parts.append(f"{role_name}: {person or '—'}")
        role_text = "｜".join(role_parts) if role_parts else ""

        for col_letter, label in DATE_COLUMNS:
            col_idx = col_letter_to_index(col_letter)
            date_str = get_cell(row, col_idx)
            target_date = parse_date(date_str, today)
            if target_date is None:
                continue

            result = classify_alert(target_date, today)
            if result is None:
                continue

            category, priority, badge = result

            alerts.append({
                "category": category,
                "priority": priority,
                "badge": badge,
                "client": client,
                "project": project,
                "label": label,
                "memo": memo,
                "role_text": role_text,
                "diff_days": (target_date - today).days,
            })

    # ソート: 超過(日数大きい順) → 本日 → 明日 → 3日後
    alerts.sort(key=lambda a: (a["priority"], a["diff_days"]))
    return alerts


def format_message(alerts: list[dict], today: date) -> str:
    """アラートリストを通知メッセージに整形"""
    date_str = today.strftime("%Y/%m/%d")

    if not alerts:
        return f"📋 サンキャク 本日のタスクアラート（{date_str}）\n\n✅ 本日のアラートはありません"

    # カテゴリ別にグループ化
    categories = {
        "overdue": ("💀 期限超過中", []),
        "today": ("🚨 本日期限", []),
        "tomorrow": ("🔥 明日期限", []),
        "in3days": ("⚠️ 3日後期限", []),
    }

    for alert in alerts:
        cat = alert["category"]
        if cat in categories:
            categories[cat][1].append(alert)

    lines = [f"📋 サンキャク 本日のタスクアラート（{date_str}）"]

    for cat_key in ["overdue", "today", "tomorrow", "in3days"]:
        title, items = categories[cat_key]
        if not items:
            continue

        lines.append("")
        lines.append(title)

        for a in items:
            suffix = ""
            if a["category"] == "overdue":
                suffix = f" ※{abs(a['diff_days'])}日超過"

            lines.append(
                f"・{a['client']} / {a['project']}　{a['label']}{suffix}"
            )
            lines.append(f"  → {a['memo']}")
            if a["role_text"]:
                lines.append(f"     {a['role_text']}")

    return "\n".join(lines)


def send_to_google_chat(message: str):
    """Google Chat Webhookに送信"""
    webhook_url = os.environ["GOOGLE_CHAT_WEBHOOK_URL"]
    payload = {"text": message}
    resp = requests.post(webhook_url, json=payload, timeout=30)
    resp.raise_for_status()
    print(f"Google Chat送信完了 (status: {resp.status_code})")


def main():
    today = date.today()
    print(f"実行日: {today}")

    try:
        header, rows = fetch_sheet_data()
        if not header:
            print("シートにデータがありません")
            send_to_google_chat(
                f"📋 サンキャク 本日のタスクアラート（{today.strftime('%Y/%m/%d')}）\n\n✅ 本日のアラートはありません"
            )
            return

        alerts = build_alerts(header, rows, today)
        message = format_message(alerts, today)

        print("--- 通知メッセージ ---")
        print(message)
        print("---")

        send_to_google_chat(message)

    except Exception as e:
        error_msg = f"⚠️ Bot実行エラー：{e}"
        print(error_msg, file=sys.stderr)
        try:
            send_to_google_chat(error_msg)
        except Exception as send_err:
            print(f"エラー通知の送信にも失敗: {send_err}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
