"""
サンキャク 案件管理タスクアラートBot
Master_DBシートを読み取り、Dir別にタスク状況をGoogle Chatに通知する。
"""

import os
import sys
from datetime import datetime, date, timedelta

import gspread
import requests
import google.auth

# ---------- 設定 ----------

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Dir名 → (表示名, Google Chat ユーザーID)
# ENABLE_MENTIONS=true で本番メンション有効化
DIR_INFO = {
    "Sho":       ("池上翔",   ""),  # TODO: Google Chat user ID
    "翔太":      ("小峰翔太", ""),  # TODO: Google Chat user ID
    "さとしゅん": ("佐藤竣",  ""),  # TODO: Google Chat user ID
}

# 除外ステータス
SKIP_STATUSES = ["7. 納品/公開待"]

# カラムインデックス (0-based)
COL = {
    "ID": 0,           # A
    "CLIENT": 1,       # B
    "PROJECT": 3,      # D
    "MEMO": 5,         # F - 次のタスク/メモ
    "STATUS": 7,       # H
    "THUMB": 8,        # I - サムネ進捗
    "NEXT_DATE": 10,   # K - 次の期日
    "DEADLINE": 12,    # M - 締切/公開
    "EDIT_START": 13,  # N - 編集着手
    "DRAFT_PERIOD": 14,# O - 初稿期間(営業日)
    "HAS_SHOOT": 15,   # P - 撮影有無
    "SHOOT1": 16,      # Q
    "SHOOT2": 20,      # U
    "SHOOT3": 24,      # Y
    "DIR": 29,         # AD
}


# ---------- ユーティリティ ----------


def get_cell(row, col_index):
    if col_index < len(row):
        return row[col_index].strip()
    return ""


def parse_date(value, today):
    """日付パース。年省略時は実行年を使用（過去日もそのまま）。"""
    if not value or not value.strip():
        return None
    value = value.strip().replace("-", "/")
    if value in ("なし", "未定", "-"):
        return None

    for fmt in ["%Y/%m/%d", "%m/%d"]:
        try:
            parsed = datetime.strptime(value, fmt).date()
            if fmt == "%m/%d":
                parsed = parsed.replace(year=today.year)
            return parsed
        except ValueError:
            continue
    return None


def add_business_days(start_date, num_days):
    """営業日を加算（土日除外）"""
    current = start_date
    added = 0
    while added < num_days:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return current


# ---------- データ取得 ----------


def fetch_sheet_data():
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    sheet_name = os.environ["SHEET_NAME"]

    creds, _ = google.auth.default(scopes=SCOPES)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    all_values = ws.get_all_values()

    if len(all_values) < 2:
        return [], []
    return all_values[0], all_values[1:]


# ---------- 案件分析 ----------


def analyze_project(row, today):
    """案件を分析してアラートメッセージのリストを返す。
    Returns: [(emoji, message, sort_priority)]
      sort_priority: 小さいほど緊急（負=超過, 0=本日, 正=残日数）
    """
    messages = []

    project_id = get_cell(row, COL["ID"])
    if not project_id:
        return messages

    status = get_cell(row, COL["STATUS"])
    if any(skip in status for skip in SKIP_STATUSES):
        return messages

    project = get_cell(row, COL["PROJECT"])
    has_shoot = get_cell(row, COL["HAS_SHOOT"])
    thumb_status = get_cell(row, COL["THUMB"])

    # --- 撮影日チェック ---
    shoot_dates = []
    for col in [COL["SHOOT1"], COL["SHOOT2"], COL["SHOOT3"]]:
        d = parse_date(get_cell(row, col), today)
        if d:
            shoot_dates.append(d)

    if has_shoot == "あり" and shoot_dates:
        future_shoots = [d for d in shoot_dates if d >= today]
        all_shot_done = len(future_shoots) == 0

        # 未来の撮影日アラート
        for sd in future_shoots:
            diff = (sd - today).days
            if diff == 0:
                messages.append(("🚨", f"「{project}」の撮影は*本日*です", 0))
            elif diff <= 14:
                messages.append(("📅", f"「{project}」は撮影日まであと*{diff}日*です", diff))

        # 撮影済み → 素材展開・サムネ発注・文言ぎめが必要
        if all_shot_done and status in [
            "0. 未着手/相談中",
            "1. 企画/撮影準備",
            "2.5. 0稿編集中",
        ]:
            latest_shoot = max(shoot_dates)
            days_since = (today - latest_shoot).days
            if days_since <= 14:
                messages.append((
                    "🎬",
                    f"「{project}」は撮影が終わったので、素材の展開と、サムネイルの発注、文言ぎめが必要です",
                    -1,
                ))

    # --- 初稿提出期限 ---
    # 先方確認以降は初稿済みなのでスキップ
    draft_done_statuses = ["4. 社内QC", "5. 先方確認", "6. 修正対応中"]
    draft_already_done = any(s in status for s in draft_done_statuses)

    edit_start = parse_date(get_cell(row, COL["EDIT_START"]), today)
    draft_period_str = get_cell(row, COL["DRAFT_PERIOD"])

    draft_deadline = None
    if edit_start and draft_period_str and not draft_already_done:
        try:
            draft_period = int(draft_period_str)
            draft_deadline = add_business_days(edit_start, draft_period)
        except ValueError:
            pass

    if draft_deadline:
        diff = (draft_deadline - today).days
        if diff < 0:
            messages.append(("💀", f"「{project}」は初稿提出期限を*{abs(diff)}日超過*しています", diff))
        elif diff == 0:
            messages.append(("🚨", f"「{project}」の初稿提出は*本日*です", 0))
        elif diff <= 10:
            messages.append(("📝", f"「{project}」は初稿提出まであと*{diff}日*", diff))

        # サムネイル文言（初稿と同じタイミング、サムネ未完了の場合）
        thumb_done = thumb_status and ("完了" in thumb_status or "なし" in thumb_status)
        if not thumb_done:
            if diff < 0:
                messages.append(("💀", f"「{project}」サムネイル文言の提出期限を*{abs(diff)}日超過*しています", diff))
            elif diff == 0:
                messages.append(("🚨", f"「{project}」サムネイル文言の提出期限は*本日*です", 0))
            elif diff <= 10:
                messages.append(("📌", f"「{project}」サムネイル文言の提出期限まであと*{diff}日*です", diff))

    # --- 締切/公開 ---
    deadline = parse_date(get_cell(row, COL["DEADLINE"]), today)
    if deadline:
        diff = (deadline - today).days
        if diff < 0:
            messages.append(("💀", f"「{project}」の締切/公開を*{abs(diff)}日超過*しています", diff))
        elif diff == 0:
            messages.append(("🚨", f"「{project}」の締切/公開は*本日*です", 0))
        elif diff <= 14:
            messages.append(("⏰", f"「{project}」の締切/公開まであと*{diff}日*", diff))

    return messages


# ---------- Dir別集計 ----------


def build_dir_alerts(rows, today):
    """Dir別 → クライアント別にアラートを構築"""
    dir_alerts = {}

    for row in rows:
        dir_name = get_cell(row, COL["DIR"])
        if not dir_name or dir_name == "未定":
            continue
        if dir_name not in DIR_INFO:
            continue

        client = get_cell(row, COL["CLIENT"])

        messages = analyze_project(row, today)

        if messages:
            if dir_name not in dir_alerts:
                dir_alerts[dir_name] = {}
            if client not in dir_alerts[dir_name]:
                dir_alerts[dir_name][client] = []
            dir_alerts[dir_name][client].extend(messages)

    return dir_alerts


# ---------- メッセージ整形 ----------


def format_message(dir_alerts, today):
    date_str = today.strftime("%Y/%m/%d")
    enable_mentions = os.environ.get("ENABLE_MENTIONS", "false").lower() == "true"

    if not dir_alerts:
        return f"📋 サンキャク 本日のタスクアラート（{date_str}）\n\n✅ 本日のアラートはありません"

    lines = [f"📋 サンキャク 本日のタスクアラート（{date_str}）"]

    for dir_name in ["Sho", "翔太", "さとしゅん"]:
        if dir_name not in dir_alerts:
            continue

        display_name, user_id = DIR_INFO[dir_name]
        if enable_mentions and user_id:
            mention = f"<users/{user_id}>"
        else:
            mention = f"{display_name}さん"
        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append(f"📣 {mention}")
        lines.append("━━━━━━━━━━━━━━")

        clients = dir_alerts[dir_name]
        for client, messages in clients.items():
            messages.sort(key=lambda m: m[2])
            lines.append("")
            lines.append(f"▸ {client}")
            for emoji, msg, _ in messages:
                lines.append(f"  {emoji} {msg}")

    return "\n".join(lines)


# ---------- 送信 ----------


def send_to_google_chat(message):
    webhook_url = os.environ["GOOGLE_CHAT_WEBHOOK_URL"]
    payload = {"text": message}
    resp = requests.post(webhook_url, json=payload, timeout=30)
    resp.raise_for_status()
    print(f"Google Chat送信完了 (status: {resp.status_code})")


# ---------- メイン ----------


def main():
    today = date.today()
    print(f"実行日: {today}")

    try:
        header, rows = fetch_sheet_data()
        if not header:
            send_to_google_chat(
                f"📋 サンキャク 本日のタスクアラート（{today.strftime('%Y/%m/%d')}）\n\n✅ データがありません"
            )
            return

        dir_alerts = build_dir_alerts(rows, today)
        message = format_message(dir_alerts, today)

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
