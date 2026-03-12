"""
サンキャク 案件管理タスクアラートBot
Master_DBシートを読み取り、Dir別にタスク状況をGoogle Chatに通知する。

RUN_MODE:
  morning  - 朝の全体通知（@全員 + Dir別全アラート）
  followup - 12時/16時の超過フォロー（超過のみ + 管理者メンション）
"""

import os
import sys
from datetime import datetime, date, timedelta

import gspread
import requests
import google.auth

# ---------- 設定 ----------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Dir名 → (表示名, Google Chat ユーザーID)
DIR_INFO = {
    "Sho":       ("池上翔",   ""),  # TODO: Google Chat user ID
    "翔太":      ("小峰翔太", ""),  # TODO: Google Chat user ID
    "さとしゅん": ("佐藤竣",  ""),  # TODO: Google Chat user ID
}

# 管理者
MANAGERS = {
    "Sho":  ("池上翔",   ""),
    "Ryu":  ("竹内竜太", ""),
}

# 除外ステータス
SKIP_STATUSES = ["7. 納品/公開待", "9. 完了", "x. ペンディング"]

# ステータス停滞とみなす日数
STALE_DAYS = 5

# 空欄チェック対象カラム（他の案件と比べて空欄が目立つ場合に指摘）
REQUIRED_FIELDS = {
    "DEADLINE": (12, "締切/公開"),
    "EDIT_START": (13, "編集着手"),
    "DRAFT_PERIOD": (14, "初稿期間"),
    "DIR": (29, "Dir"),
    "EDITOR": (30, "Editor"),
}

# カラムインデックス (0-based)
COL = {
    "ID": 0,           # A
    "CLIENT": 1,       # B
    "PROJECT": 3,      # D
    "MEMO": 5,         # F
    "STATUS": 7,       # H
    "THUMB": 8,        # I
    "NEXT_DATE": 10,   # K
    "DEADLINE": 12,    # M
    "EDIT_START": 13,  # N
    "DRAFT_PERIOD": 14,# O
    "HAS_SHOOT": 15,   # P
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
    current = start_date
    added = 0
    while added < num_days:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return current


def mention(name, user_id, enable):
    if enable and user_id:
        return f"<users/{user_id}>"
    return name


# ---------- データ取得 ----------


def get_gspread_client():
    creds, _ = google.auth.default(scopes=SCOPES)
    return gspread.authorize(creds)


def fetch_sheet_data(gc):
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    sheet_name = os.environ["SHEET_NAME"]
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return [], []
    return all_values[0], all_values[1:]


# ---------- ステータス追跡 ----------


def load_tracking(gc):
    """トラッキングシートから前回のステータスを読み込む"""
    tracking_id = os.environ.get("TRACKING_SHEET_ID", "")
    if not tracking_id:
        return {}
    try:
        sh = gc.open_by_key(tracking_id)
        ws = sh.worksheet("status_log")
        rows = ws.get_all_values()
        tracking = {}
        for r in rows[1:]:
            if len(r) >= 3 and r[0]:
                tracking[r[0]] = {"status": r[1], "last_changed": r[2]}
        return tracking
    except Exception:
        return {}


def save_tracking(gc, current_statuses, prev_tracking, today):
    """現在のステータスをトラッキングシートに保存"""
    tracking_id = os.environ.get("TRACKING_SHEET_ID", "")
    if not tracking_id:
        return

    today_str = today.isoformat()
    rows = [["project_id", "status", "last_changed"]]

    for pid, status in current_statuses.items():
        prev = prev_tracking.get(pid, {})
        if prev.get("status") == status:
            last_changed = prev.get("last_changed", today_str)
        else:
            last_changed = today_str
        rows.append([pid, status, last_changed])

    try:
        sh = gc.open_by_key(tracking_id)
        ws = sh.worksheet("status_log")
        ws.clear()
        ws.update(range_name="A1", values=rows)
    except Exception as e:
        print(f"トラッキング保存エラー: {e}", file=sys.stderr)


def detect_stale(rows, prev_tracking, today):
    """ステータスが STALE_DAYS 以上変わっていない案件を検出"""
    stale = []
    for row in rows:
        pid = get_cell(row, COL["ID"])
        if not pid:
            continue
        status = get_cell(row, COL["STATUS"])
        if any(skip in status for skip in SKIP_STATUSES):
            continue

        prev = prev_tracking.get(pid, {})
        last_changed_str = prev.get("last_changed", "")
        if not last_changed_str:
            continue

        try:
            last_changed = date.fromisoformat(last_changed_str)
        except ValueError:
            continue

        days_stale = (today - last_changed).days
        if days_stale >= STALE_DAYS:
            client = get_cell(row, COL["CLIENT"])
            project = get_cell(row, COL["PROJECT"])
            dir_name = get_cell(row, COL["DIR"])
            stale.append({
                "client": client,
                "project": project,
                "status": status,
                "days": days_stale,
                "dir": dir_name,
            })

    return stale


# ---------- 空欄チェック ----------


def detect_blanks(rows):
    """アクティブ案件で重要フィールドが空欄の案件を検出"""
    blanks = []
    for row in rows:
        pid = get_cell(row, COL["ID"])
        if not pid:
            continue
        status = get_cell(row, COL["STATUS"])
        if any(skip in status for skip in SKIP_STATUSES):
            continue

        missing = []
        for field_key, (col_idx, label) in REQUIRED_FIELDS.items():
            val = get_cell(row, col_idx)
            if not val or val in ("未定", "-"):
                # 撮影なしの案件は編集着手チェックをスキップしない
                missing.append(label)

        if missing:
            client = get_cell(row, COL["CLIENT"])
            project = get_cell(row, COL["PROJECT"])
            dir_name = get_cell(row, COL["DIR"])
            blanks.append({
                "client": client,
                "project": project,
                "missing": missing,
                "dir": dir_name,
            })

    return blanks


# ---------- 案件分析 ----------


def analyze_project(row, today):
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

    # --- 撮影日 ---
    shoot_dates = []
    for col in [COL["SHOOT1"], COL["SHOOT2"], COL["SHOOT3"]]:
        d = parse_date(get_cell(row, col), today)
        if d:
            shoot_dates.append(d)

    if has_shoot == "あり" and shoot_dates:
        future_shoots = [d for d in shoot_dates if d >= today]
        all_shot_done = len(future_shoots) == 0

        for sd in future_shoots:
            diff = (sd - today).days
            if diff == 0:
                messages.append(("🚨", f"「{project}」の撮影は*本日*です", 0, True))
            elif diff <= 14:
                messages.append(("📅", f"「{project}」は撮影日まであと*{diff}日*です", diff, False))

        if all_shot_done and status in [
            "0. 未着手/相談中",
            "1. 企画/撮影準備",
            "2. 撮影済/素材待",
            "2.5. 0稿編集中",
        ]:
            latest_shoot = max(shoot_dates)
            days_since = (today - latest_shoot).days
            if days_since <= 14:
                messages.append((
                    "🎬",
                    f"「{project}」は撮影が終わったので、素材の展開と、サムネイルの発注、文言ぎめが必要です",
                    -1, True,
                ))

    # --- 初稿提出期限 ---
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
            messages.append(("💀", f"「{project}」は初稿提出期限を*{abs(diff)}日超過*しています", diff, True))
        elif diff == 0:
            messages.append(("🚨", f"「{project}」の初稿提出は*本日*です", 0, True))
        elif diff <= 10:
            messages.append(("📝", f"「{project}」は初稿提出まであと*{diff}日*", diff, False))

        thumb_done = thumb_status and ("完了" in thumb_status or "なし" in thumb_status)
        if not thumb_done:
            if diff < 0:
                messages.append(("💀", f"「{project}」サムネイル文言の提出期限を*{abs(diff)}日超過*しています", diff, True))
            elif diff == 0:
                messages.append(("🚨", f"「{project}」サムネイル文言の提出期限は*本日*です", 0, True))
            elif diff <= 10:
                messages.append(("📌", f"「{project}」サムネイル文言の提出期限まであと*{diff}日*です", diff, False))

    # --- 締切/公開 ---
    deadline = parse_date(get_cell(row, COL["DEADLINE"]), today)
    if deadline:
        diff = (deadline - today).days
        if diff < 0:
            messages.append(("💀", f"「{project}」の締切/公開を*{abs(diff)}日超過*しています", diff, True))
        elif diff == 0:
            messages.append(("🚨", f"「{project}」の締切/公開は*本日*です", 0, True))
        elif diff <= 14:
            messages.append(("⏰", f"「{project}」の締切/公開まであと*{diff}日*", diff, False))

    return messages


# ---------- Dir別集計 ----------


def build_dir_alerts(rows, today):
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


def format_morning(dir_alerts, stale_items, blank_items, today, enable_mentions):
    date_str = today.strftime("%Y/%m/%d")

    if not dir_alerts and not stale_items and not blank_items:
        return f"📋 サンキャク 本日のタスクアラート（{date_str}）\n\n✅ 本日のアラートはありません"

    header = f"📋 サンキャク 本日のタスクアラート（{date_str}）"
    if enable_mentions:
        header += "\n<users/all>"
    lines = [header]

    # --- Dir別アラート ---
    for dir_name in ["Sho", "翔太", "さとしゅん"]:
        if dir_name not in dir_alerts:
            continue

        display_name, user_id = DIR_INFO[dir_name]
        has_urgent = any(
            m[3] for msgs in dir_alerts[dir_name].values() for m in msgs
        )

        if has_urgent and enable_mentions and user_id:
            name_line = f"📣 <users/{user_id}>"
        else:
            name_line = f"📣 {display_name}さん"

        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append(name_line)
        lines.append("━━━━━━━━━━━━━━")

        for client, messages in dir_alerts[dir_name].items():
            messages.sort(key=lambda m: m[2])
            lines.append("")
            lines.append(f"▸ {client}")
            for emoji, msg, _, _ in messages:
                lines.append(f"  {emoji} {msg}")

    # --- ステータス停滞 ---
    if stale_items:
        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append("⏸ ステータス停滞中（5日以上更新なし）")
        lines.append("━━━━━━━━━━━━━━")
        for item in stale_items:
            lines.append(
                f"  ⚠️ {item['client']} /「{item['project']}」"
                f"（{item['status']}）*{item['days']}日間*変更なし"
            )

    # --- 空欄チェック ---
    if blank_items:
        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append("📝 記入漏れ（空欄項目あり）")
        lines.append("━━━━━━━━━━━━━━")
        for item in blank_items:
            missing_str = "、".join(item["missing"])
            lines.append(
                f"  ⚠️ {item['client']} /「{item['project']}」→ {missing_str}"
            )

    return "\n".join(lines)


def format_followup(dir_alerts, today, enable_mentions):
    date_str = today.strftime("%Y/%m/%d")

    overdue_alerts = {}
    for dir_name, clients in dir_alerts.items():
        for client, messages in clients.items():
            overdue_msgs = [m for m in messages if m[2] < 0]
            if overdue_msgs:
                if dir_name not in overdue_alerts:
                    overdue_alerts[dir_name] = {}
                overdue_alerts[dir_name][client] = overdue_msgs

    if not overdue_alerts:
        return None

    manager_mentions = [
        mention(disp, uid, enable_mentions)
        for _, (disp, uid) in MANAGERS.items()
    ]
    manager_line = " ".join(manager_mentions)

    lines = [f"🔔 超過案件フォローアップ（{date_str}）"]
    lines.append(f"cc: {manager_line}")

    for dir_name in ["Sho", "翔太", "さとしゅん"]:
        if dir_name not in overdue_alerts:
            continue

        display_name, user_id = DIR_INFO[dir_name]
        name_str = mention(display_name, user_id, enable_mentions)

        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append(f"📣 {name_str}")
        lines.append("━━━━━━━━━━━━━━")

        for client, messages in overdue_alerts[dir_name].items():
            messages.sort(key=lambda m: m[2])
            lines.append("")
            lines.append(f"▸ {client}")
            for emoji, msg, _, _ in messages:
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
    run_mode = os.environ.get("RUN_MODE", "morning")
    enable_mentions = os.environ.get("ENABLE_MENTIONS", "false").lower() == "true"

    print(f"実行日: {today} / モード: {run_mode} / メンション: {enable_mentions}")

    try:
        gc = get_gspread_client()
        header, rows = fetch_sheet_data(gc)
        if not header:
            if run_mode == "morning":
                send_to_google_chat(
                    f"📋 サンキャク 本日のタスクアラート（{today.strftime('%Y/%m/%d')}）\n\n✅ データがありません"
                )
            return

        dir_alerts = build_dir_alerts(rows, today)

        # ステータス追跡
        prev_tracking = load_tracking(gc)
        current_statuses = {}
        for row in rows:
            pid = get_cell(row, COL["ID"])
            status = get_cell(row, COL["STATUS"])
            if pid:
                current_statuses[pid] = status

        stale_items = detect_stale(rows, prev_tracking, today)
        blank_items = detect_blanks(rows) if run_mode == "morning" else []

        if run_mode == "morning":
            message = format_morning(dir_alerts, stale_items, blank_items, today, enable_mentions)
        else:
            message = format_followup(dir_alerts, today, enable_mentions)
            if message is None:
                print("超過案件なし → フォロー通知スキップ")
                save_tracking(gc, current_statuses, prev_tracking, today)
                return

        print("--- 通知メッセージ ---")
        print(message)
        print("---")

        send_to_google_chat(message)

        # トラッキング保存
        save_tracking(gc, current_statuses, prev_tracking, today)

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
