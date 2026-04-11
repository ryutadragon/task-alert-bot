"""
サンキャク 案件管理タスクアラートBot
Master_DBシートを読み取り、Dir別にタスク状況をGoogle Chat + Slackに通知する。

RUN_MODE:
  morning  - 朝の全体通知（@全員 + Dir別全アラート）
  followup - 12時/16時の超過フォロー（超過のみ + 管理者メンション）
"""

import os
import sys
from datetime import datetime, date

import gspread
import jpholiday
import requests
import google.auth
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# ---------- 設定 ----------

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Dir名 → (表示名, Google Chat ユーザーID, Slack ユーザーID)
DIR_INFO = {
    "Sho":       ("池上翔",   "103766992526420785455", "U01TT135Y4A"),
    "翔太":      ("小峰翔太", "112110647990602264741", "U01TPV9MCVC"),
    "さとしゅん": ("佐藤竣",  "105169517756611623614", "U06S2LPTCQH"),
}

# 管理者
MANAGERS = {
    "Sho":  ("池上翔",   "103766992526420785455", "U01TT135Y4A"),
    "Ryu":  ("竹内竜太", "103825177855643919422", "U016ZDVNGJ0"),
}

# 除外ステータス
SKIP_STATUSES = ["7. 納品/公開待", "9. 完了", "x. ペンディング"]

# ステータス停滞とみなす日数
STALE_DAYS = 5

# カラムインデックス (0-based) — 案件管理シートの実際の列順に対応
COL = {
    "ID": 0,           # A: ID
    "CLIENT": 1,       # B: クライアント
    "PROJECT": 5,      # F: 案件名
    "DEADLINE": 7,     # H: 公開
    "MEMO": 8,         # I: 次のタスク/メモ
    "STATUS": 9,       # J: ステータス
    "THUMB": 10,       # K: サムネ進捗
    "PM": 12,          # M: PM
    "DIR": 13,         # N: Dir
    "EDITOR": 14,      # O: Editor
    "HAS_SHOOT": 19,   # T: 撮影有無
    "SHOOT1": 20,      # U: 撮影日1
    "SHOOT2": 24,      # Y: 撮影日2
    "SHOOT3": 28,      # AC: 撮影日3
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


def mention_gchat(name, user_id, enable):
    if enable and user_id:
        return f"<users/{user_id}>"
    return name


def mention_slack(name, slack_id, enable):
    if enable and slack_id:
        return f"<@{slack_id}>"
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


def _format_morning_core(dir_alerts, stale_items, today, enable_mentions, platform):
    """morning メッセージ生成。platform: 'gchat' or 'slack'"""
    date_str = today.strftime("%Y/%m/%d")
    m = mention_gchat if platform == "gchat" else mention_slack
    id_idx = 1 if platform == "gchat" else 2

    if not dir_alerts and not stale_items:
        return f"📋 サンキャク 本日のタスクアラート（{date_str}）\n\n✅ 本日のアラートはありません"

    header = f"📋 サンキャク 本日のタスクアラート（{date_str}）"
    if enable_mentions:
        header += "\n<users/all>" if platform == "gchat" else "\n<!channel>"

    any_urgent = any(
        msg[3]
        for clients in dir_alerts.values()
        for msgs in clients.values()
        for msg in msgs
    )
    ryu_info = MANAGERS["Ryu"]
    if any_urgent and enable_mentions:
        header += f"\ncc: {m(ryu_info[0], ryu_info[id_idx], True)}"

    lines = [header]

    for dir_name in ["Sho", "翔太", "さとしゅん"]:
        if dir_name not in dir_alerts:
            continue

        info = DIR_INFO[dir_name]
        display_name, uid = info[0], info[id_idx]
        has_urgent = any(
            msg[3] for msgs in dir_alerts[dir_name].values() for msg in msgs
        )

        if has_urgent and enable_mentions and uid:
            name_line = f"📣 {m(display_name, uid, True)}"
        else:
            name_line = f"📣 {display_name}さん"

        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append(name_line)
        lines.append("━━━━━━━━━━━━━━")

        for client, messages in dir_alerts[dir_name].items():
            messages.sort(key=lambda x: x[2])
            lines.append("")
            lines.append(f"▸ {client}")
            for emoji, msg, _, _ in messages:
                lines.append(f"  {emoji} {msg}")

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

    return "\n".join(lines)


def _format_followup_core(dir_alerts, today, enable_mentions, platform):
    """followup メッセージ生成。platform: 'gchat' or 'slack'"""
    date_str = today.strftime("%Y/%m/%d")
    m = mention_gchat if platform == "gchat" else mention_slack
    id_idx = 1 if platform == "gchat" else 2

    overdue_alerts = {}
    for dir_name, clients in dir_alerts.items():
        for client, messages in clients.items():
            overdue_msgs = [msg for msg in messages if msg[2] < 0]
            if overdue_msgs:
                if dir_name not in overdue_alerts:
                    overdue_alerts[dir_name] = {}
                overdue_alerts[dir_name][client] = overdue_msgs

    if not overdue_alerts:
        return None

    manager_mentions = [
        m(info[0], info[id_idx], enable_mentions)
        for _, info in MANAGERS.items()
    ]
    manager_line = " ".join(manager_mentions)

    lines = [f"🔔 超過案件フォローアップ（{date_str}）"]
    lines.append(f"cc: {manager_line}")

    for dir_name in ["Sho", "翔太", "さとしゅん"]:
        if dir_name not in overdue_alerts:
            continue

        info = DIR_INFO[dir_name]
        display_name, uid = info[0], info[id_idx]
        name_str = m(display_name, uid, enable_mentions)

        lines.append("")
        lines.append("━━━━━━━━━━━━━━")
        lines.append(f"📣 {name_str}")
        lines.append("━━━━━━━━━━━━━━")

        for client, messages in overdue_alerts[dir_name].items():
            messages.sort(key=lambda x: x[2])
            lines.append("")
            lines.append(f"▸ {client}")
            for emoji, msg, _, _ in messages:
                lines.append(f"  {emoji} {msg}")

    return "\n".join(lines)


def format_morning(dir_alerts, stale_items, today, enable_mentions):
    return _format_morning_core(dir_alerts, stale_items, today, enable_mentions, "gchat")


def format_followup(dir_alerts, today, enable_mentions):
    return _format_followup_core(dir_alerts, today, enable_mentions, "gchat")


def format_morning_slack(dir_alerts, stale_items, today, enable_mentions):
    return _format_morning_core(dir_alerts, stale_items, today, enable_mentions, "slack")


def format_followup_slack(dir_alerts, today, enable_mentions):
    return _format_followup_core(dir_alerts, today, enable_mentions, "slack")


# ---------- 送信 ----------


def send_to_google_chat(message):
    print("Google Chat送信は無効化されています")
    return


SLACK_CHANNEL_ID = "C0AR6CE7213"  # #制作進捗管理


def is_slack_silent_day(today):
    if today.weekday() >= 5:
        return True, "土日"
    if jpholiday.is_holiday(today):
        return True, f"祝日（{jpholiday.is_holiday_name(today)}）"
    return False, ""


def send_to_slack(message):
    token = os.environ.get("SLACK_BOT_TOKEN")
    if not token:
        print("SLACK_BOT_TOKEN 未設定 → Slack送信スキップ")
        return
    silent, reason = is_slack_silent_day(date.today())
    if silent:
        print(f"Slack送信スキップ（{reason}）")
        return
    client = WebClient(token=token)
    resp = client.chat_postMessage(channel=SLACK_CHANNEL_ID, text=message)
    print(f"Slack送信完了 (channel: {SLACK_CHANNEL_ID}, ts: {resp['ts']})")


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

        if run_mode == "morning":
            gchat_msg = format_morning(dir_alerts, stale_items, today, enable_mentions)
            slack_msg = format_morning_slack(dir_alerts, stale_items, today, enable_mentions)
        else:
            gchat_msg = format_followup(dir_alerts, today, enable_mentions)
            slack_msg = format_followup_slack(dir_alerts, today, enable_mentions)
            if gchat_msg is None and slack_msg is None:
                print("超過案件なし → フォロー通知スキップ")
                save_tracking(gc, current_statuses, prev_tracking, today)
                return

        print("--- Google Chat ---")
        print(gchat_msg)
        print("--- Slack ---")
        print(slack_msg)
        print("---")

        if gchat_msg:
            send_to_google_chat(gchat_msg)
        if slack_msg:
            send_to_slack(slack_msg)

        # トラッキング保存
        save_tracking(gc, current_statuses, prev_tracking, today)

    except Exception as e:
        error_msg = f"⚠️ Bot実行エラー：{e}"
        print(error_msg, file=sys.stderr)
        try:
            send_to_google_chat(error_msg)
            send_to_slack(error_msg)
        except Exception as send_err:
            print(f"エラー通知の送信にも失敗: {send_err}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
