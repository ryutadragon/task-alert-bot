# サンキャク タスクアラートBot

Googleスプレッドシートの案件管理シートを毎朝読み取り、期限が近いタスクをGoogle Chatに自動通知する「つつく係Bot」。

## 機能

- **5種類の日付カラム**をチェック（次のタスク期日・締切/公開・撮影日①②③）
- **4段階のアラート**: 3日前 → 明日 → 本日 → 超過中（毎日カウントアップ）
- ステータスが「完了」「納品/公開待」「なし」の案件は自動スキップ
- ステータスを更新しない限り毎日通知し続ける

## セットアップ手順

### 1. Google Cloud Console でService Account作成

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを選択（または新規作成）
3. **APIとサービス** → **ライブラリ** → 「Google Sheets API」を検索して**有効化**
4. **APIとサービス** → **認証情報** → **認証情報を作成** → **サービスアカウント**
5. サービスアカウント名を入力（例: `task-alert-bot`）して作成
6. 作成したサービスアカウントの詳細画面 → **キー**タブ → **鍵を追加** → **新しい鍵を作成** → **JSON**
7. ダウンロードされたJSONファイルの**中身全体**をコピー（後でGitHub Secretsに登録）

### 2. スプレッドシートへのService Account共有

1. ダウンロードしたJSONファイル内の `client_email` を確認（例: `task-alert-bot@project.iam.gserviceaccount.com`）
2. 案件管理スプレッドシートを開く
3. **共有** → 上記メールアドレスを追加 → **閲覧者**権限で共有

### 3. Google Chat Webhook URL取得

1. Google Chatの案件管理スペースを開く
2. スペース名横の **▼** → **アプリと統合**
3. **Webhook** → **Webhookを追加**
4. 名前（例: `タスクアラートBot`）を入力して作成
5. 表示されたWebhook URLをコピー

### 4. GitHubリポジトリ作成・Secrets登録

1. GitHubで新しいリポジトリを作成
2. このディレクトリの内容をpush:
   ```bash
   cd task-alert-bot
   git init
   git add .
   git commit -m "初回コミット"
   git remote add origin https://github.com/<ユーザー名>/<リポジトリ名>.git
   git push -u origin main
   ```
3. リポジトリの **Settings** → **Secrets and variables** → **Actions** → **New repository secret** で以下を登録:

| Secret名 | 内容 |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Service AccountのJSONキー（ファイルの中身全体） |
| `SPREADSHEET_ID` | スプレッドシートURLの `/d/` と `/edit` の間の文字列 |
| `SHEET_NAME` | シート名（例: `案件管理`） |
| `GOOGLE_CHAT_WEBHOOK_URL` | Google Chat Webhook URL |

### 5. 動作確認

#### 手動実行（GitHub Actions）

1. リポジトリの **Actions** タブ → **Task Alert Bot** → **Run workflow** → **Run workflow**

#### ローカルで動作確認

```bash
# 環境変数を設定
export GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'
export SPREADSHEET_ID='your-spreadsheet-id'
export SHEET_NAME='案件管理'
export GOOGLE_CHAT_WEBHOOK_URL='https://chat.googleapis.com/v1/spaces/...'

# 依存ライブラリをインストール
pip install -r requirements.txt

# 実行
python main.py
```

## スケジュール

毎日 JST 9:00（UTC 00:00）に月〜土で自動実行されます（日曜は除外）。

## 通知フォーマット例

```
📋 サンキャク 本日のタスクアラート（2026/03/12）

💀 期限超過中
・ABC Corp / 商品紹介動画　次のタスク期日 ※2日超過
  → 初稿確認依頼
     PM: 田中｜Dir: 鈴木｜Editor: 佐藤

🚨 本日期限
・DEF Inc / ブランドPV　締切/公開
  → 最終納品
     PM: 田中｜Dir: 山田｜Editor: —

🔥 明日期限
・GHI Ltd / SNS広告　撮影日①
  → 撮影準備確認
     PM: —｜Dir: 鈴木｜Editor: —
```
