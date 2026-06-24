# うおの会 イベント管理システム

浅草で毎月ゾロ目の日に開催するイベント「うおの会」の運営システム。

## 構成（予定）

| レイヤ | 技術 |
|--------|------|
| バックエンド API | Google Apps Script (GAS) |
| DB | Google スプレッドシート（2層） |
| フロント | Vercel（HTML/CSS/JS + Tailwind） |
| 外部連携 | LINE Messaging API / LIFF、Googleフォーム |

## フォルダ

```
04_イベント管理システム/
├── gas/
│   ├── Code.js          # Web API 本体（doPost / 名寄せ / LINE通知）
│   └── appsscript.json  # GAS プロジェクト設定
└── README.md
```

## スプレッドシート

### ① master_customers（顧客マスタ）

| 列 | フィールド | 備考 |
|----|-----------|------|
| A | line_user_id | 主キー（LINE連携後） |
| B | user_name | スペース除去済み氏名 |
| C | company_name | |
| D | position_category | 経営者/個人事業主/会社員/その他 |
| E | position_name | |
| F | email | 小文字・ユニークキー |
| G | phone_number | 数字のみ |
| H | referrer | |
| I | status | active / pending_match |
| J | created_at | |

### ② event_histories（イベント参加履歴）

| 列 | フィールド | 備考 |
|----|-----------|------|
| A | timestamp | |
| B | event_date | 例: 20260707 |
| C | form_email | マスタ突合キー |
| D | receipt_required | 要/不要 |
| E | receipt_name | |
| F | line_user_id | 受付時紐付け |
| G | payment_status | 初期値「未」 |
| H | attendance_status | 初期値「未」 |
| I | receipt_dl_count | 初期値 0 |

## GAS セットアップ

1. [Google Apps Script](https://script.google.com/) で新規プロジェクト作成
2. `gas/Code.js` の内容を貼り付け
3. **プロジェクトの設定 → スクリプト プロパティ** に以下を追加

| プロパティ | 値 |
|-----------|-----|
| `SPREADSHEET_ID` | 対象スプレッドシート ID |
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Messaging API トークン |
| `LINE_ADMIN_USER_ID` | 通知先 userId（魚谷さん） |
| `DEFAULT_EVENT_DATE` | （任意）フォームにイベント日がない場合の既定値 `20260707` |
| `FORM_RESPONSE_SHEET_NAME` | （任意）回答シート名。未設定時は「フォームの回答」を自動検出 |

4. スプレッドシートにシート `master_customers` / `event_histories` を作成
5. エディタから `setupSpreadsheetHeaders()` を一度実行（ヘッダー自動作成）
6. **デプロイ → 新しいデプロイ → ウェブアプリ**（アクセス: 全員）

## Googleフォーム自動連携

### 前提

- Googleフォームの回答先を **同じスプレッドシート**（`SPREADSHEET_ID`）に設定する
- GAS プロジェクトはそのスプレッドシートに **コンテナバインド** するか、Script Property で同じ ID を指定する

### フォームの推奨設問名

| 設問（タイトル） | 必須 | 備考 |
|----------------|------|------|
| 氏名 | ✅ | |
| メールアドレス | ✅ | |
| 参加希望日 | 推奨 | なければ `DEFAULT_EVENT_DATE` を使用 |
| 会社名 | | |
| 区分 | | 経営者/会社員 等 |
| 役職 | | |
| 電話番号 | | |
| 紹介者 | | |
| 領収書 | | 要/不要 |
| 領収書宛名 | | |

設問名は部分一致でも認識します（例: 「お名前」→ 氏名）。

### セットアップ手順

1. フォームの回答先を対象スプレッドシートにリンク
2. GAS エディタで `installFormSubmitTrigger()` を **1回実行**（承認が必要）
3. テスト送信 → `master_customers` と `event_histories` に行が追加されるか確認

### 手動テスト

- `processLatestFormResponse()` … 回答シートの **最終行** を処理（トリガーなしで試すとき）
- `uninstallFormSubmitTrigger()` … トリガー削除

### 処理内容（自動）

```
フォーム送信
  → onFormSubmit トリガー
  → handleEventApplication_（名寄せ）
  → master_customers（新規のみ追加）
  → event_histories（必ず追加・当日参加者リスト）
  → LINE 通知（魚谷さん）
```

同一メール・同一イベント日の **重複申込はスキップ** されます。

## API: POST イベント申込

```json
POST /exec
Content-Type: application/json

{
  "action": "event_application",
  "event_date": "20260707",
  "user_name": "山田 太郎",
  "email": "Taro.Yamada@Example.COM",
  "company_name": "株式会社サンプル",
  "position_category": "経営者",
  "position_name": "代表取締役",
  "phone_number": "090-1234-5678",
  "referrer": "紹介者名",
  "receipt_required": "要",
  "receipt_name": "山田太郎"
}
```

### 名寄せパターン

| パターン | 条件 | 処理 |
|---------|------|------|
| A | email 完全一致 | event_histories のみ追加（マスタ不更新） |
| B | email 不一致・氏名一致 | マスタ追加（pending_match）+ 履歴追加 |
| C | 両方不一致 | マスタ追加（active）+ 履歴追加 |

いずれも LINE Push で管理者に通知されます。

## 次のステップ

- [x] Googleフォーム → GAS トリガー連携
- [x] Vercel フロント（受付 LIFF）
- [x] LIFF 連携（当日受付・マスタ更新）
- [ ] pending_match 承認 UI
