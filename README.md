# うお会 イベント管理システム

浅草で毎月ゾロ目の日に開催するイベント「うお会」の運営システム。

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

4. スプレッドシートにシート `master_customers` / `event_histories` を作成
5. エディタから `setupSpreadsheetHeaders()` を一度実行（ヘッダー自動作成）
6. **デプロイ → 新しいデプロイ → ウェブアプリ**（アクセス: 全員）

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

- [ ] Googleフォーム → GAS トリガー連携
- [ ] Vercel フロント（受付・管理画面）
- [ ] LIFF 連携（当日受付・マスタ更新）
- [ ] pending_match 承認 UI
