# PTS日次レポート – Google Apps Script版

## 概要
このプロジェクトは**エンドツーエンドのサーバーレス**ワークフローを作成します：
1. 平日の06:45 JSTに実行（時間ベーストリガー）
2. QUICK APIを使用して前日のPTS上位10銘柄の**値上がり・値下がり**（`始値`, `終値`, `差額`）を取得
3. 各銘柄の24時間以内のニュース・IR情報を取得（日経、ロイター、TDnet RSS）
4. 記事を**重複除去**し、OpenAI埋め込みで**クラスタリング**、3文の日本語要約と参照URLを生成
5. 結果をGoogleスプレッドシートの**A:G列**に書き込み（シート名は`writer.gs`参照）
6. 成功・失敗時にSlack Webhookで通知送信

すべてのロジックは**Google Apps Script**ファイルに含まれており、`clasp`でデプロイ可能です。

---

## 環境
| ツール | バージョン / 備考 |
|------|----------------|
| VS Code | 1.96.4 |
| clasp  | `npm i -g @google/clasp` |
| GAS Runtime | V8 |
| OpenAI Model | `gpt-4o-mini` / temperature 0.3 |

設定が必要なスクリプトプロパティ（⚙️ **プロジェクト設定 » スクリプトプロパティ**）：
- `QUICK_API_TOKEN` – PTS価格データAPI
- `NIKKEI_API_KEY` – 日経ニュースAPI
- `OPENAI_API_KEY` – OpenAI API
- `SLACK_WEBHOOK` – Slack通知用WebhookURL  

---

## ファイル構成
```
/pts-20250614
├─ Code.gs          // メイン処理、トリガー
├─ fetchers.gs      // 株価・ニュース取得
├─ nlp.gs           // 埋め込み、クラスタリング、GPT要約
├─ writer.gs        // スプレッドシート入出力
├─ notifier.gs      // Slack通知ヘルパー
├─ util.gs          // 数学ヘルパー関数
└─ .clasp.json      // <-- scriptIdを挿入
```

---

## クイックスタート
```bash
# 1. クローン & 認証
git clone https://github.com/bikkkubo/pts-20250614.git
cd pts-20250614
clasp login

# 2. GASプロジェクト作成
clasp create --title "PTS Daily Report"

# 3. .clasp.jsonにscriptIdを挿入
# 4. コードをpush
clasp push

# 5. スクリプトプロパティを設定
# 6. installTriggers()を一度実行
```

## 詳細セットアップ手順

### 1. 依存関係のインストール
```bash
npm install -g @google/clasp
```

### 2. 認証
```bash
clasp login
```
ブラウザが開きGoogleアカウントでの認証を求められます。

### 3. 新しいGASプロジェクトの作成
```bash
clasp create --title "PTS Daily Report"
```

### 4. スクリプトIDの設定
`.clasp.json`を編集し、`"ENTER_YOURS"`を前ステップで生成された実際のスクリプトIDに置き換えてください。

### 5. コードのデプロイ
```bash
clasp push
```

### 6. スクリプトプロパティの設定
GASエディタの**プロジェクト設定 » スクリプトプロパティ**で以下を追加：

| プロパティ | 値 | 説明 |
|----------|-------|-------------|
| `QUICK_API_TOKEN` | あなたのトークン | PTS価格データAPIトークン |
| `NIKKEI_API_KEY` | あなたのキー | 日経ニュースAPIキー |
| `OPENAI_API_KEY` | あなたのキー | OpenAI API キー（埋め込み・要約用） |
| `SLACK_WEBHOOK` | webhook_url | Slack着信WebhookURL |

#### 各APIキーの取得方法：

**QUICK_API_TOKEN：**
1. QUICK公式サイトでアカウント作成
2. API利用申請を行い承認を待つ
3. 管理画面からAPIトークンを取得

**NIKKEI_API_KEY：**
1. 日経電子版APIサービスに登録
2. 開発者向けAPIプランを申し込み
3. APIキーを管理画面から取得

**OPENAI_API_KEY：**
1. OpenAI公式サイト(https://openai.com)でアカウント作成
2. API使用量に応じた課金プランを設定
3. API Keys セクションから新しいキーを生成

**SLACK_WEBHOOK：**
1. Slackワークスペースで新しいアプリを作成
2. Incoming Webhooksを有効化
3. 通知先チャンネルを選択してWebhook URLを取得

### 7. トリガーのインストール
GASエディタで`installTriggers()`関数を一度実行し、日次自動化を設定します。

### 8. セットアップのテスト
- `testSlackWebhook()`を実行してSlack連携を確認
- `main()`を手動実行して完全なワークフローをテスト

---

## 使用方法

### 手動実行
GASエディタで`main()`関数を実行してレポートを手動生成できます。

### 自動実行
インストール後、平日の06:45 JSTに自動でトリガーが実行されます。

### トリガー管理
- **インストール**: `installTriggers()`を実行
- **削除**: `removeTriggers()`を実行

---

## 出力フォーマット

レポートはGoogleスプレッドシートのA:G列にデータを書き込みます：

| 列 | 内容 | 説明 |
|--------|---------|-------------|
| A | コード | 株式銘柄コード |
| B | 始値 | 始値 |
| C | 終値 | 終値 |
| D | 差額 | 価格差 |
| E | 騰落率(%) | 変動率 |
| F | AI要約 | AI生成の日本語要約 |
| G | 情報源 | 参照URL |

---

## 実装要件

1. 各ファイルに**完全なコード** - APIキー以外にプレースホルダーなし
2. ES5互換構文（`var`使用、`const`/`let`は不確実な場合避ける）
3. K-meansクラスタリング: `k = Math.round(Math.sqrt(n))`、20エポック後または重心不変で停止
4. 要約出力形式: 3文≤400文字、その後`Sources:`とURL一覧
5. 関数は**純粋でテスト可能**、可能な限り各80行以下
6. 平日06:45のトリガーをプログラムで設定する`installTriggers()`を追加

---

## トラブルシューティング

### よくある問題

**APIレート制限**: すべての外部APIにはレート制限があります。スクリプトにはエラー処理と開発用モックデータへのフォールバックが含まれています。

**トリガーが動作しない**:
- `installTriggers()`でトリガーが正しくインストールされているか確認
- タイムゾーンが'Asia/Tokyo'に設定されているか確認
- GASエディタの実行履歴を確認

**データが見つからない**:
- すべてのスクリプトプロパティが正しく設定されているか確認
- APIキーが有効で適切な権限があるか確認
- 具体的なエラーメッセージについては実行ログを確認

**Slack通知が動作しない**:
- `testSlackWebhook()`でWebhookをテスト
- Webhook URLの形式と権限を確認
- Slackアプリの設定を確認

### デバッグ
- コード全体に`Logger.log()`文を使用
- GASエディタの**実行**タブで詳細ログを確認
- 問題を特定するため関数を手動実行

---

## ライセンス
MIT（必要に応じて変更）

---

### ⭐ これでできること
- **プロンプト版**は「一問一答」で即ファイル生成
- **README版**はコード補完AIが継続的に参照しやすいドキュメントとして機能

どちらを採用するかは運用フローに合わせて選択してください。
不足・修正点があればお知らせいただければ追記いたします。