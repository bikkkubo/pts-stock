# PTS Stock Analysis System

## 概要

Google Apps Script (GAS) を使用したPTS（Pre-Trade Session）株価分析・レポート自動生成システムです。毎日自動で株価変動の上位銘柄を取得し、AI分析とニュース情報を統合したレポートを生成します。

## 主な機能

### 1. 日次PTS分析レポート（6:45 AM実行）
- Kabutanから上昇・下落上位20銘柄を自動取得
- 複数ニュースソースから関連情報を収集
- OpenAI APIを使用した株価変動要因の分析
- Google Spreadsheetsへの自動出力

### 2. 日次PTSランキング更新（8:40 AM実行）
- 上昇・下落トップ10銘柄のランキング
- 時系列での推移追跡
- 色分け表示（上昇=緑、下落=赤）

### 3. 多角的ニュース収集
- **Kabutan**: 株探ニュース
- **TDnet**: 適時開示情報
- **企業公式IR**: 投資家向け情報
- **Yahoo Finance Japan**: 金融ニュース
- **Bloomberg Japan**: 国際金融情報
- **東洋経済オンライン**: 経済分析

## ファイル構成

```
├── Code.gs              # メイン処理・トリガー設定
├── fetchers.gs          # データ取得処理
├── nlp.gs              # OpenAI API連携・テキスト分析
├── writer.gs           # Google Sheets書き込み
├── dailyRanking.gs     # 日次ランキング機能
├── setup.gs            # システムセットアップ
├── apiKeyTest.gs       # API接続テスト
├── notifier.gs         # 通知機能
├── util.gs             # ユーティリティ関数
└── appsscript.json     # GAS設定ファイル
```

## セットアップ手順

### 1. Google Apps Script準備
1. [script.google.com](https://script.google.com) でプロジェクト作成
2. 全ファイルをコピー＆ペースト

### 2. API キー設定
```
設定 > スクリプトプロパティ:
- OPENAI_API_KEY: OpenAI APIキー (必須)
- NIKKEI_API_KEY: 日経APIキー (オプション)
```

### 3. 初回セットアップ実行
```javascript
// GAS Editorで実行
setupPtsSystem()
```

### 4. 権限承認
- Google Sheets アクセス
- 外部URL取得 (UrlFetchApp)
- スクリプトトリガー設定

## システム設定

### 自動実行スケジュール
- **メイン分析**: 毎日 6:45 AM JST
- **ランキング更新**: 毎日 8:40 AM JST

### 出力先
- Google Spreadsheets (自動作成)
  - `PTS Daily Report`: 詳細分析レポート
  - `PTS Daily Ranking`: 日次ランキング履歴

## データ構造

### PTS Daily Report
| 列 | 項目 | 説明 |
|---|---|---|
| A | コード | 銘柄コード |
| B | 銘柄名 | 会社名 |
| C | 始値 | PTS開始価格 |
| D | 終値 | PTS終了価格 |
| E | 差額 | 価格変動額 |
| F | 騰落率(%) | 変動率 |
| G | AI要約 | 変動要因分析 |
| H | 情報源 | 参照URL |

### PTS Daily Ranking
| 列 | 項目 | 説明 |
|---|---|---|
| A | 日付 | データ日付 |
| B | 分類 | TOP10上昇/TOP10下落 |
| C | 順位 | ランキング順位 |
| D-H | 銘柄情報 | コード、名前、価格など |

## AI分析機能

### OpenAI統合
- GPT-3.5-turbo使用
- ニュース記事のクラスタリング
- 株価変動要因の自動分析
- フォールバック分析（API失敗時）

### 分析内容
- 決算・業績分析
- IR発表内容評価
- 市場環境影響
- セクター特性考慮
- 企業固有要因特定

## 管理機能

### テスト・診断
```javascript
testCurrentOpenAIKey()    // API接続確認
quickSystemTest()         // システム全体テスト
showPtsSystemStatus()     // 現在状態表示
```

### トリガー管理
```javascript
setupPtsSystem()          // 全システムセットアップ
cleanupPtsSystem()        // トリガー全削除
```

## トラブルシューティング

### よくある問題

1. **OpenAI API エラー**
   - API キーの確認
   - `testCurrentOpenAIKey()` でテスト

2. **スプレッドシート権限エラー**
   - Google Sheets APIの権限承認
   - スプレッドシートの共有設定確認

3. **トリガー設定失敗**
   - `cleanupPtsSystem()` → `setupPtsSystem()` で再設定

### ログ確認
- GAS Editor > 実行 > ログを表示
- エラー詳細と実行状況を確認

## 技術仕様

- **言語**: Google Apps Script (JavaScript)
- **実行環境**: Google Cloud Platform
- **AI**: OpenAI GPT-3.5-turbo
- **データソース**: 複数金融情報サイト
- **出力**: Google Spreadsheets
- **スケジューリング**: GAS時間ベーストリガー

## ライセンス

MIT License

## 更新履歴

- **2025-06-15**: 初期リリース
  - PTS分析レポート機能
  - 日次ランキング機能
  - AI要因分析機能
  - 多角的ニュース収集