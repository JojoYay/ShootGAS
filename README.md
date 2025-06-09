# ShootGAS

Google Apps Scriptを使用したShootLiffのバックエンドサーバー

## 開発環境のセットアップ

### 1. 必要なツールのインストール
```bash
# claspのインストール
npm install -g @google/clasp

# プロジェクトの依存関係をインストール
npm install
```

### 2. Google Apps Scriptへのログイン
```bash
clasp login
```

### 3. プロジェクトの設定
- `appsscript.json`の設定を確認
- 必要な環境変数の設定

## 開発

### ローカルでの開発
```bash
# 開発サーバーの起動
npm run dev

# ビルド
npm run build

# テストの実行
npm test
```

### デプロイ
```bash
# Google Apps Scriptへのデプロイ
clasp push
```

## プロジェクト構造

```
ShootGAS/
├── src/           # ソースコード
├── resources/     # リソースファイル
├── appsscript.json  # Google Apps Scriptの設定
└── package.json   # プロジェクトの依存関係
```

## 環境変数

環境変数はデプロイ先のGAS側のpropertiesを利用します

### 必要なスクリプトプロパティ
以下のプロパティをGoogle Apps Scriptのプロジェクト設定で設定する必要があります：

| プロパティ名 | 説明 |
|------------|------|
| `calendarId` | GoogleカレンダーのID |
| `eventResults` | イベント結果の設定 |
| `reportSheet` | レポート用スプレッドシートのID |
| `liffUrl` | LINE Front-end FrameworkのURL |
| `settingSheet` | 設定用スプレッドシートのID |
| `lineAccessToken` | LINE Messaging APIのアクセストークン |
| `folderId` | メインフォルダのID |
| `archiveFolder` | アーカイブ用フォルダのID |
| `expenseFolder` | 経費関連ファイル用フォルダのID |
| `channelQr` | LINEチャンネルのQRコード情報 |
| `channelUrl` | LINEチャンネルのURL |
| `messageUsage` | メッセージ使用量の設定 |
| `chat` | チャット関連の設定 |

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルを参照してください。 