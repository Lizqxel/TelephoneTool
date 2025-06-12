# TelephoneTool - コールセンター業務効率化ツール

![バージョン](https://img.shields.io/badge/version-1.7.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![ライセンス](https://img.shields.io/badge/license-proprietary-red.svg)

## 📋 概要

TelephoneToolは、コールセンター業務の効率化を目的として開発されたWindows専用のGUIアプリケーションです。PySide6を使用して構築されており、顧客情報の管理、CTIフォーマットの生成、NTT西日本の提供エリア判定など、コールセンター業務に特化した機能を提供します。

## ✨ 主な機能

### 🎯 コア機能
- **顧客情報入力・管理** - 構造化された顧客データの入力フォーム
- **CTIフォーマット生成** - コールセンター標準フォーマットの自動生成
- **クリップボード監視** - 自動的なデータ取得・転記機能
- **Google スプレッドシート連携** - データの自動転記・同期（未実装）

### 🌐 提供エリア判定機能
- **NTT西日本 提供エリア検索** - 住所から光回線の提供可否を自動判定
- **ワンクリック検索** - 簡単操作での提供判定実行
- **詳細エラー表示** - 判定失敗時の詳細なエラー情報表示

### 📞 CTI連携機能
- **電話番号監視** - 着信番号の自動検出
- **フォーマット自動生成** - 着信情報に基づくフォーマット作成
- **リアルタイム更新** - 通話状況に応じた動的な情報更新

## 🛠️ 技術スタック

### フレームワーク・ライブラリ
```
PySide6 >= 6.6.1              # GUI フレームワーク
requests >= 2.31.0            # HTTP クライアント
selenium == 4.18.1            # Web 自動化
google-api-python-client      # Google API クライアント
webdriver-manager == 4.0.1    # WebDriver 管理
pywin32 >= 306               # Windows API
```

### 開発・テスト
```
pytest == 7.4.3              # テストフレームワーク
pillow >= 10.1.0             # 画像処理
jaconv == 0.3.4              # 日本語文字変換
python-dotenv >= 1.0.0       # 環境変数管理
```

## 🚀 インストール方法

### 前提条件
- Windows 10/11 (64bit)
- Python 3.8 以上
- Google Chrome ブラウザ（提供エリア判定機能使用時）

### 1. リポジトリのクローン
```powershell
git clone <repository-url>
cd TelephoneTool
```

### 2. 仮想環境の作成（推奨）
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. 依存関係のインストール
```powershell
pip install -r requirements.txt
```

## 📖 使用方法

### 基本的な起動方法
```powershell
python main.py
```

### モード選択
アプリケーションは以下のモードで動作します：
- **簡単モード** - 基本的な機能のみ
- **標準モード** - 全機能利用可能

### 提供エリア判定の使用方法
1. 顧客情報入力フォームに住所を入力
2. 「提供エリア判定」ボタンをクリック
3. 自動的にNTT西日本サイトにアクセスして判定実行
4. 結果がフォーマットに自動反映

### CTIフォーマット生成
1. 必要な顧客情報を入力フォームに記入
2. 「フォーマット生成」ボタンをクリック
3. 生成されたフォーマットをクリップボードにコピー

## 📦 ビルド・配布

### 実行ファイルの作成
```powershell
# PyInstaller を使用したビルド
python build_release.py

# または直接 PyInstaller を実行
pyinstaller build_release.spec
```

### インストーラーの作成
Inno Setup を使用してWindowsインストーラーを作成：
```powershell
# telephoneTool_installer.iss を Inno Setup でコンパイル
```

## 📁 プロジェクト構造

```
TelephoneTool/
├── main.py                           # メインエントリーポイント
├── requirements.txt                  # Python依存関係
├── settings.json                     # アプリケーション設定
├── ui/                              # ユーザーインターフェース
│   ├── main_window.py               # メインウィンドウ
│   ├── easy_mode_window.py          # 簡単モードウィンドウ
│   ├── settings_dialog.py           # 設定ダイアログ
│   └── ...
├── services/                        # ビジネスロジック
│   ├── area_search.py               # 提供エリア検索サービス
│   ├── cti_service.py               # CTI連携サービス
│   ├── oneclick.py                  # ワンクリック機能
│   └── ...
├── utils/                           # ユーティリティ関数
│   ├── string_utils.py              # 文字列処理
│   ├── address_utils.py             # 住所処理
│   ├── logger.py                    # ログ管理
│   └── ...
├── tests/                           # テストコード
│   ├── test_address_utils.py        # 住所処理テスト
│   └── ...
├── logs/                            # ログファイル
├── drivers/                         # WebDriverファイル
└── chrome_data/                     # Chromeユーザーデータ
```

## ⚙️ 設定

### settings.json の主要設定項目
```json
{
  "format_template": "...",           // CTIフォーマットテンプレート
  "font_size": 9,                     // UI フォントサイズ
  "delay_seconds": 0,                 // 処理遅延設定
  "browser_settings": {               // ブラウザ設定
    "headless": false,                // ヘッドレスモード
    "disable_images": true,           // 画像読み込み無効化
    "auto_close": true,               // 自動クローズ
    "page_load_timeout": 30           // ページ読み込みタイムアウト
  },
  "mode": "simple",                   // 起動モード
  "show_mode_selection": false        // モード選択ダイアログ表示
}
```

## 🐛 トラブルシューティング

### よくある問題と解決方法

#### 1. WebDriverエラー
```
selenium.common.exceptions.WebDriverException
```
**解決方法**: Chrome ブラウザを最新版に更新してください。

#### 2. 提供エリア判定が失敗する
**解決方法**: 
- インターネット接続を確認
- NTT西日本サイトのアクセス可否を確認
- ログファイル（app.log）でエラー詳細を確認

## 📝 ログ

アプリケーションは詳細なログを `app.log` ファイルに出力します：
```
2024-01-01 12:00:00,000 - INFO - アプリケーション起動
2024-01-01 12:00:01,000 - INFO - メインウィンドウ表示
2024-01-01 12:00:02,000 - ERROR - 提供エリア検索エラー: 詳細情報
```
---