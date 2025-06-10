# Excel Format Tool

## プロジェクト概要

Excel Format Tool は、指定された Excel ファイルからテーブルデータを抽出し、事前定義されたテンプレート形式で再整理・出力する専門的な Excel ファイル書式設定ツールです。このツールは、日本語環境でのインターフェース仕様書およびファイル仕様書の処理に特に適用されます。

## 主な機能

- **自動ターゲットファイル認識**：設定ファイル内の「●」マークされた行に基づいて、処理すべきファイルとテーブルを自動認識
- **マルチフォーマットサポート**：インターフェース仕様書とファイル仕様書の両方の形式をサポート
- **テンプレート化出力**：事前定義されたテンプレートを使用して標準化された Excel 出力ファイルを生成
- **バッチ処理**：複数のファイルとテーブルを同時に処理
- **インテリジェント書式変換**：自動的なヘッダーマッピングとデータ書式設定

## ファイル構造

```
format_tool/
├── main.py                    # メインプログラムエントリー
├── requirements.txt           # 依存パッケージリスト
├── .env                      # 環境変数設定
├── README.md                 # プロジェクト説明ドキュメント
├── excel_files/              # ソース Excel ファイルディレクトリ
├── target/                   # 設定ファイルディレクトリ
│   ├── Templete_File_chosen.xlsm  # ファイル選択設定
│   └── Format_list.xlsx      # 略称名と対応する表名一覧
├── templates/                # テンプレートファイルディレクトリ
│   └── Template.xlsm         # 出力テンプレート
├── output/                   # 出力ファイルディレクトリ（自動生成）
└── utils/                    # ユーティリティモジュール
    ├── common/              # 共通ツールのフォルダー
    │   ├── position_identify.py
    |   └── write_to_excel.py
    ├── process_interface_file.py  # インターフェースファイル処理のフォルダー
    └── process_file_file.py       # ファイル仕様処理のフォルダー
```

## インストール説明

### 環境要件

- Python 3.11.7
- pip パッケージマネージャー

### インストール手順

1. **仮想環境の作成**

   ```bash
   python3.11 -m venv venv
   source venv/bin/activate  # Linux/Mac
   # または
   venv\Scripts\activate     # Windows
   ```

2. **依存関係のインストール**

   ```bash
   pip install -r requirements.txt
   ```

3. **環境変数の設定**

   環境変数の設定は不要

## 基本的な使用

1. **ファイルの準備**

   - 処理したい Excel ファイルを `excel_files/` ディレクトリに配置
   - `target/Templete_File_chosen.xlsm` でターゲットファイルとテーブル情報が正しく設定されていることを確認
   - `target/Format_list.xlsx` に書式マッピング情報が含まれていることを確認

2. **仮想環境を起動**

   ```bash
   source venv/bin/activate
   ```

3. **プログラムの実行**

```bash
python main.py
```

4. **出力の確認**
   - 処理完了後、結果ファイルは `output/YYYYMMDDHHMM/` ディレクトリに保存されます
   - インターフェースファイルは `interface_excel_*.xlsx` として出力
   - ファイル仕様は `file_excel_*.xlsx` として出力
