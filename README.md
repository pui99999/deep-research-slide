# ボードゲーム攻略情報プレゼンテーション作成ツール

このツールは、ボードゲームの攻略情報をパワーポイントプレゼンテーションに自動的にまとめる Python スクリプトです。

## 機能

- 一般的なボードゲーム攻略情報をプレゼンテーションに変換（`create_presentation.py`）
- アルナック（Lost Ruins of Arnak）専用の戦略プレゼンテーション作成（`create_arnak_presentation.py`）
- タイトルスライドの自動生成
- 目次スライドの自動生成
- 攻略情報を段落ごとにスライド化
- URL などの不要な情報を自動的に除外
- 長い内容は自動的に要約
- テーマに合わせた背景画像と配色

## 必要条件

- Python 3.6 以上
- python-pptx ライブラリ
- requests ライブラリ（アルナック専用スクリプトの場合）

## インストール方法

リポジトリをクローンし、必要なライブラリをインストールします：

```bash
git clone https://github.com/yourusername/slide.git
cd slide
pip install -r requirements.txt
```

## 使用方法

### 一般的なボードゲーム攻略情報の場合

1. `game_info.txt` ファイルにボードゲームの攻略情報を記入します。

   - 最初の行にはボードゲーム名を記入してください。
   - 段落ごとに情報を整理してください。
   - 段落の最初の行は見出しとして使用されます。

2. スクリプトを実行します：

```bash
python create_presentation.py
```

3. `board_game_strategy.pptx` という名前のパワーポイントファイルが生成されます。

### アルナック専用プレゼンテーションの場合

1. 以下のコマンドを実行します：

```bash
python create_arnak_presentation.py
```

2. `arnak_strategy.pptx` という名前のパワーポイントファイルが生成されます。

## ファイル構成

- `create_presentation.py` - 一般的なボードゲーム攻略情報用スクリプト
- `create_arnak_presentation.py` - アルナック専用スクリプト
- `game_info.txt` - ボードゲームの攻略情報を記入するファイル
- `arnak_bg.jpg` - アルナック用背景画像
- `board_game_strategy.pptx` - 一般スクリプトで生成されるファイル
- `arnak_strategy.pptx` - アルナック専用スクリプトで生成されるファイル

## 注意事項

- 情報が多すぎる場合は、一部が省略される場合があります。
- 最適な結果を得るためには、`game_info.txt` の情報を段落ごとに整理してください。
- アルナック専用スクリプトは初回実行時に背景画像をダウンロードします。

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。詳細は [LICENSE](LICENSE) ファイルを参照してください。
