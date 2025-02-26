# ボードゲーム攻略情報・研究結果プレゼンテーション作成ツール

このツールは、ボードゲームの攻略情報や OpenAI の DeepResearch の結果をパワーポイントプレゼンテーションに自動的にまとめる Python スクリプトです。

## 機能

- 一般的なボードゲーム攻略情報をプレゼンテーションに変換（`create_presentation.py`）
- アルナック（Lost Ruins of Arnak）専用の戦略プレゼンテーション作成（`create_arnak_presentation.py`）
- **OpenAI の DeepResearch の結果をプレゼンテーションに変換（`create_deep_research_presentation.py`）**
- タイトルスライドの自動生成
- 目次スライドの自動生成
- 攻略情報や研究結果を段落ごとにスライド化
- URL や参考文献などの不要な情報を自動的に除外
- 長い内容は自動的に要約または複数スライドに分割
- テーマに合わせた背景画像と配色
- **複数のカラーテーマから選択可能（青、暗い、明るい、緑）**

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

### OpenAI の DeepResearch の結果をプレゼンテーションにする場合

1. DeepResearch の結果をテキストファイルに保存します（例：`research_result.txt`）。

2. 以下のコマンドを実行します：

```bash
python create_deep_research_presentation.py research_result.txt
```

3. オプションでタイトルやテーマを指定することもできます：

```bash
python create_deep_research_presentation.py research_result.txt -t "研究テーマ" --theme dark
```

4. 利用可能なオプション：

   - `-o, --output`: 出力ファイル名を指定
   - `-t, --title`: プレゼンテーションのタイトルを指定
   - `--theme`: カラーテーマを指定（blue, dark, light, green）

5. 指定したファイル名または自動生成された名前（例：`deep_research_20230401_123456.pptx`）のパワーポイントファイルが生成されます。

## ファイル構成

- `create_presentation.py` - 一般的なボードゲーム攻略情報用スクリプト
- `create_arnak_presentation.py` - アルナック専用スクリプト
- `create_deep_research_presentation.py` - **OpenAI の DeepResearch 結果用スクリプト**
- `game_info.txt` - ボードゲームの攻略情報を記入するファイル
- `arnak_bg.jpg` - アルナック用背景画像
- `board_game_strategy.pptx` - 一般スクリプトで生成されるファイル
- `arnak_strategy.pptx` - アルナック専用スクリプトで生成されるファイル

## 注意事項

- 情報が多すぎる場合は、一部が省略される場合があります。
- 最適な結果を得るためには、`game_info.txt` の情報を段落ごとに整理してください。
- アルナック専用スクリプトは初回実行時に背景画像をダウンロードします。
- **DeepResearch の結果には通常、参考文献や URL が含まれますが、これらは自動的に除外されます。**
- **参考文献セクション（「参考文献」「References」などで始まるセクション）以降の内容は完全に除外されます。**

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。詳細は [LICENSE](LICENSE) ファイルを参照してください。
