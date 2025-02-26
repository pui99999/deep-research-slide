from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import re

def create_board_game_presentation(game_info, output_file="board_game_strategy.pptx"):
    """
    ボードゲームの攻略情報をパワーポイントにまとめる関数
    
    Args:
        game_info (str): ボードゲームの攻略情報のテキスト
        output_file (str): 出力するパワーポイントファイル名
    """
    # プレゼンテーションの作成
    prs = Presentation()
    
    # スライドのレイアウト
    title_slide_layout = prs.slide_layouts[0]  # タイトルスライド
    content_slide_layout = prs.slide_layouts[1]  # タイトルと内容のスライド
    
    # タイトルスライドの作成
    title_slide = prs.slides.add_slide(title_slide_layout)
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1]
    
    # ゲーム名を抽出（最初の行または「ゲーム名:」などの形式から）
    game_name = game_info.strip().split('\n')[0]
    if ':' in game_name:
        game_name = game_name.split(':', 1)[1].strip()
    
    title.text = game_name
    subtitle.text = "ボードゲーム攻略ガイド"
    
    # テキストを段落に分割
    paragraphs = re.split(r'\n\s*\n', game_info)
    
    # 目次スライドの作成
    toc_slide = prs.slides.add_slide(content_slide_layout)
    toc_title = toc_slide.shapes.title
    toc_content = toc_slide.placeholders[1]
    
    toc_title.text = "目次"
    
    # 目次の内容を作成
    toc_text = ""
    section_titles = []
    
    for i, para in enumerate(paragraphs):
        if i == 0:  # 最初の段落はタイトルなのでスキップ
            continue
        
        # 段落の最初の行をセクションタイトルとして使用
        lines = para.strip().split('\n')
        if lines:
            section_title = lines[0].strip()
            # URLを含む行は除外
            if not section_title.startswith('http') and not 'www.' in section_title:
                section_titles.append(section_title)
                toc_text += f"• {section_title}\n"
    
    toc_content.text = toc_text
    
    # 各セクションのスライドを作成
    current_section = ""
    section_content = ""
    
    for i, para in enumerate(paragraphs):
        if i == 0:  # 最初の段落はタイトルなのでスキップ
            continue
        
        lines = para.strip().split('\n')
        if not lines:
            continue
        
        # URLを含む行は除外
        filtered_lines = [line for line in lines if not line.startswith('http') and not 'www.' in line]
        if not filtered_lines:
            continue
        
        section_title = filtered_lines[0].strip()
        
        # 新しいセクションの開始
        if section_title in section_titles:
            # 前のセクションがあれば、そのスライドを作成
            if current_section:
                create_section_slide(prs, content_slide_layout, current_section, section_content)
            
            current_section = section_title
            section_content = '\n'.join(filtered_lines[1:])
        else:
            # 同じセクションの続き
            section_content += '\n' + '\n'.join(filtered_lines)
    
    # 最後のセクションのスライドを作成
    if current_section:
        create_section_slide(prs, content_slide_layout, current_section, section_content)
    
    # まとめスライドの作成
    summary_slide = prs.slides.add_slide(content_slide_layout)
    summary_title = summary_slide.shapes.title
    summary_content = summary_slide.placeholders[1]
    
    summary_title.text = "まとめ"
    summary_content.text = f"{game_name}の攻略ポイント：\n\n• 基本ルールを理解する\n• 戦略的な思考を身につける\n• 経験を積んで上達しよう"
    
    # プレゼンテーションの保存
    prs.save(output_file)
    print(f"プレゼンテーションを {output_file} として保存しました。")

def create_section_slide(prs, layout, title, content):
    """
    セクションのスライドを作成する関数
    
    Args:
        prs: プレゼンテーションオブジェクト
        layout: スライドレイアウト
        title (str): スライドのタイトル
        content (str): スライドの内容
    """
    slide = prs.slides.add_slide(layout)
    title_shape = slide.shapes.title
    content_shape = slide.placeholders[1]
    
    title_shape.text = title
    
    # 内容が長すぎる場合は分割
    if len(content) > 1000:
        # 内容を箇条書きに変換
        bullet_points = []
        for line in content.split('\n'):
            line = line.strip()
            if line:
                if not line.startswith('•'):
                    line = f"• {line}"
                bullet_points.append(line)
        
        # 箇条書きを結合
        content = '\n'.join(bullet_points[:10])  # 最初の10項目だけ表示
        content += "\n• ..."  # 省略記号を追加
    
    content_shape.text = content

# メイン処理
if __name__ == "__main__":
    # ユーザーからの入力を受け取る
    print("ボードゲームの攻略情報をテキストファイルから読み込みます。")
    
    try:
        with open("game_info.txt", "r", encoding="utf-8") as f:
            game_info = f.read()
        
        create_board_game_presentation(game_info)
    except FileNotFoundError:
        print("game_info.txt ファイルが見つかりません。")
        print("テキストファイルを作成し、ボードゲームの攻略情報を記入してください。") 