from pptx import Presentation  # pptxモジュールからPresentationクラスをインポート
from pptx.util import Inches  # pptxモジュールからInches関数をインポート
from pptx.enum.shapes import MSO_SHAPE_TYPE  # pptxモジュールからMSO_SHAPE_TYPE列挙型をインポート
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE

def scale_position(position, scale_width, scale_height):
    return (int(position[0] * scale_width), int(position[1] * scale_height))


def scale_size(size, scale_factor):  # サイズをスケーリングする関数を定義
    return tuple(int(dim * scale_factor) for dim in size)  # スケーリングされたサイズを返す

def scale_shape(shape, scale_width, scale_height):
    shape.left, shape.top = scale_position((shape.left, shape.top), scale_width, scale_height)

    # すべての形状のサイズをスケーリング
    shape.width = int(shape.width * scale_width)
    shape.height = int(shape.height * scale_height)
    
    if shape.has_text_frame:
        text_frame = shape.text_frame
        text_frame.word_wrap = True  # テキストの折り返しを有効にする
        text_frame.auto_size = MSO_AUTO_SIZE.NONE  # 自動サイズ調整を無効にする
        
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                # フォントサイズをスケーリング
                if run.font.size:
                    run.font.size = Pt(run.font.size.pt * scale_height)  # 高さに基づいてスケーリング
                # 文字の色を保持
                run.font.color.rgb = run.font.color.rgb
            
            # 行間をスケーリング
            if paragraph.line_spacing:
                paragraph.line_spacing = Pt(paragraph.line_spacing.pt * scale_height)  # 行間をスケーリング
            
            # 段落の配置を保持
            paragraph.alignment = paragraph.alignment

def scale_slide_contents(slide, scale_width, scale_height):
    for shape in slide.shapes:
        scale_shape(shape, scale_width, scale_height)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    cell.text_frame.auto_size = MSO_AUTO_SIZE.NONE  # 自動サイズ調整を無効にする
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            # テーブル内のフォントサイズをスケーリング
                            if run.font.size:
                                run.font.size = Pt(run.font.size.pt * scale_height)  # 高さに基づいてスケーリング
                            # テーブル内の文字の色を保持
                            run.font.color.rgb = run.font.color.rgb
                        
                        # 行間をスケーリング
                        if paragraph.line_spacing:
                            paragraph.line_spacing = Pt(paragraph.line_spacing.pt * scale_height)  # 行間をスケーリング
                        
                        # テーブル内の段落の配置を保持
                        paragraph.alignment = paragraph.alignment

# プレゼンテーションを開く
prs = Presentation('Python講座 辞書編.pptx')  # プレゼンテーションファイルを開く

# 元のサイズを保存
original_width = prs.slide_width  # 元のスライドの幅を保存
original_height = prs.slide_height  # 元のスライドの高さを保存

# 新しいサイズを設定（例：16:9のワイドスクリーン）
new_width = Inches(16)  # 新しいスライドの幅を設定
new_height = Inches(9)  # 新しいスライドの高さを設定

# スケール係数を計算
scale_width = new_width / original_width  # 幅のスケール係数を計算
scale_height = new_height / original_height  # 高さのスケール係数を計算

# スライドサイズを変更
prs.slide_width = new_width  # スライドの幅を新しいサイズに変更
prs.slide_height = new_height  # スライドの高さを新しいサイズに変更

# 各スライドのコンテンツをスケーリング
for slide in prs.slides:  # プレゼンテーション内の各スライドに対して
    scale_slide_contents(slide, scale_width, scale_height)  # スライドの内容をスケーリング

# 変更を保存
prs.save('resized_and_scaled_example.pptx')  # 変更を保存
