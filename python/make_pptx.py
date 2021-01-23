# pip install Pillow
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

IMG_PATH = "../data/in/スマートグラス.jpg"
IMG_PATH2 = "../data/in/ヨガ２.jpg"
IMG_DISPLAY_HEIGHT = Inches(3) #スライドに表示するときの画像の高さ。とりあえず3インチとしておく。
# SLIDE_OUTPUT_PATH = "test.pptx" #スライドの出力先パス

#スライドオブジェクトの定義
prs = Presentation() 

TITLE_SLIDE = 0
FREE_SLIDE = 1
BRANK_SLIDE = 6


def add_slide_title(prs, lay_out, text):
    slide_layout = prs.slide_layouts[lay_out]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = text


add_slide_title(prs, TITLE_SLIDE, 'My slide' )
add_slide_title(prs, FREE_SLIDE, 'これは２枚目' )
# prs.save('../data/out/test.pptx')

#スライドサイズを取得
#
prs = Presentation() 

def add_image(slide, IMG_PATH, left_margin = 0, top_margin = 0.5, padding = 0.05):
    # スライドサイズ取得
    SLIDE_WIDTH = prs.slide_width 
    SLIDE_HEIGHT = prs.slide_height

    # スライドの画像を挿入する箇所のサイズを取得
    top_margin = Inches(top_margin)
    left_margin = Inches(left_margin)
    slide_diplay_heigth = SLIDE_HEIGHT-top_margin
    slide_diplay_width = SLIDE_WIDTH-left_margin

    # スライドの縦横比を計算
    slide_aspect_ratio = slide_diplay_width/slide_diplay_heigth
    
    #画像サイズを取得してアスペクト比を得る
    im = Image.open(IMG_PATH)
    im_width, im_height = im.size
    aspect_ratio = im_width/im_height

# 縦横比を比較して、スライドよリも横長なら横を基準にする
    if aspect_ratio >= slide_aspect_ratio:
        left_padding = (slide_diplay_width) * padding
        # 左右にpaddingが入るため、横さから、padding*2を引いた値が画像の高さになる。
        img_display_width = slide_diplay_width - (left_padding * 2)
        # 縦横比を使って、画像の縦を計算
        img_display_height = img_display_width / aspect_ratio
    else: # スライドよリも縦長なら縦を基準にする
        top_padding = (slide_diplay_heigth) * padding
        # 上下にpaddingが入るため、高さから、padding*2を引いた値が画像の高さになる。
        img_display_height = slide_diplay_heigth - (top_padding * 2)
        # 縦横比を使って、画像の横を計算
        img_display_width = aspect_ratio*img_display_height
    #センタリングする場合の画像の左上座標を計算
    left =((slide_diplay_width - img_display_width ) / 2) + left_margin
    top = ((slide_diplay_heigth - img_display_height ) / 2) + top_margin
    #画像をスライドに追加
    slide.shapes.add_picture(IMG_PATH, left, top, height = img_display_height)


prs = Presentation() 
#白紙のスライドを追加
blank_slide_layout = prs.slide_layouts[6] 
slide = prs.slides.add_slide(blank_slide_layout)
IMG_PATH = "../data/in/スマートグラス.jpg"
add_image(slide, IMG_PATH)


blank_slide_layout = prs.slide_layouts[6] 
slide = prs.slides.add_slide(blank_slide_layout)
IMG_PATH2 = "../data/in/ヨガ２.jpg"
add_image(slide, IMG_PATH2)

#スライドを出力
prs.save('../data/out/test2.pptx')


