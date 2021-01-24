# pip install Pillow
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

IMG_PATH = "../data/in/スマートグラス.jpg"
IMG_PATH2 = "../data/in/ヨガ２.jpg"


TITLE_SLIDE = prs.slide_layouts[0]
FREE_SLIDE = prs.slide_layouts[1]
BRANK_SLIDE = prs.slide_layouts[6]


def add_title(slide, text):
    title = slide.shapes.title
    title.text = text

def add_image(slide, IMG_PATH, top = 0.1, right = 0, bottom = 0, left = 0, padding = 0.05):
    # スライドサイズ取得
    SLIDE_HEIGHT = prs.slide_height
    SLIDE_WIDTH = prs.slide_width 

    # スライドの画像を挿入する箇所のサイズを取得
    margin_top = SLIDE_HEIGHT * top
    margin_bottom = SLIDE_HEIGHT * bottom
    margin_left = SLIDE_WIDTH * left
    margin_right = SLIDE_WIDTH * right

    # 余白を除いた挿入サイズを取得
    slide_diplay_heigth = SLIDE_HEIGHT - margin_top - margin_bottom 
    slide_diplay_width = SLIDE_WIDTH - margin_left - margin_right

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
    left =((slide_diplay_width - img_display_width ) / 2) + margin_left
    top = ((slide_diplay_heigth - img_display_height ) / 2) + margin_top
    #画像をスライドに追加
    slide.shapes.add_picture(IMG_PATH, left, top, height = img_display_height)

#スライドオブジェクトの定義
prs = Presentation() 

slide = prs.slides.add_slide(TITLE_SLIDE)
add_title(slide, 'My slide' )

slide = prs.slides.add_slide(FREE_SLIDE)
add_title(slide, 'これは２枚目' )

#スライドサイズを取得

IMG_PATH = "../data/in/スマートグラス.jpg"
add_image(slide, IMG_PATH, right = 0.5)
IMG_PATH2 = "../data/in/ヨガ２.jpg"
add_image(slide, IMG_PATH2,  left = 0.5)

#白紙のスライドを追加
slide = prs.slides.add_slide(BRANK_SLIDE)
add_image(slide, IMG_PATH, right = 0.6, bottom = 0.3)
add_image(slide, IMG_PATH2, right = 0.5, top = 0.6)
add_image(slide, IMG_PATH, left = 0.5,  bottom = 0.3)
add_image(slide, IMG_PATH2, left = 0.6, top = 0.6)

#白紙のスライドを追加
slide = prs.slides.add_slide(BRANK_SLIDE)

left = top = width = height = Inches(1)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

#スライドを出力
prs.save('../data/out/test2.pptx')


