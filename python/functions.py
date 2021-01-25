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
        padding = slide_diplay_width * padding
        # 左右にpaddingが入るため、横さから、padding*2を引いた値が画像の高さになる。
        img_display_width = slide_diplay_width - (padding * 2)
        # 縦横比を使って、画像の縦を計算
        img_display_height = img_display_width / aspect_ratio
    else: # スライドよリも縦長なら縦を基準にする
        padding = (slide_diplay_heigth) * padding
        # 上下にpaddingが入るため、高さから、padding*2を引いた値が画像の高さになる。
        img_display_height = slide_diplay_heigth - (padding * 2)
        # 縦横比を使って、画像の横を計算
        img_display_width = aspect_ratio*img_display_height
    #センタリングする場合の画像の左上座標を計算
    left =((slide_diplay_width - img_display_width ) / 2) + margin_left
    top = ((slide_diplay_heigth - img_display_height ) / 2) + margin_top
    #画像をスライドに追加
    slide.shapes.add_picture(IMG_PATH, left, top, height = img_display_height)

def add_text(slide, msg, top = 0.1, right = 0, bottom = 0, left = 0, padding = 0.05, font_size = 16):
    # スライドサイズ取得
    SLIDE_HEIGHT = prs.slide_height
    SLIDE_WIDTH = prs.slide_width 

    # テキストを挿入する箇所のサイズを取得
    margin_top = SLIDE_HEIGHT * top
    margin_bottom = SLIDE_HEIGHT * bottom
    margin_left = SLIDE_WIDTH * left
    margin_right = SLIDE_WIDTH * right

    # 余白を除いた挿入サイズを取得
    slide_diplay_heigth = SLIDE_HEIGHT - margin_top - margin_bottom 
    slide_diplay_width = SLIDE_WIDTH - margin_left - margin_right
    

    # 左右にpaddingが入るため、横から、padding*2を引いた値がが幅になる。
    width_padding = slide_diplay_width * padding
    text_display_width = slide_diplay_width - (width_padding * 2)

    
    # 上下にpaddingが入るため、高さから、padding*2を引いた値が高さになる。
    height_padding = slide_diplay_width * padding
    text_display_height = slide_diplay_heigth - (height_padding * 2)
    
    # 座標計算
    left =((slide_diplay_width - text_display_width ) / 2) + margin_left
    top = ((slide_diplay_heigth - text_display_height ) / 2) + margin_top

    
    textbox = slide.shapes.add_textbox(left, top, text_display_width, text_display_height)
    textFrame = textbox.text_frame
    textFrame.word_wrap = True
    p = textFrame.add_paragraph()
    p.text = msg
    p.font.size = Pt(font_size)
