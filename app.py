from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from flask_cors import CORS
import io

app = Flask(__name__)
CORS(app)

@app.route('/generate', methods=['POST'])
def generate_ppt():
    print("✅ 收到請求了！")
    data = request.get_json()
    print("📄 歌詞內容：", data)

    title = data.get("title", "").strip()
    lyricist = data.get("lyricist", "").strip()
    composer = data.get("composer", "").strip()
    singer = data.get("singer", "").strip()
    lyrics = data.get("lyrics", "").strip()

    lines = [line.strip() for line in lyrics.splitlines() if line.strip()]

    prs = Presentation()
    prs.slide_width = Inches(13.33)  # 16:9
    prs.slide_height = Inches(7.5)

    def add_cover_slide(title, lyricist, composer, singer):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        # 中央文字方塊
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.33), Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        content = []
        if title:
            content.append(title)
        if lyricist:
            content.append(f"作詞：{lyricist}")
        if composer:
            content.append(f"作曲：{composer}")
        if singer:
            content.append(f"演唱：{singer}")

        for line in content:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.size = Pt(60)
            run.font.bold = True
            run.font.name = 'Microsoft JhengHei'
            run.font.color.rgb = RGBColor(255, 255, 255)

    def add_lyrics_slide(text_lines):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.33), Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        for line in text_lines:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.size = Pt(60)
            run.font.bold = True
            run.font.name = 'Microsoft JhengHei'
            run.font.color.rgb = RGBColor(255, 255, 255)

    # 先插入封面
    add_cover_slide(title, lyricist, composer, singer)

    # 接著每 4 行歌詞一頁
    for i in range(0, len(lines), 4):
        add_lyrics_slide(lines[i:i+4])

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"{title or 'lyrics'}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

if __name__ == '__main__':
    app.run(debug=True)
