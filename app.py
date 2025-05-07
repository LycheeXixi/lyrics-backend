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
    data = request.get_json()

    title = data.get("title", "").strip()
    lyricist = data.get("lyricist", "").strip()
    composer = data.get("composer", "").strip()
    singer = data.get("singer", "").strip()
    lyrics = data.get("lyrics", "").strip()

    # 分段落处理，空行分段
    paragraphs = []
    current = []
    for line in lyrics.splitlines():
        if line.strip() == "":
            if current:
                paragraphs.append(current)
                current = []
        else:
            current.append(line.strip())
    if current:
        paragraphs.append(current)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    def add_cover_slide(title, lyricist, composer, singer):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

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
            content.append(f"原唱：{singer}")

        for line in content:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.size = Pt(60)
            run.font.bold = True
            run.font.name = 'Microsoft JhengHei'
            run.font.color.rgb = RGBColor(255, 255, 255)

    def add_lyrics_slide(lines):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.33), Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        for line in lines:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line
            run.font.size = Pt(60)
            run.font.bold = True
            run.font.name = 'Microsoft JhengHei'
            run.font.color.rgb = RGBColor(255, 255, 255)

    # 添加封面
    add_cover_slide(title, lyricist, composer, singer)

    # 每段落处理，每段最多 4 行，超过继续分页
    for para in paragraphs:
        for i in range(0, len(para), 4):
            add_lyrics_slide(para[i:i + 4])

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
