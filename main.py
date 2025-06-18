from datetime import datetime, timedelta
import os
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from threading import Thread
import time
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

# === Flask Initialization ===
app = Flask(__name__)
CORS(app)
UPLOAD_FOLDER = 'static'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# === Scheduled File Cleanup ===
def delete_old_files():
    while True:
        now = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(file_path) and now - os.path.getctime(file_path) > 600:
                os.remove(file_path)
        time.sleep(600)

Thread(target=delete_old_files, daemon=True).start()

# === Utility Functions for Slide Formatting ===
def add_title_slide(prs, title_text):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    title.text_frame.paragraphs[0].font.size = Pt(40)
    title.text_frame.paragraphs[0].font.bold = True

def add_content_slide(prs, page, is_ending=False):
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    left, top = Inches(0.5), Inches(0.5)
    width, height = Inches(9), Inches(6.5)
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.clear()

    title_run = tf.paragraphs[0].add_run()
    title_run.text = page.get("h2", "")
    title_run.font.size = Pt(28)
    title_run.font.bold = True

    for section in page.get("sections", []):
        p = tf.add_paragraph()
        p.space_before = Pt(10)
        h3 = section.get("h3", "")
        text = section.get("p", "")

        if h3:
            run_h3 = p.add_run()
            run_h3.text = f"{h3}\n"
            run_h3.font.size = Pt(20)
            run_h3.font.bold = True

        if text:
            run_p = tf.add_paragraph().add_run()
            run_p.text = text
            run_p.font.size = Pt(16)

    if is_ending:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.italic = True

# === Route to Generate PowerPoint ===
@app.route('/generate_pptx', methods=['POST'])
def generate_pptx():
    token = request.json.get("token")
    if token != "fattycat0401":
        return jsonify({"error": "Invalid token"}), 403

    data = request.json
    h1 = data.get("h1", "Untitled Presentation")
    pages = data.get("pages", [])

    prs = Presentation()
    add_title_slide(prs, h1)

    for i, page in enumerate(pages):
        is_ending = (i == len(pages) - 1)
        add_content_slide(prs, page, is_ending=is_ending)

    filename = f"presentation_{datetime.now().strftime('%Y%m%d%H%M%S')}.pptx"
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    prs.save(save_path)

    download_url = f"https://gpts-slide-api.onrender.com/static/{filename}"
    return jsonify({"download_url": download_url})

# === Serve Static Files ===
@app.route('/static/<path:filename>')
def serve_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

# === Run App with Render-Compatible Port Binding ===
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
