from flask import Flask, request, send_from_directory, jsonify
from pptx import Presentation
import os
import time
import threading
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

TEMPLATE_MAP = {
    "simple_red_gray": "簡潔紅灰風.pptx",
    "blue_watercolor": "藍色淡雅水彩風.pptx",
    "blue_polygon": "藍色多邊形風.pptx",
    "gray_modern_euro": "灰色雅緻歐美風.pptx",
    "cafe_business": "咖啡店商務風.pptx",
    "blue_grid_world": "藍色網格世界風.pptx",
    "black_simple_business": "黑色磨砂質感風.pptx",
    "cartoon": "卡通風.pptx",
    "computer_theme": "電腦題材風.pptx",
    "colorful_startup": "彩色背景創業風.pptx",
    "technology": "科技風.pptx",
    "western_business": "歐美商務風.pptx",
    "simple_white_blue_text": "簡約風_白底藍字.pptx",
    "simple_black_white_text": "簡約風_黑底白字.pptx",
    "cultural_chinese": "水墨中國風.pptx"
}

TEMPLATE_FILES = list(TEMPLATE_MAP.values())

def cleanup_expired_files_safe(folder, keep_filenames, expiry_seconds=600):
    now = time.time()
    for filename in os.listdir(folder):
        filepath = os.path.join(folder, filename)
        if (
            os.path.isfile(filepath)
            and filename.endswith(".pptx")
            and filename not in keep_filenames
        ):
            if now - os.path.getmtime(filepath) > expiry_seconds:
                try:
                    os.remove(filepath)
                    print(f"✅ 自動刪除過期檔案：{filename}")
                except Exception as e:
                    print(f"❌ 無法刪除 {filename}：{e}")

def start_safe_cleanup_thread(folder="static"):
    def loop():
        while True:
            cleanup_expired_files_safe(folder, keep_filenames=TEMPLATE_FILES)
            time.sleep(60)
    threading.Thread(target=loop, daemon=True).start()

start_safe_cleanup_thread("static")

def clone_slide(template_slide, output_presentation):
    blank_layout = output_presentation.slide_layouts[6]  # 使用空白頁面
    new_slide = output_presentation.slides.add_slide(blank_layout)

    for shape in template_slide.shapes:
        try:
            new_shape = new_slide.shapes._spTree.insert_element_before(shape.element, 'p:extLst')
        except Exception as e:
            print(f"❗️無法複製形狀：{e}")
    return new_slide

@app.route("/generate_pptx", methods=["POST"])
def generate_pptx():
    data = request.get_json()
    token = data.get("token")
    if token != "fattycat0401":
        return jsonify({"error": "Invalid token"}), 403

    template_key = data.get("template_style")
    slides_content = data.get("slides", [])
    template_file = TEMPLATE_MAP.get(template_key)
    if not template_file:
        return jsonify({"error": "Invalid template style"}), 400

    template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file)
    prs = Presentation(template_path)

    template_mapping = {
        "title": prs.slides[0],
        "content": prs.slides[1],
        "closing": prs.slides[2]
    }

    title_page = next((p for p in slides_content if p["type"] == "title"), None)
    closing_page = next((p for p in slides_content if p["type"] == "closing"), None)
    content_pages = [p for p in slides_content if p["type"] == "content"]

    ordered_slides = []
    if title_page: ordered_slides.append(title_page)
    ordered_slides.extend(content_pages)
    if closing_page: ordered_slides.append(closing_page)

    output = Presentation()
    output.slides._sldIdLst.clear()

    for page in ordered_slides:
        page_type = page.get("type")
        page_text = page.get("text", "")
        if page_type not in template_mapping:
            continue
        slide = clone_slide(template_mapping[page_type], output)
        for shape in slide.shapes:
            if shape.has_text_frame:
                shape.text = page_text
                break
# 不再手動設定 slide ID，避免錯誤
# output.slides._sldIdLst[-1].element.set("id", str(uuid.uuid4().int & (1<<32)-1))

    filename = f"slide_{template_key}_{int(time.time())}.pptx"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    output.save(filepath)
    return jsonify({"download_url": f"/static/{filename}"}), 200

@app.route('/static/<path:filename>', methods=['GET'])
def serve_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route("/", methods=["GET"])
def index():
    return "✅ SlideCrafter API is running. Use POST /generate_pptx to create a presentation."

if __name__ == "__main__":
    print("✅ SlideCrafter API is running.")
    app.run(host="0.0.0.0")