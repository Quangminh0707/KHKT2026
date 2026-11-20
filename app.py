import os
import json
import base64
import re
from flask import Flask, request, jsonify, render_template, send_from_directory
from flask_cors import CORS
import google.generativeai as genai
from docx import Document
from pypdf import PdfReader
from dotenv import load_dotenv
import requests

# 1. NẠP CẤU HÌNH BIẾN MÔI TRƯỜNG (.env)
load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'data.json')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'tailieu')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Khởi tạo Flask App
app = Flask(__name__, template_folder=BASE_DIR, static_folder=BASE_DIR)
CORS(app)

# 2. CẤU HÌNH GEMINI & MẬT MÃ QUẢN TRỊ
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
else:
    print("⚠️ CẢNH BÁO: Chưa tìm thấy GEMINI_API_KEY trong biến môi trường!")

# Mật mã quản trị giáo viên (mặc định: thptabc@2026)
ADMIN_SECRET_KEY = os.getenv("ADMIN_KEY", "thptabc@2026")

# Cấu hình tự động đẩy dữ liệu lên GitHub (Giải pháp 2)
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPO = os.getenv("GITHUB_REPO")      # Dạng: username/repo-name
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main")

# ==========================================================
# CÁC HÀM TIỆN ÍCH HỖ TRỢ (UTILITIES)
# ==========================================================

def load_data():
    """Đọc dữ liệu văn bản từ file data.json an toàn"""
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception as e:
        print(f"Lỗi khi đọc {DATA_FILE}: {e}")
        return []

def save_data(data):
    """Ghi danh sách văn bản vào data.json"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Lỗi khi ghi {DATA_FILE}: {e}")
        return False

def push_to_github(file_path_in_repo, content_bytes, commit_message):
    """Tự động đẩy file lên GitHub qua GitHub REST API (Giải pháp 2)"""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        # Nếu chưa cấu hình GitHub Token thì chỉ lưu máy chủ cục bộ, không báo lỗi
        return False

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_path_in_repo}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }

    try:
        # Kiểm tra file đã có trên GitHub chưa để lấy SHA nếu là cập nhật
        sha = None
        res_check = requests.get(url, headers=headers, timeout=10)
        if res_check.status_code == 200:
            sha = res_check.json().get("sha")

        payload = {
            "message": commit_message,
            "content": base64.b64encode(content_bytes).decode("utf-8"),
            "branch": GITHUB_BRANCH
        }
        if sha:
            payload["sha"] = sha

        res_put = requests.put(url, headers=headers, json=payload, timeout=15)
        if res_put.status_code in [200, 201]:
            print(f"✅ Đã đồng bộ lên GitHub: {file_path_in_repo}")
            return True
        else:
            print(f"⚠️ Không thể đồng bộ GitHub ({res_put.status_code}): {res_put.text}")
            return False
    except Exception as e:
        print(f"⚠️ Lỗi kết nối GitHub API: {e}")
        return False

# ==========================================================
# CÁC ROUTE PHỤC VỤ GIAO DIỆN & TÀI LIỆU
# ==========================================================

@app.route('/')
def home():
    """Trang giao diện trò chuyện chính của học sinh"""
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/admin')
def admin_page():
    """Trang giao diện Cổng Giáo Viên"""
    return send_from_directory(BASE_DIR, 'admin.html')

@app.route('/tailieu/<path:filename>', methods=['GET'])
def serve_file(filename):
    """Phục vụ file gốc Word/PDF cho khung đối soát văn bản"""
    # Chuẩn hóa tên file, tránh dấu gạch chéo ngược Windows
    safe_filename = filename.replace("\\", "/").split("/")[-1]
    return send_from_directory(UPLOAD_FOLDER, safe_filename, as_attachment=False)

@app.route('/api/verify-admin', methods=['POST'])
def verify_admin():
    """Xác thực mật khẩu khi bấm vào nút Cổng Giáo Viên trên web"""
    req = request.get_json(silent=True) or {}
    password = req.get("password", "").strip()
    if password == ADMIN_SECRET_KEY:
        return jsonify({"success": True})
    return jsonify({"success": False, "message": "Mật khẩu không chính xác!"}), 403

# ==========================================================
# CỔNG GIÁO VIÊN: TẢI VĂN BẢN & TRÍCH XUẤT (UPLOAD)
# ==========================================================

@app.route('/upload', methods=['POST'])
def upload_file():
    """Xử lý nạp văn bản Word (.docx) hoặc PDF (.pdf)"""
    # 1. Kiểm tra xác thực mật mã
    admin_key = request.headers.get("X-Admin-Key") or request.form.get("admin_key", "")
    if admin_key.strip() != ADMIN_SECRET_KEY:
        return jsonify({"success": False, "message": "Sai mật mã quản trị hoặc bạn không có quyền!"}), 403

    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Không tìm thấy tệp đính kèm!"}), 400

    file = request.files['file']
    filename = file.filename
    if not filename:
        return jsonify({"success": False, "message": "Tên tệp không hợp lệ!"}), 400

    # 2. Lưu file vào thư mục tailieu/
    save_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(save_path)

    doc_text = ""
    title_text = filename

    # 3. Đọc và trích xuất nội dung
    # 👉 Xử lý file Word (.docx) - ĐÃ SỬA LỖI SLICE INDEX:
    if filename.lower().endswith(".docx"):
        try:
            doc = Document(save_path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            if paragraphs:
                # Lấy đoạn văn đầu tiên (phần tử số 0), cắt tối đa 90 ký tự
                title_text = paragraphs[0][:90].strip()
                doc_text = " ".join(paragraphs)
        except Exception as e:
            return jsonify({"success": False, "message": f"Lỗi đọc file Word: {e}"}), 500

    # 👉 Xử lý file PDF (.pdf):
    elif filename.lower().endswith(".pdf"):
        try:
            reader = PdfReader(save_path)
            pdf_pages_text = []
            for p in reader.pages:
                extracted = p.extract_text()
                if extracted:
                    pdf_pages_text.append(extracted.strip())
            
            if pdf_pages_text:
                doc_text = " ".join(pdf_pages_text)
                lines = [l.strip() for l in pdf_pages_text[0].splitlines() if l.strip()]
                if lines:
                    title_text = lines[0][:90].strip()
        except Exception as e:
            return jsonify({"success": False, "message": f"Lỗi đọc file PDF: {e}"}), 500
    else:
        return jsonify({"success": False, "message": "Chỉ hỗ trợ file định dạng .docx hoặc .pdf!"}), 400

    if not doc_text:
        return jsonify({"success": False, "message": "Tệp không có nội dung chữ hoặc không trích xuất được!"}), 400

    # 4. Trích xuất thông tin ngày tháng bằng Regex
    time_match = re.search(r'(ngày\s+\d{1,2}\s+tháng\s+\d{1,2}(?:\s+năm\s+\d{4})?|từ\s+ngày.*?đến\s+ngày.*?\d{4}|tháng\s+\d{1,2}\s+năm\s+\d{4})', doc_text, re.IGNORECASE)
    time_info = time_match.group(0) if time_match else "Đang cập nhật"

    # 5. Lưu vào danh sách data.json
    db = load_data()
    # Đường dẫn file dạng tương đối chuẩn hóa 'tailieu/ten_file.ext'
    rel_file_path = f"tailieu/{filename}".replace("\\", "/")

    # Kiểm tra xem tài liệu cùng tên đã có chưa để cập nhật thay vì trùng lặp
    existing_index = next((i for i, item in enumerate(db) if item.get("duong_dan_file") == rel_file_path or item.get("tieu_de") == title_text), -1)

    new_entry = {
        "tieu_de": title_text,
        "noi_dung": doc_text,
        "thoi_gian": time_info,
        "duong_dan_file": rel_file_path
    }

    if existing_index >= 0:
        db[existing_index] = new_entry
    else:
        db.append(new_entry)

    save_data(db)

    # 6. GIẢI PHÁP 2: TỰ ĐỘNG ĐẨY FILE VÀ DATA.JSON LÊN GITHUB
    try:
        with open(save_path, "rb") as f:
            push_to_github(f"tailieu/{filename}", f.read(), f"Giao vien nap van ban: {filename}")
        with open(DATA_FILE, "rb") as f:
            push_to_github("data.json", f.read(), f"Cap nhat data.json sau khi them: {filename}")
    except Exception as e:
        print(f"Lỗi khi đồng bộ GitHub: {e}")

    return jsonify({
        "success": True,
        "message": "Nạp tài liệu và trích xuất dữ liệu thành công!",
        "title": title_text,
        "time_info": time_info
    })

# ==========================================================
# CỔNG HỌC SINH: HỎI ĐÁP VỚI TRỢ LÝ AI (ASK)
# ==========================================================

@app.route('/ask', methods=['POST'])
def ask():
    """Tiếp nhận câu hỏi từ học sinh, tìm kiếm tài liệu và hỏi Gemini"""
    req = request.get_json(silent=True) or {}
    question = req.get("question", "").strip()

    if not question:
        return jsonify({"answer": "Bạn hãy nhập câu hỏi để mình hỗ trợ nhé!", "source": None, "file_url": None})

    db = load_data()
    if not db:
        return jsonify({
            "answer": "Hiện tại hệ thống nhà trường chưa có dữ liệu văn bản nào. Thầy cô vui lòng vào Cổng Giáo Viên để nạp tài liệu nhé!",
            "source": None,
            "file_url": None
        })

    # Tìm kiếm văn bản liên quan dựa trên từ khóa câu hỏi
    query_words = [w.lower() for w in re.findall(r'\w+', question) if len(w) > 1]
    best_doc = None
    max_score = 0

    for doc in db:
        text = (doc.get("tieu_de", "") + " " + doc.get("noi_dung", "")).lower()
        score = sum(1 for word in query_words if word in text)
        if score > max_score:
            max_score = score
            best_doc = doc

    # Nếu không tìm thấy văn bản phù hợp
    if not best_doc or max_score == 0:
        return jsonify({
            "answer": "Xin lỗi bạn, mình chưa tìm thấy thông tin này trong các văn bản, kế hoạch chính thức hiện có của nhà trường.",
            "source": None,
            "file_url": None
        })

    # Chuẩn bị dữ liệu và URL văn bản gốc chuẩn hóa
    file_path = str(best_doc.get("duong_dan_file", "")).strip()
    if file_path:
        clean_path = file_path.replace("\\", "/").lstrip("/")
        full_file_url = f"/{clean_path}"
    else:
        full_file_url = None

    source_title = best_doc.get("tieu_de", "Văn bản nhà trường")

    # Tạo Prompt đối soát cho Gemini
    prompt = f"""Bạn là Trợ lý AI học đường thân thiện, chính xác của trường THPT.
Nhiệm vụ của bạn là giải đáp thắc mắc cho học sinh chỉ dựa trên thông tin chính thức dưới đây.

--- THÔNG TIN VĂN BẢN TRƯỜNG ---
Tiêu đề: {best_doc.get('tieu_de')}
Thời gian/Thời hạn: {best_doc.get('thoi_gian')}
Nội dung văn bản:
{best_doc.get('noi_dung')[:3500]}
--------------------------------

Câu hỏi của học sinh: "{question}"

Yêu cầu khi trả lời:
1. Xưng hô thân thiện ('mình' và 'bạn').
2. Trả lời đúng trọng tâm, ngắn gọn, dễ hiểu dựa trên văn bản trên.
3. Nếu văn bản không nhắc đến chi tiết học sinh hỏi, hãy nói rõ là văn bản chưa đề cập, không được tự suy đoán.
"""

    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(prompt)
        ai_reply = response.text.strip()
    except Exception as e:
        print(f"Lỗi gọi Gemini API: {e}")
        ai_reply = f"Dựa trên văn bản '{source_title}': Thời gian diễn ra là {best_doc.get('thoi_gian')}."

    return jsonify({
        "answer": ai_reply,
        "source": source_title,
        "file_url": full_file_url
    })

# ==========================================================
# KHỞI CHẠY MÁY CHỦ (HỖ TRỢ CẢ LOCALHOST VÀ RENDER)
# ==========================================================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
