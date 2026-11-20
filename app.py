from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import re
import os
from datetime import datetime
from dotenv import load_dotenv
import google.generativeai as genai
from docx import Document
from pypdf import PdfReader

# 1. Cấu hình biến môi trường & Gemini API
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

model = None
if API_KEY:
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel("gemini-3.5-flash-lite")
    except Exception as e:
        print(f"Lỗi khởi tạo Gemini API: {e}")
        model = None

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'tailieu')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DATA_FILE = os.path.join(BASE_DIR, "data.json")

# 🔒 MẬT MÃ QUẢN TRỊ DÀNH CHO GIÁO VIÊN
ADMIN_SECRET_KEY = os.getenv("ADMIN_KEY", "thptabc@2026")

def load_knowledge():
    """Luôn đọc lại dữ liệu mới nhất từ data.json"""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Lỗi đọc data.json: {e}")
            return []
    return []

STOP_WORDS = {
    "là", "gì", "như", "thế", "nào", "sao", "làm", "được", "không", "có", "và", 
    "của", "cho", "các", "những", "trong", "khi", "ở", "về", "đến", "tại", 
    "học", "sinh", "trường", "thpt", "em", "bạn", "thầy", "cô", "nhà"
}

def clean_words(text):
    """Tách từ khóa quan trọng từ câu hỏi"""
    words = re.findall(r'\w+', str(text).lower())
    return [w for w in words if w not in STOP_WORDS and len(w) > 1]

def safe_str(val):
    """Ép kiểu an toàn tránh lỗi nối chuỗi TypeError"""
    if val is None:
        return ""
    if isinstance(val, list):
        return " ".join([str(item) for item in val if item is not None])
    return str(val).strip()

def parse_date_to_object(date_str):
    """Chuyển chuỗi ngày thành đối tượng date để so sánh với thời gian thực"""
    if not date_str:
        return None
    
    match_end = re.search(r'đến(?:\s+hết)?\s+ngày\s+(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str, re.IGNORECASE)
    if match_end:
        d, m, y = map(int, match_end.groups())
        try:
            return datetime(y, m, d).date()
        except ValueError:
            pass

    match_slash = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
    if match_slash:
        d, m, y = map(int, match_slash.groups())
        try:
            return datetime(y, m, d).date()
        except ValueError:
            pass

    match_vn = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})(?:\s+năm\s+(\d{4}))?', date_str, re.IGNORECASE)
    if match_vn:
        d = int(match_vn.group(1))
        m = int(match_vn.group(2))
        y = int(match_vn.group(3)) if match_vn.group(3) else datetime.now().year
        try:
            return datetime(y, m, d).date()
        except ValueError:
            pass

    return None

def extract_dates(text):
    """Quét văn bản để tìm khoảng thời gian tổ chức hoặc hạn chót"""
    date_patterns = [
        r'(từ ngày\s+\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?\s+đến(?:\s+hết)?\s+ngày\s+\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
        r'(ngày\s+\d{1,2}\s+tháng\s+\d{1,2}(?:\s+năm\s+\d{4})?)',
        r'(\d{1,2}/\d{1,2}/\d{4})',
        r'(học kỳ\s+[iI12]{1,2}(?:\s+năm học\s+\d{4}[–-]\d{4})?)',
        r'(năm học\s+\d{4}[–-]\d{4})'
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(0).strip()
            
    return "Thực hiện theo kế hoạch/công văn của nhà trường"

def check_expiration(date_text):
    """Kiểm tra thời hạn sự kiện so với ngày hiện tại của máy tính"""
    target_date = parse_date_to_object(date_text)
    today = datetime.now().date()
    
    if target_date:
        if today > target_date:
            return True, f"ĐÃ KẾT THÚC / QUÁ HẠN (Thời gian diễn ra/Hạn chót: {target_date.strftime('%d/%m/%Y')})"
        elif today == target_date:
            return False, f"ĐANG DIỄN RA HÔM NAY ({target_date.strftime('%d/%m/%Y')})"
        else:
            return False, f"SẮP DIỄN RA (Thời hạn đến ngày {target_date.strftime('%d/%m/%Y')})"
    
    return False, date_text

# ----------------------------------------------------
# 1. Định tuyến giao diện & Xác thực cổng giáo viên
# ----------------------------------------------------
@app.route('/', methods=['GET'])
def home_page():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/admin', methods=['GET'])
def admin_page():
    return send_from_directory(BASE_DIR, 'admin.html')

# Tuyến đường kiểm tra mật khẩu nhanh cho nút bấm trên trang index
@app.route('/api/verify-admin', methods=['POST'])
def verify_admin():
    data = request.json or {}
    password = data.get("password", "")
    if password == ADMIN_SECRET_KEY:
        return jsonify({"success": True})
    return jsonify({"success": False, "message": "Mật mã quản trị không chính xác!"}), 403

@app.route('/tailieu/<path:filename>', methods=['GET'])
def serve_file(filename):
    # Xử lý chuẩn hóa tên file để tránh lỗi dấu gạch chéo ngược trên Windows
    safe_filename = filename.replace("\\", "/").split("/")[-1]
    return send_from_directory(UPLOAD_FOLDER, safe_filename, as_attachment=False)

# ----------------------------------------------------
# 2. Tuyến đường Hỏi đáp (/ask)
# ----------------------------------------------------
@app.route("/ask", methods=["POST"])
def ask():
    current_knowledge = load_knowledge()
    data = request.json or {}
    user_question = data.get("question", "").strip()
    user_question_lower = user_question.lower()

    if not user_question:
        return jsonify({"answer": "Vui lòng nhập câu hỏi của bạn.", "source": None, "file_url": None, "time_info": None})

    keywords_rieng_tu = ["điểm số", "điểm thi", "kỷ luật cá nhân", "học bạ", "sức khỏe", "tiền riêng"]
    if any(kw in user_question_lower for kw in keywords_rieng_tu):
        return jsonify({
            "answer": "Hệ thống chỉ hỗ trợ tra cứu thông tin chung của nhà trường và không xử lý thông tin cá nhân (như điểm số, kỷ luật). Bạn vui lòng liên hệ giáo viên chủ nhiệm hoặc ban giám hiệu nhé.",
            "source": "Quy chế bảo vệ dữ liệu học đường",
            "file_url": None,
            "time_info": None
        })

    query_keywords = clean_words(user_question)
    best_doc = None
    max_score = 0

    DAC_THU = ["ma túy", "hue-s", "trắc nghiệm", "bạo lực", "giao thông", "stem", "phù đổng", "tri ân", "hướng nghiệp", "đọc sách", "môi trường"]

    for doc in current_knowledge:
        ten_tl = safe_str(doc.get("ten_tai_lieu", ""))
        noi_dung = safe_str(doc.get("noi_dung", ""))
        thoi_gian = safe_str(doc.get("thoi_gian_to_chuc", ""))

        doc_text = f"{ten_tl} {noi_dung} {thoi_gian}".lower()
        score = sum(1 for kw in query_keywords if kw in doc_text)
        
        for dt in DAC_THU:
            if dt in user_question_lower and dt in doc_text:
                score += 5

        if any(kw in ten_tl.lower() for kw in query_keywords):
            score += 3

        if any(time_kw in user_question_lower for time_kw in ["khi nào", "thời gian", "ngày", "bao giờ", "hết hạn", "kết thúc"]):
            if any(kw in thoi_gian.lower() for kw in query_keywords):
                score += 2

        if doc.get("id", 0) > 20 and score > 0:
            score += 1

        if score > max_score:
            max_score = score
            best_doc = doc

    if best_doc and max_score >= 1:
        ten_tl_best = safe_str(best_doc.get("ten_tai_lieu", ""))
        noi_dung_best = safe_str(best_doc.get("noi_dung", ""))
        raw_time = safe_str(best_doc.get("thoi_gian_to_chuc", ""))
        
        is_expired, status_time = check_expiration(raw_time)

        expired_note = ""
        if is_expired:
            expired_note = f"\n⚠️ [LƯU Ý: Sự kiện/Kế hoạch này {status_time}. Hãy nêu rõ cho bạn học sinh biết thời gian đã kết thúc, sau đó tóm tắt nội dung kế hoạch đã lưu.]"

        if model:
            prompt = f"""
Bạn là Trợ lý AI học đường thân thiện của Trường THPT ABC.
Khi trả lời, bạn hãy xưng là "mình" và gọi người hỏi là "bạn" (tuyệt đối không xưng "em", "tôi", "thầy/cô").

Dưới đây là tài liệu chính thức từ nhà trường:
---
Tài liệu: {ten_tl_best}
Thời gian/Trạng thái: {status_time} {expired_note}
Nội dung: {noi_dung_best}
---

Câu hỏi của học sinh: {user_question}

YÊU CẦU:
1. Trả lời rõ ràng, gần gũi, đúng trọng tâm, dựa hoàn toàn vào nội dung văn bản trên.
2. Nếu văn bản ĐÃ HẾT HẠN / ĐÃ KẾT THÚC: Hãy thông báo rõ cho bạn học sinh biết là sự kiện/cuộc thi này đã kết thúc vào ngày đó trước khi tóm tắt nội dung.
3. Không tự suy diễn hay bịa đặt thêm quy định ngoài tài liệu.
"""
            try:
                ai_answer = model.generate_content(prompt).text.strip()
            except Exception as e:
                print(f"Lỗi gọi Gemini: {e}")
                if is_expired:
                    ai_answer = f"⚠️ Chào bạn, sự kiện này {status_time}.\n\nNội dung chi tiết của kế hoạch:\n{noi_dung_best}"
                else:
                    ai_answer = f"Chào bạn! Dưới đây là thông tin chi tiết:\n{noi_dung_best}\n\n📅 Thời gian: {status_time}"
        else:
            if is_expired:
                ai_answer = f"⚠️ Chào bạn, sự kiện này {status_time}.\n\nNội dung chi tiết của kế hoạch:\n{noi_dung_best}"
            else:
                ai_answer = f"Chào bạn! Dưới đây là thông tin chi tiết:\n{noi_dung_best}\n\n📅 Thời gian: {status_time}"

        file_path = safe_str(best_doc.get("duong_dan_file", ""))
        if file_path:
            clean_path = file_path.replace("\\", "/").lstrip("/")
            full_file_url = f"/{clean_path}"
        else:
             full_file_url = None 

        return jsonify({
            "answer": ai_answer,
            "source": ten_tl_best,
            "file_url": full_file_url,
            "time_info": status_time,
            "is_expired": is_expired
        })
    else:
        return jsonify({
            "answer": "Hiện tại mình chưa tìm thấy thông tin phù hợp trong các văn bản đã ban hành của nhà trường để giải đáp câu hỏi này. Bạn vui lòng kiểm tra lại bảng tin chính thức hoặc hỏi văn phòng trường nhé!",
            "source": None,
            "file_url": None,
            "time_info": None,
            "is_expired": False
        })

# ----------------------------------------------------
# 3. API Tải tệp lên (/upload) - BẢO MẬT & BÓC TÁCH CHUẨN
# ----------------------------------------------------
@app.route("/upload", methods=["POST"])
def upload_file():
    # 🔒 BẢO VỆ MÁY CHỦ: BẮT BUỘC CÓ MẬT MÃ QUẢN TRỊ
    provided_key = request.headers.get("X-Admin-Key") or request.form.get("admin_key")
    if provided_key != ADMIN_SECRET_KEY:
        return jsonify({
            "success": False, 
            "message": "Từ chối truy cập: Bạn không có quyền tải tài liệu lên hệ thống!"
        }), 403

    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Không tìm thấy tệp gửi lên."}), 400
    
    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({"success": False, "message": "Chưa chọn tệp tin."}), 400

    filename = file.filename
    save_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(save_path)

    doc_text = ""
    title_text = f"Văn bản: {filename}"

    # 1. Đọc tệp Word (.docx)
    if filename.lower().endswith(".docx"):
        try:
            doc = Document(save_path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            if paragraphs:
                title_text = paragraphs[0][:90].strip()
                doc_text = " ".join(paragraphs)
        except Exception as e:
            return jsonify({"success": False, "message": f"Lỗi đọc file Word: {e}"}), 500

    # 2. Đọc tệp PDF (.pdf)
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
                lines = [l.strip() for l in pdf_pages_text.splitlines() if l.strip()]
                if lines:
                    title_text = lines[0][:90].strip()
        except Exception as e:
            return jsonify({"success": False, "message": f"Lỗi đọc file PDF: {e}"}), 500

    else:
        doc_text = f"Thông báo: {filename}"

    extracted_time = extract_dates(doc_text)
    is_exp, status_str = check_expiration(extracted_time)

    current_knowledge = load_knowledge()

    new_item = {
        "id": len(current_knowledge) + 1,
        "ten_tai_lieu": str(title_text),
        "duong_dan_file": f"tailieu/{filename}",
        "thoi_gian_to_chuc": str(extracted_time),
        "chu_de": "Thông báo cập nhật mới",
        "noi_dung": str(doc_text) if doc_text else f"Văn bản {filename} vừa được cập nhật vào hệ thống."
    }

    current_knowledge.append(new_item)
    
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(current_knowledge, f, ensure_ascii=False, indent=2)

    return jsonify({
        "success": True, 
        "message": "Đã trích xuất và nạp văn bản thành công!",
        "filename": filename,
        "title": title_text,
        "time_info": status_str
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
