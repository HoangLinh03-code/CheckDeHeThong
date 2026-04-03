import json
import re
from difflib import SequenceMatcher

# =====================================================================
# UTILS: HÀM HỖ TRỢ LÀM SẠCH VÀ SO SÁNH
# =====================================================================
def clean_text(text):
    """Lột HTML, chuẩn hóa LaTeX, Unicode Math và khoảng trắng để so sánh chính xác hơn."""
    if not text: return ""
    text = str(text)
    
    # 1. FIX LỖI HTML: Chỉ xóa thẻ HTML thực sự (bắt đầu bằng < và chữ cái hoặc / )
    # Điều này giúp giữ lại dấu nhỏ hơn < trong các bất phương trình Toán học.
    text = re.sub(r"<[a-zA-Z/][^>]*>", "", text)
    
    # 2. Thay thế các ký tự đặc biệt của HTML
    for entity, char in [("&nbsp;", " "), ("&lt;", "<"), ("&gt;", ">"), ("&amp;", "&"), ("\xa0", " ")]:
        text = text.replace(entity, char)
        
    # 3. CHUẨN HÓA UNICODE MATH -> LATEX
    unicode_math_map = {
        "ℝ": "\\mathbb{R}",
        "ℕ": "\\mathbb{N}",
        "ℤ": "\\mathbb{Z}",
        "ℚ": "\\mathbb{Q}",
        "∖": "\\setminus",
        "∈": "\\in",
        "∉": "\\notin",
        "∪": "\\cup",
        "∩": "\\cap",
        "∅": "\\emptyset",
        "≥": "\\ge",
        "≤": "\\le",
        "→": "\\rightarrow",
        "∞": "\\infty"
    }
    for uni_char, latex_code in unicode_math_map.items():
        text = text.replace(uni_char, latex_code)
        
    # 4. CHUẨN HÓA LATEX: Bỏ các lệnh bọc ngoặc vô nghĩa & Đồng nhất các ký hiệu
    text = text.replace("\\left", "")
    text = text.replace("\\right", "")
    text = text.replace("\\geqslant", "\\ge") 
    text = text.replace("\\leqslant", "\\le") 
    text = text.replace("\\geq", "\\ge") # Đồng nhất \geq thành \ge
    text = text.replace("\\leq", "\\le") # Đồng nhất \leq thành \le
    
    # 5. Xóa SẠCH toàn bộ dấu cách (vì trong công thức Toán dấu cách không có ý nghĩa)
    text = text.replace(" ", "")
    
    # 6. Cắt bỏ dấu chấm, phẩy ở tận cùng
    text = text.strip(".,;")
    
    # 7. Chuyển tất cả về chữ thường
    return text.lower()

def calc_similarity(a, b):
    """Tính toán phần trăm giống nhau giữa 2 chuỗi (0.0 đến 1.0)"""
    return SequenceMatcher(None, clean_text(a).lower(), clean_text(b).lower()).ratio()

def find_best_match(target, candidates, threshold=0.8):
    """Tìm chuỗi giống nhất trong danh sách candidates."""
    best_score = 0
    best_candidate = None
    for cand in candidates:
        score = calc_similarity(target, cand)
        if score > best_score:
            best_score = score
            best_candidate = cand
    if best_score >= threshold:
        return best_candidate, best_score
    return None, best_score

# =====================================================================
# CORE: LOGIC SO SÁNH
# =====================================================================
def load_system_data(file_path):
    """Đọc và chuẩn hóa cấu trúc từ hệ thống JSON."""
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f).get("data", [])
    
    sys_questions = []
    for block in data:
        # Hỗ trợ cả câu đơn (typeData=0) và đoạn văn (typeData=1)
        items = [block["dataStandard"]] if block.get("typeData") == 0 else block.get("dataMaterial", {}).get("data", [])
        
        for item in items:
            sys_questions.append({
                "content": clean_text(item.get("content", "")),
                "type": item.get("typeAnswer", 0),
                "options": {opt["idOption"]: clean_text(opt.get("content", "")) for opt in item.get("options", [])},
                "answerOptionId": item.get("answerOptionId", []),
                "answerFreeText": item.get("answerFreeText", [])
            })
    return sys_questions

def compare_exams(ai_json_path, sys_json_path):
    with open(ai_json_path, "r", encoding="utf-8") as f:
        ai_data = json.load(f)
        
    sys_data = load_system_data(sys_json_path)
    
    print(f"[*] Bắt đầu đối chiếu: {len(ai_data)} câu AI vs {len(sys_data)} câu Hệ thống\n")
    
    for ai_q in ai_data:
        ai_content = ai_q.get("noi_dung_cau_hoi", "")
        cau_so = ai_q.get("cau_so", "?")
        loai_cau = ai_q.get("loai_cau", "")
        
        # 1. TÌM CÂU HỎI TRÊN HỆ THỐNG CÓ NỘI DUNG GIỐNG NHẤT
        sys_contents = [q["content"] for q in sys_data]
        matched_content, score = find_best_match(ai_content, sys_contents)
        
        if not matched_content:
            print(f"❌ CÂU {cau_so}: KHÔNG TÌM THẤY CÂU NÀY TRÊN HỆ THỐNG! (Max match: {score*100:.1f}%)")
            continue
            
        # Lấy record của câu hỏi hệ thống đã match
        sys_q = next(q for q in sys_data if q["content"] == matched_content)
        
        # 2. ĐỐI CHIẾU ĐÁP ÁN DỰA TRÊN DẠNG CÂU HỎI
        print(f"✅ Câu {cau_so} (Match {score*100:.1f}%): Đã tìm thấy trên hệ thống.")
        
        if loai_cau == "TN": # Trắc nghiệm (type 0)
            so_luong_ai = len(ai_q.get("cac_dap_an", []))
            so_luong_sys = len(sys_q["options"])
            
            if so_luong_ai != 4 or so_luong_sys != 4:
                print(f"  ⚠️ CẢNH BÁO LỖI CẤU TRÚC FILE GỐC: AI tìm thấy {so_luong_ai} đáp án, Hệ thống có {so_luong_sys} đáp án.")
                print(f"      -> Dạng TN phải có đúng 4 đáp án. Khả năng cao file gốc bị dính liền chữ (VD: gộp C và D) hoặc gõ thiếu. CẦN MỞ FILE ĐỂ SỬA!")
            
            dap_an_ai_nhan = ai_q.get("dap_an_dung", "").strip()
            
            # Lấy ra nội dung text của đáp án đúng theo AI
            noi_dung_dap_an_ai = ""
            for opt in ai_q.get("cac_dap_an", []):
                if opt["nhan"] == dap_an_ai_nhan:
                    noi_dung_dap_an_ai = clean_text(opt["noi_dung"])
                    break
                    
            if not noi_dung_dap_an_ai:
                print(f"  ⚠️ Cảnh báo: AI đánh dấu đáp án đúng là {dap_an_ai_nhan} nhưng không tìm thấy nội dung tương ứng.")
                continue

            # Tìm xem text đáp án AI này nằm ở ID nào trong các lựa chọn của Hệ thống
            sys_options_texts = list(sys_q["options"].values())
            matched_opt_text, opt_score = find_best_match(noi_dung_dap_an_ai, sys_options_texts)
            
            if matched_opt_text:
                # Tìm ID của đáp án trên hệ thống
                sys_opt_id = [k for k, v in sys_q["options"].items() if v == matched_opt_text][0]
                
                # Kiểm tra ID này có nằm trong đáp án đúng của hệ thống không
                if sys_opt_id in sys_q["answerOptionId"]:
                    print("  -> Khớp 100%: Đáp án chuẩn xác!")
                else:
                    print(f"  -> ❌ LỆCH ĐÁP ÁN: ĐỀ GỐC chọn [{noi_dung_dap_an_ai}] nhưng hệ thống chọn đáp án khác.")
            else:
                print(f"  -> ❌ LỖI: Nội dung đáp án [{noi_dung_dap_an_ai}] không tồn tại trong các lựa chọn của hệ thống.")

        elif loai_cau == "DS": # Đúng sai (type 1)
            # 1. Lấy dữ liệu thô và chuyển về chữ thường
            ai_ans_str = str(ai_q.get("dap_an_dung", "")).lower()
            sys_ans_arr = [str(x).lower() for x in sys_q["answerFreeText"]]
            
            # 2. Đổi hệ thống sang chuẩn tiếng Việt: true -> đúng, false -> sai
            sys_ans_mapped = ["đúng" if "true" in a else "sai" for a in sys_ans_arr]
            
            # 3. Dùng Regex nhặt tất cả các chữ "đúng", "sai" từ chuỗi AI
            ai_ans_extracted = re.findall(r'(đúng|sai)', ai_ans_str)
            
            # (Tùy chọn) Chuẩn bị chuỗi hiển thị cho đẹp
            ai_display = ", ".join(ai_ans_extracted) if ai_ans_extracted else ai_ans_str
            sys_display = ", ".join(sys_ans_mapped)
            
            # 4. SO SÁNH
            if ai_ans_extracted == sys_ans_mapped:
                print(f"  -> Khớp 100%: Dãy đáp án [{sys_display}]")
            else:
                print(f"  -> ❌ LỆCH ĐÁP ÁN: ĐỀ GỐC = [{ai_display}] | Hệ thống = [{sys_display}]")
                # In chi tiết từng ý để người QA (kiểm thử) dễ nhìn hơn:
                print("       Chi tiết khác biệt:")
                for i in range(max(len(ai_ans_extracted), len(sys_ans_mapped))):
                    a_val = ai_ans_extracted[i] if i < len(ai_ans_extracted) else "Trống"
                    s_val = sys_ans_mapped[i] if i < len(sys_ans_mapped) else "Trống"
                    if a_val != s_val:
                        print(f"         + Ý {chr(97+i)}: ĐỀ GỐC chọn [{a_val}] NHƯNG Hệ thống chọn [{s_val}]")
            # Soạn một đoạn Regex nhỏ để check số lượng Đúng/Sai có khớp không (Có thể mở rộng thêm)
            
        elif loai_cau == "DIEN": # Điền (type 5)
            # Lấy chuỗi gốc
            ai_dien = str(ai_q.get("dap_an_dung", ""))
            sys_dien = str(sys_q["answerFreeText"][0]) if sys_q["answerFreeText"] else ""
            
            # Dùng clean_text để lột bỏ dấu chấm, khoảng trắng thừa trước khi so sánh
            if clean_text(ai_dien) == clean_text(sys_dien):
                print(f"  -> Khớp 100%: Giá trị điền [{ai_dien}]")
            else:
                print(f"  -> ❌ LỆCH ĐÁP ÁN: ĐỀ GỐC [{ai_dien}] | Hệ thống = [{sys_dien}]")

if __name__ == "__main__":
    # Cả 2 biến đều thêm chữ 'r' ở đầu và dùng 1 dấu '\'
    AI_FILE = r"D:\CheckTool\FetchJson\dapan_tu_pdf.json"
    SYSTEM_FILE = r"D:\CheckTool\FetchJson\onluyen_data\69cc7f41b86b57e0fac47636.json"
    
    compare_exams(AI_FILE, SYSTEM_FILE)