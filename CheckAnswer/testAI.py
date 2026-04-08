"""
Tool đối chiếu đề gốc (PDF JSON) vs đề hệ thống (System JSON)
Phiên bản v4 - tối ưu hóa Hóa học, Media (Hình ảnh/Bảng), và công thức Toán học
"""

import json
import re
from difflib import SequenceMatcher

# ─────────────────────────────────────────────────────────────────────
# 1. UTILS
# ─────────────────────────────────────────────────────────────────────

HTML_ENTITIES = [
    ("&nbsp;", " "), ("&lt;", "<"), ("&gt;", ">"),
    ("&amp;", "&"), ("&#39;", "'"), ("&quot;", '"'), ("\xa0", " "),
]

UNICODE_MATH = {
    "ℝ": "\\mathbb{R}", "ℕ": "\\mathbb{N}", "ℤ": "\\mathbb{Z}", "ℚ": "\\mathbb{Q}",
    "∖": "\\setminus", "∈": "\\in", "∉": "\\notin",
    "∪": "\\cup", "∩": "\\cap", "∅": "\\emptyset",
    "≥": "\\ge", "≤": "\\le", "→": "\\rightarrow", "∞": "\\infty",
}

LATEX_NORM = [
    ("\\left", ""), ("\\right", ""),
    ("\\geqslant", "\\ge"), ("\\leqslant", "\\le"),
    ("\\geq", "\\ge"), ("\\leq", "\\le"),
]


def strip_html(text):
    """Xóa thẻ HTML thực sự (tag bắt đầu bằng chữ cái), không cắt < trong LaTeX."""
    text = re.sub(r"</?\s*[a-zA-Z][a-zA-Z0-9]*(?:\s[^>]*)?>", " ", str(text))
    for entity, char in HTML_ENTITIES:
        text = text.replace(entity, char)
    return text.strip()


def normalize(text):
    """Chuẩn hóa để so sánh: chuyển thẻ sub/sup -> xóa HTML -> unicode math -> latex -> xóa khoảng trắng -> lowercase."""
    text = str(text)
    
    # 1. Chuẩn hóa các loại dấu gạch ngang (en-dash, em-dash, minus) về 1 dấu trừ chuẩn
    text = re.sub(r'[–—−]', '-', text)
    
    # 2. Bảo toàn chỉ số trên/dưới của Hóa học trước khi xóa HTML
    text = re.sub(r'(?i)<sub>(.*?)</sub>', r'_\1', text)
    text = re.sub(r'(?i)<sup>(.*?)</sup>', r'^\1', text)
    
    text = strip_html(text)
    
    # 3. Đồng nhất kí tự Delta Unicode với LaTeX
    text = re.sub(r'[∆Δ]', r'\\Delta', text)
    
    for uni, latex in UNICODE_MATH.items():
        text = text.replace(uni, latex)
        
    text = text.replace("$", "")
    
    # 4. Đồng nhất các loại mũi tên (LaTeX và Unicode) về chuẩn \rightarrow
    text = re.sub(r'\\xrightarrow(\[.*?\])?\{.*?\}', r'\\rightarrow', text)
    text = re.sub(r'\\xrightarrow', r'\\rightarrow', text)
    text = re.sub(r'[⟶→⇌]', r'\\rightarrow', text)
    
    noisy_latex = [
        "\\left\\{", "\\right.", "\\left", "\\right", 
        "\\begin{cases}", "\\end{cases}", 
        "\\begin{array}{l}", "\\begin{array}{ll}", "\\begin{array}", "\\end{array}",
        "\\text", "\\\\"
    ]
    for noise in noisy_latex:
        text = text.replace(noise, "")
        
    text = text.replace("{", "").replace("}", "") 
    
    # 5. Xử lý ký tự bảng biểu dạng text (Markdown table)
    text = text.replace("|", "")
    text = re.sub(r'-{2,}', '', text)
    
    for old, new in LATEX_NORM:
        text = text.replace(old, new)
        
    # Xóa sạch khoảng trắng, \n, \t...
    text = re.sub(r"\s+", "", text)
    text = text.strip(".,;:!?")
    return text.lower()


def simplify_math_text(text):
    """Hàm tối giản tuyệt đối: Cạo sạch mọi định dạng LaTeX/HTML, chỉ giữ lại chữ, số và dấu = để so sánh độ tồn tại."""
    t = str(text).lower()
    t = strip_html(t)
    
    # Đồng nhất các loại mũi tên thành dấu =
    t = re.sub(r'\\xrightarrow(\[.*?\])?\{.*?\}', '=', t)
    t = re.sub(r'\\[a-z]*arrow', '=', t)
    t = re.sub(r'[⟶→⇌]', '=', t)
    
    # Đồng nhất kí hiệu đặc biệt
    t = re.sub(r'\\delta|∆|Δ', 'd', t)
    
    # Bắt chữ độ C (circ) của LaTeX và chuyển thành số 0
    t = re.sub(r'\\circ', '0', t)
    t = t.replace('°', '0')
    t = t.replace('℃', '0c') # FIX: Xử lý ký tự Unicode độ C gộp của hệ thống
    t = t.replace('℉', '0f') # (Dự phòng cho độ F)
    
    # Xóa các lệnh LaTeX (bắt đầu bằng \)
    t = re.sub(r'\\[a-z]+', '', t)
    
    # Xóa TẤT CẢ kí tự không phải chữ, số, hoặc dấu =
    t = re.sub(r'[^a-z0-9=]', '', t)
    
    # Đồng nhất hoàn toàn chữ 'o' và số '0' (Giúp $H^o$ và $H^0$ khớp nhau 100%)
    t = t.replace('o', '0')
    
    return t


def check_formula_issues(pdf_content, sys_content, sys_opts):
    """Hàm kiểm tra các vấn đề công thức Toán/Hóa (Đã khử false positive và fix lỗi window size)."""
    issues = []
    pdf_content = pdf_content or ""
    sys_content = sys_content or ""
    
    sys_full_text = sys_content
    if isinstance(sys_opts, dict):
        sys_full_text += " " + " ".join(str(v) for v in sys_opts.values())
    elif isinstance(sys_opts, list):
        sys_full_text += " " + " ".join(str(v) for v in sys_opts)

    bad_chars = ['\ufffd', '□', '▯', '\x01']
    if set([c for c in bad_chars if c in sys_full_text]):
        issues.append("❌ Lỗi hiển thị: Phát hiện ký tự rác/hộp vuông.")

    math_error_keywords = ["mathtype", "[math error]", "math processing error", "ole_link", "equation.3"]
    for kw in math_error_keywords:
        if kw in sys_full_text.lower():
            issues.append(f"❌ Lỗi convert công thức: Phát hiện dấu hiệu lỗi '{kw}'.")

    pdf_formulas = re.findall(r'\$([^\$]+)\$', pdf_content)
    # Chỉ xét công thức đủ dài (bỏ qua các biến lẻ như $x$, $A$)
    long_formulas = [f for f in pdf_formulas if len(re.sub(r'[^a-zA-Z0-9]', '', f)) >= 4]
    
    if long_formulas:
        missing_formulas = []
        # Chuyển hệ thống về dạng tối giản nhất để tìm
        sys_simple = simplify_math_text(sys_full_text)
        
        for formula in long_formulas:
            form_simple = simplify_math_text(formula)
            if not form_simple or len(form_simple) < 3: 
                continue
                
            if form_simple in sys_simple:
                continue
                
            # Fuzzy match cho chuỗi đã tối giản
            found = False
            L = len(form_simple)
            threshold = 0.85 if L > 10 else 0.9
            
            if len(sys_simple) < L:
                if SequenceMatcher(None, form_simple, sys_simple).ratio() >= threshold:
                    found = True
            else:
                # FIX: Duyệt các cửa sổ với độ dài linh hoạt từ L đến L+3 để không bị rớt điểm ratio
                for i in range(len(sys_simple) - L + 1):
                    for delta in range(0, 4): 
                        if i + L + delta <= len(sys_simple):
                            window = sys_simple[i:i+L+delta]
                            if SequenceMatcher(None, form_simple, window).ratio() >= threshold:
                                found = True
                                break
                    if found:
                        break
                        
            if not found:
                missing_formulas.append(formula)

        if missing_formulas:
            preview = ", ".join([f"${f}$" for f in missing_formulas[:2]])
            if len(missing_formulas) > 2:
                preview += " ..."
            issues.append(f"❌ Mất công thức: Đề gốc có nhưng hệ thống thiếu hoặc sai lệch (VD: {preview}).")

    return issues


def check_image_issues(pdf_co_hinh, pdf_co_bang, sys_content, sys_opts, pdf_content=""):
    """Hàm kiểm tra các vấn đề về hình ảnh và bảng biểu (Kiểm tra vị trí tuyệt đối qua tag [HÌNH_ẢNH])."""
    issues = []
    sys_content = sys_content or ""
    pdf_content = pdf_content or ""
    
    def has_image(t):
        t = str(t).lower()
        return "<img" in t or "\\includegraphics" in t
        
    def has_table(t):
        return "<table" in str(t).lower()
        
    img_in_content = has_image(sys_content)
    
    img_in_opts = False
    if isinstance(sys_opts, dict):
        img_in_opts = any(has_image(v) for v in sys_opts.values())
    elif isinstance(sys_opts, list):
        img_in_opts = any(has_image(v) for v in sys_opts)
        
    has_any_img = img_in_content or img_in_opts
    has_sys_table = has_table(sys_content)

    # --- 1. Kiểm tra Hình ảnh ---
    if pdf_co_hinh and not has_any_img:
        if has_sys_table:
            pass
        else:
            issues.append("❌ Thiếu hình ảnh (Đề gốc có hình nhưng trên hệ thống không tìm thấy thẻ <img).")
    elif not pdf_co_hinh and has_any_img:
        issues.append("⚠️ Cảnh báo: Hệ thống có hình ảnh nhưng file JSON gốc AI không đánh dấu 'co_hinh'.")

    # --- Kiểm tra vị trí hình ảnh ---
    if has_any_img:
        # Lỗi: Hình chèn nhầm vào lựa chọn
        if img_in_opts:
            issues.append("❌ Sai vị trí hình: Hình ảnh bị chèn nhầm vào các lựa chọn (A, B, C, D).")
            
        # Lỗi: Đối chiếu vị trí bằng tag [HÌNH_ẢNH]
        if img_in_content and "[HÌNH_ẢNH]" in pdf_content:
            idx_img = sys_content.lower().find("<img")
            idx_inc = sys_content.lower().find("\\includegraphics")
            first_sys_img_idx = min(idx_img, idx_inc) if (idx_img != -1 and idx_inc != -1) else max(idx_img, idx_inc)
            
            # Text trước hình ảnh trên hệ thống
            sys_before = strip_html(sys_content[:first_sys_img_idx])
            sys_before_clean = re.sub(r'[\s\.,;:!?\-\(\)\[\]{}"\']', '', sys_before)
            
            # Text trước hình ảnh trên file JSON AI
            first_pdf_img_idx = pdf_content.find("[HÌNH_ẢNH]")
            pdf_before = pdf_content[:first_pdf_img_idx]
            pdf_before_clean = strip_html(pdf_before)
            pdf_before_clean = re.sub(r'[\s\.,;:!?\-\(\)\[\]{}"\']', '', pdf_before_clean)
            
            # So sánh vị trí tương đối (Có nằm ở đầu câu hay không)
            is_sys_top = len(sys_before_clean) < 3
            is_pdf_top = len(pdf_before_clean) < 3
            
            if is_sys_top != is_pdf_top:
                if is_pdf_top:
                    issues.append("❌ Sai vị trí hình: Đề gốc hình nằm TRƯỚC chữ, nhưng hệ thống lại để hình SAU chữ.")
                else:
                    issues.append("❌ Sai vị trí hình: Đề gốc hình nằm SAU chữ, nhưng hệ thống lại để hình TRÊN CÙNG.")
    
    # --- 2. Kiểm tra Bảng biểu ---
    if pdf_co_bang and not has_sys_table:
        issues.append("❌ Thiếu bảng biểu (Đề gốc có bảng nhưng trên hệ thống không tìm thấy thẻ <table).")
    elif not pdf_co_bang and has_sys_table:
        if not pdf_co_hinh:
            issues.append("⚠️ Cảnh báo: Hệ thống có bảng nhưng file JSON gốc AI không đánh dấu 'co_bang'.")
            
    # Kiểm tra vị trí bảng biểu qua tag [BẢNG_BIỂU]
    if has_sys_table and "[BẢNG_BIỂU]" in pdf_content:
        idx_tbl = sys_content.lower().find("<table")
        
        # Text trước bảng biểu trên hệ thống
        sys_before_tbl = strip_html(sys_content[:idx_tbl])
        sys_before_tbl_clean = re.sub(r'[\s\.,;:!?\-\(\)\[\]{}"\']', '', sys_before_tbl)
        
        # Text trước bảng biểu trên file JSON AI
        first_pdf_tbl_idx = pdf_content.find("[BẢNG_BIỂU]")
        pdf_before_tbl = pdf_content[:first_pdf_tbl_idx]
        pdf_before_tbl_clean = strip_html(pdf_before_tbl)
        pdf_before_tbl_clean = re.sub(r'[\s\.,;:!?\-\(\)\[\]{}"\']', '', pdf_before_tbl_clean)
        
        # So sánh vị trí tương đối
        is_sys_top_tbl = len(sys_before_tbl_clean) < 3
        is_pdf_top_tbl = len(pdf_before_tbl_clean) < 3
        
        if is_sys_top_tbl != is_pdf_top_tbl:
            if is_pdf_top_tbl:
                issues.append("❌ Sai vị trí bảng: Đề gốc bảng nằm TRƯỚC chữ, nhưng hệ thống lại để bảng SAU chữ.")
            else:
                issues.append("❌ Sai vị trí bảng: Đề gốc bảng nằm SAU chữ, nhưng hệ thống lại để bảng TRÊN CÙNG.")
                
    return issues


def similarity(a, b):
    return SequenceMatcher(None, normalize(a), normalize(b)).ratio()


def best_match(target, candidates, threshold=0.8):
    """Trả về (index, score) của candidate khớp nhất. index=-1 nếu dưới threshold."""
    best_score, best_idx = 0.0, -1
    for i, c in enumerate(candidates):
        s = similarity(target, c)
        if s > best_score:
            best_score, best_idx = s, i
    if best_score >= threshold:
        return best_idx, best_score
    return -1, best_score


def options_similarity(pdf_opts_text, sys_opts_text):
    """
    Tính điểm tương đồng tổng hợp giữa 2 tập đáp án.
    Mỗi option PDF tìm option SYS khớp nhất, lấy trung bình.
    """
    if not pdf_opts_text or not sys_opts_text:
        return 0.0
    scores = []
    for p in pdf_opts_text:
        best = max(similarity(p, s) for s in sys_opts_text)
        scores.append(best)
    return sum(scores) / len(scores)


# ─────────────────────────────────────────────────────────────────────
# 2. LOAD DỮ LIỆU
# ─────────────────────────────────────────────────────────────────────

def flatten_pdf_questions(path):
    """
    Load file JSON đề gốc PDF, flatten cả câu đơn và câu con trong nhóm.
    Mỗi câu trả về dict chuẩn: {so, type, content, options, answer, group_context, co_hinh, co_bang}
    """
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    questions = []
    for item in raw:
        loai_block = item.get("loai_block", "").strip()
        loai_cau = item.get("loai_cau_hoi", "").strip().lower()
        qtype = _map_type_pdf(loai_cau)

        if loai_block == "nhom_cau_hoi":
            # Flatten câu con ra, gắn context chung vào mỗi câu
            group_context = item.get("du_kien_chung", "")
            group_co_hinh = item.get("co_hinh", False)
            group_co_bang = item.get("co_bang", False)
            for child in item.get("danh_sach_cau_hoi_con", []):
                child_loai = child.get("loai_cau_hoi", loai_cau).strip().lower()
                questions.append({
                    "so": child.get("cau_so", "?"),
                    "type": _map_type_pdf(child_loai),
                    "content": child.get("noi_dung_cau_hoi", ""),
                    "options": child.get("cac_lua_chon", []),
                    "answer": child.get("dap_an_dung"),
                    "group_context": group_context,
                    "is_child": True,
                    "co_hinh": child.get("co_hinh", False) or group_co_hinh,
                    "co_bang": child.get("co_bang", False) or group_co_bang,
                })
        else:
            # Câu đơn bình thường
            questions.append({
                "so": item.get("cau_so", "?"),
                "type": qtype,
                "content": item.get("noi_dung_cau_hoi", ""),
                "options": item.get("cac_lua_chon", []),
                "answer": item.get("dap_an_dung"),
                "group_context": "",
                "is_child": False,
                "co_hinh": item.get("co_hinh", False),
                "co_bang": item.get("co_bang", False)
            })
    return questions


def _map_type_pdf(loai_cau):
    mapping = {
        "trac_nghiem": "TN",
        "dung_sai": "DS",
        "tra_loi_ngan": "DIEN",
    }
    return mapping.get(loai_cau, loai_cau.upper())


def flatten_sys_questions(path):
    """
    Load file JSON hệ thống, flatten cả typeData=0 (đơn) và typeData=1 (nhóm).
    Mỗi câu trả về dict chuẩn.
    """
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f).get("data", [])

    questions = []
    for block in data:
        td = block.get("typeData", 0)
        if td == 0:
            ds = block.get("dataStandard", {})
            questions.append(_parse_sys_item(ds))
        elif td == 1:
            dm = block.get("dataMaterial", {})
            for sub in dm.get("data", []):
                questions.append(_parse_sys_item(sub))
    return questions


def _parse_sys_item(item):
    ta = item.get("typeAnswer", 0)
    type_map = {0: "TN", 1: "DS", 5: "DIEN"}
    opts = {opt["idOption"]: opt.get("content", "") for opt in item.get("options", [])}
    return {
        "content": item.get("content", ""),
        "type": type_map.get(ta, f"TYPE{ta}"),
        "options": opts,
        "answerOptionId": item.get("answerOptionId", []),
        "answerFreeText": item.get("answerFreeText", []),
    }


# ─────────────────────────────────────────────────────────────────────
# 3. MATCHING THÔNG MINH
# ─────────────────────────────────────────────────────────────────────

# Ngưỡng tối thiểu để coi nội dung câu hỏi là "có thông tin" (sau normalize)
CONTENT_MIN_LEN = 8


def find_matching_sys_q(pdf_q, sys_qs, sys_contents):
    """
    Tìm câu hệ thống tương ứng với câu PDF.
    """
    pdf_content = pdf_q["content"]
    pdf_opts_text = [o.get("noi_dung", "") for o in pdf_q["options"]]

    norm_content = normalize(pdf_content)
    if len(norm_content) >= CONTENT_MIN_LEN:
        idx, score = best_match(pdf_content, sys_contents, threshold=0.65)
        if idx >= 0:
            return sys_qs[idx], score, "content"

    if pdf_opts_text:
        best_score, best_idx = 0.0, -1
        for i, sys_q in enumerate(sys_qs):
            sys_opts_text = list(sys_q["options"].values())
            score = options_similarity(pdf_opts_text, sys_opts_text)
            if score > best_score:
                best_score, best_idx = score, i
        if best_score >= 0.75:
            return sys_qs[best_idx], best_score, "options"

    idx, score = best_match(pdf_content, sys_contents, threshold=0.0)
    return None, score, "none"


# ─────────────────────────────────────────────────────────────────────
# 4. KIỂM TRA TỪNG DẠNG CÂU
# ─────────────────────────────────────────────────────────────────────

MIN_CONTENT_LEN_OPT = 5  


def check_TN(pdf_q, sys_q, issues):
    pdf_opts = pdf_q["options"]
    sys_opts = sys_q["options"]         
    sys_opt_ids   = list(sys_opts.keys())
    sys_opt_raws  = list(sys_opts.values())
    sys_opt_norms = [normalize(v) for v in sys_opt_raws]

    n_pdf = len(pdf_opts)
    n_sys = len(sys_opts)
    if n_pdf not in (0, 4):
        issues.append(f"  ⚠️  SỐ ĐÁP ÁN ĐỀ GỐC = {n_pdf} (phải là 4) – Kiểm tra file gốc!")
    if n_sys != 4:
        lvl = "❌" if n_sys < 4 else "❌"
        issues.append(f"  {lvl} SỐ ĐÁP ÁN HỆ THỐNG = {n_sys} (phải là 4) – Thiếu/thừa đáp án!")

    for pdf_opt in pdf_opts:
        nhan     = pdf_opt.get("nhan", "")
        nd_pdf   = pdf_opt.get("noi_dung", "")
        norm_pdf = normalize(nd_pdf)

        idx, score = best_match(nd_pdf, sys_opt_raws, threshold=0.0)
        norm_sys = sys_opt_norms[idx] if idx >= 0 else ""

        if score < 0.72:
            len_pdf = len(norm_pdf)
            len_sys = len(norm_sys)
            if len_sys < MIN_CONTENT_LEN_OPT or (len_pdf > 8 and len_sys < len_pdf * 0.4):
                issues.append(
                    f"  ❌ ĐÁP ÁN [{nhan}]: Hệ thống lưu bị CỤT/THIẾU NỘI DUNG!\n"
                    f"       Đề gốc:    «{nd_pdf.strip()}»\n"
                    f"       Hệ thống:  «{strip_html(sys_opt_raws[idx]).strip()}»"
                )
            else:
                issues.append(
                    f"  ❌ ĐÁP ÁN [{nhan}] «{nd_pdf.strip()}» KHÔNG KHỚP lựa chọn nào! (max={score*100:.0f}%)\n"
                    f"       Gần nhất:  «{strip_html(sys_opt_raws[idx]).strip()}»"
                )

    dung_nhan = str(pdf_q.get("answer", "")).strip()
    if not dung_nhan or dung_nhan == "None":
        return

    noi_dung_dung = ""
    for opt in pdf_opts:
        if opt.get("nhan", "").upper() == dung_nhan.upper():
            noi_dung_dung = opt.get("noi_dung", "")
            break

    if not noi_dung_dung:
        issues.append(f"  ⚠️  Không tìm thấy nội dung đáp án đúng [{dung_nhan}] trong file gốc!")
        return

    idx, score = best_match(noi_dung_dung, sys_opt_raws, threshold=0.72)
    if idx == -1:
        return 

    sys_id = sys_opt_ids[idx]
    if sys_id not in sys_q["answerOptionId"]:
        sys_correct_texts = [sys_opts[i] for i in sys_q["answerOptionId"]]
        sys_correct_display = " / ".join(strip_html(t).strip() for t in sys_correct_texts)
        issues.append(
            f"  ❌ LỆCH ĐÁP ÁN ĐÚNG:\n"
            f"       Đề gốc [{dung_nhan}]: «{nd_pdf.strip()}»\n"
            f"       Hệ thống chọn:         «{sys_correct_display}»"
        )


def check_DS(pdf_q, sys_q, issues):
    raw_ans = str(pdf_q.get("answer", "")).lower()
    pdf_answers = re.findall(r"(đúng|sai)", raw_ans)

    sys_free = [str(x).lower() for x in sys_q["answerFreeText"]]
    sys_answers = ["đúng" if "true" in x else "sai" for x in sys_free]

    n_pdf, n_sys = len(pdf_answers), len(sys_answers)
    if n_pdf != n_sys:
        issues.append(f"  ⚠️  SỐ Ý: Đề gốc {n_pdf} ý, hệ thống {n_sys} ý!")
    if n_pdf != 4:
        issues.append(f"  ⚠️  Đề gốc parse được {n_pdf} ý đúng/sai (phải là 4)!")
    if n_sys != 4:
        issues.append(f"  ❌ Hệ thống có {n_sys} ý đúng/sai (phải là 4)!")

    diffs = []
    for i in range(max(n_pdf, n_sys)):
        a = pdf_answers[i] if i < n_pdf else "?"
        s = sys_answers[i] if i < n_sys else "?"
        if a != s:
            diffs.append(f"       Ý {chr(97+i)}: Đề gốc=[{a}]  Hệ thống=[{s}]  ← SAI")

    if diffs:
        issues.append(
            f"  ❌ LỆCH ĐÁP ÁN ĐÚNG/SAI:\n"
            f"       Đề gốc:   [{', '.join(pdf_answers)}]\n"
            f"       Hệ thống: [{', '.join(sys_answers)}]\n" +
            "\n".join(diffs)
        )


def check_DIEN(pdf_q, sys_q, issues):
    pdf_ans = str(pdf_q.get("answer", "")).strip()
    sys_ans = str(sys_q["answerFreeText"][0]).strip() if sys_q["answerFreeText"] else ""

    if normalize(pdf_ans) != normalize(sys_ans):
        issues.append(
            f"  ❌ LỆCH ĐÁP ÁN ĐIỀN:\n"
            f"       Đề gốc:   «{pdf_ans}»\n"
            f"       Hệ thống: «{sys_ans}»"
        )
    elif re.search(r"[.,]$", sys_ans):
        issues.append(f"  ⚠️  Đáp án hệ thống có dấu thừa ở cuối: «{sys_ans}» – nên xóa!")


# ─────────────────────────────────────────────────────────────────────
# 5. MAIN COMPARE
# ─────────────────────────────────────────────────────────────────────

def compare(pdf_path, sys_path):
    pdf_qs = flatten_pdf_questions(pdf_path)
    sys_qs = flatten_sys_questions(sys_path)
    sys_contents = [q["content"] for q in sys_qs]

    print("=" * 70)
    print(f"  ĐỐI CHIẾU ĐỀ – {len(pdf_qs)} câu gốc vs {len(sys_qs)} câu hệ thống")
    print("=" * 70)

    total_ok = total_warn = total_err = 0
    not_found = []
    type_mismatch = []

    for pdf_q in pdf_qs:
        cau_so = pdf_q["so"]
        qtype  = pdf_q["type"]

        print(f"\n{'─'*60}")

        sys_q, score, method = find_matching_sys_q(pdf_q, sys_qs, sys_contents)

        if sys_q is None:
            print(f"❌ CÂU {cau_so} [{qtype}]: KHÔNG TÌM THẤY trên hệ thống! (max match={score*100:.0f}%)")
            not_found.append(cau_so)
            total_err += 1
            continue

        method_tag = "" if method == "content" else f" (match qua {method})"
        print(f"✅ CÂU {cau_so} (match {score*100:.0f}%{method_tag}): TÌM THẤY – Loại [{qtype}]")

        if method == "content" and score < 0.85:
            print(f"  ⚠️  Nội dung câu hỏi khớp thấp ({score*100:.0f}%) – kiểm tra lại câu hỏi!")

        if sys_q["type"] != qtype:
            type_mismatch.append((cau_so, qtype, sys_q["type"]))
            print(f"  ⚠️  LOẠI CÂU LỆCH: Đề gốc={qtype}, Hệ thống={sys_q['type']}")

        issues = []
        
        pdf_co_hinh = pdf_q.get("co_hinh", False)
        pdf_co_bang = pdf_q.get("co_bang", False)
        sys_content = sys_q.get("content", "")
        sys_opts = sys_q.get("options", {})
        
        # Gọi hàm kiểm tra media và truyền thêm nội dung gốc để lấy tag [HÌNH_ẢNH]
        media_issues = check_image_issues(pdf_co_hinh, pdf_co_bang, sys_content, sys_opts, pdf_q.get("content", ""))
        if media_issues:
            issues.extend(media_issues)
        
        formula_issues = check_formula_issues(pdf_q.get("content", ""), sys_content, sys_opts)
        if formula_issues:
            issues.extend(formula_issues)
        
        if qtype == "TN":
            check_TN(pdf_q, sys_q, issues)
        elif qtype == "DS":
            check_DS(pdf_q, sys_q, issues)
        elif qtype == "DIEN":
            check_DIEN(pdf_q, sys_q, issues)
        else:
            issues.append(f"  ⚠️  Loại câu chưa hỗ trợ: {qtype}")

        if issues:
            for iss in issues:
                print(iss)
            errs  = sum(1 for i in issues if "❌" in i)
            warns = sum(1 for i in issues if "⚠️" in i and "❌" not in i)
            total_err  += errs
            total_warn += warns
            if errs == 0:
                total_ok += 1
        else:
            print("  ✔  Đáp án khớp 100%!")
            total_ok += 1

    # ─── TỔNG KẾT ───
    print(f"\n{'='*70}")
    print("  TỔNG KẾT")
    print(f"{'='*70}")
    print(f"  Tổng câu đề gốc   : {len(pdf_qs)}")
    print(f"  Tổng câu hệ thống : {len(sys_qs)}")
    print(f"  ✔  Khớp hoàn toàn : {total_ok}")
    print(f"  ⚠️  Có cảnh báo    : {total_warn}")
    print(f"  ❌ Có lỗi          : {total_err}")
    if not_found:
        print(f"\n  Câu KHÔNG TÌM THẤY: {not_found}")
    if type_mismatch:
        print(f"\n  Câu LỆCH LOẠI:")
        for cau, t_pdf, t_sys in type_mismatch:
            print(f"    Câu {cau}: Đề gốc={t_pdf}, Hệ thống={t_sys}")
    print(f"{'='*70}\n")


# ─────────────────────────────────────────────────────────────────────
# 6. ENTRY POINT
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys as _sys

    if len(_sys.argv) == 3:
        PDF_FILE = _sys.argv[1]
        SYS_FILE = _sys.argv[2]
    else:
        PDF_FILE = "D:\CheckTool\FetchJson\Toán ĐỀ ONLINE 10 mới 2_ai.json"
        SYS_FILE = "D:\CheckTool\FetchJson\onluyen_data\\69cc7f41b86b57e0fac47636.json"

    compare(PDF_FILE, SYS_FILE)