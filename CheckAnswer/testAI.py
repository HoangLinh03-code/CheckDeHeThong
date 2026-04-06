"""
Tool đối chiếu đề gốc (PDF JSON) vs đề hệ thống (System JSON)
Phiên bản v3 - hỗ trợ mọi môn học, mọi cấu trúc đề
  - Xử lý nhom_cau_hoi (flatten câu con ra trước khi so sánh)
  - Match bằng options khi nội dung câu hỏi quá ngắn/không có tiêu đề
  - Fix strip_html không cắt nhầm ký tự < trong LaTeX
  - Fix HTML entity &#39;
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


def check_formula_issues(pdf_content, sys_content, sys_opts):
    """
    Hàm kiểm tra toàn diện các vấn đề về công thức Toán/Hóa.
    """
    issues = []
    pdf_content = pdf_content or ""
    sys_content = sys_content or ""
    
    # Gộp toàn bộ text của hệ thống
    sys_full_text = sys_content
    if isinstance(sys_opts, dict):
        sys_full_text += " " + " ".join(str(v) for v in sys_opts.values())
    elif isinstance(sys_opts, list):
        sys_full_text += " " + " ".join(str(v) for v in sys_opts)

    # ---------------------------------------------------------
    # 1. KIỂM TRA LỖI FONT / DẤU MŨI TÊN
    # ---------------------------------------------------------
    # Đã sửa lỗi danh sách chứa chuỗi rỗng và khoảng trắng
    bad_chars = ['\ufffd', '□', '▯', '\x01']
    found_bad_chars = set([c for c in bad_chars if c in sys_full_text])
    
    if found_bad_chars:
        issues.append("❌ Lỗi hiển thị: Phát hiện ký tự rác/hộp vuông. (Thường do lỗi mũi tên -> phương trình Hóa học).")

    # ---------------------------------------------------------
    # 2. KIỂM TRA LỖI MATHTYPE / EQUATION
    # ---------------------------------------------------------
    math_error_keywords = ["mathtype", "[math error]", "math processing error", "ole_link", "equation.3"]
    for kw in math_error_keywords:
        if kw in sys_full_text.lower():
            issues.append(f"❌ Lỗi convert công thức: Phát hiện dấu hiệu lỗi '{kw}' thay vì hiển thị công thức.")
            
    latex_commands = ["\\frac", "\\sqrt", "\\int", "\\lim", "\\sum"]
    has_latex_cmd = any(cmd in sys_full_text for cmd in latex_commands)
    if has_latex_cmd and "$" not in sys_full_text and "\\(" not in sys_full_text:
         issues.append("⚠️ Lỗi render: Hệ thống chứa mã LaTeX nhưng không được bao bọc bởi thẻ toán học ($...$)")

    # ---------------------------------------------------------
    # 3. KIỂM TRA CÔNG THỨC BỊ MẤT
    # ---------------------------------------------------------
    pdf_formulas = re.findall(r'\$([^\$]+)\$', pdf_content)
    long_formulas = [f for f in pdf_formulas if len(f.replace(" ", "")) >= 4]
    
    if long_formulas:
        missing_formulas = []
        norm_sys_text = normalize(sys_full_text)
        
        for formula in long_formulas:
            norm_formula = normalize(formula)
            if norm_formula not in norm_sys_text:
                missing_formulas.append(formula)

        if missing_formulas:
            preview = ", ".join([f"${f}$" for f in missing_formulas[:2]])
            if len(missing_formulas) > 2:
                preview += " ..."
            issues.append(f"❌ Mất công thức: Đề gốc có nhưng hệ thống bị thiếu nội dung (VD bị mất: {preview}).")

    return issues

def check_image_issues(pdf_co_hinh, sys_content, sys_opts):
    """
    Hàm kiểm tra toàn diện các vấn đề về hình ảnh trong câu hỏi.
    - Nhận diện thiếu hình / dư hình so với file gốc.
    - Nhận diện hình nằm sai chỗ (trong các lựa chọn đáp án).
    - Nhận diện hình không nằm ở vị trí cuối cùng của nội dung câu hỏi.
    """
    issues = []
    
    # Ép kiểu an toàn để tránh lỗi NoneType
    sys_content = sys_content or ""
    
    # 1. Phát hiện thẻ <img trong nội dung và các lựa chọn
    img_in_content = "<img" in sys_content.lower()
    
    img_in_opts = False
    if isinstance(sys_opts, dict):
        img_in_opts = any("<img" in str(opt_text).lower() for opt_text in sys_opts.values())
    elif isinstance(sys_opts, list):
        img_in_opts = any("<img" in str(opt).lower() for opt in sys_opts)
        
    has_any_img = img_in_content or img_in_opts

    # 2. Kiểm tra tính đồng nhất số lượng hình (Có bị mất hoặc dư không)
    if pdf_co_hinh and not has_any_img:
        issues.append("❌ Thiếu hình ảnh (Đề gốc có hình nhưng trên hệ thống không tìm thấy thẻ <img).")
    elif not pdf_co_hinh and has_any_img:
        issues.append("⚠️ Cảnh báo: Hệ thống có hình ảnh nhưng file JSON gốc AI không đánh dấu 'co_hinh' (Nên check lại xem kĩ thuật viên tự chèn hay AI bắt trượt).")

    # 3. Kiểm tra vị trí đặt hình (Chỉ xét nếu trên hệ thống có hình)
    if has_any_img:
        # Lỗi 3.1: Hình lọt vào options
        if img_in_opts:
            issues.append("❌ Sai vị trí hình: Hình ảnh bị chèn nhầm vào các lựa chọn (A, B, C, D) thay vì nằm ở phần nội dung câu hỏi.")
        
        # Lỗi 3.2: Hình không nằm ở cuối cùng của nội dung câu hỏi
        if img_in_content:
            # Tìm vị trí xuất hiện cuối cùng của thẻ "<img"
            last_img_idx = sys_content.lower().rfind("<img")
            
            # Cắt lấy chuỗi từ thẻ <img cuối cùng cho đến hết đoạn HTML
            tail_content = sys_content[last_img_idx:]
            
            # Xóa sạch TẤT CẢ các thẻ HTML trong đoạn đuôi này (ví dụ </img>, </p>, <br>, </span>...)
            text_after_img = re.sub(r'<[^>]+>', '', tail_content)
            
            # Xóa tiếp các khoảng trắng, ký tự đặc biệt vô nghĩa
            text_after_img = text_after_img.replace('&nbsp;', '').replace(' ', '').replace('\n', '').strip()
            
            # Nếu đằng sau hình vẫn còn chữ (độ dài text thực tế > 3 ký tự)
            if len(text_after_img) > 3:
                # Lấy một đoạn text ngắn để báo lỗi cho dễ nhìn
                preview_text = re.sub(r'<[^>]+>', '', tail_content).replace('&nbsp;', ' ').strip()
                preview_text = (preview_text[:20] + '...') if len(preview_text) > 20 else preview_text
                issues.append(f"❌ Sai vị trí hình: Hình ảnh chưa được đặt ở dưới cùng của câu hỏi (Còn sót chữ đằng sau hình: '{preview_text}').")
                
    return issues


def strip_html(text):
    """Xóa thẻ HTML thực sự (tag bắt đầu bằng chữ cái), không cắt < trong LaTeX."""
    text = re.sub(r"</?\s*[a-zA-Z][a-zA-Z0-9]*(?:\s[^>]*)?>", " ", str(text))
    for entity, char in HTML_ENTITIES:
        text = text.replace(entity, char)
    return text.strip()


def normalize(text):
    """Chuẩn hóa để so sánh: strip HTML → unicode math → latex → xóa khoảng trắng → lowercase."""
    text = strip_html(text)
    
    for uni, latex in UNICODE_MATH.items():
        text = text.replace(uni, latex)
        
    text = text.replace("$", "")
    
    # THÊM MỚI: Xóa các định dạng LaTeX hay gây nhiễu (Bắt buộc chạy trước khi xóa ngoặc nhọn)
    noisy_latex = [
        "\\left\\{", "\\right.", "\\left", "\\right", 
        "\\begin{cases}", "\\end{cases}", 
        "\\begin{array}{l}", "\\begin{array}{ll}", "\\begin{array}", "\\end{array}",
        "\\text", "\\\\"  # Cặp gạch chéo kép \\ để xuống dòng trong hệ phương trình
    ]
    for noise in noisy_latex:
        text = text.replace(noise, "")
        
    # Bỏ qua ngoặc nhọn trong LaTeX
    text = text.replace("{", "").replace("}", "") 
    
    for old, new in LATEX_NORM:
        text = text.replace(old, new)
        
    text = re.sub(r"\s+", "", text)
    text = text.strip(".,;:!?")
    return text.lower()


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
    Mỗi câu trả về dict chuẩn: {so, type, content, options, answer, group_context}
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
                "co_hinh": item.get("co_hinh", False)
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
    Chiến lược 2 bước:
      1. Match bằng nội dung câu hỏi (content)
      2. Nếu content ngắn/thiếu → match bằng tập options
    Trả về (sys_q, score, match_method)
    """
    pdf_content = pdf_q["content"]
    pdf_opts_text = [o.get("noi_dung", "") for o in pdf_q["options"]]

    # Bước 1: thử match bằng content
    norm_content = normalize(pdf_content)
    if len(norm_content) >= CONTENT_MIN_LEN:
        idx, score = best_match(pdf_content, sys_contents, threshold=0.65)
        if idx >= 0:
            return sys_qs[idx], score, "content"

    # Bước 2: match bằng options (hữu ích khi câu hỏi chỉ là "Mark the letter...")
    if pdf_opts_text:
        best_score, best_idx = 0.0, -1
        for i, sys_q in enumerate(sys_qs):
            sys_opts_text = list(sys_q["options"].values())
            score = options_similarity(pdf_opts_text, sys_opts_text)
            if score > best_score:
                best_score, best_idx = score, i
        if best_score >= 0.75:
            return sys_qs[best_idx], best_score, "options"

    # Không tìm thấy — vẫn trả về best guess để báo cáo
    idx, score = best_match(pdf_content, sys_contents, threshold=0.0)
    return None, score, "none"


# ─────────────────────────────────────────────────────────────────────
# 4. KIỂM TRA TỪNG DẠNG CÂU
# ─────────────────────────────────────────────────────────────────────

MIN_CONTENT_LEN_OPT = 5  # ký tự sau normalize để coi option hệ thống là "bị cụt"


def check_TN(pdf_q, sys_q, issues):
    pdf_opts = pdf_q["options"]
    sys_opts = sys_q["options"]          # dict {id: raw_content}
    sys_opt_ids   = list(sys_opts.keys())
    sys_opt_raws  = list(sys_opts.values())
    sys_opt_norms = [normalize(v) for v in sys_opt_raws]

    # --- Kiểm tra số lượng đáp án ---
    n_pdf = len(pdf_opts)
    n_sys = len(sys_opts)
    if n_pdf not in (0, 4):
        issues.append(f"  ⚠️  SỐ ĐÁP ÁN ĐỀ GỐC = {n_pdf} (phải là 4) – Kiểm tra file gốc!")
    if n_sys != 4:
        lvl = "❌" if n_sys < 4 else "❌"
        issues.append(f"  {lvl} SỐ ĐÁP ÁN HỆ THỐNG = {n_sys} (phải là 4) – Thiếu/thừa đáp án!")

    # --- Kiểm tra nội dung từng đáp án ---
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

    # --- Kiểm tra đáp án đúng ---
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
        return  # Đã báo ở phần check options ở trên rồi

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

        # Cảnh báo khớp thấp chỉ khi match bằng content
        if method == "content" and score < 0.85:
            print(f"  ⚠️  Nội dung câu hỏi khớp thấp ({score*100:.0f}%) – kiểm tra lại câu hỏi!")

        # Kiểm tra loại câu
        if sys_q["type"] != qtype:
            type_mismatch.append((cau_so, qtype, sys_q["type"]))
            print(f"  ⚠️  LOẠI CÂU LỆCH: Đề gốc={qtype}, Hệ thống={sys_q['type']}")

        issues = []
        
        pdf_co_hinh = pdf_q.get("co_hinh", False)
        sys_content = sys_q.get("content", "")
        sys_opts = sys_q.get("options", {})
        
        # Gọi hàm kiểm tra hình ảnh
        image_issues = check_image_issues(pdf_co_hinh, sys_content, sys_opts)
        if image_issues:
            issues.extend(image_issues)
        
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

    # Dùng đối số dòng lệnh nếu có, không thì dùng giá trị mặc định
    if len(_sys.argv) == 3:
        PDF_FILE = _sys.argv[1]
        SYS_FILE = _sys.argv[2]
    else:
        PDF_FILE = "D:\CheckTool\FetchJson\dapan_Toan_pdf.json"
        SYS_FILE = "D:\CheckTool\FetchJson\onluyen_data\\69cf66813268c911d45bc75e.json"

    compare(PDF_FILE, SYS_FILE)