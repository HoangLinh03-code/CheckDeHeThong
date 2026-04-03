import sys
import os

# 1. Lấy đường dẫn tuyệt đối của thư mục chứa file này (CheckAnswer)
current_dir = os.path.dirname(os.path.abspath(__file__))

# 2. Lấy đường dẫn của thư mục cha (FetchJson)
project_root = os.path.dirname(current_dir)

# 3. Thêm thư mục cha vào đầu danh sách tìm kiếm của Python TRƯỚC KHI import Api
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Bây giờ Python đã có thể tìm thấy thư mục 'Api' ở thư mục gốc
import json
from Api.callAPI import get_vertex_ai_credentials, VertexClient

def extract_answers_with_ai(pdf_path, output_json_path):
    if not os.path.exists(pdf_path):
        print(f"❌ Không tìm thấy file PDF tại: {pdf_path}")
        return

    print("[1] Đang xác thực Vertex AI...")
    creds = get_vertex_ai_credentials()
    if not creds: return
        
    client = VertexClient(
        project_id=os.getenv("PROJECT_ID"), 
        creds=creds, 
        model_name="gemini-3.1-pro-preview" 
    )

    # =========================================================================
    # UNIVERSAL EXAM PROMPT - ÁP DỤNG ĐA MÔN, ĐA ĐỊNH DẠNG
    # =========================================================================
    # =========================================================================
    # UNIVERSAL EXAM PROMPT - CẬP NHẬT LẤY TOÀN BỘ CÂU HỎI VÀ ĐÁP ÁN ĐỐI CHIẾU
    # =========================================================================
    prompt = """
    Bạn là một AI chuyên gia bóc tách dữ liệu giáo dục. Nhiệm vụ của bạn là quét toàn bộ file tài liệu (đề thi) được cung cấp và trích xuất CHI TIẾT TỪNG CÂU HỎI.
    Hệ thống cần dữ liệu thô này để đối chiếu với một file JSON chuẩn nhằm tìm ra lỗi sai sót.

    *** CÁC THÔNG TIN CẦN TRÍCH XUẤT CHO MỖI CÂU: ***
    1. Nội dung câu hỏi: Trích xuất trọn vẹn phần chữ của câu hỏi. Nếu có công thức Toán/Lý/Hóa, hãy giữ nguyên định dạng LaTeX (VD: $x^2 + 1$).
    2. Hình ảnh: Xác định xem trong nội dung câu hỏi hoặc đáp án có hình ảnh, biểu đồ hay đồ thị không. Trả về true/false.
    3. Các đáp án (Lựa chọn): Lấy nội dung của tất cả các đáp án được cho (A, B, C, D hoặc Đúng/Sai).
    4. Đáp án đúng: Dựa vào các dấu hiệu (in đậm, bôi đỏ, gạch chân, khoanh tròn, hoặc bảng đáp án cuối đề) để tìm ra đâu là đáp án đúng.
    5. Lời giải: Nếu có lời giải chi tiết đi kèm, hãy trích xuất luôn.

    *** CÁC LOẠI CÂU HỎI VÀ CÁCH XỬ LÝ: ***
    1. Trắc nghiệm 1 lựa chọn (TN): Câu hỏi có 4 đáp án A, B, C, D. Lấy đủ cả 4 đáp án.
    2. Đúng / Sai (DS): Câu hỏi gồm 4 ý a, b, c, d. Xử lý mỗi ý như một đáp án riêng biệt và ghi chú ý nào là Đúng (True), ý nào Sai (False).
    3. Điền khuyết / Tự luận ngắn (DIEN): Không có lựa chọn, chỉ có nội dung câu hỏi và đáp án cuối cùng.

    *** YÊU CẦU ĐẦU RA (BẮT BUỘC): ***
    - Trả về DUY NHẤT một mảng JSON (JSON Array). Không giải thích, không thêm chữ nào khác.
    - Không thêm block code markdown như ```json.
    - QUÉT TUẦN TỰ từ câu đầu tiên đến câu cuối cùng, tuyệt đối không bỏ sót.

    *** CẤU TRÚC JSON MẪU: ***
    [
      {
        "loai_cau": "TN",
        "cau_so": 1,
        "noi_dung_cau_hoi": "Cho tam giác ABC có a=5, b=6, c=7. Diện tích tam giác ABC bằng",
        "co_hinh_anh": false,
        "cac_dap_an": [
          {"nhan": "A", "noi_dung": "2\\sqrt{6}"},
          {"nhan": "B", "noi_dung": "3\\sqrt{2}"},
          {"nhan": "C", "noi_dung": "\\sqrt{6}"},
          {"nhan": "D", "noi_dung": "\\sqrt{216}"}
        ],
        "dap_an_dung": "D",
      },
      {
        "loai_cau": "DS",
        "cau_so": 2,
        "noi_dung_cau_hoi": "Cho hàm số y = f(x) có đồ thị như hình vẽ. Các mệnh đề sau đúng hay sai?",
        "co_hinh_anh": true,
        "cac_dap_an": [
          {"nhan": "a", "noi_dung": "Hàm số đồng biến trên khoảng (0; 2)."},
          {"nhan": "b", "noi_dung": "Hàm số đạt cực đại tại x = 1."},
          {"nhan": "c", "noi_dung": "Giá trị lớn nhất của hàm số là 3."},
          {"nhan": "d", "noi_dung": "Đồ thị hàm số cắt trục tung tại điểm có tung độ bằng 1."}
        ],
        "dap_an_dung": "a-Đúng, b-Sai, c-Đúng, d-Sai",
      },
      {
        "loai_cau": "DIEN",
        "cau_so": 3,
        "noi_dung_cau_hoi": "Một người ở vị trí có toạ độ (-4;4) di chuyển...",
        "co_hinh_anh": true,
        "cac_dap_an": [],
        "dap_an_dung": "0,61",
      }
    ]
    """

    print(f"[2] Đang gửi file {pdf_path} cho AI phân tích. Vui lòng đợi...")
    
    try:
        response_text = client.send_data_to_AI(
            prompt=prompt,
            file_paths=[pdf_path],
            temperature=0.0 # Bắt buộc giữ 0.0 để AI tập trung trích xuất, không bay bổng
        )
        
        if response_text and not response_text.startswith("❌"):
            clean_json = response_text.replace("```json", "").replace("```", "").strip()
            parsed_data = json.loads(clean_json)
            
            with open(output_json_path, "w", encoding="utf-8") as f:
                json.dump(parsed_data, f, ensure_ascii=False, indent=2)
                
            print(f"✅ HOÀN THÀNH! Đã trích xuất {len(parsed_data)} bản ghi. Đã lưu vào: {output_json_path}")
            
            # Print sample để kiểm tra nhanh 3 bản ghi đầu và cuối
            print("\n--- MẪU KẾT QUẢ (3 bản ghi đầu) ---")
            print(json.dumps(parsed_data[:3], indent=2, ensure_ascii=False))
            print("...")
            
        else:
            print("❌ AI không trả về kết quả hợp lệ.")

    except json.JSONDecodeError:
        print("❌ Lỗi: AI không trả về chuẩn định dạng JSON. Kết quả thô:")
        print(response_text)
    except Exception as e:
        print(f"❌ Có lỗi xảy ra trong quá trình xử lý: {e}")

if __name__ == "__main__":
    # Test thử với file của bạn hoặc bất kỳ file môn nào khác
    pdf_file = "Toán_ĐỀ_ONLINE_10_mới_2.pdf" 
    out_file = "dapan_tu_pdf.json"
    extract_answers_with_ai(pdf_file, out_file)