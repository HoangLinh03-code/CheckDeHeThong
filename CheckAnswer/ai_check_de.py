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
from Api.schema import schema_da_nang

def load_prompt_from_file(txt_path):
    with open(txt_path, 'r', encoding='utf-8') as f:
        return f.read()

def process_exam_universal(pdf_path, output_json_path):
    # Lấy đường dẫn thư mục hiện tại
    current_dir = os.path.dirname(os.path.abspath(__file__))
    prompt_path = os.path.join(current_dir, "prompt.txt")
    
    # Đọc nội dung file text
    prompt_content = load_prompt_from_file(prompt_path)
    
    # Khởi tạo Client (giữ nguyên logic của bạn)
    creds = get_vertex_ai_credentials()
    client = VertexClient(
        project_id=os.getenv("PROJECT_ID"), 
        creds=creds, 
        model_name="gemini-3.1-pro-preview" 
    )
    
    print(f"Đang bóc tách dữ liệu từ file {os.path.basename(pdf_path)}...")
    
    # 3. Gọi hàm API và ÉP SCHEMA (truyền schema_da_nang vào đây)
    response_text = client.send_data_to_AI(
        prompt=prompt_content,
        file_paths=[pdf_path],
        temperature=0.0, # Giữ ở mức 0 để dữ liệu trích xuất chính xác, không sáng tạo
        response_schema=schema_da_nang # <--- Cực kỳ quan trọng
    )
    
    # 4. Xử lý chuỗi trả về an toàn
    try:
        # Làm sạch chuỗi nếu AI trả về kèm thẻ markdown ```json
        clean_json_str = response_text.replace("```json", "").replace("```", "").strip()
        parsed_data = json.loads(clean_json_str)
        
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(parsed_data, f, ensure_ascii=False, indent=2)
            
        print(f"✅ Bóc tách thành công! File xuất ra tại: {output_json_path}")
        
    except json.JSONDecodeError:
        print("❌ Lỗi: Kết quả trả về không đúng chuẩn JSON.")
        print(response_text)

if __name__ == "__main__":
    # Test thử với file của bạn hoặc bất kỳ file môn nào khác
    pdf_file = "D:\CheckTool\FetchJson\dist\GS 10 - Đề Onluyen Vòng 2 - HS.pdf" 
    out_file = "dapan_TA_pdf.json"
    process_exam_universal(pdf_file, out_file)