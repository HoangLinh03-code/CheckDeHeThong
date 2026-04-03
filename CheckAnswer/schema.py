# schema.py
from pydantic import BaseModel, Field
from typing import List, Optional

class DapAn(BaseModel):
    nhan: str = Field(description="Nhãn của đáp án. Ví dụ: 'A', 'B', 'a', 'b'")
    noi_dung: str = Field(description="Nội dung chi tiết của đáp án. Giữ nguyên định dạng LaTeX.")

class CauHoi(BaseModel):
    loai_cau: str = Field(description="Loại câu hỏi: Phải là 'TN', 'DS', hoặc 'DIEN'")
    cau_so: int = Field(description="Số thứ tự của câu hỏi trong đề")
    noi_dung_cau_hoi: str = Field(description="Nội dung trọn vẹn của câu hỏi. Giữ nguyên định dạng LaTeX.")
    co_hinh_anh: bool = Field(description="Trả về true nếu câu hỏi/đáp án có hình ảnh, biểu đồ.")
    cac_dap_an: List[DapAn] = Field(description="Danh sách các đáp án lựa chọn (Để trống nếu là câu Điền khuyết).")
    dap_an_dung: str = Field(description="Đáp án đúng. Câu TN: ghi nhãn. Câu DS: ghi định dạng 'a-Đúng, b-Sai...'. Câu Điền: ghi giá trị.")
    loi_giai: str = Field(description="Lời giải chi tiết nếu có. Nếu không có để chuỗi rỗng.")

class KetQuaTrichXuat(BaseModel):
    data: List[CauHoi] = Field(description="Danh sách toàn bộ các câu hỏi trích xuất được từ đề thi")