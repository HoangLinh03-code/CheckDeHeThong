<!-- @format -->

# TÀI LIỆU KỸ THUẬT: HỆ THỐNG ĐỐI CHIẾU VÀ SOÁT LỖI ĐỀ THI TỰ ĐỘNG (CHECK ĐỀ HỆ THỐNG)

## 1. TỔNG QUAN HỆ THỐNG

Mục tiêu cốt lõi của phần mềm là giải quyết bài toán "kiểm thử dữ liệu giáo dục": Tự động hóa hoàn toàn quy trình đối chiếu giữa **Tài liệu gốc của giáo viên (File Word .docx)** và **Dữ liệu đã được đưa lên hệ thống học trực tuyến (định dạng JSON)**.

Hệ thống có khả năng vượt qua rào cản về "trộn đề" (shuffling), phân tích sâu sắc các công thức Toán/Hóa (LaTeX, MathType) và nhận diện cấu trúc đa phương tiện (Hình ảnh, Bảng biểu) để phát hiện chính xác các lỗi: Sai đáp án, lệch nội dung, mất công thức, thiếu hình ảnh hoặc sai vị trí hiển thị.

### Chi tiết Dữ liệu Đầu vào và Đầu ra

- **Đầu vào (Input):**
  - **Tập tin Word (.docx):** Đề thi gốc do giáo viên soạn (chứa Text, MathType, Hình ảnh).
  - **URL/Link JSON Hệ thống:** Đường dẫn chứa dữ liệu đề thi đã được platform (vd: onluyen.vn) render.
- **Đầu ra (Output):**
  - **Tập tin PDF & JSON AI:** Hệ thống tự động convert Word sang PDF và dùng AI bóc tách thành JSON chuẩn hóa.
  - **Tập tin Báo Cáo Excel (.xlsx):** Báo cáo chi tiết từng câu hỏi, thống kê tỷ lệ khớp, và highlight cụ thể các lỗi sai lệch nội dung/đáp án.

---

## 2. KIẾN TRÚC MÃ NGUỒN VÀ CHỨC NĂNG MODULE

### Sơ đồ cấu trúc thư mục

```
CheckDeHeThong/
├── Api/                        # Lớp giao tiếp với Google Vertex AI
│ ├── callAPI.py                # Xác thực credentials, init VertexClient, xử lý file gửi lên AI
│ └── schema.py                 # Khai báo cấu trúc JSON Schema ép buộc AI phải tuân thủ
├── CheckAnswer/                # Core Logic của hệ thống
│ ├── scanAI.py                 # Điều khiển luồng AI bóc tách từ PDF sang JSON
│ ├── check_answer.py           # (Engine) Thuật toán so khớp văn bản, chuẩn hóa công thức và bắt lỗi
│ ├── prompt.txt                # System Prompt chứa luật bóc tách khắt khe cho AI
│ └── testAI.py                 # File script phục vụ debug thuật toán độc lập
├── GetData/                    # Module tương tác với API của hệ thống bên thứ 3
│ └── fetch.py                  # Tải dữ liệu JSON từ URL và parse thành cấu trúc cơ bản
├── export/                     # Module kết xuất báo cáo
│ └── export_excel.py           # Định dạng, tô màu và xuất báo cáo Excel bằng openpyxl
├── main.py                     # Điểm khởi chạy chính: Chứa giao diện PyQt5 và quản lý luồng (WorkerThread)
├── checkDe.py                  # Phiên bản xử lý logic trực tiếp từ file DOCX (XML parsing) - Backup/Alternative
├── .env                        # Chứa thông tin Service Account Google Cloud (PRIVATE_KEY, PROJECT_ID...)
├── .gitignore                  # Cấu hình bỏ qua file tạm, venv, cache khi đẩy lên Github
└── requirements.txt            # Danh sách thư viện phụ thuộc (Dependencies)
```

### Chức năng chi tiết của các Module chính

1. **`main.py` (Core UI & Threading Controller)**
   - Quản lý giao diện người dùng bằng PyQt5.
   - Xử lý tiến trình Đa luồng (Multi-threading) thông qua `WorkerThread` để giao diện không bị đơ khi AI xử lý hoặc tải dữ liệu.
   - Quản lý môi trường thực thi: Xử lý đường dẫn tương đối/tuyệt đối (`sys._MEIPASS`) để tương thích với quá trình đóng gói ra file `.exe` bằng PyInstaller.

2. **`CheckAnswer/check_answer.py` (Matching & Validation Engine)**
   - **Xử lý nhiễu (Noise Reduction):** Module chứa thuật toán làm sạch thẻ HTML, chuyển đổi Unicode Math sang LaTeX chuẩn, đồng nhất các loại dấu gạch ngang, mũi tên phản ứng hóa học (Hàm `normalize`).
   - **Fuzzy Matching:** Sử dụng `difflib.SequenceMatcher` để tìm kiếm câu hỏi tương đương giữa 2 tập dữ liệu bất chấp việc hệ thống đã trộn ngẫu nhiên thứ tự câu và thứ tự đáp án (A, B, C, D).
   - **Kiểm tra Media:** Xác minh tính toàn vẹn của Hình ảnh/Bảng biểu qua cơ chế đối chiếu Tag vị trí tương đối (`[HÌNH_ẢNH]`, `[BẢNG_BIỂU]`).

3. **`Api/callAPI.py` & `Api/schema.py` (AI Orchestrator)**
   - Tương tác với SDK mới `google.genai`.
   - Ép buộc định dạng trả về thông qua tính năng `response_schema` (Controlled Generation) được định nghĩa trong `schema_da_nang`. Đảm bảo AI luôn trả về cấu trúc JSON ổn định, loại bỏ rủi ro AI tự ý thêm bớt trường dữ liệu.

4. **`GetData/fetch.py` (Data Pipeline)**
   - Giả lập Request Headers, tải dữ liệu JSON từ các nền tảng EduTech. Phân rã cấu trúc câu đơn (`typeData=0`) và câu nhóm/bài đọc hiểu (`typeData=1`).

---

## 3. LUỒNG XỬ LÝ DỮ LIỆU CHI TIẾT (WORKFLOW)

Quy trình vòng đời đối chiếu 1 mã đề thi trải qua 5 giai đoạn hoàn toàn tự động:

1. **Fetch & Download (Tải dữ liệu hệ thống):**
   Hệ thống đọc link URL từ UI, module `fetch.py` gọi API tải JSON của đề thi trên nền tảng về máy và lưu tạm.

2. **Document Conversion (Chuyển đổi tài liệu gốc):**
   Hệ thống gọi thư viện `docx2pdf` (thông qua COM interop của Windows) để biến đổi file Word `.docx` của giáo viên thành định dạng PDF.

3. **AI Data Extraction (AI Bóc tách cấu trúc):**
   Gửi file PDF + `prompt.txt` + `schema.py` lên Google Vertex AI (Model: `gemini-3.1-pro-preview`). AI sẽ "đọc hiểu" file PDF và trả về một file `_ai.json` chứa toàn bộ metadata: Nội dung câu hỏi, Phân loại (Trắc nghiệm/Đúng Sai/Điền), Đáp án đúng, Vị trí hình ảnh.
   _(Lưu ý: Hệ thống có cơ chế Cache. Nếu file `_ai.json` đã tồn tại, nó sẽ bỏ qua bước này để tiết kiệm API Token)._

4. **Semantic Matching & Error Detection (Đối chiếu ngữ nghĩa & Bắt lỗi):**
   Module `check_answer.py` nạp cả 2 file JSON (AI và Hệ thống). Nó thực hiện vòng lặp quét chéo:
   - Dùng thuật toán so khớp (Similarity) để "ghép đôi" câu hỏi gốc với câu hỏi trên hệ thống (chống trộn đề).
   - Khi đã ghép đúng cặp, hệ thống kiểm tra logic sâu: Đáp án gốc vs Đáp án hệ thống, Tỷ lệ mất chữ, Lỗi mất công thức LaTeX, Lỗi thiếu hoặc nhảy vị trí hình ảnh.

5. **Report Generation (Kết xuất báo cáo):**
   Đẩy toàn bộ dữ liệu lỗi vào list. Gọi `export_excel.py` để tô màu, căn lề và xuất ra file Excel trực quan giúp giáo viên/QC dễ dàng rà soát.

---

## 4. CƠ CHẾ MATCHING VÀ BẮT LỖI CHUYÊN SÂU

Dự án này sở hữu engine xử lý text chuyên biệt cho dữ liệu Giáo dục (Đặc biệt là các môn Tự nhiên). Các thuật toán trong `check_answer.py`:

### 4.1. Thuật toán Đơn giản hóa Công thức (Math Simplification)

Để so sánh công thức giữa Word (được AI đọc) và Hệ thống (thường render bằng MathJax/KaTeX), hàm `simplify_math_text` thực hiện rũ bỏ toàn bộ cấu trúc LaTeX.

- Cạo sạch các lệnh LaTeX (`\frac`, `\sqrt`, `\left\{`, `\rightarrow`).
- Chỉ giữ lại ký tự chữ (a-z), số (0-9) và dấu bằng (`=`).
- Đồng nhất các ký hiệu đặc biệt (Ví dụ: `°C`, `℃`, `\circ` đều về chung 1 chuẩn `0c`).
  => **Kết quả:** Giúp phát hiện mất công thức.

### 4.2. Cơ chế Track Vị Trí Hình Ảnh (Image Absolute Positioning)

AI không thể xuất hình ảnh trực tiếp ra JSON. Thay vào đó, System Prompt ép buộc AI chèn tag `[HÌNH_ẢNH]` vào đúng tọa độ mà hình ảnh xuất hiện trong văn bản.

- Hệ thống cắt chuỗi văn bản làm 2 nửa (trước và sau tag `[HÌNH_ẢNH]`).
- So sánh độ dài chuỗi này với vị trí thẻ `<img src...>` trên hệ thống.
- Nhờ đó phát hiện được lỗi: _Hình ảnh bị nhảy lên đầu câu_ hoặc _Hình ảnh chèn nhầm vào đáp án A, B, C, D_.

### 4.3. Quản lý Dạng câu hỏi Đúng/Sai

Mỗi câu Đúng/Sai chứa 4 ý phụ (a, b, c, d). Khi hệ thống web trộn đáp án, thứ tự a, b, c, d sẽ bị xáo trộn. Thuật toán thực hiện `find_best_match_idx` cho TỪNG ý phụ để tìm ra chính xác nó đang nằm ở đâu trên hệ thống, trước khi tiến hành đối chiếu kết quả True/False.

---

## 5. HƯỚNG DẪN CÀI ĐẶT MÔI TRƯỜNG & CI/CD

### Yêu cầu hệ thống

- **Hệ điều hành:** Windows 10/11 (Bắt buộc phải có hệ điều hành Windows và cài đặt Microsoft Office Word để thư viện `docx2pdf` có thể vận hành qua giao thức `pythoncom`).
- **Python:** Phiên bản 3.10 trở lên.

### Cài đặt môi trường phát triển

1. Clone dự án từ Github:
   git clone <repo_url>
   cd CheckDeHeThong
2. Khởi tạo môi trường ảo (Virtual Environment) & Cài đặt dependencies:
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
3. Cấu hình API Key:
   - Tạo file `.env` tại folder Api.

---

## 6. HƯỚNG DẪN VẬN HÀNH DÀNH CHO USER

1. **Khởi động ứng dụng:**
   Mở Terminal, đảm bảo đã activate venv và chạy lệnh:
   ```bash
   python main.py
   ```
2. **Thao tác trên giao diện:**
   - **Link Preview JSON:** Dán đường link URL lấy từ hệ thống (VD: URL có param `?url=...`).
   - **File Đề Gốc (.docx):** Bấm nút `Chọn File` và trỏ đến file Word gốc của đề thi.
   - Bấm **🚀 BẮT ĐẦU KIỂM TRA ĐỀ**.
3. **Theo dõi:** Thanh Progress Bar và Log Console sẽ hiển thị chi tiết tiến trình (Tải data -> Convert PDF -> AI bóc tách -> Đối chiếu).
4. **Nhận kết quả:**
   Khi hoàn thành, hệ thống sẽ bật thông báo. Thư mục `export/` sẽ sinh ra file Excel `[Tên_Đề]_[ID]_KetQua.xlsx`.
