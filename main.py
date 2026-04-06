import sys
import os

class DummyStream:
    """Hút mọi lệnh print() vào hư vô khi không có màn hình console"""
    encoding = 'utf-8' # Giả lập encoding để tránh lỗi của callAPI.py
    def write(self, text): pass
    def flush(self): pass

if sys.stdout is None:
    sys.stdout = DummyStream()
if sys.stderr is None:
    sys.stderr = DummyStream()

import json
import traceback
import pythoncom
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QFileDialog, QTextEdit, QProgressBar, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal
from docx2pdf import convert

# =====================================================================
# THIẾT LẬP ĐƯỜNG DẪN MODULES
# =====================================================================
current_dir = os.path.dirname(os.path.abspath(__file__))
check_answer_dir = os.path.join(current_dir, "CheckAnswer")
get_data_dir = os.path.join(current_dir, "GetData")
export_dir = os.path.join(current_dir, "export")

for d in [check_answer_dir, get_data_dir, export_dir]:
    if d not in sys.path:
        sys.path.insert(0, d)

try:
    from GetData.fetch import load_from_url
    from CheckAnswer.ai_check_de import process_exam_universal
    from CheckAnswer.check_answer import (flatten_pdf_questions, flatten_sys_questions, find_matching_sys_q, 
                        check_image_issues, check_formula_issues, check_TN, check_DS, check_DIEN)
    from export.export_excel import export_to_excel
except ImportError as e:
    print(f"Lỗi Import: {e}. Vui lòng kiểm tra lại cấu trúc thư mục.")
    sys.exit(1)


# =====================================================================
# HÀM RÚT GỌN THÔNG BÁO LỖI CHO NGƯỜI DÙNG DỄ NHÌN
# =====================================================================
def summarize_issues(issues):
    if not issues: return "✅ Khớp hoàn toàn"
    summaries = []
    for issue in issues:
        issue_lower = issue.lower()
        if "thiếu hình ảnh" in issue_lower: summaries.append("❌ Thiếu hình ảnh")
        elif "sai vị trí hình" in issue_lower: summaries.append("❌ Sai vị trí hình")
        elif "lỗi hiển thị" in issue_lower: summaries.append("❌ Lỗi hiển thị (mất/rác ký tự)")
        elif "convert công thức" in issue_lower: summaries.append("❌ Lỗi convert công thức")
        elif "mất công thức" in issue_lower: summaries.append("❌ Mất công thức")
        elif "render" in issue_lower: summaries.append("⚠️ Lỗi render mã LaTeX")
        elif "thiếu nội dung" in issue_lower or "cụt" in issue_lower: summaries.append("❌ Lựa chọn đáp án bị thiếu/cụt")
        elif "không khớp lựa chọn" in issue_lower: summaries.append("❌ Sai nội dung phương án")
        elif "lệch đáp án" in issue_lower: summaries.append("❌ Lệch đáp án đúng")
        elif "số đáp án" in issue_lower or "số ý" in issue_lower: summaries.append("❌ Sai số lượng đáp án/ý")
        elif "dấu thừa" in issue_lower: summaries.append("⚠️ Đáp án có dấu thừa (nên xóa)")
        elif "lệch loại câu" in issue_lower: summaries.append("⚠️ Lệch loại câu hỏi")
        else: summaries.append("❌ Lỗi cấu trúc/định dạng")
    
    # Xóa trùng lặp nhưng giữ nguyên thứ tự
    seen = set()
    unique_summaries = [x for x in summaries if not (x in seen or seen.add(x))]
    return "\n".join(unique_summaries)


# =====================================================================
# WORKER THREAD
# =====================================================================
class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, docx_path, sys_link):
        super().__init__()
        self.docx_path = docx_path
        self.sys_link = sys_link

    def run(self):
        try:
            pythoncom.CoInitialize() 

            base_dir = os.path.dirname(self.docx_path)
            base_name = os.path.splitext(os.path.basename(self.docx_path))[0]
            
            # --- BƯỚC 1: TẢI DATA HỆ THỐNG ĐỂ LẤY ID ---
            self.log_signal.emit(f"🔄 Đang tải dữ liệu từ Link hệ thống...")
            self.progress_signal.emit(10)
            sys_data, json_id = load_from_url(self.sys_link)
            if not json_id: json_id = "UnknownID"
            
            # Định nghĩa các đường dẫn (File excel lưu vào folder export)
            pdf_path = os.path.join(base_dir, f"{base_name}.pdf")
            ai_json_path = os.path.join(base_dir, f"{base_name}_ai.json")
            sys_json_path = os.path.join(base_dir, f"{base_name}_{json_id}_sys.json")
            excel_path = os.path.join(export_dir, f"{base_name}_{json_id}_KetQua.xlsx")

            with open(sys_json_path, 'w', encoding='utf-8') as f:
                json.dump(sys_data, f, ensure_ascii=False, indent=2)
            self.log_signal.emit(f"✅ Tải hệ thống thành công! (Mã đề: {json_id})")
            self.progress_signal.emit(30)

            # --- BƯỚC 2: CHECK TÁI SỬ DỤNG AI JSON ---
            if os.path.exists(ai_json_path):
                self.log_signal.emit(f"♻️ TÌM THẤY file AI bóc tách từ trước ({base_name}_ai.json).")
                self.log_signal.emit(f"♻️ Tái sử dụng để tiết kiệm thời gian và Token...")
                self.progress_signal.emit(60)
            else:
                self.log_signal.emit(f"🔄 Đang chuyển đổi '{base_name}.docx' sang PDF...")
                convert(self.docx_path, pdf_path)
                
                self.log_signal.emit(f"🔄 AI đang bóc tách dữ liệu từ PDF (Sẽ mất vài phút)...")
                self.progress_signal.emit(45)
                process_exam_universal(pdf_path, ai_json_path)
                self.log_signal.emit("✅ AI bóc tách thành công!")
                self.progress_signal.emit(60)

            # --- BƯỚC 3: ĐỐI CHIẾU & XUẤT EXCEL ---
            self.log_signal.emit(f"🔄 Bắt đầu đối chiếu dữ liệu...")
            self.progress_signal.emit(80)
            self.compare_and_export_excel(ai_json_path, sys_json_path, excel_path)
            
            self.progress_signal.emit(100)
            self.log_signal.emit(f"🎉 HOÀN THÀNH! File kết quả đã được lưu trong thư mục 'export'.")
            self.finished_signal.emit(True, excel_path)

        except Exception as e:
            error_msg = traceback.format_exc()
            self.log_signal.emit(f"❌ LỖI NGHIÊM TRỌNG:\n{error_msg}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

    def compare_and_export_excel(self, pdf_path, sys_path, excel_path):
        pdf_qs = flatten_pdf_questions(pdf_path)
        sys_qs = flatten_sys_questions(sys_path)
        sys_contents = [q["content"] for q in sys_qs]

        results_data = []
        total_ok = total_warn = total_err = 0

        for pdf_q in pdf_qs:
            cau_so_goc = f"Câu {pdf_q.get('so', '?')}"
            qtype = pdf_q.get("type", "UNKNOWN")
            
            sys_q, score, method = find_matching_sys_q(pdf_q, sys_qs, sys_contents)

            if sys_q is None:
                results_data.append({
                    "STT": cau_so_goc,
                    "Câu hỏi đề gốc": cau_so_goc,
                    "Câu hỏi đề hệ thống": "-",
                    "Lỗi": "❌ Không tìm thấy trên hệ thống"
                })
                total_err += 1
                continue

            # Lấy vị trí câu trên hệ thống
            sys_idx = sys_qs.index(sys_q) + 1
            cau_so_sys = f"Câu {sys_idx}"
            
            issues = []
            
            pdf_co_hinh = pdf_q.get("co_hinh", False)
            image_issues = check_image_issues(pdf_co_hinh, sys_q.get("content", ""), sys_q.get("options", {}))
            if image_issues: issues.extend(image_issues)

            formula_issues = check_formula_issues(pdf_q.get("content", ""), sys_q.get("content", ""), sys_q.get("options", {}))
            if formula_issues: issues.extend(formula_issues)

            if qtype == "TN": check_TN(pdf_q, sys_q, issues)
            elif qtype == "DS": check_DS(pdf_q, sys_q, issues)
            elif qtype == "DIEN": check_DIEN(pdf_q, sys_q, issues)
            
            if sys_q.get("type") != qtype:
                issues.append(f"⚠️ Lệch loại câu")

            # Xử lý thống kê lỗi
            if issues:
                errs = sum(1 for i in issues if "❌" in i)
                warns = sum(1 for i in issues if "⚠️" in i and "❌" not in i)
                if errs > 0: total_err += 1
                elif warns > 0: total_warn += 1
                else: total_ok += 1
            else:
                total_ok += 1

            # Nén mảng lỗi dài dòng thành tóm tắt
            error_str = summarize_issues(issues)

            results_data.append({
                "STT": cau_so_goc,
                "Câu hỏi đề gốc": cau_so_goc,
                "Câu hỏi đề hệ thống": cau_so_sys,
                "Lỗi": error_str
            })

        summary_data = {
            "Tổng câu đề gốc (AI bóc)": len(pdf_qs),
            "Tổng câu hệ thống (Web)": len(sys_qs),
            "✅ Số câu khớp hoàn toàn": total_ok,
            "⚠️ Số câu có CẢNH BÁO (Sai định dạng nhỏ)": total_warn,
            "❌ Số câu có LỖI (Mất nội dung/Sai đáp án)": total_err
        }

        # Gọi file export_excel.py
        export_to_excel(results_data, summary_data, excel_path)


# =====================================================================
# GIAO DIỆN (UI)
# =====================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Check Đề Trắc Nghiệm - Version 1.1")
        self.resize(850, 650)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        link_layout = QHBoxLayout()
        link_label = QLabel("Link Preview JSON:")
        link_label.setFixedWidth(130)
        self.link_input = QLineEdit()
        self.link_input.setFixedHeight(30)
        link_layout.addWidget(link_label)
        link_layout.addWidget(self.link_input)
        layout.addLayout(link_layout)

        file_layout = QHBoxLayout()
        file_label = QLabel("File Đề Gốc (.docx):")
        file_label.setFixedWidth(130)
        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        self.file_input.setFixedHeight(30)
        self.btn_browse = QPushButton("Chọn File")
        self.btn_browse.setFixedHeight(30)
        self.btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(self.btn_browse)
        layout.addLayout(file_layout)

        self.btn_run = QPushButton("🚀 BẮT ĐẦU KIỂM TRA ĐỀ")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; font-size: 15px; border-radius: 5px;")
        self.btn_run.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_run)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(20)
        layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("background-color: #1e1e1e; color: #00fa9a; font-family: Consolas;")
        layout.addWidget(self.log_output)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file đề", "", "Word Documents (*.docx)")
        if file_path: self.file_input.setText(file_path)

    def log_msg(self, msg):
        self.log_output.append(msg)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def start_processing(self):
        sys_link = self.link_input.text().strip()
        docx_path = self.file_input.text().strip()

        if not sys_link or not docx_path:
            QMessageBox.warning(self, "Lỗi Nhập Liệu", "Vui lòng nhập Link và Chọn File Word!")
            return

        self.btn_run.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.link_input.setEnabled(False)
        self.log_output.clear()
        
        self.worker = WorkerThread(docx_path, sys_link)
        self.worker.log_signal.connect(self.log_msg)
        self.worker.progress_signal.connect(self.progress_bar.setValue)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success, result_path):
        self.btn_run.setEnabled(True)
        self.btn_browse.setEnabled(True)
        self.link_input.setEnabled(True)
        if success:
            QMessageBox.information(self, "Thành Công", f"Đã xuất báo cáo tại:\n{result_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())