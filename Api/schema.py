schema_da_nang = {
    "type": "ARRAY",
    "description": "Mảng tuần tự các khối dữ liệu trong đề thi từ trên xuống dưới.",
    "items": {
        "type": "OBJECT",
        "properties": {
            "loai_block": {
                "type": "STRING",
                "enum": ["cau_hoi_don", "nhom_cau_hoi"],
                "description": "Xác định khối dữ liệu là câu hỏi lẻ hay một nhóm câu hỏi dùng chung dữ kiện/đoạn văn."
            },
            
            # --- CẤU TRÚC CHO CÂU HỎI ĐƠN ---
            "loai_cau_hoi": {
                "type": "STRING",
                "enum": ["trac_nghiem", "dung_sai", "tra_loi_ngan", "tu_luan"],
                "description": "Phân loại câu hỏi."
            },
            "cau_so": {
                "type": "INTEGER",
                "description": "Số thứ tự của câu hỏi (nếu có)."
            },
            "noi_dung_cau_hoi": {
                "type": "STRING",
                "description": "Nội dung câu hỏi. Chú ý BẮT BUỘC: Nếu câu hỏi có chứa hình ảnh/đồ thị, phải chèn cụm từ [HÌNH_ẢNH] vào đúng vị trí xuất hiện của hình ảnh đó trong văn bản. Giữ nguyên định dạng LaTeX cho công thức."
            },
            "co_hinh": {
                "type": "BOOLEAN",
                "description": "Đánh dấu True nếu câu hỏi có chứa hình ảnh, đồ thị hoặc sơ đồ. Tuyệt đối KHÔNG chọn True nếu nó là BẢNG BIỂU (Table)."
            },
            "co_bang": {
                "type": "BOOLEAN",
                "description": "Đánh dấu True nếu câu hỏi có chứa bảng biểu (table)."
            },
            "cac_lua_chon": {
                "type": "ARRAY",
                "description": "Danh sách các đáp án trắc nghiệm (A, B, C, D) hoặc đúng/sai (a, b, c, d).",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "nhan": {"type": "STRING", "description": "Ký hiệu đáp án (A, B, a, b...)"},
                        "noi_dung": {"type": "STRING", "description": "Nội dung chi tiết của đáp án đó"}
                    },
                    "required": ["nhan", "noi_dung"]
                }
            },
            "dap_an_dung": {
                "type": "STRING",
                "description": "Trích xuất ký tự đáp án (A, B, C, hoặc D) dựa trên dấu hiệu định dạng (gạch chân, in đậm, highlight).",
                "nullable": True
            },

            # --- CẤU TRÚC CHO NHÓM CÂU HỎI / HỌC LIỆU DÙNG CHUNG ---
            "du_kien_chung": {
                "type": "STRING",
                "description": "Đoạn văn (Tiếng Anh) hoặc phần dữ kiện, mô tả biểu đồ chung (Toán/Lý/Hóa)."
            },
            "danh_sach_cau_hoi_con": {
                "type": "ARRAY",
                "description": "Danh sách các câu hỏi phụ thuộc vào du_kien_chung bên trên.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "loai_cau_hoi": {
                            "type": "STRING",
                            "enum": ["trac_nghiem", "dung_sai", "tra_loi_ngan", "tu_luan"]
                        },
                        "cau_so": {"type": "INTEGER"},
                        "noi_dung_cau_hoi": {
                            "type": "STRING",
                            "description": "Nội dung câu hỏi. Chú ý BẮT BUỘC: Nếu câu hỏi có chứa hình ảnh/đồ thị, phải chèn cụm từ [HÌNH_ẢNH] vào đúng vị trí xuất hiện của hình ảnh đó trong văn bản. Giữ nguyên định dạng LaTeX cho công thức."
                        },
                        "co_hinh": {
                            "type": "BOOLEAN",
                            "description": "Đánh dấu True nếu câu hỏi có chứa hình ảnh, đồ thị hoặc sơ đồ. Tuyệt đối KHÔNG chọn True nếu nó là BẢNG BIỂU (Table)."
                        },
                        "co_bang": {
                            "type": "BOOLEAN",
                            "description": "Đánh dấu True nếu câu hỏi có chứa bảng biểu (table)."
                        },
                        "cac_lua_chon": {
                            "type": "ARRAY",
                            "items": {
                                "type": "OBJECT",
                                "properties": {
                                    "nhan": {"type": "STRING"},
                                    "noi_dung": {"type": "STRING"}
                                },
                                "required": ["nhan", "noi_dung"]
                            }
                        },
                        "dap_an_dung": {
                            "type": "STRING",
                            "description": "Trích xuất đáp án dựa trên dấu hiệu định dạng (gạch chân, in đậm, highlight).",
                            "nullable": True
                        }
                    },
                    # SỬA Ở ĐÂY: Bỏ bớt "cac_lua_chon" khỏi required để tránh lỗi với câu tự luận không có A,B,C,D
                    "required": ["loai_cau_hoi", "noi_dung_cau_hoi"]
                }
            }
        },
        # SỬA QUAN TRỌNG NHẤT Ở ĐÂY: Chỉ bắt buộc "loai_block". Khi đó AI sẽ không điền thừa dữ kiện vào JSON.
        "required": ["loai_block"] 
    }
}