# Hệ Thống Đối Soát BOM Sản Xuất & Tồn Kho (BOM vs Stock Balance)

Ứng dụng web tự động hóa đối soát định mức sản xuất (BOM) và tồn kho thực tế theo cơ chế **FIFO (First-In, First-Out)**, hỗ trợ nhận diện tự động vế Trái/Phải và xuất báo cáo Excel chuyên nghiệp.

---

## 📌 Tính Năng Chính

1. **Nhận diện Thông minh Mã Sản Phẩm (Product Code):**
   - **Mã bắt đầu bằng số `1`**: 4 chữ số đầu là số đơn hàng, 2 chữ số tiếp theo là mã phân biệt vế (`19`: Bên Trái, `20`: Bên Phải). Ví dụ: `122019` (đơn 1220 bên trái), `122020` (đơn 1220 bên phải).
   - **Mã bắt đầu bằng số `8`**: Bỏ số 8, 4 chữ số tiếp theo là số đơn hàng, chữ số thứ 6 là phân biệt vế (`1`: Bên Trái, `2`: Bên Phải). Ví dụ: `815731` (đơn 1573 bên trái), `815732` (đơn 1573 bên phải).
   - Tự động lấy danh mục linh kiện (`Component Code`) và định mức (`Qty`) cho cả hai vế.

2. **Bảng Nổi Nhập Liệu & Hàng Chờ (Queue):**
   - Hộp nổi nhập liệu trực quan với 2 ô: **Đơn hàng** (tìm kiếm / chọn đơn) và **Số lượng** (bộ sản phẩm).
   - Bấm **OK**: Đơn hàng tự động nhảy xuống **Khung hàng chờ** để tính toán đối soát.
   - Thêm liên tiếp nhiều đơn hàng vào danh sách chờ tính toán.

3. **Cơ chế Trừ Tồn Kho FIFO Logic Tuyệt Đối:**
   - Đơn hàng nhập trước được ưu tiên cấp đủ vật tư trước theo đúng thứ tự.
   - Các đơn hàng tiếp theo sẽ lấy từ phần tồn kho còn lại sau khi các đơn trước đã trừ.
   - Số lượng cần = `Số lượng đơn hàng` × `Định mức BOM` (cộng gộp đủ cả Trái và Phải).

4. **Trực quan Hóa Kết Quả & Xuất Báo Cáo Excel:**
   - **Đủ vật tư**: Đánh dấu trạng thái **OK**, tô dòng **màu xanh lá**.
   - **Thiếu vật tư**: Đánh dấu trạng thái **THIẾU**, tô dòng **màu đỏ** và hiển thị rõ số lượng thiếu ở cột Ghi chú (ví dụ: `Thiếu 50`).
   - Hiển thị chi tiết vị trí **BIN** và số lô **BATCH** trong kho.
   - Hỗ trợ xuất file Excel định dạng chuyên nghiệp:
     - Xuất file riêng cho từng đơn hàng.
     - Xuất file gộp toàn bộ các đơn hàng trong hàng chờ (mỗi đơn 1 sheet).

5. **Nhập File Dữ Liệu Mới Tiện Lợi:**
   - Nút nhập file **Stock Balance** mới.
   - Nút nhập file **Production BOM** mới.
   - Tự động nhận diện cấu trúc tiêu đề và các trường dữ liệu giống file mẫu, không cần sửa đổi file gốc.
   - Tự động sử dụng 2 file dữ liệu mẫu có sẵn nếu chưa tải file mới.

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Ứng Dụng

### 1. Cài đặt thư viện
```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng web
```bash
streamlit run app.py
```
Ứng dụng sẽ chạy tại địa chỉ: `http://localhost:8501`

---

## 📂 Cấu Trúc Thư Mục
```text
├── app.py                             # Giao diện chính Streamlit
├── core_engine.py                     # Bộ xử lý logic BOM, Stock & Excel
├── streamlit_app.py                   # Điểm khởi chạy cho Streamlit Cloud
├── requirements.txt                   # Danh sách thư viện phụ thuộc
├── .streamlit/
│   └── config.toml                    # Cấu hình giao diện Streamlit
├── Production BOM List 10.2.xlsx      # File mẫu dữ liệu BOM
└── Stock Balance With Batch (3).xlsx  # File mẫu tồn kho Stock Balance
```