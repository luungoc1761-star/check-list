# Hệ Thống Đối Soát BOM Sản Xuất & Tồn Kho (BOM vs Stock 61 / Stock Balance)

Ứng dụng web tự động hóa đối soát định mức sản xuất (BOM) và tồn kho thực tế theo cơ chế **FIFO (First-In, First-Out)**, hỗ trợ nhận diện tự động vế Trái/Phải, xuất báo cáo Excel chuyên nghiệp, tích hợp **Quản lý & Phân quyền đa người dùng (RBAC)** cùng cơ chế **hoạt động ổn định, mượt mà và chính xác**.

---

## 📌 Tính Năng Nổi Bật

### 1. 🛡️ Quản Lý Người Dùng & Phân Quyền Truy Cập (RBAC)
- **4 Vai trò chuyên biệt**:
  - 👑 **Quản trị viên (Admin)**: Toàn quyền cấu hình hệ thống, quản lý tài khoản người dùng (tạo mới, đổi quyền, đặt lại mật khẩu, khóa/mở khóa), upload và khôi phục dữ liệu gốc.
  - 📋 **Kế hoạch & Sản xuất (Planner)**: Quản lý hàng chờ FIFO, nạp file Kế hoạch sản xuất, tính toán đối soát vật tư và xuất báo cáo Excel.
  - 🏬 **Thủ kho & Tồn kho (Warehouse)**: Cập nhật file Tồn kho (Stock 61 / Stock Balance), tra cứu số lượng tồn và danh mục linh kiện thiếu.
  - 👁️ **Người xem (Viewer / Công nhân)**: Chế độ chỉ đọc (Read-only) an toàn, theo dõi tiến độ đối soát và tải file Excel.
- **Xác thực an toàn & Đăng nhập 1 chạm**: Mã hóa mật khẩu an toàn (SHA-256 kèm salt ngẫu nhiên) cùng các nút Đăng nhập nhanh 1 chạm tiện dụng khi kiểm thử.
- **Sổ nhật ký hoạt động (Audit Trail)**: Ghi nhận minh bạch mọi thao tác khi chia sẻ cho nhiều người cùng dùng (Ai đăng nhập, ai nạp file, ai thêm/xóa đơn).

### 2. ⚡ Ổn Định & Chống Xung Đột Đa Người Dùng (Multi-User Workspaces)
- **Không gian Cá nhân (Private Workspace)**: Mỗi tài khoản có hàng chờ đối soát riêng biệt được lưu kiên cố. Nhiều người cùng truy cập web đồng thời hoàn toàn độc lập, không bị xung đột hay ghi đè đơn của nhau.
- **Không gian Đội nhóm (Shared Team Workspace)**: Cho phép chuyển sang hàng chờ chung của nhà máy khi cần họp giao ban kế hoạch hoặc sản xuất đồng bộ.
- **Xuất bản & Đồng bộ 1 chạm**: Planner/Admin có thể đẩy hàng chờ cá nhân lên thành hàng chờ chung hoặc sao chép hàng chờ chung về cá nhân để thử nghiệm.
- **Giao dịch file Atomic**: Ghi dữ liệu an toàn qua file tạm và rename nguyên tử, chống lỗi hỏng dữ liệu khi nhiều người cùng thao tác.

### 3. 🧠 Đối Soát Thông Minh & Lọc Dữ Liệu Trọng Tâm Cho Agent (Smart Reconciliation & Agent Context)
- **Tự động chọn Phiên bản (Variant) tối ưu**: Khi một đơn hàng có nhiều phiên bản trong BOM (ví dụ `0004` và `0104`), hệ thống tự động đối chiếu tồn kho thực tế để ưu tiên phiên bản có độ sẵn sàng 100% và cho sản lượng chạy tối đa cao nhất.
- **Nhận diện đa ngôn ngữ & đa nguồn dữ liệu mới**: Tự động chuẩn hóa các cột dữ liệu từ nhiều nguồn khác nhau (`.xlsx`, `.xls`, `.csv`, tiếng Việt, tiếng Anh, SAP/ERP), loại bỏ ký tự ẩn (`\u00A0`, `\ufeff`, `\u200b`), hậu tố số thực `.0` và dấu phẩy hàng nghìn.
- **Phân tích Đơn Cận Sẵn Sàng (Near-Ready Orders)**: Tự động lọc các đơn hàng đã đạt 95%–99% độ đồng bộ (chỉ thiếu từ 1 đến 3 mã linh kiện) để bộ phận Kế hoạch / Mua hàng ưu tiên thúc hàng mở khóa sản xuất nhanh.
- **Gợi ý Mã Linh Kiện Tương Đương Trong Kho**: Tự động phát hiện các mã linh kiện báo thiếu trong BOM nhưng trong kho đang có tồn kho lớn của mã biến thể cùng họ (khác hậu tố mạ/đóng gói như `PA`, `PD`, `PE`, `TI`, `01` vs `51`...).
- **Trích xuất Gói Dữ Liệu Chuẩn Hóa Cho AI Agent (`.JSON` & `.MD`)**: Xuất trực tiếp ngữ cảnh đối soát cấu trúc chuẩn (`schema_version: 2.0-agent-ready`) giúp các AI Agent lọc ngay dữ liệu quan trọng nhất (KPI, Top đơn chạy ngay, Top linh kiện nút cổ chai, Khuyến nghị điều độ kế hoạch).

### 4. 🚀 Trải Nghiệm Mượt Mà & Hiệu Năng Cao (Smooth UX)
- **Phân trang (Pagination) thông minh**: Tự động phân trang cho danh sách hàng trăm đơn hàng và bảng linh kiện, giúp tải nhẹ nhàng và không gây đơ lag trình duyệt.
- **Bộ lọc & Tìm kiếm tức thì**: Tìm kiếm đa năng theo mã đơn hàng, tên sản phẩm, trạng thái ĐỦ/THIẾU và sắp xếp theo số lượng chạy được.
- **Thông báo Toast tinh tế**: Phản hồi trực quan tức thời bằng `st.toast` không gây reload giật màn hình.

### 5. 📦 Logic Đối Soát FIFO BOM Trái & Phải Tuyệt Đối
- Tự động nhận diện vế Trái/Phải từ Product Code hoặc Model (`19`/`20` hoặc `1`/`2`).
- Phân bổ trừ tồn kho FIFO: Đơn hàng nhập trước ưu tiên trước; trong từng đơn hàng trừ vế Trái trước, số tồn kho còn lại tiếp tục trừ cho vế Phải sau.
- Xuất file Excel định dạng cao cấp:
  - Báo cáo chi tiết từng đơn hàng (màu xanh cho dòng OK, đỏ cho dòng THIẾU).
  - Báo cáo tổng hợp toàn bộ các đơn hàng (mỗi đơn 1 sheet).
  - Báo cáo danh sách các đơn hàng đồng bộ đủ mã sẵn sàng chạy.

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

### 3. Tài khoản đăng nhập mẫu sẵn có:
| Tài khoản | Mật khẩu | Vai trò | Quyền hạn chính |
| :--- | :--- | :--- | :--- |
| **`admin`** | `admin123` | Quản trị viên (Admin) | Toàn quyền, quản lý tài khoản, cấu hình & xóa dữ liệu |
| **`planner`** | `plan123` | Kế hoạch & ĐĐSX (Planner) | Nạp kế hoạch, thêm/xóa hàng chờ FIFO, xuất Excel |
| **`warehouse`** | `kho123` | Thủ kho & Vật tư (Warehouse) | Tải file Tồn kho mới, tra cứu mã thiếu |
| **`viewer`** | `view123` | Người xem (Viewer) | Chỉ xem báo cáo & tải Excel (Read-only) |

---

## 📂 Cấu Trúc Thư Mục
```text
├── app.py                             # Giao diện chính Streamlit đa người dùng
├── core_engine.py                     # Bộ xử lý logic BOM, Stock, FIFO & Kiểm tra toàn vẹn
├── auth_manager.py                    # Hệ thống xác thực, RBAC & Audit Log
├── storage_manager.py                 # Quản lý hàng chờ đa người dùng & an toàn giao dịch file
├── streamlit_app.py                   # Điểm khởi chạy cho Streamlit Cloud (đồng bộ app.py)
├── requirements.txt                   # Danh sách thư viện phụ thuộc
├── .persistent_data/                  # Thư mục lưu trữ an toàn (user queues, users.json, logs)
├── Production BOM List 10.2.xlsx      # File mẫu dữ liệu BOM
├── Stock 61.xlsx                      # File mẫu tồn kho Kho 61
└── Kế hoạch.xlsx                     # File mẫu Kế hoạch sản xuất
```