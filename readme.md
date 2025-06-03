# 🚀 POSM Cost Calculator

Ứng dụng này giúp bạn **tính toán chi phí và phân bổ POSM** (Point of Sales Materials) một cách **nhanh chóng** và **chính xác**.

---

## 📝 Hướng dẫn sử dụng

### 1️⃣ Chuẩn bị file dữ liệu

Bạn cần chuẩn bị **4 file Excel** sau:

| Tên file              | Nội dung & Yêu cầu                                                                                       |
|-----------------------|---------------------------------------------------------------------------------------------------------|
| `fact_display.xlsx`   | Dữ liệu trưng bày thực tế.<br/>*Gồm 3 cột:* `store`, `model`, `display` (lấy từ survey posm deployment, chọn mục No) |
| `dim_storelist.xlsx`  | Danh sách cửa hàng.<br/>*Phải có cột:* `Store name` hoặc `shop` và `address` (địa chỉ giao hàng).<br/>*Shop non prt* có thể điền tên TDS. |
| `dim_model.xlsx`      | Thông tin model sản phẩm.<br/>*Cột `priority` để phân loại ưu tiên model:*<br/>`1` - quan trọng (áp dụng buffer 30% & chia chẵn cho 5)<br/>`2` - ít quan trọng (chỉ chia chẵn cho 5) |
| `dim_posm.xlsx`       | File này phải có **2 sheet:**<br/>- `posm`<br/>- `price`                                                |

> **Lưu ý:**  
> - Trong các file Excel **không được có dữ liệu nào bị `#N/A`**.  
> - Code sẽ **bỏ qua dữ liệu bị lỗi** đó (thường là cột address do hay dùng vlookup).

---

### 2️⃣ Tải lên các file

- Ở giao diện chính, bạn sẽ thấy **4 ô để tải lên file**.
- Nhấn vào từng ô và **chọn đúng file tương ứng** từ máy tính của bạn.

---

### 3️⃣ Thực hiện tính toán

- Sau khi tải đủ 4 file, nhấn nút **`🚀 Calculate Report`** để bắt đầu tính toán.
- Hệ thống sẽ xử lý và hiển thị các bảng kết quả:

  - **Tổng hợp POSM**
  - **Tổng hợp Model-POSM**
  - **Tổng hợp theo địa chỉ**
  - **POSM theo nhóm Care** (Front Load, Dryer)
  - **POSM theo nhóm SDA** (các nhóm còn lại)

---

### 4️⃣ Tải kết quả về máy

- Sau khi tính toán xong, bạn có thể nhấn nút **`📥 Download Report as Excel`** để tải toàn bộ kết quả về dưới dạng file Excel.

---

### 5️⃣ Lưu ý

- Nếu có lỗi, **kiểm tra lại định dạng và nội dung các file Excel**.
- Nếu cần hỗ trợ, liên hệ qua **LinkedIn của tác giả ở cuối trang**.

---

> **Chúc bạn sử dụng ứng dụng hiệu quả!**