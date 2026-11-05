# HỒ SƠ ĐỐI SOÁT & TÀI LIỆU KỸ THUẬT DỰ ÁN VÍ ĐIỆN TỬ PAYKX

**Người thực hiện:** Trưởng dự án PayKX  
**Đối tượng áp dụng:** Ví điện tử thanh toán & hoàn tiền cho Sinh viên  
**Thời gian lập:** Tháng 11/2026  

---

## PHẦN 1. PHÂN TÍCH CẤU TRÚC DỮ LIỆU EXCEL (DATABASE STANDARDS)

Để đảm bảo dữ liệu giao dịch chuẩn Database, tuyệt đối không **Merge Cell** (trộn ô) và không để **Empty Rows** (dòng trống) giữa các bản ghi. Cấu trúc được chia rõ ràng thành **Field** (Cột/Trường thuộc tính) và **Record** (Dòng/Bản ghi).

### 1. Phân định Field (Trường dữ liệu - Cột)
Mỗi Field đại diện cho một thuộc tính duy nhất của giao dịch:
*   `Field 1: Trans_ID` (Mã giao dịch) – Khóa chính (Primary Key), định dạng Văn bản/Mã chuẩn (ví dụ: `TXN1001`).
*   `Field 2: User_ID` (Mã sinh viên/Mã người dùng) – Khóa ngoại liên kết với Bảng Người Dùng.
*   `Field 3: User_Name` (Tên khách hàng) – Trường cần tra cứu bằng hàm `VLOOKUP` kết hợp `TRIM`.
*   `Field 4: Service_Type` (Loại dịch vụ) – Nạp điện thoại, Thanh toán canteen, Chuyển tiền, Hoàn tiền Cashback.
*   `Field 5: Amount` (Số tiền giao dịch) – Định dạng Số (Currency/Number).
*   `Field 6: Fee_Bank` (Phí ngân hàng đối tác) – Định dạng Số.
*   `Field 7: Status` (Trạng thái) – `SUCCESS`, `PENDING`, `FAILED`.
*   `Field 8: Trans_Date` (Thời gian giao dịch) – Định dạng `YYYY-MM-DD HH:MM:SS`.

### 2. Phân định Record (Bản ghi dữ liệu - Dòng)
*   Mỗi **Record** (tương ứng từ Dòng 2 trở đi trong Sheet dữ liệu) là một giao dịch hoàn chỉnh chứa đầy đủ giá trị của tất cả các Field nêu trên.
*   **Nguyên tắc thiết kế:** 
    1. Hàng 1 (Header Row) chứa tên các Field.
    2. Từ Hàng 2 trở đi là các Record liên tục, tuyệt đối không xen kẽ các hàng Tổng (Subtotal) mid-table hoặc ô gộp (Merged Cells).

---

## PHẦN 2. TRIỂN KHAI CHI TIẾT BỘ 3 FILE (WORD, EXCEL, POWERPOINT)

### 2.1. TÀI LIỆU 1 (WORD): CẤU TRÚC SRS TÍNH NĂNG HOÀN TIỀN (5 TRANG)

Below is the structured layout and style mapping for Microsoft Word. 

```text
[MỤC LỤC TỰ ĐỘNG - AUTO TABLE OF CONTENTS]
(Tạo tự động qua References > Table of Contents > Automatic Table 1)
--------------------------------------------------------------------------------
1. TỔNG QUAN DỰ ÁN VÍ ĐIỆN TỬ PAYKX ........................................... Trang 1
   1.1 Giới thiệu hệ thống ..................................................... Trang 1
   1.2 Mục tiêu tính năng Cashback (Hoàn tiền) .................................. Trang 1
2. YÊU CẦU NĂNG LỰC HỆ THỐNG .................................................. Trang 2
   2.1 Đối tượng người dùng (Sinh viên) ........................................ Trang 2
   2.2 Ràng buộc vận hành & Ngân hàng liên kết .................................. Trang 2
3. YÊU CẦU CHỨC NĂNG (FUNCTIONAL REQUIREMENTS) ................................. Trang 3
   3.1 Chức năng FR-01: Ghi nhận giao dịch đủ điều kiện ........................ Trang 3
   3.2 Chức năng FR-02: Tính toán & Phân bổ tiền Cashback ...................... Trang 3
   3.3 Chức năng FR-03: Đối soát biến động số dư tự động ....................... Trang 4
4. YÊU CẦU PHI CHỨC NĂNG (NON-FUNCTIONAL REQUIREMENTS) ......................... Trang 4
   4.1 Hiệu năng & Tốc độ xử lý ................................................ Trang 4
   4.2 Bảo mật & Khóa giao dịch nghi vấn ...................................... Trang 5
5. PHỤ LỤC & MA TRẬN TRUY XUẤT (TRACEABILITY MATRIX) .......................... Trang 5
--------------------------------------------------------------------------------
```

#### Chi tiết thiết kế Heading Styles trong Word:
*   **Title:** `TÀI LIỆU MÔ TẢ YÊU CẦU PHẦN MỀM (SRS) - TÍNH NĂNG HOÀN TIỀN VÍ PAYKX` (Font Calibri/Arial, 24pt, Bold, Centered).
*   **Heading 1:** Dùng cho các Mục lớn (`1. TỔNG QUAN...`, `2. YÊU CẦU...`, `3. YÊU CẦU CHỨC NĂNG...`).
    *   *Formatting:* Bold, Font 16pt, Primary Color (Navy Blue `#1B365D`), Space Before 12pt, Space After 6pt. Page Break Before cho mỗi Heading 1 chính để đảm bảo dàn trang đúng 5 trang.
*   **Heading 2:** Dùng cho các Sub-sections (`3.1 Chức năng FR-01...`, `3.2 Chức năng FR-02...`).
    *   *Formatting:* Bold, Font 13pt, Secondary Color (`#4A90E2`), Space Before 6pt, Space After 4pt.
*   **Body Text:** Font Calibri 11pt, Line spacing 1.15, Paragraph spacing After 6pt.

---

### 2.2. TÀI LIỆU 2 (EXCEL): BẢNG ĐỐI SOÁT VÀ XỬ LÝ DỮ LIỆU BẰNG VLOOKUP & TRIM

#### Sheet 1: `User_Master_DB` (Bảng tra cứu người dùng gốc - Bị lỗi khoảng trắng thừa)
| Row | Column A (User_ID) | Column B (User_Name) | Column C (University) |
| :--- | :--- | :--- | :--- |
| **1** | **User_ID** | **User_Name** | **University** |
| **2** | SV001 | `"  Nguyễn Văn An  "` | ĐH Bách Khoa |
| **3** | SV002 | `"Trần Thị Bích "` | ĐH Kinh Tế Quốc Dân |
| **4** | SV003 | `" Lê Hoàng Nam"` | ĐH Quốc Gia |
| **5** | SV004 | `"Phạm Đức Minh "` | ĐH Y Hà Nội |

#### Sheet 2: `Reconciliation_Report` (Bảng tổng hợp đối soát phí & giao dịch)

Công thức áp dụng tại cột **User_Name (Cột C)** nhằm vừa xóa sạch khoảng trắng thừa, vừa tra cứu chính xác tuyệt đối:

$$\text{Formula (Ô C2):} \quad \texttt{=TRIM(VLOOKUP(B2, User\_Master\_DB!\$A\$2:\$B\$5, 2, FALSE))}$$

**Chi tiết bảng tính:**

| Row | Col A: Trans_ID | Col B: User_ID | Col C: User_Name (Formula) | Col D: Amount | Col E: Fee_Bank | Col F: Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **Trans_ID** | **User_ID** | **User_Name** | **Amount (VND)** | **Fee_Bank** | **Status** |
| **2** | TXN1001 | SV001 | `=TRIM(VLOOKUP(B2, User_Master_DB!$A$2:$B$5, 2, FALSE))` $\rightarrow$ `"Nguyễn Văn An"` | 50,000 | 1,100 | SUCCESS |
| **3** | TXN1002 | SV003 | `=TRIM(VLOOKUP(B3, User_Master_DB!$A$2:$B$5, 2, FALSE))` $\rightarrow$ `"Lê Hoàng Nam"` | 100,000 | 1,100 | SUCCESS |
| **4** | TXN1003 | SV002 | `=TRIM(VLOOKUP(B4, User_Master_DB!$A$2:$B$5, 2, FALSE))` $\rightarrow$ `"Trần Thị Bích"` | 200,000 | 2,200 | SUCCESS |
| **5** | TXN1004 | SV004 | `=TRIM(VLOOKUP(B5, User_Master_DB!$A$2:$B$5, 2, FALSE))` $\rightarrow$ `"Phạm Đức Minh"` | 20,000 | 1,100 | SUCCESS |

*Lưu ý chuẩn Database:* Không gộp dòng Header, các ô từ Row 2 đến Row 5 liên tục không có dòng trống.

---

### 2.3. TÀI LIỆU 3 (POWERPOINT): SLIDE BÁO CÁO THEO QUY TẮC 10/20/30

#### Cấu trúc Slide Báo cáo Doanh thu & Dòng tiền:
*   **Tiêu đề Slide (Assertion Title):** `"Doanh thu T10 tăng 15% nhờ ưu đãi Cashback nạp điện thoại cho Sinh viên"`
    *   *Font size:* **32pt** (Bold, Màu tương phản cao), giải thích ngay kết luận quan trọng nhất thay vì viết tiêu đề chung chung như "Báo cáo doanh thu".
*   **Phần Nội dung Trực quan (Biểu đồ - Không dùng bảng số liệu thô):**
    *   Sử dụng **Biểu đồ Cột Chồng (Stacked Column Chart)** hoặc **Biểu đồ Đường (Line Chart)** thể hiện sự tăng trưởng dòng tiền qua các tuần trong tháng.
    *   Highlight màu cam/xanh lá vào cột mốc triển khai chiến dịch Cashback Nạp điện thoại để làm nổi bật nguyên nhân tăng trưởng.
*   **Phần Tóm tắt Key Takeaways (Bullet Points):**
    *   *Font size:* **30pt** (Tuân thủ nghiêm ngặt quy tắc 10/20/30: Phông chữ $\ge 30\text{pt}$).
    *   Thanh toán dịch vụ viễn thông đóng góp 45% tổng dòng tiền vào ví.
    *   Chi phí hoàn tiền (Cashback) giảm 3% tổng bù nợ nhờ tối ưu hóa phí ngân hàng đối tác.

---

## PHẦN 3. QUY TRÌNH KIỂM THỬ 4 ĐIỂM KHI GẶP LỖI #N/A TRONG VLOOKUP

Khi thực hiện hàm `=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)` mà trả về kết quả lỗi `#N/A` (Not Available), thực hiện quy trình kiểm tra 4 điểm theo thứ tự ưu tiên sau:

```
[Điểm 1: Khoảng trắng & Ký tự ẩn]
                 │
                 ▼
[Điểm 2: Định dạng Kiểu dữ liệu (Data Type)]
                 │
                 ▼
[Điểm 3: Khung vùng tra cứu (Table Array)]
                 │
                 ▼
[Điểm 4: Tính chính xác của Khóa tìm kiếm]
```

### Bước 1: Kiểm tra Khoảng trắng thừa và Ký tự không in được (Extra Spaces & Hidden Characters)
*   **Nguyên nhân:** Mã trong ô tìm kiếm (`lookup_value`) hoặc mã trong cột đầu tiên của bảng tra cứu (`table_array`) chứa khoảng trắng (ví dụ: `"SV001 "` khác với `"SV001"`).
*   **Cách khắc phục:** 
    *   Bọc hàm `TRIM` xung quanh giá trị tìm kiếm hoặc làm sạch dữ liệu nguồn bằng `=TRIM(CLEAN(cell))`.
    *   Sử dụng hàm `=LEN(cell)` để đếm độ dài chuỗi tại 2 ô và so sánh.

### Bước 2: Kiểm tra Xung đột Kiểu dữ liệu (Data Type Mismatch)
*   **Nguyên nhân:** Mã tra cứu ở một bên đang lưu dưới dạng **Số** (Number), nhưng ở bên kia lại lưu dưới dạng **Chuỗi/Văn bản** (Text). Excel sẽ coi `1001` (Number) và `"1001"` (Text) là hai giá trị hoàn toàn khác nhau.
*   **Cách khắc phục:**
    *   Chuyển đổi kiểu dữ liệu đồng nhất bằng `Text to Columns` trong thẻ Data.
    *   Hoặc ép kiểu trực tiếp trong công thức: Thêm `&""` để biến Number thành Text (`VLOOKUP(A2&"", ...)`), hoặc nhân với `1` (`A2*1`) để biến Text thành Number.

### Bước 3: Kiểm tra Cột tìm kiếm và Vùng tham chiếu (Table Array & First Column Rule)
*   **Nguyên nhân:**
    1. Vùng `table_array` không chứa cột giá trị tìm kiếm ở **Cột đầu tiên (Cột trái cùng)**.
    2. Quên khóa tọa độ tuyệt đối (dấu `$`), dẫn đến khi kéo công thức xuống dưới, vùng tra cứu bị trượt (`A2:B5` trượt thành `A5:B8`).
*   **Cách khắc phục:**
    *   Đảm bảo giá trị `lookup_value` nằm ở cột đầu tiên của `table_array`.
    *   Nhấn `F4` để cố định vùng tra cứu (ví dụ: `$A$2:$B$5`).

### Bước 4: Kiểm tra Sự tồn tại thực tế của dữ liệu khóa (Exact Match Existence)
*   **Nguyên nhân:** Giá trị `lookup_value` hoàn toàn không tồn tại trong danh mục nguồn (do nhập sai mã, chưa cập nhật dữ liệu sinh viên mới).
*   **Cách khắc phục:**
    *   Sử dụng công thức kết hợp hàm xử lý lỗi để giao diện hiển thị chuyên nghiệp hơn:
        $$\texttt{=IFERROR(TRIM(VLOOKUP(B2, User\_Master\_DB!\$A\$2:\$B\$5, 2, FALSE)), "Không tìm thấy NV")}$$
    *   Dùng chức năng `Ctrl + F` kiểm tra thủ công xem giá trị đó có trong bảng gốc hay không.