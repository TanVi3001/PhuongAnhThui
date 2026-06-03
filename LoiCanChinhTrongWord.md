# README - Hướng dẫn rà soát và sửa lỗi số liệu BCTC Tập đoàn Thiên Long (TLG)

## 1. Mục đích

Tài liệu này dùng để rà soát lại các điểm có khả năng sai lệch trong quá trình trích xuất dữ liệu từ 3 file PDF Báo cáo tài chính hợp nhất kiểm toán của Tập đoàn Thiên Long:

- BCTC hợp nhất kiểm toán năm 2023
- BCTC hợp nhất kiểm toán năm 2024
- BCTC hợp nhất kiểm toán năm 2025

Các lỗi bên dưới có thể xuất phát từ việc AI nhận diện sai số liệu trong PDF hoặc nhập sai số vào file Word `làm dở.docx`.

Trước khi hoàn thiện báo cáo, cần mở trực tiếp file PDF gốc và đối chiếu lại đúng từng dòng được liệt kê.

---

## 2. Các lỗi logic số liệu cần kiểm tra

### 2.1. Hàng tồn kho năm 2023

**Số liệu hiện tại trong bảng:**

```text
831.899.830.043
```

**Công thức kiểm tra:**

```text
Hàng tồn kho = Giá gốc - Dự phòng giảm giá hàng tồn kho
```

Thay số:

```text
855.426.147.198 - 23.726.317.155 = 831.699.830.043
```

**Vấn đề phát hiện:**

Số liệu hiện tại đang lệch:

```text
200.000.000 VND
```

**Cần xử lý:**

Mở BCTC 2023, vào **Bảng cân đối kế toán hợp nhất**, tìm mục **Hàng tồn kho** và xác nhận lại con số tổng.

Có khả năng AI đã nhận diện nhầm số **6** thành số **8**.

---

### 2.2. Tổng Nợ ngắn hạn năm 2023

**Số liệu hiện tại trong bảng:**

```text
659.537.277.010
```

**Công thức kiểm tra:**

Cộng các khoản mục con của **Nợ ngắn hạn**, gồm:

```text
Phải trả người bán ngắn hạn
+ Người mua trả tiền trước ngắn hạn
+ Thuế và các khoản phải nộp Nhà nước
+ Phải trả người lao động
+ Chi phí phải trả ngắn hạn
+ Phải trả ngắn hạn khác
+ Vay ngắn hạn
+ Quỹ khen thưởng, phúc lợi
```

Kết quả cộng lại đang ra:

```text
659.557.276.930
```

**Vấn đề phát hiện:**

Số liệu hiện tại đang lệch khoảng:

```text
20.000.000 VND
```

**Cần xử lý:**

Mở BCTC 2023, vào **Bảng cân đối kế toán hợp nhất**, dò lại toàn bộ các khoản mục con thuộc **Nợ ngắn hạn**.

Có thể xảy ra một trong hai trường hợp:

- AI đã nhập sai một khoản mục con.
- Bảng Word đang thiếu một khoản mục con có trong báo cáo gốc.

---

### 2.3. Lợi nhuận sau thuế TNDN năm 2023

**Số liệu hiện tại trong bảng:**

```text
356.174.272.581
```

**Công thức kiểm tra:**

```text
Lợi nhuận sau thuế TNDN
= Lợi nhuận trước thuế
- Chi phí thuế TNDN hiện hành
- Chi phí thuế TNDN hoãn lại
```

Thay số:

```text
451.960.655.914 - 93.999.785.980 - 1.788.597.353
= 356.172.272.581
```

**Vấn đề phát hiện:**

Số liệu hiện tại đang lệch:

```text
2.000.000 VND
```

**Cần xử lý:**

Mở BCTC 2023, vào **Báo cáo kết quả hoạt động kinh doanh hợp nhất**, kiểm tra lại 3 dòng:

```text
Tổng lợi nhuận kế toán trước thuế
Chi phí thuế TNDN hiện hành
Chi phí thuế TNDN hoãn lại
```

Sau đó xác nhận lại dòng:

```text
Lợi nhuận sau thuế TNDN
```

Khả năng cao là kết quả đã bị gõ nhầm số **2** thành số **4**.

---

### 2.4. Lợi nhuận thuần từ hoạt động kinh doanh năm 2024

**Số liệu hiện tại trong bảng:**

```text
580.025.969.561
```

**Công thức kiểm tra:**

```text
Lợi nhuận thuần từ hoạt động kinh doanh
= Lợi nhuận gộp
+ Doanh thu hoạt động tài chính
- Chi phí tài chính
+ Phần lãi/lỗ trong công ty liên kết
- Chi phí bán hàng
- Chi phí quản lý doanh nghiệp
```

Kết quả kiểm tra đang ra:

```text
580.005.969.561
```

**Vấn đề phát hiện:**

Số liệu hiện tại đang lệch:

```text
20.000.000 VND
```

**Cần xử lý:**

Mở BCTC 2024, vào **Báo cáo kết quả hoạt động kinh doanh hợp nhất**, kiểm tra lại các dòng:

```text
Lợi nhuận gộp
Doanh thu hoạt động tài chính
Chi phí tài chính
Phần lãi/lỗ trong công ty liên kết
Chi phí bán hàng
Chi phí quản lý doanh nghiệp
Lợi nhuận thuần từ hoạt động kinh doanh
```

AI có thể đã lấy lệch số ở dòng **Chi phí bán hàng** hoặc **Chi phí quản lý doanh nghiệp**.

---

## 3. Hướng dẫn xử lý sau khi sửa số

### Bước 1. Cập nhật số liệu gốc

Mở file Word đang làm và chỉnh trực tiếp số VND bị sai vào đúng ô tương ứng.

Các dòng cần ưu tiên rà soát:

```text
Hàng tồn kho năm 2023
Nợ ngắn hạn năm 2023
Lợi nhuận sau thuế TNDN năm 2023
Lợi nhuận thuần từ hoạt động kinh doanh năm 2024
```

---

### Bước 2. Tính lại cột so sánh

Sau khi sửa số liệu gốc, bắt buộc tính lại các cột so sánh của chính dòng đó.

#### So sánh năm 2024 với năm 2023

```text
Chênh lệch số tiền = Số liệu năm 2024 - Số liệu năm 2023
```

```text
Phần trăm (%) = Chênh lệch số tiền / Số liệu năm 2023 * 100
```

#### So sánh năm 2025 với năm 2024

```text
Chênh lệch số tiền = Số liệu năm 2025 - Số liệu năm 2024
```

```text
Phần trăm (%) = Chênh lệch số tiền / Số liệu năm 2024 * 100
```

---

## 4. Lưu ý về dòng Vốn chủ sở hữu

Dòng tổng **Vốn chủ sở hữu** có thể không cộng khớp với các dòng con đang hiển thị trong file Word.

Nguyên nhân là do bảng Word có thể chưa liệt kê đầy đủ toàn bộ các khoản mục chi tiết, ví dụ:

```text
Cổ phiếu quỹ
Quỹ khác
Chênh lệch tỷ giá
Lợi ích cổ đông không kiểm soát
Các khoản mục điều chỉnh khác
```

Vì vậy, không cần cố cộng ép các dòng con cho bằng dòng tổng.

Cách xử lý đúng:

```text
Giữ nguyên số tổng Vốn chủ sở hữu đúng như BCTC gốc.
```

---

## 5. Nguyên tắc rà soát bắt buộc

Khi sửa số liệu, cần tuân thủ các nguyên tắc sau:

1. Không tự đoán số liệu.
2. Không lấy số liệu từ Internet.
3. Chỉ dùng số liệu trong 3 file PDF BCTC gốc.
4. Nếu số trong Word khác số trong PDF, ưu tiên số trong PDF gốc.
5. Sau khi sửa số liệu gốc, phải tính lại toàn bộ cột chênh lệch và phần trăm liên quan.
6. Với số âm, trình bày theo dạng kế toán nếu cần, ví dụ:

```text
(23.726.317.155)
```

---

## 6. Checklist trước khi nộp báo cáo

Trước khi hoàn thiện file Word, cần kiểm tra:

- [ ] Đã đối chiếu lại Hàng tồn kho năm 2023.
- [ ] Đã đối chiếu lại Nợ ngắn hạn năm 2023.
- [ ] Đã đối chiếu lại Lợi nhuận sau thuế TNDN năm 2023.
- [ ] Đã đối chiếu lại Lợi nhuận thuần từ hoạt động kinh doanh năm 2024.
- [ ] Đã tính lại chênh lệch 2024 so với 2023 cho các dòng bị sửa.
- [ ] Đã tính lại chênh lệch 2025 so với 2024 cho các dòng bị sửa.
- [ ] Đã kiểm tra định dạng dấu phân cách hàng nghìn.
- [ ] Đã lưu lại file Word bản cuối cùng.

---

## 7. Gợi ý đặt tên file sau khi sửa

Nên lưu bản cuối cùng với tên rõ ràng:

```text
TLG_Bang_phan_tich_BCTC_2023_2025_FINAL.docx
```

Nếu muốn giữ bản nháp cũ, có thể lưu thêm:

```text
TLG_Bang_phan_tich_BCTC_2023_2025_DRAFT.docx
```
