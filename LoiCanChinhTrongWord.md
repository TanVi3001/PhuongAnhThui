README - Hướng dẫn rà soát và sửa lỗi số liệu BCTC Tập đoàn Thiên Long (TLG)
📌 Mục đích
Tài liệu này liệt kê các điểm sai sót (do lỗi ảo giác/nhận diện sai số của AI) trong quá trình trích xuất dữ liệu từ 3 file PDF Báo cáo tài chính (2023, 2024, 2025) vào file Word (làm dở.docx).

Bạn cần mở trực tiếp các file PDF gốc, đối chiếu lại đúng các dòng được liệt kê dưới đây để cập nhật con số chính xác trước khi hoàn thiện báo cáo.

⚠️ Các lỗi logic số liệu cần khắc phục (Bắt buộc đối chiếu file gốc)
1. Lỗi Hàng tồn kho năm 2023 (Lệch 200.000.000 VND)
Số liệu trong bảng hiện tại: 831.899.830.043

Công thức kiểm tra: Hàng tồn kho = Giá gốc - Dự phòng = 855.426.147.198 - 23.726.317.155 = 831.699.830.043

Hành động: Mở BCTC 2023, Bảng Cân đối kế toán, tìm mục "Hàng tồn kho" và xác nhận lại con số tổng. AI có thể đã nhìn nhầm số 6 thành số 8.

2. Lỗi Tổng Nợ ngắn hạn năm 2023 (Lệch khoảng 20.000.000 VND)
Số liệu trong bảng hiện tại: 659.537.277.010

Công thức kiểm tra: Tổng các khoản mục con (Phải trả người bán + Người mua trả trước + Thuế + Phải trả LĐ + Chi phí phải trả + Phải trả khác + Vay NH + Quỹ KT) cộng lại ra 659.557.276.930.

Hành động: Dò lại toàn bộ các mục con của "Nợ ngắn hạn" năm 2023. Có thể AI đã gõ sai một con số hoặc file Word đang liệt kê thiếu 1 khoản mục con nào đó có trong báo cáo gốc.

3. Lỗi Lợi nhuận sau thuế TNDN năm 2023 (Lệch 2.000.000 VND)
Số liệu trong bảng hiện tại: 356.174.272.581

Công thức kiểm tra: LNST = LN trước thuế - Thuế hiện hành - Thuế hoãn lại = 451.960.655.914 - 93.999.785.980 - 1.788.597.353 = 356.172.272.581.

Hành động: Mở Báo cáo Kết quả Kinh doanh 2023, kiểm tra lại 3 con số trên để xem sai ở đâu (khả năng gõ nhầm số 2 thành số 4 ở kết quả tổng).

4. Lỗi Lợi nhuận thuần từ hoạt động kinh doanh năm 2024 (Lệch 20.000.000 VND)
Số liệu trong bảng hiện tại: 580.025.969.561

Công thức kiểm tra: LN gộp + Doanh thu TC - Chi phí TC + Lãi liên kết - Chi phí bán hàng - Chi phí QLDN = 580.005.969.561.

Hành động: Mở Báo cáo Kết quả Kinh doanh 2024. Kiểm tra lại Chi phí bán hàng và Chi phí QLDN, AI rất có thể đã lấy lệch số ở 2 mục này.

🛠 Hướng dẫn xử lý sau khi sửa số
Cập nhật số thô: Chỉnh sửa trực tiếp số VND bị sai vào các ô tương ứng trong file làm dở.docx.

Tính toán lại các cột so sánh: Với những hàng (row) bị thay đổi số liệu gốc, bạn bắt buộc phải dùng máy tính bấm lại 2 cột so sánh ở cuối bảng:

Cột Chênh lệch Số tiền VND = Số năm sau - Số năm trước.

Cột Phần trăm (%) = (Chênh lệch / Số năm trước) * 100.

Lưu ý về "Vốn chủ sở hữu": Hàng Tổng "Vốn chủ sở hữu" không cộng khớp với các hàng con bên dưới là do cấu trúc file Word đang không liệt kê đủ tất cả các dòng (ví dụ: thiếu Cổ phiếu quỹ, Quỹ khác...). Chỉ cần giữ nguyên số Tổng y hệt như BCTC gốc in ra, không cần cố cộng ép các dòng con cho bằng số tổng.