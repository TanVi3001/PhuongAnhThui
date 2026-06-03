# Hướng dẫn chạy script tạo báo cáo TLG

File Python dùng để tạo ra:

- `TLG_lam_do_da_dien_day_du.docx`
- `TLG_bang_phan_tich_markdown.md`

## 1. Cài Python

Kiểm tra máy đã có Python chưa:

```powershell
py --version
```

Nếu hiện phiên bản Python là được.

Ví dụ:

```text
Python 3.12.x
```

Nếu chưa có Python, tải tại trang chính thức của Python và nhớ tick **Add Python to PATH** khi cài.

## 2. Cài thư viện cần thiết

Mở PowerShell hoặc CMD, chạy:

```powershell
py -m pip install python-docx
```

## 3. Setup thư mục

Tạo một thư mục riêng, ví dụ:

```text
E:\PhuongAnh
```

Đặt các file vào cùng một thư mục như sau:

```text
E:\PhuongAnh
│
├── create_tlg_report_fixed.py
├── làmm dở.docx
```

Trong đó:

- `create_tlg_report_fixed.py`: file code Python đã fix.
- `làmm dở.docx`: file Word mẫu ban đầu.

## 4. Chạy code

Mở PowerShell tại thư mục chứa file:

```powershell
cd E:\PhuongAnh
```

Sau đó chạy:

```powershell
py create_tlg_report_fixed.py
```

## 5. Kết quả sau khi chạy

Nếu chạy thành công, chương trình sẽ tạo ra 2 file:

```text
TLG_lam_do_da_dien_day_du.docx
TLG_bang_phan_tich_markdown.md
```

Ý nghĩa:

- `TLG_lam_do_da_dien_day_du.docx`: file Word đã điền đầy đủ bảng.
- `TLG_bang_phan_tich_markdown.md`: bảng kết quả dạng Markdown.

## 6. Nếu file Word mẫu tên khác

Nếu file Word mẫu không tên là `làmm dở.docx`, chạy lệnh sau:

```powershell
py create_tlg_report_fixed.py --template "ten_file_word_mau.docx"
```

Ví dụ:

```powershell
py create_tlg_report_fixed.py --template "lam_do.docx"
```

## 7. Lỗi thường gặp

### Lỗi thiếu thư viện

Nếu gặp lỗi:

```text
ModuleNotFoundError: No module named 'docx'
```

Chạy lại:

```powershell
py -m pip install python-docx
```

### Lỗi không tìm thấy file Word

Nếu gặp lỗi kiểu:

```text
FileNotFoundError
```

Kiểm tra lại:

- File `.py` và file `.docx` đã nằm chung thư mục chưa.
- Tên file Word có đúng chưa.
- Nếu tên khác, dùng lệnh `--template` như mục 6.

## 8. Ghi chú

Không cần đặt file trong `/mnt/data`.

Đường dẫn `/mnt/data` chỉ dùng trong môi trường ChatGPT, không dùng trên máy Windows.
