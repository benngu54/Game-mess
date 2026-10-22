# Nạn buôn người ở Hà Nội — bài thuyết trình

Bài thuyết trình 15 slide (16:9, khoảng 10 phút) dành cho học sinh THPT.

| File | Nội dung |
|---|---|
| `Nan_buon_nguoi_o_Ha_Noi.pptx` | File PowerPoint chính: ảnh nhúng sẵn, có transition, animation và Speaker Notes |
| `Loi_thuyet_trinh.docx` | Lời thuyết trình theo từng slide |
| `Nguon_anh.txt` | Nguồn, tác giả và giấy phép của từng ảnh |
| `Tai_lieu_tham_khao.txt` | Văn bản pháp luật, số liệu, báo chí kèm ghi chú về giới hạn dữ liệu |
| `preview.pdf` | Bản xem trước (render bằng LibreOffice) |
| `source_code.py` | Mã Python tạo toàn bộ file trên |
| `assets/photos/` | Ảnh gốc (Open Images / Flickr, CC BY 2.0) |
| `assets/fonts/` | Font Be Vietnam Pro (SIL OFL) |

## Font

Tiêu đề dùng **Be Vietnam Pro**, nội dung dùng **Arial**. Trước khi mở file, hãy cài các file
`.ttf` trong `assets/fonts/` (Windows: chuột phải → *Install for all users*; macOS: mở bằng Font Book).
Nếu máy chưa có Be Vietnam Pro, PowerPoint sẽ thay bằng font khác: chữ vẫn đúng dấu nhưng tiêu đề kém đẹp hơn.
Khi trình chiếu trên máy khác, nên bật *File → Options → Save → Embed fonts in the file* rồi lưu lại.

## Tạo lại file

```bash
pip install python-pptx pillow lxml python-docx
python source_code.py
```

Hiệu ứng **Morph** ở slide 12 cần PowerPoint 2019 hoặc Microsoft 365. Các bản cũ hơn sẽ tự dùng Fade.
