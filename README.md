# Tool SXB – đối chiếu kho file với Excel tổng hợp sách xuất bản

Giao diện tkinter, 1 file `sxb_tool.py`. Yêu cầu Python 3.11 + `openpyxl`, `pymupdf`, `python-docx`.

```bash
pip install openpyxl pymupdf python-docx
python -X utf8 sxb_tool.py          # giao diện
python -X utf8 sxb_tool.py --test   # tự kiểm tra
```

## Chạy bằng file .exe (không cần Python)
Chạy `build.bat` (cần `pip install pyinstaller`) → `dist\SXBTool.exe`, 1 file ~55 MB, copy đi máy khác chạy được ngay.
Lần đầu mở mất vài giây vì exe tự giải nén. Nếu Windows Defender cảnh báo thì chọn "Run anyway" (exe PyInstaller chưa ký số).

## Thư mục kết quả
Tool tự tạo `ketqua/` cạnh file exe, mỗi bước một thư mục (có thể đổi trong từng tab):
- `1_quet_kho`: file kho `kho_*.xlsx` và cache kết quả đọc file `*.cache.sqlite` (chạy lại tab 3 không đọc lại file cũ)
- `2_doi_chieu`, `3_kiem_tra`, `4_doi_ten`: file `doi_chieu_*.xlsx` và bản Excel đánh dấu của từng bước
- `sach_da_luu`: thư mục đích mặc định khi đổi tên & lưu: mỗi sách một thư mục `Năm/thang MM/<tên mới>/<tên mới>.pdf|.docx`

## Quy trình
Nút **■ Dừng** (cuối cửa sổ) dừng bước đang chạy sau file hiện tại; tab 3, 4 vẫn ghi kết quả phần đã làm, chạy lại sẽ làm tiếp.

1. **Quét kho → Excel**: STT, tên file, định dạng, ngày sửa, năm, kích thước, đường dẫn. Header đậm, cố định dòng 1, wrap text. Quá 1.048.575 dòng thì tự tách sheet Kho 2, Kho 3. Thư mục/file không đọc được được liệt kê vào `kho_*.xlsx.loi.txt`; cảnh báo khi ảnh JPG chiếm quá nửa.
2. **Đối chiếu Excel**: khớp tên sách (bỏ dấu) với tên file, lọc theo năm Date Modified → `KHỚP` / `LỆCH NĂM` / `KHÔNG TÌM THẤY`. Xuất `doi_chieu_*.xlsx` (sheet `KHOP`, `CAN_XU_LY`, và `THU_MUC_KHO`: từng thư mục trong kho khớp sách nào / chỉ trùng tên / không khớp sách nào) và bản Excel gốc có thêm cột tình trạng.
3. **Kiểm tra nội dung**: mở PDF/DOCX/DOC (.doc đọc qua Microsoft Word nếu máy có) lấy năm xuất bản ở trang cuối, không khớp thì ở 8 trang đầu, dò "Lời nói đầu" trong 8 trang đầu. Kết quả đọc được lưu cache SQLite. File .doc mở bằng Word quá 90 giây thì Word bị tắt, file đó bỏ qua.
4. **Đổi tên & lưu**: copy/move file Word, PDF sang `Năm/thang MM/<tên mới>/`, mỗi sách một thư mục, theo tên mới `ten_sach_mmyy_R_BBT_T01` (phần tên sách tối đa 90 ký tự, hằng `DAI_TEN`). Chạy lại được: file đã có thì bỏ qua; file lỗi ghi vào Ghi chú rồi làm tiếp; kiểm tra ổ đích đủ chỗ trước khi chép.

## Chuẩn hoá Excel
`chuan_hoa_excel.py` gom các sheet năm (header lệch, tên cột khác nhau) về một bộ cột cố định:

```bash
python -X utf8 chuan_hoa_excel.py "Tong hop SXB 2006-2025.xlsx"
```
