# Tool SXB – đối chiếu kho file với Excel tổng hợp sách xuất bản

Giao diện tkinter, 1 file `sxb_tool.py`. Yêu cầu Python 3.11 + `openpyxl`, `pymupdf`, `python-docx`.

```bash
pip install openpyxl pymupdf python-docx
python -X utf8 sxb_tool.py          # giao diện
python -X utf8 sxb_tool.py --test   # tự kiểm tra
```

## Quy trình
1. **Quét kho → CSV**: tên file, định dạng, date modified, path.
2. **Đối chiếu Excel**: khớp tên sách (bỏ dấu) với tên file, lọc theo năm Date Modified → `KHỚP` / `LỆCH NĂM` / `KHÔNG TÌM THẤY`. Xuất `doi_chieu_*.xlsx` (2 sheet) và bản Excel gốc có thêm cột tình trạng.
3. **Kiểm tra nội dung**: mở PDF/DOCX đọc năm ở trang cuối, dò "Lời nói đầu" trong 8 trang đầu. Text trích được cache.
4. **Đổi tên & lưu**: copy/move file sang `Năm/thang MM` theo tên mới `ten_sach_mmyy_R_BBT_T01`.

## Chuẩn hoá Excel
`chuan_hoa_excel.py` gom các sheet năm (header lệch, tên cột khác nhau) về một bộ cột cố định:

```bash
python -X utf8 chuan_hoa_excel.py "Tong hop SXB 2006-2025.xlsx"
```
