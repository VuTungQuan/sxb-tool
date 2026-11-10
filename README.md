# Tool SXB – đối chiếu kho file với Excel tổng hợp sách xuất bản

Giao diện tkinter, toàn bộ xử lý trong `sxb_tool.py`. Có 2 bản giao diện dùng chung xử lý:
- **2 tab** (`sxb_tool.py`, `SXBTool.exe`): 1. Đối chiếu (quét kho → khớp tên → kiểm tra nội dung, một nút) | 2. Đổi tên & lưu | 3. Kiểm tra QĐXB
- **4 tab** (`sxb_tool_4tab.py`, `SXBTool_4tab.exe`): 1. Quét kho | 2. Đối chiếu Excel | 3. Kiểm tra nội dung | 4. Đổi tên & lưu | 5. Kiểm tra QĐXB

Yêu cầu Python 3.9+ (đã thử 3.11) + `openpyxl`, `pymupdf`, `python-docx`, `pywin32` (đọc .doc qua Microsoft Word).

```bash
pip install openpyxl pymupdf python-docx pywin32
python -X utf8 sxb_tool.py          # giao diện 2 tab
python -X utf8 sxb_tool_4tab.py     # giao diện 4 tab
python -X utf8 sxb_tool.py --test   # tự kiểm tra
```

## Chạy bằng file .exe (không cần Python)
Bấm đúp `build.bat`: tự tìm Python (biến `PYTHON` nếu đặt, rồi `py -3.11`, `py -3`, `python`), tự cài thư viện còn thiếu
(cả `pyinstaller`), chạy tự kiểm tra, rồi build `dist\SXBTool.exe` (2 tab) và `dist\SXBTool_4tab.exe` (4 tab), mỗi file ~37 MB,
copy đi máy khác chạy được ngay. Lần đầu mở mất vài giây vì exe tự giải nén. Nếu Windows Defender cảnh báo thì chọn
"Run anyway" (exe PyInstaller chưa ký số). Hai bản dùng chung thư mục `ketqua/` và cache khi để cạnh nhau.

## Chạy trên Ubuntu (Linux)
Build trên chính máy Ubuntu (PyInstaller không build chéo từ Windows sang Linux):
```bash
bash build.sh
```
`build.sh` tự cài gói còn thiếu (`python3-tk`, `python3-venv`, `antiword`, `catdoc`, `binutils`, `libpython3-dev` - cần sudo), tạo môi trường `.venv-linux`, cài thư viện, tự kiểm tra rồi tạo `dist-linux/SXBTool` (2 tab) và `dist-linux/SXBTool_4tab` (4 tab). Chạy: `./dist-linux/SXBTool` (cần `chmod +x` nếu chép từ Windows sang). File build trên Ubuntu 22.04 chạy được trên 22.04 và 24.04 (bản Ubuntu cũ hơn thì build lại trên chính máy đó). Không build cũng chạy được: `python3 sxb_tool.py` (hoặc `sxb_tool_4tab.py`).

Khác bản Windows: file `.doc` đọc bằng `antiword` (không có thì `catdoc`) thay cho Microsoft Word. Đường dẫn trong file kết quả là đường dẫn Linux (`/mnt/kho/...`), nên kết quả quét trên Linux dùng tiếp trên Linux.

## Thư mục kết quả
Tool tự tạo `ketqua/` cạnh file exe (có thể đổi trong từng tab):
- `1_quet_kho`: file kho `kho_*.xlsx` và cache kết quả đọc file `cache_noi_dung.sqlite` (dùng chung mọi lần chạy: file đã đọc không đọc lại)
- `2_doi_chieu`: file kết quả `doi_chieu_*.xlsx` và bản Excel tổng hợp đánh dấu (dòng KHỚP tô vàng)
- `3_doi_ten`: file kết quả sau khi đổi tên & lưu
- `4_kiem_tra_qdxb`: báo cáo `kiem_tra_qdxb_*.xlsx` và Excel đánh dấu thêm cột Kiểm tra QĐXB (cache đọc ở `1_quet_kho/cache_qdxb.sqlite`)
- `sach_da_luu`: thư mục đích mặc định khi đổi tên & lưu: mỗi sách một thư mục `Năm/thang MM/<tên mới>/<tên mới>.pdf|.docx`

## Quy trình
Nút **■ Dừng** (cuối cửa sổ) dừng bước đang chạy sau file hiện tại; vẫn ghi kết quả phần đã làm, chạy lại sẽ làm tiếp.

1. **Đối chiếu** (một nút, chạy liền các bước, có thanh tiến độ và thời gian còn lại):
   - *Quét kho → Excel*: STT, tên file, định dạng, ngày sửa, năm, kích thước, đường dẫn. Quá 1.048.575 dòng thì tự tách sheet Kho 2, Kho 3. Thư mục/file không đọc được liệt kê vào `kho_*.xlsx.loi.txt`; cảnh báo khi ảnh JPG chiếm quá nửa. Chọn ô **File kho có sẵn** để bỏ qua bước này (quét kho lớn rất lâu).
   - *Khớp tên sách*: tên sách (bỏ dấu) với tên file hoặc tên thư mục chứa file, lọc theo năm Date Modified → `KHỚP` / `LỆCH NĂM` / `KHÔNG TÌM THẤY`.
   - *Năm xuất bản (LỆCH NĂM)*: mở PDF/DOCX/DOC (.doc qua Microsoft Word nếu máy có), lấy năm xuất bản ở trang cuối, không khớp thì ở 8 trang đầu. File .doc mở bằng Word quá 90 giây thì Word bị tắt, file đó bỏ qua.
   - *Lời nói đầu (KHÔNG TÌM THẤY)*: tìm câu "cuốn sách <tên sách>" trong lời nói đầu của các file còn lại trong kho.
   - Xuất `doi_chieu_*.xlsx`: sheet `KHOP`, `CAN_XU_LY`, `THU_MUC_KHO` (từng thư mục kho khớp sách nào / chỉ trùng tên / không khớp sách nào) và `CAN_OCR` (PDF không có lớp chữ - ảnh scan, chữ dạng hình vẽ - cần OCR mới đọc được; có cột ưu tiên). Bản sao Excel tổng hợp thêm 4 cột tình trạng, dòng KHỚP tô vàng.
2. **Đổi tên & lưu**: copy/move file Word, PDF sang `Năm/thang MM/<tên mới>/`, mỗi sách một thư mục, theo tên mới `ten_sach_mmyy_R_BBT_T01` (phần tên sách tối đa 90 ký tự, hằng `DAI_TEN`). Phần `BBT` lấy theo cột "Mã quy định" (cột F) của `BBT_distinct.xlsx` (đặt cạnh Excel tổng hợp hoặc cạnh exe), khớp đúng giá trị cột J; dòng chưa điền mã thì dùng BBT bỏ dấu. File này được đọc lại mỗi lần đối chiếu, sửa mã xong chạy lại đối chiếu là có tên mới. File bìa cùng thư mục với file ruột - ảnh (.jpg, .png, .tif), file thiết kế (.ai, .psd, .eps), PDF 1-2 trang, PDF ≤ 8 trang có trang đầu nằm ngang (trải bìa + tai gập) hoặc PDF tên 'Bìa / Tờ che / Tr 2-4' - được chép vào `<tên mới>/Bia/`, tên đổi `_R_` (Ruột) thành `_B_` (Bìa), nhiều bìa cùng đuôi thì thêm `_2`, `_3`. PDF 1-2 trang không bao giờ được chọn làm ruột. Chạy lại được: file đã có thì bỏ qua; file lỗi ghi vào Ghi chú rồi làm tiếp; kiểm tra ổ đích đủ chỗ trước khi chép.
3. **Kiểm tra QĐXB**: duyệt thư mục đích đã chép ở bước 2 (bỏ thư mục `Bia`), đọc "Quyết định xuất bản số ..., ngày 04-9-2013" và "nộp lưu chiểu tháng 9-2013" ở 3 trang cuối từng sách (không có thì ở 8 trang đầu), so với cột QĐXB của Excel: tháng + năm nếu Excel ghi `mm/yyyy`, cả ngày nếu `dd/mm/yyyy`. Mỗi thư mục gắn với dòng Excel theo cột "Đường dẫn đích", không có thì theo năm + tên thư mục. Kết quả `ĐÚNG` / `SAI - Excel 01/2013, file 04/09/2013` (tô cam) / `KHÔNG ĐỌC ĐƯỢC QĐXB` (tô xám) / `KHÔNG THẤY Ở THƯ MỤC ĐÍCH`, kèm số QĐXB, tháng lưu chiểu, **tên mới theo ngày QĐXB trên file** và ghi chú khi hai dòng Excel ra cùng một QĐXB (thường do chép trùng một file). Chỉ đọc, không đổi tên hay di chuyển file. Chạy lại trên bản Excel đã đánh dấu thì ghi đè đúng các cột cũ.
   - Lưu ý năm 2013: Excel gốc không có cột QĐXB, cột ở vị trí đó là "Thời gian có sách" (bản chuẩn hoá gộp vào cột QĐXB), nên năm 2013 sẽ có nhiều dòng SAI dù sách đúng.

## Dòng LỆCH NĂM sau khi đọc nội dung
Dòng không thấy file đúng năm được tự phân loại (sheet `CAN_XU_LY`, dòng cần xem xếp lên đầu, có bộ lọc):

| Tình trạng | Ý nghĩa | Cần làm |
|---|---|---|
| `KHỚP THEO THƯ MỤC NĂM - nên xem` | Có file trong thư mục năm đúng (`SACH NAM 2012`, `Sach 2014`), đã điền sẵn file gợi ý | Xem nhanh; đúng thì sửa ô Tình trạng thành `KHỚP` rồi chạy Đổi tên & lưu |
| `LỆCH NĂM - lệch 1 năm, xem trước` | Bản thảo sửa năm trước hoặc lưu chiểu năm sau | Mở file kiểm tra |
| `LỆCH NĂM - thư mục rút gọn, cần xem` | Chỉ khớp theo tên thư mục ngắn, dễ nhầm | Mở file kiểm tra |
| `LỆCH NĂM - cần mở file kiểm tra` | Không đọc ra năm (thường là bản thảo Word không có trang lưu chiểu) | Mở file kiểm tra |
| `KHÔNG CÓ BẢN ĐÚNG NĂM` | Mọi file thuộc sách năm khác theo Date Modified, hoặc file ghi năm khác xa | Không cần mở; lý do ở cột Ghi chú |
| `KHÔNG CÓ FILE RUỘT` | Chỉ có ảnh / file thiết kế | Không cần mở |

File `~$...` (file tạm của Word) và `._...` (file phụ của macOS) không được tính là sách. Khi chép kho sang ổ khác phải giữ Date Modified (`cp -a` hoặc `rsync -a`), mất ngày gốc thì số KHỚP giảm rõ rệt.

## Chạy đa nhân
Bước kiểm tra nội dung đọc trước các file PDF/DOCX song song bằng nhiều tiến trình (mặc định = một nửa số luồng CPU, ~ số nhân thật: Xeon 16 nhân/32 luồng dùng 16 tiến trình; hằng `SO_LUONG`). Đo trên ~400 file: 26 giây đọc lần lượt còn 3,5 giây với 8 tiến trình; ổ HDD/ổ mạng chậm thì tăng ít hơn. File .doc (đọc qua Word) vẫn đọc lần lượt. Quét kho và chép file phụ thuộc tốc độ ổ đĩa nên không chạy song song. Một file PDF hỏng làm sập tiến trình đọc chỉ làm chết tiến trình con, tool tự tìm đúng file đó, ghi vào nhat_ky_loi.txt và đọc tiếp.

## Khi tool tắt đột ngột (chạy dữ liệu lớn)
Chỉ cần mở lại tool và chạy lại, phần đã làm không phải làm lại:
- **Quét kho** ghi dần vào `ketqua/1_quet_kho/quet_dang_do.csv` (mỗi thư mục xong là lưu). Chạy lại, tool hỏi "quét tiếp phần còn lại?": chọn Có thì bỏ qua các thư mục đã quét. Bấm ■ Dừng khi đang quét cũng giữ được phần đã quét.
- **Đọc nội dung** lưu cache ngay sau từng file. Trước khi đọc mỗi file tool ghi dấu "đang đọc": nếu tool tắt khi đang đọc một file (thường do file PDF hỏng làm thư viện đọc PDF sập), lần chạy sau tự nhận ra file đó, ghi vào log và `ketqua/nhat_ky_loi.txt`, coi là không đọc được và bỏ qua - không sập lại đúng chỗ cũ.
- **Đổi tên & lưu**: file đã chép thì bỏ qua, chép tiếp phần còn lại.
- `ketqua/nhat_ky_loi.txt` ghi mọi lỗi, kể cả lỗi sập ở tầng C, để biết tool tắt ở đâu.

## Chuẩn hoá Excel
`chuan_hoa_excel.py` gom các sheet năm (header lệch, tên cột khác nhau) về một bộ cột cố định:

```bash
python -X utf8 chuan_hoa_excel.py "Tong hop SXB 2006-2025.xlsx"
```
