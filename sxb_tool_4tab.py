"""Bản 4 tab của Tool SXB: 1. Quét kho | 2. Đối chiếu Excel | 3. Kiểm tra nội dung | 4. Đổi tên & lưu.
Dùng chung toàn bộ xử lý với sxb_tool.py (bản 2 tab). Build cả hai bản: build.bat"""
import sys
import sxb_tool

if __name__ == "__main__":
    import multiprocessing; multiprocessing.freeze_support()  # exe: tiến trình con đọc PDF song song
    sxb_tool._test() if "--test" in sys.argv else sxb_tool.gui(4)
