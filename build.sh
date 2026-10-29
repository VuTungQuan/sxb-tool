#!/usr/bin/env bash
# Build Tool SXB cho Ubuntu: dist-linux/SXBTool (bản 2 tab) và dist-linux/SXBTool_4tab (bản 4 tab).
# Tự cài gói hệ thống còn thiếu (cần sudo), tạo môi trường Python riêng (.venv-linux), cài thư viện, tự kiểm tra rồi build.
# Chạy: bash build.sh      (muốn dùng Python cụ thể: PYTHON=/usr/bin/python3.11 bash build.sh)
set -euo pipefail
cd "$(dirname "$0")"
echo "=== Build Tool SXB (Linux): SXBTool (2 tab) và SXBTool_4tab (4 tab) ==="

# 1. Gói hệ thống: python3 + venv + tkinter (giao diện), antiword/catdoc (đọc .doc - Linux không có Word),
#    binutils + libpython3-dev (PyInstaller cần objdump và libpython3.x.so)
can=()
for g in python3 python3-venv python3-tk antiword catdoc binutils libpython3-dev; do
    dpkg -s "$g" >/dev/null 2>&1 || can+=("$g")
done
if ((${#can[@]})); then
    echo "Cài gói còn thiếu: ${can[*]}"
    SUDO=""; [ "$(id -u)" -ne 0 ] && SUDO="sudo"
    $SUDO apt-get update -qq && $SUDO apt-get install -y -qq "${can[@]}"
fi

PY="${PYTHON:-python3}"
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' || { echo "[LỖI] Cần Python 3.9 trở lên"; exit 1; }
"$PY" -c 'import tkinter' || { echo "[LỖI] Python này thiếu tkinter: sudo apt install python3-tk"; exit 1; }

# 2. Môi trường riêng + thư viện (pymupdf ghim dưới 2.0 như bản Windows)
[ -x .venv-linux/bin/python ] || "$PY" -m venv .venv-linux
V=.venv-linux/bin/python
"$V" -m pip install -q --disable-pip-version-check --upgrade pip
"$V" -m pip install -q --disable-pip-version-check openpyxl "pymupdf>=1.24,<2" python-docx pyinstaller
"$V" -c 'import sys, fitz; print("Python", sys.version.split()[0], "- PyMuPDF", fitz.VersionBind)'

# 3. Tự kiểm tra trước khi build
echo "Tự kiểm tra (sxb_tool.py --test)..."
"$V" -X utf8 sxb_tool.py --test

# 4. Build
for x in "SXBTool sxb_tool.py" "SXBTool_4tab sxb_tool_4tab.py"; do
    set -- $x
    echo "Build dist-linux/$1..."
    "$V" -m PyInstaller --noconfirm --onefile --windowed --log-level WARN --name "$1" \
        --distpath dist-linux --workpath build-linux/work --specpath build-linux "$2"
done

# 5. Kiểm tra lại chính file vừa build (--test không mở giao diện)
for e in SXBTool SXBTool_4tab; do
    "dist-linux/$e" --test; echo "  dist-linux/$e --test: OK"
done
echo
echo "=== XONG ==="
echo "  dist-linux/SXBTool      : bản 2 tab (Đối chiếu | Đổi tên & lưu)"
echo "  dist-linux/SXBTool_4tab : bản 4 tab (Quét kho | Đối chiếu Excel | Kiểm tra nội dung | Đổi tên & lưu)"
echo "  Chạy: ./dist-linux/SXBTool  (đặt 'Tong hop SXB 2006-2025_chuan.xlsx' và BBT_distinct.xlsx cạnh file chạy hoặc thư mục cha)"
