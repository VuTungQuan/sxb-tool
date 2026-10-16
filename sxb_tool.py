"""Tool SXB: quét kho file -> CSV, đối chiếu với 'Tong hop SXB 2006-2025.xlsx', phân loại,
kiểm tra nội dung file (năm trang cuối, lời nói đầu), đổi tên & lưu theo Năm/thang MM.

Chạy:  python -X utf8 sxb_tool.py            (giao diện)
       python -X utf8 sxb_tool.py --test     (tự kiểm tra)
"""
import csv, json, os, re, shutil, sys, threading, unicodedata, datetime as dt
from collections import defaultdict

EXT_MAC_DINH = ".doc .docx .pdf"
WORD = {"doc", "docx"}
DOC_DUOC = {"pdf", "docx"}  # định dạng trích được text (.doc bỏ qua)


# ---------- chuẩn hoá tên ----------
def khong_dau(s):
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def chuan_hoa(s):
    """'Lịch sử Chính phủ- T1' -> 'lich su chinh phu t1'"""
    return re.sub(r"[^a-z0-9]+", " ", khong_dau(s).lower()).strip()


def ma_ngay(v):
    """QĐXB -> ddmmyy hoặc mmyy. datetime, '01-2015', '1/2025', '19/1/2007', '11-2-18'."""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d%m%y")
    nums = re.findall(r"\d+", str(v or ""))
    if len(nums) == 3:
        d, m, y = nums
        return f"{int(d):02d}{int(m):02d}{y[-2:]}"
    if len(nums) == 2:
        m, y = nums
        return f"{int(m):02d}{y[-2:]}"
    return ""


def ma_lan_xb(v):
    """None/1 -> M00 ; TB -> T01 ; số n>1 hoặc 'Lần 3'/'TB lần 2' -> T0n."""
    if v is None or str(v).strip() in ("", "1"):
        return "M00"
    s = str(v).strip().upper()
    n = re.findall(r"\d+", s)
    if n and int(n[0]) > 1:
        return f"T{int(n[0]):02d}"
    if s.startswith(("TB", "NB", "NỐI", "NOI", "LẦN", "LAN")):
        return "T01"
    return "M00"


def ten_file_moi(ten_sach, qdxb, bbt, lan_xb):
    parts = ["_".join(chuan_hoa(ten_sach).split())]
    if ma_ngay(qdxb):
        parts.append(ma_ngay(qdxb))
    parts += ["R", chuan_hoa(bbt or "").upper().replace(" ", "") or "XX", ma_lan_xb(lan_xb)]
    return "_".join(parts)


def thang_qdxb(v):
    """QĐXB 'mm/yyyy' hoặc 'dd/mm/yyyy' (hoặc datetime) -> 'MM'; thiếu -> '00'."""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%m")
    nums = re.findall(r"\d+", str(v or ""))
    return f"{int(nums[-2]):02d}" if len(nums) >= 2 else "00"


def duong_dan_dai(p):
    """Windows: thêm \\\\?\\ cho đường dẫn >240 ký tự."""
    if os.name == "nt" and len(p) > 240 and not p.startswith("\\\\?\\"):
        return "\\\\?\\" + os.path.abspath(p)
    return p


# ---------- quét kho ----------
COT_CSV = ["STT", "Tên file", "Định dạng", "Ngày sửa", "Năm", "Kích thước (KB)", "Thư mục", "Đường dẫn"]


def quet(thu_muc, exts, csv_path, progress=None):
    exts = {e if e.startswith(".") else "." + e for e in exts}
    goc = thu_muc
    if os.name == "nt" and not goc.startswith("\\\\?\\"):
        goc = "\\\\?\\" + os.path.abspath(goc)  # đường dẫn dài >260 ký tự
    n = 0
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        f.write("sep=,\r\n")  # Excel tiếng Việt mặc định ngăn cột bằng ';' -> báo cho Excel biết file dùng ','
        w = csv.writer(f)
        w.writerow(COT_CSV)
        for root, _dirs, files in os.walk(goc, onerror=lambda e: None):
            for fn in files:
                stem, ext = os.path.splitext(fn)
                ext = ext.lower()
                if exts and ext not in exts:
                    continue
                p = os.path.join(root, fn)
                try:
                    st = os.stat(p)
                except OSError:
                    continue
                d = dt.datetime.fromtimestamp(st.st_mtime)
                n += 1
                w.writerow([n, stem, ext.lstrip(".").upper(), d.strftime("%d/%m/%Y %H:%M"), d.year,
                            round(st.st_size / 1024), root.replace("\\\\?\\", "", 1), p.replace("\\\\?\\", "", 1)])
                if progress and n % 1000 == 0:
                    progress(n, root)
    return n


# ---------- đọc excel ----------
def _khoa_header(c):
    h = chuan_hoa(c)
    if "ten sach" in h: return "ten"
    if h.startswith("lan"): return "lan"
    if h in ("bbt", "ban bt", "ban bien tap"): return "bbt"
    if "qdxb" in h or "thoi gian" in h: return "qdxb"
    if h == "tt": return "tt"
    return None


def doc_excel(xlsx):
    """-> list dict {nam, dong, tt, ten, lan, bbt, qdxb}. Cột dò theo tên header (vị trí đổi theo năm)."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    sach = []
    for sh in wb.sheetnames:
        if not sh.strip().isdigit():
            continue
        nam, cot = int(sh), None
        for i, row in enumerate(wb[sh].iter_rows(values_only=True), 1):
            if cot is None:
                if any(c and "ten sach" in chuan_hoa(c) for c in row):
                    cot = {}
                    for idx, c in enumerate(row):
                        k = c and _khoa_header(c)
                        if k and k not in cot:
                            cot[k] = idx
                continue
            g = lambda k: row[cot[k]] if k in cot and cot[k] < len(row) else None
            ten = g("ten")
            if not ten or not str(ten).strip() or not chuan_hoa(ten):
                continue
            sach.append(dict(nam=nam, dong=i, tt=g("tt"), ten=str(ten).strip(),
                             lan=g("lan"), bbt=g("bbt"), qdxb=g("qdxb")))
    return sach


# ---------- đối chiếu ----------
def doc_csv(csv_path):
    """Đọc CSV kho (định dạng mới có header tiếng Việt, hoặc cũ ten_file/dinh_dang/date_modified/path)
    -> dict nội bộ: ten_file, dinh_dang (thường), date_modified (yyyy-mm-dd HH:MM:SS, để sort), path, key, nam."""
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        if not f.readline().startswith("sep="):
            f.seek(0)
        rows = list(csv.DictReader(f))
    out = []
    for r in rows:
        if "Tên file" in r:
            d = dt.datetime.strptime(r["Ngày sửa"], "%d/%m/%Y %H:%M")
            r = dict(ten_file=r["Tên file"], dinh_dang=r["Định dạng"].lower(),
                     date_modified=d.strftime("%Y-%m-%d %H:%M:%S"), path=r["Đường dẫn"])
        r["key"] = chuan_hoa(r["ten_file"])
        r["nam"] = int(r["date_modified"][:4])
        out.append(r)
    return out


def doi_chieu(sach, files):
    theo_key = set()
    theo_prefix = defaultdict(list)  # ponytail: so khớp "chứa" chỉ trong nhóm cùng 3 từ đầu, tránh O(sách×file)
    for f in files:
        theo_key.add(f["key"])
        theo_prefix[" ".join(f["key"].split()[:3])].append(f)

    kq = []
    for s in sach:
        key = chuan_hoa(s["ten"])
        tim = [f for f in theo_prefix.get(" ".join(key.split()[:3]), []) if key in f["key"]]  # gồm cả khớp chính xác
        kieu = "chính xác" if key in theo_key else ("chứa tên" if tim else "")
        khop_nam = [f for f in tim if f["nam"] == s["nam"]]

        r = dict(s, ten_moi=ten_file_moi(s["ten"], s["qdxb"], s["bbt"], s["lan"]),
                 kieu_khop=kieu, so_file=len(tim), file_word="", file_pdf="", ghi_chu="",
                 cac_file="\n".join(f"{f['date_modified']}  {f['path']}" for f in tim))
        if not tim:
            r["tinh_trang"] = "KHÔNG TÌM THẤY FILE"
        elif not khop_nam:
            r["tinh_trang"] = "LỆCH NĂM - cần mở file kiểm tra"
        else:
            r["tinh_trang"] = "KHỚP"
            w = [f for f in khop_nam if f["dinh_dang"] in WORD]
            p = [f for f in khop_nam if f["dinh_dang"] == "pdf"]
            if w: r["file_word"] = max(w, key=lambda f: f["date_modified"])["path"]
            if p: r["file_pdf"] = max(p, key=lambda f: f["date_modified"])["path"]
            if len(khop_nam) < len(tim):
                r["ghi_chu"] = f"{len(tim) - len(khop_nam)} file khác năm bị bỏ qua"
        kq.append(r)
    return kq


COT_KQ = [("nam", "Năm"), ("dong", "Dòng Excel"), ("tt", "TT"), ("ten", "Tên sách"), ("lan", "Lần XB"),
          ("bbt", "BBT"), ("qdxb", "QĐXB"), ("tinh_trang", "Tình trạng"), ("kieu_khop", "Kiểu khớp"),
          ("so_file", "Số file"), ("file_word", "File Word chọn"), ("file_pdf", "File PDF chọn"),
          ("ten_moi", "Tên file mới"), ("ghi_chu", "Ghi chú"), ("cac_file", "Tất cả file tìm được"),
          ("dich", "Đường dẫn đích")]


def ghi_ket_qua(kq, xlsx_goc, thu_muc_ra):
    import openpyxl
    os.makedirs(thu_muc_ra, exist_ok=True)
    tem = dt.datetime.now().strftime("%Y%m%d_%H%M")

    wb = openpyxl.Workbook()
    for ten, loc in (("KHOP", lambda r: r["tinh_trang"] == "KHỚP"),
                     ("CAN_XU_LY", lambda r: r["tinh_trang"] != "KHỚP")):
        ws = wb.create_sheet(ten)
        ws.append([h for _, h in COT_KQ])
        for r in kq:
            if loc(r):
                ws.append([_cell(r.get(k, "")) for k, _ in COT_KQ])
    del wb["Sheet"]
    ra1 = os.path.join(thu_muc_ra, f"doi_chieu_{tem}.xlsx")
    wb.save(ra1)

    # bước 6: đánh dấu tình trạng vào bản sao file excel gốc
    wb = openpyxl.load_workbook(xlsx_goc)
    theo_nam = defaultdict(list)
    for r in kq:
        theo_nam[r["nam"]].append(r)
    for nam, rs in theo_nam.items():
        ws = wb[str(nam)]
        c, hd = ws.max_column + 1, min(r["dong"] for r in rs) - 1
        for i, h in enumerate(("Tình trạng", "File chọn", "Tên file mới", "Đường dẫn đích")):
            ws.cell(hd, c + i, h)
        for r in rs:
            ws.cell(r["dong"], c, r["tinh_trang"])
            ws.cell(r["dong"], c + 1, "\n".join(x for x in (r["file_word"], r["file_pdf"]) if x))
            ws.cell(r["dong"], c + 2, r["ten_moi"])
            ws.cell(r["dong"], c + 3, r.get("dich", ""))
    ra2 = os.path.join(thu_muc_ra, f"{os.path.splitext(os.path.basename(xlsx_goc))[0]}_danh_dau_{tem}.xlsx")
    wb.save(ra2)
    return ra1, ra2


def _cell(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d/%m/%Y")
    return v


def doc_doi_chieu(xlsx):
    """Đọc lại file doi_chieu_*.xlsx (cả 2 sheet) -> list dict cùng khoá COT_KQ, sắp theo (năm, dòng)."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ten2k = {h: k for k, h in COT_KQ}
    kq = []
    for ws in wb:
        rows = ws.iter_rows(values_only=True)
        keys = [ten2k.get(h) for h in next(rows, ())]
        for row in rows:
            r = {k: ("" if v is None else v) for k, v in zip(keys, row) if k}
            if r.get("ten"):
                kq.append(r)
    kq.sort(key=lambda r: (r["nam"], r["dong"]))
    return kq


# ---------- kiểm tra nội dung ----------
def loi_noi_dau(text):
    """Text chuẩn hoá kể từ 'Lời nói đầu' / 'Lời Nhà xuất bản' / 'Lời giới thiệu' (tối đa 5000 ký tự), '' nếu không có."""
    t = chuan_hoa(text)
    m = re.search(r"\bloi (noi dau|nha xuat ban|gioi thieu|tua|mo dau)\b", t)
    return t[m.start():m.start() + 5000] if m else ""


def trich(path, so_trang=8):
    """-> {'nam': [năm 20xx ở 2 trang/20 đoạn cuối], 'lnd': lời nói đầu chuẩn hoá trong 8 trang/150 đoạn đầu}
    None nếu .doc hoặc không đọc được."""
    ext = os.path.splitext(path)[1].lower()
    p = duong_dan_dai(path)
    try:
        if ext == ".pdf":
            import fitz
            with fitz.open(p) as d:
                dau = " ".join(d[i].get_text() for i in range(min(so_trang, len(d))))
                cuoi = " ".join(d[i].get_text() for i in range(max(0, len(d) - 2), len(d)))
        elif ext == ".docx":
            import docx
            ps = [x.text for x in docx.Document(p).paragraphs if x.text.strip()]
            dau, cuoi = " ".join(ps[:150]), " ".join(ps[-20:])
        else:
            return None
    except Exception:
        return None
    return dict(nam=sorted({int(y) for y in re.findall(r"\b20\d\d\b", cuoi)}), lnd=loi_noi_dau(dau))


def doc_cache(json_path):
    try:
        with open(json_path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def ghi_cache(cache, json_path):
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def trich_cache(cache, path):
    """Như trich() nhưng nhớ theo (path, mtime) trong cache dict."""
    try:
        k = f"{path}|{int(os.path.getmtime(duong_dan_dai(path)))}"
    except OSError:
        return None
    if k not in cache:
        cache[k] = trich(path)
    return cache[k]


def _dong_file(cac_file):
    """'Tất cả file tìm được' -> [(date_modified, path)]"""
    return [(l[:19], l[21:]) for l in str(cac_file or "").splitlines() if len(l) > 21]


def _chon(r, ok, ghi_chu):
    """ok: [(date_modified, path)] đã xác nhận năm -> ghi KHỚP vào r (Word/PDF mới nhất). False nếu rỗng."""
    if not ok:
        return False
    w = [x for x in ok if x[1].rsplit(".", 1)[-1].lower() in WORD]
    p = [x for x in ok if x[1].lower().endswith(".pdf")]
    r.update(tinh_trang="KHỚP", file_word=max(w)[1] if w else "", file_pdf=max(p)[1] if p else "", ghi_chu=ghi_chu)
    return True


def xu_ly_lech_nam(r, cache):
    """Dòng LỆCH NĂM: mở từng file trong 'Tất cả file tìm được', nhận file có năm cột Năm ở trang cuối."""
    ok, ghi, thay = [], [], set()
    for d, p in _dong_file(r["cac_file"]):
        t = trich_cache(cache, p)
        if t is None:
            ghi.append(f"bỏ qua (.doc/không đọc được): {os.path.basename(p)}")
            continue
        thay.update(t["nam"])
        if r["nam"] in t["nam"]:
            ok.append((d, p))
    if not _chon(r, ok, "khớp theo nội dung"):
        ghi.insert(0, f"trang cuối không có năm {r['nam']} (thấy: {', '.join(map(str, sorted(thay))) or '-'})")
        r["ghi_chu"] = "; ".join(ghi)
    elif ghi:
        r["ghi_chu"] += "; " + "; ".join(ghi)


def loc_files(files, thu_muc="", nam=""):
    """Giới hạn file kho: pdf/docx, trong thư mục con, Date Modified trong năm/khoảng năm ('2015' hoặc '2015-2017')."""
    tm = os.path.normcase(os.path.abspath(thu_muc)) if thu_muc.strip() else ""
    ns = [int(x) for x in re.findall(r"\d{4}", nam)]
    lo, hi = (ns[0], ns[-1]) if ns else (0, 9999)
    return [f for f in files if f["dinh_dang"] in DOC_DUOC and lo <= f["nam"] <= hi
            and (not tm or os.path.normcase(f["path"]).startswith(tm))]


def xu_ly_khong_tim_thay(rows, files, cache, progress=None):
    """Dòng KHÔNG TÌM THẤY: tìm tên sách trong lời nói đầu của các file kho, rồi kiểm tra năm trang cuối."""
    theo_prefix = defaultdict(list)
    for r in rows:
        r["_key"], r["_tim"] = chuan_hoa(r["ten"]), {}
        theo_prefix[" ".join(r["_key"].split()[:3])].append(r)
    for i, f in enumerate(files, 1):
        t = trich_cache(cache, f["path"])
        if t and t["lnd"]:
            w, lnd = t["lnd"].split(), f" {t['lnd']} "
            for j in range(len(w)):  # ponytail: tra từng cụm 1-3 từ của lời nói đầu vào chỉ mục tên sách
                for n in (1, 2, 3):
                    for r in theo_prefix.get(" ".join(w[j:j + n]), ()):
                        if f" {r['_key']} " in lnd:
                            r["_tim"][f["path"]] = (f, t)
        if progress and i % 200 == 0:
            progress(i)
    for r in rows:
        tim, _ = r.pop("_tim"), r.pop("_key")
        if not tim:
            r["tinh_trang"] = "KHÔNG TÌM ĐƯỢC"
            continue
        r.update(so_file=len(tim), kieu_khop="lời nói đầu",
                 cac_file="\n".join(f"{f['date_modified']}  {f['path']}" for f, _ in tim.values()))
        ok = [(f["date_modified"], f["path"]) for f, t in tim.values() if r["nam"] in t["nam"]]
        if not _chon(r, ok, "khớp theo nội dung (lời nói đầu)"):
            r["tinh_trang"] = "LỆCH NĂM - tìm theo lời nói đầu"
            r["ghi_chu"] = f"trang cuối không có năm {r['nam']}"


# ---------- đổi tên & lưu ----------
def ke_hoach_luu(kq, thu_muc_dich):
    """-> [(r, nguồn, đích)] cho dòng KHỚP: <đích>/<Năm>/thang <MM>/<Tên file mới><đuôi>; trùng tên thêm _2, _3."""
    da_dung, kh = set(), []
    for r in kq:
        if r["tinh_trang"] != "KHỚP":
            continue
        for src in (r["file_word"], r["file_pdf"]):
            if not src:
                continue
            goc = os.path.join(thu_muc_dich, str(r["nam"]), f"thang {thang_qdxb(r['qdxb'])}", r["ten_moi"])
            ext = os.path.splitext(src)[1].lower()
            p, i = goc + ext, 1
            while p.lower() in da_dung or os.path.exists(duong_dan_dai(p)):
                i += 1
                p = f"{goc}_{i}{ext}"
            da_dung.add(p.lower())
            kh.append((r, src, p))
    return kh


def luu_file(kh, di_chuyen=False, progress=None):
    """Copy (hoặc move) theo kế hoạch, ghi đường dẫn đích vào r['dich']."""
    for i, (r, src, p) in enumerate(kh, 1):
        os.makedirs(duong_dan_dai(os.path.dirname(p)), exist_ok=True)
        (shutil.move if di_chuyen else shutil.copy2)(duong_dan_dai(src), duong_dan_dai(p))
        r["dich"] = "\n".join(x for x in (r.get("dich", ""), p) if x)
        if progress and i % 100 == 0:
            progress(i)
    return len(kh)


# ---------- giao diện ----------
def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("Tool SXB - quét kho, đối chiếu, kiểm tra nội dung, đổi tên")
    root.geometry("800x520")
    nb = ttk.Notebook(root); nb.pack(fill="both", expand=True, padx=8, pady=8)
    chon_xlsx = lambda: filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    chon_csv = lambda: filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])

    def hang(parent, r, nhan, var, chon=None):
        ttk.Label(parent, text=nhan, width=20).grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, width=70).grid(row=r, column=1, sticky="we")
        if chon:
            ttk.Button(parent, text="...", width=3, command=lambda: var.set(chon() or var.get())).grid(row=r, column=2)

    def chay(nut, ham):
        nut["state"] = "disabled"
        def run():
            try:
                ham()
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Lỗi", str(e)))
            finally:
                root.after(0, lambda: nut.config(state="normal"))
        threading.Thread(target=run, daemon=True).start()

    def logger(txt):
        return lambda s: root.after(0, lambda: (txt.insert("end", s + "\n"), txt.see("end")))

    # --- tab 1: quét ---
    t1 = ttk.Frame(nb, padding=10); nb.add(t1, text="1. Quét kho → CSV")
    v_tm, v_ext, v_csv = tk.StringVar(), tk.StringVar(value=EXT_MAC_DINH), tk.StringVar()
    hang(t1, 0, "Thư mục kho", v_tm, filedialog.askdirectory)
    hang(t1, 1, "Định dạng (cách)", v_ext)
    hang(t1, 2, "File CSV xuất", v_csv,
         lambda: filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")]))
    tt1 = ttk.Label(t1, text="", wraplength=700); tt1.grid(row=4, column=0, columnspan=3, sticky="w", pady=8)

    def quet_gui():
        if not os.path.isdir(v_tm.get()): raise ValueError("Chưa chọn thư mục kho")
        if not v_csv.get(): v_csv.set(os.path.join(os.getcwd(), f"kho_{dt.datetime.now():%Y%m%d_%H%M}.csv"))
        def prog(n, root_dir): root.after(0, lambda: tt1.config(text=f"Đã quét {n} file... {root_dir[-80:]}"))
        n = quet(v_tm.get(), v_ext.get().split(), v_csv.get(), prog)
        root.after(0, lambda: tt1.config(text=f"Xong: {n} file → {v_csv.get()}"))
    b1 = ttk.Button(t1, text="Quét", command=lambda: chay(b1, quet_gui)); b1.grid(row=3, column=1, sticky="w", pady=6)

    # --- tab 2: đối chiếu ---
    t2 = ttk.Frame(nb, padding=10); nb.add(t2, text="2. Đối chiếu Excel")
    v_xl, v_csv2, v_ra = tk.StringVar(), tk.StringVar(), tk.StringVar(value=os.getcwd())
    hang(t2, 0, "File Excel tổng hợp", v_xl, chon_xlsx)
    hang(t2, 1, "File CSV kho", v_csv2, lambda: chon_csv() or v_csv.get())
    hang(t2, 2, "Thư mục kết quả", v_ra, filedialog.askdirectory)
    tt2 = tk.Text(t2, height=9, wrap="word"); tt2.grid(row=4, column=0, columnspan=3, sticky="we", pady=8)

    def doi_chieu_gui():
        csvp = v_csv2.get() or v_csv.get()
        if not os.path.isfile(v_xl.get()): raise ValueError("Chưa chọn file Excel")
        if not os.path.isfile(csvp): raise ValueError("Chưa chọn file CSV kho")
        log = logger(tt2)
        log("Đọc Excel..."); sach = doc_excel(v_xl.get()); log(f"  {len(sach)} đầu sách")
        log("Đọc CSV..."); files = doc_csv(csvp); log(f"  {len(files)} file")
        kq = doi_chieu(sach, files)
        dem = defaultdict(int)
        for r in kq: dem[r["tinh_trang"]] += 1
        for k, v in dem.items(): log(f"  {k}: {v}")
        ra1, ra2 = ghi_ket_qua(kq, v_xl.get(), v_ra.get())
        v_dc.set(ra1); v_dc4.set(ra1)
        log(f"Kết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b2 = ttk.Button(t2, text="Đối chiếu", command=lambda: chay(b2, doi_chieu_gui)); b2.grid(row=3, column=1, sticky="w", pady=6)

    # --- tab 3: kiểm tra nội dung ---
    t3 = ttk.Frame(nb, padding=10); nb.add(t3, text="3. Kiểm tra nội dung")
    v_dc, v_xl3, v_csv3, v_ra3 = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
    v_tmc, v_nam3 = tk.StringVar(), tk.StringVar()
    v_lech, v_ktt = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    hang(t3, 0, "File doi_chieu_*.xlsx", v_dc, chon_xlsx)
    hang(t3, 1, "File Excel tổng hợp", v_xl3, chon_xlsx)
    hang(t3, 2, "File CSV kho", v_csv3, chon_csv)
    hang(t3, 3, "Thư mục kết quả", v_ra3, filedialog.askdirectory)
    ttk.Checkbutton(t3, text="Xử lý LỆCH NĂM (đọc năm ở trang cuối các file đã tìm được)", variable=v_lech)\
        .grid(row=4, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(t3, text="Xử lý KHÔNG TÌM THẤY (quét lời nói đầu các file trong kho, giới hạn bên dưới)", variable=v_ktt)\
        .grid(row=5, column=0, columnspan=3, sticky="w")
    hang(t3, 6, "  Giới hạn thư mục con", v_tmc, filedialog.askdirectory)
    hang(t3, 7, "  Giới hạn năm DM", v_nam3)
    ttk.Label(t3, text="(vd 2015 hoặc 2015-2017; để trống = toàn bộ kho)").grid(row=7, column=2, columnspan=2, sticky="w")
    tt3 = tk.Text(t3, height=8, wrap="word"); tt3.grid(row=9, column=0, columnspan=3, sticky="we", pady=8)

    def kiem_tra_gui():
        xl, csvp, ra = v_xl3.get() or v_xl.get(), v_csv3.get() or v_csv2.get() or v_csv.get(), v_ra3.get() or v_ra.get()
        for p, m in ((v_dc.get(), "file doi_chieu"), (xl, "file Excel tổng hợp"), (csvp, "file CSV kho")):
            if not os.path.isfile(p): raise ValueError(f"Chưa chọn {m}")
        log = logger(tt3)
        kq = doc_doi_chieu(v_dc.get()); files = doc_csv(csvp)
        cache_p = csvp + ".cache.json"; cache = doc_cache(cache_p)
        log(f"{len(kq)} dòng, {len(files)} file kho, cache {len(cache)} file")
        try:
            if v_lech.get():
                lech = [r for r in kq if str(r["tinh_trang"]).startswith("LỆCH NĂM")]
                log(f"LỆCH NĂM: {len(lech)} dòng...")
                for i, r in enumerate(lech, 1):
                    xu_ly_lech_nam(r, cache)
                    if i % 50 == 0: log(f"  {i}/{len(lech)}"); ghi_cache(cache, cache_p)
            if v_ktt.get():
                ktt = [r for r in kq if str(r["tinh_trang"]).startswith("KHÔNG TÌM")]
                ds = loc_files(files, v_tmc.get(), v_nam3.get())
                log(f"KHÔNG TÌM THẤY: {len(ktt)} dòng, quét lời nói đầu {len(ds)} file kho...")
                def prog(i):
                    log(f"  {i}/{len(ds)} file")
                    if i % 1000 == 0: ghi_cache(cache, cache_p)
                xu_ly_khong_tim_thay(ktt, ds, cache, prog)
        finally:
            ghi_cache(cache, cache_p)
        dem = defaultdict(int)
        for r in kq: dem[r["tinh_trang"]] += 1
        for k, v in dem.items(): log(f"  {k}: {v}")
        ra1, ra2 = ghi_ket_qua(kq, xl, ra)
        v_dc4.set(ra1)
        log(f"Kết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b3 = ttk.Button(t3, text="Kiểm tra", command=lambda: chay(b3, kiem_tra_gui)); b3.grid(row=8, column=1, sticky="w", pady=6)

    # --- tab 4: đổi tên & lưu ---
    t4 = ttk.Frame(nb, padding=10); nb.add(t4, text="4. Đổi tên & lưu")
    v_dc4, v_xl4, v_dich, v_ra4, v_move = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.BooleanVar()
    hang(t4, 0, "File doi_chieu_*.xlsx", v_dc4, chon_xlsx)
    hang(t4, 1, "File Excel tổng hợp", v_xl4, chon_xlsx)
    hang(t4, 2, "Thư mục đích", v_dich, filedialog.askdirectory)
    hang(t4, 3, "Thư mục kết quả", v_ra4, filedialog.askdirectory)
    ttk.Checkbutton(t4, text="Di chuyển file (mặc định: copy)", variable=v_move).grid(row=4, column=0, columnspan=3, sticky="w")
    tt4 = tk.Text(t4, height=10, wrap="none"); tt4.grid(row=6, column=0, columnspan=3, sticky="we", pady=8)

    def luu_gui(thu):
        xl, ra = v_xl4.get() or v_xl3.get() or v_xl.get(), v_ra4.get() or v_ra3.get() or v_ra.get()
        if not os.path.isfile(v_dc4.get()): raise ValueError("Chưa chọn file doi_chieu")
        if not os.path.isdir(v_dich.get()): raise ValueError("Chưa chọn thư mục đích")
        if not thu and not os.path.isfile(xl): raise ValueError("Chưa chọn file Excel tổng hợp")
        log = logger(tt4)
        kq = doc_doi_chieu(v_dc4.get()); kh = ke_hoach_luu(kq, v_dich.get())
        log(f"{len(kh)} file sẽ {'di chuyển' if v_move.get() else 'copy'} sang {v_dich.get()}")
        if thu:
            log("\n".join(f"{s}\n   → {p}" for _, s, p in kh) + "\n(chạy thử - chưa copy)")
            return
        luu_file(kh, v_move.get(), lambda i: log(f"  {i}/{len(kh)}"))
        ra1, ra2 = ghi_ket_qua(kq, xl, ra)
        log(f"Xong {len(kh)} file.\nKết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b4a = ttk.Button(t4, text="Chạy thử", command=lambda: chay(b4a, lambda: luu_gui(True))); b4a.grid(row=5, column=1, sticky="w", pady=6)
    b4b = ttk.Button(t4, text="Thực hiện", command=lambda: chay(b4b, lambda: luu_gui(False))); b4b.grid(row=5, column=1, sticky="e", pady=6)

    root.mainloop()


# ---------- tự kiểm tra ----------
def _test():
    import tempfile, fitz, docx, openpyxl
    assert chuan_hoa("Lịch sử Chính phủ Việt Nam- T1") == "lich su chinh phu viet nam t1"
    assert chuan_hoa("Đổi mới  công tác\n") == "doi moi cong tac"
    assert ma_ngay("01-2015") == "0115" and ma_ngay("1/2025") == "0125"
    assert ma_ngay("19/1/2007") == "190107" and ma_ngay("11-2-18") == "110218"
    assert ma_ngay(dt.datetime(2008, 1, 8)) == "080108" and ma_ngay(None) == ""
    assert ma_lan_xb(None) == "M00" and ma_lan_xb(1) == "M00" and ma_lan_xb("TB") == "T01"
    assert ma_lan_xb(2) == "T02" and ma_lan_xb("Lần 3") == "T03" and ma_lan_xb("TB lần 2") == "T02"
    assert ten_file_moi("Tên sách", "1/2021", "PL", "TB") == "ten_sach_0121_R_PL_T01"
    assert ten_file_moi("Tên sách", None, "Đ", None) == "ten_sach_R_D_M00"
    assert thang_qdxb("05/2023") == "05" and thang_qdxb("19/01/2007") == "01" and thang_qdxb(None) == "00"
    # quét -> CSV đẹp -> đọc lại ra dict nội bộ
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "Sach mau.PDF"); open(p, "wb").write(b"x" * 2048)
        t = dt.datetime(2022, 3, 5, 14, 30).timestamp(); os.utime(p, (t, t))
        csvp = os.path.join(td, "kho.csv"); assert quet(td, [".pdf"], csvp) == 1
        lines = open(csvp, encoding="utf-8-sig").read().splitlines()
        assert lines[0] == "sep=," and lines[1].split(",")[:3] == ["STT", "Tên file", "Định dạng"]
        assert lines[2].startswith("1,Sach mau,PDF,05/03/2022 14:30,2022,2,")
        f, = doc_csv(csvp)
        assert (f["dinh_dang"], f["date_modified"], f["nam"], f["path"]) == ("pdf", "2022-03-05 14:30:00", 2022, p)
    sach = [dict(nam=2025, dong=4, tt=1, ten="Dạ thưa, xứ Huế bây chừ", lan=None, bbt="QT", qdxb="1/2025"),
            dict(nam=2025, dong=5, tt=2, ten="Luật đầu tư", lan="TB", bbt="PL", qdxb="2/2025"),
            dict(nam=2025, dong=6, tt=3, ten="Không có", lan=None, bbt="PL", qdxb=None)]
    files = [dict(ten_file="Da thua xu Hue bay chu", dinh_dang="docx", date_modified="2025-03-01 10:00:00", path="a.docx"),
             dict(ten_file="Da thua xu Hue bay chu", dinh_dang="docx", date_modified="2025-05-01 10:00:00", path="b.docx"),
             dict(ten_file="Da thua xu Hue bay chu", dinh_dang="pdf", date_modified="2024-05-01 10:00:00", path="c.pdf"),
             dict(ten_file="Luat dau tu (ban in)", dinh_dang="pdf", date_modified="2023-01-01 00:00:00", path="d.pdf"),
             dict(ten_file="Da thua xu Hue bay chu - ban in", dinh_dang="pdf", date_modified="2025-02-01 00:00:00", path="e.pdf")]
    for f in files:
        f["key"] = chuan_hoa(f["ten_file"]); f["nam"] = int(f["date_modified"][:4])
    kq = doi_chieu(sach, files)
    assert kq[0]["tinh_trang"] == "KHỚP" and kq[0]["file_word"] == "b.docx" and kq[0]["file_pdf"] == "e.pdf"
    assert kq[0]["so_file"] == 4 and kq[0]["ghi_chu"] == "1 file khác năm bị bỏ qua"
    assert kq[1]["tinh_trang"].startswith("LỆCH NĂM") and kq[1]["kieu_khop"] == "chứa tên"
    assert kq[2]["tinh_trang"] == "KHÔNG TÌM THẤY FILE"

    # --- tab 3: trích text pdf/docx, cache ---
    tmp = tempfile.mkdtemp()
    pdf, dx, doc = (os.path.join(tmp, n) for n in ("a.pdf", "b.docx", "c.doc"))
    d = fitz.open()
    d.new_page().insert_text((50, 72), "Trang bia")
    d.new_page().insert_text((50, 72), "LOI NOI DAU\nCuon sach Luat dau tu nay gioi thieu...")
    d.new_page().insert_text((50, 72), "NHA XUAT BAN CHINH TRI QUOC GIA\nIn xong nam 2023")
    d.save(pdf); d.close()
    w = docx.Document()
    for s in ("Bia", "Lời Nhà xuất bản", "Cuốn sách Luật đầu tư được tái bản.", "Nội dung", "Chịu trách nhiệm xuất bản 2019"):
        w.add_paragraph(s)
    w.save(dx)
    open(doc, "wb").write(b"\xd0\xcf\x11\xe0 junk")
    t = trich(pdf)
    assert t["nam"] == [2023] and t["lnd"].startswith("loi noi dau") and "luat dau tu" in t["lnd"]
    assert trich(dx)["nam"] == [2019] and "luat dau tu" in trich(dx)["lnd"] and trich(doc) is None
    cache = {}
    assert trich_cache(cache, pdf) == t and len(cache) == 1 and trich_cache(cache, "khong_co.pdf") is None
    cj = os.path.join(tmp, "c.json"); ghi_cache(cache, cj); assert doc_cache(cj) == cache and doc_cache("x.json") == {}

    r = dict(nam=2023, ten="Luật đầu tư", tinh_trang="LỆCH NĂM - cần mở file kiểm tra", file_word="", file_pdf="", ghi_chu="",
             cac_file=f"2021-01-01 00:00:00  {pdf}\n2021-01-02 00:00:00  {doc}\n2021-01-03 00:00:00  {dx}")
    xu_ly_lech_nam(r, cache)
    assert r["tinh_trang"] == "KHỚP" and r["file_pdf"] == pdf and r["file_word"] == "" and r["ghi_chu"].startswith("khớp theo nội dung; bỏ qua")
    r2 = dict(r, nam=2020, tinh_trang="LỆCH NĂM - cần mở file kiểm tra"); xu_ly_lech_nam(r2, cache)
    assert r2["tinh_trang"].startswith("LỆCH NĂM") and "thấy: 2019, 2023" in r2["ghi_chu"]

    kho = [dict(ten_file="a", dinh_dang="pdf", date_modified="2021-01-01 00:00:00", path=pdf, nam=2021),
           dict(ten_file="b", dinh_dang="docx", date_modified="2021-06-01 00:00:00", path=dx, nam=2021),
           dict(ten_file="c", dinh_dang="doc", date_modified="2021-06-01 00:00:00", path=doc, nam=2021)]
    assert len(loc_files(kho)) == 2 and len(loc_files(kho, "", "2020")) == 0 and len(loc_files(kho, tmp, "2020-2021")) == 2
    assert loc_files(kho, os.path.join(tmp, "khac")) == []
    rows = [dict(nam=2023, ten="Luật đầu tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2023, ten="Không có", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2019, ten="Luật đầu tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2018, ten="Luật tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file="")]
    xu_ly_khong_tim_thay(rows, loc_files(kho), cache)
    assert rows[0]["tinh_trang"] == "KHỚP" and rows[0]["file_pdf"] == pdf and rows[0]["so_file"] == 2
    assert rows[1]["tinh_trang"] == "KHÔNG TÌM ĐƯỢC"
    assert rows[2]["tinh_trang"] == "KHỚP" and rows[2]["file_word"] == dx and rows[2]["file_pdf"] == ""
    assert rows[3]["tinh_trang"] == "KHÔNG TÌM ĐƯỢC"  # "luat tu" không là cụm liền trong lời nói đầu

    # --- ghi / đọc lại doi_chieu, tab 4 ---
    xl = os.path.join(tmp, "th.xlsx"); wb = openpyxl.Workbook(); ws = wb.active; ws.title = "2023"
    ws.append(["TT", "Tên sách"]); ws.append([1, "Luật đầu tư"]); ws.append([2, "Không có"]); wb.save(xl)
    kq = [dict(nam=2023, dong=2, tt=1, ten="Luật đầu tư", lan="TB", bbt="PL", qdxb="05/2023", tinh_trang="KHỚP",
               kieu_khop="chính xác", so_file=2, file_word=dx, file_pdf=pdf, ten_moi="luat_dau_tu_0523_R_PL_T01",
               ghi_chu="", cac_file=""),
          dict(nam=2023, dong=3, tt=2, ten="Không có", lan=None, bbt="PL", qdxb=None, tinh_trang="KHÔNG TÌM THẤY FILE",
               kieu_khop="", so_file=0, file_word="", file_pdf="", ten_moi="khong_co_R_PL_M00", ghi_chu="", cac_file="")]
    ra1, ra2 = ghi_ket_qua(kq, xl, os.path.join(tmp, "ra"))
    kq2 = doc_doi_chieu(ra1)
    assert [r["ten"] for r in kq2] == ["Luật đầu tư", "Không có"] and kq2[0]["file_pdf"] == pdf and kq2[0]["dong"] == 2
    kq2.append(dict(kq2[0], dong=4, qdxb=None))  # trùng tên file mới -> _2
    dich = os.path.join(tmp, "dich"); kh = ke_hoach_luu(kq2, dich)
    assert [os.path.relpath(p, dich) for _, _, p in kh] == [
        os.path.join("2023", "thang 05", "luat_dau_tu_0523_R_PL_T01.docx"), os.path.join("2023", "thang 05", "luat_dau_tu_0523_R_PL_T01.pdf"),
        os.path.join("2023", "thang 00", "luat_dau_tu_0523_R_PL_T01.docx"), os.path.join("2023", "thang 00", "luat_dau_tu_0523_R_PL_T01.pdf")]
    kq2[2]["qdxb"] = "05/2023"; kh = ke_hoach_luu(kq2, dich)
    assert kh[2][2].endswith("luat_dau_tu_0523_R_PL_T01_2.docx") and kh[3][2].endswith("_2.pdf")
    assert luu_file(kh) == 4 and all(os.path.isfile(p) for _, _, p in kh) and os.path.isfile(pdf)
    assert kq2[0]["dich"] == kh[0][2] + "\n" + kh[1][2]
    assert ke_hoach_luu(kq2[:1], dich)[0][2].endswith("_3.docx")  # chạy lại: đích đã có -> _3
    ra1, ra2 = ghi_ket_qua(kq2, xl, os.path.join(tmp, "ra"))
    ws = openpyxl.load_workbook(ra2)["2023"]
    assert ws.cell(1, 6).value == "Đường dẫn đích" and ws.cell(2, 6).value == kq2[0]["dich"] and ws.cell(2, 3).value == "KHỚP"
    shutil.rmtree(tmp, ignore_errors=True)
    print("OK")


if __name__ == "__main__":
    _test() if "--test" in sys.argv else gui()
