"""Tool SXB: quét kho file -> Excel, đối chiếu với 'Tong hop SXB 2006-2025.xlsx', phân loại,
kiểm tra nội dung file (năm trang cuối, lời nói đầu), đổi tên & lưu theo Năm/thang MM.

Chạy:  python -X utf8 sxb_tool.py            (giao diện)
       python -X utf8 sxb_tool.py --test     (tự kiểm tra)
"""
import csv, glob, html, json, os, re, shutil, sqlite3, subprocess, sys, tempfile, threading, time, unicodedata, zipfile, datetime as dt
from collections import Counter, defaultdict

EXT_MAC_DINH = ".doc .docx .pdf .jpg .jpeg"
WORD = {"doc", "docx"}
DOC_DUOC = {"pdf", "docx"}  # định dạng trích được text (.doc bỏ qua)
BIA = re.compile(r"(?:^| )(bia|to che|tr ?2 4)(?= |$)")  # file bìa / tờ che / trang 2-4 (đã chuan_hoa)
GIOI_HAN_O = 32000  # Excel chỉ chứa 32.767 ký tự/ô; dài hơn thì cắt theo dòng
KY_DAU, KY_CUOI = 25000, 8000  # .doc/.docx: ~8 trang đầu, ~2-3 trang cuối tính theo số ký tự (không bắt Word dàn trang)
GIAY_MOI_FILE = 0.25  # thời gian đọc 1 PDF/DOCX (đo trên ổ cục bộ), để ước tính trước khi quét
DUNG = threading.Event()  # nút Dừng: các vòng lặp dài kiểm tra cờ này sau mỗi file


class DaDung(Exception):
    pass


def _dung():
    if DUNG.is_set():
        raise DaDung("Đã dừng theo yêu cầu.")


def _gio(giay):
    """Số giây -> '45 giây' / '12 phút' / '3,5 giờ'."""
    if giay < 90:
        return f"{giay:.0f} giây"
    return f"{giay / 60:.0f} phút" if giay < 5400 else f"{giay / 3600:.1f} giờ".replace(".", ",")


def _ds_file(ds):
    """[(date_modified, path)] -> chuỗi 'Tất cả file tìm được': Word/PDF trước, ảnh sau (ô Excel bị cắt thì chỉ mất ảnh)."""
    ds = sorted(ds, key=lambda x: (not x[1].lower().endswith((".pdf", ".doc", ".docx")), x[1]))
    return "\n".join(f"{d}  {p}" for d, p in ds)


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


DAI_TEN = 90  # độ dài tối đa phần tên sách: thư mục sách + tên file (đều = tên mới) giữ đường dẫn dưới 260 ký tự


def ten_file_moi(ten_sach, qdxb, bbt, lan_xb):
    """Tên mới 'ten_sach_0121_R_PL_T01'; phần tên sách dài quá DAI_TEN ký tự thì cắt ở ranh giới từ."""
    ten = "_".join(chuan_hoa(ten_sach).split())
    parts = [ten if len(ten) <= DAI_TEN else ten[:DAI_TEN + 1].rsplit("_", 1)[0]]
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
COT_KHO = ["STT", "Tên file", "Định dạng", "Ngày sửa", "Năm", "Kích thước (KB)", "Đường dẫn"]
RONG_KHO = [6, 70, 10, 17, 7, 14, 100]
MAX_DONG = 1_048_575  # ponytail: Excel giới hạn 1.048.576 dòng/sheet -> quá thì mở sheet "Kho 2", "Kho 3"...


def quet(thu_muc, exts, xlsx_path, progress=None, bao_cao=None):
    """Quét kho -> file .xlsx có định dạng (header đậm, cố định dòng 1, độ rộng cột, wrap text).
    bao_cao (dict) nhận: 'loi' = [thư mục/file không đọc được], 'dinh_dang' = Counter số file theo định dạng."""
    bao_cao = {} if bao_cao is None else bao_cao
    loi, dem = bao_cao.setdefault("loi", []), bao_cao.setdefault("dinh_dang", Counter())
    bo = lambda x: str(x or "").replace("\\\\?\\", "", 1)
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Alignment, Font
    exts = {e if e.startswith(".") else "." + e for e in exts}
    goc = thu_muc
    if os.name == "nt" and not goc.startswith("\\\\?\\"):
        goc = "\\\\?\\" + os.path.abspath(goc)  # đường dẫn dài >260 ký tự
    wb = Workbook(write_only=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    def o(ws, v, **kw):
        c = WriteOnlyCell(ws, value=v)
        for k, x in kw.items(): setattr(c, k, x)
        return c

    def sheet_moi():
        ws = wb.create_sheet(f"Kho {len(wb.sheetnames) + 1}" if wb.sheetnames else "Kho")
        for cot, r in zip("ABCDEFG", RONG_KHO):
            ws.column_dimensions[cot].width = r
        ws.freeze_panes = "A2"
        ws.append([o(ws, h, font=Font(bold=True), alignment=Alignment(horizontal="center")) for h in COT_KHO])
        return ws

    ws, n = sheet_moi(), 0
    for root, _dirs, files in os.walk(goc, onerror=lambda e: loi.append(f"{bo(e.filename)}: {e.strerror}")):
        for fn in files:
            if DUNG.is_set():  # dừng: không lưu file kho dở, xoá ngay file tạm của openpyxl (quét lớn thì rất nặng)
                for w in wb.worksheets:
                    try:
                        w.close(); w._writer.cleanup()
                    except Exception:
                        pass
                _dung()
            stem, ext = os.path.splitext(fn)
            ext = ext.lower()
            if exts and ext not in exts:
                continue
            p = os.path.join(root, fn)
            try:
                st = os.stat(p)
            except OSError as e:
                loi.append(f"{bo(p)}: {e.strerror}")
                continue
            dem[ext.lstrip(".").upper()] += 1
            d = dt.datetime.fromtimestamp(st.st_mtime)
            n += 1
            if n % MAX_DONG == 1 and n > 1:
                ws = sheet_moi()
            ws.append([n, o(ws, stem, alignment=wrap), ext.lstrip(".").upper(),
                       o(ws, d.replace(second=0, microsecond=0), number_format="dd/mm/yyyy hh:mm"),
                       d.year, round(st.st_size / 1024), o(ws, p.replace("\\\\?\\", "", 1), alignment=wrap)])
            if progress and n % 1000 == 0:
                progress(n, root)
    wb.save(xlsx_path)
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
    wb.close()
    if not sach:
        raise ValueError(f"'{os.path.basename(xlsx)}' không phải file Excel tổng hợp: cần các sheet năm (2006, 2007...) "
                         "có cột 'Tên sách'. Có thể đang chọn nhầm file kho vào ô này.")
    return sach


# ---------- đối chiếu ----------
def doc_kho(path):
    """Đọc file kho: .xlsx do quet() xuất, hoặc .csv cũ (utf-8 phẩy / utf-16 tab, header tiếng Việt hoặc máy)
    -> dict nội bộ: ten_file, dinh_dang (thường), date_modified (yyyy-mm-dd HH:MM:SS, để sort), path, key, nam."""
    rows = []
    if path.lower().endswith(".xlsx"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            if list(next(it, None) or ())[:len(COT_KHO)] != COT_KHO:
                wb.close()
                raise ValueError(f"'{os.path.basename(path)}' không phải file kho do tab 1 xuất (dòng đầu phải là: "
                                 f"{', '.join(COT_KHO)}). Có thể đang chọn nhầm file Excel tổng hợp vào ô này.")
            for r in it:
                _dung()
                if r and r[1] is not None:
                    rows.append(_dong_kho(str(r[1]), str(r[2]), r[3].strftime("%Y-%m-%d %H:%M:%S"), str(r[6])))
        wb.close()
    else:
        with open(path, "rb") as f:
            utf16 = f.read(2) == b"\xff\xfe"
        with open(path, encoding="utf-16" if utf16 else "utf-8-sig", newline="") as f:
            if not utf16 and not f.readline().startswith("sep="):
                f.seek(0)
            for r in csv.DictReader(f, delimiter="\t" if utf16 else ","):
                if "Tên file" in r:
                    d = dt.datetime.strptime(r["Ngày sửa"], "%d/%m/%Y %H:%M")
                    rows.append(_dong_kho(r["Tên file"], r["Định dạng"], d.strftime("%Y-%m-%d %H:%M:%S"), r["Đường dẫn"]))
                else:
                    rows.append(_dong_kho(r["ten_file"], r["dinh_dang"], r["date_modified"], r["path"]))
    return rows


def _dong_kho(ten_file, dinh_dang, date_modified, path):
    """Một dòng kho dạng gọn (~0,6 KB): chỉ giữ tên đã chuẩn hoá. ponytail: 4 triệu file vẫn cần vài GB RAM."""
    return dict(key=chuan_hoa(ten_file), dinh_dang=sys.intern(dinh_dang.lower()), date_modified=date_modified,
                path=path, nam=int(date_modified[:4]))


def key_thu_muc(thu_muc):
    """Tên thư mục này và các thư mục cha (chuẩn hoá, bỏ ổ đĩa và hậu tố bản sao "(1)")."""
    return [chuan_hoa(re.sub(r"\s*\(\d+\)\s*$", "", x)) for x in re.split(r"[\\/]", thu_muc)
            if x and not x.endswith(":") and x != "?"]


def la_bia(path, key_sach=""):
    """File bìa / tờ che / trang 2-4 — trừ khi chính tên sách có từ đó."""
    return bool(BIA.search(chuan_hoa(os.path.splitext(os.path.basename(path))[0]))) and not BIA.search(key_sach)


def chon_file(ds, key_sach=""):
    """ds: [(date_modified, path)] -> (word, pdf, số file bìa bỏ qua). Mới nhất; bằng nhau lấy file lớn nhất."""
    def kt(p):
        try: return os.path.getsize(duong_dan_dai(p))
        except OSError: return 0
    chinh = [x for x in ds if not la_bia(x[1], key_sach)]
    w = [x for x in chinh if x[1].rsplit(".", 1)[-1].lower() in WORD]
    p = [x for x in chinh if x[1].lower().endswith(".pdf")]
    best = lambda xs: max(xs, key=lambda x: (x[0], kt(x[1])))[1] if xs else ""
    return best(w), best(p), len(ds) - len(chinh)


def doi_chieu(sach, files):
    """Khớp tên sách với tên file HOẶC tên thư mục chứa file (kho xếp mỗi sách một thư mục)."""
    theo_key, ten_sach = set(), {chuan_hoa(s["ten"]) for s in sach}
    theo_prefix = defaultdict(list)  # ponytail: so khớp "chứa" chỉ trong nhóm cùng 3 từ đầu, tránh O(sách×file)
    theo_tm = defaultdict(list)
    for f in files:
        theo_key.add(f["key"])
        theo_prefix[" ".join(f["key"].split()[:3])].append((f["key"], False, f))
        theo_tm[os.path.dirname(f["path"])].append(f)
    for tm in theo_tm:  # tên thư mục: mỗi thư mục một mục chỉ mục cho nó và từng thư mục cha, không lặp theo file
        for k in set(key_thu_muc(tm)):
            theo_key.add(k)
            theo_prefix[" ".join(k.split()[:3])].append((k, True, tm))

    kq = []
    for s in sach:
        key = chuan_hoa(s["ten"])
        tim, rut_gon = {}, False
        for k, la_tm, x in theo_prefix.get(" ".join(key.split()[:3]), ()):
            khop = f" {key} " in f" {k} "  # gồm cả khớp chính xác
            gon = not khop and la_tm and len(k.split()) >= 3 and key.startswith(k + " ") and k not in ten_sach  # thư mục tên rút gọn
            if khop or gon:
                tim.update((f["path"], f) for f in (theo_tm[x] if la_tm else (x,)))
                rut_gon |= gon
        tim = list(tim.values())
        kieu = "chính xác" if key in theo_key else "thư mục rút gọn" if rut_gon else "chứa tên" if tim else ""
        khop_nam = [f for f in tim if f["nam"] == s["nam"]]

        r = dict(s, ten_moi=ten_file_moi(s["ten"], s["qdxb"], s["bbt"], s["lan"]),
                 kieu_khop=kieu, so_file=len(tim), file_word="", file_pdf="", ghi_chu="",
                 cac_file=_ds_file((f["date_modified"], f["path"]) for f in tim))
        if not tim:
            r["tinh_trang"] = "KHÔNG TÌM THẤY FILE"
        elif not khop_nam:
            r["tinh_trang"] = "LỆCH NĂM - cần mở file kiểm tra"
        else:
            r["tinh_trang"] = "KHỚP"
            r["file_word"], r["file_pdf"], bia = chon_file([(f["date_modified"], f["path"]) for f in khop_nam], key)
            r["ghi_chu"] = "; ".join(x for x in (
                f"{len(tim) - len(khop_nam)} file khác năm bị bỏ qua" if len(khop_nam) < len(tim) else "",
                f"bỏ {bia} file bìa/tờ che" if bia else "",
                "chỉ có file bìa" if not (r["file_word"] or r["file_pdf"]) else "") if x)
        kq.append(r)
    return kq


COT_KQ = [("nam", "Năm"), ("dong", "Dòng Excel"), ("tt", "TT"), ("ten", "Tên sách"), ("lan", "Lần XB"),
          ("bbt", "BBT"), ("qdxb", "QĐXB"), ("tinh_trang", "Tình trạng"), ("kieu_khop", "Kiểu khớp"),
          ("so_file", "Số file"), ("file_word", "File Word chọn"), ("file_pdf", "File PDF chọn"),
          ("ten_moi", "Tên file mới"), ("ghi_chu", "Ghi chú"), ("cac_file", "Tất cả file tìm được"),
          ("dich", "Đường dẫn đích")]


COT_TM = [("thu_muc", "Thư mục"), ("ten_tm", "Tên thư mục"), ("so_file", "Số file"), ("loai", "Loại file"),
          ("tinh_trang", "Tình trạng"), ("nam_cuoi", "Năm XB đọc được"), ("sach_khop", "Sách khớp"),
          ("sach_khac", "Sách trùng tên chưa khớp (năm - dòng Excel)"), ("ghi_chu", "Ghi chú")]
TM_KHOP, TM_CHUA, TM_KHONG = "KHỚP", "CHƯA KHỚP - chỉ trùng tên", "KHÔNG KHỚP SÁCH NÀO"
TM_TRUNG = "BẢN TRÙNG - sách đã khớp ở thư mục khác"


def doi_chieu_thu_muc(kq, files, cache=None):
    """Mỗi thư mục trong kho -> có file được chọn cho sách nào (KHỚP), là bản trùng của sách đã khớp nơi khác,
    chỉ trùng tên sách, hay không khớp sách nào. Năm trang cuối lấy từ cache tab 3 (không mở file ở đây)."""
    cache = {} if cache is None else cache
    theo_tm, khop, khac = defaultdict(list), defaultdict(list), defaultdict(list)
    for f in files:
        theo_tm[os.path.dirname(f["path"])].append(f)
    for r in kq:
        nhan = f"{r['nam']} - dòng {r['dong']} - {r['ten']}"
        chon = {os.path.dirname(x) for x in (r["file_word"], r["file_pdf"]) if x}
        for tm in chon:
            khop[tm].append(nhan)
        for tm in {os.path.dirname(x) for _, x in _dong_file(r["cac_file"])} - chon:
            khac[tm].append((nhan, r))
    kq_tm = []
    for tm, fs in sorted(theo_tm.items()):
        doc = [f["path"] for f in fs if f["dinh_dang"] in ("pdf", "doc", "docx") and not la_bia(f["path"])]
        ts = [cache.get(_khoa_cache(x), "chưa đọc") for x in doc]
        cuoi = sorted({y for t in ts if isinstance(t, dict) for y in t["nam"]})
        dau = sorted({y for t in ts if isinstance(t, dict) for y in t.get("dau", ())})
        nam = sorted(set(cuoi) | set(dau))
        doc_duoc, chua = sum(isinstance(t, dict) for t in ts), ts.count("chưa đọc")
        ghi = (["không có file Word/PDF (chỉ ảnh, bìa, zip...)"] if not doc else
               [f"đọc được {doc_duoc}/{len(doc)} file Word/PDF"] + (["còn lại là .doc/scan không có chữ"] if doc_duoc + chua < len(doc) else [])
               + ([f"{chua} file chưa đọc - chạy tab 3"] if chua else [])
               + (["không thấy năm ở trang cuối lẫn 8 trang đầu (PDF scan/bìa, hoặc bản chính nằm trong .zip?)"] if doc_duoc and not nam else []))
        trung = any(r["tinh_trang"] == "KHỚP" and r["nam"] in nam for _, r in khac[tm])
        kq_tm.append(dict(thu_muc=tm, ten_tm=os.path.basename(tm), so_file=len(fs),
                          loai=", ".join(f"{k.upper()} {n}" for k, n in sorted(Counter(f["dinh_dang"] for f in fs).items())),
                          tinh_trang=TM_KHOP if khop[tm] else TM_TRUNG if trung else TM_CHUA if khac[tm] else TM_KHONG,
                          nam_cuoi="; ".join(f"{n}: {', '.join(map(str, ys))}" for n, ys in (("trang cuối", cuoi), ("8 trang đầu", dau)) if ys),
                          sach_khop="\n".join(khop[tm]),
                          sach_khac="\n".join(n for n, _ in khac[tm]), ghi_chu="; ".join(ghi)))
    return kq_tm


def doc_thu_muc(xlsx):
    """Đọc lại sheet THU_MUC_KHO của file doi_chieu (rỗng nếu file cũ không có)."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ten2k, rows = {h: k for k, h in COT_TM}, []
    if "THU_MUC_KHO" in wb.sheetnames:
        it = wb["THU_MUC_KHO"].iter_rows(values_only=True)
        keys = [ten2k.get(h) for h in next(it, ())]
        rows = [{k: ("" if v is None else v) for k, v in zip(keys, row) if k} for row in it]
    wb.close()
    return rows


def ghi_ket_qua(kq, xlsx_goc, thu_muc_ra, thu_muc=()):
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
    if thu_muc:
        ws = wb.create_sheet("THU_MUC_KHO")
        ws.append([h for _, h in COT_TM])
        for r in thu_muc:
            ws.append([_cell(r.get(k, "")) for k, _ in COT_TM])
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
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
            ws.cell(r["dong"], c + 1, _cell("\n".join(x for x in (r["file_word"], r["file_pdf"]) if x)))
            ws.cell(r["dong"], c + 2, r["ten_moi"])
            ws.cell(r["dong"], c + 3, _cell(r.get("dich", "")))
    ra2 = os.path.join(thu_muc_ra, f"{os.path.splitext(os.path.basename(xlsx_goc))[0]}_danh_dau_{tem}.xlsx")
    wb.save(ra2)
    return ra1, ra2


def _cell(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, str) and len(v) > GIOI_HAN_O:  # Excel báo lỗi file nếu ô dài hơn 32.767 ký tự
        giu = v[:GIOI_HAN_O].split("\n")[:-1]
        bo = v.count("\n") + 1 - len(giu)
        return "\n".join(giu or [v[:GIOI_HAN_O]]) + f"\n… cắt bớt {bo} dòng (Excel giới hạn 32.767 ký tự/ô)"
    return v


def doc_doi_chieu(xlsx):
    """Đọc lại file doi_chieu_*.xlsx (cả 2 sheet) -> list dict cùng khoá COT_KQ, sắp theo (năm, dòng)."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ten2k = {h: k for k, h in COT_KQ}
    kq = []
    for ws in (wb[n] for n in ("KHOP", "CAN_XU_LY") if n in wb.sheetnames):
        rows = ws.iter_rows(values_only=True)
        keys = [ten2k.get(h) for h in next(rows, ())]
        for row in rows:
            r = {k: ("" if v is None else v) for k, v in zip(keys, row) if k}
            if r.get("ten"):
                kq.append(r)
    wb.close()
    kq.sort(key=lambda r: (r["nam"], r["dong"]))
    return kq


# ---------- kiểm tra nội dung ----------
def loi_noi_dau(text):
    """Trong 5000 ký tự đầu của 'Lời nói đầu' / 'Lời Nhà xuất bản' / 'Lời giới thiệu', chỉ giữ các đoạn bắt đầu bằng
    'cuốn sách / bộ sách / ...' (dài 600 ký tự) — đủ để dò tên sách, cache nhỏ. '' nếu không có."""
    t = chuan_hoa(text)
    m = re.search(r"\bloi (noi dau|nha xuat ban|gioi thieu|tua|mo dau)\b", t)
    t = t[m.start():m.start() + 5000] if m else ""
    return " | ".join(t[x.start():x.start() + 600] for x in re.finditer(NHAC_SACH, t))


def nam_trang_cuoi(t):
    """Năm 20xx ở trang cuối (text đã chuan_hoa). Có khối thông tin xuất bản ('đăng ký / quyết định / giấy phép
    xuất bản', 'nộp lưu chiểu', cả font lỗi 'noap lou chieau') thì chỉ lấy năm từ đó trở đi, để bỏ năm của luật,
    tài liệu tham khảo...; không có khối này thì không lấy năm nào (trang cuối có thể là mục lục '... (1900 - 2020)')."""
    m = re.search(r"\b(quyet dinh|giay phep|dang ky)( ke hoach)? xuat ban\b|\bl\w{0,2}u chi\w{0,3}u\b", t)
    return sorted({int(y) for y in re.findall(r"\b20\d\d\b", t[m.start():])}) if m else []


def nam_trang_dau(t):
    """Năm xuất bản ở 8 trang đầu (text đã chuan_hoa), lấy theo thứ tự tin cậy, có bậc nào thì dừng ở bậc đó:
    1. khối thông tin XB đặt ở đầu sách (đăng ký / quyết định XB, lưu chiểu), tính đến chữ ISBN;
    2. biên mục CIP '... quốc gia, 2021. - 312tr.' và trang tên sách 'Nhà xuất bản ... Hà Nội - 2021';
    3. chữ ký lời NXB 'tháng 3 năm 2021 Nhà xuất bản' (có thể ký trước năm XB).
    Bỏ qua năm trong nội dung và chú thích tài liệu tham khảo ('Nxb. ..., Hà Nội, 2011, t.7')."""
    nam = lambda x: sorted({int(y) for y in re.findall(r"\b20\d\d\b", x)})
    m = re.search(r"\b(quyet dinh|giay phep|dang ky)( ke hoach)? xuat ban\b|\bl\w{0,2}u chi\w{0,3}u\b", t)
    if m and nam(t[m.start():m.start() + 300].split(" isbn ")[0]):
        return nam(t[m.start():m.start() + 300].split(" isbn ")[0])
    ys = {int(y) for pat in (r"\b(20\d\d) \d+ ?tr\b", r"\bnha xuat ban [a-z ]{0,50}?\bha noi (20\d\d)\b")
          for y in re.findall(pat, t)}
    return sorted(ys) or sorted({int(y) for y in re.findall(r"\bthang \d{1,2} nam (20\d\d) nh\w{0,2} xu", t)})


def khop_nam(t, nam):
    """Ưu tiên trang cuối, không khớp thì xét 8 trang đầu -> 'trang cuối' / '8 trang đầu' / ''."""
    return "trang cuối" if nam in t["nam"] else "8 trang đầu" if nam in t.get("dau", ()) else ""


def cac_nam(t):
    return sorted(set(t["nam"]) | set(t.get("dau", ())))


_WORD = [None, None]  # [Word.Application dùng chung (False = máy không có Word), PID của tiến trình Word đó]
WORD_HET_GIO = 90  # giây; quá thời gian (file hỏng, hộp thoại treo...) thì tắt Word, bỏ qua file đó
_KHONG_CUA_SO = 0x08000000  # CREATE_NO_WINDOW: không bật cửa sổ đen khi gọi tasklist/taskkill


def _pid_word():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, creationflags=_KHONG_CUA_SO).stdout
    return {int(x) for x in re.findall(r'"WINWORD\.EXE","(\d+)"', out, re.I)}


def _word_app():
    """Word.Application dùng chung (mở lần đầu cần); None nếu máy không có Word."""
    if _WORD[0] is None:
        try:
            import pythoncom, win32com.client
            pythoncom.CoInitialize()
            truoc = _pid_word()
            _WORD[0] = win32com.client.DispatchEx("Word.Application")
            _WORD[0].Visible, _WORD[0].DisplayAlerts = False, 0
            _WORD[1] = max(_pid_word() - truoc, default=None)  # Word riêng của tool, không đụng Word người dùng đang mở
        except Exception:
            _WORD[0] = False
    return _WORD[0] or None


def _giet_word():
    if _WORD[1]:
        subprocess.run(["taskkill", "/PID", str(_WORD[1]), "/F"], capture_output=True, creationflags=_KHONG_CUA_SO)


def _doc_word(path):
    """.doc -> (KY_DAU ký tự đầu, KY_CUOI ký tự cuối) qua Microsoft Word, mở chỉ đọc; None nếu không đọc được.
    Lấy theo số ký tự chứ không đếm trang: đếm trang bắt Word dàn trang cả cuốn, chậm ~20 lần."""
    with open(duong_dan_dai(path), "rb") as f:
        if f.read(8) != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # không phải file Word 97-2003 thật
            return None
    if not _word_app():
        return None
    mo, tam = os.path.abspath(path), None
    if len(mo) > 200:  # Word không mở được đường dẫn dài hơn ~255 ký tự (trả về rỗng, không báo lỗi) -> mở bản copy ngắn
        fd, tam = tempfile.mkstemp(suffix=".doc", prefix="sxb_"); os.close(fd)
        shutil.copyfile(duong_dan_dai(path), tam); mo = tam
    bi_giet = []
    hen = threading.Timer(WORD_HET_GIO, lambda: (bi_giet.append(1), _giet_word())); hen.start()
    d = None
    try:
        # mật khẩu giả: file có mật khẩu báo lỗi ngay thay vì hiện hộp thoại chờ người nhập
        d = _WORD[0].Documents.Open(mo, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
                                    PasswordDocument="sxb", Visible=False, NoEncodingDialog=True)
        if d is None:
            raise OSError("Word không mở được file")
        het = d.Content.End
        return d.Range(0, min(het, KY_DAU)).Text, d.Range(max(0, het - KY_CUOI), het).Text
    except Exception:
        try:
            _WORD[0].Version  # Word còn sống? (có thể bị tắt vì quá giờ hoặc tự chết)
        except Exception:
            bi_giet.append(1)
        raise
    finally:
        hen.cancel()
        try:
            d is not None and d.Close(False)
        except Exception:
            pass
        if bi_giet:  # Word đã bị tắt (quá giờ) hoặc tự chết -> lần sau mở Word mới
            _WORD[0] = _WORD[1] = None
        if tam:
            try:
                os.remove(tam)
            except OSError:
                pass


def _docx_text(path):
    """Chữ của .docx đọc thẳng từ word/document.xml (nhanh hơn python-docx nhiều lần), mỗi đoạn một dòng."""
    with zipfile.ZipFile(path) as z:
        x = z.read("word/document.xml").decode("utf-8")
    x = re.sub(r"<w:(?:tab|br|cr)\b[^>]*/>", "<w:t> </w:t>", x)  # tab/xuống dòng giữa các chữ -> khoảng trắng
    doan = (html.unescape("".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", d))) for d in x.split("</w:p>"))
    return "\n".join(d for d in doan if d.strip())


def dong_word():
    """Tắt Word sau khi chạy tab 3."""
    if _WORD[0]:
        try:
            _WORD[0].Quit(False)
        except Exception:
            _giet_word()
    _WORD[0] = _WORD[1] = None


def trich(path, so_trang=8):
    """-> {'nam': năm XB ở trang cuối, 'dau': năm XB ở 8 trang đầu, 'lnd': đoạn 'cuốn sách ...' trong lời nói đầu}.
    PDF theo trang; .docx/.doc theo số ký tự KY_DAU/KY_CUOI (.doc đọc qua Microsoft Word). None nếu không đọc được."""
    ext = os.path.splitext(path)[1].lower()
    p = duong_dan_dai(path)
    try:
        if ext == ".pdf":
            import fitz
            with fitz.open(p) as d:
                dau = " ".join(d[i].get_text() for i in range(min(so_trang, len(d))))
                cuoi = " ".join(d[i].get_text() for i in range(max(0, len(d) - 2), len(d)))
        elif ext == ".docx":
            t = _docx_text(p)
            dau, cuoi = t[:KY_DAU], t[-KY_CUOI:]
        elif ext == ".doc":
            r = _doc_word(path)
            if r is None:
                return None
            dau, cuoi = r
        else:
            return None
    except Exception:
        return None
    return dict(nam=nam_trang_cuoi(chuan_hoa(cuoi)), dau=nam_trang_dau(chuan_hoa(dau)), lnd=loi_noi_dau(dau))


class Cache:
    """Cache kết quả đọc file trong SQLite trên đĩa: không chiếm RAM, ghi từng mục, chạy lại dùng tiếp."""
    def __init__(self, path):
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.execute("create table if not exists c (k text primary key, v text)")

    def get(self, k, mac_dinh=None):
        row = self.db.execute("select v from c where k = ?", (k,)).fetchone()
        return json.loads(row[0]) if row else mac_dinh

    def __contains__(self, k):
        return self.get(k, KeyError) is not KeyError

    def __getitem__(self, k):
        v = self.get(k, KeyError)
        if v is KeyError:
            raise KeyError(k)
        return v

    def __setitem__(self, k, v):
        self.db.execute("insert or replace into c values (?, ?)", (k, json.dumps(v, ensure_ascii=False)))

    def __len__(self):
        return self.db.execute("select count(*) from c").fetchone()[0]

    def dong(self):
        self.db.commit(); self.db.close()


def doc_cache(path):
    return Cache(path)


def ghi_cache(cache, path):
    """Lưu cache xuống đĩa (Cache: commit; dict thường: chép sang file SQLite)."""
    if isinstance(cache, Cache):
        cache.db.commit()
    else:
        c = Cache(path)
        for k, v in cache.items():
            c[k] = v
        c.dong()


def _khoa_cache(path):
    """Khoá cache theo (path, mtime); None nếu file không còn. |8: đổi khi cách trích thay đổi."""
    try:
        return f"{path}|{int(os.path.getmtime(duong_dan_dai(path)))}|8"
    except OSError:
        return None


def trich_cache(cache, path):
    """Như trich() nhưng nhớ theo (path, mtime) trong cache dict."""
    k = _khoa_cache(path)
    if k is None:
        return None
    v = cache.get(k, KeyError)
    if v is KeyError:
        v = cache[k] = trich(path)
    return v


def _dong_file(cac_file):
    """'Tất cả file tìm được' -> [(date_modified, path)]"""
    return [(l[:19], l[21:]) for l in str(cac_file or "").splitlines() if len(l) > 21 and l[:4].isdigit()]


def nam_ten_file(path, key=""):
    """Năm 20xx trong tên file ('... (TB12-4-2021).pdf', '... 2021.pdf'), trừ năm nằm trong chính tên sách."""
    ten = os.path.splitext(os.path.basename(path))[0]
    return {int(y) for y in re.findall(r"(?<!\d)20\d\d(?!\d)", ten)} - {int(y) for y in re.findall(r"\b20\d\d\b", key)}


def _khong_lech(cache, p, nam, key=""):
    """File cùng thư mục được nhận khi không ghi năm khác: cả trong nội dung lẫn trong tên file."""
    fy, t = nam_ten_file(p, key), trich_cache(cache, p)
    return (not fy or nam in fy) and (not t or not cac_nam(t) or bool(khop_nam(t, nam)))


def _chon(r, ok, ghi_chu, cung_tm=()):
    """ok: [(date_modified, path)] đã xác nhận năm -> ghi KHỚP vào r. File trong cung_tm nằm cùng thư mục
    với file đã xác nhận cũng được nhận (vd .doc cạnh PDF). False nếu ok rỗng."""
    if not ok:
        return False
    tm, key = {os.path.dirname(p) for _, p in ok}, chuan_hoa(r["ten"])
    w, p, _ = chon_file(ok, key)
    w2, p2, _ = chon_file([x for x in cung_tm if os.path.dirname(x[1]) in tm], key)
    w, p = w or w2, p or p2
    r.update(tinh_trang="KHỚP", file_word=w, file_pdf=p, ghi_chu=ghi_chu if w or p else ghi_chu + "; chỉ có file bìa")
    return True


def xu_ly_lech_nam(r, cache):
    """Dòng LỆCH NĂM: mở từng file trong 'Tất cả file tìm được', nhận file có năm cột Năm ở trang cuối,
    không có thì ở 8 trang đầu (cùng các file khác trong thư mục đó)."""
    ok, ghi, thay, nguon, ds = [], [], set(), set(), _dong_file(r["cac_file"])
    for d, p in ds:
        if not p.lower().endswith((".pdf", ".doc", ".docx")):  # ảnh bìa... không có năm để đọc
            continue
        t = trich_cache(cache, p)
        if t is None:
            ghi.append(f"bỏ qua (không đọc được): {os.path.basename(p)}")
            continue
        thay.update(cac_nam(t))
        if khop_nam(t, r["nam"]):
            ok.append((d, p)); nguon.add(khop_nam(t, r["nam"]))
    nhan = f"khớp theo nội dung ({', '.join(sorted(nguon, reverse=True))})"
    if not _chon(r, ok, nhan, [x for x in ds if _khong_lech(cache, x[1], r["nam"], chuan_hoa(r["ten"]))]):
        ghi.insert(0, f"không thấy năm {r['nam']} ở trang cuối/8 trang đầu (thấy: {', '.join(map(str, sorted(thay))) or '-'})")
        r["ghi_chu"] = "; ".join(ghi)


def loc_files(files, thu_muc="", nam=""):
    """Giới hạn file kho: trong thư mục con, Date Modified trong năm/khoảng năm ('2015' hoặc '2015-2017')."""
    tm = os.path.normcase(os.path.abspath(thu_muc)) if thu_muc.strip() else ""
    ns = [int(x) for x in re.findall(r"\d{4}", nam)]
    lo, hi = (ns[0], ns[-1]) if ns else (0, 9999)
    return [f for f in files if lo <= f["nam"] <= hi
            and (not tm or os.path.normcase(f["path"]).startswith(tm))]


NHAC_SACH = r"\b(?:cuon sach|cuon|bo sach|tap sach|an pham)(?: (?:chuyen khao|tham khao|nay))? "


def nhac_ten_sach(lnd, key):
    """Lời nói đầu giới thiệu đúng cuốn này: '... xuất bản cuốn sách <tên sách>' — không chỉ nhắc qua tên sách khác."""
    return re.search(NHAC_SACH + re.escape(key) + r"\b", lnd) is not None


def xu_ly_khong_tim_thay(rows, files, cache, progress=None):
    """Dòng KHÔNG TÌM THẤY: tìm trong lời nói đầu câu giới thiệu 'cuốn sách <tên sách>', rồi kiểm tra năm trang cuối."""
    theo_prefix = defaultdict(list)
    for r in rows:
        r["_key"], r["_tim"] = chuan_hoa(r["ten"]), {}
        theo_prefix[" ".join(r["_key"].split()[:3])].append(r)
    theo_tm = defaultdict(list)
    for f in files:
        theo_tm[os.path.dirname(f["path"])].append((f["date_modified"], f["path"]))
    for i, f in enumerate(files, 1):
        _dung()
        t = f["dinh_dang"] in DOC_DUOC and trich_cache(cache, f["path"])
        if t and t["lnd"]:
            w, lnd = t["lnd"].split(), f" {t['lnd']} "
            for j in range(len(w)):  # ponytail: tra từng cụm 1-3 từ của lời nói đầu vào chỉ mục tên sách
                for n in (1, 2, 3):
                    for r in theo_prefix.get(" ".join(w[j:j + n]), ()):
                        if nhac_ten_sach(lnd, r["_key"]):
                            r["_tim"][f["path"]] = (f, t)
        if progress and i % 200 == 0:
            progress(i)
    for r in rows:
        tim, _ = r.pop("_tim"), r.pop("_key")
        if not tim:
            r["tinh_trang"] = "KHÔNG TÌM ĐƯỢC"
            continue
        r.update(so_file=len(tim), kieu_khop="lời nói đầu",
                 cac_file=_ds_file((f["date_modified"], f["path"]) for f, _ in tim.values()))
        ok = [(f["date_modified"], f["path"]) for f, t in tim.values() if khop_nam(t, r["nam"])]
        cung = [x for _, p in ok for x in theo_tm[os.path.dirname(p)] if _khong_lech(cache, x[1], r["nam"], chuan_hoa(r["ten"]))]
        if not _chon(r, ok, "khớp theo nội dung (lời nói đầu)", cung):
            r["tinh_trang"] = "LỆCH NĂM - tìm theo lời nói đầu"
            r["ghi_chu"] = f"không thấy năm {r['nam']} ở trang cuối/8 trang đầu"


# ---------- đổi tên & lưu ----------
def ke_hoach_luu(kq, thu_muc_dich):
    """-> [(r, nguồn, đích)] cho dòng KHỚP. Mỗi sách một thư mục:
    <đích>/<Năm>/thang <MM>/<Tên file mới>/<Tên file mới><đuôi>; thư mục trùng tên thêm _2, _3."""
    da_dung, kh = set(), []
    for r in kq:
        nguon = [x for x in (r["file_word"], r["file_pdf"]) if x]
        if r["tinh_trang"] != "KHỚP" or not nguon:
            continue
        goc = os.path.join(thu_muc_dich, str(r["nam"]), f"thang {thang_qdxb(r['qdxb'])}", r["ten_moi"])
        tm, i = goc, 1
        # thư mục đã có: dùng lại nếu là của chính sách này (chạy lại sau khi dừng giữa chừng), khác thì sang _2, _3
        while tm.lower() in da_dung or (os.path.exists(duong_dan_dai(tm)) and not all(
                da_co(src, os.path.join(tm, os.path.basename(tm) + os.path.splitext(src)[1].lower()), True) for src in nguon)):
            i += 1
            tm = f"{goc}_{i}"
        da_dung.add(tm.lower())
        ten = os.path.basename(tm)
        kh += [(r, src, os.path.join(tm, ten + os.path.splitext(src)[1].lower())) for src in nguon]
    return kh


def da_co(src, dst, chua_co_cung_duoc=False):
    """File đích đã là bản của nguồn (cùng dung lượng, hoặc nguồn đã bị di chuyển đi).
    chua_co_cung_duoc: đích chưa có cũng tính là được (thư mục copy dở)."""
    s, d = duong_dan_dai(src), duong_dan_dai(dst)
    if not os.path.exists(d):
        return chua_co_cung_duoc
    return not os.path.exists(s) or os.path.getsize(s) == os.path.getsize(d)


def _con_trong(p):
    while not os.path.exists(duong_dan_dai(p)) and os.path.dirname(p) != p:
        p = os.path.dirname(p)
    return shutil.disk_usage(duong_dan_dai(p)).free


def luu_file(kh, di_chuyen=False, progress=None):
    """Copy (hoặc move) theo kế hoạch, ghi đường dẫn đích vào r['dich']. File đã có thì bỏ qua (chạy lại được);
    lỗi ở file nào thì ghi vào Ghi chú rồi làm tiếp. -> (số file đã chép, số đã có sẵn, [lỗi])."""
    can = [(r, src, p) for r, src, p in kh if not da_co(src, p)]
    cung_o = lambda src, p: os.path.splitdrive(os.path.abspath(src))[0].lower() == os.path.splitdrive(os.path.abspath(p))[0].lower()
    byte = sum(os.path.getsize(duong_dan_dai(src)) for _, src, p in can
               if os.path.exists(duong_dan_dai(src)) and not (di_chuyen and cung_o(src, p)))
    if can and byte > _con_trong(can[0][2]):
        raise ValueError(f"Ổ đích không đủ chỗ: cần {byte / 2**30:.1f} GB, còn {_con_trong(can[0][2]) / 2**30:.1f} GB")
    chep, co_san, loi = 0, 0, []
    for i, (r, src, p) in enumerate(kh, 1):
        if DUNG.is_set():  # dừng: phần đã chép giữ nguyên, chạy lại sẽ làm tiếp
            break
        try:
            if da_co(src, p):
                co_san += 1
            else:
                os.makedirs(duong_dan_dai(os.path.dirname(p)), exist_ok=True)
                tam = p + ".dang_chep"  # chép ra tên tạm rồi đổi tên: dừng giữa chừng không để lại file dở mang tên thật
                (shutil.move if di_chuyen else shutil.copy2)(duong_dan_dai(src), duong_dan_dai(tam))
                os.replace(duong_dan_dai(tam), duong_dan_dai(p))
                chep += 1
            r["dich"] = "\n".join(x for x in (r.get("dich", ""), p) if x)
        except OSError as e:
            loi.append(f"{src}: {e}")
            r["ghi_chu"] = "; ".join(x for x in (r.get("ghi_chu", ""), f"LỖI lưu {os.path.basename(src)}: {e.strerror or e}") if x)
        if progress and i % 100 == 0:
            progress(i)
    return chep, co_san, loi


def moi_nhat(*mau):
    """File mới nhất (theo giờ sửa) khớp các mẫu glob; '' nếu chưa có."""
    fs = [f for m in mau for f in glob.glob(m)]
    return max(fs, key=os.path.getmtime) if fs else ""


# ---------- giao diện ----------
def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    # kết quả gom vào ketqua/ cạnh file exe (hoặc cạnh sxb_tool.py), mỗi bước một thư mục
    goc = os.path.join(os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)), "ketqua")
    TM = {k: os.path.join(goc, k) for k in ("1_quet_kho", "2_doi_chieu", "3_kiem_tra", "4_doi_ten", "sach_da_luu")}
    for d in TM.values():
        os.makedirs(d, exist_ok=True)
    # ô đầu vào mặc định = file kết quả mới nhất của bước trước
    KHO = (os.path.join(TM["1_quet_kho"], "kho_*.xlsx"), os.path.join(TM["1_quet_kho"], "kho_*.csv"))
    DC2, DC3 = (os.path.join(TM[k], "doi_chieu_*.xlsx") for k in ("2_doi_chieu", "3_kiem_tra"))

    root = tk.Tk()
    root.title("Tool SXB - quét kho, đối chiếu, kiểm tra nội dung, đổi tên")
    root.geometry("820x580")
    thanh = ttk.Frame(root); thanh.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
    b_dung = ttk.Button(thanh, text="■ Dừng", state="disabled", command=lambda: (DUNG.set(), b_dung.config(state="disabled")))
    b_dung.pack(side="right")
    ttk.Label(thanh, text="Dừng sau file đang xử lý; chạy lại sẽ làm tiếp phần còn lại.").pack(side="right", padx=8)
    nb = ttk.Notebook(root); nb.pack(fill="both", expand=True, padx=8, pady=8)
    ui = lambda f, *a: root.after(0, f, *a)  # mọi thay đổi giao diện từ luồng chạy nền đi qua luồng chính

    def hoi(msg):
        """Hỏi Có/Không từ luồng chạy nền: hộp thoại hiện ở luồng chính, luồng nền chờ câu trả lời."""
        tl, xong = [], threading.Event()
        ui(lambda: (tl.append(messagebox.askyesno("Xác nhận", msg, parent=root)), xong.set()))
        xong.wait()
        return tl[0]
    chon_xlsx = lambda: filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    chon_csv = lambda: filedialog.askopenfilename(filetypes=[("File kho", "*.xlsx *.csv")])

    def hang(parent, r, nhan, var, chon=None):
        ttk.Label(parent, text=nhan, width=20).grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, width=70).grid(row=r, column=1, sticky="we")
        if chon:
            ttk.Button(parent, text="...", width=3, command=lambda: var.set(chon() or var.get())).grid(row=r, column=2)

    dang_chay = [0]

    def chay(nut, ham, txt=None):
        nut["state"] = "disabled"; DUNG.clear(); dang_chay[0] += 1; b_dung.config(state="normal")
        def bao(msg, nhan="LỖI"):
            if isinstance(txt, tk.Text):
                txt.insert("end", f"{nhan}: {msg}\n"); txt.see("end")
            elif txt is not None:
                txt.config(text=f"{nhan}: {msg}")
            if nhan == "LỖI":
                messagebox.showerror("Lỗi", msg, parent=root)
        def xong():
            nut.config(state="normal"); dang_chay[0] -= 1
            if not dang_chay[0]:
                b_dung.config(state="disabled")
        def run():
            try:
                ham()
            except DaDung as e:
                msg = str(e)
                ui(lambda: bao(msg, "ĐÃ DỪNG"))
            except Exception as e:
                msg = str(e) or type(e).__name__
                ui(lambda: bao(msg))
            finally:
                ui(xong)
        threading.Thread(target=run, daemon=True).start()

    def logger(txt):
        return lambda s: root.after(0, lambda: (txt.insert("end", s + "\n"), txt.see("end")))

    # --- tab 1: quét ---
    t1 = ttk.Frame(nb, padding=10); nb.add(t1, text="1. Quét kho → Excel")
    v_tm, v_ext, v_csv, tu_dong = tk.StringVar(), tk.StringVar(value=EXT_MAC_DINH), tk.StringVar(), []
    hang(t1, 0, "Thư mục kho", v_tm, filedialog.askdirectory)
    hang(t1, 1, "Định dạng (cách)", v_ext)
    hang(t1, 2, "File Excel xuất", v_csv,
         lambda: filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]))
    tt1 = ttk.Label(t1, text="", wraplength=700); tt1.grid(row=4, column=0, columnspan=3, sticky="w", pady=4)
    tt1l = tk.Text(t1, height=10, wrap="word"); tt1l.grid(row=5, column=0, columnspan=3, sticky="we", pady=4)

    def quet_gui():
        if not os.path.isdir(v_tm.get()): raise ValueError("Chưa chọn thư mục kho")
        ra = v_csv.get()
        if not ra or ra in tu_dong:  # tên tự đặt -> mỗi lần quét một file mới, không ghi đè
            ra = os.path.join(TM["1_quet_kho"], f"kho_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"); tu_dong.append(ra); ui(v_csv.set, ra)
        log, bc, t0 = logger(tt1l), {}, time.time()
        def prog(n, root_dir): ui(lambda: tt1.config(text=f"Đã quét {n} file ({_gio(time.time() - t0)})... {root_dir[-80:]}"))
        try:
            n = quet(v_tm.get(), v_ext.get().split(), ra, prog, bc)
        except DaDung:
            raise DaDung(f"Đã dừng sau {sum(bc['dinh_dang'].values())} file - chưa lưu file kho vì quét chưa xong.")
        ui(v_csv2.set, ra); ui(v_csv3.set, ra); ui(lambda: tt1.config(text=""))
        log(f"Xong: {n} file trong {_gio(time.time() - t0)} → {ra}")
        log("  " + ", ".join(f"{k} {v}" for k, v in bc["dinh_dang"].most_common()))
        anh = bc["dinh_dang"]["JPG"] + bc["dinh_dang"]["JPEG"]
        if n and anh / n > 0.5:
            log(f"  CHÚ Ý: ảnh JPG chiếm {anh / n:.0%} ({anh} file). Nếu đây là ảnh scan từng trang, nên bỏ .jpg .jpeg ở ô "
                "Định dạng rồi quét lại: file kho nhỏ hơn, các bước sau nhanh và ít RAM hơn.")
        if bc["loi"]:
            ds_loi = ra + ".loi.txt"
            with open(ds_loi, "w", encoding="utf-8-sig") as f:
                f.write("\n".join(bc["loi"]))
            log(f"  CHÚ Ý: {len(bc['loi'])} thư mục/file KHÔNG đọc được (không có quyền, đường dẫn hỏng...) - "
                f"file trong đó không có trong kho. Danh sách: {ds_loi}")
            for x in bc["loi"][:10]:
                log(f"    {x}")
    b1 = ttk.Button(t1, text="Quét", command=lambda: chay(b1, quet_gui, tt1l)); b1.grid(row=3, column=1, sticky="w", pady=6)

    # --- tab 2: đối chiếu ---
    t2 = ttk.Frame(nb, padding=10); nb.add(t2, text="2. Đối chiếu Excel")
    v_xl, v_csv2, v_ra = tk.StringVar(), tk.StringVar(value=moi_nhat(*KHO)), tk.StringVar(value=TM["2_doi_chieu"])
    hang(t2, 0, "File Excel tổng hợp", v_xl, chon_xlsx)
    hang(t2, 1, "File kho (xlsx)", v_csv2, lambda: chon_csv() or v_csv.get())
    hang(t2, 2, "Thư mục kết quả", v_ra, filedialog.askdirectory)
    tt2 = tk.Text(t2, height=9, wrap="word"); tt2.grid(row=4, column=0, columnspan=3, sticky="we", pady=8)

    def doi_chieu_gui():
        csvp = v_csv2.get() or v_csv.get()
        if not os.path.isfile(v_xl.get()): raise ValueError("Chưa chọn file Excel")
        if not os.path.isfile(csvp): raise ValueError("Chưa chọn file kho")
        log = logger(tt2)
        log("Đọc Excel..."); sach = doc_excel(v_xl.get()); log(f"  {len(sach)} đầu sách")
        log("Đọc file kho..."); files = doc_kho(csvp); log(f"  {len(files)} file")
        kq = doi_chieu(sach, files)
        dem = defaultdict(int)
        for r in kq: dem[r["tinh_trang"]] += 1
        for k, v in dem.items(): log(f"  {k}: {v}")
        c = doc_cache(csvp + ".cache.sqlite"); tm = doi_chieu_thu_muc(kq, files, c); c.dong()
        for k, v in Counter(x["tinh_trang"] for x in tm).items(): log(f"  Thư mục kho - {k}: {v}")
        ra1, ra2 = ghi_ket_qua(kq, v_xl.get(), v_ra.get(), tm)
        ui(v_dc.set, ra1); ui(v_dc4.set, ra1)
        log(f"Kết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b2 = ttk.Button(t2, text="Đối chiếu", command=lambda: chay(b2, doi_chieu_gui, tt2)); b2.grid(row=3, column=1, sticky="w", pady=6)

    # --- tab 3: kiểm tra nội dung ---
    t3 = ttk.Frame(nb, padding=10); nb.add(t3, text="3. Kiểm tra nội dung")
    v_dc, v_xl3, v_ra3 = tk.StringVar(value=moi_nhat(DC2)), tk.StringVar(), tk.StringVar(value=TM["3_kiem_tra"])
    v_csv3 = tk.StringVar(value=moi_nhat(*KHO))
    v_tmc, v_nam3 = tk.StringVar(), tk.StringVar()
    v_lech, v_ktt = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    hang(t3, 0, "File doi_chieu_*.xlsx", v_dc, chon_xlsx)
    hang(t3, 1, "File Excel tổng hợp", v_xl3, chon_xlsx)
    hang(t3, 2, "File kho (xlsx)", v_csv3, chon_csv)
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
        for p, m in ((v_dc.get(), "file doi_chieu"), (xl, "file Excel tổng hợp"), (csvp, "file kho")):
            if not os.path.isfile(p): raise ValueError(f"Chưa chọn {m}")
        log = logger(tt3)
        kq = doc_doi_chieu(v_dc.get()); files = doc_kho(csvp)
        cache_p = csvp + ".cache.sqlite"; cache = doc_cache(cache_p)
        log(f"{len(kq)} dòng, {len(files)} file kho, cache {len(cache)} file")
        try:
            if v_lech.get():
                lech = [r for r in kq if str(r["tinh_trang"]).startswith("LỆCH NĂM")]
                log(f"LỆCH NĂM: {len(lech)} dòng...")
                t0 = time.time()
                for i, r in enumerate(lech, 1):
                    if DUNG.is_set():
                        break
                    xu_ly_lech_nam(r, cache)
                    if i % 50 == 0:
                        log(f"  {i}/{len(lech)} - còn khoảng {_gio((time.time() - t0) / i * (len(lech) - i))}"); ghi_cache(cache, cache_p)
            if v_ktt.get() and not DUNG.is_set():
                ktt = [r for r in kq if str(r["tinh_trang"]).startswith("KHÔNG TÌM")]
                da_khop = {os.path.dirname(x) for r in kq if r["tinh_trang"] == "KHỚP" for x in (r["file_word"], r["file_pdf"]) if x}
                ds = [f for f in loc_files(files, v_tmc.get(), v_nam3.get()) if os.path.dirname(f["path"]) not in da_khop]
                n_doc = sum(f["dinh_dang"] in DOC_DUOC for f in ds)
                uoc = n_doc * GIAY_MOI_FILE
                log(f"KHÔNG TÌM THẤY: {len(ktt)} dòng, quét lời nói đầu {n_doc} file PDF/DOCX, ước tính lần đầu ~{_gio(uoc)} "
                    "(file đã có trong cache nhanh hơn nhiều)")
                if uoc > 1800 and not hoi(
                        f"Quét lời nói đầu {n_doc} file PDF/DOCX ước tính mất khoảng {_gio(uoc)} (ổ cục bộ; ổ mạng có thể lâu hơn).\n\n"
                        "Nên giới hạn theo thư mục con hoặc năm Date Modified ở tab 3 rồi chạy từng phần.\n\nVẫn chạy toàn bộ?"):
                    log("  Bỏ qua bước quét lời nói đầu.")
                else:
                    t0 = time.time()
                    def prog(i):
                        log(f"  {i}/{len(ds)} file - còn khoảng {_gio((time.time() - t0) / i * (len(ds) - i))}")
                        if i % 1000 == 0: ghi_cache(cache, cache_p)
                    try:
                        xu_ly_khong_tim_thay(ktt, ds, cache, prog)
                    except DaDung:
                        pass
        finally:
            ghi_cache(cache, cache_p); dong_word()
        if DUNG.is_set():
            log("ĐÃ DỪNG theo yêu cầu - ghi kết quả phần đã làm; chạy lại sẽ dùng cache, làm tiếp nhanh.")
        DUNG.clear()  # ghi kết quả cho xong dù đã bấm Dừng
        dem = defaultdict(int)
        for r in kq: dem[r["tinh_trang"]] += 1
        for k, v in dem.items(): log(f"  {k}: {v}")
        tm = doi_chieu_thu_muc(kq, files, cache); cache.dong()
        for k, v in Counter(x["tinh_trang"] for x in tm).items(): log(f"  Thư mục kho - {k}: {v}")
        ra1, ra2 = ghi_ket_qua(kq, xl, ra, tm)
        ui(v_dc4.set, ra1)
        log(f"Kết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b3 = ttk.Button(t3, text="Kiểm tra", command=lambda: chay(b3, kiem_tra_gui, tt3)); b3.grid(row=8, column=1, sticky="w", pady=6)

    # --- tab 4: đổi tên & lưu ---
    t4 = ttk.Frame(nb, padding=10); nb.add(t4, text="4. Đổi tên & lưu")
    v_dc4, v_xl4, v_move = tk.StringVar(value=moi_nhat(DC2, DC3)), tk.StringVar(), tk.BooleanVar()
    v_dich, v_ra4 = tk.StringVar(value=TM["sach_da_luu"]), tk.StringVar(value=TM["4_doi_ten"])
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
            log("\n".join(f"{s}\n   → {p}{'   (đã có - bỏ qua)' if da_co(s, p) else ''}" for _, s, p in kh)
                + "\n(chạy thử - chưa copy)")
            return
        t0 = time.time()
        chep, co_san, loi = luu_file(kh, v_move.get(), lambda i: log(f"  {i}/{len(kh)} - còn khoảng {_gio((time.time() - t0) / i * (len(kh) - i))}"))
        if DUNG.is_set():
            log(f"ĐÃ DỪNG theo yêu cầu sau {chep + co_san + len(loi)}/{len(kh)} file - chạy lại sẽ bỏ qua file đã chép, làm tiếp phần còn lại.")
        ra1, ra2 = ghi_ket_qua(kq, xl, ra, doc_thu_muc(v_dc4.get()))
        log(f"Xong: chép {chep} file, {co_san} file đã có sẵn (bỏ qua), {len(loi)} lỗi."
            + "".join(f"\n  LỖI {x}" for x in loi[:50]) + f"\nKết quả: {ra1}\nExcel đánh dấu: {ra2}")
    b4a = ttk.Button(t4, text="Chạy thử", command=lambda: chay(b4a, lambda: luu_gui(True), tt4)); b4a.grid(row=5, column=1, sticky="w", pady=6)
    b4b = ttk.Button(t4, text="Thực hiện", command=lambda: chay(b4b, lambda: luu_gui(False), tt4)); b4b.grid(row=5, column=1, sticky="e", pady=6)

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
    dai = ten_file_moi("Cuốn sách " + "rất dài " * 40, "01/2021", "PL", "TB")
    assert dai.endswith("_0121_R_PL_T01") and len(dai) <= DAI_TEN + 15 and "__" not in dai and dai.startswith("cuon_sach_rat_dai")
    assert thang_qdxb("05/2023") == "05" and thang_qdxb("19/01/2007") == "01" and thang_qdxb(None) == "00"
    # quét -> xlsx có định dạng -> đọc lại ra dict nội bộ
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "Sach mau.PDF"); open(p, "wb").write(b"x" * 2048)
        t = dt.datetime(2022, 3, 5, 14, 30).timestamp(); os.utime(p, (t, t))
        kho, bc = os.path.join(td, "kho.xlsx"), {}
        assert quet(td, [".pdf"], kho, bao_cao=bc) == 1 and bc["dinh_dang"] == {"PDF": 1} and bc["loi"] == []
        bc2 = {}; assert quet(os.path.join(td, "khong_co"), [".pdf"], os.path.join(td, "k2.xlsx"), bao_cao=bc2) == 0
        assert len(bc2["loi"]) == 1 and "khong_co" in bc2["loi"][0] and not bc2["loi"][0].startswith("\\\\?")
        DUNG.set()
        try:
            quet(td, [".pdf"], os.path.join(td, "k3.xlsx")); assert False
        except DaDung:
            from openpyxl.worksheet._writer import ALL_TEMP_FILES
            assert not os.path.exists(os.path.join(td, "k3.xlsx")) and not ALL_TEMP_FILES
        finally:
            DUNG.clear()
        ws = openpyxl.load_workbook(kho)["Kho"]
        assert [c.value for c in ws[1]] == COT_KHO and ws["A1"].font.bold and ws.freeze_panes == "A2"
        assert ws.column_dimensions["B"].width == 70 and ws["B2"].alignment.wrap_text and ws["G2"].alignment.wrap_text
        assert [c.value for c in ws[2]] == [1, "Sach mau", "PDF", dt.datetime(2022, 3, 5, 14, 30), 2022, 2, p]
        f, = doc_kho(kho)
        assert (f["dinh_dang"], f["date_modified"], f["nam"], f["path"]) == ("pdf", "2022-03-05 14:30:00", 2022, p)
        cu = os.path.join(td, "cu.csv")  # csv cũ utf-8 phẩy vẫn đọc được
        open(cu, "w", encoding="utf-8-sig", newline="").write("ten_file,dinh_dang,date_modified,path\r\nx,docx,2020-01-02 03:04:05,x.docx\r\n")
        assert doc_kho(cu)[0]["nam"] == 2020
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

    # --- ô Excel 32.767 ký tự, thứ tự danh sách file, định dạng thời gian ---
    dai = _ds_file((f"2021-01-01 00:00:00", f"C:\\kho\\{i:06d}" + (".jpg" if i % 2 else ".pdf")) for i in range(2000))
    o = _cell(dai)
    assert len(dai) > 40000 and len(o) <= 32767 and o.endswith("(Excel giới hạn 32.767 ký tự/ô)")
    assert all(p.endswith(".pdf") for _, p in _dong_file(o)[:1000]) and len(_dong_file(o)) == o.count("\n")  # PDF trước, dòng ghi chú bị bỏ
    assert _cell("ngắn") == "ngắn" and _gio(30) == "30 giây" and _gio(600) == "10 phút" and _gio(9000) == "2,5 giờ"

    # --- file kết quả mới nhất ---
    with tempfile.TemporaryDirectory() as td:
        assert moi_nhat(os.path.join(td, "doi_chieu_*.xlsx")) == ""
        for i, ten in enumerate(("doi_chieu_2.xlsx", "doi_chieu_1.xlsx", "khac.xlsx")):
            f = os.path.join(td, ten); open(f, "w").close(); os.utime(f, (1000 + i, 1000 + i))
        assert os.path.basename(moi_nhat(os.path.join(td, "doi_chieu_*.xlsx"), os.path.join(td, "x_*"))) == "doi_chieu_1.xlsx"

    # --- chọn nhầm file: báo lỗi rõ ---
    with tempfile.TemporaryDirectory() as td:
        sai = os.path.join(td, "sai.xlsx"); wb = openpyxl.Workbook(); wb.active.append(["TT", "Tên sách"]); wb.save(sai)
        for ham, chu in ((doc_excel, "không phải file Excel tổng hợp"), (doc_kho, "không phải file kho")):
            try: ham(sai); assert False
            except ValueError as e: assert chu in str(e)

    # --- kho xếp mỗi sách một thư mục ---
    tm = lambda *x: dict(zip(("ten_file", "dinh_dang", "date_modified", "path", "thu_muc"), x))
    kho = [tm("BAN THAO (Thu Hoa)", "doc", "2021-03-01 00:00:00", r"K\Đèn xanh, đèn đỏ\BAN THAO (Thu Hoa).doc", "Đèn xanh, đèn đỏ"),
           tm("DEN XANH (TB)", "pdf", "2021-03-01 00:00:00", r"K\Đèn xanh, đèn đỏ\DEN XANH (TB).pdf", "Đèn xanh, đèn đỏ"),
           tm("Bia gia den xanh", "pdf", "2021-05-01 00:00:00", r"K\Đèn xanh, đèn đỏ\Bia gia den xanh.pdf", "Đèn xanh, đèn đỏ"),
           tm("CHAY (Tr2, 4)", "pdf", "2021-05-01 00:00:00", r"K\Đèn xanh, đèn đỏ\in\CHAY (Tr2, 4).pdf", r"Đèn xanh, đèn đỏ\in"),
           tm("x", "pdf", "2020-01-01 00:00:00", r"K\Đèn xanh, đèn đỏ (1)\x.pdf", "Đèn xanh, đèn đỏ (1)"),
           tm("doi ngoai", "pdf", "2021-01-01 00:00:00", r"K\Đối ngoại quốc phòng\doi ngoai.pdf", "Đối ngoại quốc phòng")]
    import io; buf = io.StringIO(); w = csv.DictWriter(buf, kho[0].keys()); w.writeheader(); w.writerows(kho)
    cp = os.path.join(tempfile.mkdtemp(), "k.csv"); open(cp, "w", encoding="utf-8-sig", newline="").write(buf.getvalue())
    files = doc_kho(cp)
    assert key_thu_muc(os.path.dirname(files[3]["path"])) == ["k", "den xanh den do", "in"]
    assert key_thu_muc(os.path.dirname(files[4]["path"])) == ["k", "den xanh den do"] and "ten_file" not in files[0]
    kq = doi_chieu([dict(nam=2021, dong=2, tt=1, ten="Đèn xanh, đèn đỏ", lan="TB", bbt="QT", qdxb="03/2021"),
                    dict(nam=2021, dong=3, tt=2, ten="Đối ngoại quốc phòng (Sách chuyên khảo)", lan=None, bbt="QT", qdxb=None),
                    dict(nam=2021, dong=4, tt=3, ten="Đèn xanh, đè", lan=None, bbt="QT", qdxb=None)], files)
    assert kq[0]["kieu_khop"] == "chính xác" and kq[0]["so_file"] == 5 and kq[0]["tinh_trang"] == "KHỚP"
    assert kq[0]["file_word"].endswith(".doc") and kq[0]["file_pdf"].endswith("DEN XANH (TB).pdf")
    assert kq[0]["ghi_chu"] == "1 file khác năm bị bỏ qua; bỏ 2 file bìa/tờ che"
    assert kq[1]["kieu_khop"] == "thư mục rút gọn" and kq[1]["file_pdf"].endswith("doi ngoai.pdf")
    kq2 = doi_chieu([dict(nam=2021, dong=2, tt=1, ten="Đối ngoại quốc phòng", lan=None, bbt="QT", qdxb=None),
                     dict(nam=2021, dong=3, tt=2, ten="Đối ngoại quốc phòng (Sách chuyên khảo)", lan=None, bbt="QT", qdxb=None)], files)
    assert kq2[0]["kieu_khop"] == "chính xác" and kq2[1]["so_file"] == 0  # đã có sách trùng tên thư mục -> không khớp rút gọn
    assert kq[2]["tinh_trang"] == "KHÔNG TÌM THẤY FILE"  # "den xanh de" chỉ khớp ký tự, không khớp nguyên từ
    assert nam_trang_cuoi("nghi quyet so 41 2017 qh14 bo luat so 100 2015 qh13 in xong va nop luu chieu thang nam 2021") == [2021]
    assert nam_trang_cuoi("in xong va noap lou chieau thaung 3 naem 2021 chiu") == [2021]
    assert nam_trang_cuoi("ky 2021 2026 dang ky xuat ban 538 2021 qd ngay 19 02 2021 nop luu chieu thang 02 nam 2020") == [2020, 2021]
    assert nam_trang_cuoi("asean 2025 so dang ky xuat ban 551 2020 cxbiph ngay 09 9 2020 ma so isbn") == [2020]
    assert nam_trang_cuoi("phu luc 1 my 1900 2020 483 phu luc 2 eu 1900 2020 518") == []  # mục lục, không có mốc XB
    assert nam_trang_dau("dang cam tu h chinh tri quoc gia 2021 280tr 21cm nam 2015 luat so 100 2015") == [2021]
    assert nam_trang_dau("nha xuat ban chinh tri quoc gia su that ha noi 2021 loi nha xuat ban nam 2020 la nam") == [2021]
    assert nam_trang_dau("gioi thieu cuon sach cung ban doc thang 4 nam 2021 nhj xuat ban nam 2004 2012") == [2021]
    assert nam_trang_dau("so dang ky ke hoach xuat ban 4139 2020 cxbiph nop luu chieu thang 10 nam 2020" + " x" * 200 + " 1900 2019") == [2020]
    assert nam_trang_dau("nghi quyet so 133 2020 qh14 ngay 17 11 2020") == []
    assert nam_trang_dau("van kien nxb chinh tri quoc gia ha noi 2011 t 7 tr 15") == []  # chú thích tài liệu tham khảo
    assert nam_trang_dau("nop luu chieu thang 4 nam 2021 ma isbn 978 bien muc h chinh tri quoc gia 2020 72tr") == [2021]
    assert nam_trang_dau("h chinh tri quoc gia 2021 244tr gioi thieu cung ban doc thang 11 nam 2019 nha xuat ban") == [2021]
    assert la_bia("Bia 1 chay.pdf") and la_bia("To che x.pdf") and la_bia("Tr 2,4 (x).pdf") and not la_bia("Bia da.pdf", "bia da")
    assert not la_bia("Kinh te bien.pdf")

    # --- tab 3: trích text pdf/docx, cache ---
    tmp = tempfile.mkdtemp()
    pdf, dx, doc = (os.path.join(tmp, n) for n in ("a.pdf", "b.docx", "c.doc"))
    d = fitz.open()
    d.new_page().insert_text((50, 72), "Trang bia")
    d.new_page().insert_text((50, 72), "LOI NOI DAU\nCuon sach Luat dau tu nay gioi thieu...")
    d.new_page().insert_text((50, 72), "NHA XUAT BAN CHINH TRI QUOC GIA\nIn xong va nop luu chieu thang 5 nam 2023")
    d.save(pdf); d.close()
    w = docx.Document()
    for s in ("Bia", "Lời Nhà xuất bản", "Cuốn sách Luật đầu tư được tái bản.", "Nội dung", "Nộp lưu chiểu tháng 6 năm 2019"):
        w.add_paragraph(s)
    w.save(dx)
    open(doc, "wb").write(b"\xd0\xcf\x11\xe0 junk")
    t = trich(pdf)
    assert t["nam"] == [2023] and t["dau"] == [2023] and t["lnd"].startswith("cuon sach luat dau tu")
    assert trich(dx)["nam"] == [2019] and "luat dau tu" in trich(dx)["lnd"] and trich(doc) is None
    assert _docx_text(dx) == "Bia\nLời Nhà xuất bản\nCuốn sách Luật đầu tư được tái bản.\nNội dung\nNộp lưu chiểu tháng 6 năm 2019"
    try:  # .doc thật: tạo bằng Word từ b.docx rồi đọc lại (bỏ qua nếu máy không có Word)
        import win32com.client
        w = _word_app()
        if w:
            doc2 = os.path.join(tmp, "e.doc"); dd = w.Documents.Open(dx); dd.SaveAs2(doc2, FileFormat=0); dd.Close(False)
            t2 = trich(doc2)
            assert t2["nam"] == [2019] and t2["dau"] == [2019] and "luat dau tu" in t2["lnd"], t2
            dai = os.path.join(tmp, ("thu muc ten sach rat dai " * 6).strip(), ("ban thao " * 10).strip() + ".doc")  # > 255 ký tự
            os.makedirs(duong_dan_dai(os.path.dirname(dai))); shutil.copy2(doc2, duong_dan_dai(dai))
            assert len(dai) > 255 and trich(dai) == t2 and not glob.glob(os.path.join(tempfile.gettempdir(), "sxb_*.doc"))
            import time
            mod, pid, goi = sys.modules[__name__], _WORD[1], []
            assert pid in _pid_word()
            # hẹn giờ: hết giờ thì gọi hàm tắt Word và lần sau mở Word mới (thay hàm tắt bằng hàm ghi nhận để kiểm tra)
            cu = mod.WORD_HET_GIO, mod._giet_word
            mod.WORD_HET_GIO, mod._giet_word = 0.001, lambda: goi.append(1)
            try:
                trich(doc2)
                assert goi and _WORD[0] is None
            finally:
                mod.WORD_HET_GIO, mod._giet_word = cu
                _WORD[1] = pid; _giet_word(); _WORD[1] = None
            # Word bị tắt thật giữa chừng -> file đó bỏ qua, lần sau tự mở Word mới
            assert trich(doc2)["nam"] == [2019]
            pid = _WORD[1]; _giet_word()
            for _ in range(50):
                if pid not in _pid_word(): break
                time.sleep(0.1)
            assert trich(doc2) is None and _WORD[0] is None
            assert trich(doc2)["nam"] == [2019] and _WORD[1] not in (None, pid)
    finally:
        dong_word()
    assert nam_ten_file("x/CHU NGHIA TU BAN 2021.pdf", "chu nghia tu ban 1900 2020") == {2021}
    assert nam_ten_file("x/BAO CAO NAM 2020 (tb2021).pdf", "bao cao nam 2020") == {2021} and nam_ten_file("x/a_090321R08.pdf") == set()
    cache = {}
    assert trich_cache(cache, pdf) == t and len(cache) == 1 and trich_cache(cache, "khong_co.pdf") is None
    cj = os.path.join(tmp, "c.sqlite"); ghi_cache(cache, cj); c = doc_cache(cj)
    assert len(c) == 1 and c[next(iter(cache))] == t and "x" not in c and c.get("x", 0) == 0
    c["moi"] = None; ghi_cache(c, cj); c.dong(); c = doc_cache(cj); assert len(c) == 2 and c["moi"] is None; c.dong()
    assert loi_noi_dau("Mở đầu. LỜI NHÀ XUẤT BẢN Năm 2020 ... xuất bản cuốn sách Đèn xanh, đèn đỏ. Trân trọng") == \
        "cuon sach den xanh den do tran trong" and loi_noi_dau("Không có lời nói đầu") == ""

    bt = os.path.join(tmp, "ban thao 2021.docx"); w = docx.Document(); w.add_paragraph("Ban thao"); w.save(bt)  # không có năm trong nội dung
    r = dict(nam=2023, ten="Luật đầu tư", tinh_trang="LỆCH NĂM - cần mở file kiểm tra", file_word="", file_pdf="", ghi_chu="",
             cac_file=f"2021-01-01 00:00:00  {pdf}\n2021-01-02 00:00:00  {doc}\n2021-01-03 00:00:00  {dx}\n2021-01-04 00:00:00  {os.path.join(tmp, 'bia.jpg')}\n2021-01-05 00:00:00  {bt}")
    xu_ly_lech_nam(r, cache)
    assert r["tinh_trang"] == "KHỚP" and r["file_pdf"] == pdf and r["ghi_chu"] == "khớp theo nội dung (trang cuối)"
    assert r["file_word"] == doc  # .doc cùng thư mục được nhận; .docx ghi năm 2019 hay tên file ghi 2021 thì không
    # trang cuối ghi lưu chiểu 2020, CIP ở đầu sách ghi 2021 -> dòng 2020 khớp trang cuối, dòng 2021 khớp 8 trang đầu
    hai = os.path.join(tmp, "hai", "h.pdf"); os.makedirs(os.path.dirname(hai)); d = fitz.open()
    d.new_page().insert_text((50, 72), "H. : Chinh tri quoc gia, 2021. - 280tr.")
    for _ in range(9): d.new_page().insert_text((50, 72), "Noi dung nam 2015")
    d.new_page().insert_text((50, 72), "In xong va nop luu chieu thang 9 nam 2020"); d.save(hai); d.close()
    assert trich(hai)["nam"] == [2020] and trich(hai)["dau"] == [2021]
    for nam, nguon in ((2020, "trang cuối"), (2021, "8 trang đầu")):
        rh = dict(r, nam=nam, tinh_trang="LỆCH NĂM", cac_file="2021-01-01 00:00:00  " + hai); xu_ly_lech_nam(rh, cache)
        assert rh["tinh_trang"] == "KHỚP" and rh["ghi_chu"] == f"khớp theo nội dung ({nguon})"
    rh = dict(r, nam=2015, tinh_trang="LỆCH NĂM", cac_file="2021-01-01 00:00:00  " + hai); xu_ly_lech_nam(rh, cache)
    assert rh["tinh_trang"] == "LỆCH NĂM" and "thấy: 2020, 2021" in rh["ghi_chu"]  # năm trong nội dung không tính
    r2 = dict(r, nam=2020, tinh_trang="LỆCH NĂM - cần mở file kiểm tra"); xu_ly_lech_nam(r2, cache)
    assert r2["tinh_trang"].startswith("LỆCH NĂM") and "thấy: 2019, 2023" in r2["ghi_chu"] and "bia.jpg" not in r2["ghi_chu"]

    kho = [dict(ten_file="a", dinh_dang="pdf", date_modified="2021-01-01 00:00:00", path=pdf, nam=2021),
           dict(ten_file="b", dinh_dang="docx", date_modified="2021-06-01 00:00:00", path=dx, nam=2021),
           dict(ten_file="c", dinh_dang="doc", date_modified="2021-06-01 00:00:00", path=doc, nam=2021)]
    assert len(loc_files(kho)) == 3 and len(loc_files(kho, "", "2020")) == 0 and len(loc_files(kho, tmp, "2020-2021")) == 3
    assert loc_files(kho, os.path.join(tmp, "khac")) == []
    rows = [dict(nam=2023, ten="Luật đầu tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2023, ten="Không có", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2019, ten="Luật đầu tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2018, ten="Luật tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file=""),
            dict(nam=2023, ten="Đầu tư", tinh_trang="KHÔNG TÌM THẤY FILE", file_word="", file_pdf="", ghi_chu="", cac_file="")]
    DUNG.set()
    try:
        xu_ly_khong_tim_thay([dict(r) for r in rows], loc_files(kho), cache); assert False
    except DaDung:
        pass
    finally:
        DUNG.clear()
    xu_ly_khong_tim_thay(rows, loc_files(kho), cache)
    assert rows[0]["tinh_trang"] == "KHỚP" and rows[0]["file_pdf"] == pdf and rows[0]["so_file"] == 2 and rows[0]["file_word"] == doc
    assert rows[1]["tinh_trang"] == "KHÔNG TÌM ĐƯỢC"
    assert rows[2]["tinh_trang"] == "KHỚP" and rows[2]["file_word"] == dx and rows[2]["file_pdf"] == ""
    assert rows[3]["tinh_trang"] == "KHÔNG TÌM ĐƯỢC"  # "luat tu" không là cụm liền trong lời nói đầu
    assert rows[4]["tinh_trang"] == "KHÔNG TÌM ĐƯỢC"  # "dau tu" có trong lời nói đầu nhưng không đứng sau "cuốn sách"
    assert nhac_ten_sach("nxb xuat ban cuon sach chuyen khao luat dau tu", "luat dau tu")
    assert not nhac_ten_sach("can cu hien phap nuoc cong hoa nam 2013", "hien phap nuoc cong hoa")

    # --- ghi / đọc lại doi_chieu, tab 4 ---
    xl = os.path.join(tmp, "th.xlsx"); wb = openpyxl.Workbook(); ws = wb.active; ws.title = "2023"
    ws.append(["TT", "Tên sách"]); ws.append([1, "Luật đầu tư"]); ws.append([2, "Không có"]); wb.save(xl)
    kq = [dict(nam=2023, dong=2, tt=1, ten="Luật đầu tư", lan="TB", bbt="PL", qdxb="05/2023", tinh_trang="KHỚP",
               kieu_khop="chính xác", so_file=2, file_word=dx, file_pdf=pdf, ten_moi="luat_dau_tu_0523_R_PL_T01",
               ghi_chu="", cac_file=""),
          dict(nam=2023, dong=3, tt=2, ten="Không có", lan=None, bbt="PL", qdxb=None, tinh_trang="KHÔNG TÌM THẤY FILE",
               kieu_khop="", so_file=0, file_word="", file_pdf="", ten_moi="khong_co_R_PL_M00", ghi_chu="", cac_file="")]
    kho_tm = [dict(dinh_dang="docx", path=dx), dict(dinh_dang="pdf", path=pdf),
              dict(dinh_dang="jpg", path=os.path.join(tmp, "x", "bia.jpg")), dict(dinh_dang="pdf", path=os.path.join(tmp, "y", "z.pdf"))]
    kq[1]["cac_file"] = "2023-01-01 00:00:00  " + os.path.join(tmp, "y", "z.pdf")
    trung = os.path.join(tmp, "z_trung", "a.pdf"); os.makedirs(os.path.dirname(trung)); shutil.copy2(pdf, trung)
    kho_tm.append(dict(dinh_dang="pdf", path=trung)); kq[0]["cac_file"] = "2023-01-01 00:00:00  " + trung
    assert doi_chieu_thu_muc(kq, kho_tm, cache)[-1]["tinh_trang"] == TM_CHUA  # bản trùng chưa đọc -> chưa biết năm
    trich_cache(cache, trung)
    tm = doi_chieu_thu_muc(kq, kho_tm, cache)
    assert [(x["ten_tm"], x["tinh_trang"], x["so_file"]) for x in tm] == [
        (os.path.basename(tmp), TM_KHOP, 2), ("x", TM_KHONG, 1), ("y", TM_CHUA, 1), ("z_trung", TM_TRUNG, 1)]
    assert tm[3]["nam_cuoi"] == "trang cuối: 2023; 8 trang đầu: 2023"
    assert tm[0]["nam_cuoi"] == "trang cuối: 2019, 2023; 8 trang đầu: 2019, 2023" and tm[0]["ghi_chu"] == "đọc được 2/2 file Word/PDF"
    assert tm[1]["ghi_chu"].startswith("không có file Word/PDF") and tm[2]["ghi_chu"].startswith("đọc được 0/1")
    assert tm[0]["loai"] == "DOCX 1, PDF 1" and tm[0]["sach_khop"] == "2023 - dòng 2 - Luật đầu tư" and tm[2]["sach_khac"].endswith("Không có")
    ra1, ra2 = ghi_ket_qua(kq, xl, os.path.join(tmp, "ra"), tm)
    assert doc_thu_muc(ra1) == tm
    kq2 = doc_doi_chieu(ra1)
    assert [r["ten"] for r in kq2] == ["Luật đầu tư", "Không có"] and kq2[0]["file_pdf"] == pdf and kq2[0]["dong"] == 2
    kq2.append(dict(kq2[0], dong=4, qdxb=None))  # trùng tên file mới -> _2
    dich = os.path.join(tmp, "dich"); kh = ke_hoach_luu(kq2, dich)
    t = "luat_dau_tu_0523_R_PL_T01"
    assert [os.path.relpath(p, dich) for _, _, p in kh] == [
        os.path.join("2023", "thang 05", t, t + ".docx"), os.path.join("2023", "thang 05", t, t + ".pdf"),
        os.path.join("2023", "thang 00", t, t + ".docx"), os.path.join("2023", "thang 00", t, t + ".pdf")]
    kq2[2]["qdxb"] = "05/2023"; kh = ke_hoach_luu(kq2, dich)  # trùng thư mục -> thư mục _2
    assert kh[2][2] == os.path.join(dich, "2023", "thang 05", t + "_2", t + "_2.docx") and kh[3][2].endswith(t + "_2.pdf")
    assert luu_file(kh) == (4, 0, []) and all(os.path.isfile(p) for _, _, p in kh) and os.path.isfile(pdf)
    assert kq2[0]["dich"] == kh[0][2] + "\n" + kh[1][2] and not glob.glob(os.path.join(dich, "**", "*.dang_chep"), recursive=True)
    # chạy lại: dùng lại đúng thư mục cũ, không chép lại
    kh2 = ke_hoach_luu([dict(r, dich="") for r in kq2], dich)
    assert [p for _, _, p in kh2] == [p for _, _, p in kh] and luu_file(kh2) == (0, 4, [])
    # thư mục trùng tên nhưng là sách khác (file khác dung lượng) -> sang _3; copy dở (thiếu PDF) -> dùng lại
    kh3 = ke_hoach_luu([dict(kq2[0], file_word="", file_pdf=hai)], dich)
    assert kh3[0][2] == os.path.join(dich, "2023", "thang 05", t + "_3", t + "_3.pdf")
    os.remove(kh[1][2]); kh4 = ke_hoach_luu([dict(kq2[0], dich="")], dich)
    assert [p for _, _, p in kh4] == [p for _, _, p in kh[:2]] and luu_file(kh4) == (1, 1, [])
    # bấm Dừng: không chép thêm file nào
    DUNG.set()
    try:
        assert luu_file(ke_hoach_luu([dict(kq2[0], ten_moi="sach_dung", dich="")], dich)) == (0, 0, [])
        assert not os.path.exists(os.path.join(dich, "2023", "thang 05", "sach_dung"))
    finally:
        DUNG.clear()
    # lỗi một file không làm dừng cả lượt
    hong = dict(kq2[0], file_word=os.path.join(tmp, "khong_co.docx"), ten_moi="sach_loi", dich="", ghi_chu="")
    chep, co_san, loi = luu_file(ke_hoach_luu([hong, dict(kq2[0], ten_moi="sach_tot", dich="")], dich))
    assert chep == 3 and len(loi) == 1 and "LỖI lưu khong_co.docx" in hong["ghi_chu"]
    ra1, ra2 = ghi_ket_qua(kq2, xl, os.path.join(tmp, "ra"))
    ws = openpyxl.load_workbook(ra2)["2023"]
    assert ws.cell(1, 6).value == "Đường dẫn đích" and ws.cell(2, 6).value == kq2[0]["dich"] and ws.cell(2, 3).value == "KHỚP"
    shutil.rmtree(tmp, ignore_errors=True)
    print("OK")


if __name__ == "__main__":
    _test() if "--test" in sys.argv else gui()
