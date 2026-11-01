# -*- coding: utf-8 -*-
"""
DOC DU LIEU TU MOT THU MUC KET QUA
===================================
Nhan dien va doc ba kieu thu muc ma phan mem sinh ra, tra ve MOT cau truc
thong nhat de tab bao cao SPC khong phai biet du lieu den tu dau.

    Charpy_<...>/     TongHop_VaDap.xlsx        bang tong hop BM31
    Keo_Sinh_<...>/   SoLieu_<N>.xlsx           sheet Co_ly_tinh
    Keo_<...>/        BaoCao/Tensile_*.docx     chi co bao cao Word

Kieu thu ba la ca kho: thu muc cua Tab 2 khong luu file so lieu nao ca, nen
phai boc nguoc bang thong so trong tung file .docx. Do thuc te 14.8 ms moi
file, 1000 file het chung 15 giay - cham nhung van dung duoc.

Cung chap nhan tro thang vao mot file .xlsx thay vi ca thu muc.
"""

import os
import re
import glob
import zipfile
import datetime
import math
import spc_luat

import openpyxl

# ------------------------------------------------------------ bang thong so
# Ten cot -> (don vi, nhan in tren bieu mau SPC, nguong mac dinh)
COT_KEO = [
    ("ReL", "N/mm2", "Gioi han chay duoi ReL", 275.0),
    ("ReH", "N/mm2", "Gioi han chay tren ReH", None),
    ("Rm", "N/mm2", "Do ben keo Rm", 480.0),
    ("A", "%", "Do gian dai tuong doi A", 20.0),
    ("Z", "%", "Do that tuong doi Z", None),
]
# Nguong duoi (LSL) cua thu keo theo MAC VAT LIEU - lay dung tu hai bieu mau
# "JIS.QT.07.BM.01 ... - Keo - SCW480.xlsx" va "... - Keo - SCW410.xlsx"
# (sheet 'Du lieu goc', o C16 / C18 / C20). Hai bieu mau chi khac nhau o mac
# vat lieu va ba nguong nay. Gioi han chay ap cho ca ReL lan ReH.
MAC_KEO = {
    "SCW480": {"ReL": 275.0, "ReH": 275.0, "Rm": 480.0, "A": 20.0},
    "SCW410": {"ReL": 235.0, "ReH": 235.0, "Rm": 410.0, "A": 21.0},
}
MAC_KEO_MAC_DINH = "SCW480"

COT_CHARPY = [
    ("Nang luong TB", "J", "Nang luong hap thu", 27.0),
    ("% mat gay TB", "%", "Phan tram mat gay", None),
]

# Vi tri cot trong sheet Co_ly_tinh cua SoLieu_<N>.xlsx (1-based)
_CLT = {"ReL": 8, "ReH": 10, "Rm": 12, "A": 13, "Z": 14}
# Vi tri cot trong bang thong so cua bao cao Word thu keo (0-based, sau o dau)
_WORD = {"ReL": 6, "ReH": 8, "Rm": 10, "A": 11, "Z": 12}

_RE_TBL = re.compile(r"<w:tbl>.*?</w:tbl>", re.S)
_RE_TR = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
_RE_TC = re.compile(r"<w:tc>.*?</w:tc>", re.S)
_RE_WT = re.compile(r"(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)")


class LoiDoc(Exception):
    """Loi giai thich duoc cho nguoi dung."""
    pass


_MOC_EXCEL = datetime.datetime(1899, 12, 30)


def _so(v):
    """Doc mot o ve so. Chap nhan ca dang '≥27' lan o bi Excel hieu nham
    thanh ngay thang.

    Vi sao phai xu ly ngay thang: mot so ban tong hop cu duoc sinh ra truoc
    khi charpy_core.sinh_tong_hop ep lai number_format tung cot, nen o
    'Nang luong trung binh' thua dinh dang NGAY THANG cua file mau. Excel
    luu 72.03 nhung tra ve datetime(1900, 3, 12, 0, 43, 12). Doi nguoc lai
    thanh so seri thi lay lai duoc dung gia tri goc, thay vi bo ca cot.
    """
    if v is None or str(v).strip() in ("", "-"):
        return None
    if isinstance(v, datetime.datetime):
        return (v - _MOC_EXCEL).total_seconds() / 86400.0
    if isinstance(v, datetime.date):
        return float((v - _MOC_EXCEL.date()).days)
    if isinstance(v, datetime.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400.0
    try:
        x = float(str(v).replace(",", ".").lstrip("≥≤><= "))
        return x if math.isfinite(x) else None
    except ValueError:
        return None


def _chu(v):
    return str(v).strip() if v not in (None, "") else None


# ---------------------------------------------------------------- nhan dien
def nhan_dien(duong_dan):
    """Tra ve ('charpy_xlsx'|'keo_xlsx'|'keo_docx'|'charpy_docx', file/thu muc)."""
    if os.path.isfile(duong_dan):
        if duong_dan.lower().endswith((".xlsx", ".xlsm")):
            wb = openpyxl.load_workbook(duong_dan, read_only=True)
            ten = wb.sheetnames
            wb.close()
            if "Co_ly_tinh" in ten:
                return "keo_xlsx", duong_dan
            return "charpy_xlsx", duong_dan
        raise LoiDoc("Chi doc duoc file .xlsx hoac ca mot thu muc ket qua.")

    if not os.path.isdir(duong_dan):
        raise LoiDoc("Khong tim thay:\n%s" % duong_dan)

    f = os.path.join(duong_dan, "TongHop_VaDap.xlsx")
    if os.path.exists(f):
        return "charpy_xlsx", f
    g = sorted(glob.glob(os.path.join(duong_dan, "SoLieu_*.xlsx")))
    if g:
        return "keo_xlsx", g[0]
    g = sorted(glob.glob(os.path.join(duong_dan, "BaoCao", "Tensile_*.docx")))
    if g:
        return "keo_docx", duong_dan
    g = sorted(glob.glob(os.path.join(duong_dan, "BaoCao", "KetQua_*.docx")))
    if g:
        raise LoiDoc(
            "Thu muc nay chi co bao cao Word thu va dap, chua co bang tong "
            "hop.\n\nHay chon thu muc co file TongHop_VaDap.xlsx, hoac chon "
            "thang file do.")
    raise LoiDoc(
        "Khong nhan ra thu muc ket qua:\n%s\n\n"
        "Thu muc phai co mot trong cac file sau:\n"
        "    TongHop_VaDap.xlsx        (thu va dap Charpy)\n"
        "    SoLieu_<so>.xlsx          (sinh bieu do keo)\n"
        "    BaoCao/Tensile_*.docx     (gop bao cao thu keo)" % duong_dan)


# ------------------------------------------------------------------ doc
def _doc_charpy_xlsx(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    ma, e_tb, p_tb = [], [], []
    r = 12                                   # dong du lieu dau tien cua BM31
    while r <= ws.max_row:
        stt = ws.cell(r, 1).value
        if stt is None or _so(stt) is None:
            break                            # het bang, gap dong ket luan
        ma.append(_chu(ws.cell(r, 2).value) or str(int(_so(stt))))
        # Gia tri dua vao SPC la Energy trung binh lam tron 1 chu so - dung
        # dang so lieu ma Tab 1 dang ghi vao file danh gia do on dinh.
        v = _so(ws.cell(r, 7).value)
        e_tb.append(round(v, 1) if v is not None else None)
        v = _so(ws.cell(r, 11).value)
        p_tb.append(round(v, 1) if v is not None else None)
        r += 1

    nguong = {"Nang luong TB": _so(ws["G11"].value),
              "% mat gay TB": _so(ws["K11"].value)}
    tt = {"vat_lieu_spc": _chu(ws["I4"].value),
          "tieu_chuan_ap_dung": _chu(ws["I6"].value),
          "ten_san_pham": _chu(ws["C6"].value)}
    return {"loai": "charpy", "nguon": path, "ma": ma,
            "cot": {"Nang luong TB": e_tb, "% mat gay TB": p_tb},
            "nguong": nguong,
            "thong_tin": {k: v for k, v in tt.items() if v}}


def _doc_keo_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Co_ly_tinh" not in wb.sheetnames:
        raise LoiDoc("File Excel khong co sheet 'Co_ly_tinh':\n%s" % path)
    ws = wb["Co_ly_tinh"]
    ma = []
    cot = dict((k, []) for k in _CLT)
    r = 3                                    # dong 1 tieu de, dong 2 don vi
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v in (None, ""):
            break
        ma.append(_chu(v))
        for k, c in _CLT.items():
            cot[k].append(_so(ws.cell(r, c).value))
        r += 1
    return {"loai": "keo", "nguon": path, "ma": ma, "cot": cot,
            "nguong": dict((k, n) for k, _u, _n, n in COT_KEO),
            "thong_tin": {}}


def _doc_bang_word(path):
    """Boc ma mau va bang thong so tu mot bao cao Word thu keo."""
    with zipfile.ZipFile(path) as z:
        x = z.read("word/document.xml").decode("utf-8")
    b = _RE_TBL.findall(x)
    if len(b) < 2:
        raise LoiDoc("Bao cao Word thieu bang thong so:\n%s" % path)

    def chu_o(tc):
        return "".join(m[1] for m in _RE_WT.findall(tc)).strip()

    hang = _RE_TR.findall(b[1])
    if len(hang) < 3:
        raise LoiDoc("Bang thong so trong bao cao Word khong du dong:\n%s" % path)
    o = [chu_o(t) for t in _RE_TC.findall(hang[2])][1:]
    h0 = _RE_TR.findall(b[0])
    ma = chu_o(_RE_TC.findall(h0[1])[1]) if len(h0) > 1 else ""
    return ma, o


def _doc_keo_docx(thu_muc, tien_do=None):
    files = sorted(glob.glob(os.path.join(thu_muc, "BaoCao", "Tensile_*.docx")))
    if not files:
        raise LoiDoc("Khong co bao cao Word nao trong:\n%s" % thu_muc)
    ma = []
    cot = dict((k, []) for k in _WORD)
    for i, f in enumerate(files):
        m, o = _doc_bang_word(f)
        ma.append(m or os.path.splitext(os.path.basename(f))[0])
        for k, j in _WORD.items():
            cot[k].append(_so(o[j]) if j < len(o) else None)
        if tien_do and (i % 25 == 0 or i == len(files) - 1):
            tien_do(i + 1, len(files))
    return {"loai": "keo", "nguon": os.path.join(thu_muc, "BaoCao"),
            "ma": ma, "cot": cot,
            "nguong": dict((k, n) for k, _u, _n, n in COT_KEO),
            "thong_tin": {}}


def doc(duong_dan, tien_do=None):
    """Doc mot thu muc ket qua (hoac mot file .xlsx). Tra ve bang thong nhat."""
    kieu, f = nhan_dien(duong_dan)
    if kieu == "charpy_xlsx":
        d = _doc_charpy_xlsx(f)
    elif kieu == "keo_xlsx":
        d = _doc_keo_xlsx(f)
    else:
        d = _doc_keo_docx(f, tien_do)

    n = len(d["ma"])
    if n == 0:
        raise LoiDoc("Khong doc duoc dong du lieu nao trong:\n%s" % d["nguon"])
    # Bo cac cot khong dung duoc, nhung GHI LAI de con bao cho nguoi dung -
    # bo cot ma im lang thi ho khong hieu vi sao thieu.
    bo = {}
    giu = {}
    for k, v in d["cot"].items():
        thieu = sum(1 for x in v if x is None)
        if thieu == n:
            bo[k] = "khong doc duoc gia tri nao"
        elif thieu:
            bo[k] = "thieu gia tri o %d/%d mau" % (thieu, n)
            giu[k] = v
        else:
            giu[k] = v
    d["cot"] = giu
    d["cot_bo_qua"] = bo
    if not d["cot"]:
        raise LoiDoc(
            "Khong co cot so lieu nao dung duoc trong:\n%s\n\n%s"
            % (d["nguon"],
               "\n".join("    %s: %s" % (k, m) for k, m in bo.items())))
    d["so_mau"] = n
    return d


def danh_sach_cot(d):
    """Cac cot cham duoc, theo dung thu tu in tren bieu mau."""
    goc = COT_CHARPY if d["loai"] == "charpy" else COT_KEO
    return [(k, u, nh) for k, u, nh, _ng in goc if k in d["cot"]]


def lay_day(d, ten_cot, bat_dau, buoc, so_diem=25):
    """Lay day con theo (mau bat dau, buoc nhay). bat_dau dem tu 1.

    Tra ve (danh_sach_chi_so_1_based, danh_sach_gia_tri).
    """
    try:
        bat_dau = spc_luat.so_nguyen(bat_dau, "Mau bat dau")
        buoc = spc_luat.so_nguyen(buoc, "Buoc nhay")
        so_diem = spc_luat.so_nguyen(so_diem, "So diem", 2)
    except ValueError as e:
        raise LoiDoc(str(e))
    v = d["cot"].get(ten_cot)
    if v is None:
        raise LoiDoc("Khong co cot '%s' trong du lieu da doc." % ten_cot)
    if bat_dau < 1:
        raise LoiDoc("Mau bat dau phai tu 1 tro len.")
    if buoc < 1:
        raise LoiDoc("Buoc nhay phai tu 1 tro len.")
    idx = [bat_dau + i * buoc for i in range(so_diem)]
    if idx[-1] > d["so_mau"]:
        raise LoiDoc(
            "Lay %d mau bat dau tu %d voi buoc nhay %d thi can toi mau thu "
            "%d, nhung du lieu chi co %d mau.\n\n"
            "Hay giam buoc nhay, giam so diem, hoac chon mau bat dau nho hon."
            % (so_diem, bat_dau, buoc, idx[-1], d["so_mau"]))
    gt = [v[i - 1] for i in idx]
    thieu = [i for i, x in zip(idx, gt) if _so(x) is None]
    if thieu:
        raise LoiDoc("Cot '%s' khong co gia tri o cac mau: %s"
                     % (ten_cot, ", ".join(str(x) for x in thieu[:8])))
    return idx, gt
