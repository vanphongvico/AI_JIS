# -*- coding: utf-8 -*-
"""
BO CAU HINH DA LUU - DUNG CHUNG CHO MOI TAB
============================================
Moi tab co DANH SACH BO CAU HINH RIENG (vd tab keo: "SCW480 - khach A",
"SCW410 - khach B"; tab va dap co danh sach khac han). Chon mot bo thi
toan bo o trong tab tu dien; sua tay sau do khong lam doi bo da luu tru
khi bam Luu.

Thong tin KHONG tu dong bo giua cac tab. Muon dung chung thi bam
"Lay thong tin tu..." - chi chep cac o cung vai tro (nguoi phe duyet,
vat lieu...), va chi khi nguoi dung chu dong bam.

Luu o:  <thu muc phan mem>/cau_hinh/<khoa tab>.json
    {"bo": {"ten bo": {...}}, "lan_truoc": {...}}

Duong dan file nam trong thu muc phan mem luu dang "%PHAN_MEM%\\file_mau\\..."
nen chuyen / chep ca thu muc sang cho khac (may khac) van tro dung file.
"""

import os
import re
import sys
import json

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import ui_theme as ui


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def thu_muc_cau_hinh():
    return os.path.join(base_dir(), "cau_hinh")


def thu_muc_file_mau():
    """Thu muc chua san cac file mau giao kem phan mem."""
    return os.path.join(base_dir(), "file_mau")


# Windows (chua bat Long Paths) khong mo duoc duong dan dai hon 259 ky tu: file
# van con ma phan mem bao "khong tim thay". Ten file mau SPC thanh phan da dai
# 85 ky tu, nen giai nen phan mem vao cho qua sau la vuot.
GIOI_HAN_DUONG_DAN = 259
CHO_DU_KET_QUA = 110        # \output\<thu muc lan chay>\<file ket qua> dai nhat


def do_dai_duong_dan_toi_da(goc=None):
    """Duong dan dai nhat: file trong file_mau/, templates/ va file ket qua du kien."""
    goc = goc or base_dir()
    dai = len(goc) + CHO_DU_KET_QUA
    for thu in ("file_mau", "templates"):
        for g, _c, fs in os.walk(os.path.join(goc, thu)):
            for f in fs:
                dai = max(dai, len(os.path.join(g, f)))
    return dai


def canh_bao_duong_dan_dai(goc=None):
    """Loi canh bao (tieng Viet) neu thu muc phan mem nam qua sau, khong thi None."""
    goc = goc or base_dir()
    n = do_dai_duong_dan_toi_da(goc)
    if n <= GIOI_HAN_DUONG_DAN:
        return None
    return ("Phần mềm đang nằm ở:\n%s\n\nĐường dẫn tới file mẫu / file kết quả dài tới "
            "%d ký tự, vượt giới hạn 260 ký tự của Windows — phần mềm có thể báo "
            "“không tìm thấy file” dù file vẫn còn.\n\nHãy chép cả thư mục phần mềm ra "
            "chỗ ngắn hơn, ví dụ:  C:\\BaoCaoCoTinh_BanGiao" % (goc, n))


def tim_file_mau(*tien_to):
    """Tim file dau tien trong file_mau/ co ten bat dau bang mot trong cac
    tien to (khong phan biet hoa thuong). Khong thay thi tra ve ''."""
    tm = thu_muc_file_mau()
    try:
        ds = sorted(os.listdir(tm))
    except OSError:
        return ""
    for tt in tien_to:
        for f in ds:
            if f.lower().startswith(tt.lower()) and not f.startswith("~$"):
                return os.path.join(tm, f)
    return ""


# ------------------------------------------- duong dan di theo phan mem
GOC_PM = "%PHAN_MEM%"
_RE_TUYET_DOI = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\)")    # C:\... hoac \\may\...


def _sang_tuong_doi(v):
    """Duong dan nam trong thu muc phan mem -> '%PHAN_MEM%\\...'."""
    if isinstance(v, dict):
        return {k: _sang_tuong_doi(x) for k, x in v.items()}
    if isinstance(v, list):
        return [_sang_tuong_doi(x) for x in v]
    if isinstance(v, str) and _RE_TUYET_DOI.match(v):
        goc = os.path.normcase(os.path.normpath(base_dir()))
        p = os.path.normpath(v)
        if os.path.normcase(p).startswith(goc + os.sep):
            return GOC_PM + p[len(goc):]
    return v


def _tim_lai(v):
    """Duong dan tuyet doi (cau hinh kieu cu) khong con: tim phan duoi cua no
    (vd file_mau\\ThanhPhan\\X.xlsx) trong thu muc phan mem hien tai."""
    goc = base_dir()
    try:
        co = {t.lower() for t in os.listdir(goc) if os.path.isdir(os.path.join(goc, t))}
    except OSError:
        return v
    phan = re.split(r"[\\/]+", v)
    for i in range(1, len(phan)):
        if phan[i].lower() in co:
            p = os.path.join(goc, *phan[i:])
            if os.path.exists(p):
                return p
    return v


def _sang_tuyet_doi(v):
    if isinstance(v, dict):
        return {k: _sang_tuyet_doi(x) for k, x in v.items()}
    if isinstance(v, list):
        return [_sang_tuyet_doi(x) for x in v]
    if isinstance(v, str):
        if v.startswith(GOC_PM):
            return os.path.normpath(base_dir() + v[len(GOC_PM):])
        if _RE_TUYET_DOI.match(v) and not os.path.exists(v):
            return _tim_lai(v)
    return v


# ------------------------------------------------------------- doc / ghi
def _file(khoa):
    return os.path.join(thu_muc_cau_hinh(), "%s.json" % khoa)


def _doc(khoa):
    try:
        with open(_file(khoa), encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict):
            d = _sang_tuyet_doi(d)
            d.setdefault("bo", {})
            return d
    except Exception:
        pass
    return {"bo": {}}


def _ghi(khoa, d):
    os.makedirs(thu_muc_cau_hinh(), exist_ok=True)
    tam = _file(khoa) + ".tmp"
    with open(tam, "w", encoding="utf-8") as f:
        json.dump(_sang_tuong_doi(d), f, ensure_ascii=False, indent=2)
    os.replace(tam, _file(khoa))


def nho_lan_truoc(khoa, du_lieu):
    """Ghi lai tham so lan chay vua roi. Loi ghi thi bo qua im lang."""
    try:
        d = _doc(khoa)
        d["lan_truoc"] = du_lieu
        _ghi(khoa, d)
    except Exception:
        pass


def doc_lan_truoc(khoa, file_cu=None):
    """file_cu: ten file ghi nho kieu cu (lan_dung_truoc_*.json o thu muc
    phan mem) - doc tam neu chua co ban moi, de may dang dung khong mat
    thong so da nho."""
    d = _doc(khoa).get("lan_truoc")
    if d is None and file_cu:
        try:
            with open(os.path.join(base_dir(), file_cu), encoding="utf-8") as f:
                d = _sang_tuyet_doi(json.load(f))
        except Exception:
            d = None
    return d


# ------------------------------------------------- chup / dat bien giao dien
def _chup(v):
    if isinstance(v, tk.Variable):
        return v.get()
    if isinstance(v, (list, tuple)):
        ra = [_chup(x) for x in v]
        return ra if all(x is not None for x in ra) else None
    if isinstance(v, dict):
        ra = {str(k): _chup(x) for k, x in v.items()}
        return ra if all(x is not None for x in ra.values()) else None
    return None


def _dat(v, gia_tri):
    if isinstance(v, tk.Variable):
        try:
            v.set(gia_tri)
        except tk.TclError:
            pass
    elif isinstance(v, (list, tuple)) and isinstance(gia_tri, list):
        for x, g in zip(v, gia_tri):
            _dat(x, g)
    elif isinstance(v, dict) and isinstance(gia_tri, dict):
        for k, x in v.items():
            if str(k) in gia_tri:
                _dat(x, gia_tri[str(k)])


def chup_bien(obj, bo_qua=()):
    """Chup moi thuoc tinh v_* (bien tkinter, hoac list/dict cua chung)."""
    ra = {}
    for ten, v in vars(obj).items():
        if not ten.startswith("v_") or ten in bo_qua:
            continue
        g = _chup(v)
        if g is not None:
            ra[ten] = g
    return ra


def dat_bien(obj, du_lieu, bo_qua=()):
    for ten, g in (du_lieu or {}).items():
        if ten in bo_qua or not hasattr(obj, ten):
            continue
        _dat(getattr(obj, ten), g)


# ------------------------------------------- "Lay thong tin tu tab khac"
# Moi tab dang ky: khoa -> (ten hien thi, ham lay {vai_tro: gia_tri},
# ham dat {vai_tro: gia_tri}). Vai tro chung:
#   nguoi_thu, nguoi_lap, nguoi_kiem_tra, nguoi_phe_duyet,
#   vat_lieu, tieu_chuan_thu, tieu_chuan_vat_lieu, san_pham
_DANG_KY = {}

TEN_VAI_TRO = {
    "nguoi_thu": "Người thử", "nguoi_lap": "Người lập",
    "nguoi_kiem_tra": "Người kiểm tra", "nguoi_phe_duyet": "Người phê duyệt",
    "vat_lieu": "Vật liệu", "tieu_chuan_thu": "Tiêu chuẩn thử",
    "tieu_chuan_vat_lieu": "Tiêu chuẩn vật liệu", "san_pham": "Sản phẩm",
}


def dang_ky_thong_tin(khoa, ten, ban_do):
    """ban_do: {vai_tro: bien tkinter} cua tab."""
    _DANG_KY[khoa] = (ten, ban_do)


# --------------------------------------------------------- thanh giao dien
class ThanhCauHinh(object):
    """Thanh 'Bo cau hinh' dat o dau mot tab.

    lay()      -> dict tham so hien tai cua tab (de luu)
    dat(dict)  -> dien dict vao cac o cua tab
    """

    def __init__(self, me, khoa, ten_tab, lay, dat):
        self.khoa = khoa
        self.ten_tab = ten_tab
        self.lay = lay
        self.dat = dat

        k = tk.Frame(me, bg=ui.COLOR_SURFACE, padx=12, pady=8,
                     highlightbackground=ui.COLOR_BORDER, highlightthickness=1)
        k.pack(fill="x", padx=12, pady=(8, 2))
        self.khung = k
        ngoai = k
        # Hang nut o tren, dong goi y/trang thai o hang rieng ben duoi -
        # de chung mot hang thi cua so hep se cat mat chu.
        k = tk.Frame(ngoai, bg=ui.COLOR_SURFACE)
        k.pack(fill="x")

        tk.Label(k, text="🗂  Bộ cấu hình", font=("Segoe UI", 9, "bold"),
                 fg=ui.COLOR_TEXT_MAIN, bg=ui.COLOR_SURFACE).pack(side="left")
        self.v_chon = tk.StringVar(value="")
        self.cb = ttk.Combobox(k, textvariable=self.v_chon, state="readonly",
                               width=34)
        self.cb.pack(side="left", padx=(8, 6), ipady=1)
        self.cb.bind("<<ComboboxSelected>>", self._ap_dung)
        ttk.Button(k, text="💾 Lưu", style="Outline.TButton",
                   command=self.luu).pack(side="left", padx=2)
        ttk.Button(k, text="🗑 Xoá", style="Outline.TButton",
                   command=self.xoa).pack(side="left", padx=2)

        self.nut_lay = ttk.Button(k, text="📋 Lấy thông tin từ tab khác ▾",
                                  style="Outline.TButton",
                                  command=self._mo_menu_lay)
        self.nut_lay.pack(side="left", padx=(12, 0))
        self.menu = tk.Menu(self.nut_lay, tearoff=0)

        self.lbl = tk.Label(ngoai, text="", font=("Segoe UI", 8),
                            fg=ui.COLOR_TEXT_MUTED, bg=ui.COLOR_SURFACE,
                            anchor="w", justify="left")
        self.lbl.pack(fill="x", pady=(4, 0))
        self._nap_ds()

    # ----------------------------------------------------------- danh sach
    def _nap_ds(self, chon=None):
        ds = sorted(_doc(self.khoa)["bo"].keys(), key=lambda s: s.lower())
        self.cb["values"] = ds
        if chon in ds:
            self.v_chon.set(chon)
        elif self.v_chon.get() not in ds:
            self.v_chon.set("")
        if not ds:
            self.lbl.configure(text="Chưa có bộ nào — điền xong bấm 💾 Lưu "
                                    "để dùng lại lần sau.")

    def _ap_dung(self, _e=None):
        ten = self.v_chon.get()
        bo = _doc(self.khoa)["bo"].get(ten)
        if bo is None:
            return
        try:
            self.dat(bo)
            self.lbl.configure(text="Đã điền theo bộ “%s”." % ten)
        except Exception as e:
            messagebox.showerror("Không áp dụng được",
                                 "Bộ cấu hình “%s” bị lỗi:\n%s" % (ten, e))

    def luu(self):
        ten = simpledialog.askstring(
            "Lưu bộ cấu hình",
            "Đặt tên cho bộ cấu hình của tab %s\n"
            "(ví dụ: SCW480 – khách A):" % self.ten_tab,
            initialvalue=self.v_chon.get(), parent=self.khung)
        if ten is None:
            return
        ten = ten.strip()
        if not ten:
            messagebox.showwarning("Thiếu tên", "Tên bộ cấu hình không được "
                                                "để trống.")
            return
        d = _doc(self.khoa)
        if ten in d["bo"] and not messagebox.askyesno(
                "Ghi đè?", "Đã có bộ “%s”. Ghi đè bằng thông số đang điền?"
                           % ten):
            return
        try:
            d["bo"][ten] = self.lay()
            _ghi(self.khoa, d)
        except Exception as e:
            messagebox.showerror("Không lưu được", str(e))
            return
        self._nap_ds(chon=ten)
        self.lbl.configure(text="Đã lưu bộ “%s”." % ten)

    def xoa(self):
        ten = self.v_chon.get()
        if not ten:
            messagebox.showinfo("Chưa chọn", "Chọn bộ cấu hình cần xoá trong "
                                             "danh sách trước.")
            return
        if not messagebox.askyesno("Xoá bộ cấu hình",
                                   "Xoá bộ “%s”?\n(Các ô đang điền giữ "
                                   "nguyên.)" % ten):
            return
        d = _doc(self.khoa)
        d["bo"].pop(ten, None)
        try:
            _ghi(self.khoa, d)
        except Exception as e:
            messagebox.showerror("Không xoá được", str(e))
            return
        self.v_chon.set("")
        self._nap_ds()
        self.lbl.configure(text="Đã xoá bộ “%s”." % ten)

    # ------------------------------------------------ lay tu tab khac
    def _mo_menu_lay(self):
        self._dung_menu_lay()
        n = self.nut_lay
        try:
            self.menu.tk_popup(n.winfo_rootx(), n.winfo_rooty() + n.winfo_height())
        finally:
            self.menu.grab_release()

    def _dung_menu_lay(self):
        self.menu.delete(0, "end")
        cua_toi = _DANG_KY.get(self.khoa)
        co = False
        for khoa, (ten, ban_do) in _DANG_KY.items():
            if khoa == self.khoa or not cua_toi:
                continue
            chung = [r for r in ban_do if r in cua_toi[1]]
            if not chung:
                continue
            co = True
            self.menu.add_command(
                label="%s  (%s)" % (ten, ", ".join(TEN_VAI_TRO.get(r, r)
                                                   for r in chung)),
                command=lambda k=khoa: self._lay_tu(k))
        if not co:
            self.menu.add_command(label="(Không tab nào có ô chung với tab "
                                        "này)", state="disabled")

    def _lay_tu(self, khoa):
        _ten, nguon = _DANG_KY[khoa]
        _ten_toi, dich = _DANG_KY[self.khoa]
        da = []
        for r, v in nguon.items():
            if r in dich:
                g = v.get()
                if str(g).strip():
                    dich[r].set(g)
                    da.append(TEN_VAI_TRO.get(r, r))
        self.lbl.configure(text=("Đã lấy từ %s: %s." % (_ten, ", ".join(da)))
                           if da else "Tab %s chưa điền ô chung nào." % _ten)
