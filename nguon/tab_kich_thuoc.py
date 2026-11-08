# -*- coding: utf-8 -*-
"""
TAB KICH THUOC SPC - BUOC 1: TAO DU LIEU & XUAT BAO CAO TONG HOP
================================================================
Chon loai san pham -> Ma san pham -> Hien bang chi tieu voi Min/Max co the sua
-> Nhap so luong mau N (vd 25, 100, 1000...) -> Tinh toan sinh du lieu sao cho
moi cum 25 gia tri lien tuc (buoc nhay 1) thoa man day du tieu chi danh gia SPC
(I1..I8, MR1..MR2, Cpk, dung sai) ngay tu dau -> Xuat file Bao cao tong hop
('BaoCao_TongHop_<SP>_<Ma>.xlsx') va tu dong nap sang Buoc 2 Danh gia SPC.
"""

import os
import queue
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import bo_cau_hinh
import kich_thuoc_core as ktcore
import xuat_kich_thuoc as xkt


SO_DIEM_DO_MAC_DINH = {
    2: 3, 3: 3, 4: 6, 5: 6, 6: 6, 7: 2, 8: 3, 9: 2,
    10: 2, 11: 2, 12: 2, 13: 2, 14: 2, 15: 2, 16: 2,
}


def _fmt_so(v):
    if v is None:
        return "-"
    fv = float(v)
    return "%.1f" % fv if round(fv, 1) == round(fv, 4) else "%g" % round(fv, 4)


class TabKichThuoc(object):
    def __init__(self, me, mo_thu_muc_fn, khi_xong=None):
        """khi_xong(duong_dan_thu_muc): goi sau khi xuat xong Bao cao tong hop -
        tab cha dung de nap san thu muc vua sinh vao tab con Buoc 2 Danh gia SPC."""
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.khi_xong = khi_xong
        self.khi_luu_dung_sai = None
        self.file_cuoi = None
        self.file_tong_hop_cuoi = None
        self.thu_muc_ket_qua = None
        self.dang_dung = False
        self.hang_doi = queue.Queue()
        self.du_lieu_master = None
        self.danh_sach_chi_tieu_hien_tai = []

        self._dung_giao_dien()
        self._nap_du_lieu_master()
        bo_cau_hinh.dang_ky_thong_tin("kich_thuoc", "Kích thước", {
            "vat_lieu": self.v_vat_lieu})
        self.nap_cau_hinh(bo_cau_hinh.doc_lan_truoc("kich_thuoc"))

    # =========================================================== GIAO DIEN
    def _dung_giao_dien(self):
        ngoai = ttk.Frame(self.me)
        ngoai.pack(fill="both", expand=True)
        canvas = tk.Canvas(ngoai, bg=ui.COLOR_BG, highlightthickness=0)
        thanh = ttk.Scrollbar(ngoai, orient="vertical", command=canvas.yview)
        than = ttk.Frame(canvas)
        than.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cua = canvas.create_window((0, 0), window=than, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(cua, width=e.width))
        canvas.configure(yscrollcommand=thanh.set)
        canvas.pack(side="left", fill="both", expand=True)
        thanh.pack(side="right", fill="y")

        ui.cuon_bang_chuot(canvas)

        self.thanh = bo_cau_hinh.ThanhCauHinh(
            than, "kich_thuoc", "Kích thước", self.lay_cau_hinh,
            self.nap_cau_hinh)
        self._khoi_chon_san_pham(than)
        self._khoi_bang_chi_tieu(than)
        self._khoi_tuy_chon(than)
        self._khoi_chay(than)

    # ------------------------------------------------------- 1. Chon san pham
    def _khoi_chon_san_pham(self, me):
        c = ui.ModernCard(me, title="Chọn sản phẩm & file mẫu", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)
        k.columnconfigure(3, weight=1)

        ttk.Label(k, text="Loại sản phẩm", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        self.v_loai = tk.StringVar(value="")
        self.cb_loai = ttk.Combobox(k, textvariable=self.v_loai,
                                     state="readonly", width=25)
        self.cb_loai.grid(row=0, column=1, sticky="w", padx=(0, 20), ipady=2)
        self.cb_loai.bind("<<ComboboxSelected>>", self._khi_doi_loai)

        ttk.Label(k, text="Mã sản phẩm", style="Surface.TLabel").grid(
            row=0, column=2, sticky="w", padx=(0, 8))
        self.v_ma = tk.StringVar(value="")
        self.cb_ma = ttk.Combobox(k, textvariable=self.v_ma,
                                   state="readonly", width=25)
        self.cb_ma.grid(row=0, column=3, sticky="w", ipady=2)
        self.cb_ma.bind("<<ComboboxSelected>>", self._khi_doi_ma)

        # File dung sai
        ttk.Label(k, text="File dung sai", style="Surface.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
        self.v_file_ds = tk.StringVar(value="")
        ttk.Entry(k, textvariable=self.v_file_ds, width=60).grid(
            row=1, column=1, columnspan=2, sticky="ew", ipady=2, pady=(10, 0))
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=self._chon_file_dung_sai).grid(
            row=1, column=3, sticky="w", padx=(8, 0), pady=(10, 0))

        # File mau Bao cao tong hop
        ttk.Label(k, text="File mẫu Tổng hợp", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        self.v_file_mau = tk.StringVar(value="")
        ttk.Entry(k, textvariable=self.v_file_mau, width=60).grid(
            row=2, column=1, columnspan=2, sticky="ew", ipady=2, pady=(6, 0))
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=self._chon_file_mau).grid(
            row=2, column=3, sticky="w", padx=(8, 0), pady=(6, 0))

        self.lbl_sp = ttk.Label(
            k, style="SurfacePhu.TLabel",
            text="Chọn file 'Dung sai kích thước.xlsx' và mẫu 'Báo cáo tổng hợp - <Sản phẩm>.xlsx' để bắt đầu."
        )
        self.lbl_sp.grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))

    # ------------------------------------------------------- 2. Bang chi tieu
    def _khoi_bang_chi_tieu(self, me):
        c = ui.ModernCard(
            me,
            title="Danh sách chỉ tiêu kích thước (double-click sửa trực tiếp thông số)",
            badge_num=2,
        )
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        cols = ("stt", "ten", "ky_hieu", "yeu_cau", "lsl", "usl", "min", "max")
        col_widths = (40, 280, 65, 160, 70, 70, 70, 70)
        col_names = ("STT", "Nội dung kiểm tra", "Ký hiệu", "Yêu cầu kỹ thuật",
                     "LSL", "USL", "Min", "Max")

        khung_tree = tk.Frame(k, bg=ui.COLOR_SURFACE)
        khung_tree.pack(fill="both", expand=True, pady=(0, 4))

        self.tree = ttk.Treeview(khung_tree, columns=cols, show="headings",
                                  height=12, selectmode="browse")
        for col, w, name in zip(cols, col_widths, col_names):
            self.tree.heading(col, text=name)
            self.tree.column(col, width=w, minwidth=w, anchor="center")
        self.tree.column("ten", anchor="w")
        self.tree.column("yeu_cau", anchor="w")

        sb = ttk.Scrollbar(khung_tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._sua_o)

        thanh_duoi = tk.Frame(k, bg=ui.COLOR_SURFACE)
        thanh_duoi.pack(fill="x", pady=(2, 0))
        ttk.Label(
            thanh_duoi,
            text="💡 Double-click vào Yêu cầu kỹ thuật, LSL, USL, Min hoặc Max để sửa trực tiếp.",
            style="SurfacePhu.TLabel",
        ).pack(side="left")
        ttk.Button(
            thanh_duoi,
            text="💾 Lưu vào file Dung sai",
            style="Outline.TButton",
            command=self._luu_vao_file_dung_sai,
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            thanh_duoi,
            text="↺ Mặc định theo file Dung sai",
            style="Outline.TButton",
            command=self._khoi_phuc_tu_file_ds,
        ).pack(side="right")

    # ------------------------------------------------------- 3. Tuy chon
    def _khoi_tuy_chon(self, me):
        c = ui.ModernCard(me, title="Tùy chọn sinh dữ liệu", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        ttk.Label(k, text="Số lượng mẫu", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        self.v_so_mau = tk.StringVar(value="1000")
        ttk.Entry(k, textvariable=self.v_so_mau, width=10).grid(
            row=0, column=1, sticky="w", ipady=2)

        ttk.Label(k, text="STT bắt đầu", style="Surface.TLabel").grid(
            row=0, column=2, sticky="w", padx=(20, 8))
        self.v_stt_dau = tk.StringVar(value="1")
        ttk.Entry(k, textvariable=self.v_stt_dau, width=8).grid(
            row=0, column=3, sticky="w", ipady=2)

        ttk.Label(k, text="Cpk tối thiểu", style="Surface.TLabel").grid(
            row=0, column=4, sticky="w", padx=(20, 8))
        self.v_cpk = tk.StringVar(value="1.33")
        ttk.Entry(k, textvariable=self.v_cpk, width=8).grid(
            row=0, column=5, sticky="w", ipady=2)

        ttk.Label(k, text="Vật liệu", style="Surface.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        self.v_vat_lieu = tk.StringVar(value="SCW480")
        ttk.Entry(k, textvariable=self.v_vat_lieu, width=14).grid(
            row=1, column=1, sticky="w", ipady=2, pady=(8, 0))

        ttk.Label(k, text="Seed (trống = ngẫu nhiên)", style="Surface.TLabel").grid(
            row=1, column=2, sticky="w", padx=(20, 8), pady=(8, 0))
        self.v_seed = tk.StringVar(value="")
        ttk.Entry(k, textvariable=self.v_seed, width=10).grid(
            row=1, column=3, sticky="w", ipady=2, pady=(8, 0))

        ttk.Label(
            k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
            text="💡 Dữ liệu sinh ra được tính toán sẵn theo từng cụm 25 giá trị liên tục (bước nhảy 1: "
                 "1–25, 26–50...) thỏa mãn đầy đủ tiêu chí đánh giá SPC (I-MR, Cpk, dung sai). "
                 "Bước ① xuất file Báo cáo tổng hợp; Bước ② kiểm tra lại và xuất các file báo cáo đánh giá SPC."
        ).grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 0))

    # ------------------------------------------------------- 4. Chay
    def _khoi_chay(self, me):
        c = ui.ModernCard(me, title="Tạo dữ liệu & Xuất Báo cáo tổng hợp", badge_num=4)
        c.pack(fill="x", padx=12, pady=(4, 8))
        k = c.get_body()

        d = tk.Frame(k, bg=ui.COLOR_SURFACE)
        d.pack(fill="x")
        self.btn_chay = ttk.Button(d, text="▶ TẠO DỮ LIỆU & XUẤT BÁO CÁO TỔNG HỢP",
                                    style="Primary.TButton",
                                    command=self._chay)
        self.btn_chay.pack(side="left")

        self.btn_dung = ttk.Button(d, text="■ Dừng", style="Outline.TButton",
                                    command=self._dung, state="disabled")
        self.btn_dung.pack(side="left", padx=(8, 0))

        self.btn_mo = ttk.Button(d, text="📂 Mở thư mục kết quả",
                                  style="Outline.TButton",
                                  command=self._mo_ket_qua,
                                  state="disabled")
        self.btn_mo.pack(side="left", padx=(12, 0))

        self.lbl_tien_trinh = ttk.Label(k, text="", style="Surface.TLabel")
        self.lbl_tien_trinh.pack(anchor="w", pady=(8, 0))

        self.console = ui.ModernTerminal(k, title="NHẬT KÝ TẠO DỮ LIỆU KÍCH THƯỚC", height=10)
        self.console.pack(fill="both", expand=True, pady=(6, 0))

    # =========================================================== LOGIC
    def _nap_du_lieu_master(self):
        """Tu dong tim va nap file dung sai mac dinh."""
        thu_muc = bo_cau_hinh.base_dir()
        cac_cho = [
            bo_cau_hinh.thu_muc_file_mau(),
            thu_muc,
            os.path.join(thu_muc, "templates"),
            os.path.dirname(thu_muc),
            os.path.join(thu_muc, "File_Mau_Va_Dung_Sai"),
        ]
        for tm in cac_cho:
            if not os.path.exists(tm):
                continue
            ds = os.path.join(tm, "Dung sai kích thước.xlsx")
            if os.path.exists(ds):
                self.v_file_ds.set(ds)
                self._doc_file_dung_sai(ds)
                break

        if not self.v_file_mau.get():
            mau = xkt.tim_mau_tong_hop(self.v_loai.get())
            if mau:
                self.v_file_mau.set(mau)

    def _chon_file_dung_sai(self):
        path = filedialog.askopenfilename(
            title="Chọn file Dung sai kích thước",
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.v_file_ds.set(path)
            self._doc_file_dung_sai(path)

    def _chon_file_mau(self):
        path = filedialog.askopenfilename(
            title="Chọn file mẫu Báo cáo tổng hợp",
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.v_file_mau.set(path)

    def _doc_file_dung_sai(self, path):
        """Doc file dung sai va cap nhat combo box."""
        try:
            self.du_lieu_master = ktcore.doc_dung_sai(path)
            loai_list = list(self.du_lieu_master.keys())
            self.cb_loai["values"] = loai_list
            if loai_list:
                self.v_loai.set(loai_list[0])
                self._khi_doi_loai(None)
            self.lbl_sp.config(
                text="✅ Đã nạp %d loại sản phẩm từ file dung sai." % len(loai_list))
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", str(e))

    def _khi_doi_loai(self, event):
        """Khi doi loai san pham -> cap nhat danh sach ma va file mau Tong hop tuong ung."""
        loai = self.v_loai.get()
        if not loai or not self.du_lieu_master:
            return
        tm_hien_tai = os.path.dirname(self.v_file_mau.get()) if self.v_file_mau.get() else None
        mau_th = xkt.tim_mau_tong_hop(loai, tm_hien_tai)
        if mau_th and os.path.isfile(mau_th):
            self.v_file_mau.set(mau_th)

        data = self.du_lieu_master.get(loai, {})
        ma_list = data.get("ma_san_pham", [])
        self.cb_ma["values"] = ma_list
        if ma_list:
            self.v_ma.set(ma_list[0])
            self._khi_doi_ma(None)

    def _khi_doi_ma(self, event):
        """Khi doi ma san pham -> cap nhat bang chi tieu."""
        loai = self.v_loai.get()
        ma = self.v_ma.get()
        if not loai or not ma or not self.du_lieu_master:
            return

        data = self.du_lieu_master.get(loai, {})
        chi_tieu_list = data.get("chi_tieu", [])

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.danh_sach_chi_tieu_hien_tai = []
        so_ct = 0
        for ct in chi_tieu_list:
            info = ct["theo_ma"].get(ma, {})
            if not info.get("yeu_cau"):
                continue

            lsl = info.get("lsl")
            usl = info.get("usl")
            min_v = info.get("min")
            max_v = info.get("max")

            lsl_str = _fmt_so(lsl)
            usl_str = _fmt_so(usl)
            min_str = _fmt_so(min_v)
            max_str = _fmt_so(max_v)

            self.tree.insert("", "end", values=(
                ct["stt"], ct["ten"].replace("\n", " "), ct["ky_hieu"],
                info["yeu_cau"].replace("\n", " "), lsl_str, usl_str, min_str, max_str
            ))

            self.danh_sach_chi_tieu_hien_tai.append({
                "stt": ct["stt"],
                "ten": ct["ten"],
                "ky_hieu": ct["ky_hieu"],
                "ky_hieu_diem": ct.get("ky_hieu_diem"),
                "so_diem_do": ct.get("so_diem_do")
                or SO_DIEM_DO_MAC_DINH.get(ct["stt"], 2),
                "min": min_v,
                "max": max_v,
                "lsl": lsl,
                "usl": usl,
                "yeu_cau": info["yeu_cau"],
                "chuan": info.get("chuan"),
                "ds_duoi": info.get("ds_duoi"),
                "ds_tren": info.get("ds_tren"),
                "_goc": {
                    "yeu_cau": info["yeu_cau"],
                    "lsl": lsl,
                    "usl": usl,
                    "min": min_v,
                    "max": max_v,
                },
            })
            so_ct += 1

        self.lbl_sp.config(
            text="📋 %s / %s — Tổng cộng %d chỉ tiêu kích thước."
                 % (loai, ma, so_ct))

    def min_max_hien_tai(self, loai, ma):
        """Tra ve {stt: (min, max)} dang co tren bang neu cung loai & ma san pham."""
        if (str(loai or "").strip().lower() == str(self.v_loai.get() or "").strip().lower()
                and str(ma or "").strip() == str(self.v_ma.get() or "").strip()):
            return {ct["stt"]: (ct.get("min"), ct.get("max"))
                    for ct in self.danh_sach_chi_tieu_hien_tai}
        return {}

    def thong_so_hien_tai(self, loai, ma):
        """Tra ve {stt: dict thong so} dang co tren bang neu cung loai & ma san pham."""
        if (str(loai or "").strip().lower() == str(self.v_loai.get() or "").strip().lower()
                and str(ma or "").strip() == str(self.v_ma.get() or "").strip()):
            return {
                ct["stt"]: {
                    "yeu_cau": ct.get("yeu_cau"),
                    "chuan": ct.get("chuan"),
                    "ds_duoi": ct.get("ds_duoi"),
                    "ds_tren": ct.get("ds_tren"),
                    "lsl": ct.get("lsl"),
                    "usl": ct.get("usl"),
                    "min": ct.get("min"),
                    "max": ct.get("max"),
                }
                for ct in self.danh_sach_chi_tieu_hien_tai
            }
        return {}

    def _sua_o(self, event):
        """Cho phep double-click sua truc tiep Yeu cau ky thuat, LSL, USL, Min, Max."""
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return

        col_idx = int(col.replace("#", ""))
        if col_idx not in (4, 5, 6, 7, 8):
            return

        bbox = self.tree.bbox(item, col)
        if not bbox:
            return
        x, y, w, h = bbox
        vals = list(self.tree.item(item, "values"))
        old_val = vals[col_idx - 1]

        entry = tk.Entry(self.tree, justify="left" if col_idx == 4 else "center")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, "" if old_val == "-" and col_idx in (5, 6) else old_val)
        entry.select_range(0, "end")
        entry.focus_set()
        da_xu_ly = [False]

        def xac_nhan(evt=None):
            if da_xu_ly[0]:
                return
            da_xu_ly[0] = True
            new_val = entry.get().strip()
            entry.destroy()

            stt = int(vals[0])
            ct_chon = next((c for c in self.danh_sach_chi_tieu_hien_tai if c["stt"] == stt), None)
            if ct_chon is None:
                return

            if col_idx == 4:
                if not new_val:
                    return
                if " ".join(new_val.split()) == " ".join(str(ct_chon.get("yeu_cau") or "").split()):
                    return
                ct_chon["yeu_cau"] = new_val
                chuan, ds_duoi, ds_tren = ktcore._phan_tich_dung_sai(new_val)
                try:
                    if chuan is None:
                        chuan = round(float(new_val.replace(",", ".")), 4)
                except ValueError:
                    pass
                lsl = round(chuan + ds_duoi, 4) if chuan is not None and ds_duoi is not None else None
                usl = round(chuan + ds_tren, 4) if chuan is not None and ds_tren is not None else None
                ct_chon["chuan"] = chuan
                ct_chon["ds_duoi"] = ds_duoi
                ct_chon["ds_tren"] = ds_tren
                ct_chon["lsl"] = lsl
                ct_chon["usl"] = usl
                vals[3] = new_val
                vals[4] = _fmt_so(lsl)
                vals[5] = _fmt_so(usl)
                self.tree.item(item, values=vals)
            elif col_idx in (5, 6):
                khoa = "lsl" if col_idx == 5 else "usl"
                khoa_ds = "ds_duoi" if col_idx == 5 else "ds_tren"
                if new_val in ("", "-"):
                    ct_chon[khoa] = None
                    ct_chon[khoa_ds] = None
                    vals[col_idx - 1] = "-"
                    self.tree.item(item, values=vals)
                else:
                    try:
                        fv = ktcore.spc_luat.so_huu_han(new_val.replace(",", "."), khoa.upper())
                        ct_chon[khoa] = fv
                        if ct_chon.get("chuan") is not None:
                            ct_chon[khoa_ds] = round(fv - ct_chon["chuan"], 4)
                        vals[col_idx - 1] = _fmt_so(fv)
                        self.tree.item(item, values=vals)
                    except ValueError:
                        pass
            else:
                khoa = "min" if col_idx == 7 else "max"
                try:
                    fv = ktcore.spc_luat.so_huu_han(new_val.replace(",", "."), "Min/Max")
                    ct_chon[khoa] = fv
                    vals[col_idx - 1] = _fmt_so(fv)
                    self.tree.item(item, values=vals)
                except ValueError:
                    pass

        def huy(evt=None):
            if da_xu_ly[0]:
                return
            da_xu_ly[0] = True
            entry.destroy()

        entry.bind("<Return>", xac_nhan)
        entry.bind("<Escape>", huy)
        entry.bind("<FocusOut>", xac_nhan)

    def _khoi_phuc_tu_file_ds(self):
        """Nap lai thong so goc tu file Dung sai kich thuoc.xlsx."""
        path = self.v_file_ds.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Chưa có file dung sai", "Không tìm thấy file Dung sai kích thước.xlsx.")
            return
        try:
            self.du_lieu_master = ktcore.doc_dung_sai(path)
            self._khi_doi_ma(None)
            self.lbl_sp.config(
                text="↺ Đã khôi phục thông số mặc định từ file Dung sai (%s / %s)."
                     % (self.v_loai.get(), self.v_ma.get())
            )
        except Exception as e:
            messagebox.showerror("Lỗi đọc file dung sai", str(e))

    def _luu_vao_file_dung_sai(self):
        """Ghi cac thong so da sua tren bang (Yeu cau ky thuat, Min, Max) vao file Dung sai."""
        path = self.v_file_ds.get().strip()
        loai = self.v_loai.get().strip()
        ma = self.v_ma.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Chưa có file dung sai", "Không tìm thấy file Dung sai kích thước.xlsx.")
            return
        if not loai or not ma or not self.danh_sach_chi_tieu_hien_tai:
            messagebox.showwarning("Chưa chọn sản phẩm", "Hãy chọn Loại sản phẩm và Mã sản phẩm trước.")
            return

        rieng_lsl_usl = [
            str(ct["stt"]) for ct in self.danh_sach_chi_tieu_hien_tai
            if ct.get("_goc")
            and ct.get("yeu_cau") == ct["_goc"].get("yeu_cau")
            and (ct.get("lsl") != ct["_goc"].get("lsl") or ct.get("usl") != ct["_goc"].get("usl"))
        ]
        ghi_chu_them = ""
        if rieng_lsl_usl:
            ghi_chu_them = (
                "\n\nLưu ý: STT %s đang sửa riêng LSL/USL mà chưa sửa cột 'Yêu cầu kỹ thuật'. "
                "File Excel Dung sai tính LSL/USL từ 'Yêu cầu kỹ thuật' (nếu chỉ muốn dùng LSL/USL riêng, "
                "bạn có thể lưu vào 'Bộ cấu hình' ở thanh trên cùng)." % ", ".join(rieng_lsl_usl)
            )

        if not messagebox.askyesno(
            "Lưu vào file Dung sai kích thước",
            "Ghi các thông số đang hiển thị của '%s / %s' vào file:\n%s\n\n"
            "(Phần mềm sẽ tự động sao lưu bản gốc vào thư mục _truoc_khi_sua trước khi ghi.)%s"
            % (loai, ma, path, ghi_chu_them),
        ):
            return

        try:
            so_o = ktcore.ghi_dung_sai(
                path, loai, ma, self.danh_sach_chi_tieu_hien_tai, luu_goc=True
            )
            if so_o == 0:
                messagebox.showinfo(
                    "Không có thay đổi",
                    "Các cột Yêu cầu kỹ thuật, Min, Max hiện tại đã trùng khớp với file Dung sai."
                    + ghi_chu_them,
                )
                return
            self.du_lieu_master = ktcore.doc_dung_sai(path)
            self._khi_doi_ma(None)
            if self.khi_luu_dung_sai:
                try:
                    self.khi_luu_dung_sai()
                except Exception:
                    pass
            self.lbl_sp.config(
                text="💾 Đã lưu %d ô thay đổi vào file Dung sai (%s / %s)." % (so_o, loai, ma)
            )
            messagebox.showinfo(
                "Đã lưu thành công",
                "Đã cập nhật %d ô thông số của '%s / %s' vào file Dung sai kích thước.xlsx.\n"
                "(Bản gốc trước khi sửa đã được lưu dự phòng trong thư mục _truoc_khi_sua.)"
                % (so_o, loai, ma),
            )
        except PermissionError:
            messagebox.showerror(
                "File đang mở",
                "Không ghi được vì file 'Dung sai kích thước.xlsx' đang mở trong Excel.\n"
                "Hãy đóng file Excel rồi bấm Lưu lại.",
            )
        except Exception as e:
            messagebox.showerror("Lỗi lưu file dung sai", str(e))

    # =========================================================== CHAY
    def _dung(self):
        self.dang_dung = True
        self.btn_dung.config(state="disabled")
        self.lbl_tien_trinh.config(text="Đang dừng...")

    def _chay(self):
        """Bat dau sinh du lieu N mau va xuat Bao cao tong hop."""
        if not self.danh_sach_chi_tieu_hien_tai:
            messagebox.showwarning("Chưa có dữ liệu",
                                    "Hãy chọn Loại và Mã sản phẩm trước.")
            return

        for ct in self.danh_sach_chi_tieu_hien_tai:
            if ct["min"] is None or ct["max"] is None:
                messagebox.showwarning(
                    "Thiếu Min/Max",
                    "Chỉ tiêu STT %d (%s) chưa có Min hoặc Max.\n"
                    "Hãy double-click vào ô để nhập." % (ct["stt"], ct["ten"]))
                return

        try:
            so_mau = ktcore.spc_luat.so_nguyen(self.v_so_mau.get().strip(), "Số lượng mẫu", 2)
            stt_dau = ktcore.spc_luat.so_nguyen(self.v_stt_dau.get().strip(), "STT bắt đầu", 1)
            cpk = ktcore.spc_luat.kiem_tra_nguong(
                None, None, str(self.v_cpk.get()).strip().replace(",", ".")
            )[2]
        except ValueError as e:
            messagebox.showwarning("Tham số chưa hợp lệ", str(e))
            return

        seed = None
        if self.v_seed.get().strip():
            try:
                seed = int(self.v_seed.get().strip())
            except ValueError:
                messagebox.showwarning("Sai Seed", "Seed phải là số nguyên (hoặc để trống).")
                return

        bo_cau_hinh.nho_lan_truoc("kich_thuoc", self.lay_cau_hinh())
        self.dang_dung = False
        self.btn_chay.config(state="disabled")
        self.btn_dung.config(state="normal")
        so_nhom = len(ktcore.mac_dinh_bat_dau(so_mau))

        self.console.clear_log()
        self.console.ghi("Bắt đầu tính toán & tạo dữ liệu kích thước chuẩn SPC...\n", tag="accent")
        self.console.ghi("Sản phẩm : %s / %s\n" % (self.v_loai.get(), self.v_ma.get()))
        self.console.ghi("Quy mô   : %d mẫu (STT %d..%d) — tương ứng %d dãy 25 mẫu liên tiếp (bước nhảy 1)\n"
                         % (so_mau, stt_dau, stt_dau + so_mau - 1, so_nhom))
        self.console.ghi("Chỉ tiêu : %d chỉ tiêu | Cpk tối thiểu: %.2f\n"
                         % (len(self.danh_sach_chi_tieu_hien_tai), cpk))
        self.console.ghi("-" * 68 + "\n")
        self.lbl_tien_trinh.config(text="Đang sinh dữ liệu %d mẫu chuẩn SPC..." % so_mau)

        t = threading.Thread(
            target=self._sinh_trong_nen, args=(so_mau, stt_dau, seed, cpk), daemon=True
        )
        t.start()
        self._poll()

    def _sinh_trong_nen(self, so_mau, stt_dau, seed, cpk):
        """Chay trong thread nen."""
        try:
            def cb(stt, ten, dat, so_cum_dat, msg):
                icon = "✅" if dat else "❌"
                ten_gon = str(ten).replace("\n", " ")[:38]
                self.hang_doi.put(("log", "%s STT %02d: %-38s — %s"
                                   % (icon, stt, ten_gon, msg), "success" if dat else "error"))

            ket_qua = ktcore.sinh_toan_bo(
                self.danh_sach_chi_tieu_hien_tai,
                so_mau=so_mau,
                callback=cb,
                seed=seed,
                cpk_toi_thieu=cpk if cpk is not None else ktcore.CPK_TOI_THIEU,
                dung=lambda: self.dang_dung,
            )

            self.hang_doi.put(("log", "-" * 68, "normal"))
            self.hang_doi.put(("log", "Đang ghi file Báo cáo tổng hợp (%d mẫu)..." % so_mau, "accent"))

            loai = self.v_loai.get()
            ma = self.v_ma.get()
            thu_muc = xkt.tao_thu_muc_xuat(san_pham=loai, ma=ma, so_mau=so_mau)
            file_mau_th = self.v_file_mau.get().strip() or None
            if file_mau_th and ("JIS.QT.07" in os.path.basename(file_mau_th) or not os.path.isfile(file_mau_th)):
                file_mau_th = xkt.tim_mau_tong_hop(loai)

            sp_c = loai.replace(" ", "_")
            ma_c = ma.replace(" ", "_")
            ten_file_th = "BaoCao_TongHop_%s_%s.xlsx" % (sp_c, ma_c)
            duong_dan_th = os.path.join(thu_muc, ten_file_th)

            thong_tin = {
                "san_pham": loai,
                "kich_thuoc": ma,
                "vat_lieu": self.v_vat_lieu.get(),
                "stt_dau": stt_dau,
                "stt_mau": [stt_dau + j for j in range(so_mau)],
            }

            xkt.xuat_bao_cao_tong_hop(ket_qua, thong_tin, duong_dan_th, duong_dan_mau_th=file_mau_th)
            self.file_tong_hop_cuoi = duong_dan_th
            self.file_cuoi = thu_muc

            so_dat = sum(1 for r in ket_qua if r["spc_result"].dat)
            so_nhom = len(ktcore.mac_dinh_bat_dau(so_mau))
            self.hang_doi.put(("log", "=" * 68, "normal"))
            self.hang_doi.put((
                "log",
                "HOÀN TẤT TẠO DỮ LIỆU: %d/%d chỉ tiêu đạt chuẩn SPC trên toàn bộ %d dãy 25 mẫu.\n"
                "  ➜ Báo cáo tổng hợp: %s\n"
                "  ➜ Chuyển sang tab '② Đánh giá SPC' để kiểm tra lại và xuất %d file báo cáo đánh giá ổn định."
                % (so_dat, len(ket_qua), so_nhom, ten_file_th, so_nhom),
                "success",
            ))

            self.hang_doi.put(("xong", thu_muc, so_dat, len(ket_qua), so_mau, so_nhom))

        except InterruptedError as e:
            self.hang_doi.put(("dung", str(e)))
        except Exception:
            self.hang_doi.put(("loi", traceback.format_exc()))

    def _poll(self):
        """Kiem tra hang doi tu thread nen."""
        try:
            while True:
                msg = self.hang_doi.get_nowait()
                if msg[0] == "log":
                    loai = msg[2] if len(msg) > 2 else "normal"
                    self.console.ghi(msg[1] + "\n", tag=loai)
                elif msg[0] == "dung":
                    self.btn_chay.config(state="normal")
                    self.btn_dung.config(state="disabled")
                    self.lbl_tien_trinh.config(text="⏹ %s" % msg[1])
                    self.console.ghi("⏹ %s\n" % msg[1], tag="warning")
                    return
                elif msg[0] == "xong":
                    self.thu_muc_ket_qua = msg[1]
                    self.btn_chay.config(state="normal")
                    self.btn_dung.config(state="disabled")
                    self.btn_mo.config(state="normal")
                    dat, tong, so_mau, so_nhom = msg[2], msg[3], msg[4], msg[5]
                    self.lbl_tien_trinh.config(
                        text="✅ Hoàn tất! Đã xuất Báo cáo tổng hợp (%d mẫu) tại: %s" % (so_mau, msg[1])
                    )
                    if self.khi_xong and self.thu_muc_ket_qua:
                        try:
                            self.khi_xong(self.thu_muc_ket_qua)
                            self.console.ghi(
                                "➜ Đã nạp sẵn thư mục vừa tạo sang bước ② Đánh giá SPC.\n",
                                tag="accent",
                            )
                        except Exception:
                            self.console.ghi(
                                "(Không nạp được sang Đánh giá SPC:\n%s)\n"
                                % traceback.format_exc(),
                                tag="error",
                            )
                    ket = "%d/%d chỉ tiêu đạt chuẩn SPC trên cả %d dãy 25 mẫu liên tiếp.\n" % (
                        dat, tong, so_nhom
                    )
                    if messagebox.askyesno(
                        "Hoàn tất tạo dữ liệu",
                        "Đã xuất Báo cáo tổng hợp (%d mẫu):\n%s\n"
                        "Thư mục đã được nạp sẵn sang bước '② Đánh giá SPC' để xuất %d file đánh giá.\n\n"
                        "Mở thư mục kết quả bây giờ?" % (so_mau, ket, so_nhom),
                    ):
                        self._mo_ket_qua()
                    return
                elif msg[0] == "loi":
                    self.console.ghi("LỖI:\n" + msg[1] + "\n", tag="error")
                    self.btn_chay.config(state="normal")
                    self.btn_dung.config(state="disabled")
                    self.lbl_tien_trinh.config(text="❌ Có lỗi xảy ra.")
                    messagebox.showerror("Lỗi", "Có lỗi khi sinh dữ liệu kích thước. "
                                                "Chi tiết ở nhật ký bên dưới.")
                    return
        except queue.Empty:
            pass
        self.me.after(100, self._poll)

    def _mo_ket_qua(self):
        """Mo thu muc ket qua."""
        if self.thu_muc_ket_qua and os.path.isdir(self.thu_muc_ket_qua):
            if self.mo_thu_muc_ngoai:
                self.mo_thu_muc_ngoai(self.thu_muc_ket_qua)
            else:
                import subprocess
                subprocess.Popen(["explorer", self.thu_muc_ket_qua])

    # =========================================================== LUU/NAP
    def lay_cau_hinh(self):
        """Tra ve dict cau hinh de luu - chi luu cac chi tieu nguoi dung da chinh sua."""
        tuy_chinh = {}
        for ct in self.danh_sach_chi_tieu_hien_tai:
            goc = ct.get("_goc") or {}
            khac = (
                ct.get("yeu_cau") != goc.get("yeu_cau")
                or ct.get("lsl") != goc.get("lsl")
                or ct.get("usl") != goc.get("usl")
                or ct.get("min") != goc.get("min")
                or ct.get("max") != goc.get("max")
            )
            if khac:
                tuy_chinh[str(ct["stt"])] = {
                    "yeu_cau": ct.get("yeu_cau"),
                    "chuan": ct.get("chuan"),
                    "ds_duoi": ct.get("ds_duoi"),
                    "ds_tren": ct.get("ds_tren"),
                    "lsl": ct.get("lsl"),
                    "usl": ct.get("usl"),
                    "min": ct.get("min"),
                    "max": ct.get("max"),
                }
        return {
            "file_dung_sai": self.v_file_ds.get(),
            "file_mau_th": self.v_file_mau.get(),
            "loai": self.v_loai.get(),
            "ma": self.v_ma.get(),
            "so_mau": self.v_so_mau.get(),
            "stt_dau": self.v_stt_dau.get(),
            "vat_lieu": self.v_vat_lieu.get(),
            "cpk": self.v_cpk.get(),
            "seed": self.v_seed.get(),
            "thong_so_tuy_chinh": tuy_chinh,
        }

    def nap_cau_hinh(self, cfg):
        """Nap cau hinh da luu."""
        if not cfg:
            return
        if cfg.get("file_dung_sai") and os.path.exists(cfg["file_dung_sai"]) \
                and cfg["file_dung_sai"] != self.v_file_ds.get():
            self.v_file_ds.set(cfg["file_dung_sai"])
            self._doc_file_dung_sai(cfg["file_dung_sai"])
        if cfg.get("loai"):
            self.v_loai.set(cfg["loai"])
            self._khi_doi_loai(None)
        f_th = cfg.get("file_mau_th")
        if f_th and os.path.exists(f_th) and "JIS.QT.07" not in os.path.basename(f_th):
            loai = self.v_loai.get()
            if not loai or loai.lower() in os.path.basename(f_th).lower():
                self.v_file_mau.set(f_th)
        if cfg.get("ma"):
            self.v_ma.set(cfg["ma"])
            self._khi_doi_ma(None)
        if cfg.get("so_mau"):
            self.v_so_mau.set(str(cfg["so_mau"]))
        if cfg.get("stt_dau"):
            self.v_stt_dau.set(str(cfg["stt_dau"]))
        if cfg.get("vat_lieu"):
            self.v_vat_lieu.set(cfg["vat_lieu"])
        if cfg.get("cpk"):
            self.v_cpk.set(cfg["cpk"])
        if cfg.get("seed") is not None:
            self.v_seed.set(str(cfg["seed"]))

        cung_sp_ma = (
            (not cfg.get("loai") or cfg.get("loai") == self.v_loai.get())
            and (not cfg.get("ma") or cfg.get("ma") == self.v_ma.get())
        )
        if cung_sp_ma:
            if "thong_so_tuy_chinh" in cfg:
                self._ap_thong_so_tuy_chinh(cfg.get("thong_so_tuy_chinh") or {})
            elif cfg.get("min_max"):
                self._ap_min_max(cfg.get("min_max") or {})

    def _ap_thong_so_tuy_chinh(self, tc):
        """Dat lai cac thong so nguoi dung da chinh sua (theo STT) vao bang va danh sach."""
        if not tc:
            return
        for ct in self.danh_sach_chi_tieu_hien_tai:
            g = tc.get(str(ct["stt"]))
            if not isinstance(g, dict):
                continue
            for k in ("yeu_cau", "chuan", "ds_duoi", "ds_tren", "lsl", "usl", "min", "max"):
                if k in g:
                    ct[k] = g[k]
        map_ct = {str(ct["stt"]): ct for ct in self.danh_sach_chi_tieu_hien_tai}
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            key = str(vals[0])
            if key in tc and key in map_ct:
                ct = map_ct[key]
                vals[3] = str(ct.get("yeu_cau") or "").replace("\n", " ")
                vals[4] = _fmt_so(ct.get("lsl"))
                vals[5] = _fmt_so(ct.get("usl"))
                vals[6] = _fmt_so(ct.get("min"))
                vals[7] = _fmt_so(ct.get("max"))
                self.tree.item(item, values=vals)

    def _ap_min_max(self, mm):
        """Dat lai Min/Max da sua (theo STT) vao bang va danh sach (tuong thich cau hinh cu)."""
        if not mm:
            return
        for ct in self.danh_sach_chi_tieu_hien_tai:
            g = mm.get(str(ct["stt"]))
            if g and len(g) == 2 and g[0] is not None and g[1] is not None:
                ct["min"], ct["max"] = g[0], g[1]
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            g = mm.get(str(vals[0]))
            if g and len(g) == 2 and g[0] is not None and g[1] is not None:
                vals[6] = _fmt_so(g[0])
                vals[7] = _fmt_so(g[1])
                self.tree.item(item, values=vals)
