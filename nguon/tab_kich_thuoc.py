# -*- coding: utf-8 -*-
"""
TAB KICH THUOC SPC - GIAO DIEN SINH DU LIEU KICH THUOC
=======================================================
Chon loai san pham -> Ma san pham -> Hien bang chi tieu voi Min/Max
co the sua -> Sinh du lieu 25 mau cho tat ca chi tieu -> Xuat file
Excel bao cao SPC chuan JIS.QT.07.BM.01.
"""

import os
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import bo_cau_hinh
import kich_thuoc_core as ktcore
import xuat_kich_thuoc as xkt


# So diem do moi chi tieu nay doc thang tu file dung sai (kich_thuoc_core).
# Bang duoi chi con la du phong khi file dung sai khong cho biet.
SO_DIEM_DO_MAC_DINH = {
    1: 3,   # Dw - duong kinh canh: do 3 diem
    2: 3,   # D - duong kinh truc: do 3 diem
    3: 6,   # t21 - chieu day canh goc: do 6 diem
    4: 6,   # t22 - chieu day canh ngoai: do 6 diem
    5: 6,   # t1 - chieu day truc: do 6 diem
    6: 2,   # nodo - chieu day co: do 2 diem
    7: 3,   # t3 - chieu day tam lap rap: do 3 diem
    8: 2,   # PL - chieu cao truc: do 2 diem
    9: 2,   # H1 - chieu cao tam lap rap: do 2 diem
    10: 2,  # H2: do 2 diem
    11: 2,  # H3: do 2 diem
    12: 2,  # H4: do 2 diem
    13: 2,  # X: do 2 diem
    14: 2,  # L1: do 2 diem
    15: 2,  # L2: do 2 diem
    16: 2,  # mac dinh
    17: 2,
    18: 2,
    19: 2,
    20: 2,
    21: 2,
    22: 2,
}


class TabKichThuoc(object):
    def __init__(self, me, mo_thu_muc_fn, khi_xong=None):
        """khi_xong(duong_dan_file): goi sau khi xuat xong bao cao - tab cha
        dung de nap san file vua sinh vao tab con Danh gia SPC."""
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.khi_xong = khi_xong
        self.file_cuoi = None
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
        c = ui.ModernCard(me, title="Chọn sản phẩm", badge_num=1)
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

        # File mau bao cao
        ttk.Label(k, text="File mẫu báo cáo", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        self.v_file_mau = tk.StringVar(value="")
        ttk.Entry(k, textvariable=self.v_file_mau, width=60).grid(
            row=2, column=1, columnspan=2, sticky="ew", ipady=2, pady=(6, 0))
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=self._chon_file_mau).grid(
            row=2, column=3, sticky="w", padx=(8, 0), pady=(6, 0))

        self.lbl_sp = ttk.Label(k, style="SurfacePhu.TLabel",
                                text="Chọn file 'Dung sai kích thước.xlsx' "
                                     "và file mẫu 'JIS.QT.07.BM.01...' để bắt đầu.")
        self.lbl_sp.grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))

    # ------------------------------------------------------- 2. Bang chi tieu
    def _khoi_bang_chi_tieu(self, me):
        c = ui.ModernCard(me, title="Danh sách chỉ tiêu kích thước (sửa Min/Max)", badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        cols = ("stt", "ten", "ky_hieu", "yeu_cau", "lsl", "usl", "min", "max")
        col_widths = (35, 260, 60, 160, 70, 70, 70, 70)
        col_names = ("STT", "Tên kích thước", "Ký hiệu", "Yêu cầu kỹ thuật",
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

        # Cho phep click doi vao o Min/Max de sua
        self.tree.bind("<Double-1>", self._sua_o)

        ttk.Label(k, text="💡 Double-click vào ô Min hoặc Max để sửa giá trị.",
                  style="SurfacePhu.TLabel").pack(anchor="w")

    # ------------------------------------------------------- 3. Tuy chon
    def _khoi_tuy_chon(self, me):
        c = ui.ModernCard(me, title="Tùy chọn", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        k.columnconfigure(1, weight=1)

        ttk.Label(k, text="Vật liệu", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        self.v_vat_lieu = tk.StringVar(value="SCW480")
        ttk.Entry(k, textvariable=self.v_vat_lieu, width=20).grid(
            row=0, column=1, sticky="w", ipady=2)

        ttk.Label(k, text="Cpk tối thiểu", style="Surface.TLabel").grid(
            row=0, column=2, sticky="w", padx=(20, 8))
        self.v_cpk = tk.StringVar(value="1.33")
        ttk.Entry(k, textvariable=self.v_cpk, width=10).grid(
            row=0, column=3, sticky="w", ipady=2)

        ttk.Label(k, text="Seed (để trống = ngẫu nhiên)", style="Surface.TLabel").grid(
            row=0, column=4, sticky="w", padx=(20, 8))
        self.v_seed = tk.StringVar(value="")
        ttk.Entry(k, textvariable=self.v_seed, width=10).grid(
            row=0, column=5, sticky="w", ipady=2)

    # ------------------------------------------------------- 4. Chay
    def _khoi_chay(self, me):
        c = ui.ModernCard(me, title="Sinh dữ liệu & Xuất báo cáo", badge_num=4)
        c.pack(fill="x", padx=12, pady=(4, 8))
        k = c.get_body()

        d = tk.Frame(k, bg=ui.COLOR_SURFACE)
        d.pack(fill="x")
        self.btn_chay = ttk.Button(d, text="▶ SINH DỮ LIỆU & XUẤT BÁO CÁO",
                                    style="Primary.TButton",
                                    command=self._chay)
        self.btn_chay.pack(side="left")

        self.btn_mo = ttk.Button(d, text="📂 Mở thư mục kết quả",
                                  style="Outline.TButton",
                                  command=self._mo_ket_qua,
                                  state="disabled")
        self.btn_mo.pack(side="left", padx=(12, 0))

        self.lbl_tien_trinh = ttk.Label(k, text="", style="Surface.TLabel")
        self.lbl_tien_trinh.pack(anchor="w", pady=(8, 0))

        # Console
        self.console = ui.ModernTerminal(k, title="NHẬT KÝ SINH DỮ LIỆU KÍCH THƯỚC", height=10)
        self.console.pack(fill="both", expand=True, pady=(6, 0))

        self.thu_muc_ket_qua = None

    # =========================================================== LOGIC
    def _nap_du_lieu_master(self):
        """Tu dong tim va nap file dung sai mac dinh."""
        import charpy_core as core
        thu_muc = core.base_dir()
        cac_cho = [
            bo_cau_hinh.thu_muc_file_mau(),     # file mau giao kem phan mem
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

        for tm in cac_cho:
            if not os.path.exists(tm):
                continue
            for f in sorted(os.listdir(tm)):
                if f.startswith("JIS.QT.07") and f.endswith(".xlsx"):
                    self.v_file_mau.set(os.path.join(tm, f))
                    break
            if self.v_file_mau.get():
                break

    def _chon_file_dung_sai(self):
        path = filedialog.askopenfilename(
            title="Chọn file Dung sai kích thước",
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.v_file_ds.set(path)
            self._doc_file_dung_sai(path)

    def _chon_file_mau(self):
        path = filedialog.askopenfilename(
            title="Chọn file mẫu báo cáo JIS.QT.07.BM.01",
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
        """Khi doi loai san pham -> cap nhat danh sach ma."""
        loai = self.v_loai.get()
        if not loai or not self.du_lieu_master:
            return
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

        # Xoa bang cu
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

            lsl_str = "%.1f" % lsl if lsl is not None else "-"
            usl_str = "%.1f" % usl if usl is not None else "-"
            min_str = "%.1f" % min_v if min_v is not None else "-"
            max_str = "%.1f" % max_v if max_v is not None else "-"

            self.tree.insert("", "end", values=(
                ct["stt"], ct["ten"], ct["ky_hieu"],
                info["yeu_cau"], lsl_str, usl_str, min_str, max_str
            ))

            # So diem do va ky hieu tung diem lay tu file dung sai (moi chi
            # tieu la mot khoi nhieu hang trong file do)
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
            })
            so_ct += 1

        self.lbl_sp.config(
            text="📋 %s / %s — Tổng cộng %d chỉ tiêu kích thước."
                 % (loai, ma, so_ct))

    def _sua_o(self, event):
        """Cho phep double-click sua gia tri Min hoac Max trong Treeview."""
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return

        # Chi cho sua cot Min (#7) va Max (#8)
        col_idx = int(col.replace("#", ""))
        if col_idx not in (7, 8):
            return

        # Lay gia tri hien tai
        vals = list(self.tree.item(item, "values"))
        old_val = vals[col_idx - 1]

        # Tao popup Entry de sua
        x, y, w, h = self.tree.bbox(item, col)
        entry = tk.Entry(self.tree, justify="center")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, old_val)
        entry.select_range(0, "end")
        entry.focus_set()

        def xac_nhan(evt=None):
            new_val = entry.get().strip()
            entry.destroy()
            try:
                fv = ktcore.spc_luat.so_huu_han(new_val.replace(",", "."), "Min/Max")
                vals[col_idx - 1] = "%.1f" % fv
                self.tree.item(item, values=vals)
                # Cap nhat danh_sach_chi_tieu
                stt = int(vals[0])
                for ct in self.danh_sach_chi_tieu_hien_tai:
                    if ct["stt"] == stt:
                        if col_idx == 7:
                            ct["min"] = fv
                        else:
                            ct["max"] = fv
                        break
            except ValueError:
                pass

        def huy(evt=None):
            entry.destroy()

        entry.bind("<Return>", xac_nhan)
        entry.bind("<Escape>", huy)
        entry.bind("<FocusOut>", xac_nhan)

    # =========================================================== CHAY
    def _chay(self):
        """Bat dau sinh du lieu va xuat bao cao."""
        if not self.danh_sach_chi_tieu_hien_tai:
            messagebox.showwarning("Chưa có dữ liệu",
                                    "Hãy chọn Loại và Mã sản phẩm trước.")
            return

        file_mau = self.v_file_mau.get()
        if not file_mau or not os.path.exists(file_mau):
            messagebox.showwarning("Chưa chọn file mẫu",
                                    "Hãy chọn file mẫu báo cáo JIS.QT.07.BM.01.")
            return

        # Kiem tra min/max
        for ct in self.danh_sach_chi_tieu_hien_tai:
            if ct["min"] is None or ct["max"] is None:
                messagebox.showwarning(
                    "Thiếu Min/Max",
                    "Chỉ tiêu STT %d (%s) chưa có Min hoặc Max.\n"
                    "Hãy double-click vào ô để nhập." % (ct["stt"], ct["ten"]))
                return

        bo_cau_hinh.nho_lan_truoc("kich_thuoc", self.lay_cau_hinh())
        self.btn_chay.config(state="disabled")
        self.console.clear_log()
        self.console.ghi("Bắt đầu sinh dữ liệu kích thước...\n", tag="accent")
        self.console.ghi("Sản phẩm: %s / %s\n" % (self.v_loai.get(), self.v_ma.get()))
        self.console.ghi("Số chỉ tiêu: %d\n" % len(self.danh_sach_chi_tieu_hien_tai))
        self.console.ghi("-" * 60 + "\n")

        seed = None
        if self.v_seed.get().strip():
            try:
                seed = int(self.v_seed.get().strip())
            except ValueError:
                pass

        try:
            cpk = ktcore.spc_luat.kiem_tra_nguong(None, None, str(self.v_cpk.get()).strip().replace(",", "."))[2]
        except ValueError as e:
            self.btn_chay.config(state="normal")
            messagebox.showwarning("Sai Cpk tối thiểu", str(e))
            return

        t = threading.Thread(target=self._sinh_trong_nen, args=(seed, cpk),
                              daemon=True)
        t.start()
        self._poll()

    def _sinh_trong_nen(self, seed, cpk=None):
        """Chay trong thread nen."""
        try:
            def cb(stt, ten, dat, lan, msg):
                icon = "✅" if dat else "❌"
                self.hang_doi.put(("log", "%s STT %02d: %s — %s"
                                   % (icon, stt, ten[:40], msg), "success" if dat else "error"))

            ket_qua = ktcore.sinh_toan_bo(
                self.danh_sach_chi_tieu_hien_tai,
                so_mau=25,
                callback=cb,
                seed=seed,
                cpk_toi_thieu=cpk if cpk is not None else ktcore.CPK_TOI_THIEU)

            # Xuat file Excel
            self.hang_doi.put(("log", "-" * 60, "normal"))
            self.hang_doi.put(("log", "Đang xuất file Excel...", "accent"))

            thu_muc = xkt.tao_thu_muc_xuat()
            file_mau = self.v_file_mau.get()
            ten_file = "BaoCao_SPC_KichThuoc_%s_%s.xlsx" % (
                self.v_loai.get().replace(" ", "_"),
                self.v_ma.get().replace(" ", "_"))
            duong_dan_xuat = os.path.join(thu_muc, ten_file)

            thong_tin = {
                "san_pham": self.v_loai.get(),
                "kich_thuoc": self.v_ma.get(),
                "vat_lieu": self.v_vat_lieu.get(),
            }

            xkt.xuat_bao_cao(ket_qua, thong_tin, file_mau, duong_dan_xuat)
            self.file_cuoi = duong_dan_xuat

            so_dat = sum(1 for r in ket_qua if r["spc_result"].dat)
            self.hang_doi.put(("log", "=" * 60, "normal"))
            self.hang_doi.put(("log",
                "HOÀN TẤT: %d/%d chỉ tiêu đạt SPC. File: %s"
                % (so_dat, len(ket_qua), ten_file), "success"))

            self.hang_doi.put(("xong", thu_muc, so_dat, len(ket_qua)))

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
                elif msg[0] == "xong":
                    self.thu_muc_ket_qua = msg[1]
                    self.btn_chay.config(state="normal")
                    self.btn_mo.config(state="normal")
                    self.lbl_tien_trinh.config(
                        text="✅ Hoàn tất! File đã lưu tại: %s" % msg[1])
                    if self.khi_xong and self.file_cuoi:
                        try:
                            self.khi_xong(self.file_cuoi)
                            self.console.ghi("➜ Đã nạp sẵn file vừa sinh sang "
                                             "bước ② Đánh giá SPC.\n", tag="accent")
                        except Exception:
                            self.console.ghi("(Không nạp được sang Đánh giá SPC:\n%s)\n"
                                             % traceback.format_exc(), tag="error")
                    dat, tong = (msg[2], msg[3]) if len(msg) > 3 else (None, None)
                    ket = ("" if tong is None else
                           "%d/%d chỉ tiêu đạt SPC.%s\n" % (dat, tong, "" if dat == tong else
                                                            " Chỉ tiêu chưa đạt xem ở nhật ký."))
                    if messagebox.askyesno("Xong", "Đã sinh xong dữ liệu và báo cáo kích thước.\n"
                                                   "%s\nMở thư mục kết quả bây giờ?" % ket):
                        self._mo_ket_qua()
                    return
                elif msg[0] == "loi":
                    self.console.ghi("LỖI:\n" + msg[1] + "\n", tag="error")
                    self.btn_chay.config(state="normal")
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
        """Tra ve dict cau hinh de luu - ca Min/Max da sua trong bang."""
        return {
            "file_dung_sai": self.v_file_ds.get(),
            "file_mau": self.v_file_mau.get(),
            "loai": self.v_loai.get(),
            "ma": self.v_ma.get(),
            "vat_lieu": self.v_vat_lieu.get(),
            "cpk": self.v_cpk.get(),
            "seed": self.v_seed.get(),
            "min_max": {str(ct["stt"]): [ct["min"], ct["max"]]
                        for ct in self.danh_sach_chi_tieu_hien_tai},
        }

    def nap_cau_hinh(self, cfg):
        """Nap cau hinh da luu. File da luu khong con thi giu file dang co
        (thuong la file tu tim thay trong file_mau/)."""
        if not cfg:
            return
        if cfg.get("file_dung_sai") and os.path.exists(cfg["file_dung_sai"]) \
                and cfg["file_dung_sai"] != self.v_file_ds.get():
            self.v_file_ds.set(cfg["file_dung_sai"])
            self._doc_file_dung_sai(cfg["file_dung_sai"])
        if cfg.get("file_mau") and os.path.exists(cfg["file_mau"]):
            self.v_file_mau.set(cfg["file_mau"])
        if cfg.get("loai"):
            self.v_loai.set(cfg["loai"])
            self._khi_doi_loai(None)
        if cfg.get("ma"):
            self.v_ma.set(cfg["ma"])
            self._khi_doi_ma(None)
        if cfg.get("vat_lieu"):
            self.v_vat_lieu.set(cfg["vat_lieu"])
        if cfg.get("cpk"):
            self.v_cpk.set(cfg["cpk"])
        if cfg.get("seed"):
            self.v_seed.set(cfg["seed"])
        self._ap_min_max(cfg.get("min_max") or {})

    def _ap_min_max(self, mm):
        """Dat lai Min/Max da sua (theo STT) vao bang va danh sach."""
        if not mm:
            return
        for ct in self.danh_sach_chi_tieu_hien_tai:
            g = mm.get(str(ct["stt"]))
            if g and len(g) == 2:
                ct["min"], ct["max"] = g[0], g[1]
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            g = mm.get(str(vals[0]))
            if g and len(g) == 2:
                vals[6] = "-" if g[0] is None else "%.1f" % g[0]
                vals[7] = "-" if g[1] is None else "%.1f" % g[1]
                self.tree.item(item, values=vals)
