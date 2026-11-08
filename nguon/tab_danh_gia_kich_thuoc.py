# -*- coding: utf-8 -*-
"""
TAB KICH THUOC - BUOC 2: DANH GIA SPC & TU DONG SUA VAO BAO CAO TONG HOP
========================================================================
Nap thu muc KichThuoc_... (Buoc 1 tao xong tu nap) hoac file Bao cao tong hop
('BaoCao_TongHop_<SP>_<Ma>.xlsx', ho tro N mau bat ky vd 25, 100, 1000 mau)
hoac file bao cao SPC le JIS.QT.07.BM.01:
  - Mac dinh chia day 25 gia tri lien tuc buoc nhay 1 (vd 1000 mau -> 40 day:
    1-25, 26-50, ..., 976-1000 -> xuat ra 40 file bao cao danh gia on dinh).
  - Kiem tra lai toan bo tieu chi SPC (8 luat I-chart, 2 luat MR-chart, Cpk,
    dung sai tung diem do) tren tat ca cac day va chi tieu.
  - Neu phat hien du lieu chua dat o Buoc 2: tu dong sua cho dat chuan SPC
    (uu tien hoan vi thu tu mau; neu vuong gia tri SPEC/I1/Cpk thi tinh chinh
    gia tri nam trong [Min, Max] an toan), sao luu ban goc vao _truoc_khi_sua/
    va GHI DE DU LIEU DA SUA VAO FILE BAO CAO TONG HOP SINH O BUOC 1 truoc khi
    xuat cac file bao cao danh gia SPC JIS.QT.07.BM.01.
"""

import copy
import os
import re
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import bo_cau_hinh
import spc_luat
import kich_thuoc_core as ktcore
import xuat_kich_thuoc as xkt

KHOA = "danh_gia_kich_thuoc"


class TabDanhGiaKichThuoc(object):
    COT = ("day", "stt", "ten", "ky_hieu", "yeu_cau", "xb", "sigma", "lcl",
           "ucl", "cpk", "ket_luan")
    TEN_COT = ("Dãy mẫu", "STT", "Chỉ tiêu", "Ký hiệu", "Yêu cầu", "X̄", "σ",
               "LCLI", "UCLI", "Cpk", "Kết luận")
    RONG = (95, 42, 210, 60, 135, 70, 52, 65, 65, 55, 250)

    def __init__(self, me, mo_thu_muc_fn, tab_sinh=None):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.tab_sinh = tab_sinh
        self.hang_doi = queue.Queue()
        self.du_lieu = None
        self.du_lieu_master = None
        self.thu_muc_cuoi = None
        self.dang_dung = False
        self.kq_theo_day = []     # [(idx_day, bd, vt, nhan_ngan, nhan_day_full, ds_sub)]
        self.map_dong = {}        # iid -> (vt, nhan_full, sub_ct)

        self._dung_giao_dien()
        self._nap_master_dung_sai()
        if self.tab_sinh is not None:
            self.tab_sinh.khi_luu_dung_sai = self._nap_master_dung_sai
        self._nap_lan_truoc()

    def _chay_nen(self, fn, *a):
        threading.Thread(target=fn, args=a, daemon=True).start()

    # =========================================================== GIAO DIEN
    def _dung_giao_dien(self):
        ngoai = ttk.Frame(self.me)
        ngoai.pack(fill="both", expand=True)
        canvas = tk.Canvas(ngoai, bg=ui.COLOR_BG, highlightthickness=0)
        thanh = ttk.Scrollbar(ngoai, orient="vertical", command=canvas.yview)
        than = ttk.Frame(canvas)
        than.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cua = canvas.create_window((0, 0), window=than, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(cua, width=e.width))
        canvas.configure(yscrollcommand=thanh.set)
        canvas.pack(side="left", fill="both", expand=True)
        thanh.pack(side="right", fill="y")
        ui.cuon_bang_chuot(canvas)

        self._khoi_nguon(than)
        self._khoi_day_va_bang(than)
        self._khoi_bieu_mau(than)
        self._khoi_chay(than)

    # ------------------------------------------------------- 1. Nguon du lieu
    def _khoi_nguon(self, me):
        c = ui.ModernCard(me, title="Thư mục hoặc file Báo cáo tổng hợp cần đánh giá", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)

        self.v_nguon = tk.StringVar(value="")
        self.v_loai = tk.StringVar(value="")
        self.v_ma = tk.StringVar(value="")
        self.v_file_jis = tk.StringVar(value="")

        ttk.Label(k, text="Thư mục / file", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(k, textvariable=self.v_nguon).grid(
            row=0, column=1, sticky="ew", ipady=2)
        ttk.Button(k, text="📂 Chọn thư mục...", style="Outline.TButton", width=16,
                   command=self.chon_thu_muc).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(k, text="📄 Chọn file .xlsx", style="Outline.TButton", width=16,
                   command=self.chon_file).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(k, text="📥 Nạp", style="Outline.TButton", width=8,
                   command=self.nap_lai).grid(row=0, column=4, padx=(6, 0))

        kh_sp = tk.Frame(k, bg=ui.COLOR_SURFACE)
        kh_sp.grid(row=1, column=0, columnspan=5, sticky="ew", pady=(8, 0))
        ttk.Label(kh_sp, text="Loại sản phẩm", style="Surface.TLabel").pack(side="left", padx=(0, 8))
        self.cb_loai = ttk.Combobox(kh_sp, textvariable=self.v_loai, state="readonly", width=20)
        self.cb_loai.pack(side="left", padx=(0, 18), ipady=2)
        self.cb_loai.bind("<<ComboboxSelected>>", self._khi_doi_loai)

        ttk.Label(kh_sp, text="Mã sản phẩm", style="Surface.TLabel").pack(side="left", padx=(0, 8))
        self.cb_ma = ttk.Combobox(kh_sp, textvariable=self.v_ma, state="readonly", width=18)
        self.cb_ma.pack(side="left", ipady=2)
        self.cb_ma.bind("<<ComboboxSelected>>", self._khi_doi_ma)

        ttk.Label(k, text="File mẫu SPC JIS", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Entry(k, textvariable=self.v_file_jis).grid(
            row=2, column=1, columnspan=3, sticky="ew", ipady=2, pady=(8, 0))
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=self._chon_file_jis).grid(
            row=2, column=4, sticky="w", padx=(6, 0), pady=(8, 0))

        self.lbl = ttk.Label(
            k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
            text="Tạo dữ liệu xong ở bước ① thì thư mục Báo cáo tổng hợp tự nạp vào đây. "
                 "Cũng có thể chọn bất kỳ thư mục hoặc file Báo cáo tổng hợp (.xlsx) nào để kiểm tra & xuất báo cáo SPC."
        )
        self.lbl.grid(row=3, column=0, columnspan=5, sticky="w", pady=(6, 0))

    # ------------------------------------------------------- 2. Day mau & Bang
    def _khoi_day_va_bang(self, me):
        c = ui.ModernCard(me, title="Chọn dãy mẫu & Kết quả đánh giá SPC từng chỉ tiêu", badge_num=2)
        c.pack(fill="both", expand=True, padx=12, pady=4)
        k = c.get_body()

        o = tk.Frame(k, bg=ui.COLOR_SURFACE)
        o.pack(fill="x")
        self.v_bd = tk.StringVar(value="1")
        self.v_buoc = tk.StringVar(value="1")
        self.v_sd = tk.StringVar(value="25")
        for i, (nhan, bien, w) in enumerate((
            ("Mẫu bắt đầu", self.v_bd, 20),
            ("Bước nhảy", self.v_buoc, 6),
            ("Số điểm (≤ 25)", self.v_sd, 6),
        )):
            ttk.Label(o, text=nhan, style="Surface.TLabel").grid(
                row=0, column=i * 2, sticky="w", padx=(0 if i == 0 else 16, 6))
            ent = ttk.Entry(o, textvariable=bien, width=w)
            ent.grid(row=0, column=i * 2 + 1, sticky="w", ipady=2)
            ent.bind("<Return>", lambda e: self.cham(hoi=True))

        ttk.Label(
            k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
            text="Mặc định Mẫu bắt đầu = 1, Bước nhảy = 1, Số điểm = 25 → tự động chia hết thành các dãy 25 mẫu nối tiếp "
                 "(ví dụ 1000 mẫu → 40 dãy: 1–25, 26–50, …, 976–1000 → xuất 40 file báo cáo ổn định). "
                 "Cũng có thể nhập nhiều điểm bắt đầu cách nhau dấu phẩy (vd: 1, 26, 51)."
        ).pack(anchor="w", pady=(4, 6))

        thanh_nut = tk.Frame(k, bg=ui.COLOR_SURFACE)
        thanh_nut.pack(fill="x", pady=(0, 6))
        ttk.Button(thanh_nut, text="🔍 Chấm lại", style="Outline.TButton",
                   command=self.cham).pack(side="left")
        ttk.Button(thanh_nut, text="⚡ Quét mọi điểm bắt đầu", style="Outline.TButton",
                   command=self.quet).pack(side="left", padx=8)
        self.btn_sua = ttk.Button(thanh_nut, text="🛠 Sửa dữ liệu cho đạt", style="Outline.TButton",
                                  command=self.sua_du_lieu)
        self.btn_sua.pack(side="left")

        # Bo loc hien thi theo day hoac chi hien dong chua dat
        kh_loc = tk.Frame(thanh_nut, bg=ui.COLOR_SURFACE)
        kh_loc.pack(side="right")
        self.v_chi_loi = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            kh_loc, text="Chỉ hiện chưa đạt", variable=self.v_chi_loi,
            command=self._ve_lai_bang
        ).pack(side="right", padx=(10, 0))
        self.v_loc_day = tk.StringVar(value="Tất cả các dãy")
        self.cb_loc_day = ttk.Combobox(
            kh_loc, textvariable=self.v_loc_day, state="readonly", width=26
        )
        self.cb_loc_day.pack(side="right", ipady=1)
        self.cb_loc_day.bind("<<ComboboxSelected>>", lambda e: self._ve_lai_bang())
        ttk.Label(kh_loc, text="Xem:", style="Surface.TLabel").pack(side="right", padx=(0, 6))

        khung = tk.Frame(k, bg=ui.COLOR_SURFACE)
        khung.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(khung, columns=self.COT, show="headings",
                                 height=12, selectmode="browse")
        for col, ten, w in zip(self.COT, self.TEN_COT, self.RONG):
            self.tree.heading(col, text=ten)
            self.tree.column(col, width=w, minwidth=w,
                             anchor="w" if col in ("ten", "yeu_cau", "ket_luan") else "center")
        sb = ttk.Scrollbar(khung, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("hong", foreground=ui.COLOR_DANGER)
        self.tree.bind("<<TreeviewSelect>>", self._khi_chon)

        ttk.Label(
            k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
            text="Chấm đúng công thức biểu mẫu JIS.QT.07.BM.01: σ = ROUND(MR̄/1,128; 2), giới hạn làm tròn 1 chữ số, "
                 "8 quy tắc I-chart + 2 quy tắc MR-chart + Cpk + nghiệm thu từng điểm đo. "
                 "Khi sửa dữ liệu chưa đạt: phần mềm sao lưu bản gốc vào _truoc_khi_sua/ và cập nhật trực tiếp vào file Báo cáo tổng hợp."
        ).pack(anchor="w", pady=(4, 0))

    # ------------------------------------------------------- 3. Thong tin bieu mau
    def _khoi_bieu_mau(self, me):
        c = ui.ModernCard(me, title="Thông tin in trên biểu mẫu SPC (JIS.QT.07.BM.01)", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        for i in (1, 3):
            k.columnconfigure(i, weight=1)

        self.v = {}
        doi = [
            ("ngay_danh_gia", "Ngày đánh giá", datetime.now().strftime("%d/%m/%Y")),
            ("du_lieu_danh_gia", "Dữ liệu đánh giá (trống = tự ghi STT)", ""),
            ("vat_lieu", "Vật liệu", "SCW480"),
            ("tieu_chuan_ap_dung", "Tiêu chuẩn áp dụng", "JIS G 5102"),
            ("nguoi_lap", "Người lập", ""),
            ("nguoi_kiem_tra", "Người kiểm tra", "Đoàn Ngọc Thoạt"),
            ("nguoi_phe_duyet", "Phê duyệt", "Phạm Văn Phong"),
            ("cpk", "Cpk tối thiểu", "1.33"),
        ]
        for idx, (khoa, nhan, mac_dinh) in enumerate(doi):
            r, col = divmod(idx, 2)
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(
                row=r, column=col * 2, sticky="w",
                padx=(0 if col == 0 else 18, 8), pady=3)
            bien = tk.StringVar(value=mac_dinh)
            self.v[khoa] = bien
            ent = ttk.Entry(k, textvariable=bien)
            ent.grid(row=r, column=col * 2 + 1, sticky="ew", ipady=2, pady=3)
            if khoa == "cpk":
                ent.bind("<Return>", lambda e: self.cham(hoi=True))

    # ------------------------------------------------------- 4. Chay & Nhat ky
    def _khoi_chay(self, me):
        d = tk.Frame(me, bg=ui.COLOR_BG, padx=4, pady=6)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   TẠO BÁO CÁO SPC",
                              style="Primary.TButton", command=self.tao_bao_cao)
        self.btn.pack(side="left")
        self.btn_dung = ttk.Button(d, text="■ Dừng", style="Outline.TButton",
                                   command=self._dung, state="disabled")
        self.btn_dung.pack(side="left", padx=(8, 0))
        ttk.Button(d, text="📂 Mở thư mục", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=(10, 0))
        self.lbl_td = ttk.Label(d, text="", style="Phu.TLabel")
        self.lbl_td.pack(side="left", padx=12)

        self.log = ui.ModernTerminal(me, title="CHI TIẾT ĐÁNH GIÁ & XUẤT BÁO CÁO SPC", height=9)
        self.log.pack(fill="both", expand=True, padx=12, pady=(2, 12))

    # ======================================================= NAP MASTER & NGUON
    def _file_dung_sai(self):
        if self.tab_sinh is not None:
            p = self.tab_sinh.v_file_ds.get().strip()
            if p and os.path.isfile(p):
                return p
        return xkt._tim_file_dung_sai_mac_dinh()

    def _nap_master_dung_sai(self):
        p = self._file_dung_sai()
        if p and os.path.isfile(p):
            try:
                self.du_lieu_master = ktcore.doc_dung_sai(p)
                loai_list = list(self.du_lieu_master.keys())
                self.cb_loai["values"] = loai_list
            except Exception:
                self.du_lieu_master = None

    def _cap_nhat_combo_sp_ma(self, sp, ma):
        if not self.du_lieu_master:
            self._nap_master_dung_sai()
        if self.du_lieu_master:
            loai_list = list(self.du_lieu_master.keys())
            self.cb_loai["values"] = loai_list
            sp_chon = next((k for k in loai_list if k.lower() == str(sp or "").strip().lower()), sp)
            self.v_loai.set(sp_chon)
            ma_list = self.du_lieu_master.get(sp_chon, {}).get("ma_san_pham", [])
            self.cb_ma["values"] = ma_list
            if ma and ma in ma_list:
                self.v_ma.set(ma)
            elif ma_list:
                self.v_ma.set(ma_list[0])
            else:
                self.v_ma.set(ma or "")
        else:
            self.v_loai.set(sp or "")
            self.v_ma.set(ma or "")

        # Tu dong chon file mau JIS theo san pham
        mau_jis = xkt.tim_mau_jis(self.v_loai.get())
        if mau_jis and os.path.isfile(mau_jis):
            self.v_file_jis.set(mau_jis)

    def _dong_bo_min_max_tu_tab_sinh(self):
        """Neu cung Loai & Ma voi Buoc 1, ke thua toan bo thong so (Yeu cau, LSL, USL, Min, Max)
        nguoi dung vua chinh truc tiep tren bang Buoc 1."""
        if not self.du_lieu or self.tab_sinh is None:
            return
        sp = self.du_lieu.get("san_pham")
        ma = self.du_lieu.get("ma")
        if hasattr(self.tab_sinh, "thong_so_hien_tai"):
            ts = self.tab_sinh.thong_so_hien_tai(sp, ma)
            if ts:
                for ct in self.du_lieu.get("chi_tieu", []):
                    info = ts.get(ct["stt"])
                    if not info:
                        continue
                    for k in ("yeu_cau", "chuan", "ds_duoi", "ds_tren", "lsl", "usl"):
                        if k in info and info[k] is not None or k in ("lsl", "usl", "ds_duoi", "ds_tren"):
                            ct[k] = info.get(k)
                    if info.get("min") is not None:
                        ct["min"] = info["min"]
                    if info.get("max") is not None:
                        ct["max"] = info["max"]
                return
        mm = self.tab_sinh.min_max_hien_tai(sp, ma)
        if not mm:
            return
        for ct in self.du_lieu.get("chi_tieu", []):
            cap = mm.get(ct["stt"])
            if cap and cap[0] is not None and cap[1] is not None:
                ct["min"], ct["max"] = cap[0], cap[1]

    def _khi_doi_loai(self, _e=None):
        loai = self.v_loai.get()
        if self.du_lieu_master and loai in self.du_lieu_master:
            ma_list = self.du_lieu_master[loai].get("ma_san_pham", [])
            self.cb_ma["values"] = ma_list
            if ma_list:
                self.v_ma.set(ma_list[0])
        mau_jis = xkt.tim_mau_jis(loai)
        if mau_jis and os.path.isfile(mau_jis):
            self.v_file_jis.set(mau_jis)
        self._khi_doi_ma()

    def _khi_doi_ma(self, _e=None):
        if not self.du_lieu:
            return
        xkt.ap_dung_dung_sai(
            self.du_lieu,
            san_pham=self.v_loai.get(),
            ma=self.v_ma.get(),
            duong_dan_ds=self._file_dung_sai(),
        )
        self._dong_bo_min_max_tu_tab_sinh()
        self.cham(hoi=False)

    def _chon_file_jis(self):
        p = filedialog.askopenfilename(
            title="Chọn file mẫu báo cáo SPC (JIS.QT.07.BM.01)",
            filetypes=[("File Excel", "*.xlsx"), ("Tất cả", "*.*")])
        if p:
            self.v_file_jis.set(p)

    def chon_thu_muc(self):
        p = filedialog.askdirectory(title="Chọn thư mục KichThuoc_... chứa Báo cáo tổng hợp")
        if p:
            self.dat_san(p)

    def chon_file(self):
        p = filedialog.askopenfilename(
            title="Chọn file Báo cáo tổng hợp hoặc file báo cáo SPC kích thước (.xlsx)",
            filetypes=[("File Excel", "*.xlsx"), ("Tất cả", "*.*")])
        if p:
            self.dat_san(p)

    def nap_lai(self):
        p = self.v_nguon.get().strip()
        if not p:
            messagebox.showinfo("Chưa chọn nguồn", "Hãy chọn thư mục hoặc file .xlsx trước.")
            return
        self.dat_san(p, giu_day=True)

    def dat_san(self, p, giu_day=False):
        """Nap thu muc hoac file bao cao va cham toan bo cac day mau (Buoc 1 goi tu dong sau khi sinh)."""
        p = str(p or "").strip()
        if not p:
            return
        self.v_nguon.set(p)
        try:
            d = xkt.doc_bao_cao(p, duong_dan_dung_sai=self._file_dung_sai())
        except Exception as e:
            self.du_lieu = None
            self.kq_theo_day = []
            self.map_dong = {}
            self.tree.delete(*self.tree.get_children())
            self.lbl.configure(text="❌ Không đọc được dữ liệu: %s" % e)
            self.log.ghi("KHÔNG ĐỌC ĐƯỢC:\n%s\n" % e, xoa=True, tag="error")
            return

        self.du_lieu = d
        self.thu_muc_cuoi = d.get("thu_muc") or os.path.dirname(d["file"])
        if d.get("vat_lieu") and not self.v["vat_lieu"].get().strip():
            self.v["vat_lieu"].set(d["vat_lieu"])

        self._cap_nhat_combo_sp_ma(d.get("san_pham"), d.get("ma"))
        xkt.ap_dung_dung_sai(d, d.get("san_pham"), self.v_ma.get(), self._file_dung_sai())
        self._dong_bo_min_max_tu_tab_sinh()

        if not giu_day:
            self.v_bd.set("1")
            self.v_buoc.set("1")
            self.v_sd.set(str(min(ktcore.SO_MAU, max(2, d.get("so_mau") or ktcore.SO_MAU))))
            self.v_loc_day.set("Tất cả các dãy")

        self.cham(hoi=False)

    # ======================================================= THAM SO & CHAM SPC
    def _tach_so_bat_dau(self, chuoi):
        ra = []
        for phan in re.split(r"[,\s;]+", str(chuoi or "").strip()):
            if not phan:
                continue
            ra.append(spc_luat.so_nguyen(phan, "Mẫu bắt đầu", 1))
        if not ra:
            raise ValueError("Hãy nhập ít nhất một số nguyên dương vào ô 'Mẫu bắt đầu' (mặc định là 1).")
        return ra

    def _tham_so(self):
        if not self.du_lieu:
            raise ValueError("Chưa nạp dữ liệu Báo cáo tổng hợp.")
        d = self.du_lieu
        so_mau = d.get("so_mau") or len(d["chi_tieu"][0].get("du_lieu") or [])
        if so_mau < 2:
            raise ValueError("Dữ liệu chỉ có %d mẫu — biểu đồ I-MR cần ít nhất 2 mẫu." % so_mau)

        cpk = spc_luat.kiem_tra_nguong(
            None, None, str(self.v["cpk"].get()).strip().replace(",", ".")
        )[2]
        buoc = spc_luat.so_nguyen(self.v_buoc.get().strip(), "Bước nhảy", 1)
        sd_nhap = spc_luat.so_nguyen(self.v_sd.get().strip(), "Số điểm", 2)
        if sd_nhap > ktcore.SO_MAU:
            raise ValueError("Số điểm mỗi biểu đồ tối đa là %d." % ktcore.SO_MAU)
        sd = min(sd_nhap, so_mau)

        bd_list = self._tach_so_bat_dau(self.v_bd.get())
        # Quy uoc thong nhat toan phan mem: nhap 1 diem bat dau + buoc nhay 1 -> tu dong
        # chia het thanh cac day sd mau lien tiep (vd 1000 mau -> 40 day: 1..25, 26..50, ..., 976..1000)
        if len(bd_list) == 1 and buoc == 1:
            s0 = bd_list[0]
            if s0 > so_mau - 1:
                raise ValueError("Mẫu bắt đầu (%d) vượt quá tổng số mẫu (%d)." % (s0, so_mau))
            bd_list = [s for s in range(s0, so_mau + 1, sd) if so_mau - s + 1 >= 2]

        cac_day = []
        da_co = set()
        for s in bd_list:
            vt = ktcore.lay_day(so_mau, s, buoc, sd)
            t_vt = tuple(vt)
            if t_vt not in da_co:
                da_co.add(t_vt)
                cac_day.append((vt[0], vt))

        thong_tin_in = {k: bien.get().strip() for k, bien in self.v.items() if k != "cpk"}
        file_jis = self.v_file_jis.get().strip() or xkt.tim_mau_jis(d.get("san_pham"), d.get("thu_muc"))

        return {
            "d": d,
            "so_mau": so_mau,
            "cpk": cpk,
            "buoc": buoc,
            "so_diem": sd,
            "cac_day": cac_day,
            "thong_tin_in": thong_tin_in,
            "file_jis": file_jis,
        }

    def cham(self, hoi=True):
        """Cham lai toan bo cac day va cap nhat bo loc + bang ket qua."""
        self.kq_theo_day = []
        self.map_dong = {}
        self.tree.delete(*self.tree.get_children())

        if not self.du_lieu:
            if hoi:
                messagebox.showinfo("Chưa có dữ liệu",
                                    "Hãy chọn thư mục hoặc file Báo cáo tổng hợp rồi bấm Nạp trước.")
            return

        try:
            ts = self._tham_so()
        except Exception as e:
            self.log.ghi("CHƯA CHẤM ĐƯỢC:\n%s\n" % e, xoa=True, tag="error")
            if hoi:
                messagebox.showwarning("Tham số chưa hợp lệ", str(e))
            return

        bo_cau_hinh.nho_lan_truoc(KHOA, self.lay_cau_hinh())
        d = ts["d"]
        stt_mau = d.get("stt_mau") or list(range(1, ts["so_mau"] + 1))
        so_day_dat = 0
        tong_ct_dat = 0
        tong_ct = 0

        lua_chon_loc = ["Tất cả các dãy (%d)" % len(ts["cac_day"])]
        w_pad = max(2, len(str(len(ts["cac_day"]))))

        for idx_day, (bd, vt) in enumerate(ts["cac_day"], start=1):
            nhan_ngan = xkt.nhan_day(stt_mau, vt, ngan=True)
            nhan_full = xkt.nhan_day(stt_mau, vt, ngan=False)
            ds_sub = ktcore.cham_day_kich_thuoc(d["chi_tieu"], vt, ts["cpk"])
            ct_dat = sum(1 for c in ds_sub if c["kq"].dat)
            tong_ct_dat += ct_dat
            tong_ct += len(ds_sub)
            dat_ca_day = (ct_dat == len(ds_sub))
            if dat_ca_day:
                so_day_dat += 1
            nhan_combo = "Dãy %0*d: %s (%s)" % (
                w_pad, idx_day, nhan_full, "ĐẠT" if dat_ca_day else ("LỖI %d" % (len(ds_sub) - ct_dat))
            )
            lua_chon_loc.append(nhan_combo)
            self.kq_theo_day.append((idx_day, bd, vt, nhan_ngan, nhan_full, nhan_combo, ds_sub))

        self.cb_loc_day["values"] = lua_chon_loc
        cur_loc = self.v_loc_day.get()
        if cur_loc not in lua_chon_loc:
            self.v_loc_day.set(lua_chon_loc[0])

        self._ve_lai_bang()

        loai_str = "Báo cáo tổng hợp" if d.get("loai_file") == "tong_hop" else "Báo cáo SPC JIS"
        n_day = len(ts["cac_day"])
        self.lbl.configure(
            text="[%s: %s] %s / %s · Tổng %d mẫu · %d chỉ tiêu · %d dãy đánh giá — %d/%d dãy ĐẠT chuẩn SPC (%d/%d lượt chỉ tiêu)."
            % (loai_str, os.path.basename(d["file"]), d["san_pham"], d["ma"],
               ts["so_mau"], len(d["chi_tieu"]), n_day, so_day_dat, n_day, tong_ct_dat, tong_ct)
        )

        if so_day_dat == n_day:
            self.log.ghi(
                "✅ ĐÃ KIỂM TRA %d DÃY (%d MẪU × %d CHỈ TIÊU = %d LƯỢT ĐÁNH GIÁ): TẤT CẢ 100%% ĐẠT CHUẨN SPC!\n"
                "Bấm “▶ TẠO BÁO CÁO SPC” để xuất %d file báo cáo đánh giá ổn định (JIS.QT.07.BM.01).\n"
                % (n_day, ts["so_mau"], len(d["chi_tieu"]), tong_ct, n_day),
                xoa=True, tag="success",
            )
        else:
            cac_day_hong = [
                "%s (%s)" % (
                    nhan_full,
                    ", ".join("%s:%s" % (c.get("ky_hieu") or c["stt"], "+".join(c["kq"].loi))
                              for c in ds_sub if not c["kq"].dat)[:60]
                )
                for _i, _b, _vt, _nn, nhan_full, _nc, ds_sub in self.kq_theo_day
                if any(not c["kq"].dat for c in ds_sub)
            ]
            mo_ta = "\n".join("  • " + s for s in cac_day_hong[:8])
            if len(cac_day_hong) > 8:
                mo_ta += "\n  • … và %d dãy khác" % (len(cac_day_hong) - 8)
            self.log.ghi(
                "⚠ PHÁT HIỆN %d/%d DÃY CHƯA ĐẠT CHUẨN SPC:\n%s\n\n"
                "➜ Bấm “🛠 Sửa dữ liệu cho đạt” hoặc “▶ TẠO BÁO CÁO SPC”: phần mềm sẽ tự động sửa dữ liệu đạt chuẩn "
                "và cập nhật trực tiếp vào file Báo cáo tổng hợp (%s).\n"
                % (n_day - so_day_dat, n_day, mo_ta, os.path.basename(d["file"])),
                xoa=True, tag="warning",
            )

    def _ve_lai_bang(self):
        """Render Treeview tu ket qua da cham theo bo loc Day va Chi hien chua dat."""
        self.tree.delete(*self.tree.get_children())
        self.map_dong = {}
        if not self.kq_theo_day:
            return

        loc_day = self.v_loc_day.get()
        chi_loi = bool(self.v_chi_loi.get())
        xem_tat_ca = loc_day.startswith("Tất cả")

        for idx_day, bd, vt, nhan_ngan, nhan_full, nhan_combo, ds_sub in self.kq_theo_day:
            if not xem_tat_ca and nhan_combo != loc_day:
                continue
            for c_idx, ct in enumerate(ds_sub):
                r = ct["kq"]
                if chi_loi and r.dat:
                    continue
                iid = "%d|%d" % (idx_day, c_idx)
                self.map_dong[iid] = (vt, nhan_full, ct)
                kl = "ĐẠT" if r.dat else "CHƯA ĐẠT: " + ", ".join(r.loi)
                self.tree.insert(
                    "", "end", iid=iid, tags=() if r.dat else ("hong",),
                    values=(
                        nhan_ngan,
                        ct["stt"],
                        str(ct.get("ten") or "").replace("\n", " "),
                        ct.get("ky_hieu") or "",
                        str(ct.get("yeu_cau") or "").replace("\n", " "),
                        "-" if r.xb is None else "%.3f" % r.xb,
                        "-" if r.sigma is None else "%.2f" % r.sigma,
                        "-" if r.lcl_i is None else "%.1f" % r.lcl_i,
                        "-" if r.ucl_i is None else "%.1f" % r.ucl_i,
                        "-" if r.cpk is None else "%.2f" % r.cpk,
                        kl,
                    )
                )

    def _khi_chon(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        info = self.map_dong.get(sel[-1])
        if not info:
            return
        vt, nhan_full, ct = info
        r = ct["kq"]
        L = [
            "%s · STT %s · %s (%s) · Yêu cầu: %s" % (
                nhan_full, ct["stt"], str(ct.get("ten") or "").replace("\n", " "),
                ct.get("ky_hieu") or "", str(ct.get("yeu_cau") or "").replace("\n", " ")
            ),
            "  LSL: %s  |  USL: %s  |  Min/Max sinh chuẩn: %s .. %s" % (
                ct.get("lsl"), ct.get("usl"), ct.get("min"), ct.get("max")
            ),
        ]
        if "DU_LIEU" in r.loi:
            L.append("  CHƯA ĐẠT: " + "; ".join(ct.get("loi_du_lieu") or ["Thiếu hoặc sai điểm đo"]))
        else:
            mr3_txt = "Có (chỉ tham khảo)" if getattr(r, "mr3", False) else "Không"
            L.append("  X̄ = %.4f   MR̄ = %.4f   σ = %.2f   UCLI = %.1f / LCLI = %.1f   UCLMR = %.2f"
                     % (r.xb, r.mrb, r.sigma, r.ucl_i, r.lcl_i, r.ucl_mr))
            L.append("  Ca = %s   Cp = %s   Cpk = %s   |   MR3 (tham khảo): %s" % (
                "-" if r.ca is None else "%.3f" % r.ca,
                "-" if r.cp is None else "%.3f" % r.cp,
                "-" if r.cpk is None else "%.3f" % r.cpk,
                mr3_txt,
            ))
            if r.dat:
                L.append("  KẾT LUẬN: ĐẠT đủ tất cả các tiêu chí đánh giá SPC")
            else:
                L.append("  KẾT LUẬN: CHƯA ĐẠT — vi phạm:")
                for m in r.loi:
                    L.append("     [%-4s] %s" % (m, spc_luat.TEN_LUAT.get(m, m)))
        self.log.ghi("\n".join(L) + "\n", xoa=True)

    # ======================================================= QUET MOI DIEM
    def quet(self):
        try:
            ts = self._tham_so()
            q = ktcore.quet_kich_thuoc(
                ts["d"]["chi_tieu"], ts["so_mau"],
                buoc=ts["buoc"], so_diem=ts["so_diem"], cpk_toi_thieu=ts["cpk"]
            )
        except Exception as e:
            messagebox.showwarning("Không quét được", str(e))
            return
        dat = [s for s, ok, _l in q if ok]
        hong = [(s, l) for s, ok, l in q if not ok]
        L = [
            "Quét %d cửa sổ trượt · bước nhảy %d · %d mẫu/cửa sổ:" % (len(q), ts["buoc"], ts["so_diem"]),
            "  • ĐẠT tất cả chỉ tiêu (%d cửa sổ): %s%s" % (
                len(dat), ", ".join(map(str, dat[:40])), " …" if len(dat) > 40 else ""
            ),
            "  • CHƯA ĐẠT (%d cửa sổ): %s%s" % (
                len(hong),
                "; ".join("bd %d (%s)" % (s, ",".join(l[:4])) for s, l in hong[:15]),
                " …" if len(hong) > 15 else "",
            ),
        ]
        if q:
            L.append("  • Tỉ lệ cửa sổ đạt: %.1f%%" % (100.0 * len(dat) / len(q)))
        self.log.ghi("\n".join(L) + "\n", xoa=True)

    # ======================================================= SUA DU LIEU CHO DAT
    def _cac_day_chua_dat(self, ts):
        """Tra ve [(vt, nhan_full, [mo_ta_loi])] cho cac day dang chon chua dat SPC."""
        d = ts["d"]
        stt_mau = d.get("stt_mau") or list(range(1, ts["so_mau"] + 1))
        hong = []
        for _bd, vt in ts["cac_day"]:
            ds_sub = ktcore.cham_day_kich_thuoc(d["chi_tieu"], vt, ts["cpk"])
            loi = [
                "%s (%s)" % (c.get("ky_hieu") or ("STT %s" % c["stt"]), ",".join(c["kq"].loi))
                for c in ds_sub if not c["kq"].dat
            ]
            if loi:
                hong.append((vt, xkt.nhan_day(stt_mau, vt), loi))
        return hong

    def _mo_ta_day_hong(self, hong):
        lines = [
            "  • %s: %s%s" % (nhan, ", ".join(loi[:5]), "…" if len(loi) > 5 else "")
            for _vt, nhan, loi in hong[:8]
        ]
        if len(hong) > 8:
            lines.append("  • … và %d dãy khác" % (len(hong) - 8))
        return "\n".join(lines)

    def _dung(self):
        self.dang_dung = True
        self.btn_dung.configure(state="disabled")
        self.lbl_td.configure(text="Đang dừng...")

    def _bat_nut(self, bat, nut_dang_chay=None, chu=None):
        st = "normal" if bat else "disabled"
        self.btn.configure(state=st, text="▶   TẠO BÁO CÁO SPC")
        self.btn_sua.configure(state=st, text="🛠 Sửa dữ liệu cho đạt")
        self.btn_dung.configure(state="disabled" if bat else "normal")
        if not bat and nut_dang_chay is not None and chu:
            nut_dang_chay.configure(text=chu)
        if bat:
            self.lbl_td.configure(text="")

    def sua_du_lieu(self):
        try:
            ts = self._tham_so()
        except Exception as e:
            messagebox.showwarning("Tham số chưa hợp lệ", str(e))
            return

        hong = self._cac_day_chua_dat(ts)
        if not hong:
            messagebox.showinfo(
                "Đã đạt chuẩn SPC",
                "Tất cả %d/%d dãy đang chọn đều đã ĐẠT chuẩn SPC 100%%, không cần sửa."
                % (len(ts["cac_day"]), len(ts["cac_day"]))
            )
            return

        if not messagebox.askyesno(
            "Sửa dữ liệu cho đạt & cập nhật Báo cáo tổng hợp",
            "Phát hiện %d/%d dãy chưa đạt chuẩn SPC:\n\n%s\n\n"
            "Phần mềm sẽ:\n"
            "1. Ưu tiên hoán vị thứ tự mẫu trong dãy lỗi; nếu chỉ tiêu vướng giá trị ngoài dung sai / Cpk "
            "thì tinh chỉnh giá trị nằm trong [Min, Max] chuẩn.\n"
            "2. Sao lưu bản gốc vào _truoc_khi_sua/ và ghi đè dữ liệu đạt chuẩn trực tiếp vào file "
            "Báo cáo tổng hợp (%s).\n\nTiếp tục?"
            % (len(hong), len(ts["cac_day"]), self._mo_ta_day_hong(hong), os.path.basename(ts["d"]["file"]))
        ):
            return

        self._bat_dau_sua(ts, xuat_sau=False)

    def _bat_dau_sua(self, ts, xuat_sau=False):
        ts["xuat_sau"] = xuat_sau
        self.dang_dung = False
        self._bat_nut(False, self.btn_sua, "Đang sửa dữ liệu…")
        self.log.ghi(
            "Đang tự động chuẩn hoá dữ liệu và cập nhật vào Báo cáo tổng hợp (%s)...\n"
            % os.path.basename(ts["d"]["file"]),
            xoa=True, tag="accent",
        )
        self._chay_nen(self._sua_nen, ts)
        self.me.after(120, self._kiem_tra_hang_doi)

    def _sua_nen(self, ts):
        try:
            d_copy = copy.deepcopy(ts["d"])
            cac_vt = [vt for _bd, vt in ts["cac_day"]]
            kq_sua = ktcore.sua_du_lieu_nhieu_day(
                d_copy["chi_tieu"],
                cac_vt,
                cpk_toi_thieu=ts["cpk"],
                cho_phep_chinh_gia_tri=True,
                dung=lambda: self.dang_dung,
                tien_do=lambda cur, tot: self.hang_doi.put(("td_sua", (cur, tot))),
            )
            if not kq_sua["dat_het"]:
                raise ValueError(
                    "Sau khi sửa vẫn còn %d dãy chưa đạt:\n%s"
                    % (len(kq_sua["con_hong"]), "\n".join(kq_sua["con_hong"][:8]))
                )
            ra_ghi = xkt.ghi_lai_bao_cao_tong_hop(d_copy, luu_goc=True)
            self.hang_doi.put(("xong_sua", (d_copy, kq_sua, ra_ghi, ts.get("xuat_sau"))))
        except InterruptedError as e:
            self.hang_doi.put(("dung", str(e)))
        except PermissionError as e:
            self.hang_doi.put((
                "loi",
                "Không ghi được vào file Báo cáo tổng hợp — thường do file đang mở trong Excel.\n"
                "Hãy đóng Excel rồi thử lại.\n\n%s" % e,
            ))
        except ValueError as e:
            self.hang_doi.put(("loi", str(e)))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    # ======================================================= XUAT BAO CAO SPC
    def tao_bao_cao(self, da_sua=False):
        """Kiem tra lai toan bo cac day truoc khi xuat. Neu phat hien day chua dat ->
        de nghi sua truc tiep vao Bao cao tong hop cua Buoc 1 roi xuat toan bo file SPC."""
        try:
            ts = self._tham_so()
        except Exception as e:
            messagebox.showwarning("Tham số chưa hợp lệ", str(e))
            return

        if not ts["file_jis"] or not os.path.isfile(ts["file_jis"]):
            messagebox.showwarning(
                "Thiếu file mẫu SPC JIS",
                "Không tìm thấy file mẫu JIS.QT.07.BM.01 cho sản phẩm '%s'.\n"
                "Hãy chọn đường dẫn ở ô 'File mẫu SPC JIS'." % ts["d"].get("san_pham")
            )
            return

        hong = [] if da_sua else self._cac_day_chua_dat(ts)
        if hong:
            chon = messagebox.askyesnocancel(
                "Phát hiện dữ liệu chưa đạt chuẩn SPC",
                "Kiểm tra lại trước khi xuất: có %d/%d dãy CHƯA ĐẠT chuẩn SPC:\n\n%s\n\n"
                "Bạn có muốn tự động sửa dữ liệu đạt chuẩn (cập nhật trực tiếp vào file "
                "Báo cáo tổng hợp ở Bước ①, sao lưu bản gốc vào _truoc_khi_sua/) "
                "rồi xuất %d file báo cáo SPC không?\n\n"
                "• Yes: Sửa vào Báo cáo tổng hợp cho đạt chuẩn rồi xuất %d file SPC\n"
                "• No : Giữ nguyên dữ liệu hiện tại và xuất %d file SPC\n"
                "• Cancel: Huỷ thao tác"
                % (len(hong), len(ts["cac_day"]), self._mo_ta_day_hong(hong),
                   len(ts["cac_day"]), len(ts["cac_day"]), len(ts["cac_day"]))
            )
            if chon is None:
                return
            if chon:
                self._bat_dau_sua(ts, xuat_sau=True)
                return

        bo_cau_hinh.nho_lan_truoc(KHOA, self.lay_cau_hinh())
        self.dang_dung = False
        self._bat_nut(False, self.btn, "Đang xuất SPC…")
        self.log.ghi(
            "Đang xuất %d file báo cáo đánh giá ổn định (JIS.QT.07.BM.01)...\n" % len(ts["cac_day"]),
            xoa=not da_sua, tag="accent",
        )
        self._chay_nen(self._xuat_nen, ts)
        self.me.after(120, self._kiem_tra_hang_doi)

    def _xuat_nen(self, ts):
        try:
            cac_vt = [vt for _bd, vt in ts["cac_day"]]
            ra_files = xkt.xuat_nhieu_bao_cao_spc(
                ts["d"],
                cac_vt,
                file_mau_jis=ts["file_jis"],
                cpk_toi_thieu=ts["cpk"],
                thong_tin_in=ts["thong_tin_in"],
                thu_muc_xuat=ts["d"].get("thu_muc"),
                buoc=ts["buoc"],
                dung=lambda: self.dang_dung,
                tien_do=lambda idx, tot, fn: self.hang_doi.put(("td_xuat", (idx, tot, fn))),
            )
            self.hang_doi.put(("xong_xuat", (ra_files, ts)))
        except InterruptedError as e:
            self.hang_doi.put(("dung", str(e)))
        except xkt.LoiMau as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put((
                "loi",
                "Không ghi được file báo cáo SPC — thường do file đang mở trong Excel.\n"
                "Hãy đóng Excel rồi thử lại.\n\n%s" % e,
            ))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    # ======================================================= HANG DOI KET QUA
    def _kiem_tra_hang_doi(self):
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "td_sua":
                    cur, tot = dl
                    self.lbl_td.configure(text="Đang kiểm tra & sửa: %d / %d" % (cur, tot))
                    continue
                if loai == "td_xuat":
                    idx, tot, fn = dl
                    self.lbl_td.configure(text="Đã xuất %d / %d: %s" % (idx, tot, fn))
                    continue

                self._bat_nut(True)
                if loai == "dung":
                    self.log.ghi("⏹ %s\n" % dl, tag="warning")
                elif loai == "xong_sua":
                    self._hoan_tat_sua(*dl)
                elif loai == "xong_xuat":
                    self._hoan_tat_xuat(*dl)
                elif loai == "loi":
                    self.log.ghi("KHÔNG HOÀN TẤT:\n%s\n" % dl, xoa=True, tag="error")
                    messagebox.showerror("Không hoàn tất", dl)
                else:
                    self.log.ghi("LỖI HỆ THỐNG:\n%s\n" % dl, xoa=True, tag="error")
                    messagebox.showerror("Lỗi", "Có lỗi ngoài dự kiến. Xem chi tiết ở nhật ký bên dưới.")
                return
        except queue.Empty:
            pass
        self.me.after(120, self._kiem_tra_hang_doi)

    def _hoan_tat_sua(self, d_moi, kq_sua, ra_ghi, xuat_sau):
        # Nap lai truc tiep tu file Bao cao tong hop vua ghi de xac nhan 100% dong bo tren o dia
        p_nap = ra_ghi.get("tong_hop") or d_moi.get("file")
        self.dat_san(p_nap, giu_day=True)

        L = [
            "✅ ĐÃ SỬA DỮ LIỆU ĐẠT CHUẨN SPC VÀ GHI ĐÈ VÀO BÁO CÁO TỔNG HỢP:",
            "  • Số dãy đạt chuẩn: %d/%d dãy (đã sửa %d dãy: %d lượt hoán vị, %d lượt tinh chỉnh giá trị)"
            % (kq_sua["so_day_dat"], kq_sua["tong_day"], kq_sua["so_day_sua"],
               kq_sua["so_ct_hoan_vi"], kq_sua["so_ct_chinh_so"]),
            "  • Đã cập nhật file Bước ①: %s" % os.path.basename(ra_ghi["tong_hop"]),
        ]
        if ra_ghi.get("sao_luu"):
            L.append("  • Bản gốc trước khi sửa lưu tại: %s" % ra_ghi["sao_luu"])
        if ra_ghi.get("spc_cu"):
            L.append("  • Đã chuyển %d file SPC cũ sang thư mục sao lưu." % len(ra_ghi["spc_cu"]))
        self.log.ghi("\n".join(L) + "\n", tag="success")

        if xuat_sau and self.du_lieu:
            self.me.after(80, lambda: self.tao_bao_cao(da_sua=True))
            return

        messagebox.showinfo(
            "Đã sửa & cập nhật Báo cáo tổng hợp",
            "\n".join(L) + "\n\nBấm “▶ TẠO BÁO CÁO SPC” để xuất bộ file báo cáo đánh giá ổn định."
        )

    def _hoan_tat_xuat(self, ra_files, ts):
        self.thu_muc_cuoi = os.path.dirname(ra_files[0]) if ra_files else ts["d"].get("thu_muc")
        hong = self._cac_day_chua_dat(ts)
        so_dat = len(ra_files) - len(hong)

        L = [
            "🎉 ĐÃ XUẤT %d FILE BÁO CÁO ĐÁNH GIÁ ỔN ĐỊNH (JIS.QT.07.BM.01) — %d/%d file ĐẠT chuẩn SPC:"
            % (len(ra_files), so_dat, len(ra_files))
        ]
        for f in ra_files[:12]:
            L.append("  • " + os.path.basename(f))
        if len(ra_files) > 12:
            L.append("  • … (%d file ở giữa) …" % (len(ra_files) - 14))
            for f in ra_files[-2:]:
                L.append("  • " + os.path.basename(f))
        L.append("  ➜ Thư mục kết quả: %s" % self.thu_muc_cuoi)
        self.log.ghi("\n".join(L) + "\n", tag="warning" if hong else "success")

        tom_tat_ds = "\n".join("• " + os.path.basename(x) for x in ra_files[:8])
        if len(ra_files) > 8:
            tom_tat_ds += "\n• … và %d file khác (tổng cộng %d file)" % (len(ra_files) - 8, len(ra_files))

        if messagebox.askyesno(
            "Hoàn tất xuất báo cáo SPC",
            "Đã xuất %d file báo cáo đánh giá ổn định (%d/%d file ĐẠT chuẩn SPC):\n\n%s\n\n"
            "Mở thư mục kết quả bây giờ?"
            % (len(ra_files), so_dat, len(ra_files), tom_tat_ds)
        ):
            self.mo_thu_muc()

    def mo_thu_muc(self):
        tm = self.thu_muc_cuoi or (self.du_lieu or {}).get("thu_muc")
        if tm and os.path.isdir(tm) and self.mo_thu_muc_ngoai:
            self.mo_thu_muc_ngoai(tm)

    # ======================================================= LUU / NAP CAU HINH
    def lay_cau_hinh(self):
        cfg = {
            k: bien.get() for k, bien in self.v.items()
            if k not in ("ngay_danh_gia", "du_lieu_danh_gia")
        }
        cfg["buoc"] = self.v_buoc.get()
        cfg["so_diem"] = self.v_sd.get()
        return cfg

    def _nap_lan_truoc(self):
        cfg = bo_cau_hinh.doc_lan_truoc(KHOA) or {}
        for k, val in cfg.items():
            if k in self.v and isinstance(val, str) and val.strip():
                self.v[k].set(val)
        if cfg.get("buoc"):
            self.v_buoc.set(str(cfg["buoc"]))
        if cfg.get("so_diem"):
            self.v_sd.set(str(cfg["so_diem"]))
