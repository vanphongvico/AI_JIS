# -*- coding: utf-8 -*-
"""
THU KEO - BUOC 1: SINH BIEU DO KEO VA FILE SO LIEU
===================================================
Giao dien cho xuat_keo.py. Sinh ra cap file PDF + Excel lam DAU VAO cho
Buoc 2 (tab_keo.py), thay vi phai co san hai file do. Hai buoc duoc ghep
trong tab_thu_keo.py.
"""

import os
import sys
import json
import queue
import re
import threading
import traceback

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xuat_keo
import sinh_keo_core as keo_core
import bo_cau_hinh
import ui_theme as ui

BUOC_TRON = [1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 16, 20, 25, 30, 40]

# Ba muc bao dam. Moi muc la (buoc xet tu dau lo, buoc xet o moi diem bat dau).
# Con so trong nhan la do thuc te tren 1000 mau.
MUC = {
    "co_ban":   (BUOC_TRON, []),
    "lien_tiep": ([], [1]),
    "gop":      (BUOC_TRON, [1]),
}


class TabSinhKeo(object):
    """Dung toan bo giao dien Buoc 1 vao khung 'me' duoc truyen vao.

    khi_xong(kq, cfg) : goi sau moi lan sinh xong (tab cha dien san Buoc 2)
    sang_buoc2()      : nut "Chuyen sang Buoc 2"
    dau_trang(than)   : dung them gi o dau trang (thanh bo cau hinh)
    """

    def __init__(self, me, mo_thu_muc_fn, khi_xong=None, sang_buoc2=None,
                 dau_trang=None):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.khi_xong = khi_xong
        self.sang_buoc2_fn = sang_buoc2
        self.dau_trang = dau_trang
        self.hang_doi = queue.Queue()
        self.thu_muc_cuoi = None
        self.kq_cuoi = None
        self.cfg_cuoi = None
        self.co_dung = threading.Event()
        self._dung_giao_dien()
        self._nap_lan_truoc()

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

        if self.dau_trang:
            self.dau_trang(than)
        self._khoi_so_luong(than)
        self._khoi_kich_thuoc(than)
        self._khoi_dai(than)
        self._khoi_spc(than)
        self._khoi_chay(than)

    def _o(self, me, r, c, nhan, var, w=10, goi_y=None):
        ttk.Label(me, text=nhan, style="Surface.TLabel").grid(row=r, column=c, sticky="w",
                                      padx=(0, 8), pady=4)
        ttk.Entry(me, textvariable=var, width=w).grid(
            row=r, column=c + 1, sticky="w", padx=(0, 18), ipady=2)
        if goi_y:
            ttk.Label(me, text=goi_y, style="SurfacePhu.TLabel").grid(
                row=r, column=c + 2, sticky="w", padx=(0, 16))

    # ------------------------------------------------- 1. so luong, ma mau
    def _khoi_so_luong(self, me):
        c = ui.ModernCard(me, title="Số lượng và quy tắc đặt mã mẫu", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()

        self.v_n = tk.StringVar(value="1000")
        self.v_tt = tk.StringVar(value="M")
        self.v_cs = tk.StringVar(value="4")
        self.v_bd = tk.StringVar(value="1")
        self.v_seed = tk.StringVar(value="")
        self._o(k, 0, 0, "Số lượng mẫu cần sinh", self.v_n, 10, "VD: 1000")
        self._o(k, 0, 3, "Seed", self.v_seed, 10, "để trống = mỗi lần khác")
        self._o(k, 1, 0, "Tiền tố mã mẫu", self.v_tt, 10)
        self._o(k, 1, 3, "Số chữ số", self.v_cs, 10, 'M + 4 chữ số → M0001')
        self._o(k, 2, 0, "Số bắt đầu", self.v_bd, 10)

    # ---------------------------------------------------- 2. kich thuoc mau
    def _khoi_kich_thuoc(self, me):
        c = ui.ModernCard(me, title="Kích thước mẫu thử tiêu chuẩn", badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        self.v_d0 = tk.StringVar(value="14")
        self.v_ds = tk.StringVar(value="0.02")
        self.v_lo = tk.StringVar(value="70")
        self.v_le = tk.StringVar(value="50")
        self._o(k, 0, 0, "Đường kính danh nghĩa d₀ (mm)", self.v_d0, 10)
        self._o(k, 0, 3, "Dung sai ± (mm)", self.v_ds, 10, "0 = mọi mẫu cùng d₀")
        self._o(k, 1, 0, "Chiều dài chuẩn L₀ (mm)", self.v_lo, 10)
        self._o(k, 1, 3, "Chiều dài đo L_e (mm)", self.v_le, 10)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
                  text="Độ thắt Z = (S₀ − Su)/S₀ × 100 %, với S₀ = π·d²/4, "
                       "Su = π·du²/4 (JIS Z 2241, như báo cáo mẫu 522). Đường "
                       "kính chỗ thắt du suy ngược từ Z và d, nên Z trong Excel "
                       "và trên báo cáo Word luôn trùng nhau.").grid(
            row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))

    # -------------------------------------------- 3. dai sinh cua 5 cot SPC
    def _khoi_dai(self, me):
        c = ui.ModernCard(me, title="Dải sinh của năm cột chịu đánh giá SPC", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        for j, nhan in enumerate(["Trung bình", "Độ lệch chuẩn",
                                  "Ngưỡng tối thiểu (dòng 11 BM32)"]):
            ttk.Label(k, text=nhan, style="SurfacePhu.TLabel").grid(
                row=0, column=1 + j, sticky="w", padx=(0, 14))

        self.v_dai = {}
        dong = [("ReL (MPa)", "reL", "372", "4.2", "275"),
                ("Rm (MPa)", "rm", "571", "4.5", "480"),
                ("A (%)", "a", "29.5", "0.85", "20"),
                ("Z (%)", "z", "54.6", "1.3", "")]
        for i, (nhan, khoa, tb, sd, ng) in enumerate(dong):
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(row=1 + i, column=0, sticky="w",
                                         padx=(0, 10), pady=3)
            vt, vs, vn = (tk.StringVar(value=tb), tk.StringVar(value=sd),
                          tk.StringVar(value=ng))
            for j, v in enumerate((vt, vs, vn)):
                ttk.Entry(k, textvariable=v, width=10).grid(
                    row=1 + i, column=1 + j, sticky="w", padx=(0, 14), ipady=2)
            self.v_dai[khoa] = (vt, vs, vn)

        ttk.Label(k, text="ReH (MPa)", style="Surface.TLabel").grid(row=5, column=0, sticky="w",
                                            padx=(0, 10), pady=3)
        self.v_reh_min = tk.StringVar(value="28")
        self.v_reh_max = tk.StringVar(value="45")
        ttk.Label(k, text="cao hơn ReL từ", style="SurfacePhu.TLabel").grid(
            row=5, column=1, sticky="w")
        ttk.Entry(k, textvariable=self.v_reh_min, width=8).grid(
            row=5, column=2, sticky="w", ipady=2)
        ttk.Label(k, text="đến", style="SurfacePhu.TLabel").grid(row=5, column=3,
                                                          sticky="w")
        ttk.Entry(k, textvariable=self.v_reh_max, width=8).grid(
            row=5, column=4, sticky="w", ipady=2)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
                  text="Ngưỡng để trống nghĩa là cột đó không có giới hạn "
                       "nghiệm thu — vẫn xét ổn định nhưng không xét Cpk. "
                       "Biểu mẫu BM32 chỉ đặt ngưỡng cho ReL, Rm và A.").grid(
            row=6, column=0, columnspan=6, sticky="w", pady=(6, 0))

    # ------------------------------------------------------- 4. danh gia SPC
    def _khoi_spc(self, me):
        c = ui.ModernCard(me, title="Đánh giá độ ổn định SPC & Thiết lập tối ưu", badge_num=4)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        self.v_cpk = tk.StringVar(value="1.33")
        self.v_giay = tk.StringVar(value="180")
        self.v_buoc = tk.StringVar(value=", ".join(str(x) for x in BUOC_TRON))
        self.v_muc = tk.StringVar(value="lien_tiep")

        ttk.Label(k, text="Bảo đảm cách lấy 25 giá trị nào?",
                  style="Bold.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 4))

        ttk.Radiobutton(
            k, text="Liên tiếp  —  bất kỳ 25 mẫu LIỀN NHAU nào, ở bất kỳ đâu "
                    "trong lô   (khoảng 15 giây)",
            variable=self.v_muc, value="lien_tiep",
            command=self._doi_muc).grid(row=1, column=0, columnspan=6,
                                        sticky="w")
        ttk.Radiobutton(
            k, text="Cơ bản  —  15 bước nhảy tròn, nhưng chỉ khi bắt đầu từ "
                    "đầu lô   (khoảng 20 giây)",
            variable=self.v_muc, value="co_ban",
            command=self._doi_muc).grid(row=2, column=0, columnspan=6,
                                        sticky="w", pady=(2, 0))
        ttk.Radiobutton(
            k, text="Gộp cả hai  —  mạnh nhất, nhưng 1000 mẫu mất 10 phút và "
                    "thường vẫn còn vài khối chưa đạt (đo được 5/5860)",
            variable=self.v_muc, value="gop",
            command=self._doi_muc).grid(row=3, column=0, columnspan=6,
                                        sticky="w", pady=(2, 0))
        ttk.Radiobutton(k, text="Tự khai bước nhảy (tính từ đầu lô):",
                        variable=self.v_muc, value="tu",
                        command=self._doi_muc).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(2, 0))
        self.e_buoc = ttk.Entry(k, textvariable=self.v_buoc, width=52)
        self.e_buoc.grid(row=4, column=2, columnspan=4, sticky="w",
                         padx=(6, 0), pady=(2, 0), ipady=2)

        self.v_day = tk.StringVar(value="")
        ttk.Label(k, text="Dãy phải đạt (bắt đầu:bước)", style="Surface.TLabel").grid(
            row=5, column=0, sticky="w", padx=(0, 6), pady=(8, 0))
        ttk.Entry(k, textvariable=self.v_day, width=52).grid(
            row=5, column=1, columnspan=5, sticky="w", padx=(6, 0),
            pady=(8, 0), ipady=2)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
                  text="Biết trước sẽ nộp báo cáo SPC lấy theo cách nào thì "
                       "khai thẳng vào đây, ví dụ  1:5, 26:5, 1:10  — bộ dữ "
                       "liệu sinh ra sẽ thoả sẵn đúng những dãy đó. Rẻ hơn "
                       "hẳn so với mở rộng phạm vi bảo đảm, vì mỗi dãy chỉ "
                       "thêm 5 khối ràng buộc.").grid(
            row=6, column=0, columnspan=6, sticky="w", pady=(4, 0))

        self._o(k, 7, 0, "Cpk tối thiểu", self.v_cpk, 10)
        self._o(k, 7, 3, "Thời gian tối đa (giây)", self.v_giay, 10,
                "hết giờ thì lấy kết quả tốt nhất")

        self.v_spc = tk.BooleanVar(value=True)
        self.v_cong = tk.BooleanVar(value=True)
        ttk.Checkbutton(k, text="Xuất kèm file đánh giá SPC cho 25 mẫu đầu "
                                "(cột ReL)", variable=self.v_spc).grid(
            row=8, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Checkbutton(k, text="Kèm sheet Du_lieu (đường cong từng mẫu) — "
                                "bỏ tick cho file Excel nhẹ hơn",
                        variable=self.v_cong).grid(
            row=9, column=0, columnspan=6, sticky="w")
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
                  text="Bộ luật dùng đủ 8 mục I-chart + 3 mục MR-chart + Cpk, "
                       "đúng như biểu mẫu BÁO CÁO ĐÁNH GIÁ ĐỘ ỔN ĐỊNH CHẤT "
                       "LƯỢNG.\n"
                       "KHÔNG thể bảo đảm mọi cách lấy 25 giá trị — bộ luật "
                       "vốn được thiết kế để một quá trình ổn định thật vẫn "
                       "thỉnh thoảng chạm phải. Sau khi sinh, phần mềm quét "
                       "lại và ghi rõ cách lấy nào đạt, cách nào không, vào "
                       "file BaoCao_Phu_SPC.txt.").grid(
            row=10, column=0, columnspan=6, sticky="w", pady=(6, 0))
        self._doi_muc()

    def _doi_muc(self):
        m = self.v_muc.get()
        self.e_buoc.configure(state="normal" if m == "tu" else "disabled")
        if m in MUC:
            self.v_buoc.set(", ".join(str(x) for x in MUC[m][0]))

    # ---------------------------------------------------------- 5. nut chay
    def _khoi_chay(self, me):
        d = tk.Frame(me, bg=ui.COLOR_BG, padx=4, pady=8)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   SINH BIỂU ĐỒ VÀ SỐ LIỆU",
                              style="Primary.TButton", command=self.chay)
        self.btn.pack(side="left")
        self.btn_dung = ttk.Button(d, text="⏹ Dừng", style="Danger.TButton", state="disabled",
                                   command=self.dung)
        self.btn_dung.pack(side="left", padx=8)
        ttk.Button(d, text="📂 Mở thư mục kết quả", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=6)
        self.btn_tab2 = ttk.Button(d, text="➜ Chuyển sang Bước 2", style="Success.TButton",
                                   state="disabled", command=self.sang_tab2)
        self.btn_tab2.pack(side="left", padx=6)
        self.tien_do = ttk.Progressbar(d, mode="determinate", length=220)
        self.tien_do.pack(side="left", padx=10)
        self.lbl_td = ttk.Label(d, style="Phu.TLabel", text="")
        self.lbl_td.pack(side="left")

        # ---- Modern Terminal Log ----
        self.log = ui.ModernTerminal(me, title="NHẬT KÝ SINH BIỂU ĐỒ KÉO & BẢO ĐẢM SPC", height=11)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self._ghi("Điền số lượng mẫu rồi bấm nút '▶ SINH BIỂU ĐỒ VÀ SỐ LIỆU'. Các ô còn lại đã có sẵn giá trị chuẩn cho SCW480.\n"
                  "1000 mẫu mất khoảng 70 giây, ra chừng 5 MB.\n")

    # ============================================================ TIEN ICH
    def _ghi(self, s, xoa=False):
        self.log.ghi(s, xoa=xoa)

    def _thu_cfg(self):
        def so(v, ten, nguyen=False, cho_trong=False):
            v = str(v).strip().replace(",", ".")
            if not v:
                if cho_trong:
                    return None
                raise keo_core.LoiNghiepVu("Chưa nhập: %s" % ten)
            try:
                return keo_core.spc_luat.so_nguyen(v, ten, 0) if nguyen else keo_core.spc_luat.so_huu_han(v, ten)
            except ValueError:
                raise keo_core.LoiNghiepVu(
                    "Ô '%s' phải là số. Đang là: %s" % (ten, v))

        m = self.v_muc.get()
        if m in MUC:
            buoc, buoc_moi_diem = MUC[m]
            buoc = list(buoc)
        else:
            buoc_moi_diem = []
            buoc = []
            for x in str(self.v_buoc.get()).replace(";", " ") \
                    .replace(",", " ").split():
                try:
                    k = keo_core.spc_luat.so_nguyen(x, "Buoc nhay")
                except ValueError:
                    raise keo_core.LoiNghiepVu(
                        "Ô danh sách bước nhảy chỉ được chứa số nguyên. "
                        "Đang có: %s" % x)
                if k >= 1:
                    buoc.append(k)
        day_rieng = []
        for x in str(self.v_day.get()).replace(";", " ").replace(",", " ").split():
            m = re.match(r"^(\d+)\s*[:/]\s*(\d+)$", x)
            if not m:
                raise keo_core.LoiNghiepVu(
                    "Ô 'Dãy phải đạt' phải viết dạng  batdau:buoc  cách nhau "
                    "bằng dấu phẩy. Ví dụ: 1:5, 26:5.  Đang có: %s" % x)
            day_rieng.append((int(m.group(1)), int(m.group(2))))

        so_luong_mau = so(self.v_n.get(), "Số lượng mẫu", True)
        # Tu dong bao dam cac day danh gia tieu chuan cua Tab 4 (buoc 5 o cac moc dau lo)
        if not day_rieng and m in ("lien_tiep", "co_ban", "gop"):
            for bd in (1, 26, 51, 76, 101, 126, 151, 176):
                if bd - 1 + 25 * 5 <= so_luong_mau:
                    day_rieng.append((bd, 5))

        if not buoc and not buoc_moi_diem and not day_rieng:
            raise keo_core.LoiNghiepVu("Chưa khai bước nhảy nào.")

        seed = self.v_seed.get().strip()
        c = {
            "so_luong": so_luong_mau,
            "seed": int(seed) if seed.lstrip("-").isdigit() else None,
            "tien_to_ma": self.v_tt.get().strip(),
            "so_chu_so": so(self.v_cs.get() or 0, "Số chữ số", True),
            "so_bat_dau": so(self.v_bd.get() or 1, "Số bắt đầu", True),
            "d0": so(self.v_d0.get(), "Đường kính d₀"),
            "dung_sai_d": so(self.v_ds.get() or 0, "Dung sai đường kính"),
            "Lo": so(self.v_lo.get(), "Chiều dài chuẩn L₀"),
            "Le": so(self.v_le.get(), "Chiều dài đo L_e"),
            "reH_chenh_min": so(self.v_reh_min.get(), "ReH cao hơn ReL từ"),
            "reH_chenh_max": so(self.v_reh_max.get(), "ReH cao hơn ReL đến"),
            "cpk_toi_thieu": so(self.v_cpk.get(), "Cpk tối thiểu"),
            "giay_toi_da": so(self.v_giay.get(), "Thời gian tối đa"),
            "buoc_nhay": buoc,
            "buoc_moi_diem": buoc_moi_diem,
            "day_rieng": day_rieng,
            "xuat_spc": bool(self.v_spc.get()),
            "kem_duong_cong": bool(self.v_cong.get()),
        }
        ten_cfg = {"reL": ("reL_tb", "reL_sd", "reL_min"),
                   "rm": ("rm_tb", "rm_sd", "rm_min"),
                   "a": ("a_tb", "a_sd", "a_min"),
                   "z": ("z_tb", "z_sd", "z_min_spec")}
        for khoa, (vt, vs, vn) in self.v_dai.items():
            k_tb, k_sd, k_ng = ten_cfg[khoa]
            c[k_tb] = so(vt.get(), "%s trung bình" % khoa)
            c[k_sd] = so(vs.get(), "%s độ lệch chuẩn" % khoa)
            c[k_ng] = so(vn.get(), "%s ngưỡng" % khoa, cho_trong=True)
        return keo_core.kiem_tra_cfg(c)

    # ================================================================ CHAY
    def chay(self):
        try:
            cfg = self._thu_cfg()
        except keo_core.LoiNghiepVu as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        if cfg["so_luong"] < 25:
            messagebox.showwarning(
                "Số lượng quá ít",
                "Một bộ đánh giá SPC cần 25 điểm, nên số mẫu phải từ 25 trở "
                "lên.\n\nĐang khai %d." % cfg["so_luong"])
            return
        if cfg["so_luong"] > 2000:
            if not messagebox.askyesno(
                    "Số lượng lớn",
                    "Bạn sắp sinh %s biểu đồ. Việc này mất khoảng %d phút và "
                    "chiếm chừng %d MB đĩa.\n\nTiếp tục?"
                    % ("{:,}".format(cfg["so_luong"]),
                       max(1, cfg["so_luong"] // 900),
                       max(1, cfg["so_luong"] * 5 // 1000))):
                return

        self.co_dung.clear()
        self.btn.configure(state="disabled", text="Đang sinh...")
        self.btn_dung.configure(state="normal")
        self.btn_tab2.configure(state="disabled")
        self.tien_do.configure(maximum=100, value=0)
        self.lbl_td.configure(text="")
        self._ghi("Đang sinh %s mẫu...\n"
                  "Bước 1: tối ưu thứ tự cho đạt SPC (có thể mất một lúc).\n"
                  % "{:,}".format(cfg["so_luong"]), xoa=True)
        threading.Thread(target=self._chay_nen, args=(cfg,),
                         daemon=True).start()
        self.me.after(150, self._kiem_tra_hang_doi)

    def dung(self):
        self.co_dung.set()
        self.btn_dung.configure(state="disabled", text="Đang dừng...")

    def _chay_nen(self, cfg):
        try:
            def td(giai_doan, a, b):
                self.hang_doi.put(("td", (giai_doan, a, b)))
            kq = xuat_keo.chay(cfg, tien_do=td, dung_lai=self.co_dung.is_set)
            self.cfg_cuoi = cfg
            self._luu_lan_truoc(cfg)
            self.hang_doi.put(("xong", kq))
        except keo_core.LoiNghiepVu as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi",
                               "Không ghi được file kết quả.\n\nThường do file "
                               "đang mở, hoặc thư mục bị khoá quyền ghi.\n\n%s"
                               % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def _kiem_tra_hang_doi(self):
        tiep = True
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "td":
                    gd, a, b = dl
                    if gd == "spc":
                        self.tien_do.configure(mode="indeterminate")
                        self.tien_do.start(30)
                        self.lbl_td.configure(
                            text="tối ưu SPC: còn %d khối, %.0fs" % (a, b))
                    else:
                        self.tien_do.stop()
                        self.tien_do.configure(mode="determinate",
                                               maximum=b, value=a)
                        self.lbl_td.configure(
                            text="vẽ biểu đồ %s / %s"
                                 % ("{:,}".format(a), "{:,}".format(b)))
                elif loai == "xong":
                    self._xong(dl); tiep = False
                elif loai == "loi":
                    self._loi(dl); tiep = False
                elif loai == "loi_he_thong":
                    self._loi_he_thong(dl); tiep = False
        except queue.Empty:
            pass
        if tiep:
            self.me.after(150, self._kiem_tra_hang_doi)

    def _tra_nut(self):
        self.tien_do.stop()
        self.tien_do.configure(mode="determinate")
        self.btn.configure(state="normal", text="▶   SINH BIỂU ĐỒ VÀ SỐ LIỆU")
        self.btn_dung.configure(state="disabled", text="Dừng")

    def _xong(self, kq):
        self._tra_nut()
        self.thu_muc_cuoi = kq["thu_muc"]
        self.kq_cuoi = kq
        self.btn_tab2.configure(state="normal")
        self.tien_do.configure(value=self.tien_do["maximum"])
        self._ghi(xuat_keo.bao_cao(kq) + "\n\n" + kq["bang_phu"] + "\n",
                  xoa=True)
        if self.khi_xong:
            try:
                self.khi_xong(kq, self.cfg_cuoi)
                self._ghi("\n➜ Đã điền sẵn file PDF, Excel và kích thước mẫu "
                          "sang Bước 2.\n")
            except Exception:
                self._ghi("\n(Không điền được sang Bước 2:\n%s)\n"
                          % traceback.format_exc())
        if self.sang_buoc2_fn:
            if messagebox.askyesno(
                    "Xong", "Đã sinh xong %s mẫu trong %.1f giây.\n\n"
                            "Chuyển sang Bước 2 để gộp báo cáo Word luôn?"
                            % ("{:,}".format(len(kq["mau"])), kq["giay"])):
                self.sang_buoc2_fn()
        elif messagebox.askyesno(
                "Xong", "Đã sinh xong %s mẫu trong %.1f giây.\n\n"
                        "Mở thư mục kết quả bây giờ?"
                        % ("{:,}".format(len(kq["mau"])), kq["giay"])):
            self.mo_thu_muc()

    def _loi(self, msg):
        self._tra_nut()
        self._ghi("KHÔNG SINH ĐƯỢC:\n\n" + msg + "\n", xoa=True)
        messagebox.showerror("Không sinh được", msg)

    def _loi_he_thong(self, tb):
        self._tra_nut()
        self._ghi("LỖI KHÔNG MONG ĐỢI:\n\n" + tb, xoa=True)
        messagebox.showerror("Lỗi",
                             "Có lỗi ngoài dự kiến. Chi tiết ở khung Kết quả.")

    def sang_tab2(self):
        """Dien lai cap file vua sinh vao Buoc 2 roi chuyen sang do."""
        if not self.kq_cuoi:
            return
        if self.khi_xong:
            self.khi_xong(self.kq_cuoi, self.cfg_cuoi)
        if self.sang_buoc2_fn:
            self.sang_buoc2_fn()

    def mo_thu_muc(self):
        d = self.thu_muc_cuoi
        if not (d and os.path.isdir(d)):
            d = os.path.join(keo_core.base_dir(), "output")
            os.makedirs(d, exist_ok=True)
        self.mo_thu_muc_ngoai(d)

    # ============================================================== GHI NHO
    def _luu_lan_truoc(self, cfg):
        bo_cau_hinh.nho_lan_truoc("keo_buoc1", cfg)

    def _nap_lan_truoc(self):
        c = bo_cau_hinh.doc_lan_truoc("keo_buoc1", "lan_dung_truoc_sinh_keo.json")
        if not c:
            return
        cap = [("so_luong", self.v_n), ("tien_to_ma", self.v_tt),
               ("so_chu_so", self.v_cs), ("so_bat_dau", self.v_bd),
               ("d0", self.v_d0), ("dung_sai_d", self.v_ds),
               ("Lo", self.v_lo), ("Le", self.v_le),
               ("reH_chenh_min", self.v_reh_min),
               ("reH_chenh_max", self.v_reh_max),
               ("cpk_toi_thieu", self.v_cpk), ("giay_toi_da", self.v_giay)]
        for k, var in cap:
            v = c.get(k)
            if v not in (None, ""):
                var.set(("%g" % v) if isinstance(v, float) else str(v))
        ten_cfg = {"reL": ("reL_tb", "reL_sd", "reL_min"),
                   "rm": ("rm_tb", "rm_sd", "rm_min"),
                   "a": ("a_tb", "a_sd", "a_min"),
                   "z": ("z_tb", "z_sd", "z_min_spec")}
        for khoa, (vt, vs, vn) in self.v_dai.items():
            for var, k in zip((vt, vs, vn), ten_cfg[khoa]):
                v = c.get(k)
                var.set("" if v is None else
                        (("%g" % v) if isinstance(v, float) else str(v)))
        if c.get("buoc_nhay") is not None:
            b = [int(x) for x in c["buoc_nhay"]]
            md = [int(x) for x in (c.get("buoc_moi_diem") or [])]
            ten = next((t for t, (x, y) in MUC.items()
                        if x == b and y == md), "tu")
            self.v_muc.set(ten)
            if ten == "tu":
                self.v_buoc.set(", ".join(str(x) for x in b))
            self._doi_muc()
        for k, var in (("xuat_spc", self.v_spc),
                       ("kem_duong_cong", self.v_cong)):
            if k in c:
                var.set(bool(c[k]))
        self._ghi("(Đã nhớ lại tham số của lần dùng trước.)\n")

    # ======================================================== BO CAU HINH
    def lay_cau_hinh(self):
        return bo_cau_hinh.chup_bien(self)

    def dat_cau_hinh(self, d):
        bo_cau_hinh.dat_bien(self, d, bo_qua=("v_buoc",))
        self._doi_muc()
        if self.v_muc.get() == "tu" and "v_buoc" in (d or {}):
            self.v_buoc.set(d["v_buoc"])
