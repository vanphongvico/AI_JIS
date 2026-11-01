# -*- coding: utf-8 -*-
"""
TAB KICH THUOC - BUOC 2: DANH GIA SPC
======================================
Doc lai mot file bao cao kich thuoc (JIS.QT.07.BM.01) da sinh - hoac file
lam tay theo dung bieu mau - cham TUNG CHI TIEU theo dung cong thuc bieu mau
(kich_thuoc_core.cham_bieu_mau) va xuat sheet 'Bao cao' cho chi tieu chon.

Vi sao can: sheet 'Bao cao' cua bieu mau chi trinh bay CHI TIEU 1. Muon co
bao cao (bieu do I-MR, bang danh gia) cho chi tieu khac thi xuat o day -
moi chi tieu mot file, cung so lieu, chi doi chi tieu dang trinh bay.
"""

import os
import re
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import spc_luat
import kich_thuoc_core as ktcore
import xuat_kich_thuoc as xkt


class TabDanhGiaKichThuoc(object):
    COT = ("stt", "ten", "ky_hieu", "yeu_cau", "xb", "sigma", "lcl", "ucl",
           "cpk", "ket_luan")
    TEN_COT = ("STT", "Chỉ tiêu", "Ký hiệu", "Yêu cầu", "X̄", "σ", "LCLI",
               "UCLI", "Cpk", "Kết luận")
    RONG = (40, 230, 60, 150, 70, 50, 65, 65, 55, 260)

    def __init__(self, me, mo_thu_muc_fn):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.du_lieu = None
        self.thu_muc_cuoi = None
        self._dung_giao_dien()

    # =========================================================== GIAO DIEN
    def _dung_giao_dien(self):
        than = ttk.Frame(self.me)
        than.pack(fill="both", expand=True)

        c = ui.ModernCard(than, title="File báo cáo kích thước cần đánh giá", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)
        self.v_file = tk.StringVar(value="")
        self.v_cpk = tk.StringVar(value="1.33")
        ttk.Label(k, text="File báo cáo (.xlsx)", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(k, textvariable=self.v_file).grid(row=0, column=1, sticky="ew", ipady=2)
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=self.chon_file).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(k, text="📥 Nạp", style="Outline.TButton", width=8,
                   command=self.nap_lai).grid(row=0, column=3, padx=(6, 0))
        ttk.Label(k, text="Cpk tối thiểu", style="Surface.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(k, textvariable=self.v_cpk, width=8).grid(
            row=1, column=1, sticky="w", pady=(6, 0), ipady=2)
        self.lbl = ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                             text="Sinh xong ở bước ① thì file vừa sinh tự nạp vào "
                                  "đây. Cũng chọn được file báo cáo làm tay theo "
                                  "đúng biểu mẫu JIS.QT.07.BM.01.")
        self.lbl.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))

        c = ui.ModernCard(than, title="Kết quả đánh giá từng chỉ tiêu", badge_num=2)
        c.pack(fill="both", expand=True, padx=12, pady=4)
        k = c.get_body()
        khung = tk.Frame(k, bg=ui.COLOR_SURFACE)
        khung.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(khung, columns=self.COT, show="headings",
                                 height=12, selectmode="extended")
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
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="Chấm đúng công thức biểu mẫu: σ = ROUND(MR̄/1,128; 2), "
                       "giới hạn làm tròn 1 chữ số, 8 mục I-chart + 2 mục "
                       "MR-chart + Cpk. Bấm vào một dòng để xem chi tiết; giữ "
                       "Ctrl hoặc Shift để chọn nhiều chỉ tiêu.").pack(anchor="w", pady=(4, 0))

        d = tk.Frame(than, bg=ui.COLOR_BG, padx=4, pady=6)
        d.pack(fill="x", padx=12)
        ttk.Button(d, text="☑ Chọn tất cả", style="Outline.TButton",
                   command=self.chon_tat_ca).pack(side="left")
        self.btn = ttk.Button(d, text="▶   XUẤT BÁO CÁO CHO CHỈ TIÊU ĐÃ CHỌN",
                              style="Primary.TButton", command=self.xuat)
        self.btn.pack(side="left", padx=8)
        ttk.Button(d, text="📂 Mở thư mục", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left")

        self.log = ui.ModernTerminal(than, title="CHI TIẾT ĐÁNH GIÁ & XUẤT BÁO CÁO", height=8)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))

    # ============================================================ HANH DONG
    def chon_file(self):
        p = filedialog.askopenfilename(
            title="Chọn file báo cáo kích thước",
            filetypes=[("File Excel", "*.xlsx"), ("Tất cả", "*.*")])
        if p:
            self.dat_san(p)

    def nap_lai(self):
        p = self.v_file.get().strip()
        if not p:
            messagebox.showinfo("Chưa có file", "Chọn file báo cáo kích thước trước.")
            return
        self.dat_san(p)

    def _cpk(self):
        try:
            return spc_luat.kiem_tra_nguong(None, None, str(self.v_cpk.get()).strip().replace(",", "."))[2]
        except ValueError:
            raise ValueError("Cpk tối thiểu phải là số hữu hạn lớn hơn 0.")

    def dat_san(self, p):
        """Nap mot file bao cao va cham tung chi tieu (tab cha goi sau khi sinh)."""
        self.v_file.set(p)
        for i in self.tree.get_children():
            self.tree.delete(i)
        try:
            cpk_min = self._cpk()
            d = xkt.doc_bao_cao(p)
        except Exception as e:
            self.du_lieu = None
            self.lbl.configure(text="Không đọc được: %s" % e)
            self.log.ghi("KHÔNG ĐỌC ĐƯỢC:\n%s\n" % e, xoa=True)
            return
        self.du_lieu = d
        dat = 0
        for k, ct in enumerate(d["chi_tieu"]):
            r = ktcore.cham_chi_tieu(ct, cpk_min)
            ct["kq"] = r
            ct["cpk_toi_thieu"] = cpk_min
            dat += r.dat
            kl = "ĐẠT" if r.dat else "CHƯA ĐẠT: " + ", ".join(r.loi)
            self.tree.insert("", "end", iid=str(k), tags=() if r.dat else ("hong",), values=(
                ct["stt"], ct["ten"].replace("\n", " "), ct["ky_hieu"], ct["yeu_cau"].replace("\n", " "),
                *("-" if v is None else fmt % v for v, fmt in ((r.xb, "%.3f"), (r.sigma, "%.2f"), (r.lcl_i, "%.1f"), (r.ucl_i, "%.1f"))),
                "-" if r.cpk is None else "%.2f" % r.cpk, kl))
        self.thu_muc_cuoi = os.path.dirname(p)
        self.lbl.configure(text="%s · mã %s · %d chỉ tiêu — %d đạt, %d chưa đạt. "
                                "Sheet Báo cáo của file này đang trình bày STT %s."
                           % (d["san_pham"], d["ma"], len(d["chi_tieu"]), dat,
                              len(d["chi_tieu"]) - dat, self._stt_dang_trinh_bay()))
        self.log.ghi("Đã đọc %s\nChọn chỉ tiêu ở bảng trên rồi bấm XUẤT BÁO CÁO để có "
                     "sheet Báo cáo (biểu đồ I-MR, bảng đánh giá) cho chỉ tiêu đó.\n"
                     % os.path.basename(p), xoa=True)

    def _stt_dang_trinh_bay(self):
        d = self.du_lieu
        for ct in d["chi_tieu"]:
            if ct["hang_x"] == d["hang_bao_cao"]:
                return ct["stt"]
        return "?"

    def _khi_chon(self, _e=None):
        sel = self.tree.selection()
        if not sel or not self.du_lieu:
            return
        ct = self.du_lieu["chi_tieu"][int(sel[-1])]
        r = ct["kq"]
        if "DU_LIEU" in r.loi:
            self.log.ghi("CHƯA ĐẠT: " + "\n".join(ct.get("loi_du_lieu", ["Thiếu điểm đo"])) + "\n", xoa=True)
            return
        L = ["STT %s · %s · %s" % (ct["stt"], ct["ten"].replace("\n", " "), ct["yeu_cau"]),
             "  LSL %s / USL %s" % (ct["lsl"], ct["usl"]),
             "  X̄ %.4f   MR̄ %.4f   σ %.2f   UCLI %.1f / LCLI %.1f   UCLMR %.2f"
             % (r.xb, r.mrb, r.sigma, r.ucl_i, r.lcl_i, r.ucl_mr),
             "  Ca %s   Cp %s   Cpk %s" % tuple("-" if v is None else "%.3f" % v
                                                for v in (r.ca, r.cp, r.cpk))]
        if r.dat:
            L.append("  KẾT LUẬN: ĐẠT đủ các mục đánh giá")
        else:
            L.append("  KẾT LUẬN: CHƯA ĐẠT — vi phạm:")
            for m in r.loi:
                L.append("     [%-3s] %s" % (m, spc_luat.TEN_LUAT.get(m, m)))
        self.log.ghi("\n".join(L) + "\n", xoa=True)

    def chon_tat_ca(self):
        self.tree.selection_set(self.tree.get_children())

    def xuat(self):
        if not self.du_lieu:
            messagebox.showwarning("Chưa có dữ liệu", "Nạp file báo cáo kích thước trước.")
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chưa chọn", "Chọn ít nhất một chỉ tiêu trong bảng.")
            return
        d = self.du_lieu
        goc = os.path.splitext(d["file"])[0]
        ra, loi = [], []
        for iid in sel:
            ct = d["chi_tieu"][int(iid)]
            kh = re.sub(r'[\\/:*?"<>|\s]+', "", ct["ky_hieu"])[:12]
            p = "%s__BaoCao_STT%02d%s.xlsx" % (goc, int(ct["stt"]), ("_" + kh) if kh else "")
            try:
                xkt.xuat_bao_cao_chi_tieu(d["file"], ct, d["hang_bao_cao"], p, ct["cpk_toi_thieu"])
                ra.append(p)
            except PermissionError:
                loi.append("%s — file đang mở, đóng lại rồi thử lại" % os.path.basename(p))
            except Exception:
                loi.append("%s\n%s" % (os.path.basename(p), traceback.format_exc()))
        self.thu_muc_cuoi = os.path.dirname(d["file"])
        self._bao_ket_qua(ra, loi)

    def _bao_ket_qua(self, ra, loi):
        self.log.ghi("Đã xuất %d file báo cáo:\n%s\n%s" % (
            len(ra), "\n".join("  " + os.path.basename(x) for x in ra),
            ("\nLỖI:\n" + "\n".join(loi)) if loi else ""), xoa=True)
        if not ra:
            if loi:
                messagebox.showerror("Không xuất được", "Không xuất được file nào:\n\n%s\n\n"
                                     "Chi tiết ở khung bên dưới."
                                     % "\n".join(x.split("\n")[0] for x in loi[:5]))
            return
        them = ("\n%d file chưa xuất được — xem khung bên dưới." % len(loi)) if loi else ""
        if messagebox.askyesno("Xong", "Đã xuất %d file báo cáo.%s\n\nMở thư mục "
                                       "bây giờ?" % (len(ra), them)):
            self.mo_thu_muc()

    def mo_thu_muc(self):
        if self.thu_muc_cuoi and os.path.isdir(self.thu_muc_cuoi):
            self.mo_thu_muc_ngoai(self.thu_muc_cuoi)
