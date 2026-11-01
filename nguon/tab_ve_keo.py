# -*- coding: utf-8 -*-
"""
THU KEO - VE BIEU DO TU SO LIEU CO SAN
=======================================
Giao dien cho ve_keo_so_lieu.py. Hai che do:
    A. So lieu tho cua may thu (file luc - hanh trinh)  -> bieu do do thuc
    B. Diem dac trung FeL, ReL, Fm, A, Z                -> duong cong MINH HOA
Tab nay dung rieng: khong dien gi sang Buoc 2 (gop bao cao Word).
"""

import os
import sys
import queue
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ve_keo_so_lieu as ve_sl
import sinh_keo_core as keo_core
import ui_theme as ui

DON_VI = [("Tự nhận", "auto"), ("kN", "kN"), ("N", "N")]
KIEU_FILE = [("Số liệu máy", "*.csv *.txt *.dat *.tsv *.xlsx *.xlsm"),
             ("Tất cả", "*.*")]


class TabVeKeo(object):
    """Dung giao dien ve bieu do tu so lieu vao khung 'me'."""

    def __init__(self, me, mo_thu_muc_fn):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.hang_doi = queue.Queue()
        self.thu_muc_cuoi = None
        self.ds_duong_dan = []
        self._dung_giao_dien()

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

        self._khoi_che_do(than)
        self.khung_a = tk.Frame(than, bg=ui.COLOR_BG)
        self.khung_b = tk.Frame(than, bg=ui.COLOR_BG)
        self._khoi_may(self.khung_a)
        self._khoi_dac_trung(self.khung_b)
        self.khung_chay = tk.Frame(than, bg=ui.COLOR_BG)
        self.khung_chay.pack(fill="both", expand=True)
        self._khoi_chay(self.khung_chay)
        self._doi_che_do()

    def _o(self, me, r, c, nhan, var, w=10, goi_y=None):
        ttk.Label(me, text=nhan, style="Surface.TLabel").grid(
            row=r, column=c, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(me, textvariable=var, width=w).grid(
            row=r, column=c + 1, sticky="w", padx=(0, 18), ipady=2)
        if goi_y:
            ttk.Label(me, text=goi_y, style="SurfacePhu.TLabel").grid(
                row=r, column=c + 2, sticky="w", padx=(0, 16))

    # ------------------------------------------------------- 1. che do
    def _khoi_che_do(self, me):
        c = ui.ModernCard(me, title="Bạn đang có số liệu gì?", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        self.v_che_do = tk.StringVar(value="may")
        ttk.Radiobutton(
            k, variable=self.v_che_do, value="may", command=self._doi_che_do,
            text="Số liệu thô của máy thử (file lực – hành trình)  →  vẽ "
                 "đúng đường cong đo được, tự tìm FeH, FeL, Fm").grid(
            row=0, column=0, sticky="w")
        ttk.Radiobutton(
            k, variable=self.v_che_do, value="minh_hoa",
            command=self._doi_che_do,
            text="Chỉ có điểm đặc trưng (FeL, ReL, Fm, A, Z)  →  vẽ đường "
                 "cong MINH HOẠ, có ghi rõ trên hình").grid(
            row=1, column=0, sticky="w", pady=(4, 0))

    # ------------------------------------------------ 2A. so lieu may
    def _khoi_may(self, me):
        c = ui.ModernCard(me, title="File số liệu của máy thử", badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        k.columnconfigure(1, weight=1)

        self.v_file = tk.StringVar(value="")
        self.v_kt = tk.StringVar(value="")
        ttk.Label(k, text="File / thư mục số liệu", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(k, textvariable=self.v_file, state="readonly").grid(
            row=0, column=1, sticky="ew", ipady=2)
        ttk.Button(k, text="📄 Chọn file...", style="Outline.TButton",
                   width=14, command=self.chon_file).grid(
            row=0, column=2, padx=(8, 0))
        ttk.Button(k, text="📂 Chọn thư mục...", style="Outline.TButton",
                   width=17, command=self.chon_thu_muc).grid(
            row=0, column=3, padx=(8, 0))
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=900, justify="left",
                  text="Nhận CSV, TXT, XLSX. Mỗi file (hoặc mỗi sheet) một "
                       "mẫu, tên file là mã mẫu; hoặc một bảng có một cột "
                       "hành trình và nhiều cột lực, mỗi cột một mẫu.").grid(
            row=1, column=1, columnspan=3, sticky="w", pady=(2, 6))

        ttk.Label(k, text="File kích thước đo tay", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(k, textvariable=self.v_kt).grid(
            row=2, column=1, sticky="ew", ipady=2)
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=14,
                   command=self.chon_kich_thuoc).grid(
            row=2, column=2, padx=(8, 0))
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=900, justify="left",
                  text="Không bắt buộc. Bảng Excel có các cột: Mau (hoặc NO), "
                       "d, du, Lo, Lu. Thiếu thì gõ thẳng vào file kết quả "
                       "sau, các cột Re, Rm, A, Z tự tính lại.").grid(
            row=3, column=1, columnspan=3, sticky="w", pady=(2, 0))

        c2 = ui.ModernCard(me, title="Cột số liệu và kích thước mẫu",
                           badge_num=3)
        c2.pack(fill="x", padx=12, pady=4)
        k2 = c2.get_body()
        self.v_cot_ht = tk.StringVar(value="")
        self.v_cot_luc = tk.StringVar(value="")
        self.v_don_vi = tk.StringVar(value=DON_VI[0][0])
        self.v_d0 = tk.StringVar(value="14")
        self.v_lo = tk.StringVar(value="70")
        self.v_bo_tut = tk.BooleanVar(value=True)
        self._o(k2, 0, 0, "Cột hành trình (mm)", self.v_cot_ht, 8)
        self._o(k2, 0, 3, "Cột lực", self.v_cot_luc, 8,
                "để trống = tự nhận theo tiêu đề; hoặc gõ B, C… / 2, 3…")
        ttk.Label(k2, text="Đơn vị lực", style="Surface.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Combobox(k2, textvariable=self.v_don_vi, state="readonly", width=8,
                     values=[x[0] for x in DON_VI]).grid(
            row=1, column=1, sticky="w", padx=(0, 18))
        self._o(k2, 2, 0, "Đường kính danh nghĩa d₀ (mm)", self.v_d0, 8)
        self._o(k2, 2, 3, "Chiều dài chuẩn L₀ (mm)", self.v_lo, 8,
                "dùng cho mẫu chưa có d, L₀ trong file kích thước")
        ttk.Checkbutton(
            k2, variable=self.v_bo_tut,
            text="FeL bỏ qua lần tụt đầu tiên ngay sau FeH (hiệu ứng chuyển "
                 "tiếp, theo ISO 6892-1 / JIS Z 2241)").grid(
            row=3, column=0, columnspan=6, sticky="w", pady=(4, 0))

    # --------------------------------------------- 2B. diem dac trung
    def _khoi_dac_trung(self, me):
        c = ui.ModernCard(me, title="Bảng điểm đặc trưng của từng mẫu",
                          badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        d = tk.Frame(k, bg=ui.COLOR_SURFACE)
        d.pack(fill="x")
        ttk.Button(d, text="📂 Nhập từ file Excel...", style="Outline.TButton",
                   command=self.nhap_excel).pack(side="left")
        ttk.Button(d, text="🧹 Xoá bảng", style="Outline.TButton",
                   command=lambda: self.txt.delete("1.0", "end")).pack(
            side="left", padx=8)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=940, justify="left",
                  text="Dán các dòng từ Excel vào ô dưới, thứ tự cột: Mã lô, "
                       "FeL (N), ReL (N/mm²), Fm (N), Rm (N/mm²), A (%), "
                       "Z (%). Mỗi dòng một mẫu.").pack(anchor="w", pady=(6, 4))
        self.txt = tk.Text(k, height=9, font=("Consolas", 10), wrap="none",
                           relief="solid", borderwidth=1)
        self.txt.pack(fill="x")
        tk.Label(k, bg=ui.COLOR_WARNING_BG, fg=ui.COLOR_WARNING, justify="left",
                 wraplength=940, padx=8, pady=6, font=("Segoe UI", 9), anchor="w",
                 text="Đường cong ở chế độ này là nội suy qua các điểm đặc "
                      "trưng, không phải số đo của máy. Mỗi trang luôn in "
                      "dòng \"ĐƯỜNG CONG MINH HOẠ – KHÔNG PHẢI DỮ LIỆU ĐO\" và "
                      "không dùng để gộp vào báo cáo thử nghiệm.").pack(
            fill="x", pady=(8, 0))

    # ---------------------------------------------------------- nut chay
    def _khoi_chay(self, me):
        d = tk.Frame(me, bg=ui.COLOR_BG, padx=4, pady=8)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   VẼ BIỂU ĐỒ", style="Primary.TButton",
                              command=self.chay)
        self.btn.pack(side="left")
        ttk.Button(d, text="📂 Mở thư mục kết quả", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=8)
        ttk.Button(d, text="📄 Mở file PDF", style="Outline.TButton",
                   command=self.mo_pdf).pack(side="left")
        self.tien_do = ttk.Progressbar(d, mode="determinate", length=220)
        self.tien_do.pack(side="left", padx=10)

        self.log = ui.ModernTerminal(me, title="KẾT QUẢ VẼ BIỂU ĐỒ", height=13)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self._ghi("Chọn loại số liệu ở mục 1, đưa số liệu vào rồi bấm "
                  "'▶ VẼ BIỂU ĐỒ'.\n")

    def _doi_che_do(self):
        self.khung_a.pack_forget()
        self.khung_b.pack_forget()
        khung = self.khung_a if self.v_che_do.get() == "may" else self.khung_b
        khung.pack(fill="x", before=self.khung_chay)

    def _ghi(self, s, xoa=False):
        self.log.ghi(s, xoa=xoa)

    # ============================================================ CHON FILE
    def _dat_duong_dan(self, ds):
        self.ds_duong_dan = list(ds)
        if len(ds) == 1:
            self.v_file.set(ds[0])
        else:
            self.v_file.set("%d file trong %s"
                            % (len(ds), os.path.dirname(ds[0])))

    def chon_file(self):
        ds = filedialog.askopenfilenames(title="Chọn file số liệu của máy thử",
                                         filetypes=KIEU_FILE)
        if ds:
            self._dat_duong_dan(ds)

    def chon_thu_muc(self):
        p = filedialog.askdirectory(title="Chọn thư mục chứa file số liệu máy")
        if p:
            self._dat_duong_dan([p])

    def chon_kich_thuoc(self):
        p = filedialog.askopenfilename(
            title="Chọn file kích thước đo tay",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Tất cả", "*.*")])
        if p:
            self.v_kt.set(p)

    def nhap_excel(self):
        p = filedialog.askopenfilename(
            title="Chọn file có các cột FeL, ReL, Fm, Rm, A, Z",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Tất cả", "*.*")])
        if not p:
            return
        try:
            van_ban = ve_sl.doc_bang_dac_trung(p)
        except ve_sl.LoiNghiepVu as e:
            messagebox.showwarning("Không nhập được", str(e))
            return
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", van_ban + "\n")
        self._ghi("Đã nhập %d dòng từ %s. Kiểm tra lại các cột trước khi vẽ.\n"
                  % (len(van_ban.splitlines()), os.path.basename(p)), xoa=True)

    # ================================================================ CHAY
    def _thu_cfg(self):
        def so(v, ten):
            v = str(v).strip().replace(",", ".")
            if not v:
                return None
            try:
                return float(v)
            except ValueError:
                raise ve_sl.LoiNghiepVu("Ô '%s' phải là số. Đang là: %s"
                                        % (ten, v))
        if not self.ds_duong_dan:
            raise ve_sl.LoiNghiepVu("Chưa chọn file hoặc thư mục số liệu máy.")
        cot_ht, cot_luc = self.v_cot_ht.get().strip(), self.v_cot_luc.get().strip()
        if bool(cot_ht) != bool(cot_luc):
            raise ve_sl.LoiNghiepVu("Ô 'Cột hành trình' và 'Cột lực' phải điền "
                                    "cả hai, hoặc để trống cả hai.")
        return {
            "duong_dan": list(self.ds_duong_dan),
            "file_kich_thuoc": self.v_kt.get().strip() or None,
            "cot_ht": cot_ht or None, "cot_luc": cot_luc or None,
            "don_vi": dict(DON_VI).get(self.v_don_vi.get(), "auto"),
            "d0": so(self.v_d0.get(), "Đường kính d₀"),
            "Lo": so(self.v_lo.get(), "Chiều dài chuẩn L₀"),
            "bo_tut_dau": bool(self.v_bo_tut.get()),
        }

    def chay(self):
        che_do = self.v_che_do.get()
        try:
            if che_do == "may":
                viec = ("may", self._thu_cfg())
            else:
                viec = ("minh_hoa",
                        ve_sl.tach_dong_dan(self.txt.get("1.0", "end")))
        except ve_sl.LoiNghiepVu as e:
            messagebox.showwarning("Thiếu hoặc sai số liệu", str(e))
            return
        self.btn.configure(state="disabled", text="Đang vẽ...")
        self.tien_do.configure(maximum=100, value=0)
        self._ghi("Đang đọc số liệu và vẽ...\n", xoa=True)
        threading.Thread(target=self._chay_nen, args=viec, daemon=True).start()
        self.me.after(150, self._kiem_tra_hang_doi)

    def _chay_nen(self, che_do, dl):
        try:
            if che_do == "may":
                kq = ve_sl.chay_do(
                    dl, tien_do=lambda a, b: self.hang_doi.put(("td", (a, b))))
                self.hang_doi.put(("xong", (kq, ve_sl.bao_cao_do(kq))))
            else:
                kq = ve_sl.chay_minh_hoa(dl)
                self.hang_doi.put(("xong", (kq, (
                    "ĐÃ VẼ XONG %d đường cong MINH HOẠ.\n  PDF  : %s\n"
                    "  Excel: %s\n\nChú ý: đây là đường cong nội suy từ điểm "
                    "đặc trưng, không phải dữ liệu đo. Trên mỗi trang đã ghi "
                    "rõ điều này."
                    % (len(kq["mau"]), kq["pdf"], kq["xlsx"])))))
        except ve_sl.LoiNghiepVu as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi", "Không ghi được file kết quả.\n\nThường "
                                      "do file đang mở, hoặc thư mục bị khoá "
                                      "quyền ghi.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def _kiem_tra_hang_doi(self):
        tiep = True
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "td":
                    self.tien_do.configure(maximum=dl[1], value=dl[0])
                else:
                    tiep = False
                    self.btn.configure(state="normal", text="▶   VẼ BIỂU ĐỒ")
                    if loai == "xong":
                        kq, chu = dl
                        self.kq_cuoi = kq
                        self.thu_muc_cuoi = kq["thu_muc"]
                        self.tien_do.configure(value=self.tien_do["maximum"])
                        self._ghi(chu + "\n", xoa=True)
                        if messagebox.askyesno(
                                "Xong", "Đã vẽ xong %d biểu đồ (PDF + Excel số liệu).\n\n"
                                        "Mở thư mục kết quả bây giờ?" % len(kq["mau"])):
                            self.mo_thu_muc()
                    elif loai == "loi":
                        self._ghi("KHÔNG VẼ ĐƯỢC:\n\n" + dl + "\n", xoa=True)
                        messagebox.showerror("Không vẽ được", dl)
                    else:
                        self._ghi("LỖI KHÔNG MONG ĐỢI:\n\n" + dl, xoa=True)
                        messagebox.showerror(
                            "Lỗi", "Có lỗi ngoài dự kiến. Chi tiết ở khung "
                                   "Kết quả.")
        except queue.Empty:
            pass
        if tiep:
            self.me.after(150, self._kiem_tra_hang_doi)

    def mo_thu_muc(self):
        d = self.thu_muc_cuoi
        if not (d and os.path.isdir(d)):
            d = os.path.join(keo_core.base_dir(), "output")
            os.makedirs(d, exist_ok=True)
        self.mo_thu_muc_ngoai(d)

    def mo_pdf(self):
        kq = getattr(self, "kq_cuoi", None)
        if not (kq and os.path.exists(kq["pdf"])):
            messagebox.showinfo("Chưa có file", "Chưa vẽ biểu đồ nào.")
            return
        try:
            os.startfile(kq["pdf"])
        except OSError as e:
            messagebox.showerror("Không mở được", str(e))
