# -*- coding: utf-8 -*-
"""
TAB THANH PHAN HOA - GIAO DIEN
==============================
Chon file du lieu chuan -> chon mac (SCW410/SCW480) -> sua Min/Max neu
can -> Sinh N mau -> xuat Tong hop + Ket qua in + SPC (xuat_thanh_phan.chay).
"""

import copy
import os
import queue
import threading
import time
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import bo_cau_hinh
import thanh_phan_core as tp
import xuat_thanh_phan as xt

KHOA = "thanh_phan"
TIEU_DE_MAC_DINH = "/ Fe 110 / E20-H / - Element Concentrations"


def _so(v, so_le):
    if v is None:
        return "-"
    return ("%." + str(so_le) + "f") % v if so_le is not None else str(v)


class TabThanhPhan(object):
    def __init__(self, me, mo_thu_muc_fn, khi_xong=None):
        """khi_xong(thu_muc): goi sau khi sinh xong - tab cha dung de nap san
        thu muc vua sinh sang buoc 2 Danh gia SPC."""
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.khi_xong = khi_xong
        self.hang_doi = queue.Queue()
        self.chuan = None
        self.ds = []
        self.thu_muc_ket_qua = None
        self.dang_dung = False
        self._dung_giao_dien()
        self._nap_mac_dinh()
        self.nap_cau_hinh(bo_cau_hinh.doc_lan_truoc(KHOA))

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

        self.thanh = bo_cau_hinh.ThanhCauHinh(than, KHOA, "Thành phần hoá",
                                              self.lay_cau_hinh, self.nap_cau_hinh)
        self._khoi_file(than)
        self._khoi_bang(than)
        self._khoi_tuy_chon(than)
        self._khoi_chay(than)

    def _dong_file(self, k, r, nhan, bien, lenh):
        ttk.Label(k, text=nhan, style="Surface.TLabel").grid(
            row=r, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        ttk.Entry(k, textvariable=bien, width=70).grid(
            row=r, column=1, columnspan=2, sticky="ew", ipady=2, pady=(6, 0))
        ttk.Button(k, text="📂 Chọn...", style="Outline.TButton", width=10,
                   command=lenh).grid(row=r, column=3, sticky="w", padx=(8, 0), pady=(6, 0))

    def _khoi_file(self, me):
        c = ui.ModernCard(me, title="Dữ liệu chuẩn & file mẫu", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)
        self.v_file_chuan = tk.StringVar()
        self.v_file_spc = tk.StringVar()
        self.v_file_word = tk.StringVar()
        self.v_file_th = tk.StringVar()
        self.v_mac = tk.StringVar()
        ttk.Label(k, text="Mác thép", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_mac = ttk.Combobox(k, textvariable=self.v_mac, state="readonly", width=18)
        self.cb_mac.grid(row=0, column=1, sticky="w", ipady=2)
        self.cb_mac.bind("<<ComboboxSelected>>", self._khi_doi_mac)
        self._dong_file(k, 1, "File dữ liệu chuẩn", self.v_file_chuan, self._chon_chuan)
        self._dong_file(k, 2, "File mẫu SPC", self.v_file_spc,
                        lambda: self._chon(self.v_file_spc, "Excel", "*.xlsx"))
        self._dong_file(k, 3, "File mẫu Kết quả in", self.v_file_word,
                        lambda: self._chon(self.v_file_word, "Word", "*.docx"))
        self._dong_file(k, 4, "File mẫu Tổng hợp", self.v_file_th,
                        lambda: self._chon(self.v_file_th, "Excel", "*.xlsx"))
        self.lbl = ttk.Label(k, style="SurfacePhu.TLabel", text="")
        self.lbl.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))

    def _khoi_bang(self, me):
        c = ui.ModernCard(me, title="Nguyên tố (double-click Min/Max để sửa)", badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        cols = ("ten", "yeu_cau", "min", "max", "so_le", "spc")
        names = ("Nguyên tố", "Yêu cầu", "Min", "Max", "Số chữ số", "Đánh giá SPC")
        widths = (80, 110, 90, 90, 80, 110)
        khung = tk.Frame(k, bg=ui.COLOR_SURFACE)
        khung.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(khung, columns=cols, show="headings", height=12)
        for col, nm, w in zip(cols, names, widths):
            self.tree.heading(col, text=nm)
            self.tree.column(col, width=w, minwidth=w, anchor="center")
        sb = ttk.Scrollbar(khung, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._sua_o)
        ttk.Label(k, style="SurfacePhu.TLabel",
                  text="💡 Fe = 100 − tổng các nguyên tố. Ceq = C + Mn/6 + Si/24 + Ni/40 "
                       "+ Cr/5 + Mo/4 + V/14. Sửa ở đây không ghi ngược vào file chuẩn.").pack(anchor="w")

    def _khoi_tuy_chon(self, me):
        c = ui.ModernCard(me, title="Tùy chọn", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        self.v_so_mau = tk.StringVar(value="25")
        self.v_stt = tk.StringVar(value="1")
        self.v_seed = tk.StringVar(value="")
        self.v_cpk = tk.StringVar(value="1.33")
        self.v_user = tk.StringVar(value="Operator")
        self.v_project = tk.StringVar(value="216955")
        self.v_ngay = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.v_gio = tk.StringVar(value="08:00")
        self.v_tieu_de = tk.StringVar(value=TIEU_DE_MAC_DINH)
        o = [("Số mẫu", self.v_so_mau, 8), ("STT bắt đầu", self.v_stt, 8),
             ("Seed (trống = ngẫu nhiên)", self.v_seed, 10), ("Cpk tối thiểu", self.v_cpk, 8),
             ("User", self.v_user, 14), ("Project bắt đầu", self.v_project, 10),
             ("Ngày thử (dd/mm/yyyy)", self.v_ngay, 12), ("Giờ bắt đầu (hh:mm)", self.v_gio, 8)]
        for i, (nhan, bien, w) in enumerate(o):
            r, cc = divmod(i, 4)
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(
                row=r, column=cc * 2, sticky="w", padx=(0 if cc == 0 else 16, 6), pady=3)
            ttk.Entry(k, textvariable=bien, width=w).grid(row=r, column=cc * 2 + 1,
                                                         sticky="w", ipady=2)
        ttk.Label(k, text="Tiêu đề máy", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", pady=3)
        ttk.Entry(k, textvariable=self.v_tieu_de, width=60).grid(
            row=2, column=1, columnspan=7, sticky="w", ipady=2)

    def _khoi_chay(self, me):
        c = ui.ModernCard(me, title="Sinh dữ liệu & xuất báo cáo", badge_num=4)
        c.pack(fill="x", padx=12, pady=(4, 8))
        k = c.get_body()
        d = tk.Frame(k, bg=ui.COLOR_SURFACE)
        d.pack(fill="x")
        self.btn_chay = ttk.Button(d, text="▶ SINH & XUẤT", style="Primary.TButton",
                                   command=self._chay)
        self.btn_chay.pack(side="left")
        self.btn_dung = ttk.Button(d, text="■ Dừng", style="Outline.TButton",
                                   command=self._dung, state="disabled")
        self.btn_dung.pack(side="left", padx=(12, 0))
        self.btn_mo = ttk.Button(d, text="📂 Mở thư mục kết quả", style="Outline.TButton",
                                 command=self._mo_ket_qua, state="disabled")
        self.btn_mo.pack(side="left", padx=(12, 0))
        self.console = ui.ModernTerminal(k, title="NHẬT KÝ THÀNH PHẦN HOÁ", height=10)
        self.console.pack(fill="both", expand=True, pady=(6, 0))

    # =========================================================== FILE / BANG
    def _thu_muc_mau(self):
        cho = [os.path.join(bo_cau_hinh.thu_muc_file_mau(), "ThanhPhan"),
               os.path.join(bo_cau_hinh.base_dir(), "Than_phan")]
        return next((t for t in cho if os.path.isdir(t)), None)

    def _nap_mac_dinh(self):
        tm = self._thu_muc_mau()
        if not tm:
            self.lbl.config(text="Chọn file 'Dữ liệu 2.xlsx' để bắt đầu.")
            return
        for bien, ten in ((self.v_file_word, "Kết quả in.docx"),
                          (self.v_file_th, "Tổng hợp thành phần.xlsx")):
            p = os.path.join(tm, ten)
            if os.path.exists(p):
                bien.set(p)
        p = os.path.join(tm, "Dữ liệu 2.xlsx")
        if os.path.exists(p):
            self.v_file_chuan.set(p)
            self._doc_chuan(p)

    def _chon(self, bien, loai, duoi):
        p = filedialog.askopenfilename(filetypes=[(loai, duoi)])
        if p:
            bien.set(p)

    def _chon_chuan(self):
        p = filedialog.askopenfilename(title="Chọn file dữ liệu chuẩn",
                                       filetypes=[("Excel", "*.xlsx")])
        if p:
            self.v_file_chuan.set(p)
            self._doc_chuan(p)

    def _doc_chuan(self, p):
        try:
            self.chuan = tp.doc_chuan(p)
        except Exception as e:
            messagebox.showerror("Lỗi đọc dữ liệu chuẩn", str(e))
            return
        self.cb_mac["values"] = list(self.chuan)
        if self.v_mac.get() not in self.chuan:
            self.v_mac.set(list(self.chuan)[0])
        self._khi_doi_mac(None)
        self.lbl.config(text="✅ Đã nạp %d mác: %s" % (len(self.chuan), ", ".join(self.chuan)))

    def _khi_doi_mac(self, event):
        mac = self.v_mac.get()
        if not self.chuan or mac not in self.chuan:
            return
        self.ds = copy.deepcopy(self.chuan[mac]["nguyen_to"])
        for tm in (self._thu_muc_mau(), os.path.dirname(self.v_file_chuan.get())):
            p = os.path.join(tm or "", xt.ten_mau_spc(mac))
            if tm and os.path.exists(p):
                self.v_file_spc.set(p)
                break
        self._ve_bang()

    def _ve_bang(self):
        self.tree.delete(*self.tree.get_children())
        for n in self.ds:
            if n["chu"] is not None:
                mn = mx = n["chu"]
            elif n["ten"] == "Fe":
                mn = mx = "100 − tổng"
            else:
                mn, mx = _so(n["min"], n["so_le"]), _so(n["max"], n["so_le"])
            self.tree.insert("", "end", iid=n["ten"], values=(
                n["ten"], n["yeu_cau"] if n["yeu_cau"] is not None else "-", mn, mx,
                n["so_le"] if n["so_le"] is not None else "-",
                "Có" if n["ten"] in tp.PHAN_TU_SPC else ""))

    def _sua_o(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or col not in ("#3", "#4"):
            return
        n = next(x for x in self.ds if x["ten"] == item)
        if n["min"] is None:
            return
        x, y, w, h = self.tree.bbox(item, col)
        e = tk.Entry(self.tree, justify="center")
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, self.tree.set(item, col))
        e.select_range(0, "end")
        e.focus_set()

        def xong(evt=None):
            s = e.get().strip().replace(",", ".")
            e.destroy()
            try:
                v = float(s)
            except ValueError:
                return
            n["min" if col == "#3" else "max"] = v
            n["so_le"] = tp.so_le(n["ten"], n["min"], n["max"])
            self._ve_bang()
        e.bind("<Return>", xong)
        e.bind("<FocusOut>", xong)
        e.bind("<Escape>", lambda evt: e.destroy())

    # =========================================================== CHAY
    def _doc_cfg_chay(self):
        try:
            so_mau = int(self.v_so_mau.get())
            stt = int(self.v_stt.get())
            project = int(self.v_project.get())
            cpk = float(self.v_cpk.get().replace(",", "."))
            tg = datetime.strptime("%s %s" % (self.v_ngay.get().strip(),
                                              self.v_gio.get().strip()), "%d/%m/%Y %H:%M")
        except ValueError:
            raise tp.LoiDuLieu("Kiểm tra lại Số mẫu, STT, Project (số nguyên), Cpk (số), "
                               "Ngày (dd/mm/yyyy) và Giờ (hh:mm).")
        if so_mau < 1:
            raise tp.LoiDuLieu("Số mẫu phải ≥ 1.")
        s = self.v_seed.get().strip()
        return {"file_chuan": self.v_file_chuan.get(), "mac": self.v_mac.get(),
                "nguyen_to": copy.deepcopy(self.ds), "file_mau_spc": self.v_file_spc.get(),
                "file_mau_word": self.v_file_word.get(),
                "file_mau_tong_hop": self.v_file_th.get(),
                "so_mau": so_mau, "stt_dau": stt,
                "seed": int(s) if s.lstrip("-").isdigit() else None,
                "cpk": cpk, "user": self.v_user.get().strip() or "Operator",
                "project": project, "thoi_gian": tg,
                "tieu_de": self.v_tieu_de.get().strip() or TIEU_DE_MAC_DINH,
                "thu_muc_goc": os.path.join(bo_cau_hinh.base_dir(), "output")}

    def _chay(self):
        try:
            cfg = self._doc_cfg_chay()
            for k in ("file_chuan", "file_mau_spc", "file_mau_word", "file_mau_tong_hop"):
                if not cfg[k] or not os.path.exists(cfg[k]):
                    raise tp.LoiDuLieu("Chưa chọn hoặc không tìm thấy file: %s" % (cfg[k] or k))
            tp.kiem_tra_chuan(cfg["nguyen_to"])
        except tp.LoiDuLieu as e:
            messagebox.showwarning("Chưa chạy được", str(e))
            return
        bo_cau_hinh.nho_lan_truoc(KHOA, self.lay_cau_hinh())
        self.dang_dung = False
        self.t0 = time.time()
        self.btn_chay.config(state="disabled")
        self.btn_dung.config(state="normal")
        self.btn_mo.config(state="disabled")
        self.console.clear_log()
        threading.Thread(target=self._nen, args=(cfg,), daemon=True).start()
        self._poll()

    def _dung(self):
        self.dang_dung = True

    def _nen(self, cfg):
        try:
            kq = xt.chay(cfg, log=lambda s, loai="normal": self.hang_doi.put(("log", s, loai)),
                         dung=lambda: self.dang_dung)
            self.hang_doi.put(("xong", kq))
        except Exception as e:
            loi_nv = isinstance(e, tp.LoiDuLieu) or type(e).__name__ in ("LoiMauWord", "LoiMauSPC")
            if loi_nv and self.dang_dung:
                self.hang_doi.put(("dung", None))
            elif loi_nv:
                self.hang_doi.put(("loi", str(e)))
            else:
                self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                m = self.hang_doi.get_nowait()
                if m[0] == "log":
                    self.console.ghi(m[1] + "\n", tag=m[2])
                    continue
                self._ket_thuc()
                if m[0] == "xong":
                    self._bao_xong(m[1])
                elif m[0] == "dung":
                    self.console.ghi("Đã dừng theo yêu cầu.\n", tag="error")
                    messagebox.showinfo("Đã dừng", "Đã dừng theo yêu cầu.\n\nThư mục kết quả "
                                                  "(nếu đã tạo) có thể chưa đủ file.")
                elif m[0] == "loi":
                    self.console.ghi("LỖI: " + m[1] + "\n", tag="error")
                    messagebox.showerror("Không sinh được", m[1])
                else:
                    self.console.ghi("LỖI KHÔNG MONG ĐỢI:\n" + m[1] + "\n", tag="error")
                    messagebox.showerror("Lỗi", "Có lỗi ngoài dự kiến. Chi tiết ở nhật ký bên dưới.")
                return
        except queue.Empty:
            pass
        self.me.after(100, self._poll)

    def _bao_xong(self, ra):
        self.thu_muc_ket_qua = ra["thu_muc"]
        self.btn_mo.config(state="normal")
        if self.khi_xong:
            try:
                self.khi_xong(ra["thu_muc"])
                self.console.ghi("➜ Đã nạp sẵn thư mục vừa sinh sang bước ② Đánh giá SPC.\n",
                                 tag="accent")
            except Exception:
                self.console.ghi("(Không nạp được sang Đánh giá SPC:\n%s)\n"
                                 % traceback.format_exc(), tag="error")
        cb =("\n\n⚠ %s" % "\n⚠ ".join(ra["canh_bao"])) if ra.get("canh_bao") else ""
        if messagebox.askyesno(
                "Xong", "Đã sinh xong %d mẫu trong %.1f giây:\n"
                        "  • Tổng hợp thành phần\n  • %d file Kết quả in\n"
                        "  • %d file đánh giá SPC%s\n\nMở thư mục kết quả bây giờ?"
                        % (len(ra["word"]), time.time() - getattr(self, "t0", time.time()),
                           len(ra["word"]), len(ra["spc"]), cb)):
            self._mo_ket_qua()

    def _ket_thuc(self):
        self.btn_chay.config(state="normal")
        self.btn_dung.config(state="disabled")

    def _mo_ket_qua(self):
        if self.thu_muc_ket_qua and os.path.isdir(self.thu_muc_ket_qua):
            if self.mo_thu_muc_ngoai:
                self.mo_thu_muc_ngoai(self.thu_muc_ket_qua)
            else:
                os.startfile(self.thu_muc_ket_qua)

    # =========================================================== LUU / NAP
    def lay_cau_hinh(self):
        return {"file_chuan": self.v_file_chuan.get(), "mac": self.v_mac.get(),
                "file_spc": self.v_file_spc.get(), "file_word": self.v_file_word.get(),
                "file_th": self.v_file_th.get(), "so_mau": self.v_so_mau.get(),
                "stt": self.v_stt.get(), "seed": self.v_seed.get(), "cpk": self.v_cpk.get(),
                "user": self.v_user.get(), "project": self.v_project.get(),
                "gio": self.v_gio.get(), "tieu_de": self.v_tieu_de.get(),
                "min_max": self._min_max_da_sua()}

    def _min_max_da_sua(self):
        """{mac: {ten: [min, max]}} - CHI nguyen to nguoi dung sua khac file
        chuan. Luu ca bang thi sua file chuan roi mo lai van thay so cu."""
        mac = self.v_mac.get()
        if not self.chuan or mac not in self.chuan:
            return {}
        goc = {n["ten"]: n for n in self.chuan[mac]["nguyen_to"]}
        sua = {n["ten"]: [n["min"], n["max"]] for n in self.ds
               if n["min"] is not None and n["ten"] in goc and
               (n["min"], n["max"]) != (goc[n["ten"]]["min"], goc[n["ten"]]["max"])}
        return {mac: sua} if sua else {}

    def nap_cau_hinh(self, cfg):
        if not cfg:
            return
        if cfg.get("file_chuan") and os.path.exists(cfg["file_chuan"]):
            self.v_file_chuan.set(cfg["file_chuan"])
            if cfg.get("mac"):
                self.v_mac.set(cfg["mac"])
            self._doc_chuan(cfg["file_chuan"])
        for k, bien in (("file_spc", self.v_file_spc), ("file_word", self.v_file_word),
                        ("file_th", self.v_file_th)):
            if cfg.get(k) and os.path.exists(cfg[k]):
                bien.set(cfg[k])
        for k, bien in (("so_mau", self.v_so_mau), ("stt", self.v_stt), ("seed", self.v_seed),
                        ("cpk", self.v_cpk), ("user", self.v_user), ("project", self.v_project),
                        ("gio", self.v_gio), ("tieu_de", self.v_tieu_de)):
            if cfg.get(k) not in (None, ""):
                bien.set(cfg[k])
        sua = (cfg.get("min_max") or {}).get(self.v_mac.get())
        sua = sua if isinstance(sua, dict) else {}      # cau hinh cu (khong theo mac): bo qua
        for n in self.ds:
            g = sua.get(n["ten"])
            if n["min"] is not None and isinstance(g, list) and len(g) == 2 and None not in g:
                n["min"], n["max"] = float(g[0]), float(g[1])
                n["so_le"] = tp.so_le(n["ten"], n["min"], n["max"])
        self._ve_bang()
