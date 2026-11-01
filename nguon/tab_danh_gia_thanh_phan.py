# -*- coding: utf-8 -*-
"""
TAB THANH PHAN HOA - BUOC 2: DANH GIA SPC
=========================================
Nap thu muc ThanhPhan_... (buoc 1 sinh xong tu nap) hoac file
TongHop_ThanhPhan.xlsx, chon day mau, cham 10 nguyen to, "Sua so cho dat"
(ghi de Tong hop + Ket qua in + file SPC trong thu muc) va xuat file SPC
cho tung day (danh_gia_thanh_phan).

Yeu cau, Min/Max lay theo bang o buoc 1 neu cung mac (ke ca o da sua),
khac mac thi lay theo file du lieu chuan.
"""

import copy
import os
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme as ui
import bo_cau_hinh
import spc_luat
import thanh_phan_core as tp
import xuat_thanh_phan as xt
import danh_gia_thanh_phan as dg

KHOA = "danh_gia_thanh_phan"


class TabDanhGiaThanhPhan(object):
    COT = ("day", "ten", "yeu_cau", "xb", "sigma", "lcl", "ucl", "cpk", "ket_luan")
    TEN_COT = ("Dãy (STT)", "Nguyên tố", "Yêu cầu", "X̄", "σ", "LCLI", "UCLI", "Cpk", "Kết luận")
    RONG = (90, 70, 80, 75, 75, 75, 75, 55, 300)

    def __init__(self, me, mo_thu_muc_fn, tab_sinh=None):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.tab_sinh = tab_sinh
        self.hang_doi = queue.Queue()
        self.du_lieu = None
        self._chuan = None
        self.kq = {}
        self._dung_giao_dien()
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
        self._khoi_day(than)
        self._khoi_bieu_mau(than)
        self._khoi_chay(than)

    def _khoi_nguon(self, me):
        c = ui.ModernCard(me, title="Thư mục hoặc file Tổng hợp cần đánh giá", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)
        self.v_nguon = tk.StringVar()
        self.v_mac = tk.StringVar()
        ttk.Label(k, text="Thư mục / file", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(k, textvariable=self.v_nguon).grid(row=0, column=1, sticky="ew", ipady=2)
        ttk.Button(k, text="📂 Chọn thư mục...", style="Outline.TButton", width=16,
                   command=self.chon_thu_muc).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(k, text="📄 Chọn file .xlsx", style="Outline.TButton", width=16,
                   command=self.chon_file).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(k, text="📥 Nạp", style="Outline.TButton", width=8,
                   command=self.nap_lai).grid(row=0, column=4, padx=(6, 0))
        ttk.Label(k, text="Mác thép", style="Surface.TLabel").grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        self.cb_mac = ttk.Combobox(k, textvariable=self.v_mac, state="readonly", width=14)
        self.cb_mac.grid(row=1, column=1, sticky="w", pady=(6, 0), ipady=2)
        self.cb_mac.bind("<<ComboboxSelected>>", lambda e: self.cham(hoi=False))
        self.lbl = ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                             text="Sinh xong ở bước ① thì thư mục vừa sinh tự nạp vào đây. "
                                  "Cũng chọn được file Tổng hợp thành phần phòng thí nghiệm "
                                  "tự điền (dòng tiêu đề có STT ở cột A).")
        self.lbl.grid(row=2, column=0, columnspan=5, sticky="w", pady=(6, 0))

    def _khoi_day(self, me):
        c = ui.ModernCard(me, title="Chọn dãy mẫu & kết quả chấm 10 nguyên tố", badge_num=2)
        c.pack(fill="both", expand=True, padx=12, pady=4)
        k = c.get_body()
        o = tk.Frame(k, bg=ui.COLOR_SURFACE)
        o.pack(fill="x")
        self.v_bd = tk.StringVar(value="1")
        self.v_buoc = tk.StringVar(value="1")
        self.v_sd = tk.StringVar(value="25")
        for i, (nhan, bien, w) in enumerate((("Mẫu bắt đầu", self.v_bd, 22),
                                              ("Bước nhảy", self.v_buoc, 6),
                                              ("Số điểm (≤ 25)", self.v_sd, 6))):
            ttk.Label(o, text=nhan, style="Surface.TLabel").grid(
                row=0, column=i * 2, sticky="w", padx=(0 if i == 0 else 16, 6))
            ttk.Entry(o, textvariable=bien, width=w).grid(row=0, column=i * 2 + 1, sticky="w", ipady=2)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="Mẫu bắt đầu tính theo thứ tự dòng trong Tổng hợp, ghi được nhiều số "
                       "cách nhau dấu phẩy (1, 26, 51 → mỗi dãy một file). Dãy cuối thiếu "
                       "mẫu thì tự ngắn lại.").pack(anchor="w", pady=(4, 6))
        n = tk.Frame(k, bg=ui.COLOR_SURFACE)
        n.pack(fill="x", pady=(0, 6))
        ttk.Button(n, text="🔍 Chấm", style="Outline.TButton",
                   command=self.cham).pack(side="left")
        ttk.Button(n, text="⚡ Quét mọi điểm bắt đầu", style="Outline.TButton",
                   command=self.quet).pack(side="left", padx=8)
        self.btn_sua = ttk.Button(n, text="🛠 Sửa số cho đạt", style="Outline.TButton",
                                  command=self.sua_so)
        self.btn_sua.pack(side="left")

        khung = tk.Frame(k, bg=ui.COLOR_SURFACE)
        khung.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(khung, columns=self.COT, show="headings", height=12)
        for col, ten, w in zip(self.COT, self.TEN_COT, self.RONG):
            self.tree.heading(col, text=ten)
            self.tree.column(col, width=w, minwidth=w,
                             anchor="w" if col == "ket_luan" else "center")
        sb = ttk.Scrollbar(khung, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("hong", foreground=ui.COLOR_DANGER)
        self.tree.bind("<<TreeviewSelect>>", self._khi_chon)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="Chấm đúng công thức file SPC xuất ra: σ = ROUND(MR̄/1,128; d+2), giới "
                       "hạn làm tròn d+1 (d = số chữ số của nguyên tố), 8 mục I-chart + 2 mục "
                       "MR-chart + Cpk. “Sửa số cho đạt” chỉ đổi số của nguyên tố chưa đạt "
                       "trong dãy chưa đạt (trong Min/Max, đúng số chữ số), tính lại Fe, giữ "
                       "Ceq; rồi GHI ĐÈ Tổng hợp, Kết quả in của mẫu bị đổi và các file SPC "
                       "trong thư mục.").pack(anchor="w", pady=(4, 0))

    def _khoi_bieu_mau(self, me):
        c = ui.ModernCard(me, title="Thông tin in trên biểu mẫu SPC", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()
        for i in (1, 3):
            k.columnconfigure(i, weight=1)
        self.v = {}
        doi = [("ngay_danh_gia", "Ngày đánh giá", datetime.now().strftime("%d/%m/%Y")),
               ("du_lieu_danh_gia", "Dữ liệu đánh giá", ""),
               ("tieu_chuan_ap_dung", "Tiêu chuẩn áp dụng", "JIS G 5102"),
               ("ten_san_pham", "Tên sản phẩm", ""),
               ("nguoi_lap", "Người lập", ""),
               ("nguoi_kiem_tra", "Người kiểm tra", "Đoàn Ngọc Thoạt"),
               ("nguoi_phe_duyet", "Phê duyệt", "Phạm Văn Phong"),
               ("cpk", "Cpk tối thiểu", "1.33")]
        for i, (khoa, nhan, md) in enumerate(doi):
            r, cc = divmod(i, 2)
            self.v[khoa] = tk.StringVar(value=md)
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(
                row=r, column=cc * 2, sticky="w", padx=(0, 8), pady=3)
            ttk.Entry(k, textvariable=self.v[khoa]).grid(
                row=r, column=cc * 2 + 1, sticky="ew", padx=(0, 16), ipady=2)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="Dữ liệu đánh giá để trống → tự ghi “Mẫu STT a–b” theo từng dãy. Ô để "
                       "trống khác giữ nguyên chữ của file mẫu.").grid(
            row=4, column=0, columnspan=4, sticky="w", pady=(4, 0))

    def _khoi_chay(self, me):
        d = tk.Frame(me, bg=ui.COLOR_BG, padx=4, pady=8)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   TẠO BÁO CÁO SPC", style="Primary.TButton",
                              command=self.tao_bao_cao)
        self.btn.pack(side="left")
        ttk.Button(d, text="📂 Mở thư mục", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=8)
        self.lbl_td = ttk.Label(d, style="Phu.TLabel", text="")
        self.lbl_td.pack(side="left", padx=10)
        self.log = ui.ModernTerminal(me, title="CHI TIẾT ĐÁNH GIÁ, SỬA SỐ & XUẤT BÁO CÁO", height=10)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))

    # =========================================================== CHUAN
    def _chuan_goc(self):
        """{mac: {...}} - cua buoc 1 neu co, khong thi doc file chuan giao kem."""
        if self.tab_sinh is not None and self.tab_sinh.chuan:
            return self.tab_sinh.chuan
        if self._chuan is None:
            p = os.path.join(bo_cau_hinh.thu_muc_file_mau(), "ThanhPhan", "Dữ liệu 2.xlsx")
            self._chuan = tp.doc_chuan(p) if os.path.exists(p) else {}
        return self._chuan

    def _ds_ceq(self, mac):
        ch = self._chuan_goc()
        if mac not in ch:
            raise tp.LoiDuLieu("Chưa có dữ liệu chuẩn của mác '%s' (nạp file 'Dữ liệu 2.xlsx' "
                               "ở bước ①)." % mac)
        ts = self.tab_sinh
        ds = copy.deepcopy(ts.ds) if ts is not None and ts.v_mac.get() == mac and ts.ds \
            else copy.deepcopy(ch[mac]["nguyen_to"])
        return ds, ch[mac]["ceq"]

    def _file_mau_spc(self, mac):
        ts = self.tab_sinh
        if ts is not None and ts.v_mac.get() == mac and os.path.exists(ts.v_file_spc.get()):
            return ts.v_file_spc.get()
        cho = [os.path.join(bo_cau_hinh.thu_muc_file_mau(), "ThanhPhan")]
        if ts is not None:
            cho = [ts._thu_muc_mau(), os.path.dirname(ts.v_file_chuan.get())] + cho
        for tm in cho:
            p = os.path.join(tm or "", xt.ten_mau_spc(mac))
            if tm and os.path.exists(p):
                return p
        raise tp.LoiDuLieu("Không tìm thấy file mẫu SPC của mác %s:\n%s"
                           % (mac, xt.ten_mau_spc(mac)))

    # =========================================================== NAP
    def chon_thu_muc(self):
        p = filedialog.askdirectory(title="Chọn thư mục ThanhPhan_…")
        if p:
            self._nap(p, hoi=True)

    def chon_file(self):
        p = filedialog.askopenfilename(title="Chọn file Tổng hợp thành phần",
                                       filetypes=[("File Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self._nap(p, hoi=True)

    def nap_lai(self):
        p = self.v_nguon.get().strip()
        if not p:
            messagebox.showinfo("Chưa có đường dẫn", "Chọn thư mục hoặc file Tổng hợp trước.")
            return
        self._nap(p, hoi=True, giu_day=True)

    def dat_san(self, p):
        """Buoc 1 sinh xong goi: nap thu muc vua sinh (khong hien hop thoai)."""
        self._nap(p, hoi=False)

    def _nap(self, p, hoi=False, giu_day=False):
        self.v_nguon.set(p)
        self.tree.delete(*self.tree.get_children())
        self.kq = {}
        try:
            d = dg.doc(p)
        except Exception as e:
            self.du_lieu = None
            msg = str(e) if isinstance(e, tp.LoiDuLieu) else traceback.format_exc()
            self.lbl.configure(text="Không đọc được: %s" % str(e).splitlines()[0])
            self.log.ghi("KHÔNG ĐỌC ĐƯỢC:\n%s\n" % msg, xoa=True, tag="error")
            if hoi:
                messagebox.showerror("Không đọc được", str(e))
            return
        self.du_lieu = d
        mac = list(self._chuan_goc())
        self.cb_mac["values"] = mac
        if d["mac"] in mac:
            self.v_mac.set(d["mac"])
        elif self.v_mac.get() not in mac and mac:
            self.v_mac.set(mac[0])
        if not giu_day:
            self.v_bd.set(", ".join(str(s) for s in dg.mac_dinh_bat_dau(len(d["mau"]))) or "1")
            self.v_buoc.set("1")
            self.v_sd.set("25")
        self.lbl.configure(text="Đã đọc %d mẫu (STT %d–%d) · mác %s · %s"
                           % (len(d["mau"]), d["stt"][0], d["stt"][-1],
                              d["mac"] or "chưa rõ — chọn ở ô Mác thép", d["file"]))
        self.cham(hoi=hoi)

    # =========================================================== THAM SO
    def _tham_so(self):
        d = self.du_lieu
        if not d:
            raise tp.LoiDuLieu("Chưa nạp thư mục / file Tổng hợp.")
        mac = self.v_mac.get()
        ds, ceq = self._ds_ceq(mac)
        tp.kiem_tra_chuan(ds)
        ds = dg.ds_theo_file(ds, d)
        try:
            cpk = spc_luat.kiem_tra_nguong(None, None, self.v["cpk"].get().strip().replace(",", "."))[2]
        except ValueError:
            raise tp.LoiDuLieu("Cpk tối thiểu phải là số lớn hơn 0.")
        try:
            buoc = spc_luat.so_nguyen(self.v_buoc.get().strip().replace(",", "."), "Bước nhảy", 1)
            sd = spc_luat.so_nguyen(self.v_sd.get().strip().replace(",", "."), "Số điểm", 2)
        except ValueError as e:
            raise tp.LoiDuLieu(str(e))
        bd = [x for x in self.v_bd.get().replace(";", " ").replace(",", " ").split()]
        if not bd:
            raise tp.LoiDuLieu("Chưa nhập mẫu bắt đầu.")
        day = []
        for s in bd:
            vt = dg.lay_day(d, s, buoc, sd)
            if vt not in [w for _b, w in day]:
                day.append((vt[0], vt))
        tt = {k: v.get().strip() for k, v in self.v.items() if k not in ("cpk", "ngay_danh_gia")}
        return {"d": d, "mac": mac, "ds": ds, "ceq": ceq, "cpk": cpk, "day": day,
                "buoc": buoc, "so_diem": sd, "ngay": self.v["ngay_danh_gia"].get().strip(),
                "thong_tin": tt, "file_spc": self._file_mau_spc(mac)}

    def _thong_tin_day(self, ts, vt):
        tt = dict(ts["thong_tin"])
        if not tt.get("du_lieu_danh_gia"):
            st = ts["d"]["stt"]
            tt["du_lieu_danh_gia"] = "Mẫu STT %d–%d" % (st[vt[0] - 1], st[vt[-1] - 1])
        return tt

    # =========================================================== CHAM
    def cham(self, hoi=True):
        self.tree.delete(*self.tree.get_children())
        self.kq = {}
        if not self.du_lieu:
            if hoi:
                messagebox.showinfo("Chưa có dữ liệu", "Chọn thư mục ThanhPhan_… hoặc file "
                                                       "Tổng hợp rồi bấm Nạp trước.")
            return
        try:
            ts = self._tham_so()
        except tp.LoiDuLieu as e:
            self.log.ghi("CHƯA CHẤM ĐƯỢC:\n%s\n" % e, xoa=True, tag="error")
            if hoi:
                messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        bo_cau_hinh.nho_lan_truoc(KHOA, self.lay_cau_hinh())
        d, nt = ts["d"], {n["ten"]: n for n in ts["ds"]}
        so_dat = 0
        for bd, vt in ts["day"]:
            r = dg.cham_day(d, ts["ds"], vt, ts["cpk"])
            so_dat += all(x.dat for x in r.values())
            for t in tp.PHAN_TU_SPC:
                x, so = r[t], nt[t]["so_le"]
                self.kq[(bd, t)] = (vt, x)

                def f(v, them):
                    return "-" if v is None else "%.*f" % (so + them, v)
                self.tree.insert("", "end", iid="%d|%s" % (bd, t),
                                 tags=() if x.dat else ("hong",), values=(
                    "%d–%d" % (d["stt"][vt[0] - 1], d["stt"][vt[-1] - 1]), t,
                    nt[t]["yeu_cau"] if nt[t]["yeu_cau"] is not None else "-",
                    f(x.xb, 1), f(x.sigma, 2), f(x.lcl_i, 1), f(x.ucl_i, 1),
                    "-" if x.cpk is None else "%.2f" % x.cpk,
                    "ĐẠT" if x.dat else "CHƯA ĐẠT: " + ", ".join(x.loi)))
        n = len(ts["day"])
        self.log.ghi("Mác %s · %d dãy — %d dãy đạt cả 10 nguyên tố, %d dãy chưa đạt.\n%s"
                     % (ts["mac"], n, so_dat, n - so_dat,
                        "Bấm 🛠 Sửa số cho đạt để sửa, hoặc ▶ TẠO BÁO CÁO SPC để xuất file.\n"
                        if so_dat < n else "Bấm ▶ TẠO BÁO CÁO SPC để xuất file.\n"),
                     xoa=True, tag="success" if so_dat == n else "warning")

    def _khi_chon(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        bd, t = sel[-1].split("|")
        vt, r = self.kq.get((int(bd), t), (None, None))
        if r is None:
            return
        st = self.du_lieu["stt"]
        L = ["%s · dãy STT %d–%d (%d mẫu)" % (t, st[vt[0] - 1], st[vt[-1] - 1], len(vt)),
             "  LSL %s / USL %s" % ("-" if r.lsl is None else r.lsl, "-" if r.usl is None else r.usl)]
        if "DU_LIEU" in r.loi:
            L.append("  CHƯA ĐẠT: có ô trống hoặc không phải số trong dãy.")
        else:
            L.append("  X̄ %.5f   MR̄ %.5f   σ %.6f   UCLI %.5f / LCLI %.5f   UCLMR %.6f"
                     % (r.xb, r.mrb, r.sigma, r.ucl_i, r.lcl_i, r.ucl_mr))
            L.append("  Ca %s   Cp %s   Cpk %s" % tuple("-" if v is None else "%.3f" % v
                                                    for v in (r.ca, r.cp, r.cpk)))
            if r.dat:
                L.append("  KẾT LUẬN: ĐẠT đủ các mục đánh giá")
            else:
                L.append("  KẾT LUẬN: CHƯA ĐẠT — vi phạm:")
                L += ["     [%-4s] %s" % (m, spc_luat.TEN_LUAT.get(m, m)) for m in r.loi]
        self.log.ghi("\n".join(L) + "\n", xoa=True)

    def quet(self):
        try:
            ts = self._tham_so()
            q = dg.quet(ts["d"], ts["ds"], ts["buoc"], ts["so_diem"], ts["cpk"])
        except tp.LoiDuLieu as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        dat = [s for s, o, _l in q if o]
        hong = [(s, l) for s, o, l in q if not o]
        L = ["Quét %d điểm bắt đầu · bước %d · %d điểm" % (len(q), ts["buoc"], ts["so_diem"]),
             "  ĐẠT cả 10 nguyên tố (%d): %s%s" % (len(dat), ", ".join(map(str, dat[:40])),
                                                  " …" if len(dat) > 40 else ""),
             "  CHƯA ĐẠT (%d): %s%s" % (len(hong), "; ".join("%d (%s)" % (s, ",".join(l))
                                                            for s, l in hong[:20]),
                                       " …" if len(hong) > 20 else "")]
        if q:
            L.append("  Tỉ lệ đạt: %.0f%%" % (100.0 * len(dat) / len(q)))
        self.log.ghi("\n".join(L) + "\n", xoa=True)

    # =========================================================== SUA SO
    def sua_so(self):
        try:
            ts = self._tham_so()
        except tp.LoiDuLieu as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        d = ts["d"]
        cac_day = [vt for _b, vt in ts["day"]]
        for _p, vt in dg.file_spc_trong_thu_muc(d, ts["mac"]):
            if vt not in cac_day:
                cac_day.append(vt)
        hong = []
        for vt in cac_day:
            r = dg.cham_day(d, ts["ds"], vt, ts["cpk"])
            loi = [t for t in tp.PHAN_TU_SPC if not r[t].dat]
            if loi:
                hong.append("   STT %d–%d: %s" % (d["stt"][vt[0] - 1], d["stt"][vt[-1] - 1],
                                                  ", ".join(loi)))
        if not hong:
            messagebox.showinfo("Đã đạt sẵn", "Cả %d dãy (gồm dãy đang chọn và dãy của các file "
                                              "SPC trong thư mục) đều đã đạt, không cần sửa."
                                % len(cac_day))
            return
        if not messagebox.askyesno(
                "Sửa số cho đạt",
                "Có %d/%d dãy chưa đạt:\n\n%s\n\nPhần mềm sẽ đổi số của nguyên tố chưa đạt "
                "(trong Min/Max, đúng số chữ số), tính lại Fe, giữ Ceq trong yêu cầu, rồi GHI "
                "ĐÈ:\n  • %s\n  • Kết quả in của các mẫu bị đổi\n  • các file SPC trong thư "
                "mục có mẫu bị đổi\n\nTiếp tục?"
                % (len(hong), len(cac_day), "\n".join(hong[:10]), os.path.basename(d["file"]))):
            return
        ts["cac_day"] = cac_day
        self._bat_nut(False, self.btn_sua, "Đang sửa số…")
        self.log.ghi("Đang tìm số mới cho các nguyên tố chưa đạt…\n", xoa=True)
        self._chay_nen(self._sua_nen, ts)
        self.me.after(150, self._kiem_tra)

    def _sua_nen(self, ts):
        try:
            d = copy.deepcopy(ts["d"])
            s = dg.sua(d, ts["ds"], ts["ceq"], ts["cac_day"], ts["cpk"])
            if s["con_loi"]:
                raise tp.LoiDuLieu("Chưa tìm được bộ số đạt — CHƯA ghi file nào:\n  %s\n\nHãy "
                                   "nới Min/Max ở bước ① hoặc giảm Cpk tối thiểu."
                                   % "\n  ".join(s["con_loi"][:10]))
            ra = dg.ghi_lai(d, ts["ds"], s["doi"], ts["file_spc"], ts["mac"], ts["cpk"],
                            ts["ngay"], lambda vt: self._thong_tin_day(ts, vt))
            self.hang_doi.put(("xong_sua", (d, s, ra, ts["ds"])))
        except tp.LoiDuLieu as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi", "Không ghi được file — thường do file đang mở trong "
                                      "Excel/Word. Đóng lại rồi bấm Sửa lần nữa.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    # =========================================================== XUAT
    def tao_bao_cao(self):
        try:
            ts = self._tham_so()
        except tp.LoiDuLieu as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        bo_cau_hinh.nho_lan_truoc(KHOA, self.lay_cau_hinh())
        self._bat_nut(False, self.btn, "Đang tạo…")
        self.log.ghi("Đang tạo %d file báo cáo SPC…\n" % len(ts["day"]), xoa=True)
        self._chay_nen(self._xuat_nen, ts)
        self.me.after(150, self._kiem_tra)

    def _xuat_nen(self, ts):
        try:
            d, ra = ts["d"], []
            for i, (bd, vt) in enumerate(ts["day"]):
                p = os.path.join(d["thu_muc"], dg.ten_file_day(ts["mac"], bd, ts["buoc"], len(vt)))
                _p, spc = dg.xuat_spc_day(d, ts["ds"], vt, ts["mac"], ts["file_spc"], ts["cpk"],
                                          ts["ngay"], self._thong_tin_day(ts, vt), p)
                ra.append((p, vt, spc))
                self.hang_doi.put(("td", (i + 1, len(ts["day"]))))
            self.hang_doi.put(("xong_xuat", ra))
        except tp.LoiDuLieu as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi", "Không ghi được file — thường do file đang mở trong "
                                      "Excel. Đóng lại rồi thử lại.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    # =========================================================== HANG DOI
    def _kiem_tra(self):
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "td":
                    self.lbl_td.configure(text="%d / %d" % dl)
                    continue
                self._bat_nut(True)
                if loai == "xong_sua":
                    self._bao_sua(*dl)
                elif loai == "xong_xuat":
                    self._bao_xuat(dl)
                elif loai == "loi":
                    self.log.ghi("KHÔNG XONG:\n%s\n" % dl, xoa=True, tag="error")
                    messagebox.showerror("Không xong", dl)
                else:
                    self.log.ghi("LỖI KHÔNG MONG ĐỢI:\n%s\n" % dl, xoa=True, tag="error")
                    messagebox.showerror("Lỗi", "Có lỗi ngoài dự kiến. Chi tiết ở khung bên dưới.")
                return
        except queue.Empty:
            pass
        self.me.after(150, self._kiem_tra)

    def _bao_sua(self, d, s, ra, ds):
        self._nap(self.v_nguon.get(), giu_day=True)        # doc lai + cham lai
        so_mau = len({v for v, _t in s["doi"]})
        L = ["ĐÃ SỬA %d ô số liệu ở %d mẫu:" % (len(s["doi"]), so_mau), dg.tom_tat_doi(d, s["doi"], ds),
             "", "Đã ghi đè:", "  " + os.path.basename(ra["tong_hop"])]
        L += ["  KetQuaIn/" + os.path.basename(p) for p in ra["word"][:8]]
        if len(ra["word"]) > 8:
            L.append("  … (%d file Kết quả in)" % len(ra["word"]))
        L += ["  " + os.path.basename(p) for p in ra["spc"]]
        L += ["⚠ " + c for c in ra["canh_bao"]]
        self.log.ghi("\n".join(L) + "\n", tag="success")
        messagebox.showinfo("Đã sửa xong", "Đã đổi %d ô số liệu ở %d mẫu, mọi dãy đều đạt.\n\n"
                                           "Đã ghi đè Tổng hợp, %d file Kết quả in, %d file SPC."
                            % (len(s["doi"]), so_mau, len(ra["word"]), len(ra["spc"])))

    def _bao_xuat(self, ra):
        st = self.du_lieu["stt"]
        hong = [(p, vt, [t for t in tp.PHAN_TU_SPC if not spc[t].dat]) for p, vt, spc in ra]
        hong = [h for h in hong if h[2]]
        L = ["ĐÃ TẠO %d FILE BÁO CÁO SPC — %d dãy đạt, %d dãy chưa đạt" % (len(ra), len(ra) - len(hong), len(hong))]
        for p, vt, spc in ra:
            loi = ["%s(%s)" % (t, ",".join(spc[t].loi)) for t in tp.PHAN_TU_SPC if not spc[t].dat]
            L.append("  %s  ·  STT %d–%d  ·  %s" % (os.path.basename(p), st[vt[0] - 1], st[vt[-1] - 1],
                                                 "ĐẠT" if not loi else "CHƯA ĐẠT: " + " ".join(loi)))
        L.append("Thư mục: %s" % self.du_lieu["thu_muc"])
        self.log.ghi("\n".join(L) + "\n", xoa=True, tag="warning" if hong else "success")
        if hong:
            messagebox.showwarning("Có dãy chưa đạt", "Đã tạo %d file, trong đó %d dãy CHƯA ĐẠT "
                                                     "(file tick 'Không đạt').\n\nBấm 🛠 Sửa số cho "
                                                     "đạt rồi tạo lại." % (len(ra), len(hong)))
        elif messagebox.askyesno("Xong", "Đã tạo %d file báo cáo SPC, cả %d dãy đều ĐẠT.\n\n"
                                         "Mở thư mục bây giờ?" % (len(ra), len(ra))):
            self.mo_thu_muc()

    def _bat_nut(self, bat, nut=None, chu=None):
        """Bat/tat hai nut chay; nut dang chay (btn | btn_sua) hien chu."""
        self.btn.configure(state="normal" if bat else "disabled", text="▶   TẠO BÁO CÁO SPC")
        self.btn_sua.configure(state="normal" if bat else "disabled", text="🛠 Sửa số cho đạt")
        if nut is not None and chu:
            nut.configure(text=chu)
        if bat:
            self.lbl_td.configure(text="")

    def mo_thu_muc(self):
        tm = (self.du_lieu or {}).get("thu_muc")
        if tm and os.path.isdir(tm) and self.mo_thu_muc_ngoai:
            self.mo_thu_muc_ngoai(tm)

    # =========================================================== CAU HINH
    def lay_cau_hinh(self):
        return {k: v.get() for k, v in self.v.items()
                if k not in ("ngay_danh_gia", "du_lieu_danh_gia")}

    def _nap_lan_truoc(self):
        d = bo_cau_hinh.doc_lan_truoc(KHOA) or {}
        for k, v in d.items():
            if k in self.v and isinstance(v, str) and v.strip():
                self.v[k].set(v)
