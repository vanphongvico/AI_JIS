# -*- coding: utf-8 -*-
"""
TAB 4 - TAO BAO CAO DANH GIA DO ON DINH SPC TU THU MUC KET QUA
===============================================================
Chon mot thu muc da sinh (Charpy hoac keo), chon cot, mau bat dau va buoc
nhay, phan mem lay ra 25 mau tuong ung roi xuat file theo bieu mau
TPL_SPC.xlsx - ghi thang vao chinh thu muc do.

Ket qua cham duoc bao trung thuc: dat thi ghi dat, khong dat thi liet ke
ro tung muc vi pham. Nut "Quet diem bat dau" cho thay ngay day nao dat va
day nao khong, khoi phai mo tung file.
"""

import os
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import doc_thu_muc as dtm
import tao_spc
import spc_luat
import bo_cau_hinh
import ui_theme as ui


TEN_LOAI = {"charpy": "thử va đập Charpy", "keo": "thử kéo"}


class TabBaoCaoSPC(object):
    """Phan danh gia SPC - nam thanh TAB CON trong tab Va dap va tab Thu keo.

    khoa : khoa luu bo cau hinh / lan truoc rieng cho tung noi dat tab
    loai : "charpy" | "keo" | None - chi nhan dung loai du lieu cua tab cha
    """

    def __init__(self, me, mo_thu_muc_fn, khoa="spc", ten="Đánh giá SPC",
                 loai=None):
        self.me = me
        self.mo_thu_muc_ngoai = mo_thu_muc_fn
        self.khoa, self.ten, self.loai = khoa, ten, loai
        self.goi_y = {
            "charpy": "Chọn thư mục Charpy_… trong output. Sinh xong ở bước ① "
                      "thì thư mục vừa sinh tự điền vào đây.",
            "keo": "Chọn thư mục Keo_Sinh_… (Bước 1) hoặc Keo_… (Bước 2). "
                   "Sinh xong ở Bước 1 thì thư mục vừa sinh tự điền vào đây.",
        }.get(loai, "Chọn thư mục Charpy_… , Keo_Sinh_… hoặc Keo_… mà phần "
                    "mềm đã sinh ra.")
        self.hang_doi = queue.Queue()
        self.du_lieu = None
        self.v_cot = {}
        self.v_lsl = {}
        # Mac vat lieu de danh gia thu keo (SCW480 / SCW410) - quyet dinh LSL
        self.v_mac = tk.StringVar(value=dtm.MAC_KEO_MAC_DINH)
        self._dung_giao_dien()
        bo_cau_hinh.dang_ky_thong_tin(self.khoa, self.ten, {
            "nguoi_lap": self.v["nguoi_lap"],
            "nguoi_kiem_tra": self.v["nguoi_kiem_tra"],
            "nguoi_phe_duyet": self.v["nguoi_phe_duyet"],
            "vat_lieu": self.v["vat_lieu_spc"],
            "tieu_chuan_vat_lieu": self.v["tieu_chuan_ap_dung"],
            "san_pham": self.v["ten_san_pham"]})
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

        self.thanh = bo_cau_hinh.ThanhCauHinh(
            than, self.khoa, self.ten, self.lay_cau_hinh, self.dat_cau_hinh)
        self._khoi_nguon(than)
        self._khoi_chon(than)
        self._khoi_bieu_mau(than)
        self._khoi_chay(than)

    # ------------------------------------------------------- 1. nguon
    def _khoi_nguon(self, me):
        c = ui.ModernCard(me, title="Thư mục hoặc file dữ liệu nguồn", badge_num=1)
        c.pack(fill="x", padx=12, pady=(8, 4))
        k = c.get_body()
        k.columnconfigure(1, weight=1)

        self.v_tm = tk.StringVar(value="")
        ttk.Label(k, text="Thư mục kết quả", style="Surface.TLabel").grid(row=0, column=0, sticky="w",
                                                  padx=(0, 8))
        ttk.Entry(k, textvariable=self.v_tm).grid(row=0, column=1, sticky="ew",
                                                  ipady=2)
        ttk.Button(k, text="📂 Chọn thư mục...", style="Outline.TButton", width=16,
                   command=self.chon_thu_muc).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(k, text="📄 Chọn file .xlsx", style="Outline.TButton", width=16,
                   command=self.chon_file).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(k, text="📥 Nạp", style="Outline.TButton", width=8,
                   command=self.nap_lai).grid(row=0, column=4, padx=(6, 0))
        self.lbl_nguon = ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920,
                                   justify="left",
                                   text=self.goi_y)
        self.lbl_nguon.grid(row=1, column=0, columnspan=5, sticky="w",
                            pady=(6, 0))

    # -------------------------------------------------- 2. chon day 25 mau
    def _khoi_chon(self, me):
        c = ui.ModernCard(me, title="Chọn dãy mẫu & Kiểm tra độ ổn định", badge_num=2)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        # Thu keo: chon mac vat lieu TRUOC, LSL cac cot dien theo mac do
        if self.loai == "keo":
            km = tk.Frame(k, bg=ui.COLOR_SURFACE)
            km.grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 8))
            ttk.Label(km, text="Đánh giá theo mác vật liệu",
                      style="Surface.TLabel").pack(side="left", padx=(0, 8))
            cb = ttk.Combobox(km, textvariable=self.v_mac, state="readonly",
                              width=10, values=sorted(dtm.MAC_KEO, reverse=True))
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", self._ap_mac)
            ttk.Label(km, style="SurfacePhu.TLabel",
                      text="   File xuất theo biểu mẫu của mác đã chọn   ·   "
                           + "   ·   ".join(
                          "%s: ReH/ReL ≥ %g, Rm ≥ %g, A ≥ %g"
                          % (m, n["ReL"], n["Rm"], n["A"])
                          for m, n in sorted(dtm.MAC_KEO.items(),
                                             reverse=True))).pack(side="left")

        self.khung_cot = tk.Frame(k, bg=ui.COLOR_SURFACE)
        self.khung_cot.grid(row=1, column=0, columnspan=8, sticky="w")
        ttk.Label(self.khung_cot, style="SurfacePhu.TLabel",
                  text="(chọn thư mục trước)").pack(anchor="w")

        self.v_bd = tk.StringVar(value="1")
        self.v_buoc = tk.StringVar(value="1")
        self.v_sd = tk.StringVar(value="25")
        o = [("Mẫu bắt đầu", self.v_bd, 22), ("Bước nhảy", self.v_buoc, 8),
             ("Số điểm", self.v_sd, 8)]
        for i, (nhan, var, w) in enumerate(o):
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(row=2, column=i * 2, sticky="w",
                                         padx=(0, 6), pady=(10, 0))
            ttk.Entry(k, textvariable=var, width=w).grid(
                row=2, column=i * 2 + 1, sticky="w", padx=(0, 20),
                pady=(10, 0), ipady=2)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="Với Bước nhảy = 1 và nhập 1 mẫu bắt đầu (VD: 1), phần "
                       "mềm tự chia thành các khoảng nối tiếp (1–25, 26–50, "
                       "51–75, 76–100...) để xuất đủ các file đánh giá hết dữ "
                       "liệu. Cũng có thể nhập nhiều số cách nhau dấu phẩy "
                       "(ví dụ 1, 26, 51).").grid(
            row=3, column=0, columnspan=8, sticky="w", pady=(6, 0))

        d = tk.Frame(k, bg=ui.COLOR_SURFACE)
        d.grid(row=4, column=0, columnspan=8, sticky="w", pady=(8, 0))
        ttk.Button(d, text="🔍 Xem trước kết quả chấm", style="Outline.TButton",
                   command=self.xem_truoc).pack(side="left")
        ttk.Button(d, text="⚡ Quét mọi điểm bắt đầu", style="Outline.TButton",
                   command=self.quet).pack(side="left", padx=8)
        self.btn_sua = ttk.Button(d, text="🛠 Sửa dữ liệu cho đạt", style="Outline.TButton",
                                  command=self.sua_du_lieu)
        self.btn_sua.pack(side="left")
        self.dung_sua = threading.Event()
        self.btn_dung_sua = ttk.Button(d, text="Dừng sửa", state="disabled", command=self.dung_sua.set)
        self.btn_dung_sua.pack(side="left", padx=6)
        ttk.Label(k, style="SurfacePhu.TLabel", wraplength=920, justify="left",
                  text="“Sửa dữ liệu cho đạt” sắp lại THỨ TỰ các mẫu trong "
                       "thư mục cho tới khi những dãy đang chọn đều đạt — "
                       "giữ nguyên giá trị và mã mẫu; MR, sigma, Cpk có thể "
                       "thay đổi theo thứ tự. Bản gốc được chép sang thư mục con "
                       "_truoc_khi_sua trước khi ghi đè.").grid(
            row=5, column=0, columnspan=8, sticky="w", pady=(6, 0))

    def _ap_mac(self, _e=None):
        """Dien LSL cac cot theo mac vat lieu dang chon (chi thu keo).

        O 'Vat lieu' tren bieu mau doi theo khi nguoi dung vua chon mac, hoac
        khi o do dang de trong / dang ghi dung ten mot mac - con neu nguoi
        dung da go chu rieng thi giu nguyen."""
        mac = self.v_mac.get()
        ng = dtm.MAC_KEO.get(mac)
        if self.loai != "keo" or not ng:
            return
        for k, var in self.v_lsl.items():
            if k in ng:
                var.set("%g" % ng[k])
        vl = getattr(self, "v", {}).get("vat_lieu_spc")
        if vl is not None and (_e is not None or not vl.get().strip()
                               or vl.get().strip() in dtm.MAC_KEO):
            vl.set(mac)

    # ------------------------------------------------- 3. thong tin bieu mau
    def _khoi_bieu_mau(self, me):
        c = ui.ModernCard(me, title="Thông tin in trên biểu mẫu SPC", badge_num=3)
        c.pack(fill="x", padx=12, pady=4)
        k = c.get_body()

        for col_idx in (1, 3):
            k.columnconfigure(col_idx, weight=1)
        self.v_cpk = tk.StringVar(value="1.33")
        self.v = {}
        doi = [("ngay_danh_gia", "Ngày đánh giá",
                datetime.now().strftime("%d/%m/%Y")),
               ("du_lieu_danh_gia", "Dữ liệu đánh giá", ""),
               ("vat_lieu_spc", "Vật liệu", "SCW480"),
               ("tieu_chuan_ap_dung", "Tiêu chuẩn áp dụng", "JIS G5102 - 1991"),
               ("ten_san_pham", "Tên sản phẩm", "ISHIZUE/礎"),
               ("nguoi_lap", "Người lập", "Vũ Vương Việt"),
               ("nguoi_kiem_tra", "Người kiểm tra", "Đoàn Ngọc Thoạt"),
               ("nguoi_phe_duyet", "Phê duyệt", "Phạm Văn Phong")]
        for i, (khoa, nhan, md) in enumerate(doi):
            r, c_idx = i // 2, (i % 2) * 2
            self.v[khoa] = tk.StringVar(value=md)
            ttk.Label(k, text=nhan, style="Surface.TLabel").grid(row=r, column=c_idx, sticky="w",
                                         padx=(0, 8), pady=3)
            ttk.Entry(k, textvariable=self.v[khoa]).grid(
                row=r, column=c_idx + 1, sticky="ew", padx=(0, 16), ipady=2)
        ttk.Label(k, text="Cpk tối thiểu", style="Surface.TLabel").grid(row=4, column=0, sticky="w",
                                                padx=(0, 8), pady=(6, 0))
        ttk.Entry(k, textvariable=self.v_cpk, width=10).grid(
            row=4, column=1, sticky="w", pady=(6, 0), ipady=2)

    # ---------------------------------------------------------- 4. nut chay
    def _khoi_chay(self, me):
        d = tk.Frame(me, bg=ui.COLOR_BG, padx=4, pady=8)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   TẠO BÁO CÁO SPC",
                              style="Primary.TButton", command=self.chay)
        self.btn.pack(side="left")
        ttk.Button(d, text="📂 Mở thư mục", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=8)
        self.lbl_td = ttk.Label(d, style="Phu.TLabel", text="")
        self.lbl_td.pack(side="left", padx=10)

        # ---- Modern Terminal Log ----
        self.log = ui.ModernTerminal(me, title="NHẬT KÝ TẠO BÁO CÁO SPC & QUÉT ĐỘ ỔN ĐỊNH", height=12)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self._ghi("Chọn thư mục kết quả rồi chọn cột, mẫu bắt đầu và bước nhảy.\n"
                  "File báo cáo SPC ghi thẳng vào chính thư mục đó.\n")

    # ============================================================ TIEN ICH
    def _ghi(self, s, xoa=False):
        self.log.ghi(s, xoa=xoa)

    def chon_thu_muc(self):
        p = filedialog.askdirectory(title="Chọn thư mục kết quả")
        if p:
            self.v_tm.set(p)
            self._nap(p)

    def chon_file(self):
        p = filedialog.askopenfilename(
            title="Chọn file bảng tổng hợp",
            filetypes=[("File Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self.v_tm.set(p)
            self._nap(p)

    def nap_lai(self):
        p = self.v_tm.get().strip()
        if not p:
            messagebox.showinfo("Chưa có đường dẫn",
                                "Chọn thư mục hoặc file kết quả trước.")
            return
        self._nap(p)

    def dat_san(self, p, nap=True):
        """Tab cha goi sau khi sinh xong: dien san thu muc vua sinh.
        nap=False (thu muc nhieu file Word, doc lau) thi chi dien duong dan,
        nguoi dung bam Nap khi can."""
        self.v_tm.set(p)
        if nap:
            self._nap(p)
        else:
            self.du_lieu = None
            self.lbl_nguon.configure(
                text="Đã điền sẵn thư mục vừa sinh. Bấm  📥 Nạp  để đọc dữ liệu.")

    def _nap(self, p):
        self._ghi("Đang đọc %s ...\n" % p, xoa=True)
        self.me.update_idletasks()
        try:
            d = dtm.doc(p)
        except dtm.LoiDoc as e:
            self.du_lieu = None
            self.lbl_nguon.configure(text="Không đọc được: %s"
                                          % str(e).splitlines()[0])
            self._ghi("KHÔNG ĐỌC ĐƯỢC:\n\n%s\n" % e, xoa=True)
            messagebox.showerror("Không đọc được", str(e))
            return
        except Exception:
            self.du_lieu = None
            self._ghi("LỖI KHÔNG MONG ĐỢI:\n\n" + traceback.format_exc(),
                      xoa=True)
            return

        if self.loai and d.get("loai") != self.loai:
            self.du_lieu = None
            msg = ("Đây là dữ liệu %s, không phải %s.\nHãy đánh giá dữ liệu này "
                   "ở tab tương ứng." % (TEN_LOAI.get(d.get("loai"), d.get("loai")),
                                        TEN_LOAI.get(self.loai, self.loai)))
            self.lbl_nguon.configure(text=msg.replace("\n", " "))
            self._ghi(msg + "\n", xoa=True)
            messagebox.showwarning("Sai loại dữ liệu", msg)
            return
        self.du_lieu = d
        cot = dtm.danh_sach_cot(d)
        canh = ""
        if d.get("cot_bo_qua"):
            canh = "\nCột bỏ qua: " + "; ".join(
                "%s (%s)" % (k, m) for k, m in d["cot_bo_qua"].items())
        self.lbl_nguon.configure(
            text="Đã đọc %s mẫu · loại %s · nguồn %s\nCác cột chấm được: %s%s"
                 % ("{:,}".format(d["so_mau"]),
                    "thử va đập Charpy" if d["loai"] == "charpy" else "thử kéo",
                    os.path.basename(d["nguon"]),
                    ", ".join(k for k, _u, _n in cot), canh))
        self._dung_o_cot(cot, d)
        for khoa in ("vat_lieu_spc", "tieu_chuan_ap_dung", "ten_san_pham"):
            if d["thong_tin"].get(khoa):
                self.v[khoa].set(d["thong_tin"][khoa])
        if not self.v["du_lieu_danh_gia"].get():
            self.v["du_lieu_danh_gia"].set(os.path.basename(
                p.rstrip(os.sep)))
        self._ghi("Đã đọc %s mẫu.  Nguồn: %s\n\n"
                  "Bấm 'Xem trước kết quả chấm' để soi dãy đang chọn, hoặc "
                  "'Quét mọi điểm bắt đầu' để xem điểm nào đạt.\n"
                  % ("{:,}".format(d["so_mau"]), d["nguon"]), xoa=True)

    def _dung_o_cot(self, cot, d):
        """Dung lai danh sach cot va o nguong moi khi doc thu muc moi."""
        for w in self.khung_cot.winfo_children():
            w.destroy()
        self.v_cot, self.v_lsl = {}, {}
        ttk.Label(self.khung_cot, text="Cột đánh giá  ", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w")
        ttk.Label(self.khung_cot, text="Giới hạn dưới LSL",
                  style="SurfacePhu.TLabel").grid(row=0, column=2, sticky="w",
                                           padx=(24, 0))
        # Thu keo: tick san ba cot cua bieu mau (gioi han chay, Rm, A) - dung
        # nhung cot se duoc cham khi xuat. Loai khac: tick cot dau tien.
        mac_dinh = {cot[0][0]} if cot else set()
        if self.loai == "keo":
            try:
                mac_dinh = set(tao_spc.cot_bieu_mau_keo(d).values())
            except dtm.LoiDoc:
                pass
        for i, (k, u, _nh) in enumerate(cot):
            self.v_cot[k] = tk.BooleanVar(value=(k in mac_dinh))
            ttk.Checkbutton(self.khung_cot, text="%s (%s)" % (k, u),
                            variable=self.v_cot[k]).grid(
                row=i + 1, column=0, columnspan=2, sticky="w", pady=1)
            ng = d["nguong"].get(k)
            self.v_lsl[k] = tk.StringVar(value="" if ng is None else "%g" % ng)
            ttk.Entry(self.khung_cot, textvariable=self.v_lsl[k],
                      width=10).grid(row=i + 1, column=2, sticky="w",
                                     padx=(24, 0), ipady=1)
        gc = ("Để trống LSL = cột không có ngưỡng nghiệm thu, vẫn chấm ổn "
              "định nhưng không chấm Cpk.")
        ttk.Label(self.khung_cot, style="SurfacePhu.TLabel", wraplength=900,
                  justify="left", text=gc).grid(
            row=len(cot) + 1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self._ap_mac()

    # ------------------------------------------------------- doc tham so
    def _thu(self):
        if not self.du_lieu:
            raise dtm.LoiDoc("Chưa chọn thư mục dữ liệu.")
        cot = [k for k, v in self.v_cot.items() if v.get()]

        def so(v, ten):
            v = str(v).strip().replace(",", ".")
            try:
                return spc_luat.so_nguyen(v, ten)
            except ValueError:
                raise dtm.LoiDoc("Ô '%s' phải là số. Đang là: %s" % (ten, v))

        bd = []
        for x in str(self.v_bd.get()).replace(";", " ").replace(",", " ").split():
            bd.append(so(x, "Mẫu bắt đầu"))
        if not bd:
            raise dtm.LoiDoc("Chưa nhập mẫu bắt đầu.")
        buoc = so(self.v_buoc.get(), "Bước nhảy")
        sd = so(self.v_sd.get(), "Số điểm")
        if not (2 <= sd <= 25):
            raise dtm.LoiDoc(
                "Số điểm phải từ 2 đến 25 — biểu mẫu SPC chỉ chứa được 25 "
                "điểm. Đang khai %d." % sd)
        # Thu keo du 25 diem: bieu mau luon cham du ba chi tieu, nen Xem
        # truoc / Quet / Sua / Tao bao cao deu phai xet dung ba cot do -
        # truoc day chi xet cot da tick nen Sua bao dat ma Xuat van truot.
        if self._dung_bieu_mau_keo(sd):
            bm = list(self._cot_bm().values())
            cot = bm + [k for k in cot if k not in bm]
        if not cot:
            raise dtm.LoiDoc("Chưa tick cột nào để đánh giá.")
        # Neu chi nhap 1 mau bat dau voi buoc nhay 1, tu dong chia thanh cac
        # khoang lien tiep (vd 100 mau, bd=1, buoc=1, sd=25 -> 1, 26, 51, 76)
        if len(bd) == 1 and buoc == 1:
            s0 = bd[0]
            het = self.du_lieu["so_mau"] - (sd - 1) * buoc
            if het >= s0:
                bd = list(range(s0, het + 1, sd))
        try:
            cpk = spc_luat.kiem_tra_nguong(None, None, str(self.v_cpk.get()).replace(",", "."))[2]
        except ValueError:
            raise dtm.LoiDoc("Ô 'Cpk tối thiểu' phải là số.")
        ng = {}
        for k in self.v_lsl:
            s = str(self.v_lsl[k].get()).strip().replace(",", ".")
            try:
                ng[k] = spc_luat.so_huu_han(s, "LSL " + k) if s else None
            except ValueError as e:
                raise dtm.LoiDoc(str(e))
        cfg = dict((k, v.get().strip()) for k, v in self.v.items()
                   if v.get().strip())
        # thu muc dich: chinh thu muc nguoi dung da chon
        p = self.v_tm.get().strip()
        tm = p if os.path.isdir(p) else os.path.dirname(p)
        bo_cau_hinh.nho_lan_truoc(self.khoa, self.lay_cau_hinh())
        return cot, bd, buoc, sd, cpk, ng, cfg, tm

    # ================================================================ LENH
    def xem_truoc(self):
        try:
            cot, bd, buoc, sd, cpk, ng, _cfg, _tm = self._thu()
        except dtm.LoiDoc as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        d = self.du_lieu
        L = []
        for k in cot:
            for s in bd:
                try:
                    idx, gt, r = tao_spc.cham_day(d, k, s, buoc, sd,
                                                  ng.get(k), None, cpk)
                except dtm.LoiDoc as e:
                    L.append("Cột %s, bắt đầu %d: %s\n" % (k, s, e))
                    continue
                L.append("Cột %s · bắt đầu %d · bước %d · %d điểm%s"
                         % (k, s, buoc, sd,
                            "" if ng.get(k) is None else " · %sLSL %g" % (
                                (self.v_mac.get() + " · ")
                                if self.loai == "keo" else "", ng[k])))
                L.append("  Mẫu   : %s ... %s"
                         % (", ".join(str(x) for x in idx[:6]), idx[-1]))
                L.append("  Mã mẫu: %s ..."
                         % ", ".join(d["ma"][i - 1] for i in idx[:5]))
                L.append("  Giá trị: %s ..."
                         % ", ".join("%g" % x for x in gt[:10]))
                L.append("  X-bar %.4f   MR-bar %.4f   sigma %.4f%s"
                         % (r.xb, r.mrb, r.sigma,
                            ("   Cpk %.3f" % r.cpk)
                            if r.cpk is not None else ""))
                L.append("  UCLI %.4f / LCLI %.4f   UCLMR %.4f"
                         % (r.ucl_i, r.lcl_i, r.ucl_mr))
                if r.dat:
                    L.append("  KẾT LUẬN: ĐẠT tất cả các mục đánh giá")
                else:
                    L.append("  KẾT LUẬN: CHƯA ĐẠT — vi phạm %d mục:"
                             % len(r.loi))
                    for m in r.loi:
                        L.append("     [%-3s] %s"
                                 % (m, spc_luat.TEN_LUAT.get(m, m)))
                L.append("")
        self._ghi("\n".join(L), xoa=True)

    def quet(self):
        try:
            cot, _bd, buoc, sd, cpk, ng, _cfg, _tm = self._thu()
        except dtm.LoiDoc as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        d = self.du_lieu
        L = []
        for k in cot:
            try:
                q = tao_spc.quet_diem_bat_dau(d, k, buoc, sd, ng.get(k), None,
                                              cpk)
            except dtm.LoiDoc as e:
                L.append("Cột %s: %s\n" % (k, e))
                continue
            dat = [s for s, o, _l in q if o]
            hong = [s for s, o, _l in q if not o]
            L.append("Cột %s · bước nhảy %d · %d điểm — quét %d điểm bắt đầu"
                     % (k, buoc, sd, len(q)))
            L.append("  ĐẠT  (%d): %s%s"
                     % (len(dat), ", ".join(str(x) for x in dat[:30]),
                        " ..." if len(dat) > 30 else ""))
            L.append("  HỎNG (%d): %s%s"
                     % (len(hong), ", ".join(str(x) for x in hong[:30]),
                        " ..." if len(hong) > 30 else ""))
            if q:
                L.append("  Tỉ lệ đạt: %.0f%%" % (100.0 * len(dat) / len(q)))
            L.append("")
        self._ghi("\n".join(L), xoa=True)

    # ------------------------------------------------- sua du lieu cho dat
    def sua_du_lieu(self):
        import sua_du_lieu as sdl
        try:
            cot, bd, buoc, sd, cpk, ng, _cfg, tm = self._thu()
        except dtm.LoiDoc as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return

        # chi sua khi that su co day chua dat
        chua = self._day_chua_dat(cot, bd, buoc, sd, cpk, ng)
        if chua is None:
            return
        if not chua:
            messagebox.showinfo(
                "Đã đạt sẵn",
                "Cả %d dãy đang chọn đều đã đạt, không cần sửa gì."
                % (len(cot) * len(bd)))
            return

        mo_ta = "\n".join("   cột %s, bắt đầu %d — vi phạm %s"
                          % (k, s, ", ".join(l)) for k, s, l in chua[:8])
        if not messagebox.askyesno(
                "Sửa dữ liệu trong thư mục",
                "Có %d/%d dãy chưa đạt:\n\n%s\n\n"
                "Phần mềm sẽ SẮP LẠI THỨ TỰ các mẫu trong thư mục cho tới khi "
                "mọi dãy đang chọn đều đạt mới áp dụng. PDF giữ nguyên các "
                "trang gốc, sắp theo thứ tự mới.\n\nGiữ nguyên giá trị và mã "
                "mẫu; MR, sigma, Cpk có thể đổi. %s\nBản gốc được chép sang "
                "thư mục con _truoc_khi_sua.\n\nTiếp tục?"
                % (len(chua), len(cot) * len(bd), mo_ta, self._loi_nhan_word())):
            return
        self._bat_dau_sua(cot, bd, buoc, sd, cpk, ng, tm)

    def _loi_nhan_word(self):
        return ("Báo cáo SPC cũ được lưu vào bản sao và cần xuất lại; báo "
                "cáo Word từng mẫu giữ nguyên (mã và giá trị không đổi).")

    def _day_chua_dat(self, cot, bd, buoc, sd, cpk, ng):
        """Cham lai moi (cot x diem bat dau). Tra ve list (cot, bat_dau,
        ma_loi) cua cac day chua dat, hoac None neu khong lay duoc day."""
        chua = []
        for k in cot:
            for s in bd:
                try:
                    _i, _g, r = tao_spc.cham_day(self.du_lieu, k, s, buoc, sd,
                                                 ng.get(k), None, cpk)
                except dtm.LoiDoc as e:
                    messagebox.showwarning("Không lấy được dãy", str(e))
                    return None
                if not r.dat:
                    chua.append((k, s, r.loi))
        return chua

    def _bat_dau_sua(self, cot, bd, buoc, sd, cpk, ng, tm, xuat_sau=False):
        """Chay sap lai thu tu o nen. xuat_sau=True: sua dat thi tu xuat
        luon bao cao SPC (nut Tao bao cao goi)."""
        import sua_du_lieu as sdl
        lua_chon = [(k, s, buoc) for k in cot for s in bd]
        self._xuat_sau = xuat_sau
        self.btn_sua.configure(state="disabled", text="Đang sửa...")
        self.dung_sua.clear()
        self.btn_dung_sua.configure(state="normal")
        self.btn.configure(state="disabled")
        self._ghi("Đang sắp lại thứ tự dữ liệu...\n", xoa=True)
        threading.Thread(target=self._sua_nen,
                         args=(sdl, tm, lua_chon, sd, cpk, ng),
                         daemon=True).start()
        self.me.after(150, self._kiem_tra)

    def _sua_nen(self, sdl, tm, lua_chon, so_diem, cpk, nguong=None):
        try:
            d = sdl.doc_day_du(tm)
            # cham theo dung LSL dang hien tren man hinh (theo mac da chon)
            d["nguong"].update(nguong or {})
            khoi = sdl.xay_rang_buoc(d, lua_chon, so_diem, giu_bao_dam=True)

            def td(con, giay):
                self.hang_doi.put(("td_sua", (con, giay)))
            stop = getattr(self, "dung_sua", threading.Event()).is_set
            con, vong, giay, doi = sdl.sua(d, khoi, cpk, 180, None, td, stop)
            if stop():
                raise sdl.LoiSua("Đã dừng sửa. Không ghi đè dữ liệu gốc.")

            if con:
                raise sdl.LoiSua("Còn %d khối chưa đạt sau %.1f giây. Không ghi đè dữ liệu gốc." % (con, giay))

            def td2(i, n):
                self.hang_doi.put(("td", (i, n)))
            kg = sdl.ghi_lai(d, tien_do=td2)
            self.hang_doi.put(("xong_sua", (
                sdl.tom_tat(d, khoi, con, vong, giay, doi, kg), tm, con)))
        except (sdl.LoiSua, dtm.LoiDoc) as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi", "Không ghi được file.\n\nThường do file "
                                      "đang mở trong Excel hoặc trình đọc "
                                      "PDF.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def chay(self, da_sua=False):
        """Danh gia lai du lieu roi xuat bao cao SPC.

        Truoc khi xuat, cham lai moi day se xuat. Co day chua dat thi de
        nghi tu sap lai thu tu mau (giu nguyen gia tri) cho dat roi moi
        xuat. da_sua=True: vua sua xong, xuat luon khong hoi lai.
        """
        try:
            cot, bd, buoc, sd, cpk, ng, cfg, tm = self._thu()
        except dtm.LoiDoc as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        tong = len(cot) * len(bd)
        self._bm = None
        if self._dung_bieu_mau_keo(sd):
            # thu keo du 25 diem: moi diem bat dau mot file theo bieu mau
            # cua mac, gom gioi han chay - Rm - A; cot khac (Z...) ra file rieng
            try:
                cot_bm = self._cot_bm()
            except dtm.LoiDoc as e:
                messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
                return
            trong = set(cot_bm.values())
            self._bm = (self.v_mac.get(), ng, trong, cot_bm["ReH"])
            tong = (1 + len([k for k in cot if k not in trong])) * len(bd)

        chua = self._day_chua_dat(cot, bd, buoc, sd, cpk, ng)
        if chua is None:
            return
        if chua and not da_sua:
            mo_ta = "\n".join("   cột %s, bắt đầu %d — vi phạm %s"
                              % (k, s, ", ".join(l)) for k, s, l in chua[:8])
            if len(chua) > 8:
                mo_ta += "\n   ... và %d dãy khác" % (len(chua) - 8)
            dau = ("Đánh giá lại: %d/%d dãy CHƯA ĐẠT:\n\n%s\n\n"
                   % (len(chua), len(cot) * len(bd), mo_ta))
            if os.path.isdir(self.v_tm.get().strip()):
                tl = messagebox.askyesnocancel(
                    "Phát hiện dãy chưa đạt",
                    dau + "• Yes: tự SẮP LẠI THỨ TỰ mẫu cho đạt rồi xuất %d "
                    "file (giữ nguyên giá trị và mã mẫu; bản gốc chép sang "
                    "_truoc_khi_sua; %s)\n"
                    "• No: xuất luôn với dữ liệu hiện tại (file ghi rõ chưa "
                    "đạt)\n• Cancel: huỷ" % (tong, self._loi_nhan_word()))
                if tl is None:
                    return
                if tl:
                    self._bat_dau_sua(cot, bd, buoc, sd, cpk, ng, tm,
                                      xuat_sau=True)
                    return
            elif not messagebox.askyesno(
                    "Phát hiện dãy chưa đạt",
                    dau + "Nguồn là một file lẻ nên không tự sửa được (chỉ "
                    "sửa được khi chọn cả thư mục kết quả).\n\nVẫn xuất %d "
                    "file với dữ liệu hiện tại?" % tong):
                return
        elif tong > 20 and not da_sua and not messagebox.askyesno(
                "Nhiều file", "Đánh giá lại: cả %d dãy đều ĐẠT.\n\nSẽ sinh "
                              "%d file báo cáo SPC. Tiếp tục?"
                              % (len(cot) * len(bd), tong)):
            return
        self.btn.configure(state="disabled", text="Đang tạo...")
        self._ghi("Đang tạo %d file báo cáo SPC...\n" % tong, xoa=True)
        threading.Thread(
            target=self._chay_nen,
            args=(cot, bd, buoc, sd, cpk, ng, cfg, tm), daemon=True).start()
        self.me.after(150, self._kiem_tra)

    def _cot_bm(self):
        """Ba cot cua bieu mau keo. Hang gioi han chay lay ReL khi nguoi
        dung tick ReL ma bo tick ReH; con lai uu tien ReH."""
        def tick(k):
            return k in self.v_cot and self.v_cot[k].get()
        chay = "ReL" if tick("ReL") and not tick("ReH") else None
        return tao_spc.cot_bieu_mau_keo(self.du_lieu, chay)

    def _dung_bieu_mau_keo(self, so_diem):
        """Thu keo va du 25 diem thi xuat theo bieu mau cua mac vat lieu
        (mot file ba chi tieu). It diem hon thi bieu mau khong chua vua -
        quay ve file tung cot nhu truoc."""
        return (self.loai == "keo"
                and so_diem == tao_spc.SO_DIEM_BIEU_MAU_KEO
                and self.v_mac.get() in dtm.MAC_KEO)

    def _chay_nen(self, cot, bd, buoc, sd, cpk, ng, cfg, tm):
        try:
            def td(i, n):
                self.hang_doi.put(("td", (i, n)))
            kq = []
            if getattr(self, "_bm", None):
                mac, ng_bm, trong, chay = self._bm
                kq = tao_spc.sinh_bieu_mau_keo(
                    self.du_lieu, mac, bd, buoc, tm, cfg, ng_bm, cpk, td,
                    chay)
                cot = [k for k in cot if k not in trong]
            if cot:
                kq += tao_spc.sinh_nhieu(self.du_lieu, cot, bd, buoc, tm, cfg,
                                         sd, ng, cpk, td)
            self.hang_doi.put(("xong", (kq, tm)))
        except dtm.LoiDoc as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi", "Không ghi được file.\n\nThường do file "
                                      "đang mở trong Excel.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def _kiem_tra(self):
        tiep = True
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "td":
                    i, n = dl
                    self.lbl_td.configure(text="%d / %d" % (i, n))
                elif loai == "td_sua":
                    con, giay = dl
                    self.lbl_td.configure(text="còn %d khối, %.0fs"
                                               % (con, giay))
                elif loai == "xong_sua":
                    tom, tm, con = dl
                    self._tra_nut()
                    self._ghi(tom + "\n", xoa=True)
                    # doc lai du lieu vua sua, giu nguyen cot tick va LSL
                    tick = dict((k, v.get()) for k, v in self.v_cot.items())
                    lsl = dict((k, v.get()) for k, v in self.v_lsl.items())
                    self._nap(tm)
                    for k, v in tick.items():
                        if k in self.v_cot:
                            self.v_cot[k].set(v)
                    for k, v in lsl.items():
                        if k in self.v_lsl:
                            self.v_lsl[k].set(v)
                    self._ghi(tom + "\n\nĐã đọc lại dữ liệu sau khi sửa.\n")
                    xuat_sau, self._xuat_sau = getattr(self, "_xuat_sau", False), False
                    if xuat_sau and not con and self.du_lieu:
                        self.me.after(100, lambda: self.chay(da_sua=True))
                        return
                    if con:
                        messagebox.showwarning(
                            "Chưa xong",
                            "Hết thời gian mà vẫn còn %d khối chưa đạt.\n\n"
                            "Dữ liệu gốc không được ghi đè."
                            % con)
                    else:
                        messagebox.showinfo(
                            "Xong", "Đã sắp lại thứ tự, mọi dãy đang chọn "
                                    "đều đạt.\n\nBản gốc nằm trong thư mục "
                                    "con _truoc_khi_sua.")
                    tiep = False
                elif loai == "xong":
                    kq, tm = dl
                    self._tra_nut()
                    self._ghi(tao_spc.bao_cao_chu(kq)
                              + "\nThư mục: %s\n" % tm, xoa=True)
                    hong = sum(1 for k in kq if not k["spc"].dat)
                    so_file = len(set(k["file"] for k in kq))
                    if hong:
                        messagebox.showwarning(
                            "Có dãy chưa đạt",
                            "Đã tạo %d file (%d dãy chỉ tiêu), trong đó %d "
                            "dãy CHƯA ĐẠT.\n\nChi tiết từng mục vi phạm ở "
                            "khung Kết quả."
                            % (so_file, len(kq), hong))
                    elif messagebox.askyesno(
                            "Xong", "Đã tạo %d file báo cáo SPC, cả %d dãy đều ĐẠT.\n\n"
                                    "Mở thư mục bây giờ?"
                                    % (so_file, len(kq))):
                        self.mo_thu_muc()
                    tiep = False
                elif loai == "loi":
                    self._xuat_sau = False
                    self._tra_nut()
                    self._ghi("KHÔNG TẠO ĐƯỢC:\n\n" + dl + "\n", xoa=True)
                    messagebox.showerror("Không tạo được", dl)
                    tiep = False
                elif loai == "loi_he_thong":
                    self._xuat_sau = False
                    self._tra_nut()
                    self._ghi("LỖI KHÔNG MONG ĐỢI:\n\n" + dl, xoa=True)
                    messagebox.showerror("Lỗi", "Có lỗi ngoài dự kiến. Chi tiết ở khung Kết quả.")
                    tiep = False
        except queue.Empty:
            pass
        if tiep:
            self.me.after(150, self._kiem_tra)

    def _tra_nut(self):
        self.btn_dung_sua.configure(state="disabled")
        self.btn.configure(state="normal", text="▶   TẠO BÁO CÁO SPC")
        self.btn_sua.configure(state="normal", text="Sửa dữ liệu cho đạt")
        self.lbl_td.configure(text="")

    def mo_thu_muc(self):
        p = self.v_tm.get().strip()
        d = p if os.path.isdir(p) else os.path.dirname(p)
        if d and os.path.isdir(d):
            self.mo_thu_muc_ngoai(d)

    # ======================================================== BO CAU HINH
    # Thu muc nguon, cot va LSL doi theo tung bo du lieu; ngay danh gia va
    # ten du lieu cung vay - khong luu vao bo cau hinh.
    KHONG_LUU = ("v_tm", "v_cot", "v_lsl")
    KHONG_LUU_O = ("ngay_danh_gia", "du_lieu_danh_gia")

    def lay_cau_hinh(self):
        d = bo_cau_hinh.chup_bien(self, bo_qua=self.KHONG_LUU)
        for k in self.KHONG_LUU_O:
            d.get("v", {}).pop(k, None)
        return d

    def dat_cau_hinh(self, d):
        bo_cau_hinh.dat_bien(self, d, bo_qua=self.KHONG_LUU)
        if self.loai == "keo":
            # bo cau hinh luu truoc khi co o chon mac: suy ra tu o Vat lieu
            vl = self.v["vat_lieu_spc"].get().strip()
            if "v_mac" not in (d or {}) and vl in dtm.MAC_KEO:
                self.v_mac.set(vl)
            if self.v_mac.get() not in dtm.MAC_KEO:
                self.v_mac.set(dtm.MAC_KEO_MAC_DINH)
            self._ap_mac()

    def _nap_lan_truoc(self):
        d = bo_cau_hinh.doc_lan_truoc(self.khoa)
        if d:
            self.dat_cau_hinh(d)
