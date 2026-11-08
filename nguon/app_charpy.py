# -*- coding: utf-8 -*-
"""
PHAN MEM SINH DU LIEU TEST VA DAP CHARPY
=========================================
Giao dien cua so. Logic nam trong charpy_core.py.
"""

import os
import sys
import json
import queue
import threading
import traceback
import subprocess
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from _lib_noi_bo import kich_hoat as _kich_hoat_giai_ma
    _kich_hoat_giai_ma()          # khong lam gi neu chi co file .py that
    import charpy_core as core
    import bo_cau_hinh
    from tab_thu_keo import TabThuKeo
    from tab_bao_cao_spc import TabBaoCaoSPC
    from tab_kich_thuoc import TabKichThuoc
    from tab_thanh_phan import TabThanhPhan
    from tab_danh_gia_thanh_phan import TabDanhGiaThanhPhan
    from tab_danh_gia_kich_thuoc import TabDanhGiaKichThuoc
    import ui_theme as ui
except Exception as _loi_khoi_dong:
    # Bao loi bang cua so (khong dung console vi pythonw khong co man
    # hinh dong lenh) roi thoat gon gang, thay vi treo may hoac dong
    # im lim khong ro nguyen nhan.
    _r = tk.Tk()
    _r.withdraw()
    messagebox.showerror(
        "Không khởi động được",
        "Phần mềm không khởi động được.\n\n%s\n\n"
        "Hãy chạy lại CHAY_PHAN_MEM.bat, hoặc liên hệ người phát triển."
        % _loi_khoi_dong)
    _r.destroy()
    sys.exit(1)

TIEU_DE = "Phần mềm tự động hoá báo cáo thử nghiệm cơ tính"
PHU_DE = "Phòng NCCN  •  Thử va đập Charpy  •  Thử kéo  •  Kích thước  •  Thành phần hoá  •  Đánh giá SPC trong từng tab"


def mo_bang_windows(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Lỗi", "Không mở được:\n%s\n\n%s" % (path, e))


class App:
    def __init__(self, root):
        self.root = root
        self.hang_doi = queue.Queue()
        self.thu_muc_cuoi = None

        root.title(TIEU_DE)
        root.geometry("1100x850")
        root.minsize(980, 720)
        
        ui.apply_app_styles(root)
        self._giao_dien()
        self._nap_lan_truoc()

    # ----------------------------------------------------------- giao dien
    def _giao_dien(self):
        # Header ung dung hien dai
        self.header = ui.ModernHeader(self.root, TIEU_DE, PHU_DE, badge_text="🟢 SẴN SÀNG")
        self.header.pack(fill="x")

        # Notebook tabs
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=12, pady=(10, 10))
        # Nhom theo LOAI THU NGHIEM. Moi tab co cac buoc con, buoc cuoi la
        # Danh gia SPC cua chinh loai thu do (sinh xong tu nap san).
        tab1 = ttk.Frame(self.nb)
        tab2 = ttk.Frame(self.nb)
        tab3 = ttk.Frame(self.nb)
        tab4 = ttk.Frame(self.nb)
        tab5 = ttk.Frame(self.nb)
        self.nb.add(tab1, text="  📊 1. Va đập Charpy  ")
        self.nb.add(tab2, text="  📈 2. Thử kéo  ")
        self.nb.add(tab3, text="  📐 3. Kích thước  ")
        self.nb.add(tab5, text="  🧪 4. Thành phần hoá  ")
        self.nb.add(tab4, text="  💡 5. Hướng dẫn  ")

        f1, f2 = self._tab_con(tab1, "① Sinh dữ liệu", "② Đánh giá SPC")
        self.nb_charpy = f1.master
        self._tab_sinh(f1)
        self.tab_spc_charpy = TabBaoCaoSPC(f2, mo_bang_windows,
                                           khoa="spc_charpy",
                                           ten="Đánh giá SPC va đập",
                                           loai="charpy")

        self.tab_thu_keo = TabThuKeo(tab2, mo_bang_windows)

        f1, f2 = self._tab_con(tab3, "① Sinh dữ liệu & xuất báo cáo",
                               "② Đánh giá SPC")
        self.nb_kich_thuoc = f1.master
        self.tab_kich_thuoc = TabKichThuoc(f1, mo_bang_windows)
        self.tab_danh_gia_kt = TabDanhGiaKichThuoc(f2, mo_bang_windows,
                                                   tab_sinh=self.tab_kich_thuoc)
        self.tab_kich_thuoc.khi_xong = self.tab_danh_gia_kt.dat_san
        f1, f2 = self._tab_con(tab5, "① Sinh dữ liệu & xuất báo cáo",
                               "② Đánh giá SPC")
        self.nb_thanh_phan = f1.master
        self.tab_thanh_phan = TabThanhPhan(f1, mo_bang_windows)
        self.tab_danh_gia_tp = TabDanhGiaThanhPhan(f2, mo_bang_windows,
                                                   tab_sinh=self.tab_thanh_phan)
        self.tab_thanh_phan.khi_xong = self.tab_danh_gia_tp.dat_san
        self._tab_huong_dan(tab4)

    def _tab_con(self, me, *ten):
        """Notebook tab con (giong tab Thu keo). Tra ve cac khung."""
        nb = ttk.Notebook(me)
        nb.pack(fill="both", expand=True, padx=4, pady=(6, 0))
        khung = []
        for t in ten:
            f = ttk.Frame(nb)
            nb.add(f, text="  %s  " % t)
            khung.append(f)
        return khung

    # ---------------------------------------------------- TAB 1: SINH DU LIEU
    def _tab_sinh(self, me):
        ngoai = ttk.Frame(me)
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

        self.thanh_cau_hinh = bo_cau_hinh.ThanhCauHinh(
            than, "charpy", "Va đập Charpy",
            self.lay_cau_hinh, self.dat_cau_hinh)

        # Row 1: 2 Cards (Khoi 1 & Khoi 2)
        r1 = ttk.Frame(than)
        r1.pack(fill="x", padx=12, pady=(8, 4))
        r1.columnconfigure(0, weight=4)
        r1.columnconfigure(1, weight=6)

        # ---- Khoi 1: So luong & ma mau ----
        c1 = ui.ModernCard(r1, title="Số lượng và quy tắc đặt mã mẫu", badge_num=1)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        k1 = c1.get_body()
        k1.columnconfigure(1, weight=1)

        self.v_n = tk.StringVar(value="1000")
        self.v_stt = tk.StringVar(value="1")
        self.v_tt = tk.StringVar(value="")
        self.v_cs = tk.StringVar(value="0")

        self._o(k1, 0, "Số lượng mẫu cần sinh", self.v_n, "VD: 1000")
        self._o(k1, 1, "Số thứ tự bắt đầu", self.v_stt, "VD: 1")
        self._o(k1, 2, "Tiền tố mã mẫu", self.v_tt, "để trống = chỉ dùng số")
        self._o(k1, 3, "Số chữ số của mã", self.v_cs, "0 = không đệm số 0")
        ttk.Label(k1, style="SurfacePhu.TLabel", wraplength=380, justify="left",
                  text="Ví dụ: để trống tiền tố, 0 chữ số → mã là 1, 2, 3...\n"
                       "Tiền tố VF, 5 chữ số → VF00001, VF00002...").grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ---- Khoi 2: Khoang so lieu, khai rieng cho tung lan thu ----
        c2 = ui.ModernCard(r1, title="Khoảng số liệu sinh ngẫu nhiên (3 lần thử)", badge_num=2)
        c2.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        k2 = c2.get_body()
        for c in (2, 3, 4, 5):
            k2.columnconfigure(c, weight=1)

        self.v_chung = tk.BooleanVar(value=True)
        ttk.Checkbutton(k2, text="Cả 3 lần thử dùng chung một cấu hình",
                        variable=self.v_chung, command=self.doi_dung_chung).grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 6))

        for j, nhan in enumerate(["Energy từ", "đến", "Các giá trị A (mm)",
                                  "Các giá trị B (mm)"]):
            ttk.Label(k2, text=nhan, style="SurfacePhu.TLabel").grid(
                row=1, column=2 + j, sticky="w", padx=(0, 6))

        self.v_lan = []          # 3 dong x 4 o
        self.e_lan = []
        mac_dinh = [("50", "90", "7, 7.5, 8, 8.5", "4.5, 5, 5.5")] * 3
        for i in range(3):
            ttk.Label(k2, text="Lần %d" % (i + 1), style="Surface.TLabel").grid(
                row=2 + i, column=0, sticky="w", padx=(0, 8), pady=2)
            dong_v, dong_e = [], []
            for j in range(4):
                v = tk.StringVar(value=mac_dinh[i][j])
                e = ttk.Entry(k2, textvariable=v,
                              width=7 if j < 2 else 17)
                e.grid(row=2 + i, column=2 + j, sticky="ew", padx=(0, 5), ipady=2)
                dong_v.append(v)
                dong_e.append(e)
            self.v_lan.append(dong_v)
            self.e_lan.append(dong_e)

        self.v_nd = tk.StringVar(value="0")
        self.v_seed = tk.StringVar(value="")
        ttk.Label(k2, text="Nhiệt độ thử (°C)", style="Surface.TLabel").grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Entry(k2, textvariable=self.v_nd, width=8).grid(
            row=6, column=2, sticky="w", pady=(8, 0), ipady=2)
        ttk.Label(k2, text="Seed", style="Surface.TLabel").grid(
            row=6, column=3, sticky="e", padx=(0, 6), pady=(8, 0))
        ttk.Entry(k2, textvariable=self.v_seed, width=10).grid(
            row=6, column=4, sticky="w", pady=(8, 0), ipady=2)
        ttk.Label(k2, text="để trống = mỗi lần khác", style="SurfacePhu.TLabel").grid(
            row=6, column=5, sticky="w", pady=(8, 0))

        try:
            bt = core.BangTra()
            gy = ("A chỉ nhận: %s\nB chỉ nhận: %s"
                  % (", ".join("%g" % x for x in bt.A),
                     ", ".join("%g" % x for x in bt.B)))
        except Exception:
            gy = "Không đọc được bảng tra phần trăm mặt gãy."
        ttk.Label(k2, style="SurfacePhu.TLabel", wraplength=540, justify="left",
                  text=gy).grid(row=7, column=0, columnspan=6, sticky="w",
                                pady=(6, 0))
        self.doi_dung_chung()

        # ---- Khoi 3: Thong tin tren bao cao ----
        c3 = ui.ModernCard(than, title="Thông tin cố định in trên báo cáo", badge_num=3)
        c3.pack(fill="x", padx=12, pady=4)
        k3 = c3.get_body()
        for c in (1, 3):
            k3.columnconfigure(c, weight=1)

        self.v_ngay = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.v_tester = tk.StringVar(value="Tran Dang Khoa")
        self.v_duyet = tk.StringVar(value="Doan Ngoc Thoat")
        self.v_tc = tk.StringVar(value="JIS Z 2422:2023")
        self.v_vl = tk.StringVar(value="Steel")
        self.v_kt = tk.StringVar(value="55x10x10 mm")

        doi = [("Ngày thử", self.v_ngay), ("Người thử", self.v_tester),
               ("Người phê duyệt", self.v_duyet), ("Tiêu chuẩn", self.v_tc),
               ("Vật liệu", self.v_vl), ("Kích thước mẫu", self.v_kt)]
        for i, (nhan, var) in enumerate(doi):
            r, c = i // 2, (i % 2) * 2
            ttk.Label(k3, text=nhan, style="Surface.TLabel").grid(
                row=r, column=c, sticky="w", padx=(0, 8), pady=3)
            ttk.Entry(k3, textvariable=var).grid(
                row=r, column=c + 1, sticky="ew", padx=(0, 16), ipady=2)

        # ---- Khoi 4: Danh gia SPC ----
        c5 = ui.ModernCard(than, title="Đánh giá độ ổn định SPC (I-MR chart)", badge_num=4)
        c5.pack(fill="x", padx=12, pady=4)
        k5 = c5.get_body()
        k5.columnconfigure(1, weight=1)

        self.v_file_mau = tk.StringVar(value="")
        self.v_lsl = tk.StringVar(value="27")
        self.v_cpk = tk.StringVar(value="1.33")

        ttk.Label(k5, text="File mẫu SPC (đọc LSL, vật liệu, sản phẩm)", style="Surface.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(k5, textvariable=self.v_file_mau).grid(row=0, column=1,
                                                         sticky="ew", ipady=2)
        ttk.Button(k5, text="📂 Chọn...", style="Outline.TButton", width=12,
                   command=self.chon_file_mau).grid(row=0, column=2, padx=(8, 0))
        ttk.Label(k5, style="SurfacePhu.TLabel",
                  text="Để trống thì dùng ngưỡng gõ tay bên dưới.").grid(
            row=1, column=1, sticky="w", pady=(2, 4))

        ttk.Label(k5, text="Giới hạn dưới LSL (J)", style="Surface.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(k5, textvariable=self.v_lsl, width=12).grid(
            row=2, column=1, sticky="w", ipady=2)
        ttk.Label(k5, text="Cpk tối thiểu", style="Surface.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 8), pady=(4, 0))
        ttk.Entry(k5, textvariable=self.v_cpk, width=12).grid(
            row=3, column=1, sticky="w", ipady=2)
        self.lbl_spc = ttk.Label(k5, style="SurfacePhu.TLabel", wraplength=900,
                                 justify="left", text="")
        self.lbl_spc.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        bo_cau_hinh.dang_ky_thong_tin("charpy", "Va đập Charpy", {
            "nguoi_thu": self.v_tester, "nguoi_phe_duyet": self.v_duyet,
            "tieu_chuan_thu": self.v_tc, "vat_lieu": self.v_vl})

        # ---- Khoi 5: Chon xuat ----
        k4 = tk.Frame(than, bg=ui.COLOR_SURFACE, padx=14, pady=8,
                      highlightbackground=ui.COLOR_BORDER, highlightthickness=1)
        k4.pack(fill="x", padx=12, pady=4)
        self.v_word = tk.BooleanVar(value=True)
        self.v_th = tk.BooleanVar(value=True)
        self.v_gon = tk.BooleanVar(value=True)
        self.v_spc = tk.BooleanVar(value=True)
        
        tk.Label(k4, text="Các định dạng file xuất ra:", font=("Segoe UI", 9, "bold"),
                 fg=ui.COLOR_TEXT_MAIN, bg=ui.COLOR_SURFACE).pack(side="left", padx=(0, 16))
        ttk.Checkbutton(k4, text="Báo cáo Word từng mẫu (.docx)",
                        variable=self.v_word).pack(side="left", padx=8)
        ttk.Checkbutton(k4, text="Bảng tổng hợp BM31 (.xlsx)",
                        variable=self.v_th).pack(side="left", padx=8)
        ttk.Checkbutton(k4, text="Bảng tổng hợp gọn (.xlsx)",
                        variable=self.v_gon).pack(side="left", padx=8)
        ttk.Checkbutton(k4, text="Báo cáo độ ổn định SPC (.xlsx)",
                        variable=self.v_spc).pack(side="left", padx=8)

        # ---- Thanh hanh dong (Action Bar) ----
        d = tk.Frame(than, bg=ui.COLOR_BG, padx=4, pady=8)
        d.pack(fill="x", padx=12)
        self.btn = ttk.Button(d, text="▶   SINH DỮ LIỆU", style="Primary.TButton",
                              command=self.chay)
        self.btn.pack(side="left")
        ttk.Button(d, text="📂 Mở thư mục kết quả", style="Outline.TButton",
                   command=self.mo_thu_muc).pack(side="left", padx=10)
        self.tien_do = ttk.Progressbar(d, mode="determinate", length=320)
        self.tien_do.pack(side="left", padx=10)
        self.lbl_td = ttk.Label(d, style="Phu.TLabel", text="")
        self.lbl_td.pack(side="left")

        # ---- Modern Terminal Log ----
        self.log = ui.ModernTerminal(than, title="NHẬT KÝ SINH DỮ LIỆU VA ĐẬP CHARPY & ĐỐI CHIẾU SPC", height=10)
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self._ghi("Sẵn sàng. Điền tham số rồi bấm nút '▶ SINH DỮ LIỆU'.\n"
                  "Sinh 1000 file mất khoảng 30 giây.\n")

    # ----------------------------------------------------- TAB 3: HUONG DAN
    HUONG_DAN = """
╔══════════════════════════════════════════════════════════════════════════╗
║  PHẦN MỀM TỰ ĐỘNG HOÁ BÁO CÁO THỬ NGHIỆM CƠ TÍNH                         ║
╚══════════════════════════════════════════════════════════════════════════╝

  Sinh số liệu GIẢ LẬP để kiểm thử hệ thống khác.
  Không phải kết quả đo thực tế.

  Các tab xếp theo LOẠI THỬ NGHIỆM:

     TAB 1  Va đập Charpy
            ① Sinh dữ liệu   N báo cáo Word + bảng tổng hợp + báo cáo SPC
            ② Đánh giá SPC   chọn dãy 25 mẫu, chấm I-MR, xuất file SPC

     TAB 2  Thử kéo
            ① Bước 1  Sinh biểu đồ & số liệu  → 1 file PDF + 1 file Excel
            ② Bước 2  Gộp báo cáo Word        → mỗi mẫu một file Word
            ③ Đánh giá SPC                    → chấm dãy 25 mẫu, file SPC
            Sinh xong Bước 1, file PDF, Excel, kích thước mẫu tự điền sang
            Bước 2, và thư mục vừa sinh tự nạp vào bước ③.

     TAB 3  Kích thước
            ① Sinh dữ liệu & xuất báo cáo JIS.QT.07.BM.01
            ② Đánh giá SPC   chấm từng chỉ tiêu, xuất sheet Báo cáo cho
                             chỉ tiêu bất kỳ (biểu mẫu gốc chỉ có chỉ tiêu 1)

     TAB 4  Thành phần hoá
            ① Sinh dữ liệu   chọn mác SCW410 / SCW480 → bảng tổng hợp + kết
                             quả in từng mẫu + file SPC (mỗi file 25 mẫu)
            ② Đánh giá SPC   chấm 10 nguyên tố theo dãy bất kỳ, sửa số
                             cho đạt (ghi đè vào thư mục), xuất file SPC

  Bước Đánh giá SPC nằm ngay trong tab của từng loại thử, chỉ nhận đúng
  loại dữ liệu của tab đó. Sinh xong thì kết quả vừa sinh tự nạp sẵn.

  Cuộn xuống để xem hướng dẫn chi tiết từng tab.


┌──────────────────────────────────────────────────────────────────────────┐
│  BỘ CẤU HÌNH ĐÃ LƯU  (có ở đầu mọi tab)                                  │
└──────────────────────────────────────────────────────────────────────────┘

  Mỗi tab có danh sách bộ cấu hình RIÊNG — tab Va đập và tab Thử kéo không
  dùng chung, vì mỗi loại thử thường khác người làm, khác vật liệu.

    🗂 Bộ cấu hình  ▼     Chọn một bộ → mọi ô trong tab tự điền theo bộ đó.
                          Sau đó vẫn sửa tay thoải mái; bộ đã lưu không đổi
                          trừ khi bấm Lưu.
    💾 Lưu                Lưu các ô đang điền thành một bộ, đặt tên tuỳ ý
                          (ví dụ "SCW480 – khách A"). Trùng tên thì hỏi
                          trước khi ghi đè.
    🗑 Xoá                Xoá bộ đang chọn (các ô đang điền giữ nguyên).
    📋 Lấy thông tin từ tab khác
                          Chỉ khi bạn bấm: chép các ô CÙNG VAI TRÒ (người
                          phê duyệt, vật liệu, tiêu chuẩn...) từ tab khác
                          sang. Không bao giờ tự đồng bộ.

  Ngoài ra mỗi tab tự NHỚ thông số của lần chạy trước.
  Tất cả nằm trong thư mục  cau_hinh/  cạnh phần mềm.


╔══════════════════════════════════════════════════════════════════════════╗
║  TAB 1 — VA ĐẬP CHARPY                                                   ║
╚══════════════════════════════════════════════════════════════════════════╝


┌──────────────────────────────────────────────────────────────────────────┐
│  DÙNG TRONG 3 BƯỚC                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  1.  Điền số lượng mẫu ở ô đầu tiên (ô "Số lượng mẫu cần sinh").

  2.  Chọn file mẫu SPC ở Khối 4 — phần mềm tự đọc ngưỡng LSL, vật liệu,
      tên sản phẩm từ file đó. Không có file thì gõ tay ô LSL.

  3.  Bấm nút xanh  ▶ SINH DỮ LIỆU.

  Các ô còn lại đã điền sẵn giá trị hợp lý, không cần đụng tới.


┌──────────────────────────────────────────────────────────────────────────┐
│  SINH RA NHỮNG GÌ                                                        │
└──────────────────────────────────────────────────────────────────────────┘

  output/Charpy_<ngày giờ>_<số mẫu>mau/
     │
     ├─  BaoCao/KetQua_1.docx, KetQua_2.docx ...    mỗi mẫu một báo cáo Word
     ├─  TongHop_VaDap.xlsx                          bảng tổng hợp BM31
     ├─  TongHop_VaDap_Gon.xlsx                      bảng tổng hợp gọn (chỉ bảng số liệu)
     └─  DanhGia_DoOnDinh_SPC.xlsx                   báo cáo I-MR chart

  Mỗi lần chạy một thư mục riêng, không bao giờ ghi đè lên nhau.
  Sinh 1000 file mất khoảng 30 giây.


┌──────────────────────────────────────────────────────────────────────────┐
│  GIẢI THÍCH TỪNG Ô                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  KHỐI 1 — SỐ LƯỢNG VÀ MÃ MẪU

    Số lượng mẫu           Sinh bao nhiêu mẫu, ra bấy nhiêu file Word
    Số thứ tự bắt đầu      Mặc định 1. Đổi khi muốn nối tiếp lô trước
    Tiền tố mã mẫu         Để trống  → mã là  1, 2, 3...
                           Gõ "VF" + 5 chữ số  → VF00001, VF00002...

  KHỐI 2 — KHOẢNG SỐ LIỆU  (khai riêng cho từng lần thử)

    Mỗi mẫu thử 3 lần. Bảng có 3 dòng, mỗi dòng khai riêng cho một lần:

              Energy từ   đến    Các giá trị A (mm)   Các giá trị B (mm)
      Lần 1      52        58      7, 7.5              4.5, 5
      Lần 2      54        60      8, 8.5              5, 5.5
      Lần 3      50        56      7.5, 8, 8.5         4.5, 5.5

    [x] Cả 3 lần dùng chung một cấu hình
           Bật (mặc định) → chỉ cần điền dòng 1, hai dòng dưới mờ đi và
           tự ăn theo. Tắt → khai riêng từng dòng.

    LƯU Ý:  A và B chỉ dùng để tra ra % MẶT GÃY.
            SPC chỉ đánh giá NĂNG LƯỢNG HẤP THỤ.
            Nên muốn dễ đạt SPC thì chỉnh khoảng Energy, không phải A/B.

    Nhiệt độ thử           In lên báo cáo và cột °C của bảng tổng hợp
    Seed                   Điền một số → lần sau chạy lại ra Y HỆT bộ dữ
                           liệu cũ. Rất tiện khi cần tái hiện một ca lỗi

  KHỐI 3 — THÔNG TIN IN TRÊN BÁO CÁO

    Ngày thử, người thử, người phê duyệt, tiêu chuẩn, vật liệu, kích thước.
    Phần mềm nhớ lại lần dùng trước nên hầu như không phải gõ lại.

  KHỐI 4 — ĐÁNH GIÁ ĐỘ ỔN ĐỊNH SPC

    File mẫu SPC           Chọn file "Va đập - SCW410 - ....xlsx" — có
                           sẵn trong thư mục file_mau/, lần đầu mở phần
                           mềm tự chọn luôn. Phần mềm tự đọc:
                              ô M11 → giới hạn dưới LSL
                              ô E9  → vật liệu       (SCW410, SCW480...)
                              ô E11 → tên sản phẩm
                              ô E10 → tiêu chuẩn áp dụng
                           Đổi mác thép chỉ cần đổi file, không sửa gì thêm

    Giới hạn dưới LSL      Dùng khi không chọn file mẫu
    Cpk tối thiểu          Mặc định 1.33


┌──────────────────────────────────────────────────────────────────────────┐
│  PHẦN MỀM TỰ KIỂM TRA NHƯ THẾ NÀO                                        │
└──────────────────────────────────────────────────────────────────────────┘

  Sinh xong, phần mềm tự đối chiếu với tiêu chuẩn SPC. Không đạt thì nó
  RANDOM LẠI và sinh bộ khác, lặp tới khi đạt (tối đa 300 lần).

  Điều kiện phải thoả cùng lúc:

     1.  Mọi điểm nằm trong các giới hạn LSL/USL đã khai.
     2.  Đạt đủ 8 luật I-chart và 3 luật MR-chart.
     3.  Có đủ biến thiên để ước lượng sigma = MR̄ / 1.128.
     4.  Cpk đạt ngưỡng đã khai (mặc định 1.33).
         Một phía: dùng CPL hoặc CPU. Hai phía: Cpk = min(CPL, CPU).

  File mẫu SPC chỉ chứa 25 điểm. Nhiều hơn 25 mẫu thì phần mềm tự chia
  thành nhiều file SPC, mỗi file tối đa 25 mẫu và mỗi nhóm phải đạt riêng.


┌──────────────────────────────────────────────────────────────────────────┐
│  KHOẢNG ENERGY ẢNH HƯỞNG TỚI Cpk                                         │
└──────────────────────────────────────────────────────────────────────────┘

  Khoảng random càng rộng thì σ càng lớn, Cpk càng thấp.

        Khoảng          Cpk ước tính      Đạt ngay lần đầu
        ─────────────────────────────────────────────────
        50 – 90 J           2.15               87 %      ← mặc định
        40 – 90 J           1.52               63 %
        30 – 90 J           1.10               13 %      ← có cảnh báo

  Nếu khoảng quá rộng, phần mềm hỏi trước khi chạy và gợi ý khoảng hẹp hơn.


┌──────────────────────────────────────────────────────────────────────────┐
│  LỖI THƯỜNG GẶP                                                          │
└──────────────────────────────────────────────────────────────────────────┘

  "Bảng tra không có cặp kích thước A=..., B=..."
        → Giá trị A hoặc B không nằm trong bảng tra.
          Xem danh sách cho phép ngay dưới ô nhập ở Khối 2.

  "Đã thử 300 lần nhưng không sinh được ... thoả mãn SPC"
        → Khoảng Energy quá rộng so với ngưỡng LSL.
          Thu hẹp khoảng theo gợi ý phần mềm đưa ra.

  "Không ghi được file kết quả"
        → File đang mở trong Excel. Đóng Excel rồi bấm lại.

  "Windows protected your PC"  (nền xanh, lần đầu mở)
        → Bấm  More info  rồi  Run anyway.  Chỉ xảy ra một lần.




╔══════════════════════════════════════════════════════════════════════════╗
║  TAB 2 — THỬ KÉO  ·  BƯỚC 1 — SINH BIỂU ĐỒ & SỐ LIỆU                     ║
╚══════════════════════════════════════════════════════════════════════════╝

  Tab Thử kéo có hai bước con, làm lần lượt:

        BƯỚC 1  ──►  BƯỚC 2  ──►  báo cáo Word
      sinh PDF+Excel    gộp lại

  Bộ cấu hình của tab Thử kéo lưu cả hai bước cùng lúc.

┌──────────────────────────────────────────────────────────────────────────┐
│  DÙNG TRONG 3 BƯỚC                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  1.  Điền "Số lượng mẫu cần sinh" ở Khối 1. Các ô còn lại đã điền sẵn
      đúng cho SCW480, lần đầu không cần đụng tới.

  2.  Bấm nút xanh  ▶ SINH BIỂU ĐỒ VÀ SỐ LIỆU.

  3.  Xong, phần mềm tự điền file PDF, Excel, d₀, L₀, hệ số giãn dài sang
      Bước 2 và hỏi có chuyển sang luôn không (hoặc bấm  "➜ Chuyển sang
      Bước 2"). Ở đó bấm  ▶ GỘP BÁO CÁO  là ra Word.

┌──────────────────────────────────────────────────────────────────────────┐
│  SINH RA NHỮNG GÌ                                                        │
└──────────────────────────────────────────────────────────────────────────┘

  output/Keo_Sinh_<ngày giờ>_<số mẫu>mau/
     │
     ├─  BieuDo_<N>.pdf                      N trang, mỗi trang một biểu đồ
     ├─  SoLieu_<N>.xlsx                     5 sheet, xem bên dưới
     ├─  DanhGia_DoOnDinh_SPC_ReL_mau1-25.xlsx   file SPC cho 25 mẫu đầu
     └─  BaoCao_Phu_SPC.txt                  bảng phủ mọi bước nhảy

  Các sheet trong file Excel:
       Tong_hop     thông số đặc trưng (lực theo kN) — Bước 2 đọc sheet này
       Co_ly_tinh   d, du, FeL/FeH/Fm (N), ReL, ReH, Rm, A, Z — điền BM32
       Du_lieu      đường cong từng mẫu, lấy mẫu mỗi 0,5 mm
       Bieu_do_mau  ba mẫu đầu, để xem nhanh
       Huong_dan    ghi chú đơn vị và công thức

  1000 mẫu mất khoảng 70 giây, PDF chừng 5 MB.

┌──────────────────────────────────────────────────────────────────────────┐
│  GIẢI THÍCH TỪNG KHỐI                                                    │
└──────────────────────────────────────────────────────────────────────────┘

  KHỐI 1 — Số lượng, mã mẫu, seed.  Tối thiểu 25 mẫu (một bộ SPC cần 25
           điểm). Seed điền số thì lần sau chạy lại ra y hệt.

  KHỐI 2 — Kích thước mẫu thử.  d₀ mặc định 14, dung sai ±0,02.
           Độ thắt Z = (S₀ − Su)/S₀ × 100 %   (JIS Z 2241, như mẫu 522)
                  S₀ = π·d²/4      Su = π·du²/4
           du suy ngược từ Z và d, nên Z trong Excel và trên báo cáo Word
           luôn trùng nhau.

  KHỐI 3 — Dải sinh của năm cột chịu SPC.  Mỗi cột khai trung bình, độ
           lệch chuẩn, và ngưỡng tối thiểu (dòng 11 của BM32).
           Để trống ô ngưỡng = cột đó không có giới hạn nghiệm thu, vẫn
           xét ổn định nhưng không xét Cpk. BM32 chỉ đặt ngưỡng cho ReL
           (≥275), Rm (≥480) và A (≥20).

           Biểu đồ chỉ phụ thuộc ReH, ReL, Rm (độ cao ba điểm) và A (vị trí
           đứt). Lực lúc đứt = 70–78 % Fm, không phụ thuộc Z.

  KHỐI 4 — Đánh giá độ ổn định SPC.  Đây là phần quan trọng nhất.

           Phần mềm dùng ĐỦ BỘ 11 mục của biểu mẫu BÁO CÁO ĐÁNH GIÁ ĐỘ
           ỔN ĐỊNH CHẤT LƯỢNG: 8 mục I-chart, 3 mục MR-chart, cộng Cpk.
           Rồi nó SẮP LẠI THỨ TỰ các mẫu cho tới khi mọi cách lấy 25 giá
           trị trong phạm vi đã chọn đều đạt cả 11 mục, trên cả 5 cột.

           Việc sắp lại chỉ ĐỔI THỨ TỰ, không đổi tập giá trị — nên dải
           số và ngưỡng nghiệm thu giữ nguyên. MR̄, sigma, Cpk và các
           giới hạn kiểm soát có thể đổi; phần mềm phải chấm lại.
           Hết thời gian mà còn fail thì kết quả vẫn ghi rõ CHƯA ĐẠT.

           BỐN MỨC BẢO ĐẢM

              Liên tiếp   bất kỳ 25 mẫu LIỀN NHAU nào, ở bất kỳ đâu trong
              (mặc định)  lô. Lấy mẫu 1–25, hay 301–325, hay 700–724 đều
                          đạt. Khoảng 15 giây cho 1000 mẫu.

              Cơ bản      15 bước nhảy tròn, nhưng CHỈ khi bắt đầu từ đầu
                          lô: mẫu 1, 41, 81... thì đạt, còn mẫu 301, 341,
                          381... thì không. Khoảng 20 giây.

              Gộp cả hai  cả hai điều trên. Mạnh nhất nhưng mất chừng 10
                          phút và thường vẫn còn vài khối chưa đạt.

              Tự khai     danh sách bước nhảy của riêng bạn, tính từ đầu lô.

           ĐIỀU PHẢI BIẾT

           KHÔNG thể bảo đảm MỌI cách lấy 25 giá trị. Bộ luật SPC vốn được
           thiết kế để một quá trình ổn định THẬT vẫn thỉnh thoảng chạm
           phải — đó là bản chất của kiểm soát thống kê. Với hàng chục
           nghìn cửa sổ khả dĩ, việc không cửa sổ nào chạm là bất khả.

           Hơn nữa: càng tối ưu cho nhiều cách lấy, các cách lấy NGOÀI
           danh sách càng tệ đi. Dữ liệu chưa sắp xếp đạt khoảng 70%; sau
           khi sắp cho một tập lớn, các cách lấy ngoài tập chỉ còn 4–29%.

           BẢNG PHỦ

           Xong việc, phần mềm quét lại và ghi vào BaoCao_Phu_SPC.txt:

              buoc      bat dau tu dau lo     bat dau tu giua lo
                 1      1/1 DAT               88/88 (100%)  <- da bao dam
                 2      2/2 DAT               13/76 (17%)
                 5      1/5                   16/80 (20%)

           Hai cột tách riêng vì chúng cho kết quả rất khác nhau. Bảng này
           cho bạn thấy chính xác cách lấy nào được bảo đảm, cách nào không.

┌──────────────────────────────────────────────────────────────────────────┐
│  LỖI THƯỜNG GẶP Ở BƯỚC 1                                                 │
└──────────────────────────────────────────────────────────────────────────┘

  "Số lượng mẫu phải từ 25 trở lên"
        → Một bộ đánh giá SPC cần đúng 25 điểm.

  "Không sinh được giá trị thoả ngưỡng nghiệm thu"
        → Trung bình khai quá sát ngưỡng tối thiểu, hoặc độ lệch chuẩn
          quá lớn. Nâng trung bình lên hoặc giảm độ lệch xuống.

  Còn khối chưa đạt sau khi hết giờ
        → Tăng ô "Thời gian tối đa", hoặc bớt bước nhảy trong danh sách,
          hoặc chuyển từ "Kỹ" về "Nhanh".




╔══════════════════════════════════════════════════════════════════════════╗
║  TAB 2 · BƯỚC 2 — GỘP BÁO CÁO WORD                                       ║
╚══════════════════════════════════════════════════════════════════════════╝

  Nếu vừa sinh ở Bước 1 thì hai file PDF, Excel và kích thước mẫu đã được
  điền sẵn — chỉ việc bấm  ▶ GỘP BÁO CÁO. Dùng được cả với PDF/Excel có
  sẵn từ nguồn khác.

  Bài toán: có sẵn 1000 biểu đồ kéo nằm trong một file PDF 1000 trang, và
  1000 dòng số liệu tương ứng nằm trong một file Excel. Cần 1000 file Word
  báo cáo, mỗi file gồm một bảng số liệu và đúng cái biểu đồ của nó.

  Phần mềm ghép theo THỨ TỰ:

        trang PDF thứ 1  +  dòng Excel thứ 1   →   file Word thứ 1
        trang PDF thứ 2  +  dòng Excel thứ 2   →   file Word thứ 2
        ...

  Nếu PDF và Excel lệch số lượng, phần mềm ghép được bao nhiêu cặp thì
  báo bấy nhiêu và ghi rõ ở khung Kết quả.


┌──────────────────────────────────────────────────────────────────────────┐
│  DÙNG TRONG 3 BƯỚC                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  1.  Khối 1 — chọn file PDF biểu đồ và file Excel số liệu.
      Chọn xong, dòng chữ xám ngay bên dưới báo luôn: PDF có bao nhiêu
      biểu đồ, Excel có bao nhiêu dòng. Hai số này nên bằng nhau.

  2.  Khối 2 — chọn "Toàn bộ", hoặc "Chỉ xuất ... file" nếu muốn thử
      trước vài file cho nhanh.

  3.  Bấm nút xanh  ▶ GỘP BÁO CÁO.

  Các khối 3, 4, 5 đã điền sẵn giá trị hợp lý, lần đầu không cần đụng tới.


┌──────────────────────────────────────────────────────────────────────────┐
│  SINH RA NHỮNG GÌ                                                        │
└──────────────────────────────────────────────────────────────────────────┘

  output/Keo_<ngày giờ>_<số mẫu>mau/
     │
     └─  BaoCao/Tensile_S2-1.docx, Tensile_S2-2.docx ...

  Mỗi file Word gồm đúng ba phần như file mẫu:
       · dòng Standard và Sample
       · bảng 13 cột thông số cơ lý tính
       · biểu đồ kéo của chính mẫu đó, chèn căn giữa bên dưới bảng

  Tốc độ khoảng nửa giây một file. 1000 file mất chừng 8 phút và chiếm
  chừng 90 MB đĩa. Đang chạy vẫn bấm Dừng được, phần đã sinh vẫn giữ.


┌──────────────────────────────────────────────────────────────────────────┐
│  GIẢI THÍCH TỪNG Ô                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  KHỐI 1 — BA FILE ĐẦU VÀO

    File PDF biểu đồ    Mỗi trang là một biểu đồ kéo
    File Excel số liệu  Đọc sheet "Tong_hop". Phần mềm tự dò tên cột nên
                        thứ tự cột không quan trọng. Cột bắt buộc:
                           FeH_Load_kN, FeL_Load_kN, Fm_Load_kN
                        Cột nên có (để tính A% và Z% cho chuẩn):
                           Break_Position_mm, Break_Top_Load_kN
    File Word mẫu       Để trống → dùng bản có sẵn trong thư mục templates.
                        Chọn file khác khi đổi mẫu báo cáo. File mẫu phải
                        có 2 bảng và có sẵn một ảnh biểu đồ để thay.

  KHỐI 2 — XUẤT BAO NHIÊU FILE

    Toàn bộ                Bao nhiêu biểu đồ ra bấy nhiêu file
    Chỉ xuất ... file      Dùng khi muốn thử trước 5–10 file
    Bắt đầu từ biểu đồ số  Mặc định 1. Đang chạy dở 1000 file mà dừng ở
                           file 400 thì lần sau điền 401 để chạy tiếp

  KHỐI 3 — MÃ MẪU VÀ TÊN FILE

    Lấy mã có sẵn trong Excel   Mã lấy từ cột đầu (M0001, M0002...)
    Tự đặt                      Tiền tố + số bắt đầu + số chữ số
                                   "S2-", 1, 0   →  S2-1, S2-2, S2-3...
                                   "M",   1, 4   →  M0001, M0002...
    Tiêu chuẩn                  In vào dòng Standard của báo cáo

  KHỐI 4 — KÍCH THƯỚC MẪU THỬ VÀ CÁCH TÍNH

    Đường kính danh nghĩa d₀    Gõ 14, 12, 10 hay 8 tuỳ mẫu
    Dung sai gia công ±         Mỗi mẫu lấy một đường kính ngẫu nhiên
                                trong khoảng d₀ ± dung sai. Để 0 nếu muốn
                                mọi mẫu dùng đúng một đường kính
    Chiều dài chuẩn L₀          Cố định theo tiêu chuẩn, mặc định 70 mm
    Chiều dài đo L_e            Cố định theo tiêu chuẩn, mặc định 50 mm

    Số liệu FeH/FeL/Fm lấy từ

       "Số ghi trên biểu đồ PDF"  ← nên chọn cái này
             Đọc thẳng ba con số đã in sẵn trên biểu đồ. Bảng số liệu
             trong báo cáo vì thế LUÔN khớp 100% với biểu đồ in kèm.

       "Cột số liệu trong file Excel"
             Lấy từ Excel. Chỉ nên dùng khi chắc chắn Excel và PDF được
             sinh ra từ cùng một bộ dữ liệu.

       Dù chọn cách nào, phần mềm vẫn đối chiếu chéo hai nguồn và báo ở
       khung Kết quả nếu có mẫu lệch nhau quá 0,5 kN.

    Hệ số quy đổi độ giãn dài   Lu = L₀ + hệ số × quãng đường lúc đứt
    Ứng suất đứt thực (MPa)     Su = lực lúc đứt / ứng suất đứt thực
                                rồi suy ra du và Z%

    Nút "Tự hiệu chuẩn theo file mẫu"
             Phần mềm đọc ngược các số có sẵn trong file Word mẫu
             (L₀, Lu, du) cùng dòng số liệu đầu tiên của Excel để tính ra
             hai hệ số trên. Bấm một lần là xong, không cần hiểu công thức.

    Nếu Excel không có cột lực đứt thì phần mềm chuyển sang sinh Z% ngẫu
    nhiên trong khoảng bạn khai ở hai ô cuối khối.

  KHỐI 5 — ẢNH BIỂU ĐỒ

    Chất lượng ảnh              150 DPI nhẹ file, 200 DPI khuyên dùng,
                                300 DPI khi cần in ấn
    Tự cắt viền trắng           Cắt bỏ lề trắng, biểu đồ chiếm trọn chiều
                                ngang trang Word
    Nén ảnh                     Giảm khoảng một nửa dung lượng, nhìn bằng
                                mắt thường không khác
    Cắt bỏ dòng ghi chú chứa    Biểu đồ mẫu có dòng "Chart ID: 0001" ở
                                chân trang. Phần mềm cắt bỏ trước khi chèn
                                vào báo cáo. Để trống ô này nếu muốn giữ


┌──────────────────────────────────────────────────────────────────────────┐
│  CÔNG THỨC BƯỚC 2 DÙNG                                                   │
└──────────────────────────────────────────────────────────────────────────┘

     S₀  = π·d²/4                        d lấy ngẫu nhiên quanh d₀
     ReH = FeH × 1000 / S₀    (MPa)
     ReL = FeL × 1000 / S₀    (MPa)
     Rm  = Fm  × 1000 / S₀    (MPa)
     Lu  = L₀ + hệ số × quãng đường lúc đứt
     A   = (Lu − L₀) / L₀ × 100    (%)
     Su  = lực lúc đứt × 1000 / ứng suất đứt thực
     du  = căn(4·Su/π)
     Z   = (S₀ − Su) / S₀ × 100    (%)


┌──────────────────────────────────────────────────────────────────────────┐
│  LỖI THƯỜNG GẶP Ở BƯỚC 2                                                 │
└──────────────────────────────────────────────────────────────────────────┘

  "Thiếu thư viện đọc PDF (pypdfium2)"
        → Máy chưa cài thư viện đọc PDF. Đóng phần mềm, chạy lại
          CHAY_PHAN_MEM.bat — nó tự cài. Hoặc mở Command Prompt gõ:
              python -m pip install pypdfium2 pillow

  "Sheet Tong_hop thiếu cột bắt buộc: ..."
        → File Excel không có cột FeH_Load_kN / FeL_Load_kN / Fm_Load_kN.
          Kiểm tra lại tên cột ở dòng đầu tiên của sheet.

  "PDF có 1000 biểu đồ nhưng Excel có 998 dòng số liệu"
        → Hai file lệch nhau. Phần mềm vẫn chạy với 998 cặp đầu, nhưng
          nên kiểm tra lại nguồn dữ liệu.

  "Có N biểu đồ lệch số liệu giữa PDF và Excel"
        → Hai file được sinh ra từ hai lần chạy khác nhau. Nếu đang để
          nguồn là "Số ghi trên biểu đồ PDF" thì báo cáo vẫn đúng, đây
          chỉ là lời nhắc kiểm tra lại file Excel.

  "Không nhận dạng được ô ... trong file Word mẫu"
        → File Word mẫu đã bị sửa cấu trúc bảng. Để trống ô File Word
          mẫu để dùng lại bản gốc templates/TPL_TensileTest.docx.
╔══════════════════════════════════════════════════════════════════════════╗
║  TAB 3 — KÍCH THƯỚC                                                      ║
╚══════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────┐
│  DÙNG TRONG 4 BƯỚC                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  1.  File dung sai và file mẫu báo cáo JIS.QT.07.BM.01 có sẵn trong thư
      mục file_mau/ — phần mềm tự nạp khi mở. Muốn dùng file khác thì
      bấm  📂 Chọn...
  2.  Chọn Loại sản phẩm (ALKTOP II / ISHIZUE / ISHIZUE OMEGA) và Mã.
      Bảng chỉ tiêu hiện ra kèm Min/Max.
  3.  Double-click ô Min hoặc Max để sửa nếu cần.
  4.  Bấm  ▶ SINH DỮ LIỆU & XUẤT BÁO CÁO.

  Kết quả:  output/KichThuoc_<ngày giờ>/BaoCao_SPC_KichThuoc_<loại>_<mã>.xlsx

  Bộ cấu hình của tab này lưu cả loại, mã sản phẩm và các Min/Max đã sửa.

  File xuất ra giữ đúng form biểu mẫu JIS.QT.07.BM.01:
    · Số điểm đo, ký hiệu từng điểm, yêu cầu kỹ thuật lấy thẳng từ file
      dung sai. Tên chỉ tiêu: ISHIZUE dùng tên tiếng Việt của biểu mẫu,
      ALKTOP II / OMEGA giữ nguyên tên trong file dung sai.
    · Khối tính toán dùng đúng công thức của biểu mẫu; ALKTOP (22 chỉ tiêu),
      OMEGA (17) thì khối tự dài ra.
    · Phần mềm chấm SPC đúng như công thức trong file (σ làm tròn 2 số,
      giới hạn làm tròn 1 số) — báo Đạt thì mở bằng Excel cũng Đạt.
    · Sheet Báo cáo trình bày chỉ tiêu 1, như biểu mẫu gốc. Cần báo cáo
      cho chỉ tiêu khác thì dùng bước ② Đánh giá SPC (xem dưới).
    · Mã còn thiếu Min/Max trong file dung sai: phần mềm không cho chạy,
      phải double-click nhập Min/Max trước.
    · Mở file bằng Excel để xem số liệu: Excel tự tính lại khi mở.


╔══════════════════════════════════════════════════════════════════════════╗
║  TAB 4 — THÀNH PHẦN HOÁ                                                  ║
╚══════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────┐
│  DÙNG TRONG 4 BƯỚC                                                       │
└──────────────────────────────────────────────────────────────────────────┘

  1.  File dữ liệu chuẩn (Dữ liệu 2.xlsx), hai file mẫu SPC (SCW410,
      SCW480), file mẫu Kết quả in và file mẫu Tổng hợp có sẵn trong thư
      mục file_mau/ThanhPhan/ — phần mềm tự nạp khi mở. Muốn dùng file
      khác thì bấm  📂 Chọn...
  2.  Chọn Mác thép. Bảng nguyên tố hiện Yêu cầu / Min / Max lấy từ file
      chuẩn, ô File mẫu SPC tự đổi sang đúng mẫu của mác đó.
      Double-click ô Min hoặc Max để sửa nếu cần (không ghi ngược vào
      file chuẩn).
  3.  Điền Số mẫu, STT bắt đầu, Cpk tối thiểu (mặc định 1,33). User,
      Project bắt đầu, Ngày thử, Giờ bắt đầu in ở chân từng kết quả in:
      Project tăng 1 sau mỗi mẫu, giờ tăng vài phút.
  4.  Bấm  ▶ SINH & XUẤT.

  Kết quả:  output/ThanhPhan_<mác>_<ngày giờ>_<N>mau/
     ├─  TongHop_ThanhPhan.xlsx         bảng tổng hợp N mẫu, Ceq giữ công thức
     ├─  KetQuaIn/KetQuaIn_<STT>.docx   kết quả in máy quang phổ từng mẫu:
     │                                  3 lần đo, Average, RSD
     └─  DanhGia_SPC_ThanhPhan_<mác>_01_mau1-25.xlsx …
                                        mỗi file 25 mẫu (60 mẫu → 25 + 25 + 10)

  · Average trên kết quả in trùng đúng số trong bảng tổng hợp và file SPC.
  · Nhóm 25 mẫu nào cũng phải đạt SPC. Chưa đạt thì phần mềm không xuất
    file nào và báo rõ nguyên tố, mục bị lỗi.
  · Ngoài 10 luật, số sinh ra còn tránh những hình trông như không đạt: hai
    mẫu liền nhau trùng số (điểm MR chạm 0), 9 điểm MR liên tiếp cùng phía
    đường tâm, điểm sát giới hạn kiểm soát. Biểu đồ I có thêm đường tâm X̄
    (nét đen) mà file mẫu gốc thiếu.
  · Số mẫu chia 25 dư 1 (ví dụ 26): mẫu cuối không có file SPC, vì biểu đồ
    I-MR cần ít nhất 2 điểm.
  · Mục 6 (Ca, Cp, Cpk): số tính sẵn nằm ở bảng bên phải, ngoài vùng in.
    Bảng trên trang in để trống để chép tay sau khi in.
  · Mở file SPC bằng Excel để xem số liệu: Excel tự tính lại khi mở. Một
    số công thức lỗi của file mẫu SPC được sửa trên bản xuất ra; file mẫu
    gốc giữ nguyên.
  · Bộ cấu hình của tab này lưu cả mác thép và các Min/Max đã sửa.

┌──────────────────────────────────────────────────────────────────────────┐
│  ② ĐÁNH GIÁ SPC                                                          │
└──────────────────────────────────────────────────────────────────────────┘

  Sinh xong ở ① thì thư mục vừa sinh tự nạp sang đây. Cũng chọn được thư
  mục ThanhPhan_… cũ, hoặc chọn thẳng file Tổng hợp (kể cả file phòng thí
  nghiệm tự điền: dòng tiêu đề có STT ở cột A, mỗi cột một nguyên tố).

  1.  Mác thép tự nhận theo tên thư mục. Yêu cầu, Min/Max lấy theo bảng
      ở ① (cùng mác) hoặc theo file dữ liệu chuẩn.
  2.  Mẫu bắt đầu mặc định 1, 26, 51… (nhóm 25 như lúc sinh). Ghi số
      khác để chấm dãy khác; Bước nhảy, Số điểm (≤ 25). Mẫu bắt đầu tính
      theo thứ tự dòng trong Tổng hợp.
  3.  Bảng chấm hiện từng dãy × 10 nguyên tố (C Si Mn P S Ni Cr Mo V Cu):
      X̄, σ, LCLI/UCLI, Cpk, kết luận; dòng chưa đạt tô đỏ, bấm vào dòng
      để xem chi tiết từng luật. ⚡ Quét: điểm bắt đầu nào đạt cả 10.
  4.  🛠 Sửa số cho đạt: chỉ đổi số của nguyên tố chưa đạt trong dãy chưa
      đạt (dãy đang chọn + dãy của các file SPC đang có trong thư mục).
      Số mới nằm trong Min/Max, đúng số chữ số; Fe tính lại, Ceq giữ
      trong yêu cầu. Hỏi xác nhận rồi GHI ĐÈ (không giữ bản cũ):
         TongHop_ThanhPhan.xlsx, KetQuaIn của mẫu bị đổi (3 lần đo mới,
         trung bình = số mới, giữ User/Project/giờ), các file SPC có mẫu
         bị đổi (xuất lại). Nhật ký ghi từng số cũ → số mới.
  5.  ▶ TẠO BÁO CÁO SPC: mỗi dãy một file
         DanhGia_SPC_ThanhPhan_<mác>_bd<mẫu bắt đầu>_buoc<b>_<n>mau.xlsx
      trong chính thư mục đó, cùng bộ xuất với ① (đã sửa lỗi file mẫu).
      Dãy chưa đạt vẫn xuất, tick đúng "Không đạt".
  · Thông tin in (Người lập, Người kiểm tra, Phê duyệt, Tên sản phẩm,
    Tiêu chuẩn) ghi lên 8 sheet nguyên tố; ô để trống giữ chữ của file mẫu.
    Dữ liệu đánh giá để trống → tự ghi "Mẫu STT a–b".


┌──────────────────────────────────────────────────────────────────────────┐
│  LỖI THƯỜNG GẶP                                                          │
└──────────────────────────────────────────────────────────────────────────┘

  "Không sinh được bộ số liệu đạt SPC sau ... lần thử"
        → Khoảng Min/Max quá hẹp. Nới Min/Max hoặc giảm Cpk tối thiểu.

  "Ceq không thể đạt yêu cầu ..."  /  "... vẫn có mẫu Ceq vượt yêu cầu"
        → Giảm Min/Max của C, Mn, Cr… để Ceq lọt yêu cầu kỹ thuật.

  "...: Max ... vượt yêu cầu ..."  /  "...: Min (...) lớn hơn Max (...)"
        → Ô Min/Max vừa sửa sai. Double-click sửa lại, hoặc chọn lại mác
          để lấy lại số trong file chuẩn.


╔══════════════════════════════════════════════════════════════════════════╗
║  ĐÁNH GIÁ SPC — BƯỚC CUỐI TRONG TỪNG TAB                                 ║
╚══════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────────┐
│  VA ĐẬP (TAB 1 · ②)  VÀ  THỬ KÉO (TAB 2 · ③)                             │
└──────────────────────────────────────────────────────────────────────────┘

  1.  Thư mục vừa sinh đã tự nạp sẵn. Muốn đánh giá lần chạy khác thì chọn
      thư mục (Charpy_… ở tab Va đập; Keo_Sinh_… hoặc Keo_… ở tab Thử kéo)
      hoặc file bảng tổng hợp .xlsx, rồi bấm  📥 Nạp.
      Thư mục Keo_… (báo cáo Word) chỉ điền sẵn, chưa đọc — bấm 📥 Nạp.
  2.  Tick cột cần đánh giá, khai mẫu bắt đầu, bước nhảy, số điểm (25).
      Bấm  🔍 Xem trước  để soi một dãy, hoặc  ⚡ Quét mọi điểm bắt đầu
      để xem dãy nào đạt, dãy nào không.
  3.  Điền thông tin in trên biểu mẫu rồi bấm  ▶ TẠO BÁO CÁO SPC.
      File ghi thẳng vào chính thư mục đã chọn.

┌──────────────────────────────────────────────────────────────────────────┐
│  KÍCH THƯỚC (TAB 3 · ②)                                                  │
└──────────────────────────────────────────────────────────────────────────┘

  1.  File báo cáo vừa sinh đã tự nạp; bảng hiện từng chỉ tiêu với X̄, σ,
      LCLI/UCLI, Cpk và kết luận (chưa đạt thì ghi rõ vi phạm mục nào).
      Bấm vào một dòng để xem chi tiết.
  2.  Chọn các chỉ tiêu cần báo cáo (Ctrl/Shift để chọn nhiều, hoặc
      ☑ Chọn tất cả) rồi bấm  ▶ XUẤT BÁO CÁO CHO CHỈ TIÊU ĐÃ CHỌN.
      Mỗi chỉ tiêu ra một file …__BaoCao_STT05_t1.xlsx cạnh file gốc —
      cùng số liệu, sheet Báo cáo (biểu đồ I-MR) trình bày chỉ tiêu đó.

  Kết quả chấm được báo trung thực: không đạt thì liệt kê rõ từng mục.
"""

    def _tab_huong_dan(self, me):
        khung = ttk.Frame(me, padding=(12, 10, 12, 10))
        khung.pack(fill="both", expand=True)

        card = ui.ModernCard(khung, title="Tài liệu & Cẩm nang hướng dẫn sử dụng phần mềm", badge_num="i")
        card.pack(fill="both", expand=True)
        cb = card.get_body()

        t = scrolledtext.ScrolledText(cb, font=("Segoe UI", 9), wrap="word",
                                      bg="#FFFFFF", fg=ui.COLOR_TEXT_MAIN,
                                      relief="flat", borderwidth=0)
        t.pack(fill="both", expand=True)
        t.insert("1.0", self.HUONG_DAN)
        t.configure(state="disabled")

    def _o(self, me, r, nhan, var, goi_y):
        ttk.Label(me, text=nhan, style="Surface.TLabel").grid(row=r * 2, column=0, sticky="w", pady=(3, 0))
        ttk.Entry(me, textvariable=var).grid(row=r * 2 + 1, column=0, sticky="ew",
                                            ipady=2)
        if goi_y:
            ttk.Label(me, text=goi_y, style="SurfacePhu.TLabel").grid(
                row=r * 2 + 1, column=1, sticky="w", padx=(8, 0))

    def _ghi(self, s, xoa=False):
        self.log.ghi(s, xoa=xoa)

    # ------------------------------------------------------------ doc tham so
    def _thu_cfg(self):
        def so(v, ten, nguyen=False):
            v = str(v).strip().replace(",", ".")
            if not v:
                raise core.LoiNghiepVu("Chưa nhập: %s" % ten)
            try:
                return core.spc_luat.so_nguyen(v, ten, 0) if nguyen else core.spc_luat.so_huu_han(v, ten)
            except ValueError:
                raise core.LoiNghiepVu("Ô '%s' phải là số. Đang là: %s" % (ten, v))

        def danh_sach(v, ten):
            xs = [x.strip().replace(",", ".") for x in
                  str(v).replace(";", " ").replace(",", " ").split()]
            try:
                ra = [core.spc_luat.so_huu_han(x, ten) for x in xs if x]
            except ValueError:
                raise core.LoiNghiepVu(
                    "Ô '%s' chỉ được chứa số, cách nhau bằng dấu cách hoặc phẩy."
                    % ten)
            if not ra:
                raise core.LoiNghiepVu("Chưa nhập: %s" % ten)
            return ra

        chung = bool(self.v_chung.get())
        lan = []
        for i in range(3):
            k = 0 if chung else i
            ten = "Lần %d" % (k + 1)
            lan.append({
                "e_min": so(self.v_lan[k][0].get(), "%s — Energy từ" % ten),
                "e_max": so(self.v_lan[k][1].get(), "%s — Energy đến" % ten),
                "A": danh_sach(self.v_lan[k][2].get(), "%s — các giá trị A" % ten),
                "B": danh_sach(self.v_lan[k][3].get(), "%s — các giá trị B" % ten),
            })
            if lan[-1]["e_min"] > lan[-1]["e_max"]:
                raise core.LoiNghiepVu(
                    "%s: Energy nhỏ nhất đang lớn hơn Energy lớn nhất." % ten)

        seed = self.v_seed.get().strip()
        return {
            "so_luong": so(self.v_n.get(), "Số lượng mẫu", True),
            "stt_bat_dau": so(self.v_stt.get(), "Số thứ tự bắt đầu", True),
            "tien_to_ma": self.v_tt.get().strip(),
            "so_chu_so": so(self.v_cs.get() or 0, "Số chữ số của mã", True),
            "cau_hinh_lan": lan,
            "e_min": lan[0]["e_min"], "e_max": lan[0]["e_max"],
            "danh_sach_A": lan[0]["A"], "danh_sach_B": lan[0]["B"],
            "nhiet_do_thu": so(self.v_nd.get(), "Nhiệt độ thử"),
            "seed": int(seed) if seed.isdigit() else None,
            "ngay_thu": self.v_ngay.get().strip(),
            "nguoi_thu": self.v_tester.get().strip(),
            "nguoi_duyet": self.v_duyet.get().strip(),
            "tieu_chuan": self.v_tc.get().strip(),
            "vat_lieu": self.v_vl.get().strip(),
            "kich_thuoc_mau": self.v_kt.get().strip(),
            "sinh_word": bool(self.v_word.get()),
            "sinh_tong_hop": bool(self.v_th.get()),
            "sinh_tong_hop_gon": bool(self.v_gon.get()),
            "sinh_spc": bool(self.v_spc.get()),
            "file_mau_spc": self.v_file_mau.get().strip() or None,
            "lsl": so(self.v_lsl.get(), "Giới hạn dưới LSL"),
            "cpk_toi_thieu": so(self.v_cpk.get(), "Cpk tối thiểu"),
        }

    def doi_dung_chung(self):
        """Bat/tat hai dong duoi khi tick 'dung chung cho ca 3 lan'."""
        chung = bool(self.v_chung.get())
        for i in (1, 2):
            for j in range(4):
                try:
                    self.e_lan[i][j].configure(
                        state="disabled" if chung else "normal")
                except Exception:
                    pass
                if chung:
                    self.v_lan[i][j].set(self.v_lan[0][j].get())

    def chon_file_mau(self):
        p = filedialog.askopenfilename(
            title="Chọn file mẫu đánh giá độ ổn định SPC",
            filetypes=[("File Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not p:
            return
        self._doc_file_mau(p)

    def _doc_file_mau(self, p, im_lang=False):
        self.v_file_mau.set(p)
        try:
            d = core.doc_cau_hinh_tu_file_mau(p)
            self.v_lsl.set("%g" % d["lsl"])
            self.lbl_spc.configure(
                text="Đọc từ sheet “%s”:  vật liệu %s  ·  sản phẩm %s  ·  "
                     "LSL %g J  ·  USL %s"
                     % (d["sheet_da_doc"], d.get("vat_lieu_spc"),
                        d.get("ten_san_pham"), d["lsl"],
                        ("%g J" % d["usl"]) if d.get("usl") else "không có"))
        except Exception as e:
            self.lbl_spc.configure(text="Không đọc được file mẫu: %s" % e)
            if not im_lang:
                messagebox.showerror("Không đọc được file mẫu", str(e))

    # ---------------------------------------------------------- bo cau hinh
    KHONG_LUU = ("v_ngay",)     # ngay thu doi theo tung lan

    def lay_cau_hinh(self):
        return bo_cau_hinh.chup_bien(self, bo_qua=self.KHONG_LUU)

    def dat_cau_hinh(self, d):
        bo_cau_hinh.dat_bien(self, d, bo_qua=self.KHONG_LUU + ("v_lsl",))
        self.doi_dung_chung()
        p = self.v_file_mau.get().strip()
        if p and os.path.exists(p):
            self._doc_file_mau(p, im_lang=True)
        if d.get("v_lsl") not in (None, ""):
            self.v_lsl.set(d["v_lsl"])

    # ----------------------------------------------------------------- chay
    def chay(self):
        try:
            cfg = self._thu_cfg()
        except core.LoiNghiepVu as e:
            messagebox.showwarning("Thiếu hoặc sai tham số", str(e))
            return
        if not (cfg["sinh_word"] or cfg["sinh_tong_hop"] or cfg["sinh_spc"]
                or cfg["sinh_tong_hop_gon"]):
            messagebox.showwarning("Chưa chọn gì",
                                   "Hãy tick ít nhất một loại file muốn sinh.")
            return
        if cfg["sinh_spc"]:
            cb = core.canh_bao_khoang(dict(core.CFG_MAC_DINH, **cfg))
            if cb and not messagebox.askyesno(
                    "Khoảng Energy hơi rộng so với ngưỡng",
                    cb + "\n\nPhần mềm vẫn chạy được — nó sẽ sinh lại nhiều lần "
                         "cho tới khi đạt, nhưng dữ liệu sẽ nghiêng về phía ít "
                         "dao động.\n\nVẫn chạy?"):
                return
        if cfg["so_luong"] > 3000 and cfg["sinh_word"]:
            if not messagebox.askyesno(
                    "Số lượng lớn",
                    "Bạn đang sinh %s file Word. Việc này mất khoảng %d phút "
                    "và chiếm chừng %d MB đĩa.\n\nTiếp tục?"
                    % ("{:,}".format(cfg["so_luong"]),
                       max(1, cfg["so_luong"] // 2000),
                       cfg["so_luong"] * 33 // 1000)):
                return

        self.btn.configure(state="disabled", text="Đang sinh...")
        self.tien_do.configure(maximum=cfg["so_luong"], value=0)
        self.lbl_td.configure(text="")
        self._ghi("Đang sinh %s mẫu...\n" % "{:,}".format(cfg["so_luong"]), xoa=True)
        threading.Thread(target=self._chay_nen, args=(cfg,), daemon=True).start()
        self.root.after(120, self._kiem_tra_hang_doi)

    def _chay_nen(self, cfg):
        try:
            def td(i, n):
                self.hang_doi.put(("tien_do", (i, n)))
            kq = core.chay(cfg, tien_do=td)
            if cfg.get("sinh_spc") and kq.get("thu_muc"):
                import spc_bao_cao
                spc_bao_cao.lam_sach_tick_thu_muc(kq["thu_muc"])
            self._luu_lan_truoc(cfg)
            self.hang_doi.put(("xong", kq))
        except core.LoiNghiepVu as e:
            self.hang_doi.put(("loi", str(e)))
        except PermissionError as e:
            self.hang_doi.put(("loi",
                               "Không ghi được file kết quả.\n\nThường do file "
                               "đang mở, hoặc thư mục bị khoá quyền ghi.\n\n%s" % e))
        except Exception:
            self.hang_doi.put(("loi_he_thong", traceback.format_exc()))

    def _kiem_tra_hang_doi(self, lap=True):
        try:
            while True:
                loai, dl = self.hang_doi.get_nowait()
                if loai == "tien_do":
                    i, n = dl
                    self.tien_do.configure(value=i)
                    self.lbl_td.configure(text="%s / %s file"
                                          % ("{:,}".format(i), "{:,}".format(n)))
                elif loai == "xong":
                    self._xong(dl)
                elif loai == "loi":
                    self._loi(dl)
                elif loai == "loi_he_thong":
                    self._loi_he_thong(dl)
        except queue.Empty:
            pass
        if lap:
            self.root.after(120, self._kiem_tra_hang_doi)

    def _xong(self, kq):
        self.btn.configure(state="normal", text="▶   SINH DỮ LIỆU")
        self.thu_muc_cuoi = kq["thu_muc"]
        try:                       # nap san thu muc vua sinh vao buoc Danh gia
            co_bang = os.path.exists(os.path.join(kq["thu_muc"], "TongHop_VaDap.xlsx"))
            self.tab_spc_charpy.dat_san(kq["thu_muc"], nap=co_bang)
        except Exception:
            pass
        self.tien_do.configure(value=self.tien_do["maximum"])
        self._ghi(core.bao_cao(kq) + "\n", xoa=True)
        if messagebox.askyesno("Xong",
                               "Đã sinh xong %s mẫu trong %.1f giây.\n\n"
                               "Mở thư mục kết quả bây giờ?"
                               % ("{:,}".format(len(kq["ds"])), kq["giay"])):
            self.mo_thu_muc()

    def _loi(self, msg):
        self.btn.configure(state="normal", text="▶   SINH DỮ LIỆU")
        self._ghi("KHÔNG SINH ĐƯỢC:\n\n" + msg + "\n", xoa=True)
        messagebox.showerror("Không sinh được", msg)

    def _loi_he_thong(self, tb):
        self.btn.configure(state="normal", text="▶   SINH DỮ LIỆU")
        self._ghi("LỖI KHÔNG MONG ĐỢI:\n\n" + tb, xoa=True)
        messagebox.showerror("Lỗi", "Có lỗi ngoài dự kiến. Chi tiết ở khung dưới.")

    def mo_thu_muc(self):
        d = self.thu_muc_cuoi
        if not (d and os.path.isdir(d)):
            d = os.path.join(core.base_dir(), "output")
            os.makedirs(d, exist_ok=True)
        mo_bang_windows(d)

    # -------------------------------------------------------------- ghi nho
    def _luu_lan_truoc(self, cfg):
        bo_cau_hinh.nho_lan_truoc("charpy", cfg)

    def _nap_lan_truoc(self):
        c = bo_cau_hinh.doc_lan_truoc("charpy", "lan_dung_truoc.json")
        p = (c or {}).get("file_mau_spc") or ""
        if not (p and os.path.exists(p)):
            # may moi: lay file mau SPC giao kem trong thu muc file_mau/
            p = bo_cau_hinh.tim_file_mau("Va đập", "Va dap")
        if p:
            self._doc_file_mau(p, im_lang=True)
        if not c:
            return
        cap = [("so_luong", self.v_n), ("stt_bat_dau", self.v_stt),
               ("tien_to_ma", self.v_tt), ("so_chu_so", self.v_cs),
               ("nhiet_do_thu", self.v_nd), ("nguoi_thu", self.v_tester),
               ("nguoi_duyet", self.v_duyet), ("tieu_chuan", self.v_tc),
               ("vat_lieu", self.v_vl), ("kich_thuoc_mau", self.v_kt),
               ("lsl", self.v_lsl), ("cpk_toi_thieu", self.v_cpk)]
        for k, var in cap:
            if c.get(k) not in (None, ""):
                v = c[k]
                var.set(("%g" % v) if isinstance(v, float) else str(v))

        lan = c.get("cau_hinh_lan")
        if lan and len(lan) == 3:
            for i, x in enumerate(lan):
                self.v_lan[i][0].set("%g" % x["e_min"])
                self.v_lan[i][1].set("%g" % x["e_max"])
                self.v_lan[i][2].set(", ".join("%g" % v for v in x["A"]))
                self.v_lan[i][3].set(", ".join("%g" % v for v in x["B"]))
            khac = any(lan[i] != lan[0] for i in (1, 2))
            self.v_chung.set(not khac)
            self.doi_dung_chung()
        self._ghi("(Đã nhớ lại tham số của lần dùng trước.)\n")


def main():
    root = tk.Tk()
    App(root)
    cb = bo_cau_hinh.canh_bao_duong_dan_dai()
    if cb:
        root.after(400, lambda: messagebox.showwarning("Thư mục phần mềm nằm quá sâu", cb))
    root.mainloop()


if __name__ == "__main__":
    main()
