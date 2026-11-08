# -*- coding: utf-8 -*-
"""
TAB THU KEO - BA BUOC TRONG MOT TAB
====================================
    Buoc 1  Sinh bieu do & so lieu   (tab_sinh_keo.TabSinhKeo)
    Buoc 2  Gop bao cao Word         (tab_keo.TabKeo)
    Buoc 3  Danh gia SPC             (tab_bao_cao_spc.TabBaoCaoSPC, loai keo)
    Ve tu so lieu co san             (tab_ve_keo.TabVeKeo) - dung rieng,
                                     khong noi voi ba buoc tren

Sinh xong o Buoc 1 thi file PDF, Excel va kich thuoc mau tu dien sang Buoc 2,
dong thoi thu muc vua sinh tu nap vao Buoc 3. Gop xong o Buoc 2 thi thu muc
ket qua cung tu nap vao Buoc 3 - Buoc 2 ghi kem SoLieu_<N>.xlsx nen doc
nhanh, khong phai boc lai tung file Word.

Mot bo cau hinh cua tab nay luu ca Buoc 1 va Buoc 2; Buoc 3 co bo cau hinh
rieng (thong tin in tren bieu mau SPC).
"""

import tkinter as tk
from tkinter import ttk

import bo_cau_hinh
import ui_theme as ui
from tab_sinh_keo import TabSinhKeo
from tab_keo import TabKeo
from tab_bao_cao_spc import TabBaoCaoSPC
from tab_ve_keo import TabVeKeo

KHOA = "thu_keo"


class TabThuKeo(object):
    def __init__(self, me, mo_thu_muc_fn):
        self.me = me

        self.nb = ttk.Notebook(me)
        self.nb.pack(fill="both", expand=True, padx=4, pady=(6, 0))
        f1 = ttk.Frame(self.nb)
        f2 = ttk.Frame(self.nb)
        f3 = ttk.Frame(self.nb)
        f4 = ttk.Frame(self.nb)
        self.nb.add(f1, text="  ① Bước 1 — Sinh biểu đồ & số liệu  ")
        self.nb.add(f2, text="  ② Bước 2 — Gộp báo cáo Word  ")
        self.nb.add(f3, text="  ③ Đánh giá SPC  ")
        self.nb.add(f4, text="  📈 Vẽ biểu đồ từ số liệu có sẵn  ")

        # Thanh bo cau hinh dat o dau ca hai buoc, dung chung mot danh sach
        self.thanh = []
        self.buoc1 = TabSinhKeo(f1, mo_thu_muc_fn,
                                khi_xong=self._buoc1_xong,
                                sang_buoc2=self.sang_buoc2,
                                dau_trang=self._thanh_cau_hinh)
        self.buoc2 = TabKeo(f2, mo_thu_muc_fn,
                            dau_trang=self._thanh_cau_hinh,
                            khi_xong=self._buoc2_xong)
        self.danh_gia = TabBaoCaoSPC(f3, mo_thu_muc_fn, khoa="spc_keo",
                                     ten="Đánh giá SPC thử kéo", loai="keo")
        self.ve = TabVeKeo(f4, mo_thu_muc_fn)
        self.nb.bind("<<NotebookTabChanged>>", self._dong_bo_thanh)

        bo_cau_hinh.dang_ky_thong_tin(KHOA, "Thử kéo", {
            "tieu_chuan_thu": self.buoc2.v_tc,
        })

    def _thanh_cau_hinh(self, than):
        self.thanh.append(bo_cau_hinh.ThanhCauHinh(
            than, KHOA, "Thử kéo", self.lay_cau_hinh, self.dat_cau_hinh))

    def _dong_bo_thanh(self, _e=None):
        """Hai thanh cung doc mot file - nap lai danh sach khi doi buoc."""
        for t in self.thanh:
            t._nap_ds()

    # ---------------------------------------------------------------
    def _buoc1_xong(self, kq, cfg):
        self.buoc2.nhan_tu_buoc1(kq, cfg)
        self.danh_gia.dat_san(kq["thu_muc"], nap=True)

    def _buoc2_xong(self, thu_muc):
        self.danh_gia.dat_san(thu_muc, nap=True)

    def sang_buoc2(self):
        self.nb.select(1)

    def lay_cau_hinh(self):
        return {"buoc1": self.buoc1.lay_cau_hinh(),
                "buoc2": self.buoc2.lay_cau_hinh()}

    def dat_cau_hinh(self, d):
        self.buoc1.dat_cau_hinh(d.get("buoc1") or {})
        self.buoc2.dat_cau_hinh(d.get("buoc2") or {})
