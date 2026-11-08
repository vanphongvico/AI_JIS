# -*- coding: utf-8 -*-
"""
TAO BAO CAO DANH GIA DO ON DINH SPC TU DU LIEU DA CO
=====================================================
Ghep doc_thu_muc (doc du lieu) voi spc_luat (cham) va spc (ghi file mau).

Khac voi Tab 1 va Tab 3 - hai tab do SINH du lieu roi moi danh gia - module
nay lam viec tren du lieu DA CO: chon mot day 25 mau theo (mau bat dau, buoc
nhay), cham, roi xuat file theo dung bieu mau TPL_SPC.xlsx.

Ket qua cham duoc bao TRUNG THUC. Neu day chon ra khong dat thi file van
duoc sinh nhung ghi ro dat hay khong, va vi pham nhung muc nao.
"""

import os
import re

import doc_thu_muc as dtm
import spc_luat
MAC_DINH = object()


def ten_file(ten_cot, bat_dau, buoc, so_diem):
    an = re.sub(r"[^A-Za-z0-9]+", "_", ten_cot).strip("_")
    return "DanhGia_DoOnDinh_SPC_%s_bd%d_buoc%d_%dmau.xlsx" % (
        an, bat_dau, buoc, so_diem)


def cham_day(d, ten_cot, bat_dau, buoc, so_diem=25, lsl=MAC_DINH, usl=None,
             cpk_toi_thieu=1.33):
    """Lay day con roi cham. Tra ve (chi_so, gia_tri, ket_qua_spc)."""
    idx, gt = dtm.lay_day(d, ten_cot, bat_dau, buoc, so_diem)
    if lsl is MAC_DINH:
        lsl = d["nguong"].get(ten_cot)
    try:
        r = spc_luat.cham(gt, lsl, usl, cpk_toi_thieu)
    except ValueError as e:
        raise dtm.LoiDoc(str(e))
    return idx, gt, r


def quet_diem_bat_dau(d, ten_cot, buoc, so_diem=25, lsl=MAC_DINH, usl=None,
                      cpk_toi_thieu=1.33, toi_da=None):
    """Thu MOI diem bat dau hop le voi buoc nhay dang chon.

    Tra ve list (mau_bat_dau, dat, danh_sach_ma_loi). Dung de nguoi dung
    thay ngay day nao dat, day nao khong - khong phai mo tung cai.
    """
    try:
        buoc = spc_luat.so_nguyen(buoc, "Buoc nhay")
        so_diem = spc_luat.so_nguyen(so_diem, "So diem", 2)
    except ValueError as e:
        raise dtm.LoiDoc(str(e))
    het = d["so_mau"] - (so_diem - 1) * buoc
    if het < 1:
        raise dtm.LoiDoc(
            "Voi buoc nhay %d va %d diem thi can it nhat %d mau, du lieu chi "
            "co %d." % (buoc, so_diem, (so_diem - 1) * buoc + 1, d["so_mau"]))
    if lsl is MAC_DINH:
        lsl = d["nguong"].get(ten_cot)
    v = d["cot"][ten_cot]
    ra = []
    for s in range(1, min(het, toi_da or het) + 1):
        gt = [v[s - 1 + i * buoc] for i in range(so_diem)]
        try:
            r = spc_luat.cham(gt, lsl, usl, cpk_toi_thieu)
            ra.append((s, r.dat, list(r.loi)))
        except ValueError:
            ra.append((s, False, ["DU_LIEU"]))
    return ra


def sinh_bao_cao(d, ten_cot, bat_dau, buoc, thu_muc, cfg=None, so_diem=25,
                 lsl=MAC_DINH, usl=None, cpk_toi_thieu=1.33):
    """Sinh MOT file bao cao SPC. Tra ve (duong_dan, ket_qua_cham)."""
    import spc as _spc

    idx, gt, r = cham_day(d, ten_cot, bat_dau, buoc, so_diem, lsl, usl,
                          cpk_toi_thieu)
    nhan = dict((k, (u, nh)) for k, u, nh in dtm.danh_sach_cot(d))
    don_vi, ten_chi_tieu = nhan.get(ten_cot, ("", ten_cot))

    c = dict(cfg or {})
    c["lsl"] = r.lsl
    c["cpk_toi_thieu"] = cpk_toi_thieu
    c["usl"] = usl
    c["chi_tieu_danh_gia"] = c.get("chi_tieu_danh_gia") or ten_chi_tieu
    c["don_vi"] = c.get("don_vi") or don_vi
    for k, v in (d.get("thong_tin") or {}).items():
        c.setdefault(k, v)

    os.makedirs(thu_muc, exist_ok=True)
    p = os.path.join(thu_muc, ten_file(ten_cot, bat_dau, buoc, so_diem))
    _spc.sinh_bao_cao_spc(gt, c, p, {
        "n": r.n, "X": r.X, "MR": r.MR, "xb": r.xb, "mrb": r.mrb,
        "sigma": r.sigma, "ucl_i": r.ucl_i, "lcl_i": r.lcl_i,
        "ucl_mr": r.ucl_mr, "lcl_mr": 0.0, "lsl": c["lsl"], "usl": usl,
        "cpk": r.cpk,
        "cp": r.cp, "ca": r.ca})
    import spc_bao_cao
    spc_bao_cao.lam_sach_tick_file(p)
    return p, idx, gt, r


def sinh_nhieu(d, cac_cot, cac_bat_dau, buoc, thu_muc, cfg=None, so_diem=25,
               nguong=None, cpk_toi_thieu=1.33, tien_do=None):
    """Sinh nhieu file mot luot: moi (cot x diem bat dau) mot file.

    Tra ve list dict, moi phan tu la mot file da sinh kem ket qua cham.
    """
    nguong = nguong or {}
    ra = []
    tong = len(cac_cot) * len(cac_bat_dau)
    i = 0
    for ten_cot in cac_cot:
        for bd in cac_bat_dau:
            p, idx, gt, r = sinh_bao_cao(
                d, ten_cot, bd, buoc, thu_muc, cfg, so_diem,
                nguong.get(ten_cot, d["nguong"].get(ten_cot)), None,
                cpk_toi_thieu)
            ra.append({"file": p, "cot": ten_cot, "bat_dau": bd, "buoc": buoc,
                       "chi_so": idx, "gia_tri": gt, "spc": r})
            i += 1
            if tien_do:
                tien_do(i, tong)
    return ra


# ------------------------------------------------ thu keo: bieu mau theo mac
SO_DIEM_BIEU_MAU_KEO = 25


def cot_bieu_mau_keo(d, chay=None):
    """Cot du lieu dua vao ba chi tieu cua bieu mau thu keo.

    Tra ve {"ReH": ten cot nguon, "Rm": "Rm", "A": "A"}. Hang gioi han chay
    lay `chay` neu khai ("ReH"/"ReL") va du lieu co cot do; khong thi lay
    ReH, du lieu khong co ReH thi lay ReL. Thieu chi tieu nao -> LoiDoc.
    """
    uu_tien = (chay,) if chay in ("ReH", "ReL") else ()
    chay = next((k for k in uu_tien + ("ReH", "ReL") if k in d["cot"]), None)
    thieu = [t for t, k in (("giới hạn chảy (ReH/ReL)", chay),
                            ("Rm", "Rm" if "Rm" in d["cot"] else None),
                            ("A", "A" if "A" in d["cot"] else None)) if not k]
    if thieu:
        raise dtm.LoiDoc(
            "Biểu mẫu SPC thử kéo cần đủ ba chỉ tiêu giới hạn chảy, Rm và A. "
            "Dữ liệu đang thiếu: %s." % ", ".join(thieu))
    return {"ReH": chay, "Rm": "Rm", "A": "A"}


def ten_file_keo(mac, bat_dau, buoc):
    return "DanhGia_DoOnDinh_SPC_Keo_%s_bd%d_buoc%d_%dmau.xlsx" % (
        re.sub(r"[^A-Za-z0-9]+", "", mac), bat_dau, buoc, SO_DIEM_BIEU_MAU_KEO)


def sinh_bieu_mau_keo(d, mac, cac_bat_dau, buoc, thu_muc, cfg=None,
                      nguong=None, cpk_toi_thieu=1.33, tien_do=None,
                      chay=None):
    """Thu keo: moi diem bat dau MOT file theo bieu mau cua mac vat lieu
    (SCW480 / SCW410), gom ca ba chi tieu gioi han chay - Rm - A.

    nguong: LSL theo ten cot nguon; cot nao khong khai thi lay nguong cua mac.
    Tra ve list dict cung dang voi sinh_nhieu (moi chi tieu mot phan tu, ba
    phan tu cua cung mot diem bat dau tro vao cung mot file).
    """
    import xuat_spc_keo

    cot = cot_bieu_mau_keo(d, chay)
    mac_ng = dtm.MAC_KEO.get(mac) or {}
    nguong = nguong or {}
    lsl = {}
    for ct, k in cot.items():
        lsl[ct] = nguong.get(k, mac_ng.get(k))
    os.makedirs(thu_muc, exist_ok=True)
    ra = []
    for i, bd in enumerate(cac_bat_dau):
        p = os.path.join(thu_muc, ten_file_keo(mac, bd, buoc))
        gia_tri, cham = {}, []
        for ct, k in cot.items():
            idx, gt, r = cham_day(d, k, bd, buoc, SO_DIEM_BIEU_MAU_KEO,
                                  lsl[ct], None, cpk_toi_thieu)
            gia_tri[ct] = gt
            cham.append({"file": p, "cot": k, "bat_dau": bd, "buoc": buoc,
                         "chi_so": idx, "gia_tri": gt, "spc": r})
        try:
            xuat_spc_keo.xuat(mac, gia_tri, lsl, dict(cfg or {}, cpk_toi_thieu=cpk_toi_thieu), p, cot["ReH"])
        except xuat_spc_keo.LoiMau as e:
            raise dtm.LoiDoc(str(e))
        ra += cham
        if tien_do:
            tien_do(i + 1, len(cac_bat_dau))
    return ra


def bao_cao_chu(kq_list):
    """Tom tat dang chu cho khung Ket qua."""
    L = []
    dat = sum(1 for k in kq_list if k["spc"].dat)
    L.append("DA SINH %d FILE BAO CAO SPC  -  CHAM %d DAY: %d dat, %d chua dat"
             % (len(set(k["file"] for k in kq_list)), len(kq_list), dat,
                len(kq_list) - dat))
    L.append("")
    for k in kq_list:
        r = k["spc"]
        L.append("%s" % os.path.basename(k["file"]))
        L.append("   cot %s | mau %s ... %d  (buoc %d)"
                 % (k["cot"], ", ".join(map(str, k["chi_so"][:3])),
                    k["chi_so"][-1], k["buoc"]))
        L.append("   X-bar %.4f   MR-bar %.4f   sigma %.4f%s"
                 % (r.xb, r.mrb, r.sigma,
                    ("   Cpk %.3f" % r.cpk) if r.cpk is not None else ""))
        if r.dat:
            L.append("   KET LUAN: DAT tat ca cac muc danh gia")
        else:
            L.append("   KET LUAN: CHUA DAT - vi pham %d muc:" % len(r.loi))
            for m in r.loi:
                L.append("      [%-3s] %s" % (m, spc_luat.TEN_LUAT.get(m, m)))
        L.append("")
    return "\n".join(L)
