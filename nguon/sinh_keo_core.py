# -*- coding: utf-8 -*-
"""
SINH BO SO LIEU THU KEO + BIEU DO  (dau vao cho Tab 2)
=======================================================
Sinh ra DONG THOI hai file khop nhau tuyet doi:
    - 1 file PDF N trang, moi trang mot bieu do Load-Position
    - 1 file Excel so lieu (sheet Tong_hop) dung dinh dang Tab 2 doc duoc

DAO CHIEU TINH TOAN so voi tensile_core.py
-------------------------------------------
tensile_core.py di theo chieu:   luc do duoc  ->  ung suat
Module nay di NGUOC LAI:         ung suat     ->  luc

Ly do: nam dai luong phai chiu rang buoc SPC (ReL, ReH, Rm, A, Z) deu la
ung suat va ti le, khong phai luc. Sinh thang chung roi suy nguoc ra luc
thi kiem soat duoc SPC; lam theo chieu cu thi khong.

Chuoi suy dien:
    S0 = pi*d^2/4                     d ngau nhien trong d0 +- dung sai
    FeL = ReL*S0/1000  (kN)           FeH, Fm tuong tu
    Lu  = L0*(1 + A/100)
    Su  = S0*(1 - Z/100)              <-- Z = (S0 - Su)/S0 (JIS Z 2241)
    du  = sqrt(4*Su/pi) = d*sqrt(1 - Z/100)
    luc luc dut (bieu do) = 70-78% Fm  - KHONG phu thuoc Z (xem
                                         tham_so_duong_cong)

Ve cot Z
--------
Z theo dinh nghia JIS Z 2241 (ti le DIEN TICH), dung nhu bao cao mau 522:
    Z = (S0 - Su)/S0 * 100,   S0 = pi*d^2/4,   Su = pi*du^2/4
Ca phan mem dung mot dinh nghia nay: o nhap Z, file Excel, bao cao Word.
Chon d, du lam tron 0,01 mm, tinh lai Z theo hai duong kinh nay TRUOC khi
toi uu SPC. Kich thuoc di cung hang du lieu khi hoan vi; Buoc 2 doc lai
d, du trong Excel se tinh ra cung so Z.

(Cong thuc cu cua bieu mau BM32, Z = (d0 - du)/d0, van con: dat
cfg['z_theo_dien_tich'] = False.)
"""

import os
import sys
import math
import random
from datetime import datetime

import spc_luat

# ---------------------------------------------------------------- mac dinh
CFG_MAC_DINH = {
    "so_luong": 1000,
    "seed": None,

    # --- kich thuoc mau thu (hinh tru 14A) ---
    "d0": 14.0,
    "dung_sai_d": 0.02,
    "Lo": 70.0,
    "Le": 50.0,

    # --- dai sinh cua nam cot chiu SPC ---
    "reL_tb": 372.0, "reL_sd": 4.2,
    "reH_chenh_min": 28.0, "reH_chenh_max": 45.0,
    "rm_tb": 571.0, "rm_sd": 4.5,
    "a_tb": 29.5, "a_sd": 0.85,
    "z_tb": 54.6, "z_sd": 1.3,     # Z theo dien tich (mau 522: 61.0)

    # --- nguong nghiem thu in o dong 11 cua BM32 ---
    "reL_min": 275.0,
    "rm_min": 480.0,
    "a_min": 20.0,
    "reH_min": None,
    "z_min_spec": None,

    # --- danh gia SPC ---
    "cpk_toi_thieu": 1.33,
    # buoc_nhay      : xet tu dau lo (re)
    # buoc_moi_diem  : xet o moi diem bat dau (dat, nhung manh hon han)
    "buoc_nhay": [1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 16, 20, 25, 30, 40],
    "buoc_moi_diem": [1],
    "giay_toi_da": 120,
    "bo_qua_luat": (),
    "z_theo_dien_tich": True,      # Z = (S0 - Su)/S0 nhu JIS Z 2241

    # --- tham so bieu do ---
    "he_so_gian_dai": 0.772,      # Lu = L0 + he_so * quang duong luc dut
    "ung_suat_dut": 1005.0,       # MPa, de suy luc luc dut tu Su
    "ti_le_fm": 0.669,            # fm_pos / dut_pos (mau 522: 20,592/30,79)
    "feh_pos_tb": 3.88, "feh_pos_sd": 0.05,
    # Be rong doan chay deo FeH -> FeL. Mau 522: 3,863 -> 4,876 = 1,013 mm
    # (bo 1000 bieu do cu chi 0,449 mm - kieu do khong con dung).
    "fel_lech_tb": 1.00, "fel_lech_sd": 0.06,
    "dut_duoi_ti_le": 0.84,
    # Luc luc dut tren bieu do = ti le nay x Fm (mau that 522: 58.8/80.6 = 73%)
    "dut_tren_ti_le_min": 0.70, "dut_tren_ti_le_max": 0.78,


    "tien_to_ma": "M",
    "so_chu_so": 4,
    "so_bat_dau": 1,

    # --- thong tin in tren file danh gia do on dinh SPC ---
    "xuat_spc": True,
    "cot_spc": ["ReL"],
    "vat_lieu_spc": "SCW480",
    "tieu_chuan_ap_dung": "JIS G5102 - 1991",
    "ten_san_pham": "ISHIZUE/礎",
    "nguoi_lap": "Vũ Vương Việt",
    "nguoi_kiem_tra": "Đoàn Ngọc Thoạt",
    "nguoi_phe_duyet": "Phạm Văn Phong",
}

COT_SPC = ["ReL", "ReH", "Rm", "A", "Z"]

# ReL / ReH / Rm sinh co 1 chu so thap phan (vd 572.8) - file Excel ghi so
# le nay, con bao cao Word (Buoc 2) in so nguyen lam tron NUA LEN (573).
# Ba cot nay phai dat SPC o CA HAI dang, xem _cot_kep.
SO_LE_UNG_SUAT = 1
COT_LAM_TRON_BAO_CAO = (0, 1, 2)       # chi so cot ReL, ReH, Rm trong COT_SPC


def lam_tron_bao_cao(x):
    """So nguyen in tren bao cao: lam tron NUA LEN, 572.5 -> 573.

    Hai nac: ve 1 chu so truoc roi moi ve so nguyen. Buoc 2 tinh lai ung
    suat tu luc Fm da lam tron 3 chu so nen 572.5 co the thanh 572.4993 -
    lam tron thang se ra 572. round() cua Python cung khong dung duoc: no
    lam tron ve so CHAN (572.5 -> 572).
    """
    if x is None:
        return None
    x1 = round(float(x), 1)
    return int(math.floor(x1 + 0.5 + 1e-9)) if x1 >= 0 else -int(math.floor(-x1 + 0.5 + 1e-9))


def base_dir():
    """Thu muc phan mem. Chay bang .exe thi __file__ nam trong thu muc tam
    cua PyInstaller (bi xoa khi tat) - phai lay theo file .exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class LoiNghiepVu(Exception):
    pass


class DaDung(LoiNghiepVu):
    pass


def kiem_tra_dung(dung_lai):
    if dung_lai and dung_lai():
        raise DaDung("Da dung tac vu. Khong xuat bo file dang do.")


def kiem_tra_cfg(cfg):
    c = dict(CFG_MAC_DINH)
    c.update(cfg or {})
    try:
        for k, default in CFG_MAC_DINH.items():
            if isinstance(default, (int, float)) and not isinstance(default, bool) and not (k in ("reL_min", "rm_min", "a_min") and c[k] is None):
                c[k] = spc_luat.so_huu_han(c[k], k)
        for k in ("so_luong", "so_bat_dau", "so_chu_so"):
            c[k] = spc_luat.so_nguyen(c[k], k, 0 if k == "so_chu_so" else 1)
        if not 25 <= c["so_luong"] <= 20000:
            raise ValueError("So luong mau phai tu 25 den 20000.")
        for k in ("d0", "Lo", "Le", "reL_tb", "rm_tb", "reL_sd", "rm_sd", "a_sd", "z_sd", "a_tb", "he_so_gian_dai", "cpk_toi_thieu", "giay_toi_da"):
            if c[k] <= 0:
                raise ValueError(k + " phai > 0.")
        if c["dung_sai_d"] < 0 or round(c["d0"] - c["dung_sai_d"], 2) <= 0:
            raise ValueError("Can d0 > dung sai >= 0 va duong kinh lam tron > 0.")
        if not 0 < c["z_tb"] < 100:
            raise ValueError("Z trung binh phai nam trong (0, 100).")
        if not 0 <= c["reH_chenh_min"] <= c["reH_chenh_max"]:
            raise ValueError("Chenh ReH phai khong am, min <= max.")
        for k in ("ti_le_fm", "dut_duoi_ti_le", "dut_tren_ti_le_min", "dut_tren_ti_le_max"):
            if not 0 < c[k] < 1:
                raise ValueError(k + " phai nam trong (0, 1).")
        if c["dut_tren_ti_le_min"] > c["dut_tren_ti_le_max"] or c["dut_duoi_ti_le"] * 1.03 >= 1:
            raise ValueError("Ti le luc dut khong hop le.")
        if c["feh_pos_tb"] <= 0 or c["fel_lech_tb"] <= 0 or c["feh_pos_sd"] < 0 or c["fel_lech_sd"] < 0:
            raise ValueError("Vi tri chay va do lech khong hop le.")
        for k in ("reL_min", "reH_min", "rm_min", "a_min", "z_min_spec"):
            if c.get(k) is not None:
                c[k] = spc_luat.so_huu_han(c[k], k)
        if c.get("z_min_spec") is not None and not 0 <= c["z_min_spec"] < 100:
            raise ValueError("Nguong Z phai nam trong [0, 100).")
        for key in ("buoc_nhay", "buoc_moi_diem"):
            c[key] = [spc_luat.so_nguyen(v, key) for v in c.get(key) or []]
        c["day_rieng"] = [(spc_luat.so_nguyen(a, "Mau bat dau"), spc_luat.so_nguyen(b, "Buoc nhay")) for a, b in c.get("day_rieng") or []]
        for a, b in c["day_rieng"]:
            if a + 24 * b > c["so_luong"]:
                raise ValueError("Day rieng vuot qua so mau da khai.")
    except (TypeError, ValueError, OverflowError) as e:
        raise LoiNghiepVu(str(e))
    return c


class HangGiaTri(list):
    """Kich thuoc di cung hang khi hoan vi, khong boc lai sau cham SPC."""
    mau = None


def _lay(cfg, k):
    if k in ("reL_min", "rm_min", "a_min"):
        return cfg.get(k, CFG_MAC_DINH[k])
    v = cfg.get(k)
    return CFG_MAC_DINH[k] if v is None else v


# ------------------------------------------------------- 1. SINH GIA TRI
def sinh_gia_tri(cfg, rnd, dung_lai=None):
    """Sinh N hang [ReL, ReH, Rm, A, Z]. Chua sap xep thu tu.

    ReH sinh bang ReL cong mot luong duong nen luon lon hon ReL - dung ban
    chat vat ly cua gioi han chay tren / chay duoi.
    """
    cfg = kiem_tra_cfg(cfg)
    n = int(_lay(cfg, "so_luong"))
    if n < 25:
        raise LoiNghiepVu("So luong mau phai tu 25 tro len (mot bo SPC can "
                          "25 diem). Dang khai %d." % n)
    if n > 20000:
        raise LoiNghiepVu("So luong mau qua lon (toi da 20.000).")

    reL_tb, reL_sd = _lay(cfg, "reL_tb"), _lay(cfg, "reL_sd")
    rm_tb, rm_sd = _lay(cfg, "rm_tb"), _lay(cfg, "rm_sd")
    a_tb, a_sd = _lay(cfg, "a_tb"), _lay(cfg, "a_sd")
    z_tb, z_sd = _lay(cfg, "z_tb"), _lay(cfg, "z_sd")
    ch_min, ch_max = _lay(cfg, "reH_chenh_min"), _lay(cfg, "reH_chenh_max")
    nguong = [_lay(cfg, "reL_min"), cfg.get("reH_min"), _lay(cfg, "rm_min"),
              _lay(cfg, "a_min"), cfg.get("z_min_spec")]

    ds = []
    for _ in range(n):
        kiem_tra_dung(dung_lai)
        for _lan in range(200):
            reL = round(rnd.gauss(reL_tb, reL_sd), SO_LE_UNG_SUAT)
            reH = round(reL + rnd.uniform(ch_min, ch_max), SO_LE_UNG_SUAT)
            rm = round(rnd.gauss(rm_tb, rm_sd), SO_LE_UNG_SUAT)
            a = round(rnd.gauss(a_tb, a_sd), 1)
            z = round(rnd.gauss(z_tb, z_sd), 1)
            if not (0 < z < 100 and a > 0 and reL > 0):
                continue
            hang = HangGiaTri([reL, reH, rm, a, z])
            # rang buoc vat ly + nguong nghiem thu
            if rm <= reH + 20:
                continue
            hang.mau = suy_ra(hang, cfg, rnd)
            hang[3], hang[4] = hang.mau["A"], hang.mau["Z"]
            if any(g is not None and v < g for v, g in zip(hang, nguong)):
                continue
            ds.append(hang)
            break
        else:
            raise LoiNghiepVu(
                "Khong sinh duoc gia tri thoa nguong nghiem thu. Hay kiem tra "
                "lai cac o trung binh / do lech so voi nguong toi thieu.")
    return ds


# --------------------------------------------- 2. RANG BUOC VA TOI UU THU TU
def xay_khoi(n, buoc_tu_dau, buoc_moi_diem=(), day_rieng=(), so_diem=25):
    """Liet ke moi day con can dat SPC.

    Hai kieu bao dam khac han nhau ve suc manh va ve gia:

    buoc_tu_dau    voi buoc nhay k, chi xet cac day BAT DAU TU DAU LO:
                   o, o+k, o+2k, ... voi o chay tu 0 den k-1.
                   Re (buoc k cho k khoi) nhung KHONG phu truong hop nguoi
                   ta lay 25 mau tu giua lo, vi du 301, 311, 321...

    buoc_moi_diem  voi buoc nhay k, xet MOI diem bat dau s trong ca lo.
                   Dat nhat nhung la cai nguoi dung thuc su can: "lay 25
                   mau lien tiep bat ky o dau cung dat" chinh la k = 1 o
                   day. Do thuc te: k=1 het 11 giay, them k=2 het 50 giay,
                   them k=5 nua thi khong hoi tu trong 5 phut.

    day_rieng      danh sach cap (mau_bat_dau, buoc_nhay) khai tay, mau dem
                   tu 1 nhu tren giao dien. Moi cap chi them DUNG MOT day
                   nen re nhat, dung khi biet truoc se lay mau kieu nao -
                   vi du se nop bao cao SPC lay tu mau 26 buoc 5 thi khai
                   (26, 5) o day, bo du lieu sinh ra da thoa san.

    Tra ve list (tuple_chi_so, chi_so_cot), da bo trung lap.
    """
    thay = set()
    ra = []

    def them(idx):
        if idx[0] >= 0 and idx[-1] < n and idx not in thay:
            thay.add(idx)
            for c in range(len(COT_SPC)):
                ra.append((idx, c))

    for k in sorted(set(int(x) for x in (buoc_moi_diem or ()) if int(x) >= 1)):
        for s in range(n - so_diem * k + k + 1):
            them(tuple(range(s, s + so_diem * k, k)))
    for k in sorted(set(int(x) for x in (buoc_tu_dau or ()) if int(x) >= 1)):
        for o in range(k):
            them(tuple(range(o, o + so_diem * k, k)))
    for bd, k in (day_rieng or ()):
        bd, k = int(bd), int(k)
        if bd >= 1 and k >= 1:
            them(tuple(range(bd - 1, bd - 1 + so_diem * k, k)))
    return ra


def _dem_loi_kep(X, c, lsl, cpk_min, bo_qua):
    """So loi SPC cua mot day. Voi ReL/ReH/Rm cham CA so le (file Excel) LAN
    so da lam tron (bao cao Word): du sau nay ai lay so o dau de danh gia
    (Tab 4 doc Excel hay doc bao cao Word) thi cung dat."""
    loi = spc_luat.dem_loi(X, lsl, None, cpk_min, bo_qua)
    if c in COT_LAM_TRON_BAO_CAO:
        loi += spc_luat.dem_loi([lam_tron_bao_cao(x) for x in X], lsl, None,
                                cpk_min, bo_qua)
    return loi


def _lsl_cot(cfg):
    return [_lay(cfg, "reL_min"), cfg.get("reH_min"), _lay(cfg, "rm_min"),
            _lay(cfg, "a_min"), cfg.get("z_min_spec")]


def toi_uu_thu_tu(ds, cfg, rnd, tien_do=None, dung_lai=None):
    """Hoan vi thu tu cac HANG cho toi khi moi khoi rang buoc deu dat.

    Cac luat hinh dang chi phu thuoc THU TU, khong phu thuoc TAP GIA TRI.
    Nen hoan vi khong lam thay doi dai so, nguong nghiem thu hay phan bo -
    chi thay doi thu tu xuat hien.

    Nuoc di uu tien la doi cho hai vi tri TRONG CUNG mot khoi dang hong:
    khi do tap gia tri va X-bar cua khoi khong doi; MR, sigma va Cpk van
    thay doi. Nuoc di nay lam thuat toan hoi tu nhanh
    hon han so voi doi cho ngau nhien toan mang.

    Tra ve (so_khoi_con_hong, so_vong, giay).
    """
    n = len(ds)
    khoi = xay_khoi(n, _lay(cfg, "buoc_nhay"), cfg.get("buoc_moi_diem") or (),
                    cfg.get("day_rieng") or ())
    if not khoi:
        return 0, 0, 0.0
    lsl = _lsl_cot(cfg)
    cpk_min = float(_lay(cfg, "cpk_toi_thieu"))
    bo_qua = tuple(cfg.get("bo_qua_luat") or ())
    han = float(_lay(cfg, "giay_toi_da"))

    thuoc = {}
    for bi, (idx, c) in enumerate(khoi):
        for i in idx:
            thuoc.setdefault(i, []).append(bi)

    def diem(bi):
        idx, c = khoi[bi]
        return _dem_loi_kep([ds[i][c] for i in idx], c, lsl[c], cpk_min,
                            bo_qua)

    chi_phi = [diem(b) for b in range(len(khoi))]
    t0 = datetime.now()
    vong = 0
    while True:
        hong = [b for b, x in enumerate(chi_phi) if x]
        if not hong:
            break
        giay = (datetime.now() - t0).total_seconds()
        if giay > han:
            break
        kiem_tra_dung(dung_lai)
        vong += 1
        if tien_do and vong % 500 == 0:
            tien_do(len(hong), giay)

        idx, _c = khoi[rnd.choice(hong)]
        i = rnd.choice(idx)
        j = rnd.choice(idx) if rnd.random() < 0.6 else rnd.randrange(n)
        if i == j:
            continue
        cham = set(thuoc.get(i, ()))
        cham.update(thuoc.get(j, ()))
        truoc = 0
        for b in cham:
            truoc += chi_phi[b]
        ds[i], ds[j] = ds[j], ds[i]
        moi = {}
        sau = 0
        for b in cham:
            q = diem(b)
            moi[b] = q
            sau += q
        if sau <= truoc:
            for b, q in moi.items():
                chi_phi[b] = q
        else:
            ds[i], ds[j] = ds[j], ds[i]

    return sum(1 for x in chi_phi if x), vong, \
        (datetime.now() - t0).total_seconds()


def quet_phu(ds, cfg, buoc_toi_da=None, so_mau_thu=60, dung_lai=None):
    """Cham lai du lieu de in bang phu.

    Quet HAI kieu lay mau, vi chung cho ket qua rat khac nhau:
      - 'tu dau lo'  : day bat dau tu k mau dau tien
      - 'giua lo'    : day bat dau tu bat ky dau trong lo (lay mau so_mau_thu
                       diem bat dau rai deu, du de thay buc tranh that)

    Tra ve list (buoc, dat_tu_dau, tong_tu_dau, dat_giua, tong_giua).
    Muc dich la de nguoi dung nhin thay chinh xac cach lay nao duoc bao dam
    va cach nao khong - khong de ai phai doan.
    """
    n = len(ds)
    lsl = _lsl_cot(cfg)
    cpk_min = float(_lay(cfg, "cpk_toi_thieu"))
    bo_qua = tuple(cfg.get("bo_qua_luat") or ())
    kmax = buoc_toi_da or (n // 25)

    def dat(idx):
        return all(not _dem_loi_kep([ds[i][c] for i in idx], c, lsl[c],
                                    cpk_min, bo_qua)
                   for c in range(len(COT_SPC)))

    ra = []
    for k in range(1, kmax + 1):
        kiem_tra_dung(dung_lai)
        t1 = d1 = 0
        for o in range(k):
            idx = list(range(o, o + 25 * k, k))
            if idx[-1] < n:
                t1 += 1
                d1 += dat(idx)
        het = n - 25 * k + k
        t2 = d2 = 0
        if het > 0:
            buoc_s = max(1, het // so_mau_thu)
            for s in range(0, het + 1, buoc_s):
                idx = list(range(s, s + 25 * k, k))
                if idx[-1] < n:
                    t2 += 1
                    d2 += dat(idx)
        if t1 or t2:
            ra.append((k, d1, t1, d2, t2))
    return ra


# ------------------------------------------------- 3. SUY RA KICH THUOC, LUC
def _z_dien_tich(d, du):
    return (1.0 - (du / d) ** 2) * 100.0


def _du_khop_z(d, Z):
    """Chon du tren luoi 0,01 mm gan Z muc tieu nhat.

    Luoi khong luon bieu dien duoc Z muc tieu; suy_ra tinh lai Z thuc.
    """
    goc = round(d * math.sqrt(max(0.0, 1.0 - Z / 100.0)), 2)
    ung = [round(goc + k * 0.01, 2) for k in range(-3, 4)]
    return min(ung, key=lambda du: abs((1.0 - (du / d) ** 2) * 100.0 - Z))


def suy_ra(hang, cfg, rnd):
    """Tu [ReL, ReH, Rm, A, Z] suy ra day du thong so mot mau."""
    if getattr(hang, "mau", None) is not None:
        return dict(hang.mau)
    d0 = float(_lay(cfg, "d0"))
    dsai = abs(float(_lay(cfg, "dung_sai_d")))
    Lo = float(_lay(cfg, "Lo"))
    Le = float(_lay(cfg, "Le"))

    reL, reH, rm, A, Z = hang
    if d0 <= dsai or Lo <= 0 or Le <= 0 or not 0 < Z < 100 or A <= 0:
        raise LoiNghiepVu("Kich thuoc hoac A/Z khong hop le.")
    d = round(rnd.uniform(d0 - dsai, d0 + dsai), 2) if dsai > 0 else d0
    S0 = math.pi * d * d / 4.0

    if _lay(cfg, "z_theo_dien_tich"):
        du = _du_khop_z(d, Z)
        # Hiem (~0,1%) khong co du 0,01 mm nao cho lai dung Z: rut lai d
        # trong khoang dung sai (d von ngau nhien, khong anh huong SPC)
        for _ in range(30):
            if dsai <= 0 or "%.1f" % _z_dien_tich(d, du) == "%.1f" % Z:
                break
            d = round(rnd.uniform(d0 - dsai, d0 + dsai), 2)
            du = _du_khop_z(d, Z)
        S0 = math.pi * d * d / 4.0
    else:
        du = round(d * (1.0 - Z / 100.0), 2)   # cong thuc cu cua BM32
    Su = math.pi * du * du / 4.0
    Lu = round(Lo * (1.0 + A / 100.0), 2)
    Z = round(_z_dien_tich(d, du) if _lay(cfg, "z_theo_dien_tich") else (1 - du / d) * 100, 1)
    A = round((Lu / Lo - 1) * 100, 1)

    return {
        "d": d, "du": du, "S0": S0, "Su": Su,
        "Lo": Lo, "Lu": round(Lu, 2), "Le": Le,
        "ReL": reL, "ReH": reH, "Rm": rm, "A": A, "Z": Z,
        "FeL": reL * S0 / 1000.0,
        "FeH": reH * S0 / 1000.0,
        "Fm": rm * S0 / 1000.0,
    }


def tham_so_duong_cong(m, cfg, rnd):
    """Tu thong so mot mau, dung tham so ve bieu do.

    Quang duong luc dut suy nguoc tu do gian dai A qua he so quy doi, dung
    dung quan he ma tensile_core.py dang dung theo chieu nguoc:
        Lu = Lo + he_so * dut_pos   ->   dut_pos = (Lu - Lo)/he_so
    Nho vay bieu do va cot A trong bang luon ke cung mot cau chuyen.

    Luc luc dut KHONG suy tu Z/du: Buoc 2 doc thang cot du_mm trong Excel
    de in bao cao, khong can luc dut, nen o day chi can bieu do giong mau
    that: luc dut 70-78% Fm (mau 522: 58.8/80.6 = 73%).
    """
    hs = float(_lay(cfg, "he_so_gian_dai"))
    dut_pos = (m["Lu"] - m["Lo"]) / hs
    fm_pos = dut_pos * float(_lay(cfg, "ti_le_fm")) * rnd.uniform(0.99, 1.01)
    feh_pos = rnd.gauss(_lay(cfg, "feh_pos_tb"), _lay(cfg, "feh_pos_sd"))
    fel_pos = feh_pos + abs(rnd.gauss(_lay(cfg, "fel_lech_tb"),
                                      _lay(cfg, "fel_lech_sd")))
    dut_tren = m["Fm"] * rnd.uniform(float(_lay(cfg, "dut_tren_ti_le_min")),
                                     float(_lay(cfg, "dut_tren_ti_le_max")))
    dut_duoi = dut_tren * _lay(cfg, "dut_duoi_ti_le") * rnd.uniform(0.97, 1.03)
    if not 0 < feh_pos < fel_pos < fm_pos < dut_pos:
        raise LoiNghiepVu("Thong so khong tao duoc duong cong: can 0 < FeH_pos < FeL_pos < Fm_pos < dut_pos. Kiem tra A/Lo/he so gian dai.")
    return {
        "feh": round(m["FeH"], 3), "feh_pos": round(feh_pos, 3),
        "fel": round(m["FeL"], 3), "fel_pos": round(fel_pos, 3),
        "fm": round(m["Fm"], 3), "fm_pos": round(fm_pos, 3),
        "dut_pos": round(dut_pos, 3),
        "dut_tren": round(dut_tren, 3), "dut_duoi": round(dut_duoi, 3),
        # Bon he so hinh dang duoi day khop tu anh bieu do mau trong
        # 'S2-1 TENSILE TESTING 522.docx' (1.80 / 3.65 / 3.45 / 0,20 mm),
        # tha ra mot khoang hep quanh do de moi mau khong giong het nhau.
        # Rang cua doan chay deo tu xe dich theo tung mau, xem
        # ve_bieu_do_pdf.MAU_CHAY.
        "mu_dan_hoi": round(rnd.uniform(1.74, 1.86), 3),
        "k_hoa_ben": round(rnd.uniform(3.50, 3.80), 3),
        "mu_that_co": round(rnd.uniform(3.35, 3.55), 3),
        "chan_chay": round(rnd.uniform(0.17, 0.23), 3),
    }


def ma_mau(cfg, i):
    tt = str(_lay(cfg, "tien_to_ma") or "")
    cs = int(_lay(cfg, "so_chu_so") or 0)
    stt = int(_lay(cfg, "so_bat_dau")) + i
    return tt + (("%0*d" % (cs, stt)) if cs else str(stt))
