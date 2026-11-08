# -*- coding: utf-8 -*-
"""
BO LUAT DANH GIA DO ON DINH SPC (I-MR chart)
=============================================
Cai dat DAY DU cac muc danh gia in tren bieu mau JIS.QT.07.BM.01
"BAO CAO DANH GIA DO ON DINH CHAT LUONG":

    Bang I-Chart  (o B64:C72 cua bieu mau) - 8 muc
    Bang MR-Chart (o C46:D49 cua bieu mau) - 3 muc
    Muc 6 "Danh gia nang luc"              - Ca / Cp / Cpk

Module nay CHI TINH VA CHAM DIEM, khong sinh file. Dung chung cho:
    - Tab 1 (Charpy)  : thay cho bo 4 dieu kien rut gon truoc day
    - Tab 3 (sinh keo): cham 5 cot ReL / ReH / Rm / A / Z

Hai muc MR-chart trong bieu mau ghi bang loi van, khong kem nguong so.
Nguong dung o day khai o MR2_SO_DIEM / MR3_SO_DIEM / MR3_HE_SO ngay duoi,
sua mot cho la doi ca phan mem.

MR3 ("dao dong dot ngot keo dai") CHI DE THAM KHAO: bieu mau khong co o tick
cho muc nay va nguong la do phan mem tu dat, nen MR3 KHONG tinh vao ket luan
DAT / KHONG DAT (truoc day co tinh: bao cao hien moi o deu dat ma ket luan
lai khong dat). Ket qua MR3 nam o KetQuaSPC.mr3 de in ghi chu.
"""

import math

D2 = 1.128        # he so d2 cho nhom co 1 quan sat  -> sigma = MR-bar / d2
D4 = 3.267        # he so D4 cho gioi han tren cua MR chart

# --- Nguong cho hai muc MR-chart ma bieu mau khong ghi ro con so ---
MR2_SO_DIEM = 6   # 6 gia tri MR lien tiep tang hoac giam
MR3_SO_DIEM = 3   # "Dao dong dot ngot keo dai"    = 3 gia tri MR lien tiep
MR3_HE_SO = 2.0   #                                  cung vuot 2 x MR-bar

# Diem cach duong tam duoi EPS * sigma coi nhu NAM TREN duong tam.
#
# Khong co nguong nay thi mot bo so lieu co X-bar tron chan (vi du 25 gia tri
# quanh 370.0 cho dung X-bar = 370.0) se cho ket qua khac nhau chi vi thu tu
# cong don lam tron khac nhau o bit cuoi cung: diem "bang dung X-bar" luc thi
# duoc tinh la tren, luc thi duoi, lam dut hay noi mot chuoi 9 diem cua muc 4.
# Mot cong cu QC khong duoc phep co hanh vi nhu vay.
EPS = 1e-9

TEN_LUAT = {
    "I1": "Diem nam ngoai UCLI/LCLI",
    "I2": "2/3 diem lien tiep roi vao vung A cung phia",
    "I3": "4/5 diem lien tiep roi vao vung B cung phia hoac xa hon",
    "I4": "9 diem lien tiep nam cung mot phia duong tam",
    "I5": "6 diem lien tiep tang dan hoac giam dan",
    "I6": "14 diem lien tiep luan phien tang va giam",
    "I7": "15 diem lien tiep nam trong vung C",
    "I8": "8 diem lien tiep nam ben ngoai vung C",
    "MR1": "Diem vuot UCLMR",
    "MR2": "6 diem MR lien tiep tang dan hoac giam dan",
    "MR3": "Dao dong dot ngot keo dai (chi tham khao)",
    "Cpk": "Cpk chua dat nguong",
    "SPEC": "Co so do nam ngoai gioi han nghiem thu LSL/USL",
    "SIGMA": "Khong du bien thien de danh gia nang luc",
    "DU_LIEU": "Thieu du lieu hoac co gia tri khong hop le",
}

# Thu tu in ra bao cao, bam theo thu tu in tren bieu mau
THU_TU = ["I1", "I2", "I3", "I4", "I5", "I6", "I7", "I8",
          "MR1", "MR2", "MR3", "Cpk", "SPEC", "SIGMA", "DU_LIEU"]


def so_huu_han(v, ten="Gia tri"):
    try:
        x = float(v)
    except (TypeError, ValueError, OverflowError):
        raise ValueError("%s phai la so huu han." % ten)
    if isinstance(v, bool) or not math.isfinite(x):
        raise ValueError("%s phai la so huu han." % ten)
    return x


def so_nguyen(v, ten="So luong", toi_thieu=1):
    x = so_huu_han(v, ten)
    if x != int(x) or x < toi_thieu:
        raise ValueError("%s phai la so nguyen >= %d." % (ten, toi_thieu))
    return int(x)


def kiem_tra_nguong(lsl, usl, cpk=1.33):
    lsl = None if lsl is None else so_huu_han(lsl, "LSL")
    usl = None if usl is None else so_huu_han(usl, "USL")
    cpk = so_huu_han(cpk, "Cpk toi thieu")
    if cpk <= 0 or (lsl is not None and usl is not None and lsl >= usl):
        raise ValueError("Can LSL < USL va Cpk toi thieu > 0.")
    return lsl, usl, cpk


class KetQuaSPC(object):
    """Chi so I-MR va danh sach ma luat bi vi pham."""

    __slots__ = ("n", "X", "MR", "xb", "mrb", "sigma", "ucl_i", "lcl_i",
                 "ucl_mr", "lcl_mr", "lsl", "usl", "cpk", "cp", "ca", "loi",
                 "mr3")

    def __init__(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)

    @property
    def dat(self):
        return not self.loi

    def mo_ta_loi(self):
        return [TEN_LUAT.get(m, m) for m in self.loi]


def tinh(X, lsl=None, usl=None):
    """Tinh cac chi so I-MR. Khong cham luat.

    sigma uoc luong theo pham vi di dong: sigma = MR-bar / 1.128, dung nhu
    o E23 cua bieu mau (khong phai do lech chuan mau).
    """
    X = [so_huu_han(v, "Diem %d" % (i + 1)) for i, v in enumerate(X)]
    lsl, usl, _ = kiem_tra_nguong(lsl, usl)
    n = len(X)
    if n < 2:
        raise ValueError("Can it nhat 2 diem du lieu moi danh gia duoc SPC.")
    xb = math.fsum(X) / n
    MR = [abs(X[i] - X[i - 1]) for i in range(1, n)]
    mrb = math.fsum(MR) / len(MR)
    sg = mrb / D2

    r = KetQuaSPC(n=n, X=list(X), MR=MR, xb=xb, mrb=mrb, sigma=sg,
                  ucl_i=xb + 3 * sg, lcl_i=xb - 3 * sg,
                  ucl_mr=D4 * mrb, lcl_mr=0.0,
                  lsl=lsl, usl=usl, cpk=None, cp=None, ca=None, loi=[],
                  mr3=False)
    if sg > 0:
        phia = []
        if lsl is not None:
            phia.append((xb - lsl) / (3 * sg))
        if usl is not None:
            phia.append((usl - xb) / (3 * sg))
        r.cpk = min(phia) if phia else None
        if usl is not None and lsl is not None:
            r.cp = (usl - lsl) / (6 * sg)
            r.ca = abs(xb - (usl + lsl) / 2) / ((usl - lsl) / 2)
    return r


def _z(X, xb, sg):
    """Toa do chuan hoa, co ap nguong EPS quanh duong tam (xem chu thich EPS)."""
    ra = []
    for x in X:
        t = (x - xb) / sg
        ra.append(0.0 if -EPS < t < EPS else t)
    return ra


def cham(X, lsl=None, usl=None, cpk_toi_thieu=1.33, bo_qua=()):
    """Cham mot bo du lieu theo day du cac muc trong bieu mau.

    bo_qua: tap ma luat khong xet, vi du ("MR3",).
    Tra ve KetQuaSPC, xem .dat / .loi / .mo_ta_loi().
    """
    lsl, usl, cpk_toi_thieu = kiem_tra_nguong(lsl, usl, cpk_toi_thieu)
    r = tinh(X, lsl, usl)
    X = r.X
    n = r.n
    loi = []
    if any((lsl is not None and x < lsl) or
           (usl is not None and x > usl) for x in X):
        loi.append("SPEC")
    if r.sigma <= 0:
        r.loi = loi + ["SIGMA"]
        if n >= 15 and "I7" not in bo_qua:
            r.loi.append("I7")
        return r

    z = _z(X, r.xb, r.sigma)

    # --- I1: diem nam ngoai +-3 sigma ---
    for t in z:
        if t > 3.0 or t < -3.0:
            loi.append("I1"); break

    # --- I2: 2/3 diem lien tiep trong vung A (|z|>2) cung phia ---
    for i in range(n - 2):
        tren = duoi = 0
        for j in range(i, i + 3):
            if z[j] > 2.0: tren += 1
            elif z[j] < -2.0: duoi += 1
        if tren >= 2 or duoi >= 2:
            loi.append("I2"); break

    # --- I3: 4/5 diem lien tiep tu vung B tro ra (|z|>1) cung phia ---
    for i in range(n - 4):
        tren = duoi = 0
        for j in range(i, i + 5):
            if z[j] > 1.0: tren += 1
            elif z[j] < -1.0: duoi += 1
        if tren >= 4 or duoi >= 4:
            loi.append("I3"); break

    # --- I4: 9 diem lien tiep cung phia duong tam ---
    # run dem SO DIEM. Diem nam dung tren duong tam (z = 0) cat chuoi.
    run = 0; dau = 0
    for t in z:
        d = 1 if t > 0 else (-1 if t < 0 else 0)
        if d == 0: run = 0
        elif d == dau: run += 1
        else: run = 1
        dau = d
        if run >= 9:
            loi.append("I4"); break

    # --- I5: 6 diem lien tiep tang dan hoac giam dan ---
    tang = giam = 1
    for i in range(1, n):
        if X[i] > X[i - 1]: tang += 1; giam = 1
        elif X[i] < X[i - 1]: giam += 1; tang = 1
        else: tang = giam = 1
        if tang >= 6 or giam >= 6:
            loi.append("I5"); break

    # --- I6: 14 diem lien tiep luan phien tren/duoi ---
    # alt dem SO DIEM: 14 diem tuong ung 13 hieu so doi dau lien tiep.
    alt = 1; truoc = 0
    for i in range(1, n):
        d = X[i] - X[i - 1]
        s = 1 if d > 0 else (-1 if d < 0 else 0)
        if s == 0: alt = 1
        elif truoc != 0 and s == -truoc: alt += 1
        else: alt = 2
        truoc = s
        if alt >= 14:
            loi.append("I6"); break

    # --- I7: 15 diem lien tiep trong vung C (|z| < 1) ---
    run = 0
    for t in z:
        run = run + 1 if -1.0 < t < 1.0 else 0
        if run >= 15:
            loi.append("I7"); break

    # --- I8: 8 diem lien tiep ngoai vung C ---
    run = 0
    for t in z:
        run = run + 1 if (t >= 1.0 or t <= -1.0) else 0
        if run >= 8:
            loi.append("I8"); break

    # --- MR1: diem vuot UCLMR ---
    for m in r.MR:
        if m > r.ucl_mr:
            loi.append("MR1"); break

    # --- MR2: 6 diem MR lien tiep tang hoac giam ---
    run = giam = 1
    for i in range(1, len(r.MR)):
        run = run + 1 if r.MR[i] > r.MR[i - 1] else 1
        giam = giam + 1 if r.MR[i] < r.MR[i - 1] else 1
        if max(run, giam) >= MR2_SO_DIEM:
            loi.append("MR2"); break

    # --- MR3: dao dong dot ngot keo dai - CHI THAM KHAO, khong vao loi ---
    run = 0; nguong = MR3_HE_SO * r.mrb
    for m in r.MR:
        run = run + 1 if m > nguong else 0
        if run >= MR3_SO_DIEM:
            r.mr3 = True; break

    # --- Danh gia nang luc ---
    if r.cpk is not None and r.cpk < cpk_toi_thieu:
        loi.append("Cpk")

    r.loi = [m for m in THU_TU if m in loi and m not in bo_qua]
    return r


def dem_loi(X, lsl=None, usl=None, cpk_toi_thieu=1.33, bo_qua=()):
    """So luat bi vi pham. Dung lam ham chi phi cho vong toi uu thu tu:
    tra ve so dem (khong phai True/False) de tim kiem co do doc de bam."""
    return len(cham(X, lsl, usl, cpk_toi_thieu, bo_qua).loi)


def bang_chu(r, cpk_toi_thieu=1.33):
    """Bao cao chu cho mot bo du lieu, bam theo bo cuc muc 3-6 cua bieu mau."""
    L = []
    L.append("  So diem du lieu       : %d" % r.n)
    L.append("  Trung binh      X-bar : %.4f" % r.xb)
    L.append("  Pham vi di dong MR-bar: %.4f" % r.mrb)
    L.append("  Do lech chuan   sigma : %.6f   (MR-bar / %.3f)" % (r.sigma, D2))
    L.append("  I-Chart   UCLI / LCLI : %.4f / %.4f" % (r.ucl_i, r.lcl_i))
    L.append("  MR-Chart  UCLMR       : %.4f" % r.ucl_mr)
    if r.lsl is not None:
        L.append("  Gioi han duoi   LSL   : %g" % r.lsl)
    if r.usl is not None:
        L.append("  Gioi han tren   USL   : %g" % r.usl)
    if r.cpk is not None:
        L.append("  Nang luc qua trinh Cpk: %.3f  (yeu cau >= %g)"
                 % (r.cpk, cpk_toi_thieu))
    else:
        L.append("  Cpk: khong ap dung / khong du bien thien")
    if r.dat:
        L.append("  KET LUAN: THOA MAN TAT CA CAC MUC DANH GIA")
    else:
        L.append("  KET LUAN: CHUA DAT - vi pham %d muc:" % len(r.loi))
        for m in r.loi:
            L.append("     [%-3s] %s" % (m, TEN_LUAT.get(m, m)))
    return "\n".join(L)
