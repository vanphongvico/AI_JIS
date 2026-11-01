# -*- coding: utf-8 -*-
"""
KICH THUOC CORE - SINH DU LIEU KICH THUOC & DANH GIA SPC
=========================================================
Doc file master data "Dung sai kich thuoc.xlsx", sinh du lieu 25 mau
cho N chi tieu kich thuoc, cham SPC I-MR DUNG NHU CONG THUC cua bieu mau
JIS.QT.07.BM.01 (sheet "Du lieu goc") - de phan mem bao "Dat" thi mo file
bang Excel cung "Dat".

Cong thuc bieu mau (hang x / RM cua moi chi tieu):
    x    = AVERAGE(cac diem do cua mau)          - KHONG lam tron
    RM   = ABS(x(i) - x(i-1))
    sigma= ROUND(MR-bar / 1.128, 2)
    +-k sigma  = ROUND(X-bar +- k*sigma, 1)       k = 1, 2, 3
    UCLRM= ROUND(3.267 * MR-bar, 2)
    Ca = (X-bar - To)/(T/2),  Cp = T/(6 sigma),  Cpk = (1-|Ca|)*Cp
"""

import re
import math
import random

import openpyxl

try:
    import spc_luat
except ImportError:
    from . import spc_luat

# So mau mac dinh
SO_MAU = 25

# Moi lan sinh mot bo 25 mau roi doi cho toi da SO_LAN_DOI_CHO lan cho dat;
# chua dat thi sinh bo khac, toi da SO_LAN_SINH_LAI bo.
SO_LAN_SINH_LAI = 40
BUOC_DO = 0.1          # so do lam tron 0,1 mm
SO_LAN_DOI_CHO = 800
MAX_RETRY = SO_LAN_SINH_LAI

# Cpk toi thieu yeu cau
CPK_TOI_THIEU = 1.33

D2 = 1.128
D4 = 3.267

# Hai so cach nhau duoi muc nay coi la bang nhau (xem cham_bieu_mau)
SAI_SO = 1e-7


# ---------------------------------------------------------------------------
# 1. DOC FILE MASTER DATA "DUNG SAI KICH THUOC"
# ---------------------------------------------------------------------------
_SO = r"([-+]?\d+(?:\.\d+)?)"


def _phan_tich_dung_sai(chuoi_yeu_cau):
    """Phan tich chuoi yeu cau ky thuat thanh (gia_tri_chuan, ds_duoi, ds_tren).

    Vi du:
        '900.0(0~+15mm)'   -> (900.0, 0, 15)
        '355.6(-1~+3mm)'   -> (355.6, -1, 3)
        '45±0.5'           -> (45, -0.5, 0.5)
        '6.0(0㎜～)'       -> (6.0, 0, None)      chi co gioi han duoi
        '100.0'            -> (100.0, None, None) kich thuoc tham khao
    Khong doc duoc thi tra ve (None, None, None).
    """
    if chuoi_yeu_cau is None:
        return None, None, None
    s = str(chuoi_yeu_cau).strip()
    if not s:
        return None, None, None

    m = re.match(_SO + r"\s*±\s*" + _SO, s)
    if m:
        tol = abs(float(m.group(2)))
        return float(m.group(1)), -tol, tol

    m = re.match(_SO + r"\s*\(\s*" + _SO + r"\s*(?:mm|㎜)?\s*[～~]\s*"
                 + _SO + r"\s*(?:mm|㎜)\s*\)", s)
    if m:
        return float(m.group(1)), float(m.group(2)), float(m.group(3))

    m = re.match(_SO + r"\s*\(\s*" + _SO + r"\s*(?:mm|㎜)?\s*[～~]\s*\)", s)
    if m:                                    # mot phia: chi co dung sai duoi
        return float(m.group(1)), float(m.group(2)), None

    m = re.fullmatch(_SO, s)
    if m:                                    # chi co con so, khong dung sai
        return float(m.group(1)), None, None
    return None, None, None


def _chu_yeu_cau(v):
    """Chuoi 'Yeu cau' de in. O so (ket qua cong thuc trong file dung sai,
    vd 100.00000000000001) thi in gon lai nhu Excel hien thi: '100'."""
    if isinstance(v, (int, float)):
        return "%g" % round(float(v), 2)
    return str(v).strip() if v is not None else ""


def _so_hang_khoi(ws, r, r_ke):
    """So hang (= so diem do) cua chi tieu bat dau o hang r.

    Uu tien o gop cua cot B/C (moi chi tieu la mot khoi gop nhieu hang).
    Khong gop thi tinh toi truoc STT ke tiep.
    """
    for mr in ws.merged_cells.ranges:
        if mr.min_row == r and mr.min_col <= 3 <= mr.max_col or \
                mr.min_row == r and mr.min_col == 2:
            return mr.max_row - mr.min_row + 1
    if r_ke is not None:
        return r_ke - r
    return 1


def doc_dung_sai(duong_dan):
    """Doc file 'Dung sai kich thuoc.xlsx' tra ve cau truc du lieu.

    Ket qua:
        {
            "ALKTOP II": {
                "ma_san_pham": ["1031M", ...],
                "chi_tieu": [
                    {"stt": 1, "ten": "Dw(翼径)", "ky_hieu": "AD",
                     "ky_hieu_diem": ["AD", "BE", "CF"], "so_diem_do": 3,
                     "theo_ma": {"1031M": {"yeu_cau": "310.0(0～+15mm)",
                                 "min": 310.5, "max": 313.5, "chuan": 310.0,
                                 "ds_duoi": 0, "ds_tren": 15,
                                 "lsl": 310.0, "usl": 325.0}, ...}},
                    ...
                ]
            }, ...
        }
    So diem do va ky hieu tung diem lay THANG tu file (moi chi tieu la mot
    khoi nhieu hang, moi hang mot vi tri do) - khong cai cung trong code.
    Ten chi tieu lay nguyen van, khong tu dat lai.
    """
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    ket_qua = {}

    for ten_sheet in wb.sheetnames:
        ws = wb[ten_sheet]

        ma_list, cot_ma = [], []
        for c in range(6, ws.max_column + 1, 3):
            val = ws.cell(7, c).value
            if val is not None and str(val).strip():
                ma_list.append(str(val).strip())
                cot_ma.append(c)
        if not ma_list:
            continue

        hang_stt = []
        for r in range(10, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
                if ws.cell(r, 3).value is not None:
                    hang_stt.append(r)

        chi_tieu_list = []
        for i, r in enumerate(hang_stt):
            r_ke = hang_stt[i + 1] if i + 1 < len(hang_stt) else None
            n = max(1, _so_hang_khoi(ws, r, r_ke))
            ten = str(ws.cell(r, 3).value).strip()
            kh_diem = []
            for x in range(r, r + n):
                k = ws.cell(x, 5).value
                kh_diem.append(str(k).strip() if k is not None else "")
            # o gop: chi hang dau co ky hieu -> cac hang sau dung chung
            if not any(kh_diem[1:]):
                kh_diem = [kh_diem[0]] * n

            theo_ma = {}
            for idx, ma in enumerate(ma_list):
                c = cot_ma[idx]
                raw = ws.cell(r, c).value
                yeu_cau = _chu_yeu_cau(raw)
                chuan, ds_duoi, ds_tren = _phan_tich_dung_sai(
                    raw if not isinstance(raw, (int, float)) else repr(float(raw)))
                if isinstance(raw, (int, float)):
                    chuan = round(float(raw), 2)
                lsl = chuan + ds_duoi if chuan is not None and ds_duoi is not None else None
                usl = chuan + ds_tren if chuan is not None and ds_tren is not None else None

                def _so(v):
                    try:
                        return float(v) if v is not None and str(v).strip() != "" else None
                    except (ValueError, TypeError):
                        return None

                theo_ma[ma] = {
                    "yeu_cau": yeu_cau,
                    "min": _so(ws.cell(r, c + 1).value),
                    "max": _so(ws.cell(r, c + 2).value),
                    "chuan": chuan, "ds_duoi": ds_duoi, "ds_tren": ds_tren,
                    "lsl": round(lsl, 4) if lsl is not None else None,
                    "usl": round(usl, 4) if usl is not None else None,
                }

            chi_tieu_list.append({
                "stt": int(float(ws.cell(r, 2).value)),
                "ten": ten,
                "ky_hieu": kh_diem[0],
                "ky_hieu_diem": kh_diem,
                "so_diem_do": n,
                "theo_ma": theo_ma,
            })

        ket_qua[ten_sheet] = {"ma_san_pham": ma_list, "chi_tieu": chi_tieu_list}

    wb.close()
    return ket_qua


# ---------------------------------------------------------------------------
# 2. CHAM SPC DUNG NHU CONG THUC BIEU MAU
# ---------------------------------------------------------------------------
def _round_excel(x, n):
    """ROUND cua Excel: lam tron nua len (xa so 0), khong phai banker's."""
    k = 10 ** n
    return math.floor(abs(x) * k + 0.5 + 1e-9) / k * (1 if x >= 0 else -1)


def chi_so_bieu_mau(X, lsl=None, usl=None):
    """Tinh dung cac o AL..BC cua hang x trong sheet 'Du lieu goc'."""
    base = spc_luat.tinh(X, lsl, usl)
    X = base.X
    n = len(X)
    xb = math.fsum(X) / n
    MR = [abs(X[i] - X[i - 1]) for i in range(1, n)]
    mrb = math.fsum(MR) / len(MR)
    sg = _round_excel(mrb / D2, 2)
    d = {"X": list(X), "MR": MR, "xb": xb, "mrb": mrb, "sigma": sg,
         "u3": _round_excel(xb + 3 * sg, 1), "u2": _round_excel(xb + 2 * sg, 1),
         "u1": _round_excel(xb + 1 * sg, 1), "l1": _round_excel(xb - 1 * sg, 1),
         "l2": _round_excel(xb - 2 * sg, 1), "l3": _round_excel(xb - 3 * sg, 1),
         "ucl_mr": _round_excel(D4 * mrb, 2), "ca": None, "cp": None,
         "cpk": None, "cpl": None, "cpu": None}
    if sg > 0:
        if lsl is not None:
            d["cpl"] = (xb - lsl) / (3 * sg)
        if usl is not None:
            d["cpu"] = (usl - xb) / (3 * sg)
        if lsl is not None and usl is not None and usl > lsl:
            to, t = (lsl + usl) / 2.0, usl - lsl
            d["ca"] = (xb - to) / (t / 2.0)
            d["cp"] = t / (6 * sg)
            d["cpk"] = (1 - abs(d["ca"])) * d["cp"]
        elif lsl is not None:
            d["cpk"] = d["cpl"]                  # mot phia duoi
        elif usl is not None:
            d["cpk"] = d["cpu"]
    return d


def cham_bieu_mau(X, lsl=None, usl=None, cpk_toi_thieu=CPK_TOI_THIEU, d=None):
    """Cham 8 muc I-chart + 2 muc MR-chart in tren sheet 'Bao cao', theo
    dung cac duong gioi han DA LAM TRON cua bieu mau, cong Cpk.

    Tra ve spc_luat.KetQuaSPC (dung chung voi phan con lai cua phan mem).
    Chi tieu khong co dung sai (kich thuoc tham khao) thi khong xet Cpk.
    """
    lsl, usl, cpk_toi_thieu = spc_luat.kiem_tra_nguong(lsl, usl, cpk_toi_thieu)
    if d is None:
        d = chi_so_bieu_mau(X, lsl, usl)
    r = spc_luat.KetQuaSPC(
        n=len(X), X=list(X), MR=d["MR"], xb=d["xb"], mrb=d["mrb"],
        sigma=d["sigma"], ucl_i=d["u3"], lcl_i=d["l3"], ucl_mr=d["ucl_mr"],
        lcl_mr=0.0, lsl=lsl, usl=usl, cpk=d["cpk"], cp=d["cp"], ca=d["ca"],
        loi=[])
    if d["sigma"] <= 0:
        r.loi = ["SIGMA"] + (["I7"] if len(X) >= 15 else [])
        if any((lsl is not None and x < lsl) or (usl is not None and x > usl) for x in X):
            r.loi.append("SPEC")
        return r
    xb = d["xb"]
    # Moi so sanh deu co dung sai SAI_SO: diem cach duong duoi muc do coi
    # nhu NAM TREN duong. So do lam tron 0,1 nen x (trung binh 2-3 diem) rat
    # hay rot dung tren duong +-1 sigma; Python va Excel cong so thuc lech
    # nhau o bit cuoi, khong co dung sai thi ben nay tinh trong vung, ben
    # kia tinh ngoai vung. Quy uoc: 'vuot duong' phai vuot han; 'ngoai
    # vung C' tinh ca diem nam tren duong.
    tren = lambda x, g: x > g + SAI_SO      # noqa: E731
    duoi = lambda x, g: x < g - SAI_SO      # noqa: E731
    loi = []

    if any(tren(x, d["u3"]) or duoi(x, d["l3"]) for x in X):
        loi.append("I1")
    if any(sum(tren(X[j], d["u2"]) for j in range(i, i + 3)) >= 2 or
           sum(duoi(X[j], d["l2"]) for j in range(i, i + 3)) >= 2
           for i in range(len(X) - 2)):
        loi.append("I2")
    if any(sum(tren(X[j], d["u1"]) for j in range(i, i + 5)) >= 4 or
           sum(duoi(X[j], d["l1"]) for j in range(i, i + 5)) >= 4
           for i in range(len(X) - 4)):
        loi.append("I3")
    run, dau = 0, 0
    for x in X:
        s = 1 if tren(x, xb) else (-1 if duoi(x, xb) else 0)
        run = 0 if s == 0 else (run + 1 if s == dau else 1)
        dau = s
        if run >= 9:
            loi.append("I4"); break
    if _chuoi_don_dieu(X, 6):
        loi.append("I5")
    alt, truoc = 1, 0
    for i in range(1, len(X)):
        s = tren(X[i], X[i - 1]) - duoi(X[i], X[i - 1])
        alt = 1 if s == 0 else (alt + 1 if truoc and s == -truoc else 2)
        truoc = s
        if alt >= 14:
            loi.append("I6"); break
    run = 0
    for x in X:
        run = run + 1 if (tren(x, d["l1"]) and duoi(x, d["u1"])) else 0
        if run >= 15:
            loi.append("I7"); break
    run = 0
    for x in X:
        run = run + 1 if not (tren(x, d["l1"]) and duoi(x, d["u1"])) else 0
        if run >= 8:
            loi.append("I8"); break

    if any(tren(m, d["ucl_mr"]) for m in d["MR"]):
        loi.append("MR1")
    if _chuoi_don_dieu(d["MR"], 6):
        loi.append("MR2")

    if d["cpk"] is not None and d["cpk"] < cpk_toi_thieu:
        loi.append("Cpk")
    if lsl is not None and duoi(min(X), lsl) or usl is not None and tren(max(X), usl):
        loi.append("SPEC")
    r.loi = sorted(set(loi), key=lambda m: spc_luat.THU_TU.index(m)
                   if m in spc_luat.THU_TU else 99)
    return r


def _chuoi_don_dieu(A, so_diem):
    tang = giam = 1
    for i in range(1, len(A)):
        if A[i] > A[i - 1] + SAI_SO:
            tang += 1; giam = 1
        elif A[i] < A[i - 1] - SAI_SO:
            giam += 1; tang = 1
        else:
            tang = giam = 1
        if tang >= so_diem or giam >= so_diem:
            return True
    return False


def cham_chi_tieu(ct, cpk_toi_thieu=CPK_TOI_THIEU):
    """SPC tren trung binh; nghiem thu tren moi diem do rieng le."""
    lsl, usl, cpk = spc_luat.kiem_tra_nguong(ct.get("lsl"), ct.get("usl"), cpk_toi_thieu)
    X = ct.get("X", ct.get("xbar", []))
    raw = ct.get("du_lieu", [])
    so_diem = ct.get("so_diem_do")
    try:
        if ct.get("loi_du_lieu") or len(X) != SO_MAU:
            raise ValueError("Thieu so do")
        if raw:
            if len(raw) != len(X) or any(not p or (so_diem and len(p) != so_diem) for p in raw):
                raise ValueError("Thieu diem do")
            for p in raw:
                for v in p:
                    spc_luat.so_huu_han(v)
            raw = [[float(v) for v in p] for p in raw]
            X = [math.fsum(p) / len(p) for p in raw]
        r = cham_bieu_mau(X, lsl, usl, cpk)
    except ValueError:
        r = spc_luat.KetQuaSPC(**{k: None for k in spc_luat.KetQuaSPC.__slots__})
        r.n, r.X, r.MR, r.lsl, r.usl = len(X), X, [], lsl, usl
        r.loi = ["DU_LIEU"]
        return r
    if any((lsl is not None and v < lsl - SAI_SO) or
           (usl is not None and v > usl + SAI_SO) for p in raw for v in p):
        if "SPEC" not in r.loi:
            r.loi.append("SPEC")
    return r


# ---------------------------------------------------------------------------
# 3. SINH DU LIEU NGAU NHIEN CO KIEM SOAT
# ---------------------------------------------------------------------------
def _sinh_mau_gauss(mu, sigma, lo, hi, n):
    """Sinh n gia tri phan phoi chuan N(mu, sigma^2), clamp trong [lo, hi]."""
    vals = []
    for _ in range(n):
        v = random.gauss(mu, sigma)
        v = max(lo, min(hi, v))
        vals.append(round(v, 1))
    return vals


def sinh_du_lieu_1_chi_tieu(min_val, max_val, lsl, usl, so_mau=SO_MAU,
                             so_diem_do=1, cpk_toi_thieu=CPK_TOI_THIEU,
                             seed=None):
    """Sinh du lieu cho 1 chi tieu kich thuoc.

    Moi mau co so_diem_do diem do, lam tron 0,1 nhu so do that, nam trong
    [Min, Max]. Gia tri dua vao I-MR la TRUNG BINH KHONG LAM TRON cua cac
    diem (dung o x = AVERAGE(...) cua bieu mau).

    Tra ve (x_list, du_lieu_diem_do, ket_qua_spc).
    """
    if seed is not None:
        random.seed(seed)
    if min_val is None or max_val is None:
        raise ValueError("Min va Max khong duoc de trong")
    min_val = spc_luat.so_huu_han(min_val, "Min")
    max_val = spc_luat.so_huu_han(max_val, "Max")
    if min_val >= max_val:
        raise ValueError("Min phai nho hon Max de co bien thien.")
    so_mau = spc_luat.so_nguyen(so_mau, "So mau", 2)
    so_diem_do = spc_luat.so_nguyen(so_diem_do, "So diem do")
    lsl, usl, cpk_toi_thieu = spc_luat.kiem_tra_nguong(lsl, usl, cpk_toi_thieu)

    tam = (min_val + max_val) / 2
    ban_kinh = (max_val - min_val) / 2
    # Do phan tan toi thieu bang mot buoc do (0,1 mm): khoang Min-Max hep
    # (vd 28,1-28,3) ma lay 1/6 khoang thi moi diem lam tron ve cung mot so,
    # 25 mau trung nhau, sigma = 0 va bieu mau chia cho 0.
    sigma_x = max(ban_kinh / 6.0, BUOC_DO * 0.6)
    sigma_sub = max(ban_kinh / 8.0, BUOC_DO * 0.5) if so_diem_do > 1 else 0

    def phat(mau):
        X = [math.fsum(p) / len(p) for p in mau]
        d = chi_so_bieu_mau(X, lsl, usl)          # tinh mot lan, dung hai cho
        r = cham_bieu_mau(X, lsl=lsl, usl=usl, cpk_toi_thieu=cpk_toi_thieu, d=d)
        return len(r.loi), r

    tot = None                                   # (diem phat, mau, ket qua)
    for _ in range(SO_LAN_SINH_LAI):
        mau = []
        for _ in range(so_mau):
            t = max(min_val, min(max_val, random.gauss(tam, sigma_x)))
            if so_diem_do == 1:
                mau.append([round(t, 1)])
            else:
                mau.append(_sinh_mau_gauss(t, sigma_sub, min_val, max_val,
                                           so_diem_do))
        p, r = phat(mau)
        # Sap lai THU TU 25 mau (giu nguyen tap so do): cac luat I-chart va
        # MR-chart phu thuoc thu tu, va sigma = MR-bar/1.128 cung vay. Doi
        # cho hai mau, giu lai neu khong te hon - leo doi dan toi 0 loi.
        for _ in range(SO_LAN_DOI_CHO):
            if p == 0:
                break
            i, j = random.sample(range(so_mau), 2)
            mau[i], mau[j] = mau[j], mau[i]
            p2, r2 = phat(mau)
            if p2 <= p:
                p, r = p2, r2
            else:
                mau[i], mau[j] = mau[j], mau[i]
        if tot is None or p < tot[0]:
            tot = (p, [list(x) for x in mau], r)
        if p == 0:
            break
    _p, du_lieu_sub, r = tot
    x_list = [math.fsum(p) / len(p) for p in du_lieu_sub]
    return x_list, du_lieu_sub, r


def sinh_toan_bo(danh_sach_chi_tieu, so_mau=SO_MAU, callback=None, seed=None,
                 cpk_toi_thieu=CPK_TOI_THIEU):
    """Sinh du lieu dong thoi cho tat ca chi tieu.

    danh_sach_chi_tieu: list cua dict, moi dict co it nhat
        {"stt", "ten", "ky_hieu", "min", "max", "lsl", "usl", "so_diem_do"}
    cac khoa khac (yeu_cau, chuan, ds_duoi, ds_tren, ky_hieu_diem...) duoc
    giu nguyen sang ket qua de buoc xuat file dung.
    callback: ham callback(stt, ten, dat, lan_thu, thong_bao)
    """
    if seed is not None:
        random.seed(seed)

    ket_qua = []
    for ct in danh_sach_chi_tieu:
        stt, ten = ct["stt"], ct.get("ten", "")
        min_v, max_v = ct["min"], ct["max"]
        so_diem_do = ct.get("so_diem_do", 1)
        if min_v is None or max_v is None:
            if callback:
                callback(stt, ten, False, 0, "BO QUA: Min/Max chua xac dinh")
            continue

        x, du_lieu, r = sinh_du_lieu_1_chi_tieu(
            min_v, max_v, ct.get("lsl"), ct.get("usl"), so_mau, so_diem_do,
            cpk_toi_thieu=cpk_toi_thieu)

        kq = dict(ct)
        kq.update({"xbar": x, "du_lieu": du_lieu,
                   "so_diem_do": so_diem_do, "cpk_toi_thieu": cpk_toi_thieu})
        r = kq["spc_result"] = cham_chi_tieu(kq, cpk_toi_thieu)

        if callback:
            if r.dat:
                trang_thai = "DAT" + ("" if r.cpk is None else
                                      "  (Cpk %.2f)" % r.cpk)
            else:
                trang_thai = "CHUA DAT sau %d lan: %s" % (
                    MAX_RETRY, ", ".join(r.loi))
            callback(stt, ten, r.dat, 0, trang_thai)

        ket_qua.append(kq)
    return ket_qua
