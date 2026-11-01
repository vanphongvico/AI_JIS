# -*- coding: utf-8 -*-
"""
THANH PHAN HOA CORE - SINH DU LIEU THANH PHAN HOA & DANH GIA SPC
=================================================================
Doc file du lieu chuan "Du lieu 2.xlsx" (moi mac thep mot khoi: dong ten
nguyen to, roi 3 dong Yeu cau / Min / Max), sinh N mau thanh phan hoa
gia lap, cham SPC I-MR cho 10 nguyen to cua bieu mau JIS.QT.07.BM.01
theo tung nhom 25 mau, va sinh 3 lan do cho ban in may quang pho.

Lam tron cua bieu mau (ROUND(sigma,2), ROUND(gioi han,1)) khong dung
duoc cho thanh phan hoa (S ~ 0,003 -> sigma = 0). O day lam tron theo so
chu so d cua tung nguyen to:
    sigma = ROUND(MR-bar/1.128, d+2)
    X-bar +- k sigma = ROUND(..., d+1)
    UCLMR = ROUND(3.267*MR-bar, d+2)
File SPC xuat ra (xuat_thanh_phan_spc) ghi DUNG cac cong thuc nay.
"""

import math
import random
import re
import unicodedata

import openpyxl

try:
    import spc_luat
    import kich_thuoc_core as kt
except ImportError:
    from . import spc_luat
    from . import kich_thuoc_core as kt

SO_MAU_NHOM = 25
CPK_TOI_THIEU = 1.33
D2, D4 = 1.128, 3.267

# 10 nguyen to tren sheet "Du lieu goc" cua bieu mau, dung thu tu hang 10-19
PHAN_TU_SPC = ("C", "Si", "Mn", "P", "S", "Ni", "Cr", "Mo", "V", "Cu")
# So chu so in tren ban in may quang pho (Ket qua in.docx)
SO_LE_MAY = {"S": 4, "B": 4, "Ce": 4, "Mg": 4, "Nb": 4, "Pb": 4, "Sn": 4,
             "Ti": 4, "W": 4, "Zn": 4, "N": 4, "Ca": 4, "Fe": 2}
SO_LE_MAC_DINH = 3
HE_SO_CEQ = (("C", 1.0), ("Mn", 6.0), ("Si", 24.0), ("Ni", 40.0),
             ("Cr", 5.0), ("Mo", 4.0), ("V", 14.0))

SO_LAN_SINH_LAI = 40
SO_LAN_DOI_CHO = 800


class LoiDuLieu(Exception):
    pass


def _nfc(s):
    return unicodedata.normalize("NFC", str(s)).strip()


_SO = r"(\d+(?:\.\d+)?)"


def phan_tich_yeu_cau(v):
    """'≤ 0.22' -> (None, 0.22); '≥ x' -> (x, None); 'a ~ b' -> (a, b);
    '-' / trong -> (None, None); so 0.4 -> (None, 0.4) (Ceq ghi so)."""
    if v is None:
        return None, None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None, float(v)
    s = _nfc(v).replace(",", ".")
    if s in ("", "-", "–"):
        return None, None
    m = re.match(r"^(?:≤|<=|<)\s*" + _SO + r"$", s)
    if m:
        return None, float(m.group(1))
    m = re.match(r"^(?:≥|>=|>)\s*" + _SO + r"$", s)
    if m:
        return float(m.group(1)), None
    m = re.match(r"^" + _SO + r"\s*(?:~|÷|–|-)\s*" + _SO + r"$", s)
    if m:
        return float(m.group(1)), float(m.group(2))
    try:
        return None, float(s)
    except ValueError:
        raise LoiDuLieu("Không hiểu yêu cầu kỹ thuật '%s'." % v)


def _so_le_cua(v):
    s = ("%.10f" % float(v)).rstrip("0")
    return len(s.split(".")[1]) if "." in s else 0


def so_le(ten, mn, mx):
    """So chu so = max(so chu so cua Min/Max, so chu so in tren ban in may)."""
    if ten == "Fe":
        return 2
    return max(_so_le_cua(mn), _so_le_cua(mx), SO_LE_MAY.get(ten, SO_LE_MAC_DINH))


def doc_chuan(duong_dan):
    """{mac: {"nguyen_to": [...], "ceq": {...} | None}} - do khoi theo chu
    o cot A (dong ten mac, ngay sau la dong 'Yeu cau'), khong co dinh so dong."""
    wb = openpyxl.load_workbook(duong_dan, data_only=False)
    ws = wb.worksheets[0]
    ra = {}
    nhan_dong = ("yêu cầu", "min", "max")
    for r in range(1, ws.max_row + 1):
        mac = ws.cell(r, 1).value
        if not isinstance(mac, str) or not mac.strip() or ws.cell(r, 2).value is None:
            continue
        # Dong ten mac: chu o cot A khong phai nhan Yeu cau/Min/Max, va DONG
        # NGAY DUOI la 'Yeu cau' (khoi cach nhau 1 dong trong cung doc dung).
        if _nfc(mac).lower() in nhan_dong or \
                _nfc(ws.cell(r + 1, 1).value or "").lower() != "yêu cầu":
            continue
        nhan = {}
        for k in (2, 3):
            v = ws.cell(r + k, 1).value
            if v is not None and _nfc(v).lower() in nhan_dong:
                nhan[_nfc(v).lower()] = r + k
        r_yc, r_min, r_max = r + 1, nhan.get("min"), nhan.get("max")
        if not (r_min and r_max):
            raise LoiDuLieu("Khối %s thiếu dòng Min hoặc Max." % mac.strip())
        ds, ceq = [], None
        for c in range(2, ws.max_column + 1):
            ten = ws.cell(r, c).value
            if ten is None or not str(ten).strip():
                break
            ten = _nfc(ten)
            yc = ws.cell(r_yc, c).value
            lsl, usl = phan_tich_yeu_cau(yc)
            if ten.lower() == "ceq":
                ceq = {"ten": "Ceq", "yeu_cau": yc, "lsl": lsl, "usl": usl}
                continue
            mn, mx = ws.cell(r_min, c).value, ws.cell(r_max, c).value
            nt = {"ten": ten, "yeu_cau": yc, "lsl": lsl, "usl": usl,
                  "min": None, "max": None, "so_le": None, "chu": None}
            if ten == "Fe":
                nt["so_le"] = 2
            elif isinstance(mn, str) or isinstance(mx, str):
                nt["chu"] = _nfc(mn if isinstance(mn, str) else mx)
            elif mn is None or mx is None:
                raise LoiDuLieu("%s / %s: thiếu Min hoặc Max." % (mac.strip(), ten))
            else:
                nt["min"], nt["max"] = float(mn), float(mx)
                nt["so_le"] = so_le(ten, mn, mx)
            ds.append(nt)
        ra[_nfc(mac)] = {"nguyen_to": ds, "ceq": ceq}
    wb.close()
    if not ra:
        raise LoiDuLieu("Không tìm thấy khối mác thép nào trong file dữ liệu chuẩn "
                        "(cột A: tên mác, dòng dưới: 'Yêu cầu', 'Min', 'Max').")
    return ra


def kiem_tra_chuan(ds):
    """Bao loi tieng Viet neu Min/Max khong hop le. Khong tra ve gi."""
    co = {n["ten"]: n for n in ds}
    thieu = [t for t in PHAN_TU_SPC if t not in co]
    if thieu:
        raise LoiDuLieu("Dữ liệu chuẩn thiếu nguyên tố cần đánh giá SPC: %s."
                        % ", ".join(thieu))
    if "Fe" not in co:
        raise LoiDuLieu("Dữ liệu chuẩn thiếu cột Fe.")
    for nt in ds:
        ten = nt["ten"]
        if ten in PHAN_TU_SPC and nt["chu"] is not None:
            raise LoiDuLieu("%s phải là số (đang là '%s') để đánh giá SPC."
                            % (ten, nt["chu"]))
        if nt["min"] is None:
            continue
        mn, mx = nt["min"], nt["max"]
        if mn < 0:
            raise LoiDuLieu("%s: Min không được âm (%g)." % (ten, mn))
        if mn > mx:
            raise LoiDuLieu("%s: Min (%g) lớn hơn Max (%g)." % (ten, mn, mx))
        if ten in PHAN_TU_SPC and mn == mx:
            raise LoiDuLieu("%s: Min = Max = %g - SPC cần có biến thiên." % (ten, mn))
        if nt["usl"] is not None and mx > nt["usl"]:
            raise LoiDuLieu("%s: Max %g vượt yêu cầu %s." % (ten, mx, nt["yeu_cau"]))
        if nt["lsl"] is not None and mn < nt["lsl"]:
            raise LoiDuLieu("%s: Min %g vượt yêu cầu %s." % (ten, mn, nt["yeu_cau"]))


# ---------------------------------------------------------------------------
# 2. CHI SO SPC (lam tron theo so chu so) VA SINH MAU
# ---------------------------------------------------------------------------
def chi_so(X, lsl, usl, so_le):
    """Dung cong thuc ghi vao file SPC xuat ra (xem dau file)."""
    r = kt._round_excel
    n = len(X)
    xb = math.fsum(X) / n
    MR = [abs(X[i] - X[i - 1]) for i in range(1, n)]
    mrb = math.fsum(MR) / len(MR)
    sg = r(mrb / D2, so_le + 2)
    g = so_le + 1
    d = {"X": list(X), "MR": MR, "xb": xb, "mrb": mrb, "sigma": sg,
         "u3": r(xb + 3 * sg, g), "u2": r(xb + 2 * sg, g), "u1": r(xb + sg, g),
         "l1": r(xb - sg, g), "l2": r(xb - 2 * sg, g), "l3": r(xb - 3 * sg, g),
         "ucl_mr": r(D4 * mrb, so_le + 2), "ca": None, "cp": None, "cpk": None}
    if sg > 0:
        if lsl is not None and usl is not None and usl > lsl:
            to, t = (lsl + usl) / 2.0, usl - lsl
            d["ca"] = (xb - to) / (t / 2.0)
            d["cp"] = t / (6 * sg)
            d["cpk"] = (1 - abs(d["ca"])) * d["cp"]
        elif usl is not None:
            d["cpk"] = (usl - xb) / (3 * sg)
        elif lsl is not None:
            d["cpk"] = (xb - lsl) / (3 * sg)
    return d


def cham(X, lsl, usl, so_le, cpk_toi_thieu=CPK_TOI_THIEU):
    return kt.cham_bieu_mau(X, lsl, usl, cpk_toi_thieu, d=chi_so(X, lsl, usl, so_le))


SAT_GIOI_HAN = 0.9      # diem qua 90 % khoang tam -> gioi han kiem soat: trong nhu cham


def _so_sat(X, so_le):
    """So diem nam sat gioi han kiem soat (chua vuot): MR > 90 % UCLMR, hoac I
    qua 90 % khoang X-bar -> UCLI/LCLI. Dung duong da lam tron nhu bieu mau."""
    d = chi_so(X, None, None, so_le)
    xb, k = d["xb"], SAT_GIOI_HAN
    return (sum(1 for m in d["MR"] if m > k * d["ucl_mr"]) +
            sum(1 for x in X if x > xb + k * (d["u3"] - xb) or x < xb - k * (xb - d["l3"])))


def loi_hinh_dang(X, so_le):
    """Hien tuong bieu mau KHONG cham nhung nhin bieu do thay nhu khong dat:
      'MR0' - hai mau lien nhau trung so: diem MR = 0 nam de len duong LCLMR;
      'MR9' - 9 diem MR lien tiep cung phia duong tam CLMR (trong nhu bi dich);
      'SAT' - diem sat gioi han kiem soat (> 90 % khoang tu duong tam).
    Chi dung khi SINH so; ket qua SPC in tren bieu mau van theo cham()."""
    MR = [abs(b - a) for a, b in zip(X, X[1:])]
    if not MR:
        return []
    loi = []
    if any(m < 0.5 * 10 ** -so_le for m in MR):
        loi.append("MR0")
    tb, ss = math.fsum(MR) / len(MR), kt.SAI_SO
    run = dau = 0
    for m in MR:
        s = 1 if m > tb + ss else (-1 if m < tb - ss else 0)
        run = 0 if s == 0 else (run + 1 if s == dau else 1)
        dau = s
        if run >= 9:
            loi.append("MR9")
            break
    if _so_sat(X, so_le):
        loi.append("SAT")
    return loi


def _la_so(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def tinh_fe(mau):
    return round(100 - math.fsum(v for k, v in mau.items() if k != "Fe" and _la_so(v)), 2)


def tinh_ceq(mau):
    return math.fsum(mau[k] / h for k, h in HE_SO_CEQ if _la_so(mau.get(k)))


def sinh_cot_spc(nt, n, cpk_toi_thieu, rng):
    """n gia tri cho mot nguyen to SPC, dat het luat neu tim duoc.
    Sinh Gauss quanh tam khoang roi doi cho (giong kich_thuoc_core)."""
    mn, mx, d = nt["min"], nt["max"], nt["so_le"]
    tam, bk = (mn + mx) / 2.0, (mx - mn) / 2.0
    sg = max(bk / 3.0, 0.6 * 10 ** -d)

    def mot():
        return round(max(mn, min(mx, rng.gauss(tam, sg))), d)

    def day():
        """n gia tri, mau sau khong trung mau ngay truoc (MR = 0)."""
        X = []
        for _ in range(n):
            v = mot()
            for _ in range(20):
                if not X or v != X[-1]:
                    break
                v = mot()
            X.append(v)
        return X

    nua = 0.5 * 10 ** -d

    def phat(X):
        # Luat bieu mau + hinh dang bieu do (diem MR = 0, chuoi 9 MR cung phia,
        # diem sat gioi han kiem soat)
        r = cham(X, nt["lsl"], nt["usl"], d, cpk_toi_thieu)
        so_0 = sum(1 for a, b in zip(X, X[1:]) if abs(b - a) < nua)
        return len(r.loi) + so_0 + ("MR9" in loi_hinh_dang(X, d)) + _so_sat(X, d), r

    tot = None
    for _ in range(SO_LAN_SINH_LAI):
        X = day()
        p, r = phat(X)
        for _ in range(SO_LAN_DOI_CHO):
            if p == 0:
                break
            i, j = rng.sample(range(n), 2)
            X[i], X[j] = X[j], X[i]
            p2, r2 = phat(X)
            if p2 <= p:
                p, r = p2, r2
            else:
                X[i], X[j] = X[j], X[i]
        if tot is None or p < tot[0]:
            tot = (p, list(X), r)
        if p == 0:
            break
    return tot[1], tot[2]


def sinh_toan_bo(ds, so_mau, cpk_toi_thieu=CPK_TOI_THIEU, seed=None, ceq=None,
                 callback=None, dung=None):
    """Sinh N mau, chia nhom 25; moi nhom cham SPC rieng 10 nguyen to.

    Tra ve {"mau": [dict]*N, "nhom": [{"dau","cuoi","spc": {ten: KetQuaSPC}}],
            "nguyen_to": ds, "ceq": ceq, "dat": bool}.
    Nhom chi co 1 mau thi khong cham SPC (I-MR can >= 2 diem).
    """
    so_mau = spc_luat.so_nguyen(so_mau, "Số mẫu", 1)
    kiem_tra_chuan(ds)
    kiem_tra_ceq(ds, ceq)
    rng = random.Random(seed)
    mau = []
    nhom = []
    for dau in range(0, so_mau, SO_MAU_NHOM):
        cuoi = min(dau + SO_MAU_NHOM, so_mau)
        # Ceq ghep tu 7 nguyen to SPC: doi mot mau se pha SPC cua ca cot, nen
        # mau nao vuot Ceq thi sinh lai CA NHOM.
        for _ in range(SO_LAN_THU_CEQ):
            mau_nhom, spc = _sinh_nhom(ds, cuoi - dau, cpk_toi_thieu, rng)
            sai = [k for k, m in enumerate(mau_nhom) if _ngoai_ceq(m, ceq)]
            if not sai:
                break
        else:
            raise LoiDuLieu("Sau %d lần sinh lại, nhóm mẫu %d-%d vẫn có mẫu Ceq vượt yêu "
                            "cầu %s. Hãy thu hẹp Max của C, Mn, Cr…"
                            % (SO_LAN_THU_CEQ, dau + 1, cuoi, ceq["yeu_cau"]))
        mau.extend(mau_nhom)
        nhom.append({"dau": dau, "cuoi": cuoi, "spc": spc})
        if callback:
            callback(cuoi, so_mau, spc)
        if dung and dung():
            raise LoiDuLieu("Đã dừng theo yêu cầu.")
    return {"mau": mau, "nhom": nhom, "nguyen_to": ds, "ceq": ceq,
            "dat": all(r.dat for g in nhom for r in g["spc"].values())}


SO_LAN_THU_CEQ = 30


def _sinh_nhom(ds, n, cpk_toi_thieu, rng):
    """n mau (list dict) + {ten: KetQuaSPC} cho mot nhom."""
    mau = [dict() for _ in range(n)]
    spc = {}
    for nt in ds:
        ten = nt["ten"]
        if ten == "Fe":
            continue
        if nt["chu"] is not None:
            X = [nt["chu"]] * n
        elif ten in PHAN_TU_SPC and n >= 2:
            X, r = sinh_cot_spc(nt, n, cpk_toi_thieu, rng)
            spc[ten] = r
        else:
            X = [round(rng.uniform(nt["min"], nt["max"]), nt["so_le"]) for _ in range(n)]
        for m, v in zip(mau, X):
            m[ten] = v
    for m in mau:
        m["Fe"] = tinh_fe(m)
    return mau, spc


def _ngoai_ceq(mau, ceq):
    if not ceq:
        return False
    v = tinh_ceq(mau)
    return (ceq["usl"] is not None and v > ceq["usl"] + 1e-12) or \
        (ceq["lsl"] is not None and v < ceq["lsl"] - 1e-12)


def kiem_tra_ceq(ds, ceq):
    """Bao loi NGAY (truoc khi sinh) neu Ceq khong the dat: moi nguyen to o
    Min ma Ceq van > gioi han tren (hoac o Max ma van < gioi han duoi)."""
    if not ceq:
        return
    mn = {n["ten"]: n["min"] for n in ds if n["min"] is not None}
    mx = {n["ten"]: n["max"] for n in ds if n["max"] is not None}
    if ceq["usl"] is not None and tinh_ceq(mn) > ceq["usl"] + 1e-12:
        raise LoiDuLieu("Ceq không thể đạt yêu cầu %s: ngay cả khi mọi nguyên tố ở Min, "
                        "Ceq = %.3f. Hãy giảm Min của C, Mn, Cr…" % (ceq["yeu_cau"], tinh_ceq(mn)))
    if ceq["lsl"] is not None and tinh_ceq(mx) < ceq["lsl"] - 1e-12:
        raise LoiDuLieu("Ceq không thể đạt yêu cầu %s: ngay cả khi mọi nguyên tố ở Max, "
                        "Ceq = %.3f." % (ceq["yeu_cau"], tinh_ceq(mx)))


# ---------------------------------------------------------------------------
# 3. BA LAN DO CHO BAN IN MAY QUANG PHO
# ---------------------------------------------------------------------------
def _rsd_muc_tieu(a, rng):
    """RSD (%) hop ly theo muc ham luong - nguyen to vet dao dong nhieu hon."""
    if a >= 0.1:
        return rng.uniform(0.1, 2.0)
    if a >= 0.01:
        return rng.uniform(1.0, 5.0)
    return rng.uniform(3.0, 15.0)


def _rsd(lan, a):
    m = math.fsum(lan) / len(lan)
    s = math.sqrt(math.fsum((x - m) ** 2 for x in lan) / (len(lan) - 1))
    return round(s / a * 100, 2) if a else 0.0


def ba_lan_do(a, so_le, rng):
    """Tinh theo DON VI BUOC (so nguyen) nen trung binh 3 lan BANG DUNG a:
    R3 = 3A - R1 - R2."""
    buoc = 10.0 ** -so_le
    A = int(round(a / buoc))
    if A <= 0:
        return [a] * 3, 0.0
    # It nhat ~1 buoc lam tron: RSD nho hon thi 3 lan do in ra trung nhau
    # (RSD 0.00) - ban in that hiem khi nhu vay.
    sd = max(A * _rsd_muc_tieu(a, rng) / 100.0, 0.8)
    for _ in range(50):
        R1 = int(round(A + rng.gauss(0, sd)))
        R2 = int(round(A + rng.gauss(0, sd)))
        R3 = 3 * A - R1 - R2
        if min(R1, R2, R3) > 0 and abs(R3 - A) <= 3 * sd + 1 and \
                (A < 5 or not R1 == R2 == R3):
            break
    else:
        R1 = R2 = R3 = A
    lan = [round(R * buoc, so_le) for R in (R1, R2, R3)]
    return lan, _rsd(lan, a)


def ba_lan_do_mau(mau, ds, rng):
    """{ten: {"lan": [3 lan], "tb": Average, "rsd": %, "so_le": d|None}}
    cho mot mau - ca Fe (= 100 - tong tung lan) va nguyen to chu."""
    ra = {}
    for nt in ds:
        ten = nt["ten"]
        if ten == "Fe":
            continue
        a = mau[ten]
        if nt["chu"] is not None:
            ra[ten] = {"lan": [a] * 3, "tb": a, "rsd": round(rng.uniform(5, 80), 2),
                       "so_le": None}
        else:
            lan, rsd = ba_lan_do(a, nt["so_le"], rng)
            ra[ten] = {"lan": lan, "tb": a, "rsd": rsd, "so_le": nt["so_le"]}
    ra["Fe"] = fe_ba_lan(ra, mau["Fe"])
    return ra


def fe_ba_lan(ra, fe):
    """3 lan do Fe theo 3 lan do cac nguyen to so trong `ra` (khong tinh Fe):
    {"lan", "tb", "rsd", "so_le"}. Dung chung cho ban in moi va khi sua so."""
    so = [k for k, v in ra.items() if k != "Fe" and _la_so(v["tb"])]
    F = int(round(fe * 100))
    # Fe tung lan = 100 - tong lan do do (don vi 0,01), nhung tong 3 lan phai = 3F
    # de trung binh dung bang Fe. Phan du lam tron chia cho lan dang lech it nhat
    # (truoc day don het vao lan 3 -> lan 3 cong lai lech 100 toi 0,025).
    e = [(100 - math.fsum(ra[k]["lan"][i] for k in so)) * 100 for i in range(3)]
    L = [int(round(x)) for x in e]
    buoc = 1 if 3 * F > sum(L) else -1
    for _ in range(abs(3 * F - sum(L))):
        i = max(range(3), key=lambda j: (e[j] - L[j]) * buoc)
        L[i] += buoc
    lan = [round(x / 100.0, 2) for x in L]
    return {"lan": lan, "tb": fe, "rsd": _rsd(lan, fe), "so_le": 2}
