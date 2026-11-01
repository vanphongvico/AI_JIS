# -*- coding: utf-8 -*-
"""
BO GHI PDF BIEU DO THU KEO (Load - Position)
=============================================
Ghi thang file PDF vector, KHONG dung thu vien ngoai.

Vi sao tu viet ma khong dung ReportLab:
  - Noi dung mot trang chi gom duong thang, duong Bezier va chu Helvetica.
    Helvetica la font chuan cua PDF nen khong phai nhung font.
  - Lop chu la thu Tab 2 doc nguoc lai de lay FeH/FeL/Fm. Tu viet thi kiem
    soat duoc chinh xac THU TU cac doan chu, dung cai ma
    tensile_core.doc_so_lieu_tren_trang() dang can.
  - Khong them thu vien -> CHAY_PHAN_MEM.bat va file .exe khong phinh them.

Hinh hoc lay dung theo bieu do mau trong 'S2-1 TENSILE TESTING 522.docx'
(anh chup man hinh may thu, 774 x 648 px). Moi toa do duoi day tinh bang
PIXEL CUA ANH MAU, goc tren-trai, truc y huong xuong - do thang tu anh ra
nen doi chieu duoc tung con so. 1 px = 1 pt tren trang PDF.
    Goc toa do (0 mm, 0 kN)  : x = 27, y = 617
    Vung ve                  : rong 712.5, cao 596.5
    Luoi                     : 2 mm x 5 kN, xam 128, net 1
    Truc                     : net 2, mui ten tam giac dac
Pham vi truc TU CO GIAN theo tung mau (mau goc: 0-36 mm, 0-90 kN), xem
pham_vi_truc().
"""

import math
import random
import zlib

# ---------------------------------------------------------------- hinh hoc
RONG_ANH, CAO_ANH = 774.0, 648.0     # dung kich thuoc anh mau
CAO_CHAN = 26.0                      # dai duoi chua dong 'Chart ID' (Tab 2 cat)
RONG_TRANG, CAO_TRANG = RONG_ANH, CAO_ANH + CAO_CHAN

GOC_X, GOC_Y = 27.0, 617.0           # pixel cua diem (0 mm, 0 kN)
RONG_VUNG, CAO_VUNG = 712.5, 596.5   # tu goc toi vach luoi cuoi cung
BUOC_LUOI_X, BUOC_LUOI_Y = 2.0, 5.0

XAM = 128 / 255.0
NET_LUOI, NET_TRUC, NET_CONG, NET_DAN = 1.0, 2.0, 1.0, 1.0
CO_CHU_TRUC = 11                     # so tren truc
CO_CHU = 12                          # nhan truc + khoi chu FeH/FeL/Fm
DONG_CACH = 14.0                     # khoang cach dong trong khoi chu
# Font may thu (MS Sans Serif) hep hon Helvetica: chu so va chu 'm' cua no
# chi bang ~87% be ngang. Ep ngang khoi chu va so tren truc cho khop.
EP_NGANG_CHU = 87
BAN_KINH_DIEM = 5.0
KAPPA = 0.5523                       # ve hinh tron bang 4 cung Bezier

# mau cham tron, do tu anh mau
MAU_FEH = (128 / 255.0, 1.0, 128 / 255.0)
MAU_FEL = (1.0, 1.0, 128 / 255.0)
MAU_FM = (1.0, 0.0, 0.0)

# Be rong ky tu Helvetica (phan nghin em) - chi can cho viec can le chu
_RONG = {" ": 278, ":": 278, ".": 278, "(": 333, ")": 333, "-": 333,
         "/": 278, "k": 500, "N": 722, "m": 833, "L": 611, "o": 556,
         "a": 556, "d": 556, "P": 667, "s": 500, "i": 222, "t": 278,
         "n": 556, "C": 722, "h": 556, "r": 333, "I": 278, "D": 722,
         "F": 611, "e": 556, "H": 722, "M": 833}
for _c in "0123456789":
    _RONG[_c] = 556


def _be_rong(s, co):
    return sum(_RONG.get(c, 556) for c in s) * co / 1000.0


def _thoat(s):
    return str(s).replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")


def pham_vi_truc(t):
    """(x_max mm, y_max kN) cua truc, tu co gian theo mau.

    Luat suy tu bieu do mau: Fm 80,634 kN -> truc 90; diem dut 30,79 mm ->
    truc 36. Tuc la tran Fm len boi so 5 roi them mot o luoi, va tran diem
    dut len boi so 2 roi them hai o luoi. Luat nay cung khop bo 1000 bieu do
    cu (Fm ~88 kN -> 95). Muon co dinh thi dat t['x_max'] / t['y_max'].
    """
    x_max = t.get("x_max") or (math.floor(t["dut_pos"] / BUOC_LUOI_X) * BUOC_LUOI_X
                               + 3 * BUOC_LUOI_X)
    y_max = t.get("y_max") or (math.floor(t["fm"] / BUOC_LUOI_Y) * BUOC_LUOI_Y
                               + 2 * BUOC_LUOI_Y)
    return float(x_max), float(y_max)


# ------------------------------------------------------------- duong cong
# Doan chay deo FeH -> FeL lay theo bieu do mau, CHUAN HOA:
#     s = (x - feh_pos)/(fel_pos - feh_pos)      0 tai FeH, 1 tai FeL
#     r = (y - fel)/(feh - fel)                  1 tai FeH, 0 tai FeL
# Tai dinh FeH tut gan thang dung xuong ~0,47 (chua toi mot nua do chenh),
# roi di xuong theo RANG CUA ba bon nhip, cham day dung tai FeL. FeL la
# diem THAP NHAT cua ca doan - moi diem khac giu r > 0 de khi doc lai
# bieu do, cham vang FeL van la day that su.
MAU_CHAY = ((0.00, 1.00), (0.03, 0.47), (0.08, 0.30), (0.13, 0.40),
            (0.18, 0.33), (0.23, 0.18), (0.28, 0.10), (0.33, 0.03),
            (0.38, 0.15), (0.43, 0.22), (0.48, 0.10), (0.53, 0.11),
            (0.58, 0.05), (0.63, 0.015), (0.68, 0.08), (0.72, 0.12),
            (0.77, 0.19), (0.82, 0.16), (0.90, 0.07), (1.00, 0.00))

# Doan dan hoi, chuan hoa v = x/feh_pos, w = y/feh:
#     w = A + B*(1 - exp(-v/TAU)) + (1 - A - B*(1 - exp(-1/TAU))) * v^mu
# Mot khuc "lot may" dung nhanh len ~3 kN roi ~7 kN trong 0,3 mm dau, sau
# do cong lom len toi FeH. Khop anh mau sai so RMS 0,21 kN (duoi 1 pixel).
DH_A, DH_B, DH_TAU = 0.0517, 0.1328, 0.1125

MU_DAN_HOI_MD = 1.80      # luy thua doan dan hoi
K_HOA_BEN_MD = 3.65       # he so ham bao hoa Voce doan hoa ben
MU_THAT_CO_MD = 3.45      # luy thua doan that co
CHAN_CHAY_MD = 0.20       # mm, khuc nam ngang sau FeL truoc khi hoa ben

SO_DIEM_DAN_HOI = 80
SO_DIEM_HOA_BEN = 160
SO_DIEM_THAT_CO = 120


def _hat(t):
    """Hat ngau nhien CO DINH theo mau: cung t -> cung rang cua.

    duong_cong() duoc goi hai lan cho mot mau (ve PDF va lay mau sheet
    Du_lieu) nen khong duoc phep ra hai duong khac nhau.
    """
    if t.get("hat_chay") is not None:
        return int(t["hat_chay"])
    return int(round(t["feh"] * 1000)) * 7919 + int(round(t["fel_pos"] * 1000))


def duong_cong(t):
    """Sinh day diem (position_mm, load_kN) cua mot duong cong thu keo.

    t la dict tham so cua mot mau, xem sinh_keo_core.tham_so_duong_cong().
    Chi bat buoc 9 khoa feh/feh_pos/fel/fel_pos/fm/fm_pos/dut_pos/dut_tren/
    dut_duoi; cac he so hinh dang thieu thi lay mac dinh do tu anh mau.

    Duong cong gom 5 doan noi lien nhau:
        1. dan hoi    0 -> FeH       : lot may + luy thua mu ~1.80
        2. chay deo   FeH -> FeL     : tut thang + rang cua (MAU_CHAY)
        3. chan chay  FeL -> +0,20mm : nam ngang, thanh cai moc nho
        4. hoa ben    -> Fm          : ham BAO HOA Voce, k ~3.65
        5. that co    Fm -> diem dut : luy thua mu ~3.45
    roi ket bang MOT nhat gay thang dung xuong dut_duoi.
    """
    p_eh, f_eh = t["feh_pos"], t["feh"]
    p_el, f_el = t["fel_pos"], t["fel"]
    p_m, f_m = t["fm_pos"], t["fm"]
    p_b, f_bt, f_bb = t["dut_pos"], t["dut_tren"], t["dut_duoi"]
    mu_dh = t.get("mu_dan_hoi", MU_DAN_HOI_MD)
    k_hb = t.get("k_hoa_ben", K_HOA_BEN_MD)
    mu_tc = t.get("mu_that_co", MU_THAT_CO_MD)
    p_hb = p_el + t.get("chan_chay", CHAN_CHAY_MD)   # bat dau hoa ben

    d = [(0.0, 0.0)]

    # 1. dan hoi. Diem (0, A*feh) nam ngay tren truc tung nen khuc dung
    # thang tu goc len bi truc che, dung nhu anh mau.
    he_so_mu = 1.0 - DH_A - DH_B * (1.0 - math.exp(-1.0 / DH_TAU))
    for i in range(SO_DIEM_DAN_HOI):
        v = i / float(SO_DIEM_DAN_HOI)
        w = DH_A + DH_B * (1.0 - math.exp(-v / DH_TAU)) + he_so_mu * v ** mu_dh
        d.append((p_eh * v, f_eh * w))

    # 2. chay deo: mau rang cua, xe dich nhe theo tung mau
    rnd = random.Random(_hat(t))
    rong, cao = p_el - p_eh, f_eh - f_el
    for k, (s, r) in enumerate(MAU_CHAY):
        if 1 < k < len(MAU_CHAY) - 1:
            s += rnd.uniform(-0.012, 0.012)
            r = max(0.01, r + rnd.uniform(-0.035, 0.035))
        d.append((p_eh + rong * s, f_el + cao * r))

    # 3 + 4. hoa ben (Voce) - diem dau chinh la (p_hb, fel)
    chia = 1.0 - math.exp(-k_hb)
    for i in range(SO_DIEM_HOA_BEN):
        u = i / float(SO_DIEM_HOA_BEN)
        d.append((p_hb + (p_m - p_hb) * u,
                  f_el + (f_m - f_el) * (1.0 - math.exp(-k_hb * u)) / chia))

    # 5. that co
    for i in range(SO_DIEM_THAT_CO + 1):
        u = i / float(SO_DIEM_THAT_CO)
        d.append((p_m + (p_b - p_m) * u, f_m - (f_m - f_bt) * u ** mu_tc))

    # nhat gay: mot doan thang dung duy nhat
    d.append((p_b, f_bb))
    return d


# ------------------------------------------------------------ noi dung trang
class Trang(object):
    """Mot trang bieu do."""

    def __init__(self, t, chart_id, so_mau="1"):
        self.t = t
        self.chart_id = chart_id
        self.so_mau = str(so_mau)


def _y(y_anh):
    """Doi toa do y cua anh mau (huong xuong) sang PDF (huong len)."""
    return CAO_TRANG - y_anh


def _hinh_tron(x, y, r):
    k = r * KAPPA
    return ("%.4f %.4f m %.4f %.4f %.4f %.4f %.4f %.4f c "
            "%.4f %.4f %.4f %.4f %.4f %.4f c %.4f %.4f %.4f %.4f %.4f %.4f c "
            "%.4f %.4f %.4f %.4f %.4f %.4f c f\n"
            % (x, y + r,
               x + k, y + r, x + r, y + k, x + r, y,
               x + r, y - k, x + k, y - r, x, y - r,
               x - k, y - r, x - r, y - k, x - r, y,
               x - r, y + k, x - k, y + r, x, y + r))


def _chu(s, x, y_anh, co=CO_CHU, font="F1", ep=100):
    """Chu tai (x, duong chan chu y_anh) theo toa do anh mau.

    ep: % be ngang chu (toan tu Tz). Khong lam thay doi lop chu Tab 2 doc.
    """
    return ("BT /%s %d Tf %.3f TL %d Tz 1 0 0 1 %.4f %.4f Tm (%s) Tj ET\n"
            % (font, co, co * 1.2, ep, x, _y(y_anh), _thoat(s)))


def _doan(x1, y1, x2, y2):
    return "n %.4f %.4f m %.4f %.4f l S\n" % (x1, _y(y1), x2, _y(y2))


# Vi tri khoi chu so voi tam cham tron, do tu anh mau (pixel, y huong xuong):
#   (ten, thu tu dong, dx chu, dy dong dau, dx dau dan, dy dau dan)
# FeH ghi len tren-phai, dong ten o CUOI; FeL xuong duoi-phai; Fm xuong
# duoi-trai, ngay duoi dinh. Dau dan chi ve toi goc tren-trai cua chu.
KHOI_CHU = (
    ("FeH", ("load", "pos", "ten"), 19.5, -60.5, 19.0, -42.0, MAU_FEH),
    ("FeL", ("ten", "pos", "load"), 45.5, 39.5, 45.0, 28.5, MAU_FEL),
    ("Fm", ("ten", "pos", "load"), -45.5, 28.5, -45.2, 16.8, MAU_FM),
)


def noi_dung_trang(tr):
    """Dung content stream cua mot trang.

    THU TU CAC DOAN CHU LA MOT PHAN CUA GIAO KEO voi Tab 2:
    tensile_core.doc_so_lieu_tren_trang() bat cac chuoi 'Load:' va
    'Position:' theo THU TU XUAT HIEN roi ghep voi nhan FeH/FeL/Fm cung
    theo thu tu. Vi vay khoi FeH phai ra truoc, roi FeL, roi Fm - va moi
    khoi phai co dung mot 'Load:' va mot 'Position:'. Doi thu tu o day la
    hong ca Tab 2.
    """
    t = tr.t
    x_max, y_max = pham_vi_truc(t)
    sx, sy = RONG_VUNG / x_max, CAO_VUNG / y_max

    def px(mm):
        return GOC_X + mm * sx

    def py(kn):                       # toa do anh mau
        return GOC_Y - kn * sy

    o = []
    o.append("1 1 1 rg n 0 0 %.4f %.4f re f\n" % (RONG_TRANG, CAO_TRANG))

    # --- luoi ---
    o.append("%.6f %.6f %.6f RG %.2f w 0 J\n" % (XAM, XAM, XAM, NET_LUOI))
    n = int(round(x_max / BUOC_LUOI_X))
    m = int(round(y_max / BUOC_LUOI_Y))
    for i in range(1, n + 1):
        x = px(i * BUOC_LUOI_X)
        o.append(_doan(x, py(y_max), x, GOC_Y - 1))
    for i in range(1, m + 1):
        y = py(i * BUOC_LUOI_Y)
        o.append(_doan(GOC_X + 1, y, px(x_max), y))

    # --- truc va mui ten tam giac dac ---
    o.append("0 0 0 RG 0 0 0 rg %.2f w\n" % NET_TRUC)
    o.append(_doan(GOC_X - 1, GOC_Y, RONG_ANH - 15, GOC_Y))
    o.append(_doan(GOC_X, GOC_Y + 1, GOC_X, 8))
    ax = RONG_ANH - 14
    o.append("n %.4f %.4f m %.4f %.4f l %.4f %.4f l h f\n"
             % (ax, _y(GOC_Y), ax - 11, _y(GOC_Y - 6), ax - 11, _y(GOC_Y + 6)))
    o.append("n %.4f %.4f m %.4f %.4f l %.4f %.4f l h f\n"
             % (GOC_X, _y(0), GOC_X - 6, _y(11), GOC_X + 6, _y(11)))

    # --- so tren truc ---
    for i in range(n + 1):
        s = "%g" % (i * BUOC_LUOI_X)
        lech = -3.0 if i == 0 else -1.0
        o.append(_chu(s, px(i * BUOC_LUOI_X) + lech
                      - _be_rong(s, CO_CHU_TRUC) * EP_NGANG_CHU / 200.0,
                      639, CO_CHU_TRUC, ep=EP_NGANG_CHU))
    for i in range(m + 1):
        s = "%g" % (i * BUOC_LUOI_Y)
        o.append(_chu(s, 1.5, py(i * BUOC_LUOI_Y) + 2, CO_CHU_TRUC,
                      ep=EP_NGANG_CHU))
    o.append(_chu("Load(kN)", 32, 47))
    o.append(_chu("Position(mm)", 639, 604))

    # --- cham tron + dau dan (ve truoc de duong cong de len tren) ---
    diem = {"FeH": (t["feh_pos"], t["feh"]), "FeL": (t["fel_pos"], t["fel"]),
            "Fm": (t["fm_pos"], t["fm"])}
    for ten, _thu_tu, _dx, _dy, ldx, ldy, mau in KHOI_CHU:
        cx, cy = px(diem[ten][0]), py(diem[ten][1])
        o.append("%.4f %.4f %.4f rg\n" % mau)
        o.append(_hinh_tron(cx, _y(cy), BAN_KINH_DIEM))
        o.append("0 0 0 RG %.2f w\n" % NET_DAN)
        o.append(_doan(cx, cy, cx + ldx, cy + ldy))

    # --- duong cong ---
    d = duong_cong(t)
    o.append("0 0 0 RG %.2f w 1 j 1 J\n" % NET_CONG)
    o.append("n %.4f %.4f m " % (px(d[0][0]), _y(py(d[0][1]))))
    for x, y in d[1:]:
        o.append("%.4f %.4f l " % (px(x), _y(py(y))))
    o.append("S\n")

    # --- khoi chu. Thu tu FeH -> FeL -> Fm la bat buoc (xem tren) ---
    o.append("0 0 0 rg\n")
    for ten, thu_tu, dx, dy, _ldx, _ldy, _mau in KHOI_CHU:
        p, f = diem[ten]
        cx, cy = px(p), py(f)
        noi = {"ten": "%s:" % ten, "pos": "Position:%.3f mm" % p,
               "load": "Load:%.3f kN" % f}
        for i, k in enumerate(thu_tu):
            o.append(_chu(noi[k], cx + dx, cy + dy + i * DONG_CACH,
                          ep=EP_NGANG_CHU))

    # --- so mau canh chan vet gay ---
    o.append(_chu(tr.so_mau, px(t["dut_pos"]) + 3, py(t["dut_duoi"]) + 4.5,
                  CO_CHU_TRUC, ep=EP_NGANG_CHU))
    # --- dong ghi chu ky thuat o chan trang (Tab 2 se cat bo dong nay) ---
    o.append(_chu("Chart ID: %s" % tr.chart_id, GOC_X, CAO_TRANG - 8, 8))
    return "".join(o).encode("latin-1")


# --------------------------------------------------------------- ghi file
def viet_pdf(duong_dan, danh_sach_trang, nen=True, tien_do=None):
    """Ghi mot file PDF nhieu trang. danh_sach_trang: list[Trang]."""
    buf = bytearray()
    vi_tri = {}

    def them(so, noi_dung):
        vi_tri[so] = len(buf)
        buf.extend(b"%d 0 obj\n" % so)
        buf.extend(noi_dung)
        buf.extend(b"\nendobj\n")

    buf.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    n = len(danh_sach_trang)
    so_trang = [5 + 2 * i for i in range(n)]

    them(1, b"<< /Type /Catalog /Pages 2 0 R >>")
    kids = b" ".join(b"%d 0 R" % s for s in so_trang)
    them(2, b"<< /Type /Pages /Count %d /Kids [ %s ] >>" % (n, kids))
    them(3, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
            b"/Encoding /WinAnsiEncoding >>")
    them(4, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
            b"/Encoding /WinAnsiEncoding >>")

    for i, tr in enumerate(danh_sach_trang):
        ps, cs = so_trang[i], so_trang[i] + 1
        them(ps, b"<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 %.4f %.4f ] "
                 b"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                 b"/Contents %d 0 R >>" % (RONG_TRANG, CAO_TRANG, cs))
        # trang tu ve noi dung (ve_keo_so_lieu) thi co ham noi_dung() rieng
        raw = tr.noi_dung() if hasattr(tr, "noi_dung") else noi_dung_trang(tr)
        if nen:
            data = zlib.compress(raw, 9)
            them(cs, b"<< /Length %d /Filter /FlateDecode >>\nstream\n"
                     % len(data) + data + b"\nendstream")
        else:
            them(cs, b"<< /Length %d >>\nstream\n" % len(raw) + raw
                 + b"\nendstream")
        if tien_do and (i % 25 == 0 or i == n - 1):
            tien_do(i + 1, n)

    tong = max(vi_tri) + 1
    bd_xref = len(buf)
    buf.extend(b"xref\n0 %d\n" % tong)
    buf.extend(b"0000000000 65535 f \n")
    for s in range(1, tong):
        buf.extend(b"%010d 00000 n \n" % vi_tri.get(s, 0))
    buf.extend(b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
               % (tong, bd_xref))

    with open(duong_dan, "wb") as f:
        f.write(bytes(buf))
    return duong_dan
