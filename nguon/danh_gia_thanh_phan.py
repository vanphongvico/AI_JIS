# -*- coding: utf-8 -*-
"""
DANH GIA SPC THANH PHAN HOA TU THU MUC DA SINH
==============================================
Doc lai thu muc ThanhPhan_<mac>_<ngay gio>_<N>mau/ (hoac thang file
TongHop_ThanhPhan.xlsx - ca file phong thi nghiem tu dien cung bo cuc), chon
day mau theo (mau bat dau, buoc nhay, so diem), cham 10 nguyen to bang DUNG
cong thuc file SPC xuat ra (thanh_phan_core.cham), roi:

  - xuat file SPC cho tung day bang chinh bo xuat cua tab Sinh
    (xuat_thanh_phan_spc.xuat_spc - da sua 7 loi file mau);
  - "Sua so cho dat": doi so CUA NGUYEN TO CHUA DAT, chi trong cac day chua
    dat. So moi nam trong [Min, Max] cua chuan, dung so chu so; Fe tinh lai,
    Ceq giu trong yeu cau. Ghi de Tong hop, Ket qua in cua mau bi doi (3 lan
    do moi, trung binh = so moi, giu User/Project/gio) va cac file SPC dang
    co trong thu muc (xuat lai cho khop so moi).

"Mau bat dau" la THU TU DONG trong bang Tong hop (1 = dong dau), giong tab
danh gia Va dap; STT in tren bang co the khac neu luc sinh STT bat dau khac 1.
"""

import os
import re
import copy
import random

import openpyxl

try:
    import spc_luat
    import thanh_phan_core as tp
    import xuat_thanh_phan_spc as xs
    import xuat_thanh_phan_word as xw
except ImportError:
    from . import spc_luat
    from . import thanh_phan_core as tp
    from . import xuat_thanh_phan_spc as xs
    from . import xuat_thanh_phan_word as xw

TEN_TONG_HOP = "TongHop_ThanhPhan.xlsx"
SO_DIEM_TOI_DA = 25
_RE_MAC = re.compile(r"^ThanhPhan_(.+?)_\d{4}-\d{2}-\d{2}_")
_RE_FILE_DAY = re.compile(r"^DanhGia_SPC_ThanhPhan_(.+)_bd(\d+)_buoc(\d+)_(\d+)mau\.xlsx$")
_RE_FILE_NHOM = re.compile(r"^DanhGia_SPC_ThanhPhan_(.+)_\d+_mau(\d+)-(\d+)\.xlsx$")

# Tim kiem khi sua so (moi nguyen to): so lan khoi dong lai x so buoc moi lan
SO_LAN_THU_SUA = 30
SO_BUOC_SUA = 2500
# Trong so diem phat: luat/Ceq >> hinh dang bieu do >> so mau bi doi
P_LUAT, P_HINH, P_DOI = 1000, 10, 1


# ---------------------------------------------------------------------------
# 1. DOC
# ---------------------------------------------------------------------------
def _mac_tu_ten(thu_muc):
    m = _RE_MAC.match(os.path.basename(os.path.normpath(thu_muc)))
    return m.group(1) if m else None


def _so_le_dinh_dang(fmt):
    m = re.match(r"^0\.(0+)$", fmt or "")
    return len(m.group(1)) if m else None


def doc(duong_dan):
    """Doc Tong hop. Tra ve {"file", "thu_muc", "mac" (doan theo ten thu muc,
    co the None), "stt", "hang", "cot" {ten: cot}, "mau" [{ten: gia tri}],
    "so_le" {ten: so chu so theo dinh dang o dong dau | None}}."""
    if os.path.isdir(duong_dan):
        thu_muc, f = duong_dan, os.path.join(duong_dan, TEN_TONG_HOP)
        if not os.path.exists(f):
            raise tp.LoiDuLieu(
                "Thư mục không có file %s:\n%s\n\nHãy chọn thư mục ThanhPhan_… do bước ① "
                "sinh ra, hoặc chọn thẳng file Tổng hợp." % (TEN_TONG_HOP, duong_dan))
    elif os.path.isfile(duong_dan) and duong_dan.lower().endswith((".xlsx", ".xlsm")):
        f = duong_dan
        thu_muc = os.path.dirname(os.path.abspath(duong_dan))
    else:
        raise tp.LoiDuLieu("Không tìm thấy thư mục hoặc file .xlsx:\n%s" % duong_dan)

    wb = openpyxl.load_workbook(f, data_only=True)
    try:
        ws = wb.worksheets[0]
        r0 = next((r for r in range(1, ws.max_row + 1)
                   if str(ws.cell(r, 1).value).strip().upper() == "STT"), None)
        if r0 is None:
            raise tp.LoiDuLieu("File %s không có dòng tiêu đề 'STT' ở cột A."
                               % os.path.basename(f))
        cot = {}
        for c in range(2, ws.max_column + 1):
            v = ws.cell(r0, c).value
            if v is not None and str(v).strip() and tp._nfc(v) != "Ceq":
                cot[tp._nfc(v)] = c
        thieu = [t for t in tp.PHAN_TU_SPC if t not in cot]
        if thieu:
            raise tp.LoiDuLieu("File %s thiếu cột nguyên tố: %s."
                               % (os.path.basename(f), ", ".join(thieu)))
        stt, hang, mau, so_le = [], [], [], {}
        for r in range(r0 + 1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if not tp._la_so(v):
                break                                   # het bang
            m = {}
            for ten, c in cot.items():
                o = ws.cell(r, c)
                x = o.value
                if isinstance(x, int) and not isinstance(x, bool):
                    x = float(x)
                m[ten] = x.strip() if isinstance(x, str) else x
                if not mau:
                    so_le[ten] = _so_le_dinh_dang(o.number_format)
            stt.append(int(v))
            hang.append(r)
            mau.append(m)
    finally:
        wb.close()
    if not mau:
        raise tp.LoiDuLieu("File %s không có dòng số liệu nào dưới dòng 'STT'."
                           % os.path.basename(f))
    return {"file": f, "thu_muc": thu_muc, "mac": _mac_tu_ten(thu_muc),
            "stt": stt, "hang": hang, "cot": cot, "mau": mau, "so_le": so_le}


def ds_theo_file(ds, d):
    """Ban sao danh sach nguyen to, so chu so lay theo dinh dang o cua file
    (= so chu so luc sinh, ke ca khi Min/Max da sua). Dinh dang General -> giu."""
    ra = copy.deepcopy(ds)
    for n in ra:
        k = d["so_le"].get(n["ten"])
        if n["min"] is not None and k is not None:
            n["so_le"] = k
    return ra


# ---------------------------------------------------------------------------
# 2. CHON DAY, CHAM, QUET
# ---------------------------------------------------------------------------
def mac_dinh_bat_dau(so_mau, so_diem=SO_DIEM_TOI_DA):
    """Chia nhom nhu luc sinh: 1, 26, 51... (nhom chi 1 mau bo - I-MR can >= 2)."""
    return [s for s in range(1, so_mau + 1, so_diem) if so_mau - s + 1 >= 2]


def lay_day(d, bat_dau, buoc=1, so_diem=SO_DIEM_TOI_DA):
    """Thu tu dong (1-based) cua day. Het du lieu thi day ngan lai."""
    try:
        bat_dau = spc_luat.so_nguyen(bat_dau, "Mẫu bắt đầu", 1)
        buoc = spc_luat.so_nguyen(buoc, "Bước nhảy", 1)
        so_diem = spc_luat.so_nguyen(so_diem, "Số điểm", 2)
    except ValueError as e:
        raise tp.LoiDuLieu(str(e))
    if so_diem > SO_DIEM_TOI_DA:
        raise tp.LoiDuLieu("Số điểm tối đa %d — biểu mẫu SPC chỉ chứa %d mẫu. Đang khai %d."
                           % (SO_DIEM_TOI_DA, SO_DIEM_TOI_DA, so_diem))
    n = len(d["mau"])
    vt = [v for v in range(bat_dau, bat_dau + buoc * so_diem, buoc) if v <= n]
    if len(vt) < 2:
        raise tp.LoiDuLieu("Bắt đầu từ mẫu %d (bước %d) chỉ lấy được %d mẫu — I-MR cần ít "
                           "nhất 2 điểm. Dữ liệu có %d mẫu." % (bat_dau, buoc, len(vt), n))
    return vt


def _cot(d, t, vi_tri):
    return [d["mau"][v - 1].get(t) for v in vi_tri]


def _cham_cot(X, nt, cpk):
    if not all(tp._la_so(x) for x in X):
        return spc_luat.KetQuaSPC(
            n=len(X), X=X, MR=None, xb=None, mrb=None, sigma=None, ucl_i=None,
            lcl_i=None, ucl_mr=None, lcl_mr=None, lsl=nt["lsl"], usl=nt["usl"],
            cpk=None, cp=None, ca=None, loi=["DU_LIEU"])
    return tp.cham(X, nt["lsl"], nt["usl"], nt["so_le"], cpk)


def cham_day(d, ds, vi_tri, cpk=tp.CPK_TOI_THIEU):
    """{ten: KetQuaSPC} cho 10 nguyen to cua mot day."""
    nt = {n["ten"]: n for n in ds}
    return {t: _cham_cot(_cot(d, t, vi_tri), nt[t], cpk) for t in tp.PHAN_TU_SPC}


def quet(d, ds, buoc=1, so_diem=SO_DIEM_TOI_DA, cpk=tp.CPK_TOI_THIEU):
    """Moi diem bat dau co du so_diem mau: [(bat_dau, dat_ca_10, [nguyen to loi])]."""
    het = len(d["mau"]) - (so_diem - 1) * buoc
    if het < 1:
        raise tp.LoiDuLieu("Với bước nhảy %d và %d điểm cần ít nhất %d mẫu, dữ liệu chỉ có %d."
                           % (buoc, so_diem, (so_diem - 1) * buoc + 1, len(d["mau"])))
    ra = []
    for s in range(1, het + 1):
        r = cham_day(d, ds, list(range(s, s + buoc * so_diem, buoc)), cpk)
        hong = [t for t in tp.PHAN_TU_SPC if not r[t].dat]
        ra.append((s, not hong, hong))
    return ra


# ---------------------------------------------------------------------------
# 3. XUAT FILE SPC
# ---------------------------------------------------------------------------
def ten_file_day(mac, bat_dau, buoc, so_diem):
    return "DanhGia_SPC_ThanhPhan_%s_bd%d_buoc%d_%dmau.xlsx" % (mac, bat_dau, buoc, so_diem)


def xuat_spc_day(d, ds, vi_tri, mac, file_mau_spc, cpk, ngay, thong_tin, duong_dan):
    """Xuat file SPC cho mot day (dat hay khong deu xuat, tick dung ket qua).
    Tra ve (duong_dan, {ten: KetQuaSPC})."""
    spc = cham_day(d, ds, vi_tri, cpk)
    thieu = [t for t, r in spc.items() if "DU_LIEU" in r.loi]
    if thieu:
        raise tp.LoiDuLieu("Dãy bắt đầu mẫu %d: %s có ô trống hoặc không phải số — chưa xuất "
                           "được. Bấm 'Sửa số cho đạt' hoặc sửa file Tổng hợp."
                           % (vi_tri[0], ", ".join(thieu)))
    xs.xuat_spc(file_mau_spc, duong_dan, [d["mau"][v - 1] for v in vi_tri], ds, spc, mac,
                cpk, ngay, thong_tin)
    return duong_dan, spc


def file_spc_trong_thu_muc(d, mac):
    """File SPC dang co trong thu muc (cua tab Sinh: _01_mau1-25; cua tab nay:
    _bd1_buoc1_25mau) kem day mau cua chung: [(duong_dan, vi_tri)]."""
    tm, n, ra = d["thu_muc"], len(d["mau"]), []
    if not tm or not os.path.isdir(tm):
        return ra
    for f in sorted(os.listdir(tm)):
        m = _RE_FILE_DAY.match(f)
        if m and m.group(1) == mac:
            bd, buoc, so = (int(x) for x in m.groups()[1:])
            vt = list(range(bd, bd + buoc * so, buoc))
            if vt and vt[-1] <= n:
                ra.append((os.path.join(tm, f), vt))
            continue
        m = _RE_FILE_NHOM.match(f)
        if m and m.group(1) == mac:
            a, b = int(m.group(2)), int(m.group(3))
            vt = [i + 1 for i, s in enumerate(d["stt"]) if a <= s <= b]
            if len(vt) == b - a + 1 and len(vt) >= 2:
                ra.append((os.path.join(tm, f), vt))
    return ra


# ---------------------------------------------------------------------------
# 4. SUA SO CHO DAT
# ---------------------------------------------------------------------------
def _phat_hinh(X, so_le):
    """Hinh dang bieu do bieu mau khong cham nhung nhin nhu khong dat (giong luc sinh)."""
    nua = 0.5 * 10 ** -so_le
    return (sum(1 for a, b in zip(X, X[1:]) if abs(b - a) < nua) +
            ("MR9" in tp.loi_hinh_dang(X, so_le)) + tp._so_sat(X, so_le))


def _ngoai_ceq(m, ceq):
    return bool(ceq) and tp._ngoai_ceq(m, ceq)


def _sua_cot(d, nt, P, W, ceq, cpk, rng, ceq_dong=()):
    """Tim gia tri moi cho nguyen to nt tai cac dong P sao cho moi day trong W
    dat. Ceq xet o dong bi doi va o ceq_dong. Khong ghi vao d.
    Tra ve (gia tri {v: x}, so luat con loi, so mau Ceq vuot)."""
    t, mn, mx, so = nt["ten"], nt["min"], nt["max"], nt["so_le"]
    lsl, usl = nt["lsl"], nt["usl"]
    tam, bk = (mn + mx) / 2.0, (mx - mn) / 2.0
    sg = max(bk / 3.0, 0.6 * 10 ** -so)

    def moi():
        return round(max(mn, min(mx, rng.gauss(tam, sg))), so)

    def hong(x):
        return (not tp._la_so(x) or (lsl is not None and x < lsl - 1e-12) or
                (usl is not None and x > usl + 1e-12))

    goc = {v: d["mau"][v - 1].get(t) for v in P}
    he_so = dict(tp.HE_SO_CEQ).get(t) if ceq else None
    nen = {}                                    # Ceq cua mau khi bo nguyen to t
    if he_so:
        for v in P:
            m = dict(d["mau"][v - 1])
            m.pop(t, None)
            nen[v] = tp.tinh_ceq(m)

    def ngoai_ceq(v, x):
        c = nen[v] + x / he_so
        return ((ceq["usl"] is not None and c > ceq["usl"] + 1e-12) or
                (ceq["lsl"] is not None and c < ceq["lsl"] - 1e-12))

    def phat(c):
        luat = hinh = vuot = 0
        for w in W:
            X = [c[v] if v in c else d["mau"][v - 1][t] for v in w]
            luat += len(tp.cham(X, lsl, usl, so, cpk).loi)
            hinh += _phat_hinh(X, so)
        if he_so:
            vuot = sum(1 for v in P if (v in ceq_dong or c[v] != goc[v]) and ngoai_ceq(v, c[v]))
        doi = sum(1 for v in P if c[v] != goc[v])
        return (luat + vuot) * P_LUAT + hinh * P_HINH + doi * P_DOI, (luat, vuot)

    tot = None
    for lan in range(SO_LAN_THU_SUA):
        # Lan dau giu so cu (chi thay o hong); lan sau sinh moi ca phan duoc doi
        c = {v: (moi() if hong(goc[v]) else goc[v]) for v in P} if lan == 0 else \
            {v: moi() for v in P}
        p, luat = phat(c)
        for _ in range(SO_BUOC_SUA):
            if p == 0:
                break
            k = rng.random()
            if k < 0.4 and len(P) >= 2:
                i, j = rng.sample(P, 2)
                c2 = dict(c)
                c2[i], c2[j] = c[j], c[i]
            elif k < 0.85:
                c2 = dict(c)
                c2[rng.choice(P)] = moi()
            else:                                # tra lai so cu neu duoc
                lui = [v for v in P if c[v] != goc[v] and not hong(goc[v])]
                if not lui:
                    continue
                c2 = dict(c)
                v = rng.choice(lui)
                c2[v] = goc[v]
            p2, luat2 = phat(c2)
            if p2 <= p:
                c, p, luat = c2, p2, luat2
        if tot is None or p < tot[0]:
            tot = (p, c, luat)
        # Ceq vuot co the do nguyen to KHAC dang hong (sua sau) - chi thu lai mot lan
        if tot[2][0] == 0 and (tot[2][1] == 0 or lan >= 1):
            break
    return tot[1], tot[2][0], tot[2][1]


def sua(d, ds, ceq, cac_day, cpk=tp.CPK_TOI_THIEU, seed=None):
    """Sua TRUC TIEP d["mau"] cho moi day trong cac_day dat ca 10 nguyen to.

    Chi doi nguyen to chua dat, chi o cac dong thuoc day chua dat cua nguyen to
    do; day chon chong len nhau thi moi day chung dong deu phai dat. Mau bi doi
    phai co Ceq trong yeu cau; Fe tinh lai.
    Tra ve {"doi": {(dong, ten): (cu, moi)}, "con_loi": [chu]}.
    """
    rng = random.Random(seed)
    nt = {n["ten"]: n for n in ds}
    cac_day = [list(w) for w in cac_day]
    truoc = copy.deepcopy(d["mau"])
    con_loi = []

    def ap(t, c):
        for v, x in c.items():
            d["mau"][v - 1][t] = x

    for t in tp.PHAN_TU_SPC:
        hong = [w for w in cac_day if not _cham_cot(_cot(d, t, w), nt[t], cpk).dat]
        if not hong:
            continue
        P = sorted({v for w in hong for v in w})
        W = [w for w in cac_day if set(w) & set(P)]
        c, luat, _vuot = _sua_cot(d, nt[t], P, W, ceq, cpk, rng)
        ap(t, c)
        if luat:
            for w in W:
                r = _cham_cot(_cot(d, t, w), nt[t], cpk)
                if not r.dat:
                    con_loi.append("%s – dãy bắt đầu mẫu %d: %s" % (t, w[0], ", ".join(r.loi)))

    # Ceq: xet sau cung (Ceq ghep tu 7 nguyen to). Mau bi doi ma con vuot thi ha
    # lan luot tung nguyen to cua Ceq, khong duoc lam hong luat.
    def doi_dong():
        return sorted({k + 1 for k in range(len(truoc)) for t in tp.PHAN_TU_SPC
                       if truoc[k].get(t) != d["mau"][k].get(t)})
    vuot = [v for v in doi_dong() if _ngoai_ceq(d["mau"][v - 1], ceq)]
    for t, _h in tp.HE_SO_CEQ:
        if not vuot:
            break
        P = sorted({x for w in cac_day if set(w) & set(vuot) for x in w})
        W = [w for w in cac_day if set(w) & set(P)]
        c, luat, _v = _sua_cot(d, nt[t], P, W, ceq, cpk, rng, ceq_dong=set(vuot))
        if not luat:
            ap(t, c)
            vuot = [v for v in doi_dong() if _ngoai_ceq(d["mau"][v - 1], ceq)]
    if vuot:
        con_loi.append("Ceq vượt yêu cầu %s ở mẫu STT %s" % (
            ceq["yeu_cau"], ", ".join(str(d["stt"][v - 1]) for v in vuot[:10])))

    doi = {}
    for v in doi_dong():
        for t in tp.PHAN_TU_SPC:
            if truoc[v - 1].get(t) != d["mau"][v - 1].get(t):
                doi[(v, t)] = (truoc[v - 1].get(t), d["mau"][v - 1][t])
        if "Fe" in d["mau"][v - 1]:
            d["mau"][v - 1]["Fe"] = tp.tinh_fe(d["mau"][v - 1])
    return {"doi": doi, "con_loi": con_loi}


# ---------------------------------------------------------------------------
# 5. GHI LAI THU MUC
# ---------------------------------------------------------------------------
def _ghi_tong_hop(d, doi):
    wb = openpyxl.load_workbook(d["file"])
    ws = wb.worksheets[0]
    for v, t in doi:
        r = d["hang"][v - 1]
        ws.cell(r, d["cot"][t]).value = d["mau"][v - 1][t]
        if "Fe" in d["cot"]:
            ws.cell(r, d["cot"]["Fe"]).value = d["mau"][v - 1]["Fe"]
    wb.save(d["file"])


def ghi_lai(d, ds, doi, file_mau_spc, mac, cpk, ngay, thong_tin=None, seed=None):
    """Ghi de so da sua: Tong hop, Ket qua in cua mau bi doi, file SPC dang co
    trong thu muc ma co mau bi doi. thong_tin: dict, hoac ham(vi_tri) -> dict
    (thong tin rieng tung day). Tra ve {"tong_hop", "word", "spc", "canh_bao"}."""
    ra = {"tong_hop": d["file"], "word": [], "spc": [], "canh_bao": []}
    if not doi:
        return ra
    _ghi_tong_hop(d, doi)

    nt = {n["ten"]: n for n in ds}
    rng = random.Random(None if seed is None else "%s-in" % seed)
    thieu = []
    for v in sorted({v for v, _t in doi}):
        stt = d["stt"][v - 1]
        p = os.path.join(d["thu_muc"], "KetQuaIn", "KetQuaIn_%d.docx" % stt)
        if not os.path.exists(p):
            thieu.append(stt)
            continue
        w = xw.MauWord(p)
        sl = w.doc_so_lieu()
        m = d["mau"][v - 1]
        for vv, t in doi:
            if vv == v and t in sl:
                lan, rsd = tp.ba_lan_do(m[t], nt[t]["so_le"], rng)
                sl[t] = {"lan": lan, "tb": m[t], "rsd": rsd, "so_le": nt[t]["so_le"]}
        if "Fe" in sl and tp._la_so(m.get("Fe")):
            sl["Fe"] = tp.fe_ba_lan(sl, m["Fe"])
        w.ghi_so_lieu(p, sl)
        ra["word"].append(p)
    if thieu:
        ra["canh_bao"].append("Không thấy file Kết quả in của mẫu STT %s — chỉ sửa Tổng hợp."
                              % ", ".join(str(s) for s in thieu[:10]))

    vt = {v for v, _t in doi}
    for p, w in file_spc_trong_thu_muc(d, mac):
        if set(w) & vt:
            xuat_spc_day(d, ds, w, mac, file_mau_spc, cpk, ngay,
                         thong_tin(w) if callable(thong_tin) else thong_tin, p)
            ra["spc"].append(p)
    return ra


def tom_tat_doi(d, doi, ds):
    """Dong chu cho nhat ky: 'C: mau 5 0.250 -> 0.152; ...' theo nguyen to."""
    nt = {n["ten"]: n for n in ds}
    L = []
    for t in tp.PHAN_TU_SPC:
        o = sorted((v, cu, moi) for (v, tt), (cu, moi) in doi.items() if tt == t)
        if not o:
            continue
        so = nt[t]["so_le"]

        def f(x):
            return ("%.*f" % (so, x)) if tp._la_so(x) else ("(trống)" if x is None else str(x))
        L.append("  %s — đổi %d mẫu: %s%s" % (
            t, len(o), "; ".join("STT %d: %s → %s" % (d["stt"][v - 1], f(cu), f(moi))
                                 for v, cu, moi in o[:6]), " …" if len(o) > 6 else ""))
    return "\n".join(L)

