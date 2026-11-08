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
  - "Sua cho dat": SAP LAI THU TU cac mau (ca dong mau di cung nhau) cho toi
    khi cac day dang chon dat - KHONG doi so nao, nen Fe, Ceq va Ket qua in
    cua tung mau giu nguyen. Ghi de Tong hop (ban goc sao luu vao
    _truoc_khi_sua/), file SPC cu chuyen vao ban sao luu vi phu thuoc thu tu.

"Mau bat dau" la THU TU DONG trong bang Tong hop (1 = dong dau), giong tab
danh gia Va dap; STT in tren bang co the khac neu luc sinh STT bat dau khac 1.
"""

import os
import re
import copy
import random
import shutil
import time
from datetime import datetime

import openpyxl

try:
    import spc_luat
    import thanh_phan_core as tp
    import xuat_thanh_phan_spc as xs
except ImportError:
    from . import spc_luat
    from . import thanh_phan_core as tp
    from . import xuat_thanh_phan_spc as xs

TEN_TONG_HOP = "TongHop_ThanhPhan.xlsx"
SO_DIEM_TOI_DA = 25
_RE_MAC = re.compile(r"^ThanhPhan_(.+?)_\d{4}-\d{2}-\d{2}_")
_RE_FILE_DAY = re.compile(r"^DanhGia_SPC_ThanhPhan_(.+)_bd(\d+)_buoc(\d+)_(\d+)mau\.xlsx$")
_RE_FILE_NHOM = re.compile(r"^DanhGia_SPC_ThanhPhan_(.+)_\d+_mau(\d+)-(\d+)\.xlsx$")



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
# 4. SUA CHO DAT: SAP LAI THU TU MAU (KHONG DOI SO NAO)
# ---------------------------------------------------------------------------
# Truoc day "Sua so cho dat" THAY SO bang so ngau nhien trong Min/Max va viet
# lai Ket qua in kem 3 lan do tu tao - lam gia so lieu neu file la cua phong
# thi nghiem. Gio chi doi cho CA DONG mau (STT + moi nguyen to di cung nhau):
# thanh phan, Fe, Ceq va Ket qua in cua tung mau giu nguyen.
GIAY_SUA_TOI_DA = 120


def nhan_day(d, vt, ngan=False):
    """'Mau STT a-b' khi STT cua day lien tiep deu; sau khi sap lai thu tu
    STT co the khong con lien tiep -> ghi theo dong cua Tong hop."""
    st = [d["stt"][v - 1] for v in vt]
    buoc = vt[1] - vt[0] if len(vt) > 1 else 1
    if all(b - a == buoc for a, b in zip(st, st[1:])):
        return ("%d–%d" if ngan else "Mẫu STT %d–%d") % (st[0], st[-1])
    return ("dòng %d–%d" if ngan else "Dòng %d–%d của Tổng hợp") % (vt[0], vt[-1])


def _doi_cho(d, goc, i, j):
    for k in ("mau", "stt"):
        d[k][i], d[k][j] = d[k][j], d[k][i]
    goc[i], goc[j] = goc[j], goc[i]


def sua_hoan_vi(d, ds, cac_day, cpk=tp.CPK_TOI_THIEU, giay_toi_da=GIAY_SUA_TOI_DA,
                seed=None, dung=None, tien_do=None):
    """Sap lai THU TU mau cho toi khi moi day trong cac_day dat ca 10 nguyen to.

    Cac day mac dinh 25 mau (1, 26, 51...) dang dat thi giu dat. Sua truc
    tiep d["mau"], d["stt"]; d["thu_tu_goc"][k] = chi so dong goc (0-based)
    cua mau dang dung o vi tri k. Hoan vi khong doi gia tri nao, nen khong
    sua duoc mau co so ngoai dung sai hoac o trong.
    Tra ve {"con_hong": [nhan day], "vong", "giay", "doi", "so_khoi"}.
    """
    rnd = random.Random(seed)
    nt = {n["ten"]: n for n in ds}
    n = len(d["mau"])
    goc = d.setdefault("thu_tu_goc", list(range(n)))

    def cost(w):
        return sum(len(_cham_cot([d["mau"][v - 1].get(t) for v in w], nt[t], cpk).loi)
                   for t in tp.PHAN_TU_SPC)

    khoi, thay = [], set()
    for w in cac_day:
        w = tuple(w)
        if w not in thay:
            thay.add(w)
            khoi.append(w)
    for s in mac_dinh_bat_dau(n):
        w = tuple(lay_day(d, s, 1, SO_DIEM_TOI_DA))
        if w not in thay and cost(w) == 0:
            thay.add(w)
            khoi.append(w)
    thuoc = {}
    for b, w in enumerate(khoi):
        for v in w:
            thuoc.setdefault(v - 1, []).append(b)

    chi_phi = [cost(w) for w in khoi]
    t0 = time.time()
    vong = 0
    while True:
        hong = [b for b, x in enumerate(chi_phi) if x]
        if not hong:
            break
        giay = time.time() - t0
        if giay > giay_toi_da or (dung and dung()):
            break
        vong += 1
        if tien_do and vong % 100 == 0:
            tien_do(len(hong), giay)
        w = khoi[rnd.choice(hong)]
        i = rnd.choice(w) - 1
        j = (rnd.choice(w) - 1) if rnd.random() < 0.6 else rnd.randrange(n)
        if i == j:
            continue
        cham = set(thuoc.get(i, ())) | set(thuoc.get(j, ()))
        truoc = sum(chi_phi[b] for b in cham)
        _doi_cho(d, goc, i, j)
        moi = {b: cost(khoi[b]) for b in cham}
        if sum(moi.values()) <= truoc:
            for b, q in moi.items():
                chi_phi[b] = q
        else:
            _doi_cho(d, goc, i, j)
    return {"con_hong": [nhan_day(d, list(khoi[b])) for b, x in enumerate(chi_phi) if x],
            "vong": vong, "giay": time.time() - t0, "so_khoi": len(khoi),
            "doi": sum(1 for k, g in enumerate(goc) if k != g)}


# ---------------------------------------------------------------------------
# 5. GHI THU TU MOI
# ---------------------------------------------------------------------------
def ghi_hoan_vi(d):
    """Ghi thu tu moi vao Tong hop (ca dong di cung nhau; o cong thuc theo dong
    nhu Ceq de nguyen). Sao luu Tong hop vao _truoc_khi_sua/<thoi diem>/ va
    CHUYEN cac file SPC cu vao do - chung phu thuoc thu tu nen het hieu luc.
    Ket qua in giu nguyen: moi file la cua mot STT, so khong doi.
    Tra ve {"tong_hop", "sao_luu", "spc_cu": [ten file]}."""
    goc = d.get("thu_tu_goc") or []
    ra = {"tong_hop": d["file"], "sao_luu": None, "spc_cu": []}
    if goc == list(range(len(goc))):
        return ra
    tm = d["thu_muc"]
    bk = os.path.join(tm, "_truoc_khi_sua", datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f"))
    os.makedirs(bk)
    shutil.copy2(d["file"], os.path.join(bk, os.path.basename(d["file"])))

    def cong_thuc(v):
        return isinstance(v, str) and v.startswith("=")

    wb = openpyxl.load_workbook(d["file"])
    ws = wb.worksheets[0]
    so_cot = ws.max_column
    cu = [[(ws.cell(r, c).value, copy.copy(ws.cell(r, c)._style)) for c in range(1, so_cot + 1)]
          for r in d["hang"]]
    for k, g in enumerate(goc):
        for c in range(1, so_cot + 1):
            o = ws.cell(d["hang"][k], c)
            v, kieu = cu[g][c - 1]
            if cong_thuc(o.value) or cong_thuc(v):
                continue
            o.value = v
            o._style = kieu
    wb.save(d["file"])

    for f in sorted(os.listdir(tm)):
        if f.startswith("DanhGia_SPC_ThanhPhan_") and f.lower().endswith(".xlsx"):
            shutil.move(os.path.join(tm, f), os.path.join(bk, f))
            ra["spc_cu"].append(f)
    with open(os.path.join(bk, "GHI_CHU.txt"), "w", encoding="utf-8") as f:
        f.write("Ban goc truoc khi sap lai thu tu mau (%s).\n"
                "Chi doi cho ca dong mau, khong doi so nao. File SPC cu phu thuoc thu "
                "tu nen da chuyen vao day - can xuat lai.\n"
                "Vi tri moi -> STT: %s\n" % (os.path.basename(d["file"]),
                                               ", ".join(str(s) for s in d["stt"])))
    d["thu_tu_goc"] = list(range(len(goc)))
    ra["sao_luu"] = bk
    return ra
