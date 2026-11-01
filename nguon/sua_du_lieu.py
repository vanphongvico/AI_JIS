# -*- coding: utf-8 -*-
"""
SUA DU LIEU TRONG THU MUC CHO DAT SPC
======================================
Khi mot day 25 mau lay ra khong thoa cac muc danh gia, module nay sap lai
THU TU cac mau trong thu muc cho toi khi thoa, roi ghi lai cac file du lieu.

Cach lam: HOAN VI, khong bia so moi
-----------------------------------
Hoan vi giu nguyen cac gia tri va ma mau; trung binh toan lo khong doi.
MR-bar, sigma, Cpk, gioi han kiem soat va ket qua luat co the thay doi.
Khong the sua loi nghiem thu tung so do bang cach doi thu tu.

Mot ban ghi di lien mot khoi (ca so lieu lan tham so bieu do), nen sau khi
hoan vi thi bieu do van khop bang so lieu. Ma mau goc duoc giu lai de truy vet; vi tri moi ghi trong metadata.

An toan
-------
Truoc khi ghi de, toan bo file bi thay doi duoc chep sang thu muc con
'_truoc_khi_sua/'. Neu ket qua khong vua y thi chep nguoc lai la xong.
"""

import os
import json
import glob
import shutil
import random
import tempfile
import copy
import hashlib
from datetime import datetime

import openpyxl

import spc_luat
import doc_thu_muc as dtm

COT_KEO = ["ReL", "ReH", "Rm", "A", "Z"]
COT_CHARPY = ["Nang luong TB", "% mat gay TB"]

# Vi tri cot trong hai file du lieu (1-based)
_CLT = {"ReL": 8, "ReH": 10, "Rm": 12, "A": 13, "Z": 14}    # Co_ly_tinh
_BM31_E, _BM31_P = 7, 11                                    # TongHop_VaDap


class LoiSua(Exception):
    pass


# ------------------------------------------------------------ 1. DOC DAY DU
def doc_day_du(thu_muc):
    """Doc toan bo ban ghi cua tung mau, du de ghi lai file sau khi hoan vi."""
    if not os.path.isdir(thu_muc):
        raise LoiSua("Chi sua duoc khi chon ca THU MUC ket qua, khong sua "
                     "duoc khi chi chon mot file le.")
    f = os.path.join(thu_muc, "TongHop_VaDap.xlsx")
    if os.path.exists(f):
        d = _doc_charpy(thu_muc, f)
        d["_tep_nguon"] = _tep_nguon(thu_muc)
        return d
    g = sorted(glob.glob(os.path.join(thu_muc, "SoLieu_*.xlsx")))
    if g:
        if len(g) > 1:
            raise LoiSua("Thu muc co nhieu file SoLieu: can xac dinh mot bo du lieu duy nhat truoc khi sua.")
        d = _doc_keo(thu_muc, g[0])
        d["_tep_nguon"] = _tep_nguon(thu_muc)
        return d

    # Neu thu muc chi chua file Word BaoCao/Tensile_*.docx (thu muc Tab 2 cu)
    # Tu dong tai tao file SoLieu de nguoi dung co the sua duoc ngay
    # Doc/kiem tra khong duoc tu y tao hoac ghi file nguon.

    raise LoiSua(
        "Thu muc nay khong co file du lieu de sua:\n%s\n\n"
        "Chi sua duoc thu muc co TongHop_VaDap.xlsx (thu va dap) hoac "
        "SoLieu_<so>.xlsx (sinh bieu do keo).\n\n"
        "Thu muc cua Tab 2 chi chua bao cao Word nen khong sua duoc - hay "
        "sua o thu muc Keo_Sinh_... goc roi chay lai Tab 2." % thu_muc)


def _khoi_phuc_excel_tu_word(thu_muc):
    """Tu dong tai tao SoLieu_<N>.xlsx tu cac file Word trong thu muc BaoCao."""
    files = sorted(glob.glob(os.path.join(thu_muc, "BaoCao", "Tensile_*.docx")))
    if not files:
        return None
    wb = openpyxl.Workbook()
    ws_th = wb.active
    ws_th.title = "Tong_hop"
    ws_th.append(["Mau", "FeH_Position_mm", "FeH_Load_kN", "FeL_Position_mm",
                  "FeL_Load_kN", "Fm_Position_mm", "Fm_Load_kN",
                  "Break_Position_mm", "Break_Top_Load_kN",
                  "Break_Bottom_Load_kN", "d_mm", "du_mm"])

    ws_cl = wb.create_sheet("Co_ly_tinh")
    ws_cl.append(["NO", "d (mm)", "du (mm)", "Lo (mm)", "Lu (mm)", "Le (mm)",
                  "FeL (N)", "ReL (MPa)", "FeH (N)", "ReH (MPa)",
                  "Fm (N)", "Rm (MPa)", "A (%)", "Z (%)", "Ket_luan"])
    ws_cl.append([None, "mm", "mm", "mm", "mm", "mm", "N", "MPa", "N", "MPa",
                  "N", "MPa", "%", "%", None])

    for f in files:
        try:
            ma, o = dtm._doc_bang_word(f)
        except Exception:
            continue
        ma_mau = ma or os.path.splitext(os.path.basename(f))[0]
        d_val = dtm._so(o[0]) if len(o) > 0 else None
        du_val = dtm._so(o[1]) if len(o) > 1 else None
        Lo_val = dtm._so(o[2]) if len(o) > 2 else None
        Lu_val = dtm._so(o[3]) if len(o) > 3 else None
        Le_val = dtm._so(o[4]) if len(o) > 4 else None
        fel_val = dtm._so(o[5]) if len(o) > 5 else None
        rel_val = dtm._so(o[6]) if len(o) > 6 else None
        feh_val = dtm._so(o[7]) if len(o) > 7 else None
        reh_val = dtm._so(o[8]) if len(o) > 8 else None
        fm_val = dtm._so(o[9]) if len(o) > 9 else None
        rm_val = dtm._so(o[10]) if len(o) > 10 else None
        a_val = dtm._so(o[11]) if len(o) > 11 else None
        z_val = dtm._so(o[12]) if len(o) > 12 else None

        ws_th.append([ma_mau, None, feh_val, None, fel_val, None, fm_val,
                      None, None, None, d_val, du_val])
        # Word in luc theo kN; Co_ly_tinh ghi N (Tong_hop giu kN)
        def n(v):
            return None if v is None else int(round(v * 1000))
        ws_cl.append([ma_mau, d_val, du_val, Lo_val, Lu_val, Le_val,
                      n(fel_val), rel_val, n(feh_val), reh_val, n(fm_val), rm_val,
                      a_val, z_val, "OK"])

    f_out = os.path.join(thu_muc, "SoLieu_%d.xlsx" % len(files))
    wb.save(f_out)
    return f_out


def _doc_charpy(thu_muc, f):
    ws = openpyxl.load_workbook(f).worksheets[0]
    bg = []
    r = 12
    while r <= ws.max_row:
        v = ws.cell(r, 1).value
        if v is None or dtm._so(v) is None:
            break
        bg.append({"dong": r,
                   "o": [ws.cell(r, c).value for c in range(1, 12)]})
        r += 1
    if not bg:
        raise LoiSua("Khong doc duoc dong du lieu nao trong:\n%s" % f)
    for b in bg:
        e = dtm._so(b["o"][_BM31_E - 1])
        p = dtm._so(b["o"][_BM31_P - 1])
        b["cot"] = {"Nang luong TB": None if e is None else round(e, 1),
                    "% mat gay TB": None if p is None else round(p, 1)}
    return {"loai": "charpy", "thu_muc": thu_muc, "file_xlsx": f,
            "ban_ghi": bg, "ten_cot": COT_CHARPY,
            "nguong": {"Nang luong TB": dtm._so(ws["G11"].value),
                       "% mat gay TB": dtm._so(ws["K11"].value)},
            "meta": _doc_meta(thu_muc)}


def _doc_keo(thu_muc, f):
    wb = openpyxl.load_workbook(f)
    if "Tong_hop" not in wb.sheetnames or "Co_ly_tinh" not in wb.sheetnames:
        raise LoiSua("File so lieu thieu sheet Tong_hop hoac Co_ly_tinh:\n%s"
                     % f)
    th, cl = wb["Tong_hop"], wb["Co_ly_tinh"]
    dau_th = [c.value for c in th[1]]
    bg = []
    r = 2
    while r <= th.max_row:
        if th.cell(r, 1).value in (None, ""):
            break
        bg.append({"th": [th.cell(r, c).value
                          for c in range(1, len(dau_th) + 1)]})
        r += 1
    r = 3
    i = 0
    while r <= cl.max_row and i < len(bg):
        if cl.cell(r, 1).value in (None, ""):
            break
        bg[i]["cl"] = [cl.cell(r, c).value for c in range(1, 16)]
        r += 1
        i += 1
    if i != len(bg):
        raise LoiSua("Sheet Tong_hop va Co_ly_tinh lech so dong (%d / %d)."
                     % (len(bg), i))
    for b in bg:
        b["cot"] = dict((k, dtm._so(b["cl"][c - 1])) for k, c in _CLT.items())
    return {"loai": "keo", "thu_muc": thu_muc, "file_xlsx": f,
            "dau_th": dau_th, "ban_ghi": bg, "ten_cot": COT_KEO,
            "nguong": dict((k, n) for k, _u, _nh, n in dtm.COT_KEO),
            "meta": _doc_meta(thu_muc)}


def _doc_meta(thu_muc):
    p = os.path.join(thu_muc, "_thongtin_sinh.json")
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


# ------------------------------------------------------- 2. XAY RANG BUOC
def xay_rang_buoc(d, lua_chon, so_diem=25, giu_bao_dam=True):
    """Liet ke cac khoi phai dat.

    lua_chon: list (ten_cot, mau_bat_dau, buoc_nhay) - dem tu 1.
    giu_bao_dam: co gang giu lai pham vi bao dam ghi trong _thongtin_sinh.json
                 de viec sua khong pha vo thu von da dat truoc do.
    """
    n = len(d["ban_ghi"])
    try:
        so_diem = spc_luat.so_nguyen(so_diem, "So diem", 2)
    except ValueError as e:
        raise LoiSua(str(e))
    ten = d["ten_cot"]
    thay = set()
    ra = []

    def them(idx, c):
        k = (idx, c)
        if idx[0] >= 0 and idx[-1] < n and k not in thay:
            thay.add(k)
            ra.append(k)

    for tc, bd, b in lua_chon:
        try:
            bd = spc_luat.so_nguyen(bd, "Mau bat dau")
            b = spc_luat.so_nguyen(b, "Buoc nhay")
        except ValueError as e:
            raise LoiSua(str(e))
        if tc not in ten:
            raise LoiSua("Khong co cot '%s' trong du lieu." % tc)
        c = ten.index(tc)
        idx = tuple(range(bd - 1, bd - 1 + so_diem * b, b))
        if idx[-1] >= n:
            raise LoiSua(
                "Day 'cot %s, bat dau %d, buoc %d' can toi mau thu %d nhung "
                "du lieu chi co %d mau." % (tc, bd, b, idx[-1] + 1, n))
        them(idx, c)

    m = d.get("meta") or {}
    if giu_bao_dam and m:
        for idx, name in m.get("khoi_da_kiem_tra", []):
            if name in ten:
                them(tuple(idx), ten.index(name))
        moi = range(len(ten))
        for k in m.get("buoc_moi_diem") or []:
            k = int(k)
            for s in range(n - so_diem * k + k + 1):
                for c in moi:
                    them(tuple(range(s, s + so_diem * k, k)), c)
        for k in m.get("buoc_nhay") or []:
            k = int(k)
            for o in range(k):
                for c in moi:
                    them(tuple(range(o, o + so_diem * k, k)), c)
        for x in m.get("day_rieng") or []:
            bd, k = int(x[0]), int(x[1])
            for c in moi:
                them(tuple(range(bd - 1, bd - 1 + so_diem * k, k)), c)
    return ra


# ------------------------------------------------------------- 3. HOAN VI
def sua(d, khoi, cpk_toi_thieu=1.33, giay_toi_da=120, seed=None,
        tien_do=None, dung_lai=None):
    """Hoan vi cac ban ghi cho toi khi moi khoi deu dat.

    Tra ve (so_khoi_con_hong, so_vong, giay, so_mau_da_doi_cho).
    """
    bg = d["ban_ghi"]
    n = len(bg)
    ten = d["ten_cot"]
    lsl = [d["nguong"].get(t) for t in ten]
    rnd = random.Random(seed)
    ban_dau = [id(b) for b in bg]
    d.pop("_kiem_tra_sua", None)
    for i, b in enumerate(bg):
        b.setdefault("vi_tri_goc", i)
    meta = d.get("meta") or {}
    old_limits = meta.get("nguong") or d["nguong"]
    old_cpk = meta.get("cpk_toi_thieu") or cpk_toi_thieu
    try:
        spc_luat.kiem_tra_nguong(None, None, cpk_toi_thieu)
        spc_luat.kiem_tra_nguong(None, None, old_cpk)
    except ValueError as e:
        raise LoiSua(str(e))

    def cost(idx, c, limit, threshold):
        v = [bg[i]["cot"].get(ten[c]) for i in idx]
        try:
            if d["loai"] == "keo":
                import sinh_keo_core as sk
                return sk._dem_loi_kep(v, sk.COT_SPC.index(ten[c]), limit, threshold, ())
            return spc_luat.dem_loi(v, limit, None, threshold)
        except (ValueError, TypeError):
            return 100

    # Giu moi cot dang dat tren cac day dang chon va cac nhom lien tiep.
    protect = set()
    windows = {idx for idx, _ in khoi}
    windows.update(tuple(range(i, i + 25)) for i in range(0, n - 24, 25))
    for idx in windows:
        for c in range(len(ten)):
            if cost(idx, c, old_limits.get(ten[c]), old_cpk) == 0:
                protect.add((idx, c))
                if (idx, c) not in khoi:
                    khoi.append((idx, c))
    # Metadata bao dam cu phai duoc cham lai voi dung nguong cu.
    if meta:
        protect.update(khoi)

    thuoc = {}
    for bi, (idx, _c) in enumerate(khoi):
        for i in idx:
            thuoc.setdefault(i, []).append(bi)

    def diem(bi):
        idx, c = khoi[bi]
        q = cost(idx, c, lsl[c], cpk_toi_thieu)
        if (idx, c) in protect:
            q += cost(idx, c, old_limits.get(ten[c]), old_cpk)
        return q

    chi_phi = [diem(b) for b in range(len(khoi))]
    t0 = datetime.now()
    vong = 0
    while True:
        hong = [b for b, x in enumerate(chi_phi) if x]
        if not hong:
            break
        giay = (datetime.now() - t0).total_seconds()
        if giay > giay_toi_da:
            break
        if dung_lai is not None and dung_lai():
            break
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
        bg[i], bg[j] = bg[j], bg[i]
        moi = {}
        sau_ = 0
        for b in cham:
            q = diem(b)
            moi[b] = q
            sau_ += q
        if sau_ <= truoc:
            for b, q in moi.items():
                chi_phi[b] = q
        else:
            bg[i], bg[j] = bg[j], bg[i]

    doi = sum(1 for a, b in zip(ban_dau, [id(x) for x in bg]) if a != b)
    # Recheck all constraints before authorizing any write.
    chi_phi = [diem(b) for b in range(len(khoi))]
    if not any(chi_phi) and not (dung_lai and dung_lai()):
        d["_kiem_tra_sua"] = {"hash": _dau_van(d), "cpk": cpk_toi_thieu,
                              "khoi": [(list(idx), ten[c]) for idx, c in khoi]}
    return (sum(1 for x in chi_phi if x), vong,
            (datetime.now() - t0).total_seconds(), doi)


# ------------------------------------------------------------- 4. GHI LAI
def _sao_luu(thu_muc, cac_file):
    bk = os.path.join(thu_muc, "_truoc_khi_sua")
    os.makedirs(bk, exist_ok=True)
    ra = []
    for f in cac_file:
        if os.path.exists(f):
            dich = os.path.join(bk, os.path.basename(f))
            if not os.path.exists(dich):        # giu ban goc dau tien
                shutil.copy2(f, dich)
            ra.append(dich)
    return bk, ra


def _dau_van(d):
    return hashlib.sha256(json.dumps(d["ban_ghi"], sort_keys=True, default=str).encode()).hexdigest()


def _tep_nguon(root):
    ra = {}
    for here, dirs, files in os.walk(root):
        dirs[:] = [x for x in dirs if x != "_truoc_khi_sua"]
        for name in files:
            path = os.path.join(here, name)
            with open(path, "rb") as f:
                ra[os.path.relpath(path, root)] = hashlib.file_digest(f, "sha256").hexdigest()
    return ra


def ghi_lai(d, danh_lai_ma=False, tien_do=None):
    """Chuan bi trong thu muc tam; sao luu tron bo, rollback neu ghi loi.

    Giu ma mau goc de doi chieu trang PDF. Bao cao phu thuoc cu chi con
    trong ban sao luu; khong duoc tiep tuc coi la bao cao cua thu tu moi.
    """
    proof = d.get("_kiem_tra_sua") or {}
    if proof.get("hash") != _dau_van(d):
        raise LoiSua("Chua sua dat/du lieu da thay doi: khong ghi de file goc.")
    if danh_lai_ma:
        raise LoiSua("Giu ma mau goc khi sap lai de bao toan truy vet Excel/PDF.")
    root = os.path.abspath(d["thu_muc"])
    original = d.get("_tep_nguon")
    if original is not None and original != _tep_nguon(root):
        raise LoiSua("File nguon da thay doi trong luc sua. Nap lai du lieu; khong ghi de.")
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f")
    backup = os.path.join(root, "_truoc_khi_sua", stamp)
    # Moi file/directory hien hanh deu duoc sao luu, tru lich su da co.
    entries = [p for p in os.listdir(root) if p != "_truoc_khi_sua"]
    os.makedirs(backup)
    for name in entries:
        src, dst = os.path.join(root, name), os.path.join(backup, name)
        if os.path.isdir(src):
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)
    replaced, archived = [], []
    with tempfile.TemporaryDirectory(prefix="spc_sua_") as stage:
        w = copy.deepcopy(d)
        w["thu_muc"] = stage
        w["file_xlsx"] = os.path.join(stage, os.path.basename(d["file_xlsx"]))
        shutil.copy2(d["file_xlsx"], w["file_xlsx"])
        if d["loai"] == "charpy":
            result = _ghi_charpy(w, False)
        else:
            result = _ghi_keo(w, False, tien_do)
            import pypdfium2 as pdfium
            for name in entries:
                if not (name.startswith("BieuDo_") and name.lower().endswith(".pdf")):
                    continue
                with pdfium.PdfDocument(os.path.join(root, name)) as src:
                    if len(src) != len(d["ban_ghi"]):
                        raise LoiSua("PDF lech so mau: khong ghi de. " + name)
                    with pdfium.PdfDocument.new() as dest:
                        dest.import_pages(src, [b["vi_tri_goc"] for b in d["ban_ghi"]])
                        dest.save(os.path.join(stage, name))
                result["file_da_ghi"].append(os.path.join(stage, name))
        meta = dict(d.get("meta") or {})
        meta.update({"ngay_sua": stamp, "so_mau": len(d["ban_ghi"]),
                     "nguong": d["nguong"], "cpk_toi_thieu": proof["cpk"],
                     "khoi_da_kiem_tra": proof["khoi"],
                     "anh_xa_vi_tri": [b["vi_tri_goc"] + 1 for b in d["ban_ghi"]]})
        for key in ("buoc_nhay", "buoc_moi_diem", "day_rieng"):
            meta.pop(key, None)  # chi cong bo cac khoi thuc su vua kiem tra
        for name, value in [("_thongtin_sinh.json", meta)]:
            with open(os.path.join(stage, name), "w", encoding="utf-8") as f:
                json.dump(value, f, ensure_ascii=False, indent=2)
            result["file_da_ghi"].append(os.path.join(stage, name))
        note = "DA SAP LAI THU TU. Bao cao SPC/Word/bang gon cu da luu tai " + backup + ". Can xuat lai bao cao.\nCpk va gioi han I-MR co the thay doi theo thu tu.\n"
        with open(os.path.join(stage, "SAU_KHI_SUA.txt"), "w", encoding="utf-8") as f:
            f.write(note)
        result["file_da_ghi"].append(os.path.join(stage, "SAU_KHI_SUA.txt"))
        try:
            if original is not None and original != _tep_nguon(root):
                raise LoiSua("File nguon da thay doi. Khong ap dung ket qua sua.")
            for f in result["file_da_ghi"]:
                target = os.path.join(root, os.path.basename(f))
                replaced.append(os.path.basename(f))
                shutil.copy2(f, target)
            for name in entries:
                if (name == "BaoCao" or name.lower().endswith(".docx") or
                    "SPC" in name or name.startswith("TongHop_VaDap_Gon")) and name not in replaced:
                    # Ban sao hoan chinh da ton tai trong backup.
                    archived.append(name)
                    path = os.path.join(root, name)
                    if os.path.commonpath([root, os.path.abspath(path)]) != root:
                        raise LoiSua("Duong dan luu tru khong nam trong thu muc ket qua.")
                    if os.path.isdir(path):
                        shutil.rmtree(path)
                    else:
                        os.remove(path)
        except Exception:
            for name in set(replaced + archived):
                src, dst = os.path.join(backup, name), os.path.join(root, name)
                if os.path.isdir(src):
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                elif os.path.isfile(src):
                    shutil.copy2(src, dst)
                elif os.path.isfile(dst):
                    os.remove(dst)
            raise
    return {"file_da_ghi": [os.path.join(root, x) for x in replaced],
            "sao_luu": backup, "canh_bao": note}


def _ghi_charpy(d, danh_lai_ma):
    f = d["file_xlsx"]
    bk, _ = _sao_luu(d["thu_muc"], [f])
    wb = openpyxl.load_workbook(f)
    ws = wb.worksheets[0]
    for k, b in enumerate(d["ban_ghi"]):
        r = 12 + k
        gt = list(b["o"])
        if danh_lai_ma:
            gt[0] = k + 1                       # STT theo vi tri moi
        for c, v in enumerate(gt, 1):
            ws.cell(r, c).value = v
    wb.save(f)
    return {"file_da_ghi": [f], "sao_luu": bk,
            "canh_bao": _canh_bao_word(d["thu_muc"], "KetQua_*.docx")}


def _ghi_keo(d, danh_lai_ma, tien_do=None):
    f = d["file_xlsx"]
    pdf = sorted(glob.glob(os.path.join(d["thu_muc"], "BieuDo_*.pdf")))
    bk, _ = _sao_luu(d["thu_muc"], [f] + pdf)

    wb = openpyxl.load_workbook(f)
    th, cl = wb["Tong_hop"], wb["Co_ly_tinh"]
    for k, b in enumerate(d["ban_ghi"]):
        for c, v in enumerate(b["th"], 1):
            th.cell(2 + k, c).value = v
        for c, v in enumerate(b["cl"], 1):
            cl.cell(3 + k, c).value = v
    if danh_lai_ma:
        # Ma mau chay lai theo vi tri: M0001 van la mau dau tien.
        goc = [b["th"][0] for b in d["ban_ghi"]]
        moi = sorted(goc)
        for k in range(len(d["ban_ghi"])):
            th.cell(2 + k, 1).value = moi[k]
            cl.cell(3 + k, 1).value = moi[k]
    # Reorder curve columns from the original workbook; never redraw measured curves.
    if "Du_lieu" in wb.sheetnames:
        ws = wb["Du_lieu"]
        cols = {str(ws.cell(1, c).value): [ws.cell(r, c).value for r in range(1, ws.max_row + 1)]
                for c in range(2, ws.max_column + 1)}
        for k, b in enumerate(d["ban_ghi"], 2):
            values = cols.get(str(b["th"][0]) + "_Load_kN")
            if values is None:
                raise LoiSua("Thieu duong cong cua mau " + str(b["th"][0]))
            for row, value in enumerate(values, 1):
                ws.cell(row, k).value = value
        if "Bieu_do_mau" in wb.sheetnames:
            example = wb["Bieu_do_mau"]
            for row in range(1, ws.max_row + 1):
                for col in range(1, min(4, ws.max_column) + 1):
                    example.cell(row, col).value = ws.cell(row, col).value
    if "Huong_dan" in wb.sheetnames:
        wb["Huong_dan"].append(["Da sap lai thu tu; giu ma mau goc. Can xuat lai bao cao SPC/Word."])
    wb.save(f)

    return {"file_da_ghi": [f], "sao_luu": bk, "canh_bao": None}


def _canh_bao_word(thu_muc, mau):
    n = len(glob.glob(os.path.join(thu_muc, "BaoCao", mau)))
    if n:
        return ("Thu muc con BaoCao/ dang co %d bao cao Word sinh TRUOC khi "
                "sua, gio khong con khop du lieu nua. Hay xoa roi chay lai "
                "Tab 2." % n)
    return None


def tom_tat(d, khoi, con_hong, vong, giay, doi, kq_ghi):
    L = []
    L.append("=" * 66)
    L.append("  DA SAP LAI THU TU DU LIEU CHO DAT SPC")
    L.append("=" * 66)
    L.append("  So mau           : %s" % "{:,}".format(len(d["ban_ghi"])))
    L.append("  Khoi rang buoc   : %d" % len(khoi))
    L.append("  Da doi cho       : %d mau" % doi)
    L.append("  Thoi gian        : %d vong, %.1f giay" % (vong, giay))
    if con_hong:
        L.append("  !! CON %d KHOI CHUA DAT - het thoi gian truoc khi xong."
                 % con_hong)
    else:
        L.append("  Ket qua          : moi khoi deu DAT")
    L.append("-" * 66)
    L.append("  KHONG co con so nao bi bia them: chi doi cho cac mau cho nhau,")
    L.append("  tap gia tri giu nguyen; MR, sigma, Cpk co the thay doi.")
    L.append("-" * 66)
    for f in kq_ghi["file_da_ghi"]:
        L.append("  Da ghi lai : %s" % os.path.basename(f))
    L.append("  Ban goc    : %s" % os.path.basename(kq_ghi["sao_luu"]) + os.sep)
    if kq_ghi.get("canh_bao"):
        L.append("")
        L.append("  LUU Y: %s" % kq_ghi["canh_bao"])
    L.append("=" * 66)
    return "\n".join(L)
