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

import os
import re
import math
import random
import unicodedata

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
_DON_VI = r"(?:mm|㎜|°|º)?"


def _phan_tich_dung_sai(chuoi_yeu_cau):
    """Phan tich chuoi yeu cau ky thuat thanh (gia_tri_chuan, ds_duoi, ds_tren).

    Vi du:
        '900.0(0~+15mm)'       -> (900.0, 0, 15)
        '355.6(-1~+3mm)'       -> (355.6, -1, 3)
        '217.0\\n(-2～+2mm)'   -> (217.0, -2, 2)
        '45±0.5'               -> (45, -0.5, 0.5)
        '0.0\\n(-0,5～+0.5°)'  -> (0.0, -0.5, 0.5)
        '6.0(0㎜～)'           -> (6.0, 0, None)      chi co gioi han duoi
        '100.0' / '98.1mm'     -> (100.0, None, None) kich thuoc tham khao
    Khong doc duoc thi tra ve (None, None, None).
    """
    if chuoi_yeu_cau is None:
        return None, None, None
    s = str(chuoi_yeu_cau).strip()
    if not s:
        return None, None, None
    s = re.sub(r"[\r\n]+", " ", s)
    s = re.sub(r"(\d),(\d)", r"\1.\2", s)

    m = re.match(_SO + r"\s*±\s*" + _SO + r"\s*" + _DON_VI, s)
    if m:
        tol = abs(float(m.group(2)))
        return float(m.group(1)), -tol, tol

    m = re.match(_SO + r"\s*\(\s*" + _SO + r"\s*" + _DON_VI + r"\s*[～~]\s*"
                 + _SO + r"\s*" + _DON_VI + r"\s*\)", s)
    if m:
        return float(m.group(1)), float(m.group(2)), float(m.group(3))

    m = re.match(_SO + r"\s*\(\s*" + _SO + r"\s*" + _DON_VI + r"\s*[～~]\s*\)", s)
    if m:                                    # mot phia: chi co dung sai duoi
        return float(m.group(1)), float(m.group(2)), None

    m = re.match(_SO + r"\s*\(\s*[～~]\s*" + _SO + r"\s*" + _DON_VI + r"\s*\)", s)
    if m:                                    # mot phia: chi co dung sai tren
        return float(m.group(1)), None, float(m.group(2))

    m = re.match(
        _SO + r"\s*\(?\s*([-+]\d+(?:\.\d+)?)\s*" + _DON_VI
        + r"\s*(?:[,/;]|\s+)\s*([-+]\d+(?:\.\d+)?)\s*" + _DON_VI + r"\s*\)?$",
        s,
    )
    if m:
        d1, d2 = float(m.group(2)), float(m.group(3))
        return float(m.group(1)), min(d1, d2), max(d1, d2)

    m = re.fullmatch(_SO + r"\s*" + _DON_VI, s)
    if m:                                    # chi co con so, khong dung sai
        return float(m.group(1)), None, None
    return None, None, None


def _chu_yeu_cau(v):
    """Chuoi 'Yeu cau' de in. O so (ket qua cong thuc trong file dung sai,
    vd 100.00000000000001) thi in gon lai nhu Excel hien thi: '100'."""
    if isinstance(v, (int, float)):
        return "%g" % round(float(v), 2)
    return str(v).strip() if v is not None else ""


def goi_y_min_max(lsl, usl, chuan=None):
    """De xuat khoang [Min, Max] an toan khi file dung sai de trong Min/Max."""
    if lsl is not None and usl is not None and usl > lsl:
        t = usl - lsl
        tam = (lsl + usl) / 2.0
        if t <= 1.0:
            return round(tam, 1), round(tam + 0.1, 1)
        span = max(0.2, min(3.0, round(t * 0.22, 1)))
        lo = round(tam - span / 2.0, 1)
        hi = round(lo + span, 1)
        return lo, hi
    if lsl is not None:
        return round(lsl + 1.0, 1), round(lsl + 2.0, 1)
    if usl is not None:
        return round(usl - 2.0, 1), round(usl - 1.0, 1)
    if chuan is not None:
        return round(chuan - 0.5, 1), round(chuan + 0.5, 1)
    return None, None


def _rut_ky_hieu_chi_tieu(ten, kh_cot, ten_sheet=""):
    """Rut ky hieu chi tieu (Dw, d, d1, T1, t1, 15°...) tu ten hoac cot ky hieu.

    Trong sheet ALKTOP II cua file Dung sai, cot E la ky hieu DIEM DO (AD, BE,
    CF...), con ky hieu CHI TIEU nam o dau ten ('Dw(翼径)', '掘削刃補強部 角度(15°)').
    """
    kh_cot_s = str(kh_cot or "").strip()
    if str(ten_sheet).strip().upper() == "ALKTOP II":
        s = unicodedata.normalize("NFKC", str(ten or "")).strip()
        m = re.match(r"^([A-Za-z]+\d*)\s*\(", s)
        if m:
            return m.group(1)
        m = re.search(r"\((\d+(?:\.\d+)?°)\)", s)
        if m:
            return m.group(1)
    return kh_cot_s


def ten_sheet_chi_tieu(danh_sach_chi_tieu):
    """Tao danh sach ten sheet bao cao cho tung chi tieu theo dung quy uoc mau:
    - Chuan hoa NFKC (D₁ -> D1), bo ky hieu do (°), bo ky tu cam cua Excel.
    - Neu trung khong phan biet hoa/thuong voi chi tieu phia sau (vd T1 va t1,
      T2 va t2, h2 va H2 trong ALKTOP II) thi them dau cham '.' vao chi tieu
      dung truoc ('T1.', 'T2.', 'h2.').
    """
    goc = []
    for ct in danh_sach_chi_tieu:
        kh = unicodedata.normalize("NFKC", str(ct.get("ky_hieu") or "")).strip()
        kh = kh.replace("°", "").replace("º", "").strip()
        kh = re.sub(r'[\\/*?:\[\]]+', "", kh)
        if kh in ("", "-"):
            m_goc = re.search(r"\((\d+(?:\.\d+)?)\s*[°º]\)", str(ct.get("ten") or ""))
            if m_goc:
                kh = m_goc.group(1)
        goc.append(kh or ("STT%s" % ct.get("stt", "")))

    ket_qua = []
    da_dung = set()
    for i, s in enumerate(goc):
        co_trung_sau = any(
            s.lower() == goc[j].lower() and s != goc[j]
            for j in range(i + 1, len(goc))
        )
        ten = (s + ".") if co_trung_sau else s
        while ten.lower() in da_dung:
            ten = ten + "."
        da_dung.add(ten.lower())
        ket_qua.append(ten[:31])
    return ket_qua


def _doc_ten_tu_file_phu(duong_dan_ds):
    """Tim ten song ngu Viet-Nhat va ky hieu chi tieu tu file 'Bao cao tong hop'
    hoac 'JIS.QT.07.BM.01' nam cung thu muc voi file dung sai (neu co)."""
    ket_qua = {}
    thu_muc = os.path.dirname(os.path.abspath(duong_dan_ds))
    cac_thu_muc = [thu_muc]
    try:
        import bo_cau_hinh
        tm_mau = bo_cau_hinh.thu_muc_file_mau()
        if tm_mau and os.path.isdir(tm_mau) and tm_mau not in cac_thu_muc:
            cac_thu_muc.append(tm_mau)
    except Exception:
        pass

    for tm in cac_thu_muc:
        if not os.path.isdir(tm):
            continue
        # 1. Doc tu Bao cao tong hop - <SP>.xlsx hoac Bao cao tong hop.xlsx
        th_chung = os.path.join(tm, "Báo cáo tổng hợp.xlsx")
        if os.path.isfile(th_chung):
            try:
                wb = openpyxl.load_workbook(th_chung, data_only=True)
                for sh in wb.sheetnames:
                    ws = wb[sh]
                    mp = {}
                    for r in range(5, ws.max_row + 1):
                        stt = ws.cell(r, 1).value
                        ten = ws.cell(r, 2).value
                        kh = ws.cell(r, 6).value
                        if isinstance(stt, (int, float)) and ten:
                            mp[int(stt)] = (str(ten).strip(),
                                            str(kh).strip() if kh is not None else "")
                    if mp:
                        ket_qua.setdefault(sh.strip(), mp)
                wb.close()
            except Exception:
                pass

        for fn in sorted(os.listdir(tm)):
            if fn.startswith("Báo cáo tổng hợp - ") and fn.endswith(".xlsx"):
                sp = fn[len("Báo cáo tổng hợp - "):-len(".xlsx")].strip()
                try:
                    wb = openpyxl.load_workbook(os.path.join(tm, fn), data_only=True)
                    ws = wb.worksheets[0]
                    mp = {}
                    for r in range(5, ws.max_row + 1):
                        stt = ws.cell(r, 1).value
                        ten = ws.cell(r, 2).value
                        kh = ws.cell(r, 6).value
                        if isinstance(stt, (int, float)) and ten:
                            mp[int(stt)] = (str(ten).strip(),
                                            str(kh).strip() if kh is not None else "")
                    wb.close()
                    if mp:
                        ket_qua[sp] = mp
                except Exception:
                    pass
    return ket_qua


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
                    {"stt": 2, "ten": "Đường kính cánh/ 翼径", "ky_hieu": "Dw",
                     "ky_hieu_diem": ["AD", "BE", "CF"], "so_diem_do": 3,
                     "theo_ma": {"1031M": {"yeu_cau": "310.0(0～+15mm)",
                                 "min": 310.5, "max": 313.5, "chuan": 310.0,
                                 "ds_duoi": 0, "ds_tren": 15,
                                 "lsl": 310.0, "usl": 325.0}, ...}},
                    ...
                ]
            }, ...
        }
    STT chuan hoa bat dau tu 2 theo bieu mau moi (JIS & Bao cao tong hop).
    """
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    ten_phu_map = _doc_ten_tu_file_phu(duong_dan)
    ket_qua = {}

    for ten_sheet in wb.sheetnames:
        ws = wb[ten_sheet]

        ma_list, cot_ma = [], []
        for c in range(6, ws.max_column + 1, 3):
            val = ws.cell(7, c).value
            if val is not None and str(val).strip():
                ma_s = str(val).strip()
                if ma_s in ma_list:
                    idx_cu = ma_list.index(ma_s)
                    cot_ma[idx_cu] = c
                else:
                    ma_list.append(ma_s)
                    cot_ma.append(c)
        if not ma_list:
            continue

        hang_stt = []
        for r in range(10, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
                if ws.cell(r, 3).value is not None:
                    hang_stt.append(r)

        if not hang_stt:
            continue

        stt_dau = int(float(ws.cell(hang_stt[0], 2).value))
        cong_stt = 1 if stt_dau == 1 else 0
        ten_phu_sp = ten_phu_map.get(ten_sheet.strip(), {})

        chi_tieu_list = []
        for i, r in enumerate(hang_stt):
            r_ke = hang_stt[i + 1] if i + 1 < len(hang_stt) else None
            n = max(1, _so_hang_khoi(ws, r, r_ke))
            ten_goc_ds = str(ws.cell(r, 3).value).strip()
            stt_bc = int(float(ws.cell(r, 2).value)) + cong_stt

            kh_diem = []
            for x in range(r, r + n):
                k = ws.cell(x, 5).value
                kh_diem.append(str(k).strip() if k is not None else "")
            if not any(kh_diem[1:]):
                kh_diem = [kh_diem[0]] * n

            ky_hieu_ct = _rut_ky_hieu_chi_tieu(ten_goc_ds, kh_diem[0], ten_sheet)
            ten_hien_thi = ten_goc_ds
            if stt_bc in ten_phu_sp:
                ten_th, kh_th = ten_phu_sp[stt_bc]
                if ten_th:
                    ten_hien_thi = ten_th
                if kh_th:
                    ky_hieu_ct = kh_th

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

                lsl_r = round(lsl, 4) if lsl is not None else None
                usl_r = round(usl, 4) if usl is not None else None
                min_v = _so(ws.cell(r, c + 1).value)
                max_v = _so(ws.cell(r, c + 2).value)
                tu_goi_y = False
                if (min_v is None or max_v is None) and yeu_cau:
                    g_min, g_max = goi_y_min_max(lsl_r, usl_r, chuan)
                    if min_v is None:
                        min_v = g_min
                    if max_v is None:
                        max_v = g_max
                    tu_goi_y = True

                theo_ma[ma] = {
                    "yeu_cau": yeu_cau,
                    "min": min_v,
                    "max": max_v,
                    "tu_goi_y": tu_goi_y,
                    "chuan": chuan, "ds_duoi": ds_duoi, "ds_tren": ds_tren,
                    "lsl": lsl_r,
                    "usl": usl_r,
                }

            chi_tieu_list.append({
                "stt": stt_bc,
                "ten": ten_hien_thi,
                "ten_ds": ten_goc_ds,
                "ky_hieu": ky_hieu_ct,
                "ky_hieu_diem": kh_diem,
                "so_diem_do": n,
                "theo_ma": theo_ma,
            })

        ket_qua[ten_sheet] = {"ma_san_pham": ma_list, "chi_tieu": chi_tieu_list}

    wb.close()
    return ket_qua


def ghi_dung_sai(duong_dan, loai_sp, ma_sp, danh_sach_chi_tieu, luu_goc=True):
    """Ghi cap nhat thong so (Yeu cau, Min, Max) cua (loai_sp, ma_sp) truc tiep vao
    file 'Dung sai kich thuoc.xlsx' bang XML de giu nguyen 100% cong thuc va dinh dang."""
    import os
    import shutil
    import zipfile
    from datetime import datetime
    import xuat_kich_thuoc as xkt

    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    try:
        ten_sheet = next(
            (s for s in wb.sheetnames if s.strip().lower() == str(loai_sp or "").strip().lower()),
            None,
        )
        if not ten_sheet:
            raise ValueError("Không tìm thấy sheet '%s' trong file dung sai." % loai_sp)
        ws = wb[ten_sheet]

        ma_list, cot_ma = [], []
        for c in range(6, ws.max_column + 1, 3):
            val = ws.cell(7, c).value
            if val is not None and str(val).strip():
                ma_s = str(val).strip()
                if ma_s in ma_list:
                    cot_ma[ma_list.index(ma_s)] = c
                else:
                    ma_list.append(ma_s)
                    cot_ma.append(c)
        if str(ma_sp) not in ma_list:
            raise ValueError("Không tìm thấy mã '%s' trong sheet '%s'." % (ma_sp, ten_sheet))
        c_yc = cot_ma[ma_list.index(str(ma_sp))]

        hang_stt = []
        for r in range(10, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
                if ws.cell(r, 3).value is not None:
                    hang_stt.append(r)
        if not hang_stt:
            raise ValueError("Sheet '%s' không có dòng chỉ tiêu nào." % ten_sheet)

        stt_dau = int(float(ws.cell(hang_stt[0], 2).value))
        cong_stt = 1 if stt_dau == 1 else 0

        def _so_cu(v):
            try:
                return float(v) if v is not None and str(v).strip() != "" else None
            except (ValueError, TypeError):
                return None

        map_stt = {}
        for r in hang_stt:
            stt_bc = int(float(ws.cell(r, 2).value)) + cong_stt
            yc_cu = _chu_yeu_cau(ws.cell(r, c_yc).value)
            mn_cu = _so_cu(ws.cell(r, c_yc + 1).value)
            mx_cu = _so_cu(ws.cell(r, c_yc + 2).value)
            map_stt[stt_bc] = (r, yc_cu, mn_cu, mx_cu)
    finally:
        wb.close()

    def _norm_yc(s):
        return " ".join(str(s or "").split())

    thay_doi = []
    for ct in danh_sach_chi_tieu:
        stt = ct["stt"]
        if stt not in map_stt:
            continue
        r, yc_cu, mn_cu, mx_cu = map_stt[stt]
        yc_moi = str(ct.get("yeu_cau") or "").strip()
        if yc_moi and _norm_yc(yc_moi) != _norm_yc(yc_cu):
            try:
                val_yc = float(yc_moi)
            except ValueError:
                val_yc = yc_moi
            thay_doi.append((r, c_yc, val_yc))

        mn_moi = ct.get("min")
        if mn_moi is not None and (mn_cu is None or abs(float(mn_moi) - mn_cu) > 1e-9):
            thay_doi.append((r, c_yc + 1, float(mn_moi)))

        mx_moi = ct.get("max")
        if mx_moi is not None and (mx_cu is None or abs(float(mx_moi) - mx_cu) > 1e-9):
            thay_doi.append((r, c_yc + 2, float(mx_moi)))

    if not thay_doi:
        return 0

    if luu_goc and os.path.isfile(duong_dan):
        tm_bk = os.path.join(os.path.dirname(os.path.abspath(duong_dan)), "_truoc_khi_sua")
        os.makedirs(tm_bk, exist_ok=True)
        ten_goc, duoi = os.path.splitext(os.path.basename(duong_dan))
        dich_bk = os.path.join(tm_bk, "%s_%s%s" % (ten_goc, datetime.now().strftime("%Y%m%d_%H%M%S"), duoi))
        shutil.copy2(duong_dan, dich_bk)

    with zipfile.ZipFile(duong_dan, "r") as z:
        paths = xkt._sheet_paths(z)
        p_sh = paths[ten_sheet]
        entries = [(info.filename, z.read(info.filename)) for info in z.infolist()]

    by_name = dict(entries)
    dau, hang, duoi_xml = xkt._tach_hang(by_name[p_sh].decode("utf-8"))
    for r, col_idx, val_moi in thay_doi:
        ref = "%s%d" % (xkt._cot(col_idx), r)
        attr, nd = hang.get(r, ("", ""))
        nd = xkt._dat_o(nd, ref, xkt._o(ref, None, val_moi))
        hang[r] = (attr, nd)

    xml_moi = dau + "".join(
        xkt._hang(r_k, hang[r_k][0], [hang[r_k][1]]) for r_k in sorted(hang)
    ) + duoi_xml
    by_name[p_sh] = xml_moi.encode("utf-8")

    # Yeu cau Excel tinh lai neu co o cong thuc phu thuoc
    if "xl/workbook.xml" in by_name:
        wb_xml = by_name["xl/workbook.xml"].decode("utf-8")
        wb_xml = re.sub(
            r"<calcPr([^>]*?)\s*/>",
            lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)),
            wb_xml,
        )
        by_name["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tam = duong_dan + ".tmp"
    with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
        for fn, _old in entries:
            out.writestr(fn, by_name[fn], compress_type=zipfile.ZIP_DEFLATED)
    os.replace(tam, duong_dan)
    return len(thay_doi)


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
        loi=[], mr3=False)
    if d["sigma"] <= 0:
        r.loi = ["SIGMA"] + (["I7"] if len(X) >= 15 else [])
        if any((lsl is not None and x < lsl) or (usl is not None and x > usl) for x in X):
            r.loi.append("SPEC")
        return r
    xb = d["xb"]
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

    # MR3 chi mang tinh tham khao (khong dua vao r.loi)
    run_mr, dau_mr = 0, 0
    for m in d["MR"]:
        s = 1 if tren(m, d["mrb"]) else (-1 if duoi(m, d["mrb"]) else 0)
        run_mr = 0 if s == 0 else (run_mr + 1 if s == dau_mr else 1)
        dau_mr = s
        if run_mr >= 9:
            r.mr3 = True
            break

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


def cham_chi_tieu(ct, cpk_toi_thieu=CPK_TOI_THIEU, so_mau_yeu_cau=None):
    """SPC tren trung binh; nghiem thu tren moi diem do rieng le.
    Ho tro day mau tu 2..SO_MAU diem (mac dinh cho phep 2..25 diem)."""
    lsl, usl, cpk = spc_luat.kiem_tra_nguong(ct.get("lsl"), ct.get("usl"), cpk_toi_thieu)
    X = ct.get("X", ct.get("xbar", []))
    raw = ct.get("du_lieu", [])
    so_diem = ct.get("so_diem_do")
    try:
        if ct.get("loi_du_lieu"):
            raise ValueError("Thieu so do")
        if raw:
            n_raw = len(raw)
            if so_mau_yeu_cau is not None and n_raw != so_mau_yeu_cau:
                raise ValueError("Khong du so mau yeu cau")
            if n_raw < 2 or n_raw > SO_MAU:
                raise ValueError("So mau ngoai pham vi 2..25")
            if any(not p or (so_diem and len(p) != so_diem) for p in raw):
                raise ValueError("Thieu diem do")
            for p in raw:
                for v in p:
                    spc_luat.so_huu_han(v)
            raw = [[float(v) for v in p] for p in raw]
            X = [math.fsum(p) / len(p) for p in raw]
        else:
            if so_mau_yeu_cau is not None and len(X) != so_mau_yeu_cau:
                raise ValueError("Khong du so mau yeu cau")
            if len(X) < 2 or len(X) > SO_MAU:
                raise ValueError("So mau ngoai pham vi 2..25")
        r = cham_bieu_mau(X, lsl, usl, cpk)
    except ValueError:
        r = spc_luat.KetQuaSPC(**{k: None for k in spc_luat.KetQuaSPC.__slots__})
        r.n, r.X, r.MR, r.lsl, r.usl = len(X), X, [], lsl, usl
        r.loi = ["DU_LIEU"]
        r.mr3 = False
        return r
    if any((lsl is not None and v < lsl - SAI_SO) or
           (usl is not None and v > usl + SAI_SO) for p in raw for v in p):
        if "SPEC" not in r.loi:
            r.loi.append("SPEC")
    return r


def khoang_min_max_an_toan(ct):
    """Tra ve khoang (min_v, max_v) hop le va an toan trong [lsl, usl] cho chi tieu."""
    lsl, usl, chuan = ct.get("lsl"), ct.get("usl"), ct.get("chuan")
    min_v, max_v = ct.get("min"), ct.get("max")
    so = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)  # noqa: E731
    g_min, g_max = goi_y_min_max(lsl, usl, chuan)
    if not so(min_v):
        min_v = g_min
    if not so(max_v):
        max_v = g_max
    if min_v is not None and lsl is not None and min_v < lsl:
        min_v = lsl
    if max_v is not None and usl is not None and max_v > usl:
        max_v = usl
    if min_v is None or max_v is None or min_v >= max_v:
        min_v, max_v = g_min, g_max
    if min_v is None or max_v is None or min_v >= max_v:
        base = chuan if so(chuan) else (lsl if so(lsl) else (usl if so(usl) else 10.0))
        min_v, max_v = round(base - 0.5, 1), round(base + 0.5, 1)
    return float(min_v), float(max_v)


# ---------------------------------------------------------------------------
# 3. SINH DU LIEU NGAU NHIEN CO KIEM SOAT, CHIA DAY & HOAN VI / TU SUA SPC
# ---------------------------------------------------------------------------
def _sinh_mau_gauss(mu, sigma, lo, hi, n):
    """Sinh n gia tri phan phoi chuan N(mu, sigma^2), clamp trong [lo, hi]."""
    vals = []
    for _ in range(n):
        v = random.gauss(mu, sigma)
        v = max(lo, min(hi, v))
        vals.append(round(v, 1))
    return vals


def _sinh_1_nhom_chi_tieu(min_val, max_val, lsl, usl, so_mau=SO_MAU,
                          so_diem_do=1, cpk_toi_thieu=CPK_TOI_THIEU):
    """Sinh 1 nhom (2..25 mau) cho 1 chi tieu sao cho dat 100% luat SPC."""
    lo_gen = max(min_val, lsl) if lsl is not None else min_val
    hi_gen = min(max_val, usl) if usl is not None else max_val
    if lo_gen >= hi_gen:
        lo_gen, hi_gen = min_val, max_val

    tam = (lo_gen + hi_gen) / 2.0
    ban_kinh = (hi_gen - lo_gen) / 2.0
    sigma_x = max(ban_kinh / 6.0, BUOC_DO * 0.6)
    sigma_sub = max(ban_kinh / 8.0, BUOC_DO * 0.5) if so_diem_do > 1 else 0.0

    def phat(mau):
        X = [math.fsum(p) / len(p) for p in mau]
        d = chi_so_bieu_mau(X, lsl, usl)
        r = cham_bieu_mau(X, lsl=lsl, usl=usl, cpk_toi_thieu=cpk_toi_thieu, d=d)
        if any((lsl is not None and v < lsl - SAI_SO) or
               (usl is not None and v > usl + SAI_SO) for p in mau for v in p):
            if "SPEC" not in r.loi:
                r.loi.append("SPEC")
        diem_phu = 0.1 if getattr(r, "mr3", False) else 0.0
        return len(r.loi) + diem_phu, r

    tot = None
    cac_cau_hinh = [(tam, sigma_x, sigma_sub, SO_LAN_SINH_LAI)]
    # Cau hinh du phong neu Min/Max rong hoac lech tam gay kho Cpk
    if lsl is not None and usl is not None and usl > lsl:
        tam_chuan = max(lo_gen + 0.05, min(hi_gen - 0.05, (lsl + usl) / 2.0))
        sg_hep = max(min(ban_kinh / 8.0, (usl - lsl) / (6.0 * max(cpk_toi_thieu, 1.33) * 1.35)),
                     BUOC_DO * 0.45)
        cac_cau_hinh.append((tam_chuan, sg_hep, sg_hep * 0.7 if so_diem_do > 1 else 0.0, 30))

    for mu_cur, sg_cur, sg_sub_cur, so_lan in cac_cau_hinh:
        for _ in range(so_lan):
            mau = []
            for _ in range(so_mau):
                t = max(lo_gen, min(hi_gen, random.gauss(mu_cur, sg_cur)))
                if so_diem_do == 1:
                    v1 = round(t, 1)
                    for _ in range(10):
                        if not mau or v1 != mau[-1][0] or (hi_gen - lo_gen) < BUOC_DO:
                            break
                        v1 = round(max(lo_gen, min(hi_gen, random.gauss(mu_cur, sg_cur))), 1)
                    mau.append([v1])
                else:
                    mau.append(_sinh_mau_gauss(t, sg_sub_cur, lo_gen, hi_gen, so_diem_do))
            p, r = phat(mau)
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
        if tot is not None and int(tot[0]) == 0:
            break

    _p, du_lieu_sub, r = tot
    x_list = [math.fsum(p) / len(p) for p in du_lieu_sub]
    return x_list, du_lieu_sub, r


def sinh_du_lieu_1_chi_tieu(min_val, max_val, lsl, usl, so_mau=SO_MAU,
                             so_diem_do=1, cpk_toi_thieu=CPK_TOI_THIEU,
                             seed=None):
    """Sinh du lieu cho 1 chi tieu kich thuoc (ho tro N mau bat ky).

    Neu so_mau > 25 (vd 1000 mau), tu dong chia thanh cac cum 25 gia tri lien
    tuc (buoc nhay 1: 1-25, 26-50, ..., 976-1000) va sinh sao cho MOI CUM deu
    thoa man day du luat SPC & Cpk ngay tu dau.

    Tra ve (x_list, du_lieu_diem_do, ket_qua_spc_tong_quat).
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

    if so_mau <= SO_MAU:
        return _sinh_1_nhom_chi_tieu(
            min_val, max_val, lsl, usl, so_mau=so_mau,
            so_diem_do=so_diem_do, cpk_toi_thieu=cpk_toi_thieu
        )

    x_all, raw_all, ds_kq = [], [], []
    for dau in range(0, so_mau, SO_MAU):
        m = min(SO_MAU, so_mau - dau)
        if m >= 2:
            x_sub, raw_sub, r_sub = _sinh_1_nhom_chi_tieu(
                min_val, max_val, lsl, usl, so_mau=m,
                so_diem_do=so_diem_do, cpk_toi_thieu=cpk_toi_thieu
            )
            x_all.extend(x_sub)
            raw_all.extend(raw_sub)
            ds_kq.append(r_sub)
        else:
            # Du 1 mau le cuoi cung (I-MR can >= 2 diem): sinh 1 mau hop le trong [Min, Max]
            lo_gen = max(min_val, lsl) if lsl is not None else min_val
            hi_gen = min(max_val, usl) if usl is not None else max_val
            if lo_gen >= hi_gen:
                lo_gen, hi_gen = min_val, max_val
            tam = (lo_gen + hi_gen) / 2.0
            sg = max((hi_gen - lo_gen) / 6.0, BUOC_DO * 0.5)
            p_le = _sinh_mau_gauss(tam, sg, lo_gen, hi_gen, so_diem_do)
            raw_all.append(p_le)
            x_all.append(math.fsum(p_le) / len(p_le))

    # Tong hop ket qua dai dien tu cac cum 25 mau
    r_tong = ds_kq[0]
    loi_hop = []
    cpk_min_thuc = None
    for r_i in ds_kq:
        for ma_loi in r_i.loi:
            if ma_loi not in loi_hop:
                loi_hop.append(ma_loi)
        if r_i.cpk is not None:
            cpk_min_thuc = r_i.cpk if cpk_min_thuc is None else min(cpk_min_thuc, r_i.cpk)
    r_tong.loi = loi_hop
    if cpk_min_thuc is not None:
        r_tong.cpk = cpk_min_thuc
    return x_all, raw_all, r_tong


def _tim_bo_ba_t2_h1_h2(danh_sach_chi_tieu):
    """Tim bo ba chi tieu (idx_t2, idx_h1, idx_h2) cua ALKTOP II co quan he
    cong thuc t2 = h1 - h2 (cung so diem do, h1 va h2 la kich thuoc tham khao)."""
    if not danh_sach_chi_tieu:
        return None
    vi_tri_kh = {}
    for idx, ct in enumerate(danh_sach_chi_tieu):
        kh = str(ct.get("ky_hieu") or "").strip()
        if kh in ("t2", "h1", "h2") and kh not in vi_tri_kh:
            vi_tri_kh[kh] = idx
    if len(vi_tri_kh) != 3:
        return None
    i_t2, i_h1, i_h2 = vi_tri_kh["t2"], vi_tri_kh["h1"], vi_tri_kh["h2"]
    c_t2, c_h1, c_h2 = danh_sach_chi_tieu[i_t2], danh_sach_chi_tieu[i_h1], danh_sach_chi_tieu[i_h2]
    n_t2 = max(1, int(c_t2.get("so_diem_do") or 1))
    n_h1 = max(1, int(c_h1.get("so_diem_do") or 1))
    n_h2 = max(1, int(c_h2.get("so_diem_do") or 1))
    if not (n_t2 == n_h1 == n_h2):
        return None
    if c_h1.get("lsl") is not None or c_h1.get("usl") is not None:
        return None
    if c_h2.get("lsl") is not None or c_h2.get("usl") is not None:
        return None
    return (i_t2, i_h1, i_h2)


def _khoang_sinh_bo_ba_t2_h1_h2(ct_t2, ct_h1, ct_h2):
    """Tinh khoang sinh dong bo (lo_t2, hi_t2, lo_h1, hi_h1, lo_h2, hi_h2)
    sao cho t2 = h1 - h2 luon dat dung sai/Cpk va h1, h2 co bien thien tu nhien."""
    min_t2, max_t2 = khoang_min_max_an_toan(ct_t2)
    min_h1, max_h1 = khoang_min_max_an_toan(ct_h1)
    min_h2, max_h2 = khoang_min_max_an_toan(ct_h2)
    lsl_t2, usl_t2 = ct_t2.get("lsl"), ct_t2.get("usl")

    le_h = min(0.25, max(0.1, (max_h1 - min_h1) * 0.2, (max_h2 - min_h2) * 0.2))
    lo_h = round(min_h1 - max_h2 + le_h, 2)
    hi_h = round(max_h1 - min_h2 - le_h, 2)

    # Uu tien 1: Giao truc tiep giua [min_t2, max_t2] va khoang hieu [h1 - h2]
    lo_1, hi_1 = max(min_t2, lo_h), min(max_t2, hi_h)
    if hi_1 - lo_1 >= 0.45:
        return lo_1, hi_1, min_h1, max_h1, min_h2, max_h2

    # Uu tien 2: Giao giua khoang an toan Cpk trong [lsl_t2, usl_t2] va [h1 - h2]
    if lsl_t2 is not None and usl_t2 is not None and usl_t2 - lsl_t2 > 1.6:
        le_ds = min(0.8, (usl_t2 - lsl_t2) * 0.2)
        lo_2 = max(round(lsl_t2 + le_ds, 2), lo_h)
        hi_2 = min(round(usl_t2 - le_ds, 2), hi_h)
        if hi_2 - lo_2 >= 0.45:
            return lo_2, hi_2, min_h1, max_h1, min_h2, max_h2

    # Uu tien 3: Giu nguyen [min_t2, max_t2] va [min_h2, max_h2], tinh tien tham khao h1
    lo_h1_moi = round(min_h2 + min_t2, 1)
    hi_h1_moi = round(max_h2 + max_t2, 1)
    if hi_h1_moi - lo_h1_moi < 0.6:
        hi_h1_moi = round(lo_h1_moi + 0.8, 1)
    return min_t2, max_t2, lo_h1_moi, hi_h1_moi, min_h2, max_h2


def _sinh_1_nhom_t2_h1_h2(ct_t2, ct_h1, ct_h2, so_mau=SO_MAU, cpk_toi_thieu=CPK_TOI_THIEU):
    """Sinh dong bo 1 nhom (2..25 mau) cho (t2, h1, h2) thoa t2 == round(h1 - h2, 1)
    va ca 3 chi tieu deu dat 100% luat SPC + Cpk."""
    so_diem = max(1, int(ct_t2.get("so_diem_do") or 1))
    lsl_t2, usl_t2 = ct_t2.get("lsl"), ct_t2.get("usl")
    lo_t2, hi_t2, lo_h1, hi_h1, lo_h2, hi_h2 = _khoang_sinh_bo_ba_t2_h1_h2(ct_t2, ct_h1, ct_h2)

    def _cham_raw(raw_list, lsl=None, usl=None, exact_diff=None):
        X = [math.fsum(p) / len(p) for p in raw_list]
        r = cham_bieu_mau(X, lsl=lsl, usl=usl, cpk_toi_thieu=cpk_toi_thieu)
        if any((lsl is not None and v < lsl - SAI_SO) or
               (usl is not None and v > usl + SAI_SO) for p in raw_list for v in p):
            if "SPEC" not in r.loi:
                r.loi.append("SPEC")
        if exact_diff is not None:
            X_ex = [math.fsum(p) / len(p) for p in exact_diff]
            r_ex = cham_bieu_mau(X_ex, lsl=lsl, usl=usl, cpk_toi_thieu=cpk_toi_thieu)
            for ml in r_ex.loi:
                if ml not in r.loi:
                    r.loi.append(ml)
        return r

    tot = None
    for _ in range(12):
        _xt2, mau_t2, _rt2 = _sinh_1_nhom_chi_tieu(
            lo_t2, hi_t2, lsl_t2, usl_t2,
            so_mau=so_mau, so_diem_do=so_diem, cpk_toi_thieu=cpk_toi_thieu
        )

        def _sinh_cap_h_cho_mau(p_t2):
            p_h1, p_h2, p_t2_moi, p_ex = [], [], [], []
            for v_t2 in p_t2:
                a_h2 = max(lo_h2, round(lo_h1 - v_t2, 4))
                b_h2 = min(hi_h2, round(hi_h1 - v_t2, 4))
                if a_h2 > b_h2:
                    a_h2 = b_h2 = round((lo_h2 + hi_h2) / 2.0, 1)
                tam_2 = (a_h2 + b_h2) / 2.0
                sg_2 = max((b_h2 - a_h2) / 4.5, BUOC_DO * 0.5)
                v2 = round(max(a_h2, min(b_h2, random.gauss(tam_2, sg_2))), 1)
                v1 = round(v2 + v_t2, 1)
                vt = round(v1 - v2, 1)
                p_h1.append(v1)
                p_h2.append(v2)
                p_t2_moi.append(vt)
                p_ex.append(v1 - v2)
            return p_t2_moi, p_h1, p_h2, p_ex

        for _ in range(25):
            m_t2, m_h1, m_h2, m_ex = [], [], [], []
            for j in range(so_mau):
                pt, p1, p2, pex = _sinh_cap_h_cho_mau(mau_t2[j])
                m_t2.append(pt)
                m_h1.append(p1)
                m_h2.append(p2)
                m_ex.append(pex)

            r_t2 = _cham_raw(m_t2, lsl_t2, usl_t2, exact_diff=m_ex)
            r_h1 = _cham_raw(m_h1)
            r_h2 = _cham_raw(m_h2)
            diem = len(r_t2.loi) + len(r_h1.loi) + len(r_h2.loi)

            # Neu con vuong luat I-MR tren h1/h2/t2 -> hoan vi cot chung hoac doi nhe h2
            for step in range(250):
                if diem == 0:
                    break
                if step % 2 == 0 and so_mau >= 2:
                    a, b = random.sample(range(so_mau), 2)
                    m_t2[a], m_t2[b] = m_t2[b], m_t2[a]
                    m_h1[a], m_h1[b] = m_h1[b], m_h1[a]
                    m_h2[a], m_h2[b] = m_h2[b], m_h2[a]
                    m_ex[a], m_ex[b] = m_ex[b], m_ex[a]
                    rt_n = _cham_raw(m_t2, lsl_t2, usl_t2, exact_diff=m_ex)
                    r1_n = _cham_raw(m_h1)
                    r2_n = _cham_raw(m_h2)
                    d_n = len(rt_n.loi) + len(r1_n.loi) + len(r2_n.loi)
                    if d_n <= diem:
                        diem, r_t2, r_h1, r_h2 = d_n, rt_n, r1_n, r2_n
                    else:
                        m_t2[a], m_t2[b] = m_t2[b], m_t2[a]
                        m_h1[a], m_h1[b] = m_h1[b], m_h1[a]
                        m_h2[a], m_h2[b] = m_h2[b], m_h2[a]
                        m_ex[a], m_ex[b] = m_ex[b], m_ex[a]
                else:
                    j_pick = random.randrange(so_mau)
                    old_pt, old_p1, old_p2, old_pex = m_t2[j_pick], m_h1[j_pick], m_h2[j_pick], m_ex[j_pick]
                    m_t2[j_pick], m_h1[j_pick], m_h2[j_pick], m_ex[j_pick] = _sinh_cap_h_cho_mau(old_pt)
                    rt_n = _cham_raw(m_t2, lsl_t2, usl_t2, exact_diff=m_ex)
                    r1_n = _cham_raw(m_h1)
                    r2_n = _cham_raw(m_h2)
                    d_n = len(rt_n.loi) + len(r1_n.loi) + len(r2_n.loi)
                    if d_n <= diem:
                        diem, r_t2, r_h1, r_h2 = d_n, rt_n, r1_n, r2_n
                    else:
                        m_t2[j_pick], m_h1[j_pick], m_h2[j_pick], m_ex[j_pick] = old_pt, old_p1, old_p2, old_pex

            if tot is None or diem < tot[0]:
                tot = (
                    diem,
                    ([math.fsum(p) / len(p) for p in m_t2], [list(p) for p in m_t2], r_t2),
                    ([math.fsum(p) / len(p) for p in m_h1], [list(p) for p in m_h1], r_h1),
                    ([math.fsum(p) / len(p) for p in m_h2], [list(p) for p in m_h2], r_h2),
                )
            if diem == 0:
                break
        if tot is not None and tot[0] == 0:
            break

    return tot[1], tot[2], tot[3]


def _sinh_nhieu_nhom_t2_h1_h2(ct_t2, ct_h1, ct_h2, so_mau=SO_MAU, cpk_toi_thieu=CPK_TOI_THIEU):
    """Sinh dong bo N mau bat ky (vd 25, 100, 1000) cho bo ba (t2, h1, h2)."""
    so_diem = max(1, int(ct_t2.get("so_diem_do") or 1))
    if so_mau <= SO_MAU:
        return _sinh_1_nhom_t2_h1_h2(ct_t2, ct_h1, ct_h2, so_mau=so_mau, cpk_toi_thieu=cpk_toi_thieu)

    acc = [([], [], []) for _ in range(3)]
    for dau in range(0, so_mau, SO_MAU):
        m = min(SO_MAU, so_mau - dau)
        if m >= 2:
            res3 = _sinh_1_nhom_t2_h1_h2(ct_t2, ct_h1, ct_h2, so_mau=m, cpk_toi_thieu=cpk_toi_thieu)
            for idx_k in range(3):
                acc[idx_k][0].extend(res3[idx_k][0])
                acc[idx_k][1].extend(res3[idx_k][1])
                acc[idx_k][2].append(res3[idx_k][2])
        else:
            lo_t2, hi_t2, lo_h1, hi_h1, lo_h2, hi_h2 = _khoang_sinh_bo_ba_t2_h1_h2(ct_t2, ct_h1, ct_h2)
            p_t2, p_h1, p_h2 = [], [], []
            for _ in range(so_diem):
                vt = round((lo_t2 + hi_t2) / 2.0, 1)
                v2 = round(max(lo_h2, min(hi_h2, (lo_h2 + hi_h2) / 2.0)), 1)
                v1 = round(v2 + vt, 1)
                vt = round(v1 - v2, 1)
                p_t2.append(vt)
                p_h1.append(v1)
                p_h2.append(v2)
            for idx_k, p_le in enumerate((p_t2, p_h1, p_h2)):
                acc[idx_k][1].append(p_le)
                acc[idx_k][0].append(math.fsum(p_le) / len(p_le))

    ra_3 = []
    for x_all, raw_all, ds_kq in acc:
        r_tong = ds_kq[0]
        loi_hop = []
        cpk_min = None
        for r_i in ds_kq:
            for ml in r_i.loi:
                if ml not in loi_hop:
                    loi_hop.append(ml)
            if r_i.cpk is not None:
                cpk_min = r_i.cpk if cpk_min is None else min(cpk_min, r_i.cpk)
        r_tong.loi = loi_hop
        if cpk_min is not None:
            r_tong.cpk = cpk_min
        ra_3.append((x_all, raw_all, r_tong))
    return tuple(ra_3)


def sinh_toan_bo(danh_sach_chi_tieu, so_mau=SO_MAU, callback=None, seed=None,
                 cpk_toi_thieu=CPK_TOI_THIEU, dung=None):
    """Sinh du lieu dong thoi cho tat ca chi tieu voi N mau bat ky (vd 25, 100, 1000).
    Dam bao moi cum 25 gia tri lien tuc (buoc nhay 1) deu dat chuan SPC ngay tu dau."""
    if seed is not None:
        random.seed(seed)
    so_mau = spc_luat.so_nguyen(so_mau, "Số lượng mẫu", 2)
    so_cum = sum(1 for d in range(0, so_mau, SO_MAU) if so_mau - d >= 2)

    bo_ba = _tim_bo_ba_t2_h1_h2(danh_sach_chi_tieu)
    du_lieu_bo_ba = {}
    if bo_ba is not None:
        i_t2, i_h1, i_h2 = bo_ba
        res_t2, res_h1, res_h2 = _sinh_nhieu_nhom_t2_h1_h2(
            danh_sach_chi_tieu[i_t2], danh_sach_chi_tieu[i_h1], danh_sach_chi_tieu[i_h2],
            so_mau=so_mau, cpk_toi_thieu=cpk_toi_thieu
        )
        du_lieu_bo_ba = {i_t2: res_t2, i_h1: res_h1, i_h2: res_h2}

    ket_qua = []
    for idx_ct, ct in enumerate(danh_sach_chi_tieu):
        if dung and dung():
            raise InterruptedError("Đã dừng sinh dữ liệu theo yêu cầu.")
        stt, ten = ct["stt"], ct.get("ten", "")
        so_diem_do = ct.get("so_diem_do", 1)
        min_v, max_v = khoang_min_max_an_toan(ct)

        if idx_ct in du_lieu_bo_ba:
            x, du_lieu, r_tong = du_lieu_bo_ba[idx_ct]
        else:
            x, du_lieu, r_tong = sinh_du_lieu_1_chi_tieu(
                min_v, max_v, ct.get("lsl"), ct.get("usl"), so_mau, so_diem_do,
                cpk_toi_thieu=cpk_toi_thieu)

        # Kiem tra lai tung cum 25 mau bang cham_chi_tieu (ke ca SPEC tung diem le)
        so_cum_dat = 0
        loi_cac_cum = []
        cpk_thap_nhat = None
        for d0 in range(0, so_mau, SO_MAU):
            m = min(SO_MAU, so_mau - d0)
            if m < 2:
                continue
            sub = dict(ct)
            sub.update({
                "du_lieu": du_lieu[d0:d0 + m],
                "X": x[d0:d0 + m],
                "xbar": x[d0:d0 + m],
                "so_diem_do": so_diem_do,
            })
            rk = cham_chi_tieu(sub, cpk_toi_thieu)
            if rk.dat:
                so_cum_dat += 1
            else:
                for ml in rk.loi:
                    if ml not in loi_cac_cum:
                        loi_cac_cum.append(ml)
            if rk.cpk is not None:
                cpk_thap_nhat = rk.cpk if cpk_thap_nhat is None else min(cpk_thap_nhat, rk.cpk)

        r_tong.loi = loi_cac_cum
        if cpk_thap_nhat is not None:
            r_tong.cpk = cpk_thap_nhat

        kq = dict(ct)
        kq.update({"min": min_v, "max": max_v, "xbar": x, "X": x,
                   "du_lieu": du_lieu, "so_diem_do": so_diem_do,
                   "cpk_toi_thieu": cpk_toi_thieu,
                   "so_cum": so_cum, "so_cum_dat": so_cum_dat,
                   "spc_result": r_tong})

        if callback:
            if r_tong.dat:
                cpk_txt = "" if cpk_thap_nhat is None else "  (Cpk min %.2f)" % cpk_thap_nhat
                if so_cum > 1:
                    trang_thai = "ĐẠT cả %d/%d dãy 25 mẫu%s" % (so_cum_dat, so_cum, cpk_txt)
                else:
                    trang_thai = "ĐẠT%s" % cpk_txt
            else:
                trang_thai = "CHƯA ĐẠT (%d/%d dãy): %s" % (
                    so_cum_dat, so_cum, ", ".join(loi_cac_cum))
            callback(stt, ten, r_tong.dat, so_cum_dat, trang_thai)

        ket_qua.append(kq)
    return ket_qua


# ---------------------------------------------------------------------------
# 4. CHIA DAY MAU, CHAM NHIEU DAY, QUET VA TU DONG SUA DU LIEU CHO DAT
# ---------------------------------------------------------------------------
def mac_dinh_bat_dau(so_mau, so_diem=SO_MAU):
    """Chia nhom lien tiep buoc nhay 1: 1, 26, 51... (bo nhom le chi co 1 mau)."""
    return [s for s in range(1, so_mau + 1, so_diem) if so_mau - s + 1 >= 2]


def lay_day(so_mau, bat_dau, buoc=1, so_diem=SO_MAU):
    """Tra ve danh sach chi so mau (1-based) cua mot day danh gia."""
    bat_dau = spc_luat.so_nguyen(bat_dau, "Mẫu bắt đầu", 1)
    buoc = spc_luat.so_nguyen(buoc, "Bước nhảy", 1)
    so_diem = spc_luat.so_nguyen(so_diem, "Số điểm", 2)
    if so_diem > SO_MAU:
        raise ValueError("Số điểm tối đa là %d (biểu mẫu SPC chỉ chứa tối đa %d mẫu)." % (SO_MAU, SO_MAU))
    vt = [v for v in range(bat_dau, bat_dau + buoc * so_diem, buoc) if v <= so_mau]
    if len(vt) < 2:
        raise ValueError("Bắt đầu từ mẫu %d (bước nhảy %d) chỉ lấy được %d mẫu — biểu đồ I-MR cần ít nhất 2 điểm (tổng có %d mẫu)."
                         % (bat_dau, buoc, len(vt), so_mau))
    return vt


def cat_day_chi_tieu(ct, vi_tri):
    """Cat du lieu cua 1 chi tieu theo danh sach vi_tri (1-based: 1..N)."""
    raw_all = ct.get("du_lieu") or []
    x_all = ct.get("X") or ct.get("xbar") or []
    so_diem = max(1, int(ct.get("so_diem_do") or 1))
    raw_sub, x_sub, errs = [], [], []
    for idx in vi_tri:
        j = idx - 1
        if raw_all:
            p = raw_all[j] if 0 <= j < len(raw_all) else None
            hop_le = (
                p is not None
                and len(p) == so_diem
                and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in p)
            )
            if not hop_le:
                errs.append("Mẫu %d thiếu/sai điểm đo" % idx)
                raw_sub.append(list(p) if p else [None] * so_diem)
                x_sub.append(None)
            else:
                pf = [float(v) for v in p]
                raw_sub.append(pf)
                x_sub.append(math.fsum(pf) / len(pf))
        else:
            v = x_all[j] if 0 <= j < len(x_all) else None
            if not (isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)):
                errs.append("Mẫu %d thiếu/sai điểm đo" % idx)
                x_sub.append(None)
            else:
                x_sub.append(float(v))
    sub = dict(ct)
    sub["du_lieu"] = raw_sub
    sub["X"] = x_sub
    sub["xbar"] = list(x_sub)
    sub["loi_du_lieu"] = errs
    return sub


def cham_day_kich_thuoc(danh_sach_chi_tieu, vi_tri, cpk_toi_thieu=CPK_TOI_THIEU):
    """Cham tat ca chi tieu tren mot day vi_tri (1-based).
    Tra ve danh sach cac ban sao chi tieu cua day do (da gan 'kq' va 'spc_result')."""
    ra = []
    for ct in danh_sach_chi_tieu:
        sub = cat_day_chi_tieu(ct, vi_tri)
        r = cham_chi_tieu(sub, cpk_toi_thieu)
        sub["kq"] = sub["spc_result"] = r
        sub["cpk_toi_thieu"] = cpk_toi_thieu
        ra.append(sub)
    return ra


def quet_kich_thuoc(danh_sach_chi_tieu, so_mau, buoc=1, so_diem=SO_MAU,
                    cpk_toi_thieu=CPK_TOI_THIEU):
    """Quet moi diem bat dau co du so_diem mau: [(bat_dau, dat_het, [ky_hieu_loi])]."""
    het = so_mau - (so_diem - 1) * buoc
    if het < 1:
        raise ValueError("Với bước nhảy %d và %d điểm cần ít nhất %d mẫu, dữ liệu hiện có %d mẫu."
                         % (buoc, so_diem, (so_diem - 1) * buoc + 1, so_mau))
    ra = []
    for s in range(1, het + 1):
        vt = list(range(s, s + buoc * so_diem, buoc))
        ds_sub = cham_day_kich_thuoc(danh_sach_chi_tieu, vt, cpk_toi_thieu)
        hong = ["%s" % (c.get("ky_hieu") or ("STT%s" % c["stt"])) for c in ds_sub if not c["kq"].dat]
        ra.append((s, not hong, hong))
    return ra


def sua_hoan_vi_kich_thuoc(danh_sach_chi_tieu, cpk_toi_thieu=CPK_TOI_THIEU,
                           so_buoc=6000, seed=None):
    """Sap xep lai thu tu cac mau trong 1 nhom (2..25 mau, KHONG doi so do)
    de cac chi tieu dat SPC."""
    if seed is not None:
        random.seed(seed)

    ds_moi = []
    for ct in danh_sach_chi_tieu:
        ban_sao = dict(ct)
        if ct.get("du_lieu"):
            ban_sao["du_lieu"] = [list(p) for p in ct["du_lieu"]]
        if ct.get("X"):
            ban_sao["X"] = list(ct["X"])
        ds_moi.append(ban_sao)

    n_mau_nhom = len(ds_moi[0].get("du_lieu") or []) if ds_moi else 0
    co_the_sua = []
    for ct in ds_moi:
        r = cham_chi_tieu(ct, cpk_toi_thieu)
        ct["kq"] = r
        ct["spc_result"] = r
        if ("DU_LIEU" not in r.loi and "SPEC" not in r.loi
                and 2 <= len(ct.get("du_lieu", [])) == n_mau_nhom <= SO_MAU):
            co_the_sua.append(ct)

    if n_mau_nhom >= 2 and co_the_sua:
        def tong_loi(perm):
            diem = 0
            for ct in co_the_sua:
                raw_p = [ct["du_lieu"][k] for k in perm]
                X_p = [math.fsum(p) / len(p) for p in raw_p]
                r = cham_bieu_mau(X_p, ct.get("lsl"), ct.get("usl"), cpk_toi_thieu)
                diem += len(r.loi)
            return diem

        perm = list(range(n_mau_nhom))
        p_hien_tai = tong_loi(perm)
        perm_tot = list(perm)
        p_tot = p_hien_tai

        # Buoc 1: Hoan vi chung cac cot mau tren toan bo chi tieu
        if p_tot > 0:
            for _ in range(min(so_buoc, 2500)):
                if p_tot == 0:
                    break
                i, j = random.sample(range(n_mau_nhom), 2)
                perm[i], perm[j] = perm[j], perm[i]
                p2 = tong_loi(perm)
                if p2 <= p_hien_tai:
                    p_hien_tai = p2
                    if p2 < p_tot:
                        p_tot = p2
                        perm_tot = list(perm)
                else:
                    perm[i], perm[j] = perm[j], perm[i]

            for ct in co_the_sua:
                ct["du_lieu"] = [ct["du_lieu"][k] for k in perm_tot]
                ct["X"] = [math.fsum(p) / len(p) for p in ct["du_lieu"]]
                ct["xbar"] = list(ct["X"])
                r = cham_chi_tieu(ct, cpk_toi_thieu)
                ct["kq"] = ct["spc_result"] = r

        # Buoc 2: Voi chi tieu van con loi, hoan vi mau cua rieng chi tieu do
        # (Neu la bo ba t2 = h1 - h2 thi hoan vi dong bo ca 3 de giu nguyen cong thuc)
        bo_ba_hv = _tim_bo_ba_t2_h1_h2(ds_moi)
        tap_bo_ba_hv = set(bo_ba_hv) if bo_ba_hv else set()
        if bo_ba_hv and any(not ds_moi[idx_b]["kq"].dat for idx_b in bo_ba_hv):
            cts_3 = [ds_moi[idx_b] for idx_b in bo_ba_hv]
            if all(c in co_the_sua for c in cts_3):
                maus_3 = [[list(p) for p in c["du_lieu"]] for c in cts_3]

                def phat_3(m3):
                    tong_p = 0
                    for idx_k, c_k in enumerate(cts_3):
                        X_k = [math.fsum(p) / len(p) for p in m3[idx_k]]
                        tong_p += len(cham_bieu_mau(X_k, c_k.get("lsl"), c_k.get("usl"), cpk_toi_thieu).loi)
                    return tong_p

                p3 = phat_3(maus_3)
                tot_3 = [[list(p) for p in m_k] for m_k in maus_3]
                tot_p3 = p3
                for _ in range(so_buoc):
                    if tot_p3 == 0:
                        break
                    i, j = random.sample(range(n_mau_nhom), 2)
                    for m_k in maus_3:
                        m_k[i], m_k[j] = m_k[j], m_k[i]
                    p3_moi = phat_3(maus_3)
                    if p3_moi <= p3:
                        p3 = p3_moi
                        if p3_moi < tot_p3:
                            tot_p3 = p3_moi
                            tot_3 = [[list(p) for p in m_k] for m_k in maus_3]
                    else:
                        for m_k in maus_3:
                            m_k[i], m_k[j] = m_k[j], m_k[i]
                for c_k, tot_mk in zip(cts_3, tot_3):
                    c_k["du_lieu"] = tot_mk
                    c_k["X"] = [math.fsum(p) / len(p) for p in tot_mk]
                    c_k["xbar"] = list(c_k["X"])
                    c_k["kq"] = c_k["spc_result"] = cham_chi_tieu(c_k, cpk_toi_thieu)

        for ct in co_the_sua:
            if any(ct is ds_moi[idx_b] for idx_b in tap_bo_ba_hv):
                continue
            r = ct["kq"]
            if r.dat:
                continue
            mau = [list(p) for p in ct["du_lieu"]]
            lsl, usl = ct.get("lsl"), ct.get("usl")

            def phat_1(m_list):
                X_l = [math.fsum(p) / len(p) for p in m_list]
                kq_r = cham_bieu_mau(X_l, lsl, usl, cpk_toi_thieu)
                return len(kq_r.loi), kq_r

            p1, r1 = phat_1(mau)
            tot_mau = [list(p) for p in mau]
            tot_p = p1
            for _ in range(so_buoc):
                if tot_p == 0:
                    break
                i, j = random.sample(range(n_mau_nhom), 2)
                mau[i], mau[j] = mau[j], mau[i]
                p2, _r2 = phat_1(mau)
                if p2 <= p1:
                    p1 = p2
                    if p2 < tot_p:
                        tot_p = p2
                        tot_mau = [list(p) for p in mau]
                else:
                    mau[i], mau[j] = mau[j], mau[i]
            ct["du_lieu"] = tot_mau
            ct["X"] = [math.fsum(p) / len(p) for p in tot_mau]
            ct["xbar"] = list(ct["X"])
            ct["kq"] = ct["spc_result"] = cham_chi_tieu(ct, cpk_toi_thieu)

    so_dat = 0
    tong_loi_con = 0
    khong_the_sua = []
    for ct_goc, ct_moi in zip(danh_sach_chi_tieu, ds_moi):
        ct_goc["du_lieu"] = ct_moi.get("du_lieu")
        ct_goc["X"] = ct_moi.get("X")
        ct_goc["xbar"] = ct_moi.get("xbar")
        ct_goc["kq"] = ct_moi["kq"]
        ct_goc["spc_result"] = ct_moi["spc_result"]
        if ct_moi["kq"].dat:
            so_dat += 1
        else:
            tong_loi_con += len(ct_moi["kq"].loi)
            khong_the_sua.append("STT %s (%s: %s)" % (
                ct_moi.get("stt"), ct_moi.get("ky_hieu"), ", ".join(ct_moi["kq"].loi)
            ))

    return {
        "dat_het": so_dat == len(ds_moi),
        "so_dat": so_dat,
        "tong_loi": tong_loi_con,
        "khong_the_sua": khong_the_sua,
        "chi_tieu": danh_sach_chi_tieu,
    }


def sua_du_lieu_nhieu_day(danh_sach_chi_tieu, cac_day, cpk_toi_thieu=CPK_TOI_THIEU,
                          cho_phep_chinh_gia_tri=True, seed=None, dung=None,
                          tien_do=None):
    """Kiem tra va tu dong sua du lieu tren nhieu day mau (vd 40 day cua 1000 mau)
    truc tiep vao danh_sach_chi_tieu (du_lieu cua Bao cao tong hop):
      - Uu tien 1: Hoan vi thu tu mau trong day bi loi (giu nguyen so do).
      - Uu tien 2: Neu van con chi tieu bi vuong SPEC / DU_LIEU / I1 / Cpk ma
                   hoan vi khong the cuu duoc va cho_phep_chinh_gia_tri=True,
                   tinh chinh / sinh lai gia tri chuan trong [Min, Max] cho dung
                   day do de dam bao 100% dat chuan SPC.
    Tra ve dict ket qua chi tiet.
    """
    if seed is not None:
        random.seed(seed)
    if not danh_sach_chi_tieu or not cac_day:
        return {"dat_het": True, "so_day_dat": len(cac_day), "tong_day": len(cac_day),
                "so_day_sua": 0, "so_ct_hoan_vi": 0, "so_ct_chinh_so": 0, "con_hong": []}

    # Dam bao moi chi tieu deu co danh sach X day du
    for ct in danh_sach_chi_tieu:
        raw = ct.get("du_lieu") or []
        if raw and (not ct.get("X") or len(ct["X"]) != len(raw)):
            ct["X"] = [
                math.fsum(p) / len(p)
                if (p and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in p))
                else None
                for p in raw
            ]
            ct["xbar"] = list(ct["X"])

    tat_ca_vt = [idx for vt in cac_day for idx in vt]
    khong_giao_nhau = len(tat_ca_vt) == len(set(tat_ca_vt))
    bo_ba = _tim_bo_ba_t2_h1_h2(danh_sach_chi_tieu)
    tap_bo_ba = set(bo_ba) if bo_ba else set()

    so_day_sua = 0
    so_ct_hoan_vi = 0
    so_ct_chinh_so = 0

    if khong_giao_nhau:
        for d_idx, vt in enumerate(cac_day):
            if dung and dung():
                raise InterruptedError("Đã dừng sửa dữ liệu theo yêu cầu.")
            if tien_do:
                tien_do(d_idx + 1, len(cac_day))

            ds_sub = cham_day_kich_thuoc(danh_sach_chi_tieu, vt, cpk_toi_thieu)
            loi_truoc = [i for i, c in enumerate(ds_sub) if not c["kq"].dat]
            if not loi_truoc:
                continue

            so_day_sua += 1
            # Buoc 1: Hoan vi thu tu mau trong day vt
            sua_hoan_vi_kich_thuoc(ds_sub, cpk_toi_thieu=cpk_toi_thieu, so_buoc=3500)
            loi_sau_hv = [i for i, c in enumerate(ds_sub) if not c["kq"].dat]
            so_ct_hoan_vi += max(0, len(loi_truoc) - len(loi_sau_hv))

            # Buoc 2: Tinh chinh gia tri trong [Min, Max] cho chi tieu van con loi
            if loi_sau_hv and cho_phep_chinh_gia_tri:
                da_xu_ly_bo_ba = False
                for c_i in loi_sau_hv:
                    if c_i in tap_bo_ba:
                        if da_xu_ly_bo_ba:
                            continue
                        da_xu_ly_bo_ba = True
                        i_t2, i_h1, i_h2 = bo_ba
                        res3 = _sinh_1_nhom_t2_h1_h2(
                            ds_sub[i_t2], ds_sub[i_h1], ds_sub[i_h2],
                            so_mau=len(vt), cpk_toi_thieu=cpk_toi_thieu
                        )
                        for idx_b, (x_m, raw_m, _r_m) in zip(bo_ba, res3):
                            sub_b = ds_sub[idx_b]
                            sub_b["du_lieu"] = raw_m
                            sub_b["X"] = x_m
                            sub_b["xbar"] = list(x_m)
                            sub_b["loi_du_lieu"] = []
                            sub_b["kq"] = sub_b["spc_result"] = cham_chi_tieu(sub_b, cpk_toi_thieu)
                        so_ct_chinh_so += sum(1 for idx_b in bo_ba if idx_b in loi_sau_hv)
                        continue
                    sub_ct = ds_sub[c_i]
                    min_v, max_v = khoang_min_max_an_toan(sub_ct)
                    x_m, raw_m, r_m = _sinh_1_nhom_chi_tieu(
                        min_v, max_v, sub_ct.get("lsl"), sub_ct.get("usl"),
                        so_mau=len(vt), so_diem_do=max(1, int(sub_ct.get("so_diem_do") or 1)),
                        cpk_toi_thieu=cpk_toi_thieu
                    )
                    sub_ct["du_lieu"] = raw_m
                    sub_ct["X"] = x_m
                    sub_ct["xbar"] = list(x_m)
                    sub_ct["loi_du_lieu"] = []
                    sub_ct["kq"] = sub_ct["spc_result"] = cham_chi_tieu(sub_ct, cpk_toi_thieu)
                    so_ct_chinh_so += 1

            # Ghi nguoc du lieu day vt da chuan hoa vao danh_sach_chi_tieu
            for ct_goc, sub_ct in zip(danh_sach_chi_tieu, ds_sub):
                for k, sample_idx in enumerate(vt):
                    j = sample_idx - 1
                    ct_goc["du_lieu"][j] = list(sub_ct["du_lieu"][k])
                    ct_goc["X"][j] = sub_ct["X"][k]
                ct_goc["xbar"] = list(ct_goc["X"])
    else:
        # Truong hop cac day giao nhau (nguoi dung nhap nhieu diem bat dau chong lap)
        tap_idx = sorted(set(tat_ca_vt))
        da_sua_bo_ba_giao = False
        for ct_i, ct in enumerate(danh_sach_chi_tieu):
            if dung and dung():
                raise InterruptedError("Đã dừng sửa dữ liệu theo yêu cầu.")
            if tien_do:
                tien_do(ct_i + 1, len(danh_sach_chi_tieu))

            if ct_i in tap_bo_ba:
                if da_sua_bo_ba_giao:
                    continue
                da_sua_bo_ba_giao = True
                cts_b = [danh_sach_chi_tieu[idx_b] for idx_b in bo_ba]

                def tong_loi_3():
                    tong = 0
                    for vt in cac_day:
                        for c_b in cts_b:
                            tong += len(cham_chi_tieu(cat_day_chi_tieu(c_b, vt), cpk_toi_thieu).loi)
                    return tong

                p3_cur = tong_loi_3()
                if p3_cur == 0:
                    continue
                if len(tap_idx) >= 2:
                    for _ in range(4000):
                        if p3_cur == 0:
                            break
                        a, b = random.sample(tap_idx, 2)
                        ja, jb = a - 1, b - 1
                        for c_b in cts_b:
                            c_b["du_lieu"][ja], c_b["du_lieu"][jb] = c_b["du_lieu"][jb], c_b["du_lieu"][ja]
                            c_b["X"][ja], c_b["X"][jb] = c_b["X"][jb], c_b["X"][ja]
                        p3_new = tong_loi_3()
                        if p3_new <= p3_cur:
                            p3_cur = p3_new
                        else:
                            for c_b in cts_b:
                                c_b["du_lieu"][ja], c_b["du_lieu"][jb] = c_b["du_lieu"][jb], c_b["du_lieu"][ja]
                                c_b["X"][ja], c_b["X"][jb] = c_b["X"][jb], c_b["X"][ja]
                    for c_b in cts_b:
                        c_b["xbar"] = list(c_b["X"])
                    if p3_cur == 0:
                        so_ct_hoan_vi += 1
                if p3_cur > 0 and cho_phep_chinh_gia_tri:
                    for _ in range(25):
                        res3 = _sinh_nhieu_nhom_t2_h1_h2(
                            cts_b[0], cts_b[1], cts_b[2],
                            so_mau=len(tap_idx), cpk_toi_thieu=cpk_toi_thieu
                        )
                        for c_b, (x_m, raw_m, _) in zip(cts_b, res3):
                            for k, idx in enumerate(tap_idx):
                                c_b["du_lieu"][idx - 1] = raw_m[k]
                                c_b["X"][idx - 1] = x_m[k]
                            c_b["xbar"] = list(c_b["X"])
                        p3_cur = tong_loi_3()
                        for _ in range(2000):
                            if p3_cur == 0:
                                break
                            a, b = random.sample(tap_idx, 2)
                            ja, jb = a - 1, b - 1
                            for c_b in cts_b:
                                c_b["du_lieu"][ja], c_b["du_lieu"][jb] = c_b["du_lieu"][jb], c_b["du_lieu"][ja]
                                c_b["X"][ja], c_b["X"][jb] = c_b["X"][jb], c_b["X"][ja]
                            p3_new = tong_loi_3()
                            if p3_new <= p3_cur:
                                p3_cur = p3_new
                            else:
                                for c_b in cts_b:
                                    c_b["du_lieu"][ja], c_b["du_lieu"][jb] = c_b["du_lieu"][jb], c_b["du_lieu"][ja]
                                    c_b["X"][ja], c_b["X"][jb] = c_b["X"][jb], c_b["X"][ja]
                        for c_b in cts_b:
                            c_b["xbar"] = list(c_b["X"])
                        if p3_cur == 0:
                            break
                    so_ct_chinh_so += 1
                    so_day_sua += 1
                continue

            lsl, usl = ct.get("lsl"), ct.get("usl")
            so_diem = max(1, int(ct.get("so_diem_do") or 1))
            min_v, max_v = khoang_min_max_an_toan(ct)

            def tong_loi_ct():
                tong = 0
                for vt in cac_day:
                    sub = cat_day_chi_tieu(ct, vt)
                    tong += len(cham_chi_tieu(sub, cpk_toi_thieu).loi)
                return tong

            p_cur = tong_loi_ct()
            if p_cur == 0:
                continue

            # Sua cac mau bi thieu so do hoac ngoai dung sai truoc
            if cho_phep_chinh_gia_tri:
                da_sua_diem = False
                lo_g = max(min_v, lsl) if lsl is not None else min_v
                hi_g = min(max_v, usl) if usl is not None else max_v
                if lo_g >= hi_g:
                    lo_g, hi_g = min_v, max_v
                tam_g = (lo_g + hi_g) / 2.0
                sg_g = max((hi_g - lo_g) / 6.0, BUOC_DO * 0.5)
                for idx in tap_idx:
                    j = idx - 1
                    p_val = ct["du_lieu"][j] if j < len(ct["du_lieu"]) else None
                    hop_le = (
                        p_val is not None and len(p_val) == so_diem
                        and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in p_val)
                        and not any((lsl is not None and v < lsl - SAI_SO) or (usl is not None and v > usl + SAI_SO) for v in p_val)
                    )
                    if not hop_le:
                        p_moi = _sinh_mau_gauss(tam_g, sg_g, lo_g, hi_g, so_diem)
                        ct["du_lieu"][j] = p_moi
                        ct["X"][j] = math.fsum(p_moi) / len(p_moi)
                        da_sua_diem = True
                if da_sua_diem:
                    ct["xbar"] = list(ct["X"])
                    so_ct_chinh_so += 1
                    p_cur = tong_loi_ct()

            # Hoan vi tren tap_idx
            if p_cur > 0 and len(tap_idx) >= 2:
                for _ in range(4000):
                    if p_cur == 0:
                        break
                    a, b = random.sample(tap_idx, 2)
                    ja, jb = a - 1, b - 1
                    ct["du_lieu"][ja], ct["du_lieu"][jb] = ct["du_lieu"][jb], ct["du_lieu"][ja]
                    ct["X"][ja], ct["X"][jb] = ct["X"][jb], ct["X"][ja]
                    p_new = tong_loi_ct()
                    if p_new <= p_cur:
                        p_cur = p_new
                    else:
                        ct["du_lieu"][ja], ct["du_lieu"][jb] = ct["du_lieu"][jb], ct["du_lieu"][ja]
                        ct["X"][ja], ct["X"][jb] = ct["X"][jb], ct["X"][ja]
                ct["xbar"] = list(ct["X"])
                if p_cur == 0:
                    so_ct_hoan_vi += 1

            # Neu van con vuong Cpk / I1 tren cac cua so chong lap -> sinh lai day chuan nhe
            if p_cur > 0 and cho_phep_chinh_gia_tri:
                for _ in range(30):
                    x_m, raw_m, _ = sinh_du_lieu_1_chi_tieu(
                        min_v, max_v, lsl, usl, so_mau=len(tap_idx),
                        so_diem_do=so_diem, cpk_toi_thieu=cpk_toi_thieu
                    )
                    for k, idx in enumerate(tap_idx):
                        ct["du_lieu"][idx - 1] = raw_m[k]
                        ct["X"][idx - 1] = x_m[k]
                    ct["xbar"] = list(ct["X"])
                    p_cur = tong_loi_ct()
                    for _ in range(2500):
                        if p_cur == 0:
                            break
                        a, b = random.sample(tap_idx, 2)
                        ja, jb = a - 1, b - 1
                        ct["du_lieu"][ja], ct["du_lieu"][jb] = ct["du_lieu"][jb], ct["du_lieu"][ja]
                        ct["X"][ja], ct["X"][jb] = ct["X"][jb], ct["X"][ja]
                        p_new = tong_loi_ct()
                        if p_new <= p_cur:
                            p_cur = p_new
                        else:
                            ct["du_lieu"][ja], ct["du_lieu"][jb] = ct["du_lieu"][jb], ct["du_lieu"][ja]
                            ct["X"][ja], ct["X"][jb] = ct["X"][jb], ct["X"][ja]
                    ct["xbar"] = list(ct["X"])
                    if p_cur == 0:
                        break
                so_ct_chinh_so += 1
                so_day_sua += 1

    # Cap nhat lai loi_du_lieu tren danh_sach_chi_tieu goc
    for ct in danh_sach_chi_tieu:
        raw = ct.get("du_lieu") or []
        so_d = max(1, int(ct.get("so_diem_do") or 1))
        ct["loi_du_lieu"] = [
            "Mẫu %d thiếu/sai điểm đo" % (j + 1)
            for j, p in enumerate(raw)
            if not (p and len(p) == so_d and all(
                isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in p
            ))
        ]

    # Kiem tra tong ket toan bo cac_day sau khi sua
    con_hong = []
    so_day_dat = 0
    for vt in cac_day:
        ds_sub = cham_day_kich_thuoc(danh_sach_chi_tieu, vt, cpk_toi_thieu)
        loi_ct = ["%s(%s)" % (c.get("ky_hieu") or c["stt"], ",".join(c["kq"].loi))
                  for c in ds_sub if not c["kq"].dat]
        if not loi_ct:
            so_day_dat += 1
        else:
            con_hong.append("Mẫu %d–%d: %s" % (vt[0], vt[-1], "; ".join(loi_ct[:5])))

    return {
        "dat_het": len(con_hong) == 0,
        "so_day_dat": so_day_dat,
        "tong_day": len(cac_day),
        "so_day_sua": so_day_sua,
        "so_ct_hoan_vi": so_ct_hoan_vi,
        "so_ct_chinh_so": so_ct_chinh_so,
        "con_hong": con_hong,
    }
