# -*- coding: utf-8 -*-
"""
XUAT KICH THUOC - XUAT BAO CAO SPC KICH THUOC & BAO CAO TONG HOP
================================================================
1. Xuat bao cao SPC chuan JIS.QT.07.BM.01 (mau moi):
   - Sheet 'Du lieu goc': khoi so do (STT bat dau tu 2) + khoi tinh toan x/RM.
   - Moi chi tieu 1 sheet bao cao rieng (ten sheet = ky hieu chi tieu: Dw, d,
     d1, T1., t1...), co day du bieu do I-MR; cac o tick ☐ va bang danh gia
     nang luc chinh (F85:M87) de trong de in ra chep tay, gia tri Ca/Cp/Cpk
     va cap danh gia ghi sang bang phu ben canh (P85:R87, ngoai vung in).
   - Sheet 'Danh gia' o cuoi.
   Sua truc tiep XML trong goi .xlsx de giu nguyen 100% checkbox (ctrlProps),
   vmlDrawing, logo va dinh dang bieu do.
2. Xuat file 'Bao cao tong hop - <San pham>_<Ma>.xlsx' theo dung bieu mau moi.
3. Doc lai ca 2 loai file (JIS.QT.07.BM.01 va Bao cao tong hop) de danh gia SPC
   va sap xep lai thu tu 25 mau khi can.
"""

import os
import re
import math
import copy
import shutil
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

try:
    import bo_cau_hinh as core
except ImportError:
    from . import bo_cau_hinh as core


SO_MAU = 25
COT_MAU_DAU = 13                         # M (trong sheet Du lieu goc)
COT_MAU_CUOI = COT_MAU_DAU + SO_MAU - 1  # AK

TH_COT_MAU_DAU = 7                       # G (trong Bao cao tong hop)
TH_COT_MAU_CUOI = TH_COT_MAU_DAU + SO_MAU - 1  # AE

G_SO_DO_DAU = 11      # hang dau khoi so do trong sheet Du lieu goc


class LoiMau(Exception):
    pass


def _cot(c):
    return get_column_letter(c)


def _quote_sheet(ten_sh):
    """Dat ten sheet trong dau nhay don neu co chu so, dau cham, dau gach...
    theo dung quy tac cong thuc cua Excel (vd: Dw! nhung 't1'!, 'H1'!, 'T1.'!)."""
    s = str(ten_sh).replace("'", "''")
    if re.fullmatch(r"[A-Za-z]+", s):
        return s
    return "'%s'" % s


# ------------------------------------------------------------ doc file goc
def _sheet_paths(z):
    """{ten sheet: duong dan xml trong goi} theo dung thu tu trong workbook."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    dich = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    dich.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    ra = {}
    for ten, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        if rid in dich:
            t = dich[rid].lstrip("/")
            ra[_bo_escape(ten)] = t if t.startswith("xl/") else "xl/" + t
    return ra


def _bo_escape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
            .replace("&apos;", "'").replace("&amp;", "&"))


def _tim_sheet(paths, *tu):
    for ten, p in paths.items():
        if any(t.lower() in ten.lower() for t in tu):
            return ten, p
    raise LoiMau("File mẫu không có sheet %s." % " / ".join(tu))


def _tach_hang(xml):
    """(dau, {so hang: (thuoc tinh, noi dung)}, duoi) cua sheetData."""
    a = xml.index("<sheetData>") + len("<sheetData>")
    b = xml.index("</sheetData>")
    hang = {}
    for m in re.finditer(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', xml[a:b], re.S):
        attr = re.sub(r'\s*spans="[^"]*"', "", m.group(2))
        hang[int(m.group(1))] = (attr, m.group(3) or "")
    return xml[:a], hang, xml[b:]


def _kieu_o(noi_dung):
    """{chu cot: style id} cua mot hang mau."""
    return {c: s for c, s in re.findall(r'<c r="([A-Z]+)\d+"[^>]*? s="(\d+)"', noi_dung)}


# ----------------------------------------------------------- tao o / hang
def _o(ref, s, gia_tri=None, cong_thuc=None, chu_kq=False):
    s_attr = ' s="%s"' % s if s is not None else ""
    if cong_thuc is not None:
        t = ' t="str"' if chu_kq else ""
        if not chu_kq and isinstance(gia_tri, (int, float)) and not isinstance(gia_tri, bool) and math.isfinite(gia_tri):
            v_txt = repr(float(gia_tri)) if not float(gia_tri).is_integer() else str(int(gia_tri))
            return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, s_attr, escape(cong_thuc), v_txt)
        return '<c r="%s"%s%s><f>%s</f></c>' % (ref, s_attr, t, escape(cong_thuc))
    if gia_tri is None or gia_tri == "":
        return '<c r="%s"%s/>' % (ref, s_attr)
    if isinstance(gia_tri, (int, float)) and not isinstance(gia_tri, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, repr(float(gia_tri))
                                             if not float(gia_tri).is_integer()
                                             else int(gia_tri))
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, s_attr, escape(str(gia_tri)))


def _kiem_tra_khop_t2_h1_h2(ds):
    """Neu ds co bo ba chi tieu 't2', 'h1', 'h2' (vd ALKTOP II) voi cung so diem do
    va gia tri do thoa t2 == round(h1 - h2, 1), tra ve (idx_t2, idx_h1, idx_h2)
    de ghi cong thuc Excel '=+<cot><r_h1>-<cot><r_h2>' giong het file mau."""
    if not ds:
        return None
    vi_tri_kh = {}
    for idx, ct in enumerate(ds):
        kh = str(ct.get("ky_hieu") or (ct.get("ky_hieu_diem") or [""])[0] or "").strip()
        if kh in ("t2", "h1", "h2") and kh not in vi_tri_kh:
            vi_tri_kh[kh] = idx
    if len(vi_tri_kh) != 3:
        return None
    i_t2, i_h1, i_h2 = vi_tri_kh["t2"], vi_tri_kh["h1"], vi_tri_kh["h2"]
    c_t2, c_h1, c_h2 = ds[i_t2], ds[i_h1], ds[i_h2]
    n_t2 = max(1, int(c_t2.get("so_diem_do") or 1))
    n_h1 = max(1, int(c_h1.get("so_diem_do") or 1))
    n_h2 = max(1, int(c_h2.get("so_diem_do") or 1))
    if not (n_t2 == n_h1 == n_h2):
        return None
    dl_t2 = c_t2.get("du_lieu") or []
    dl_h1 = c_h1.get("du_lieu") or []
    dl_h2 = c_h2.get("du_lieu") or []
    n_mau = min(len(dl_t2), len(dl_h1), len(dl_h2))
    if n_mau == 0:
        return None
    co_so_do = False
    for j in range(n_mau):
        p_t2, p_h1, p_h2 = dl_t2[j], dl_h1[j], dl_h2[j]
        if not (p_t2 and p_h1 and p_h2 and len(p_t2) == n_t2 and len(p_h1) == n_t2 and len(p_h2) == n_t2):
            return None
        for vt, v1, v2 in zip(p_t2, p_h1, p_h2):
            if not all(isinstance(x, (int, float)) and not isinstance(x, bool) and math.isfinite(x)
                       for x in (vt, v1, v2)):
                return None
            if abs((float(v1) - float(v2)) - float(vt)) > 0.06:
                return None
            co_so_do = True
    return (i_t2, i_h1, i_h2) if co_so_do else None


_RE_HANG_TU_CT = re.compile(r"([+-]?)(?:\$?([A-Z]+)\$?(\d+)|(\d+(?:\.\d+)?))", re.I)


def _gia_tri_so_hoac_cong_thuc(val, lay_o_fn, _depth=0):
    """Tra ve float tu gia tri so hoac cong thuc Excel don gian (vd '=+G43-G45',
    '=M49-M51', '=ROUND(G43-G45,1)') khi doc bang openpyxl(data_only=False)."""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val) if math.isfinite(val) else None
    if not isinstance(val, str) or _depth > 6:
        return None
    s = val.strip()
    if not s:
        return None
    if not s.startswith("="):
        try:
            fv = float(s)
            return fv if math.isfinite(fv) else None
        except ValueError:
            return None
    expr = s[1:].replace(" ", "")
    lam_tron = None
    m_rnd = re.match(r"^ROUND\((.+),(\d+)\)$", expr, re.I)
    if m_rnd:
        expr = m_rnd.group(1)
        lam_tron = int(m_rnd.group(2))
    if not expr:
        return None
    vi_tri = 0
    tong = 0.0
    for m in _RE_HANG_TU_CT.finditer(expr):
        if m.start() != vi_tri:
            return None
        vi_tri = m.end()
        dau = -1.0 if m.group(1) == "-" else 1.0
        if m.group(2):
            c_idx = column_index_from_string(m.group(2).upper())
            r_idx = int(m.group(3))
            gt_con = _gia_tri_so_hoac_cong_thuc(lay_o_fn(r_idx, c_idx), lay_o_fn, _depth + 1)
            if gt_con is None:
                return None
            tong += dau * gt_con
        else:
            tong += dau * float(m.group(4))
    if vi_tri != len(expr) or not math.isfinite(tong):
        return None
    return round(tong, lam_tron if lam_tron is not None else 6)


def _hang(r, attr, cac_o):
    return '<row r="%d"%s>%s</row>' % (r, attr, "".join(cac_o))


def _doi_so_hang(noi_dung, r_moi):
    """Chep mot hang mau sang so hang moi (chi doi tham chieu o)."""
    return re.sub(r'(<c r="[A-Z]+)\d+"', r'\g<1>%d"' % r_moi, noi_dung)


def _dat_o(noi_dung, ref, o_moi):
    """Thay mot o trong noi dung hang (giu kieu o cu neu o moi khong co s)."""
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % ref, noi_dung, re.S)
    if not m:
        col_moi = column_index_from_string(re.match(r"[A-Z]+", ref).group(0))
        for mc in re.finditer(r'<c r="([A-Z]+)\d+"', noi_dung):
            if column_index_from_string(mc.group(1)) > col_moi:
                return noi_dung[:mc.start()] + o_moi + noi_dung[mc.start():]
        return noi_dung + o_moi
    s = re.search(r' s="(\d+)"', m.group(1))
    if s and ' s="' not in o_moi:
        o_moi = o_moi.replace('<c r="%s"' % ref, '<c r="%s" s="%s"' % (ref, s.group(1)), 1)
    return noi_dung[:m.start()] + o_moi + noi_dung[m.end():]


# ------------------------------------------------------- ten tieng Viet
def _ten_trong_file_goc(duong_dan_mau):
    """{stt: ten} va ten san pham in o T4 cua file goc."""
    wb = openpyxl.load_workbook(duong_dan_mau, data_only=True)
    ws = wb.worksheets[0]
    sp = ws.cell(4, 20).value                      # T4
    ten = {}
    for r in range(G_SO_DO_DAU, ws.max_row + 1):
        if ws.cell(r, 12).value == "x":
            break
        for c_stt, c_ten in ((6, 7), (5, 6), (3, 4)):
            stt, t = ws.cell(r, c_stt).value, ws.cell(r, c_ten).value
            if isinstance(stt, (int, float)) and t and not isinstance(t, (int, float)):
                ten[int(stt)] = str(t)
                break
    wb.close()
    return ten, (str(sp).strip() if sp else "")


def _xac_dinh_bo_cuc_so_do(hang, duoi, g_tinh_dau):
    """Nhan dien bo cuc cot cua khoi so do (hang 11..) trong sheet Du lieu goc.

    Tra ve dict:
        cot_stt, cot_ten, cot_ten_cuoi, cot_kh, cot_yc, cot_yc_cuoi, cot_trong_dau
    """
    nd_x0 = hang[g_tinh_dau][1]
    gop_dau = re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', duoi)

    # Kiem tra theo cong thuc D{x0} (=F11, =G11, =D11) va E{x0} (=J11, =K11)
    if "<f>F11</f>" in nd_x0:
        # ALKTOP II hoac ISHIZUE OMEGA: STT=E, Ten=F:I, KH=J, YC=K:L
        return {
            "cot_stt": "E", "cot_ten": "F", "cot_ten_cuoi": "I",
            "cot_kh": "J", "cot_yc": "K", "cot_yc_cuoi": "L",
            "cac_cot_truoc_m": list("CDEFGHIJKL"),
        }
    if "<f>G11</f>" in nd_x0:
        # ISHIZUE (STT=F, Ten=G:J, KH=K, YC=L) hoac SFP (STT=F, Ten=G:I, KH=J, YC=K:L)
        if "<f>J11</f>" in nd_x0 or "G11:I11" in gop_dau:
            return {
                "cot_stt": "F", "cot_ten": "G", "cot_ten_cuoi": "I",
                "cot_kh": "J", "cot_yc": "K", "cot_yc_cuoi": "L",
                "cac_cot_truoc_m": list("CDEFGHIJKL"),
            }
        return {
            "cot_stt": "F", "cot_ten": "G", "cot_ten_cuoi": "J",
            "cot_kh": "K", "cot_yc": "L", "cot_yc_cuoi": "L",
            "cac_cot_truoc_m": list("CDEFGHIJKL"),
        }
    # Mau cu: STT=C, Ten=D:J, KH=K, YC=L
    return {
        "cot_stt": "C", "cot_ten": "D", "cot_ten_cuoi": "J",
        "cot_kh": "K", "cot_yc": "L", "cot_yc_cuoi": "L",
        "cac_cot_truoc_m": list("CDEFGHIJKL"),
    }


# ======================================================== DU LIEU GOC
def _dung_du_lieu_goc(xml, ds, thong_tin, ten_goc):
    dau, hang, duoi = _tach_hang(xml)

    # Tim hang bat dau khoi tinh toan (g_tinh_dau) va ket thuc (g_tinh_cuoi)
    g_tinh_dau = next(
        (r for r, (_, nd) in sorted(hang.items())
         if r > 15 and any(x in nd for x in ("<f>D11</f>", "<f>F11</f>", "<f>G11</f>"))),
        None
    )
    if g_tinh_dau is None:
        raise LoiMau("File mẫu không đúng cấu trúc biểu mẫu JIS.QT.07.BM.01 "
                     "(không tìm thấy khối tính toán x/RM trong Dữ liệu gốc).")

    g_trong = g_tinh_dau - 3
    g_td1 = g_tinh_dau - 2
    g_td2 = g_tinh_dau - 1
    cac_hang_sigma = [r for r, (_, nd) in hang.items() if r >= g_tinh_dau and "1.128" in nd]
    g_tinh_cuoi = (max(cac_hang_sigma) + 1) if cac_hang_sigma else (g_tinh_dau + 1)
    g_hang_cuoi = max(hang)

    bo_cuc = _xac_dinh_bo_cuc_so_do(hang, duoi, g_tinh_dau)
    c_stt = bo_cuc["cot_stt"]
    c_ten, c_ten_cuoi = bo_cuc["cot_ten"], bo_cuc["cot_ten_cuoi"]
    c_kh = bo_cuc["cot_kh"]
    c_yc, c_yc_cuoi = bo_cuc["cot_yc"], bo_cuc["cot_yc_cuoi"]

    k_do_dau = _kieu_o(hang[G_SO_DO_DAU][1])
    k_do_giua = _kieu_o(hang.get(G_SO_DO_DAU + 1, hang[G_SO_DO_DAU])[1])
    k_do_cuoi = _kieu_o(hang.get(G_SO_DO_DAU + 2, hang[G_SO_DO_DAU])[1])
    a_do = hang[G_SO_DO_DAU][0]

    r_x_mau = g_tinh_dau + 2 if (g_tinh_dau + 2) in hang and (g_tinh_dau + 2) < g_tinh_cuoi else g_tinh_dau
    r_rm_mau = r_x_mau + 1
    k_x, k_rm = _kieu_o(hang[r_x_mau][1]), _kieu_o(hang[r_rm_mau][1])
    k_x_cuoi = _kieu_o(hang.get(g_tinh_cuoi - 1, hang[r_x_mau])[1])
    k_rm_cuoi = _kieu_o(hang.get(g_tinh_cuoi, hang[r_rm_mau])[1])
    a_x = hang[g_tinh_dau][0]

    n_mau = min(SO_MAU, max((len(ct.get("du_lieu") or []) for ct in ds), default=SO_MAU))
    n_mau = max(2, n_mau)
    cot_cuoi_thuc = _cot(COT_MAU_DAU + n_mau - 1)
    stt_mau_day = thong_tin.get("stt_mau")

    ra = []
    # ---- hang 1-10: chep nguyen, dien T4/T5/T6 va tieu de STT mau M10..AK10
    for r in range(1, G_SO_DO_DAU):
        if r not in hang:
            continue
        attr, nd = hang[r]
        gt = {4: thong_tin.get("san_pham"), 5: thong_tin.get("kich_thuoc"),
              6: thong_tin.get("vat_lieu")}.get(r)
        if gt is not None:
            nd = _dat_o(nd, "T%d" % r, _o("T%d" % r, None, gt))
        if r == 10 and (stt_mau_day or n_mau < SO_MAU):
            for j in range(SO_MAU):
                c_ref = "%s10" % _cot(COT_MAU_DAU + j)
                if j < n_mau:
                    lbl_stt = stt_mau_day[j] if (stt_mau_day and j < len(stt_mau_day)) else (j + 1)
                    nd = _dat_o(nd, c_ref, _o(c_ref, None, lbl_stt))
                else:
                    nd = _dat_o(nd, c_ref, _o(c_ref, None, None))
        ra.append(_hang(r, attr, [nd]))

    gop = [m for m in re.findall(r'<mergeCell ref="([^"]+)"/>', duoi)
           if int(re.search(r"\d+", m).group()) < G_SO_DO_DAU]

    # ---- KHOI SO DO
    vi_tri = []                      # (dong dau, dong cuoi) cua tung chi tieu
    r_tam = G_SO_DO_DAU
    for ct_tam in ds:
        n_tam = ct_tam["so_diem_do"]
        vi_tri.append((r_tam, r_tam + n_tam - 1))
        r_tam += n_tam
    bo_ba_goc = _kiem_tra_khop_t2_h1_h2(ds)

    r = G_SO_DO_DAU
    for idx_ct, ct in enumerate(ds):
        n = ct["so_diem_do"]
        kh_ct = ct.get("ky_hieu") or (ct.get("ky_hieu_diem") or [""])[0]
        ten = ten_goc.get(ct["stt"], ct["ten"])
        la_t2_ct = bo_ba_goc is not None and idx_ct == bo_ba_goc[0]
        d0_h1 = vi_tri[bo_ba_goc[1]][0] if la_t2_ct else None
        d0_h2 = vi_tri[bo_ba_goc[2]][0] if la_t2_ct else None
        for i in range(n):
            if n == 1 or i == 0:
                k_do = k_do_dau
            elif i == n - 1:
                k_do = k_do_cuoi
            else:
                k_do = k_do_giua

            o = []
            for c in bo_cuc["cac_cot_truoc_m"]:
                if c == c_stt:
                    val = ct["stt"] if i == 0 else None
                elif c == c_ten:
                    val = ten if i == 0 else None
                elif c == c_kh:
                    val = kh_ct if i == 0 else None
                elif c == c_yc:
                    val = ct["yeu_cau"] if i == 0 else None
                else:
                    val = None
                st = k_do.get(c)
                if st is not None or val is not None:
                    o.append(_o("%s%d" % (c, r), st, val))
            for j in range(SO_MAU):
                cot_m = _cot(COT_MAU_DAU + j)
                ref = "%s%d" % (cot_m, r)
                gt = ct["du_lieu"][j][i] if (j < n_mau and j < len(ct["du_lieu"]) and i < len(ct["du_lieu"][j])) else None
                ct_o = ("+%s%d-%s%d" % (cot_m, d0_h1 + i, cot_m, d0_h2 + i)) if (la_t2_ct and gt is not None) else None
                o.append(_o(ref, k_do.get(cot_m, k_do_dau.get("M")), gt, cong_thuc=ct_o))
            ra.append(_hang(r, a_do, o))
            r += 1
        d0, d1 = vi_tri[idx_ct]
        if n > 1:
            gop.append("%s%d:%s%d" % (c_stt, d0, c_stt, d1))
            gop.append("%s%d:%s%d" % (c_ten, d0, c_ten_cuoi, d1))
            gop.append("%s%d:%s%d" % (c_kh, d0, c_kh, d1))
            gop.append("%s%d:%s%d" % (c_yc, d0, c_yc_cuoi, d1))
        else:
            if c_ten != c_ten_cuoi:
                gop.append("%s%d:%s%d" % (c_ten, d0, c_ten_cuoi, d1))
            if c_yc != c_yc_cuoi:
                gop.append("%s%d:%s%d" % (c_yc, d0, c_yc_cuoi, d1))

    # ---- hang trong + 2 hang tieu de (dich so hang)
    lech = r - g_trong
    td1 = g_td1 + lech
    td2 = g_td2 + lech
    for g in (g_trong, g_td1, g_td2):
        if g in hang:
            attr, nd = hang[g]
            nd = _doi_so_hang(nd, g + lech)
        else:
            attr, nd = "", ""
        if g == g_td1:
            nd = _dat_o(nd, "J%d" % (g + lech), _o("J%d" % (g + lech), None, "LSL"))
            nd = _dat_o(nd, "K%d" % (g + lech), _o("K%d" % (g + lech), None, "USL"))
        if g == g_td2:
            nd = _dat_o(nd, "U%d" % (g + lech), _o("U%d" % (g + lech), None, 9))
        ra.append(_hang(g + lech, attr, [nd]))
    gop.append("BA%d:BC%d" % (td2, td2))

    # ---- KHOI TINH TOAN
    x0 = g_tinh_dau + lech
    hang_x_moi = []
    for k, (ct, (d0, d1)) in enumerate(zip(ds, vi_tri)):
        x, m = x0 + 2 * k, x0 + 2 * k + 1
        hang_x_moi.append((x, m))
        cuoi = k == len(ds) - 1
        kx, km = (k_x_cuoi, k_rm_cuoi) if cuoi else (k_x, k_rm)
        hai_phia = ct.get("lsl") is not None and ct.get("usl") is not None
        mot_phia = ct.get("lsl") is not None and ct.get("usl") is None
        mot_phia_tren = ct.get("lsl") is None and ct.get("usl") is not None
        o_j = (
            _o("J%d" % x, kx.get("J"), cong_thuc="G%d+H%d" % (x, x))
            if (ct.get("chuan") is not None and ct.get("ds_duoi") is not None)
            else (_o("J%d" % x, kx.get("J"), ct.get("lsl")) if ct.get("lsl") is not None else _o("J%d" % x, kx.get("J")))
        )
        o_k = (
            _o("K%d" % x, kx.get("K"), cong_thuc="G%d+I%d" % (x, x))
            if (ct.get("chuan") is not None and ct.get("ds_tren") is not None)
            else (_o("K%d" % x, kx.get("K"), ct.get("usl")) if ct.get("usl") is not None else _o("K%d" % x, kx.get("K")))
        )
        o = [_o("C%d" % x, kx.get("C"), ct["stt"]),
             _o("D%d" % x, kx.get("D"), cong_thuc="%s%d" % (c_ten, d0), chu_kq=True),
             _o("E%d" % x, kx.get("E"), cong_thuc="%s%d" % (c_kh, d0), chu_kq=True),
             _o("F%d" % x, kx.get("F"), cong_thuc="%s%d" % (c_yc, d0), chu_kq=True),
             _o("G%d" % x, kx.get("G"), ct.get("chuan")),
             _o("H%d" % x, kx.get("H"), ct.get("ds_duoi")),
             _o("I%d" % x, kx.get("I"), ct.get("ds_tren")),
             o_j,
             o_k,
             _o("L%d" % x, kx.get("L"), "x")]
        for j in range(SO_MAU):
            c = _cot(COT_MAU_DAU + j)
            if j < n_mau:
                o.append(_o("%s%d" % (c, x), kx.get(c),
                            cong_thuc="AVERAGE(%s%d:%s%d)" % (c, d0, c, d1)))
            else:
                o.append(_o("%s%d" % (c, x), kx.get(c)))
        f = {
            "AL": "AVERAGE(M%d:%s%d)" % (x, cot_cuoi_thuc, x),
            "AM": "ROUND(AL%d/1.128,2)" % m,
            "AN": "ROUND(AL%d+3*AM%d,1)" % (x, x),
            "AO": "ROUND(AL%d+2*AM%d,1)" % (x, x),
            "AP": "ROUND(AL%d+1*AM%d,1)" % (x, x),
            "AQ": "AL%d" % x,
            "AR": "ROUND(AL%d-1*AM%d,1)" % (x, x),
            "AS": "ROUND(AL%d-2*AM%d,1)" % (x, x),
            "AT": "ROUND(AL%d-3*AM%d,1)" % (x, x),
            "AU": "ROUND(3.267*AL%d,2)" % m,
        }
        for c, ct_f in f.items():
            o.append(_o("%s%d" % (c, x), kx.get(c), cong_thuc=ct_f))
        o.append(_o("AV%d" % x, kx.get("AV"), 0))
        if hai_phia:
            nang_luc = {
                "AW": "(J%d+K%d)/2" % (x, x), "AX": "K%d-J%d" % (x, x),
                "AY": 'IFERROR((AQ%d-AW%d)/(AX%d/2),"-")' % (x, x, x),
                "AZ": 'IFERROR(AX%d/(6*AM%d),"-")' % (x, x),
                "BA": 'IFERROR((1-ABS(AY%d))*AZ%d,"-")' % (x, x),
                "BB": 'IFERROR((AQ%d-J%d)/(3*AM%d),"-")' % (x, x, x),
                "BC": 'IFERROR((K%d-AQ%d)/(3*AM%d),"-")' % (x, x, x),     # Cpk tren
            }
        elif mot_phia:                          # chi co gioi han duoi
            nang_luc = {"AW": None, "AX": None, "AY": None, "AZ": None,
                        "BA": 'IFERROR((AQ%d-J%d)/(3*AM%d),"-")' % (x, x, x),
                        "BB": 'IFERROR((AQ%d-J%d)/(3*AM%d),"-")' % (x, x, x), "BC": None}
        elif mot_phia_tren:                     # chi co gioi han tren
            nang_luc = {"AW": None, "AX": None, "AY": None, "AZ": None,
                        "BA": 'IFERROR((K%d-AQ%d)/(3*AM%d),"-")' % (x, x, x),
                        "BB": None, "BC": 'IFERROR((K%d-AQ%d)/(3*AM%d),"-")' % (x, x, x)}
        else:                                   # kich thuoc tham khao
            nang_luc = dict.fromkeys(("AW", "AX", "AY", "AZ", "BA", "BB", "BC"))
        for c in ("AW", "AX", "AY", "AZ", "BA", "BB", "BC"):
            o.append(_o("%s%d" % (c, x), kx.get(c), cong_thuc=nang_luc[c])
                     if nang_luc[c] else _o("%s%d" % (c, x), kx.get(c), "-"))
        ra.append(_hang(x, a_x, o))

        o = [_o("%s%d" % (c, m), km.get(c)) for c in "CDEFGHIJK"]
        o.append(_o("L%d" % m, km.get("L"), "RM"))
        o.append(_o("M%d" % m, km.get("M"), "-"))
        for j in range(1, SO_MAU):
            c, cp = _cot(COT_MAU_DAU + j), _cot(COT_MAU_DAU + j - 1)
            if j < n_mau:
                o.append(_o("%s%d" % (c, m), km.get(c),
                            cong_thuc="ABS(%s%d-%s%d)" % (c, x, cp, x)))
            else:
                o.append(_o("%s%d" % (c, m), km.get(c)))
        o.append(_o("AL%d" % m, km.get("AL"), cong_thuc="AVERAGE(M%d:%s%d)" % (m, cot_cuoi_thuc, m)))
        for c in ("BA", "BB", "BC"):
            o.append(_o("%s%d" % (c, m), km.get(c)))
        ra.append(_hang(m, a_x, o))

    het_tinh = x0 + 2 * len(ds) - 1
    # ---- cac hang trong con lai sau khoi tinh toan (chep kieu o)
    lech_duoi = het_tinh - g_tinh_cuoi
    for g in range(g_tinh_cuoi + 1, g_hang_cuoi + 1):
        if g in hang:
            attr, nd = hang[g]
            ra.append(_hang(g + lech_duoi, attr, [_doi_so_hang(nd, g + lech_duoi)]))
    hang_cuoi = max(het_tinh, g_hang_cuoi + lech_duoi)

    # ---- mergeCells
    duoi = re.sub(r"<mergeCells[^>]*>.*?</mergeCells>",
                  '<mergeCells count="%d">%s</mergeCells>' % (
                      len(gop), "".join('<mergeCell ref="%s"/>' % g for g in gop)),
                  duoi, flags=re.S)

    # ---- dinh dang co dieu kien
    dxf = re.findall(r'<cfRule type="cellIs" dxfId="(\d+)"[^>]*operator="(\w+)"', duoi)
    dxf_duoi = next((d for d, op in dxf if op == "lessThan"), "26")
    dxf_tren = next((d for d, op in dxf if op == "greaterThan"), "25")
    cf, uu_tien = [], 1
    for ct, (d0, d1) in zip(ds, vi_tri):
        luat = ""
        if ct.get("lsl") is not None:
            luat += ('<cfRule type="cellIs" dxfId="%s" priority="%d" operator="lessThan">'
                     '<formula>%s</formula></cfRule>' % (dxf_duoi, uu_tien, repr(float(ct["lsl"]))))
            uu_tien += 1
        if ct.get("usl") is not None:
            luat += ('<cfRule type="cellIs" dxfId="%s" priority="%d" operator="greaterThan">'
                     '<formula>%s</formula></cfRule>' % (dxf_tren, uu_tien, repr(float(ct["usl"]))))
            uu_tien += 1
        if luat:
            cf.append('<conditionalFormatting sqref="M%d:AK%d">%s</conditionalFormatting>'
                      % (d0, d1, luat))
    if "<conditionalFormatting" in duoi:
        duoi = re.sub(r"(<conditionalFormatting.*?</conditionalFormatting>)+",
                      "@@CF@@", duoi, count=1, flags=re.S)
        duoi = re.sub(r"<conditionalFormatting.*?</conditionalFormatting>", "", duoi, flags=re.S)
        duoi = duoi.replace("@@CF@@", "".join(cf))

    dau = re.sub(r'<dimension ref="[^"]+"/>', '<dimension ref="C1:BC%d"/>' % hang_cuoi, dau)
    dau = re.sub(r'<sheetView ([^>]*?)topLeftCell="[^"]*"', r'<sheetView \1topLeftCell="A1"', dau)
    dau = re.sub(r'<selection [^>]*/>', '<selection activeCell="A1" sqref="A1"/>', dau)
    return dau + "".join(ra) + duoi, {
        "x1": x0, "rm1": x0 + 1, "td1": td1, "td2": td2,
        "x_cuoi": het_tinh - 1, "hang_x_moi": hang_x_moi,
        "g_td1": g_td1, "g_td2": g_td2, "g_tinh_dau": g_tinh_dau,
        "n_mau": n_mau,
    }


# ============================================================== SHEET CHI TIEU
def _doi_hang_bao_cao(xml, ten_goc_sheet, bang):
    """Doi so hang trong moi tham chieu 'Du lieu goc'!<cot><hang> theo bang
    {hang cu: hang moi}."""
    ten = re.escape(escape(ten_goc_sheet, {"'": "&apos;"}))
    ten2 = re.escape(ten_goc_sheet)

    def doi(m):
        so = int(m.group(3))
        return "%s%s%d" % (m.group(1), m.group(2), bang.get(so, so))

    for t in {ten, ten2}:
        xml = re.sub(r"(" + "(?:'|&apos;)" + t + r"(?:'|&apos;)!\$?[A-Z]{1,3})(\$?)(\d+)",
                     doi, xml)
    return xml


def _ghi_chu_xml(xml, ref, chu):
    """Ghi chuoi inlineStr vao o ref tren sheet XML, giu nguyen style id s="..."."""
    def _thay(m):
        m_s = re.search(r' s="(\d+)"', m.group(1))
        s_attr = (' s="%s"' % m_s.group(1)) if m_s else ""
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            ref, s_attr, escape(str(chu))
        )
    return re.sub(r'<c r="%s"([^>/]*)(?:/>|>.*?</c>)' % ref, _thay, xml, count=1, flags=re.S)


def _xoa_o_xml(xml, ref):
    """Xoa gia tri/cong thuc cua o ref tren sheet XML, giu nguyen style id s="..."."""
    def _thay(m):
        m_s = re.search(r' s="(\d+)"', m.group(1))
        s_attr = (' s="%s"' % m_s.group(1)) if m_s else ""
        return '<c r="%s"%s/>' % (ref, s_attr)
    return re.sub(r'<c r="%s"([^>/]*)(?:/>|>.*?</c>)' % ref, _thay, xml, count=1, flags=re.S)


def _ghi_cong_thuc_xml(xml, ref, cong_thuc):
    """Ghi cong thuc vao o ref tren sheet XML, giu nguyen style id s="..."."""
    def _thay(m):
        m_s = re.search(r' s="(\d+)"', m.group(1))
        s_attr = (' s="%s"' % m_s.group(1)) if m_s else ""
        return '<c r="%s"%s><f>%s</f></c>' % (ref, s_attr, escape(cong_thuc))
    moi, n = re.subn(r'<c r="%s"([^>/]*)(?:/>|>.*?</c>)' % ref, _thay, xml, count=1, flags=re.S)
    if n:
        return moi
    r_num = int(re.search(r"\d+$", ref).group(0))
    return re.sub(
        r'(<row r="%d"[^>]*>)(.*?)(</row>)' % r_num,
        lambda mr: mr.group(1) + _dat_o(mr.group(2), ref, _o(ref, None, cong_thuc=cong_thuc)) + mr.group(3),
        xml, count=1, flags=re.S,
    )


O_TICK_TRONG = (
    "I50", "K50", "I51", "K51",
    "K68", "M68", "K69", "M69", "K70", "M70", "K71", "M71",
    "K72", "M72", "K73", "M73", "K74", "M74", "K75", "M75",
    "B78", "B79", "D89", "H89", "F95", "K95", "F96", "K96",
)

O_THONG_TIN_IN = (
    ("du_lieu_danh_gia", "E8"),
    ("tieu_chuan_ap_dung", "E10"),
    ("nguoi_lap", "D99"),
    ("nguoi_kiem_tra", "G99"),
    ("nguoi_phe_duyet", "J99"),
)


def _chuyen_bang_nang_luc_sang_phai(xml):
    """Muc 6 Danh gia nang luc:
    - Bang chinh trong vung in (F85:M87) de TRONG hoan toan de in ra chep tay.
    - Bang phu ben canh ngoai vung in (O84:W87 -> P85:R87) ghi day du Gia tri, Cap, Danh gia.
    - Tat ca cac o tick ☐ de trong (khong tich ho), E80 & C93 trong vung in de trong.
    """
    o_ca, o_cp, o_cpk = (
        ("BK5", "BL5", "BM5") if "BM5,BN5,BO5" in xml else ("BL5", "BM5", "BN5")
    )

    # 1. Xoa sach bang chinh F85:H87 trong vung in (giu nguyen vien/kieu s="...")
    for r in (85, 86, 87):
        for c in ("F", "G", "H"):
            xml = _xoa_o_xml(xml, "%s%d" % (c, r))

    # 2. Ghi cong thuc Gia tri / Cap / Danh gia sang bang ben canh (P85:R87)
    cong_thuc_ben_canh = (
        ("P85", o_ca),
        ("Q85", 'IF(ISNUMBER(P85),IF(ABS(P85)<=0.125,"A",IF(ABS(P85)<=0.25,"B",IF(ABS(P85)<=0.5,"C","D"))),"-")'),
        ("R85", 'IF(Q85="A","Rất tốt",IF(Q85="B","Tốt",IF(Q85="C","Kém",IF(Q85="D","Không đạt",IF(ISNUMBER(P87),"Không tính (chỉ có giới hạn một phía)","Không có giới hạn kỹ thuật")))))'),
        ("P86", o_cp),
        ("Q86", 'IF(ISNUMBER(P86),IF(P86>=2,"A+",IF(P86>=1.67,"A",IF(P86>=1.33,"B",IF(P86>=1,"C","D")))),"-")'),
        ("R86", 'IF(Q86="A+","Rất tốt",IF(Q86="A","Tốt",IF(Q86="B","Đạt",IF(Q86="C","Cảnh báo",IF(Q86="D","Không đạt",IF(ISNUMBER(P87),"Không tính (chỉ có giới hạn một phía)","Không có giới hạn kỹ thuật"))))))'),
        ("P87", o_cpk),
        ("Q87", 'IF(ISNUMBER(P87),IF(P87>=1.67,"A+",IF(P87>=1.33,"A",IF(P87>=1,"B",IF(P87>=0.67,"C","D")))),"-")'),
        ("R87", 'IF(Q87="A+","Rất tốt",IF(Q87="A","Đạt tốt",IF(Q87="B","Đạt",IF(Q87="C","Không đạt",IF(Q87="D","Không đạt nghiêm trọng","Không có giới hạn kỹ thuật")))))'),
    )
    for ref, cf in cong_thuc_ben_canh:
        xml = _ghi_cong_thuc_xml(xml, ref, cf)

    # 3. Dam bao cac o tick va nhan xet trong vung in luon de trong ☐ (ke ca khi nap lai file cu)
    if "☑" in xml or "☒" in xml:
        for ref in O_TICK_TRONG:
            xml = _ghi_chu_xml(xml, ref, "☐")
    xml = _xoa_o_xml(xml, "E80")
    xml = _xoa_o_xml(xml, "C93")

    # 4. Chuan hoa noi dung quy tac 6 tren bieu do I
    xml = _ghi_chu_xml(
        xml,
        "B73",
        "14 điểm liên tiếp luân phiên tăng và giảm\n連続する14点が交互に増加・減少している。",
    )
    return xml


def _sua_sheet_chi_tieu(xml, ten_goc_sheet, bang_map, thong_tin=None, n_mau=SO_MAU):
    """Tro cac tham chieu toi hang td1/td2/x/RM cua chi tieu, chuyen bang nang luc
    sang bang phu ben phai (de trong bang chinh & cac o tick), dien thong tin in."""
    xml = _doi_hang_bao_cao(xml, ten_goc_sheet, bang_map)
    xml = _chuyen_bang_nang_luc_sang_phai(xml)
    # Dam bao K8 hien thi ten + ky hieu
    xml = re.sub(r'(<c r="K8"[^>]*>)<f[^>]*>Q5</f>',
                 r'\1<f>CONCATENATE(Q5," ","(",R5,")")</f>', xml)

    tt = thong_tin or {}
    ngay = str(tt.get("ngay_danh_gia") or "").strip() or datetime.now().strftime("%d/%m/%Y")
    xml = _ghi_chu_xml(xml, "E7", ngay)
    for khoa, ref in O_THONG_TIN_IN:
        val = str(tt.get(khoa) or "").strip()
        if val:
            xml = _ghi_chu_xml(xml, ref, val)

    # Neu nhom cuoi co it hon 25 mau (2 <= n_mau < 25): xoa cac o thua va thu gon TRANSPOSE
    if 2 <= n_mau < SO_MAU:
        for k in range(n_mau, SO_MAU):
            c_z = _cot(26 + k)  # Z=26 (mau 1) .. AX=50 (mau 25)
            xml = _xoa_o_xml(xml, "%s5" % c_z)
            xml = _xoa_o_xml(xml, "%s6" % c_z)
            d_stt, d_dl = (16, 17) if k < 13 else (19, 20)
            c_tb = _cot(2 + (k if k < 13 else k - 13))
            xml = _xoa_o_xml(xml, "%s%d" % (c_tb, d_stt))
            xml = _xoa_o_xml(xml, "%s%d" % (c_tb, d_dl))
        c_cuoi_z = _cot(26 + n_mau - 1)
        xml = re.sub(
            r'<f t="array" ref="AA10:AA33">TRANSPOSE\(AA6:AX6\)</f>',
            '<f t="array" ref="AA10:AA%d">TRANSPOSE(AA6:%s6)</f>' % (8 + n_mau, c_cuoi_z),
            xml, count=1
        )
        xml = re.sub(
            r'<f t="array" ref="AG10:AG34">TRANSPOSE\(Z5:AX5\)</f>',
            '<f t="array" ref="AG10:AG%d">TRANSPOSE(Z5:%s5)</f>' % (9 + n_mau, c_cuoi_z),
            xml, count=1
        )

    # Bo tabSelected="1" tren sheet con de Excel khong chon nhom nhieu sheet
    xml = re.sub(r'\s*tabSelected="1"', "", xml)
    return xml


def _xoa_gia_tri_cu(xml):
    """Bo gia tri tinh san o moi o cong thuc va trong vung cong thuc mang,
    de Excel tinh lai khi mo file (fullCalcOnLoad)."""
    xml = re.sub(r"(<f[^>]*>[^<]*</f>|<f[^>]*/>)<v>[^<]*</v>", r"\1", xml)
    for ref in re.findall(r'<f t="array" ref="([A-Z]+\d+:[A-Z]+\d+)"', xml):
        a, b = ref.split(":")
        ca, ra_ = re.match(r"([A-Z]+)(\d+)", a).groups()
        cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
        c0, c1 = column_index_from_string(ca), column_index_from_string(cb)
        for r in range(int(ra_), int(rb) + 1):
            for c in range(c0, c1 + 1):
                o = "%s%d" % (_cot(c), r)
                if o == a:
                    continue
                xml = re.sub(r'<c r="%s"((?: [a-z]+="[^"]*")*)>(?!<f)<v>[^<]*</v></c>' % o,
                             r'<c r="%s"\1/>' % o, xml)
    return xml


def _sua_chart_xml(chart_xml, ten_sheet_moi, n_mau=SO_MAU):
    """Cap nhat tham chieu sheet trong <c:f> cua chart XML, xoa numCache cu,
    va thu hep vung du lieu bieu do neu n_mau < 25."""
    q_sh = escape(_quote_sheet(ten_sheet_moi), {"'": "&apos;", '"': "&quot;"})
    chart_xml = re.sub(
        r"(<c:f>)(?:(?:'[^']*'|&apos;.*?&apos;|[^!]+)!)" r"(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)(</c:f>)",
        lambda m: "%s%s!%s%s" % (m.group(1), q_sh, m.group(2), m.group(3)),
        chart_xml
    )
    chart_xml = re.sub(r"<c:numCache>.*?</c:numCache>", "", chart_xml, flags=re.S)
    if 2 <= n_mau < SO_MAU:
        chart_xml = re.sub(
            r"(\$[A-Z]{1,2}\$10:\$[A-Z]{1,2}\$)(33|34)\b",
            lambda m: "%s%d" % (m.group(1), 9 + n_mau - (1 if m.group(2) == "33" else 0)),
            chart_xml
        )
        chart_xml = re.sub(
            r'<c:max val="(24|25)"/>',
            lambda m: '<c:max val="%d"/>' % (n_mau - (1 if m.group(1) == "24" else 0)),
            chart_xml
        )
    return chart_xml


_CACHE_MAU_JIS = {}


def _lay_cache_mau_jis(duong_dan_mau):
    """Doc va cache noi dung zip + metadata cua file mau JIS.QT.07.BM.01 theo (abspath, mtime)
    de xuat hang chuc file SPC (vd 40 file cho 1000 mau) cuc nhanh."""
    p_abs = os.path.abspath(duong_dan_mau)
    mtime = os.path.getmtime(p_abs)
    key = (p_abs, mtime)
    if key in _CACHE_MAU_JIS:
        return _CACHE_MAU_JIS[key]

    with zipfile.ZipFile(p_abs, "r") as z:
        paths = _sheet_paths(z)
        entries = [(info.filename, z.read(info.filename)) for info in z.infolist()]
    by_name = dict(entries)
    ten_dl, p_dl = _tim_sheet(paths, "liệu gốc", "lieu goc")
    ten_dg, p_dg = _tim_sheet(paths, "Đánh giá", "Danh gia")
    cac_sheet_ct = [(t, p) for t, p in paths.items() if t not in (ten_dl, ten_dg)]
    if not cac_sheet_ct:
        raise LoiMau("File mẫu không có sheet báo cáo chỉ tiêu nào.")
    ten_mau_ct, p_mau_ct = cac_sheet_ct[0]
    xml_mau_ct_clean = _xoa_gia_tri_cu(by_name[p_mau_ct].decode("utf-8"))
    ten_goc, sp_goc = _ten_trong_file_goc(p_abs)

    if len(_CACHE_MAU_JIS) >= 8:
        _CACHE_MAU_JIS.clear()
    data = {
        "entries": entries,
        "by_name": by_name,
        "paths": paths,
        "ten_dl": ten_dl,
        "p_dl": p_dl,
        "ten_dg": ten_dg,
        "p_dg": p_dg,
        "cac_sheet_ct": cac_sheet_ct,
        "ten_mau_ct": ten_mau_ct,
        "p_mau_ct": p_mau_ct,
        "xml_mau_ct_clean": xml_mau_ct_clean,
        "ten_goc": ten_goc,
        "sp_goc": sp_goc,
    }
    _CACHE_MAU_JIS[key] = data
    return data


# ============================================================ XUAT BAO CAO JIS
def xuat_bao_cao(ket_qua_sinh, thong_tin, duong_dan_mau, duong_dan_xuat):
    """Xuat file Excel bao cao SPC kich thuoc chuan JIS.QT.07.BM.01 (mau moi:
    moi chi tieu 1 sheet bao cao rieng; hoac tuong thich mau cu neu truyen vao).
    """
    if not ket_qua_sinh:
        raise LoiMau("Không có chỉ tiêu nào để xuất.")

    import kich_thuoc_core as kt

    cm = _lay_cache_mau_jis(duong_dan_mau)
    by_name = cm["by_name"]
    entries = cm["entries"]
    paths = cm["paths"]
    ten_dl, p_dl = cm["ten_dl"], cm["p_dl"]
    ten_dg, p_dg = cm["ten_dg"], cm["p_dg"]
    cac_sheet_ct = cm["cac_sheet_ct"]
    ten_mau_ct, p_mau_ct = cm["ten_mau_ct"], cm["p_mau_ct"]

    ten_goc = dict(cm["ten_goc"])
    sp_goc = cm["sp_goc"]
    sp = str(thong_tin.get("san_pham") or "").strip()
    if not (sp and sp_goc and sp.lower() == sp_goc.lower()
            and len(ten_goc) == len(ket_qua_sinh)):
        ten_goc = {}

    xml_dl, vi_tri = _dung_du_lieu_goc(by_name[p_dl].decode("utf-8"),
                                       ket_qua_sinh, thong_tin, ten_goc)
    n_mau = vi_tri.get("n_mau", SO_MAU)

    xml_mau_ct = cm["xml_mau_ct_clean"]
    # Tim hang x/RM ma sheet mau chi tieu dang tham chieu toi
    m_rx = re.search(r"!\$?M\$?(\d+)", xml_mau_ct)
    g_rx_mau = int(m_rx.group(1)) if m_rx else vi_tri["g_tinh_dau"]

    # Kiem tra xem mau la mau moi (nhieu sheet chi tieu / cot Q5) hay mau cu (1 sheet 'Bao cao' / cot P5)
    la_mau_moi = (len(cac_sheet_ct) > 1) or (ten_mau_ct not in ("Báo cáo", "Bao cao"))

    if not la_mau_moi:
        # Tuong thich nguoc voi mau 1 sheet 'Bao cao' cu
        bang_map = {
            vi_tri["g_td1"]: vi_tri["td1"],
            vi_tri["g_td2"]: vi_tri["td2"],
            g_rx_mau: vi_tri["x1"],
            g_rx_mau + 1: vi_tri["rm1"],
        }
        xml_bc = _sua_sheet_chi_tieu(xml_mau_ct, ten_dl, bang_map,
                                     thong_tin=thong_tin, n_mau=n_mau)

        wb = by_name["xl/workbook.xml"].decode("utf-8")
        wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                    lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                        r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
        idx = list(paths).index(ten_dl)
        wb = re.sub(r'(<definedName name="_xlnm\.Print_Area" localSheetId="%d">[^<]*!)'
                    r'\$[A-Z]+\$\d+:\$[A-Z]+\$\d+' % idx,
                    lambda m: "%s$J$%d:$W$%d" % (m.group(1), vi_tri["td1"], vi_tri["x_cuoi"]), wb)
        rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "",
                      by_name["xl/_rels/workbook.xml.rels"].decode("utf-8"))
        ct_xml = re.sub(r'<Override [^>]*calcChain[^>]*/>', "",
                        by_name["[Content_Types].xml"].decode("utf-8"))
        thay = {p_dl: xml_dl, p_mau_ct: xml_bc, "xl/workbook.xml": wb,
                "xl/_rels/workbook.xml.rels": rels, "[Content_Types].xml": ct_xml}
        tam = duong_dan_xuat + ".tmp"
        with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
            for fn, raw_bytes in entries:
                if fn == "xl/calcChain.xml":
                    continue
                data = thay[fn].encode("utf-8") if fn in thay else raw_bytes
                out.writestr(fn, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tam, duong_dan_xuat)
        return duong_dan_xuat

    # --- MAU MOI: MOI CHI TIEU 1 SHEET BAO CAO RIENG ---
    ds_ten_sh = kt.ten_sheet_chi_tieu(ket_qua_sinh)
    n_ct = len(ket_qua_sinh)

    # Lay cac thanh phan mau tu sheet2 (drawing2, chart1, chart2, style1/2, colors1/2)
    drawing_mau = by_name["xl/drawings/drawing2.xml"].decode("utf-8")
    chart1_mau = by_name["xl/charts/chart1.xml"].decode("utf-8")
    chart2_mau = by_name["xl/charts/chart2.xml"].decode("utf-8")
    style1_mau = by_name.get("xl/charts/style1.xml")
    colors1_mau = by_name.get("xl/charts/colors1.xml")
    style2_mau = by_name.get("xl/charts/style2.xml")
    colors2_mau = by_name.get("xl/charts/colors2.xml")
    xml_dg = re.sub(r'\s*tabSelected="1"', "", by_name[p_dg].decode("utf-8"))

    goi_moi = {}
    # 1. Sheet 1: Du lieu goc
    goi_moi["xl/worksheets/sheet1.xml"] = xml_dl.encode("utf-8")

    # 2. Sheets 2 .. N+1: N chi tieu
    for idx, (ct, ten_sh, (rx, rrm)) in enumerate(
        zip(ket_qua_sinh, ds_ten_sh, vi_tri["hang_x_moi"]), start=1
    ):
        sh_num = idx + 1
        c1_num = 2 * idx - 1
        c2_num = 2 * idx

        bang_map = {
            vi_tri["g_td1"]: vi_tri["td1"],
            vi_tri["g_td2"]: vi_tri["td2"],
            g_rx_mau: rx,
            g_rx_mau + 1: rrm,
        }
        xml_ct = _sua_sheet_chi_tieu(xml_mau_ct, ten_dl, bang_map,
                                     thong_tin=thong_tin, n_mau=n_mau)

        goi_moi["xl/worksheets/sheet%d.xml" % sh_num] = xml_ct.encode("utf-8")
        goi_moi["xl/worksheets/_rels/sheet%d.xml.rels" % sh_num] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%d.xml"/>'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings2.bin"/>'
            '</Relationships>' % sh_num
        ).encode("utf-8")

        goi_moi["xl/drawings/drawing%d.xml" % sh_num] = drawing_mau.encode("utf-8")
        goi_moi["xl/drawings/_rels/drawing%d.xml.rels" % sh_num] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart%d.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart%d.xml"/>'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpeg"/>'
            '</Relationships>' % (c2_num, c1_num)
        ).encode("utf-8")

        goi_moi["xl/charts/chart%d.xml" % c1_num] = _sua_chart_xml(chart1_mau, ten_sh, n_mau=n_mau).encode("utf-8")
        goi_moi["xl/charts/chart%d.xml" % c2_num] = _sua_chart_xml(chart2_mau, ten_sh, n_mau=n_mau).encode("utf-8")

        if style1_mau and colors1_mau:
            goi_moi["xl/charts/style%d.xml" % c1_num] = style1_mau
            goi_moi["xl/charts/colors%d.xml" % c1_num] = colors1_mau
            goi_moi["xl/charts/_rels/chart%d.xml.rels" % c1_num] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors%d.xml"/>'
                '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style%d.xml"/>'
                '</Relationships>' % (c1_num, c1_num)
            ).encode("utf-8")
        if style2_mau and colors2_mau:
            goi_moi["xl/charts/style%d.xml" % c2_num] = style2_mau
            goi_moi["xl/charts/colors%d.xml" % c2_num] = colors2_mau
            goi_moi["xl/charts/_rels/chart%d.xml.rels" % c2_num] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors%d.xml"/>'
                '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style%d.xml"/>'
                '</Relationships>' % (c2_num, c2_num)
            ).encode("utf-8")

    # 3. Sheet N+2: Danh gia
    sh_dg_num = n_ct + 2
    goi_moi["xl/worksheets/sheet%d.xml" % sh_dg_num] = xml_dg.encode("utf-8")

    # 4. Cap nhat xl/workbook.xml
    wb = by_name["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r'activeTab="\d+"', 'activeTab="0"', wb)
    sheets_xml = ['<sheet name="%s" sheetId="1" r:id="rId1"/>' % escape(ten_dl, {'"': "&quot;"})]
    for idx, ten_sh in enumerate(ds_ten_sh, start=1):
        sheets_xml.append(
            '<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
            % (escape(ten_sh, {'"': "&quot;"}), idx + 1, idx + 1)
        )
    sheets_xml.append(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
        % (escape(ten_dg, {'"': "&quot;"}), sh_dg_num, sh_dg_num)
    )
    wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>%s</sheets>" % "".join(sheets_xml), wb, flags=re.S)

    def_names = [
        '<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s!$A$12:$M$112</definedName>'
        % (n_ct + 1, escape(_quote_sheet(ten_dg))),
        '<definedName name="_xlnm.Print_Area" localSheetId="0">%s!$J$%d:$W$%d</definedName>'
        % (escape(_quote_sheet(ten_dl)), vi_tri["td1"], vi_tri["x_cuoi"]),
    ]
    for idx, ten_sh in enumerate(ds_ten_sh, start=1):
        def_names.append(
            '<definedName name="_xlnm.Print_Area" localSheetId="%d">%s!$A$1:$N$101</definedName>'
            % (idx, escape(_quote_sheet(ten_sh)))
        )
    wb = re.sub(r"<definedNames>.*?</definedNames>",
                "<definedNames>%s</definedNames>" % "".join(def_names), wb, flags=re.S)
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                    r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
    goi_moi["xl/workbook.xml"] = wb.encode("utf-8")

    # 5. Cap nhat xl/_rels/workbook.xml.rels
    rels_goc = by_name["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels_khac = re.findall(
        r'<Relationship [^>]*Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/(?!worksheet|calcChain)[^"]+"[^>]*/>',
        rels_goc
    )
    rels_moi = [
        '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
        % (i, i) for i in range(1, sh_dg_num + 1)
    ]
    next_rid = sh_dg_num + 1
    for rk in rels_khac:
        rk_new = re.sub(r'Id="rId\d+"', 'Id="rId%d"' % next_rid, rk)
        rels_moi.append(rk_new)
        next_rid += 1
    goi_moi["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
        % "".join(rels_moi)
    ).encode("utf-8")

    # 6. Cap nhat [Content_Types].xml
    ct_goc = by_name["[Content_Types].xml"].decode("utf-8")
    ct_goc = re.sub(
        r'<Override PartName="/xl/(?:worksheets/sheet\d+|drawings/drawing(?:[2-9]|\d{2,})|charts/(?:chart|style|colors)\d+|calcChain)\.xml"[^>]*/>',
        "", ct_goc
    )
    ct_them = []
    for i in range(1, sh_dg_num + 1):
        ct_them.append(
            '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i
        )
    for idx in range(1, n_ct + 1):
        sh_num = idx + 1
        c1_num, c2_num = 2 * idx - 1, 2 * idx
        ct_them.append(
            '<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % sh_num
        )
        for c_num in (c1_num, c2_num):
            ct_them.append(
                '<Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % c_num
            )
            if style1_mau and colors1_mau:
                ct_them.append(
                    '<Override PartName="/xl/charts/style%d.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>' % c_num
                )
                ct_them.append(
                    '<Override PartName="/xl/charts/colors%d.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>' % c_num
                )
    ct_goc = ct_goc.replace("</Types>", "".join(ct_them) + "</Types>")
    goi_moi["[Content_Types].xml"] = ct_goc.encode("utf-8")

    # 7. Dong goi file .xlsx
    bo_qua_prefix = (
        "xl/worksheets/sheet",
        "xl/worksheets/_rels/sheet",
        "xl/drawings/drawing",
        "xl/drawings/_rels/drawing",
        "xl/charts/",
        "xl/calcChain.xml",
    )
    giu_nguyen = {
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
    }

    tam = duong_dan_xuat + ".tmp"
    with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
        for fn, raw_bytes in entries:
            if fn in goi_moi:
                continue
            if fn not in giu_nguyen and fn.startswith(bo_qua_prefix):
                continue
            if fn.endswith(".vml"):
                s_vml = raw_bytes.decode("utf-8", errors="ignore")
                s_vml = re.sub(r"\s*<[A-Za-z0-9_]*:?Checked>[^<]*</[A-Za-z0-9_]*:?Checked>", "", s_vml)
                s_vml = re.sub(r"\s*<[A-Za-z0-9_]*:?Checked\s*/>", "", s_vml)
                raw_bytes = s_vml.encode("utf-8")
            elif "ctrlProps/" in fn and fn.endswith(".xml"):
                s_cp = raw_bytes.decode("utf-8", errors="ignore")
                raw_bytes = re.sub(r'\s+checked="[^"]*"', "", s_cp).encode("utf-8")
            out.writestr(fn, raw_bytes, compress_type=zipfile.ZIP_DEFLATED)
        for fn, data in goi_moi.items():
            out.writestr(fn, data, compress_type=zipfile.ZIP_DEFLATED)
    os.replace(tam, duong_dan_xuat)
    return duong_dan_xuat


# ============================================== XUAT BAO CAO TONG HOP (MAU MOI)
def tim_mau_tong_hop(san_pham, thu_muc_goi_y=None):
    """Tim file mau 'Bao cao tong hop - <SP>.xlsx' hoac 'Bao cao tong hop.xlsx'."""
    sp = str(san_pham or "").strip()
    cac_tm = []
    if thu_muc_goi_y and os.path.isdir(thu_muc_goi_y):
        cac_tm.append(thu_muc_goi_y)
    try:
        import bo_cau_hinh
        tm_fm = bo_cau_hinh.thu_muc_file_mau()
        if tm_fm and os.path.isdir(tm_fm):
            cac_tm.append(tm_fm)
    except Exception:
        pass
    cac_tm.append(os.path.join(core.base_dir(), "file_mau"))

    for tm in cac_tm:
        if not os.path.isdir(tm):
            continue
        if sp:
            p_rieng = os.path.join(tm, "Báo cáo tổng hợp - %s.xlsx" % sp)
            if os.path.isfile(p_rieng):
                return p_rieng
        p_chung = os.path.join(tm, "Báo cáo tổng hợp.xlsx")
        if os.path.isfile(p_chung):
            return p_chung
        for fn in sorted(os.listdir(tm)):
            if fn.startswith("Báo cáo tổng hợp") and fn.endswith(".xlsx") and not fn.startswith("~$"):
                return os.path.join(tm, fn)
    return None


def xuat_bao_cao_tong_hop(ket_qua_sinh, thong_tin, duong_dan_xuat, duong_dan_mau_th=None):
    """Xuat file 'Bao cao tong hop - <San pham>_<Ma>.xlsx' voi N mau bat ky (vd 25, 100, 1000):
        - Hang 1: A1:C1 = 'Tên sản phẩm', D1:F1 = <san_pham>
        - Hang 2: A2:C2 = 'Kích thước',   D2:F2 = <kich_thuoc>
        - Hang 4: A4 = 'STT\\n番号', B4:E4 = 'Nội dung kiểm tra\\n検査内容',
                  F4 = 'Ký hiệu\\n記号', G4..<cot(6+N)>4 = stt_mau (1..N)
        - Hang 5..: Moi chi tieu chiem n = so_diem_do dong, gop doc A, B:E, F
                    va dien N x n gia tri do tu cot G sang ngang.
    """
    if not ket_qua_sinh:
        raise LoiMau("Không có chỉ tiêu nào để xuất Báo cáo tổng hợp.")

    sp = str(thong_tin.get("san_pham") or "").strip()
    ma = str(thong_tin.get("kich_thuoc") or "").strip()
    n_mau = max(1, max((len(ct.get("du_lieu") or []) for ct in ket_qua_sinh), default=SO_MAU))
    stt_dau = int(thong_tin.get("stt_dau") or 1)
    stt_mau = thong_tin.get("stt_mau")
    if not stt_mau or len(stt_mau) < n_mau:
        stt_mau = [stt_dau + j for j in range(n_mau)]

    if not duong_dan_mau_th or not os.path.isfile(duong_dan_mau_th):
        duong_dan_mau_th = tim_mau_tong_hop(sp)

    if duong_dan_mau_th and os.path.isfile(duong_dan_mau_th):
        wb = openpyxl.load_workbook(duong_dan_mau_th)
        if sp in wb.sheetnames:
            ws = wb[sp]
            for sname in list(wb.sheetnames):
                if sname != sp:
                    wb.remove(wb[sname])
        else:
            ws = wb.worksheets[0]
            if sp:
                ws.title = sp[:31]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = (sp or "TongHop")[:31]
        ws["A1"] = "Tên sản phẩm"
        ws["A2"] = "Kích thước"
        ws["A4"] = "STT\n番号"
        ws["B4"] = "Nội dung kiểm tra\n検査内容"
        ws["F4"] = "Ký hiệu\n記号"
        for j in range(SO_MAU):
            ws.cell(4, TH_COT_MAU_DAU + j, j + 1)

    # Lay kieu tieu de hang 4 (cot dau G=7, cot giua H=8, cot cuoi AE=31)
    c_last_mau = min(ws.max_column, TH_COT_MAU_CUOI)
    hdr_dau = copy.copy(ws.cell(4, TH_COT_MAU_DAU)._style) if ws.max_column >= TH_COT_MAU_DAU else None
    hdr_giua = copy.copy(ws.cell(4, min(TH_COT_MAU_DAU + 1, max(ws.max_column, TH_COT_MAU_DAU)))._style) if hdr_dau is not None else None
    hdr_cuoi = copy.copy(ws.cell(4, max(c_last_mau, TH_COT_MAU_DAU))._style) if hdr_dau is not None else None

    # Lay kieu o mau o hang 5 (dau khoi), hang 6 (giua khoi), hang 7 (cuoi khoi)
    mau_style = {}
    for idx_r, r_mau in enumerate((5, 6, 7)):
        if r_mau <= ws.max_row:
            d_st = {
                c: copy.copy(ws.cell(r_mau, c)._style)
                for c in range(1, min(ws.max_column, TH_COT_MAU_DAU - 1) + 1)
            }
            if ws.max_column >= TH_COT_MAU_DAU:
                d_st["m_dau"] = copy.copy(ws.cell(r_mau, TH_COT_MAU_DAU)._style)
                d_st["m_giua"] = copy.copy(ws.cell(r_mau, min(TH_COT_MAU_DAU + 1, ws.max_column))._style)
                d_st["m_cuoi"] = copy.copy(ws.cell(r_mau, max(c_last_mau, TH_COT_MAU_DAU))._style)
            mau_style[idx_r] = d_st

    # Xoa cac o gop va du lieu cu tu hang 5 tro xuong
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 5:
            ws.merged_cells.remove(mr)
    if ws.max_row >= 5:
        ws.delete_rows(5, ws.max_row - 4)

    # Xoa cot du thua neu n_mau nho hon so cot cu tren file mau
    cot_cuoi_moi = TH_COT_MAU_DAU + n_mau - 1
    if ws.max_column > cot_cuoi_moi:
        ws.delete_cols(cot_cuoi_moi + 1, ws.max_column - cot_cuoi_moi)

    # Dien tieu de STT mau o hang 4 (G4 .. cot_cuoi_moi) va dat do rong cot dong deu
    for j in range(n_mau):
        col_idx = TH_COT_MAU_DAU + j
        h_cell = ws.cell(4, col_idx, stt_mau[j])
        if hdr_dau is not None:
            st_h = hdr_dau if j == 0 else (hdr_cuoi if j == n_mau - 1 else hdr_giua)
            h_cell._style = copy.copy(st_h)
        col_ltr = _cot(col_idx)
        cur_w = ws.column_dimensions[col_ltr].width
        if not cur_w or cur_w < 8.5:
            ws.column_dimensions[col_ltr].width = 8.5

    # Dien thong tin chung hang 1, 2
    ws["D1"] = sp
    ws["D2"] = ma
    existing_merges = {str(m) for m in ws.merged_cells.ranges}
    for ref in ("A1:C1", "D1:F1", "A2:C2", "D2:F2", "B4:E4"):
        if ref not in existing_merges:
            ws.merge_cells(ref)

    # Dien tung chi tieu tu hang 5
    hang_dau_ct = []
    r_tam = 5
    for ct_tam in ket_qua_sinh:
        hang_dau_ct.append(r_tam)
        r_tam += max(1, int(ct_tam.get("so_diem_do") or 1))
    bo_ba_th = _kiem_tra_khop_t2_h1_h2(ket_qua_sinh)

    r = 5
    for idx_ct, ct in enumerate(ket_qua_sinh):
        n = max(1, int(ct.get("so_diem_do") or 1))
        kh_ct = ct.get("ky_hieu") or (ct.get("ky_hieu_diem") or [""])[0]
        ten_ct = ct.get("ten") or ""
        du_lieu = ct.get("du_lieu") or []
        la_t2_th = bo_ba_th is not None and idx_ct == bo_ba_th[0]
        r_h1_th = hang_dau_ct[bo_ba_th[1]] if la_t2_th else None
        r_h2_th = hang_dau_ct[bo_ba_th[2]] if la_t2_th else None

        for i in range(n):
            kieu_hang = mau_style.get(0 if i == 0 else (2 if i == n - 1 else 1)) or mau_style.get(0, {})
            for c in range(1, TH_COT_MAU_DAU):
                cell = ws.cell(r, c)
                if c in kieu_hang:
                    cell._style = copy.copy(kieu_hang[c])
                cell.value = None

            if i == 0:
                ws.cell(r, 1, ct["stt"])
                ws.cell(r, 2, ten_ct)
                ws.cell(r, 6, kh_ct)

            st_m_dau = kieu_hang.get("m_dau")
            st_m_giua = kieu_hang.get("m_giua")
            st_m_cuoi = kieu_hang.get("m_cuoi")
            for j in range(n_mau):
                col_idx_j = TH_COT_MAU_DAU + j
                gt = du_lieu[j][i] if (j < len(du_lieu) and i < len(du_lieu[j])) else None
                if la_t2_th and gt is not None:
                    col_ltr_j = _cot(col_idx_j)
                    gt_ghi = "=+%s%d-%s%d" % (col_ltr_j, r_h1_th + i, col_ltr_j, r_h2_th + i)
                else:
                    gt_ghi = gt
                c_cell = ws.cell(r, col_idx_j, gt_ghi)
                if st_m_dau is not None:
                    st_m = st_m_dau if j == 0 else (st_m_cuoi if j == n_mau - 1 else st_m_giua)
                    c_cell._style = copy.copy(st_m)
                c_cell.number_format = "0.0"

            ws.row_dimensions[r].height = 26.25 if (n == 1 and "\n" in str(ten_ct)) else 17.25
            r += 1

        d0, d1 = r - n, r - 1
        if n > 1:
            ws.merge_cells(start_row=d0, start_column=1, end_row=d1, end_column=1)
            ws.merge_cells(start_row=d0, start_column=2, end_row=d1, end_column=5)
            ws.merge_cells(start_row=d0, start_column=6, end_row=d1, end_column=6)
        else:
            ws.merge_cells(start_row=d0, start_column=2, end_row=d1, end_column=5)

    tam = duong_dan_xuat + ".tmp"
    wb.save(tam)
    wb.close()
    os.replace(tam, duong_dan_xuat)
    return duong_dan_xuat


# ===================================================== DOC & DANH GIA FILE
def _tim_file_dung_sai_mac_dinh(thu_muc_goi_y=None):
    cac_tm = []
    if thu_muc_goi_y and os.path.isdir(thu_muc_goi_y):
        cac_tm.append(thu_muc_goi_y)
    try:
        import bo_cau_hinh
        tm_fm = bo_cau_hinh.thu_muc_file_mau()
        if tm_fm and os.path.isdir(tm_fm):
            cac_tm.append(tm_fm)
    except Exception:
        pass
    cac_tm.append(os.path.join(core.base_dir(), "file_mau"))
    cac_tm.append(core.base_dir())
    for tm in cac_tm:
        p = os.path.join(tm, "Dung sai kích thước.xlsx")
        if os.path.isfile(p):
            return p
    return None


def _tra_dung_sai_cho_tong_hop(san_pham, ma, duong_dan_ds=None):
    """Tra cuu thong tin dung sai {stt: info_dict}, ma_chon, sp_key tu Dung sai kich thuoc.xlsx."""
    import kich_thuoc_core as kt
    if not duong_dan_ds or not os.path.isfile(duong_dan_ds):
        duong_dan_ds = _tim_file_dung_sai_mac_dinh()
    if not duong_dan_ds or not os.path.isfile(duong_dan_ds):
        return {}, "", str(san_pham or "").strip()
    master = kt.doc_dung_sai(duong_dan_ds)
    sp_key = next((k for k in master if k.lower() == str(san_pham or "").strip().lower()), None)
    if not sp_key and master:
        sp_key = next(iter(master))
    if not sp_key:
        return {}, "", str(san_pham or "").strip()
    data = master[sp_key]
    ma_list = data.get("ma_san_pham", [])
    ma_chon = str(ma or "").strip()
    if ma_chon not in ma_list:
        ma_chon = next((m for m in ma_list if ma_chon and ma_chon in m), ma_list[0] if ma_list else "")
    bang = {}
    for ct in data.get("chi_tieu", []):
        info = dict(ct["theo_ma"].get(ma_chon, {}))
        info["ky_hieu_diem"] = ct.get("ky_hieu_diem")
        info["ten_master"] = ct.get("ten")
        info["so_diem_do_master"] = ct.get("so_diem_do")
        bang[ct["stt"]] = info
    return bang, ma_chon, sp_key


def ap_dung_dung_sai(du_lieu_doc, san_pham=None, ma=None, duong_dan_ds=None):
    """Cap nhat lai dung sai (yeu_cau, lsl, usl, min, max) cho du_lieu_doc khi nguoi dung
    doi Loai san pham hoac Ma san pham o Buoc 2."""
    sp = san_pham if san_pham is not None else du_lieu_doc.get("san_pham", "")
    ma_in = ma if ma is not None else du_lieu_doc.get("ma", "")
    bang_ds, ma_tra, sp_key = _tra_dung_sai_cho_tong_hop(sp, ma_in, duong_dan_ds)
    if sp_key:
        du_lieu_doc["san_pham"] = sp_key
    if ma_tra:
        du_lieu_doc["ma"] = ma_tra
    if not bang_ds:
        return du_lieu_doc
    for ct in du_lieu_doc.get("chi_tieu", []):
        info = bang_ds.get(ct["stt"])
        if not info:
            continue
        for k in ("yeu_cau", "chuan", "ds_duoi", "ds_tren", "lsl", "usl", "min", "max"):
            if k in info:
                ct[k] = info.get(k)
        if info.get("ky_hieu_diem"):
            ct["ky_hieu_diem"] = info["ky_hieu_diem"]
        if not ct.get("ten") and info.get("ten_master"):
            ct["ten"] = info["ten_master"]
    return du_lieu_doc


def tim_file_nguon_trong_thu_muc(thu_muc):
    """Tim file BaoCao_TongHop*.xlsx (uu tien) hoac file .xlsx nguon trong thu muc."""
    if not thu_muc or not os.path.isdir(thu_muc):
        raise LoiMau("Thư mục không tồn tại: %s" % thu_muc)
    cac_xlsx = [
        fn for fn in sorted(os.listdir(thu_muc))
        if fn.lower().endswith(".xlsx") and not fn.startswith("~$")
    ]
    if not cac_xlsx:
        raise LoiMau("Trong thư mục không có file .xlsx nào: %s" % thu_muc)
    for fn in cac_xlsx:
        low = fn.lower()
        if low.startswith("baocao_tonghop") or low.startswith("báo cáo tổng hợp"):
            return os.path.join(thu_muc, fn)
    for fn in cac_xlsx:
        if not fn.startswith("BaoCao_SPC_KichThuoc_"):
            return os.path.join(thu_muc, fn)
    return os.path.join(thu_muc, cac_xlsx[0])


def _chuyen_stt_nguyen(val, mac_dinh):
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(round(val))
    if isinstance(val, str):
        s = val.strip()
        if s.lstrip("-").isdigit():
            return int(s)
    return mac_dinh


def doc_bao_cao(duong_dan, duong_dan_dung_sai=None, san_pham_chon=None, ma_chon=None):
    """Doc thu muc hoac file bao cao kich thuoc:
    - Thu muc KichThuoc_... -> tu dong tim file BaoCao_TongHop*.xlsx ben trong.
    - File 'Bao cao tong hop - <SP>.xlsx' / 'BaoCao_TongHop_<SP>_<Ma>.xlsx' (N mau bat ky: 25, 100, 1000...).
    - File bao cao SPC JIS.QT.07.BM.01 (co sheet 'Du lieu goc', <= 25 mau).
    Tra ve {"san_pham", "ma", "vat_lieu", "chi_tieu": [...], "so_mau", "stt_mau",
            "loai_file", "file", "thu_muc"}.
    """
    import math
    import kich_thuoc_core as kt

    if os.path.isdir(duong_dan):
        duong_dan = tim_file_nguon_trong_thu_muc(duong_dan)
    duong_dan = os.path.abspath(duong_dan)
    thu_muc = os.path.dirname(duong_dan)

    wb = openpyxl.load_workbook(duong_dan, data_only=False)
    try:
        ten_dl = next((n for n in wb.sheetnames if "liệu gốc" in n.lower()), None)
        if ten_dl:
            g = wb[ten_dl]
            # Xac dinh cac cot mau thuc te tren M..AK (hang 10 hoac dong so do dau tien)
            cot_mau = [
                c for c in range(COT_MAU_DAU, COT_MAU_CUOI + 1)
                if g.cell(10, c).value is not None or g.cell(G_SO_DO_DAU, c).value is not None
            ]
            if not cot_mau:
                cot_mau = list(range(COT_MAU_DAU, COT_MAU_CUOI + 1))
            stt_mau = [_chuyen_stt_nguyen(g.cell(10, c).value, idx + 1) for idx, c in enumerate(cot_mau)]

            sp_doc = str(san_pham_chon or g["T4"].value or "").strip()
            ma_doc = str(ma_chon or g["T5"].value or "").strip()
            bang_ds, ma_tra, sp_key = _tra_dung_sai_cho_tong_hop(sp_doc, ma_doc, duong_dan_dung_sai)
            if not sp_doc and sp_key:
                sp_doc = sp_key
            if not ma_doc and ma_tra:
                ma_doc = ma_tra

            ds = []
            for r in range(11, g.max_row + 1):
                if g.cell(r, 12).value != "x":
                    continue
                cong_thuc_m = str(g.cell(r, 13).value or "").strip()
                m = re.match(r"^=AVERAGE\(\$?M\$?(\d+)(?::\$?M\$?(\d+))?\)$", cong_thuc_m, re.I)
                if m:
                    a = int(m.group(1))
                    b = int(m.group(2)) if m.group(2) else a
                else:
                    m2 = re.match(r"^=\$?M\$?(\d+)$", cong_thuc_m, re.I)
                    if not m2:
                        continue
                    a = b = int(m2.group(1))

                def _lay_gia_tri_tham_chieu(col_idx, mac_dinh_col):
                    f_str = str(g.cell(r, col_idx).value or "").strip()
                    m_ref = re.match(r"^=\$?([A-Z]+)\$?(\d+)$", f_str, re.I)
                    if m_ref:
                        c_i = column_index_from_string(m_ref.group(1))
                        r_i = int(m_ref.group(2))
                        return g.cell(r_i, c_i).value
                    return g.cell(a, mac_dinh_col).value

                ten_val = _lay_gia_tri_tham_chieu(4, 6)
                kh_val = _lay_gia_tri_tham_chieu(5, 10)
                yc_val = _lay_gia_tri_tham_chieu(6, 11)

                lay_o_goc = lambda r_i, c_i: g.cell(r_i, c_i).value  # noqa: E731
                X, raw, errors = [], [], []
                for idx_c, c in enumerate(cot_mau):
                    v = [_gia_tri_so_hoac_cong_thuc(g.cell(x, c).value, lay_o_goc) for x in range(a, b + 1)]
                    raw.append(v)
                    valid = all(isinstance(t, (int, float)) and not isinstance(t, bool) and math.isfinite(t) for t in v)
                    X.append(math.fsum(v) / len(v) if valid else None)
                    if not valid:
                        errors.append("Mẫu %d thiếu/sai điểm đo" % (idx_c + 1))

                G, H, I = (g.cell(r, c).value for c in (7, 8, 9))
                so = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)  # noqa: E731
                yc_str = str(yc_val or "").strip()
                if not so(G) and yc_str:
                    G, H, I = kt._phan_tich_dung_sai(yc_str)

                lsl = round(G + H, 4) if so(G) and so(H) else None
                usl = round(G + I, 4) if so(G) and so(I) else None
                stt_val = g.cell(r, 3).value
                stt_int = int(stt_val) if so(stt_val) else (len(ds) + 2)
                info = bang_ds.get(stt_int, {})

                ds.append({
                    "stt": stt_int,
                    "ten": str(ten_val or info.get("ten_master") or "").strip(),
                    "ky_hieu": str(kh_val or "").strip(),
                    "ky_hieu_diem": info.get("ky_hieu_diem"),
                    "yeu_cau": yc_str or info.get("yeu_cau", ""),
                    "chuan": G if so(G) else info.get("chuan"),
                    "ds_duoi": H if so(H) else info.get("ds_duoi"),
                    "ds_tren": I if so(I) else info.get("ds_tren"),
                    "lsl": lsl if lsl is not None else info.get("lsl"),
                    "usl": usl if usl is not None else info.get("usl"),
                    "min": info.get("min"),
                    "max": info.get("max"),
                    "X": X,
                    "xbar": list(X),
                    "hang_x": r,
                    "hang_rm": r + 1,
                    "du_lieu": raw,
                    "so_diem_do": b - a + 1,
                    "loi_du_lieu": errors,
                })
            if not ds:
                raise LoiMau("Không đọc được chỉ tiêu nào trong sheet “Dữ liệu gốc”.")
            cac_sh_ct = [n for n in wb.sheetnames if n != ten_dl and "đánh giá" not in n.lower()]
            return {
                "san_pham": sp_doc,
                "ma": ma_doc,
                "vat_lieu": str(g["T6"].value or "SCW480").strip(),
                "chi_tieu": ds,
                "so_mau": len(cot_mau),
                "stt_mau": stt_mau,
                "cac_sheet_chi_tieu": cac_sh_ct,
                "hang_bao_cao": ds[0]["hang_x"],
                "loai_file": "jis",
                "file": duong_dan,
                "thu_muc": thu_muc,
            }

        # --- Kiem tra xem co phai file 'Bao cao tong hop' khong ---
        ws = wb.worksheets[0]
        a1 = str(ws["A1"].value or "").strip().lower()
        a4 = str(ws["A4"].value or "").strip().lower()
        if "tên sản phẩm" in a1 or "stt" in a4:
            sp = str(san_pham_chon or ws["D1"].value or ws.title or "").strip()
            ma = str(ma_chon or ws["D2"].value or "").strip()
            bang_ds, ma_tra, sp_key = _tra_dung_sai_cho_tong_hop(sp, ma, duong_dan_dung_sai)
            if sp_key:
                sp = sp_key
            if ma_tra and (not ma or ma_chon is None):
                ma = ma_tra

            hang_stt = []
            for r in range(5, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if isinstance(v, (int, float)) or (isinstance(v, str) and v.strip().isdigit()):
                    hang_stt.append(r)

            if not hang_stt:
                raise LoiMau("File Báo cáo tổng hợp không có dòng chỉ tiêu nào từ hàng 5.")

            # Xac dinh tat ca cot mau tu cot G (7) sang ngang (ho tro 25, 100, 1000... mau)
            r_dau = hang_stt[0]
            cot_mau = [
                c for c in range(TH_COT_MAU_DAU, ws.max_column + 1)
                if ws.cell(4, c).value is not None or ws.cell(r_dau, c).value is not None
            ]
            if not cot_mau:
                cot_mau = list(range(TH_COT_MAU_DAU, TH_COT_MAU_CUOI + 1))
            stt_mau = [_chuyen_stt_nguyen(ws.cell(4, c).value, idx + 1) for idx, c in enumerate(cot_mau)]

            # Doc nhanh theo hang bang iter_rows(values_only=True) khi so cot lon (vd 1000 cot)
            max_c_doc = max(cot_mau[-1], 6)
            max_r_doc = ws.max_row
            bang_gia_tri = list(ws.iter_rows(
                min_row=1, max_row=max_r_doc, min_col=1, max_col=max_c_doc, values_only=True
            ))

            def _cell_val(r_idx, c_idx):
                ri = r_idx - 1
                if 0 <= ri < len(bang_gia_tri) and 0 <= c_idx - 1 < len(bang_gia_tri[ri]):
                    return bang_gia_tri[ri][c_idx - 1]
                return None

            ds = []
            for idx, r0 in enumerate(hang_stt):
                r_ke = hang_stt[idx + 1] if idx + 1 < len(hang_stt) else None
                n = kt._so_hang_khoi(ws, r0, r_ke)
                r1 = r0 + n - 1
                stt_int = int(float(_cell_val(r0, 1)))
                ten_val = str(_cell_val(r0, 2) or "").strip()
                kh_val = str(_cell_val(r0, 6) or "").strip()

                X, raw, errors = [], [], []
                for j, c in enumerate(cot_mau):
                    v = [_gia_tri_so_hoac_cong_thuc(_cell_val(x, c), _cell_val) for x in range(r0, r1 + 1)]
                    raw.append(v)
                    valid = all(isinstance(t, (int, float)) and not isinstance(t, bool) and math.isfinite(t) for t in v)
                    X.append(math.fsum(v) / len(v) if valid else None)
                    if not valid:
                        errors.append("Mẫu %d thiếu/sai điểm đo" % (j + 1))

                info = bang_ds.get(stt_int, {})
                ds.append({
                    "stt": stt_int,
                    "ten": ten_val or info.get("ten_master", ""),
                    "ky_hieu": kh_val,
                    "ky_hieu_diem": info.get("ky_hieu_diem"),
                    "yeu_cau": info.get("yeu_cau", ""),
                    "chuan": info.get("chuan"),
                    "ds_duoi": info.get("ds_duoi"),
                    "ds_tren": info.get("ds_tren"),
                    "lsl": info.get("lsl"),
                    "usl": info.get("usl"),
                    "min": info.get("min"),
                    "max": info.get("max"),
                    "X": X,
                    "xbar": list(X),
                    "hang_x": None,
                    "hang_rm": None,
                    "du_lieu": raw,
                    "so_diem_do": n,
                    "loi_du_lieu": errors,
                })

            return {
                "san_pham": sp,
                "ma": ma,
                "vat_lieu": "SCW480",
                "chi_tieu": ds,
                "so_mau": len(cot_mau),
                "stt_mau": stt_mau,
                "cac_sheet_chi_tieu": kt.ten_sheet_chi_tieu(ds),
                "hang_bao_cao": None,
                "loai_file": "tong_hop",
                "file": duong_dan,
                "thu_muc": thu_muc,
            }

        raise LoiMau("File không có sheet “Dữ liệu gốc” và cũng không phải "
                     "mẫu “Báo cáo tổng hợp” kích thước.")
    finally:
        wb.close()


def tim_mau_jis(san_pham, thu_muc_goi_y=None):
    """Tim file mau JIS.QT.07.BM.01 theo dung ten san pham."""
    sp = str(san_pham or "").strip()
    cac_tm = []
    if thu_muc_goi_y and os.path.isdir(thu_muc_goi_y):
        cac_tm.append(thu_muc_goi_y)
    try:
        import bo_cau_hinh
        tm_fm = bo_cau_hinh.thu_muc_file_mau()
        if tm_fm and os.path.isdir(tm_fm):
            cac_tm.append(tm_fm)
    except Exception:
        pass
    cac_tm.append(os.path.join(core.base_dir(), "file_mau"))

    for tm in cac_tm:
        if not os.path.isdir(tm):
            continue
        if sp:
            for fn in sorted(os.listdir(tm)):
                if fn.startswith("JIS.QT.07") and fn.endswith(".xlsx") and not fn.startswith("~$"):
                    if f"- {sp}.xlsx".lower() in fn.lower():
                        return os.path.join(tm, fn)
        for fn in sorted(os.listdir(tm)):
            if fn.startswith("JIS.QT.07") and fn.endswith(".xlsx") and not fn.startswith("~$"):
                return os.path.join(tm, fn)
    return None


def nhan_day(stt_mau, vt, ngan=False):
    """Nhan hien thi cho mot day mau (vd 'Mẫu STT 1–25' hoac '1–25')."""
    if not vt:
        return ""
    st = [stt_mau[v - 1] if (stt_mau and 0 <= v - 1 < len(stt_mau)) else v for v in vt]
    if len(vt) == 1:
        return ("%s" if ngan else "Mẫu STT %s") % st[0]
    buoc = vt[1] - vt[0]
    lien_tiep = all(b - a == buoc for a, b in zip(vt, vt[1:])) and all(
        isinstance(x, int) for x in st
    ) and all(b - a == buoc for a, b in zip(st, st[1:]))
    if lien_tiep:
        if buoc == 1:
            return ("%d–%d" if ngan else "Mẫu STT %d–%d") % (st[0], st[-1])
        return ("%d..%d (b%d)" if ngan else "Mẫu STT %d..%d (bước %d)") % (st[0], st[-1], buoc)
    return ("cột %d–%d" if ngan else "Mẫu cột %d–%d") % (vt[0], vt[-1])


def ten_file_spc_day(san_pham, ma, idx_day, tong_day, stt_mau, vt, buoc=1):
    """Dat ten file bao cao SPC cho tung day (vd _01_mau1-25.xlsx .. _40_mau976-1000.xlsx)."""
    sp_clean = re.sub(r'[\\/:*?"<>|\s]+', "_", str(san_pham or "SPC").strip())
    ma_clean = re.sub(r'[\\/:*?"<>|\s]+', "_", str(ma or "SPC").strip())
    a = stt_mau[vt[0] - 1] if (stt_mau and 0 <= vt[0] - 1 < len(stt_mau)) else vt[0]
    b = stt_mau[vt[-1] - 1] if (stt_mau and 0 <= vt[-1] - 1 < len(stt_mau)) else vt[-1]
    if buoc == 1:
        w = max(2, len(str(max(1, tong_day))))
        return "BaoCao_SPC_KichThuoc_%s_%s_%0*d_mau%s-%s.xlsx" % (
            sp_clean, ma_clean, w, idx_day, a, b
        )
    return "BaoCao_SPC_KichThuoc_%s_%s_bd%d_buoc%d_%dmau.xlsx" % (
        sp_clean, ma_clean, vt[0], buoc, len(vt)
    )


def luu_ban_goc_truoc_khi_sua(duong_dan):
    """Sao luu file goc vao thu muc con _truoc_khi_sua/ truoc khi ghi de."""
    if not duong_dan or not os.path.isfile(duong_dan):
        return None
    tm_goc = os.path.join(os.path.dirname(os.path.abspath(duong_dan)), "_truoc_khi_sua")
    os.makedirs(tm_goc, exist_ok=True)
    dich = os.path.join(tm_goc, os.path.basename(duong_dan))
    if not os.path.exists(dich):
        shutil.copy2(duong_dan, dich)
    return dich


def ghi_lai_bao_cao_tong_hop(du_lieu_doc, luu_goc=True):
    """Sao luu ban goc vao <thu_muc>/_truoc_khi_sua/<thoi_diem>/, chuyen cac file SPC cu
    vao do, va ghi de du lieu da sua dat chuan SPC truc tiep vao file Bao cao tong hop
    sinh o Buoc 1.
    Tra ve {"tong_hop", "sao_luu", "spc_cu": [ten_file]}."""
    file_nguon = du_lieu_doc.get("file")
    loai_file = du_lieu_doc.get("loai_file", "tong_hop")
    sp = du_lieu_doc.get("san_pham") or "ISHIZUE"
    ma = du_lieu_doc.get("ma") or ""
    vat_lieu = du_lieu_doc.get("vat_lieu") or "SCW480"
    ds = du_lieu_doc["chi_tieu"]
    stt_mau = du_lieu_doc.get("stt_mau")
    thu_muc = du_lieu_doc.get("thu_muc") or (
        os.path.dirname(os.path.abspath(file_nguon)) if file_nguon else tao_thu_muc_xuat(san_pham=sp, ma=ma)
    )
    sp_clean = re.sub(r'[\\/:*?"<>|\s]+', "_", sp.strip())
    ma_clean = re.sub(r'[\\/:*?"<>|\s]+', "_", ma.strip()) if ma else "SPC"

    if loai_file == "tong_hop" and file_nguon:
        file_th = file_nguon
    else:
        file_th = os.path.join(thu_muc, "BaoCao_TongHop_%s_%s.xlsx" % (sp_clean, ma_clean))

    bk = None
    spc_cu = []
    if luu_goc and os.path.isdir(thu_muc):
        bk = os.path.join(thu_muc, "_truoc_khi_sua", datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f"))
        os.makedirs(bk, exist_ok=True)
        if file_nguon and os.path.isfile(file_nguon):
            shutil.copy2(file_nguon, os.path.join(bk, os.path.basename(file_nguon)))
        if file_th != file_nguon and os.path.isfile(file_th):
            shutil.copy2(file_th, os.path.join(bk, os.path.basename(file_th)))
        for fn in sorted(os.listdir(thu_muc)):
            p_f = os.path.join(thu_muc, fn)
            if (fn.startswith("BaoCao_SPC_KichThuoc_") and fn.lower().endswith(".xlsx")
                    and os.path.isfile(p_f) and os.path.abspath(p_f) != os.path.abspath(file_nguon or "")):
                shutil.move(p_f, os.path.join(bk, fn))
                spc_cu.append(fn)
        with open(os.path.join(bk, "GHI_CHU.txt"), "w", encoding="utf-8") as f:
            f.write(
                "Bản gốc trước khi sửa dữ liệu kích thước cho đạt chuẩn SPC (%s).\n"
                "Dữ liệu mới đã được cập nhật trực tiếp vào: %s\n"
                % (os.path.basename(file_nguon or file_th), os.path.basename(file_th))
            )

    thong_tin_th = {
        "san_pham": sp,
        "kich_thuoc": ma,
        "vat_lieu": vat_lieu,
        "stt_mau": stt_mau,
    }
    xuat_bao_cao_tong_hop(ds, thong_tin_th, file_th)

    # Neu dau vao la 1 file JIS le (<= 25 mau) thi cap nhat luon file JIS do de dong bo
    if loai_file == "jis" and file_nguon and os.path.isfile(file_nguon) and len(ds[0].get("du_lieu") or []) <= SO_MAU:
        xuat_bao_cao(ds, thong_tin_th, file_nguon, file_nguon)

    du_lieu_doc["file_tong_hop"] = file_th
    if loai_file == "tong_hop":
        du_lieu_doc["file"] = file_th
    return {"tong_hop": file_th, "sao_luu": bk, "spc_cu": spc_cu}


def xuat_nhieu_bao_cao_spc(du_lieu_doc, cac_day, file_mau_jis=None, cpk_toi_thieu=1.33,
                           thong_tin_in=None, thu_muc_xuat=None, buoc=1,
                           dung=None, tien_do=None):
    """Xuat hang loat file bao cao danh gia on dinh JIS.QT.07.BM.01 cho danh sach cac_day
    (vd 1000 mau buoc nhay 1 -> 40 file: _01_mau1-25.xlsx .. _40_mau976-1000.xlsx)."""
    import kich_thuoc_core as kt

    if not cac_day:
        raise LoiMau("Chưa có dãy mẫu nào để xuất báo cáo SPC.")

    sp = du_lieu_doc.get("san_pham") or "ISHIZUE"
    ma = du_lieu_doc.get("ma") or ""
    vat_lieu = str((thong_tin_in or {}).get("vat_lieu") or du_lieu_doc.get("vat_lieu") or "SCW480").strip()
    stt_mau_all = du_lieu_doc.get("stt_mau") or list(range(1, (du_lieu_doc.get("so_mau") or SO_MAU) + 1))

    if not thu_muc_xuat:
        thu_muc_xuat = du_lieu_doc.get("thu_muc") or (
            os.path.dirname(os.path.abspath(du_lieu_doc["file"]))
            if du_lieu_doc.get("file") else tao_thu_muc_xuat(san_pham=sp, ma=ma)
        )
    os.makedirs(thu_muc_xuat, exist_ok=True)

    if not file_mau_jis or not os.path.isfile(file_mau_jis):
        if du_lieu_doc.get("loai_file") == "jis" and du_lieu_doc.get("file") and os.path.isfile(du_lieu_doc["file"]):
            file_mau_jis = du_lieu_doc["file"]
        else:
            file_mau_jis = tim_mau_jis(sp, thu_muc_xuat)
    if not file_mau_jis or not os.path.isfile(file_mau_jis):
        raise LoiMau("Không tìm thấy file mẫu JIS.QT.07.BM.01 cho sản phẩm '%s' trong file_mau/." % sp)

    tt_goc = dict(thong_tin_in or {})
    dl_danh_gia_user = str(tt_goc.get("du_lieu_danh_gia") or "").strip()
    tong = len(cac_day)
    ra_files = []

    for idx, vt in enumerate(cac_day, start=1):
        if dung and dung():
            raise InterruptedError("Đã dừng xuất báo cáo SPC theo yêu cầu.")

        ds_day = kt.cham_day_kich_thuoc(du_lieu_doc["chi_tieu"], vt, cpk_toi_thieu)
        thieu = [c.get("ky_hieu") or ("STT %s" % c["stt"]) for c in ds_day if "DU_LIEU" in c["kq"].loi]
        if thieu:
            raise LoiMau(
                "%s có ô trống hoặc sai điểm đo ở: %s. Hãy bấm 'Sửa dữ liệu cho đạt' trước khi xuất."
                % (nhan_day(stt_mau_all, vt), ", ".join(thieu[:8]))
            )

        stt_day = [stt_mau_all[v - 1] if 0 <= v - 1 < len(stt_mau_all) else v for v in vt]
        tt_day = dict(tt_goc)
        tt_day.update({
            "san_pham": sp,
            "kich_thuoc": ma,
            "vat_lieu": vat_lieu,
            "stt_mau": stt_day,
            "du_lieu_danh_gia": dl_danh_gia_user or nhan_day(stt_mau_all, vt),
        })

        ten_f = ten_file_spc_day(sp, ma, idx, tong, stt_mau_all, vt, buoc=buoc)
        duong_dan_f = os.path.join(thu_muc_xuat, ten_f)
        xuat_bao_cao(ds_day, tt_day, file_mau_jis, duong_dan_f)
        ra_files.append(duong_dan_f)

        if tien_do:
            tien_do(idx, tong, ten_f)

    return ra_files


def cap_nhat_va_xuat_bao_cao(du_lieu_doc, duong_dan_xuat_jis=None,
                             duong_dan_xuat_th=None, cpk_toi_thieu=1.33,
                             luu_goc=True):
    """Ghi cap nhat (hoac xuat moi) ca file JIS.QT.07.BM.01 va file Bao cao tong hop
    (ham giu tuong thich cho truong hop 1 nhom 25 mau)."""
    import kich_thuoc_core as kt

    file_nguon = du_lieu_doc.get("file")
    loai_file = du_lieu_doc.get("loai_file", "jis")
    sp = du_lieu_doc.get("san_pham") or "ISHIZUE"
    ma = du_lieu_doc.get("ma") or ""
    vat_lieu = du_lieu_doc.get("vat_lieu") or "SCW480"
    ds = du_lieu_doc["chi_tieu"]

    so_mau = len(ds[0].get("du_lieu") or []) if ds else SO_MAU
    if so_mau > SO_MAU and not duong_dan_xuat_jis:
        if luu_goc:
            ghi_lai_bao_cao_tong_hop(du_lieu_doc, luu_goc=True)
        cac_day = [kt.lay_day(so_mau, s, 1, SO_MAU) for s in kt.mac_dinh_bat_dau(so_mau)]
        return xuat_nhieu_bao_cao_spc(du_lieu_doc, cac_day, cpk_toi_thieu=cpk_toi_thieu)

    for ct in ds:
        ct["cpk_toi_thieu"] = cpk_toi_thieu
        ct["spc_result"] = ct["kq"] = kt.cham_chi_tieu(ct, cpk_toi_thieu)

    thong_tin = {
        "san_pham": sp,
        "kich_thuoc": ma,
        "vat_lieu": vat_lieu,
        "stt_mau": du_lieu_doc.get("stt_mau"),
    }
    thu_muc = os.path.dirname(os.path.abspath(file_nguon)) if file_nguon else tao_thu_muc_xuat(san_pham=sp, ma=ma)
    sp_clean = sp.replace(" ", "_")
    ma_clean = ma.replace(" ", "_") if ma else "SPC"

    if loai_file == "jis":
        mau_jis = file_nguon
        if not duong_dan_xuat_jis:
            duong_dan_xuat_jis = file_nguon
        if not duong_dan_xuat_th:
            cand_th = os.path.join(thu_muc, "BaoCao_TongHop_%s_%s.xlsx" % (sp_clean, ma_clean))
            if os.path.exists(cand_th):
                duong_dan_xuat_th = cand_th
    else:
        mau_jis = tim_mau_jis(sp, thu_muc)
        if not mau_jis:
            raise LoiMau("Không tìm thấy file mẫu JIS.QT.07.BM.01 trong file_mau/ để xuất báo cáo SPC.")
        if not duong_dan_xuat_th:
            duong_dan_xuat_th = file_nguon
        if not duong_dan_xuat_jis:
            duong_dan_xuat_jis = os.path.join(
                thu_muc, "BaoCao_SPC_KichThuoc_%s_%s.xlsx" % (sp_clean, ma_clean)
            )

    if luu_goc:
        if duong_dan_xuat_jis and os.path.isfile(duong_dan_xuat_jis):
            luu_ban_goc_truoc_khi_sua(duong_dan_xuat_jis)
        if duong_dan_xuat_th and os.path.isfile(duong_dan_xuat_th):
            luu_ban_goc_truoc_khi_sua(duong_dan_xuat_th)

    ra_files = []
    if duong_dan_xuat_jis:
        xuat_bao_cao(ds, thong_tin, mau_jis, duong_dan_xuat_jis)
        ra_files.append(duong_dan_xuat_jis)
    if duong_dan_xuat_th:
        xuat_bao_cao_tong_hop(ds, thong_tin, duong_dan_xuat_th)
        ra_files.append(duong_dan_xuat_th)
    return ra_files


def xuat_bao_cao_chi_tieu(duong_dan, ct, hang_hien_tai, duong_dan_xuat, cpk_toi_thieu=1.33):
    """Ham giu lai de tuong thich nguoc khi file dau vao la mau 1 sheet 'Bao cao' cu."""
    z = zipfile.ZipFile(duong_dan)
    paths = _sheet_paths(z)
    ten_dl, _p = _tim_sheet(paths, "liệu gốc", "lieu goc")
    try:
        _t, p_bc = _tim_sheet(paths, "Báo cáo", "Bao cao")
    except LoiMau:
        z.close()
        d = doc_bao_cao(duong_dan)
        cap_nhat_va_xuat_bao_cao(d, duong_dan_xuat_jis=duong_dan_xuat,
                                 cpk_toi_thieu=cpk_toi_thieu, luu_goc=False)
        return duong_dan_xuat

    xml = _doi_hang_bao_cao(z.read(p_bc).decode("utf-8"), ten_dl,
                            {hang_hien_tai: ct["hang_x"],
                             hang_hien_tai + 1: ct["hang_rm"]})
    xml = _chuyen_bang_nang_luc_sang_phai(xml)
    xml = _xoa_gia_tri_cu(xml)
    import kich_thuoc_core as kt
    r = kt.cham_chi_tieu(ct, cpk_toi_thieu)
    if "DU_LIEU" in r.loi:
        z.close()
        raise LoiMau("Chỉ tiêu thiếu/sai dữ liệu: bổ sung đủ điểm đo trước khi xuất.")
    wb = z.read("xl/workbook.xml").decode("utf-8")
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                    r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
    rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "",
                  z.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
    ctt = re.sub(r'<Override [^>]*calcChain[^>]*/>', "",
                 z.read("[Content_Types].xml").decode("utf-8"))
    thay = {p_bc: xml, "xl/workbook.xml": wb, "xl/_rels/workbook.xml.rels": rels,
            "[Content_Types].xml": ctt}
    tam = duong_dan_xuat + ".tmp"
    with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            out.writestr(info, thay[info.filename].encode("utf-8")
                         if info.filename in thay else z.read(info.filename),
                         compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tam, duong_dan_xuat)
    return duong_dan_xuat


def tao_thu_muc_xuat(thu_muc_goc=None, san_pham=None, ma=None, so_mau=None):
    """Tao thu muc output voi ten co san pham, ma, ngay gio va so mau."""
    if thu_muc_goc is None:
        thu_muc_goc = os.path.join(core.base_dir(), "output")
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if san_pham:
        sp_c = re.sub(r'[\\/:*?"<>|\s]+', "_", str(san_pham).strip())
        ma_c = re.sub(r'[\\/:*?"<>|\s]+', "_", str(ma).strip()) if ma else ""
        duoi_mau = ("_%dmau" % int(so_mau)) if so_mau else ""
        ten_tm = "KichThuoc_%s%s_%s%s" % (
            sp_c, ("_" + ma_c) if ma_c else "", now, duoi_mau
        )
    else:
        ten_tm = "KichThuoc_SPC_%s" % now
    duong_dan = os.path.join(thu_muc_goc, ten_tm)
    os.makedirs(duong_dan, exist_ok=True)
    return duong_dan
