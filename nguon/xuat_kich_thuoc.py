# -*- coding: utf-8 -*-
"""
XUAT KICH THUOC - XUAT BAO CAO SPC KICH THUOC RA FILE EXCEL
=============================================================
Dung lai sheet "Du lieu goc" cua bieu mau JIS.QT.07.BM.01 theo DUNG san
pham dang chon, giu nguyen form cua file goc.

Vi sao sua thang XML ma khong dung openpyxl
-------------------------------------------
openpyxl doc-roi-ghi lai file mau se LAM MAT 4 checkbox (ctrlProps), kieu
dang bieu do (chart style/colors) va vmlDrawing. Nen o day mo file .xlsx
nhu mot file nen, chi viet lai:
    xl/worksheets/<Du lieu goc>.xml   dung lai toan bo khoi so do + khoi tinh
    xl/worksheets/<Bao cao>.xml       tro lai dung hang cua chi tieu 1
    xl/workbook.xml                   vung in + bat Excel tinh lai khi mo
moi phan con lai (logo, bieu do, checkbox, sheet Danh gia) chep nguyen.

Bo cuc sheet "Du lieu goc" (y nhu file goc, chi so hang thay doi):
    hang 1-10      tieu de, chep nguyen (dien T4/T5/T6)
    hang 11-...    KHOI SO DO: moi chi tieu so hang = so diem do
    +1 hang trong, +2 hang tieu de (vd 56/57/58 trong file goc)
    ...            KHOI TINH TOAN: moi chi tieu 2 hang x / RM, cong thuc
                   y nhu file goc (vd 59-88 trong file goc)
    ...            cac hang trong con lai, chep nguyen kieu o

Hai loi cong thuc cua file goc duoc SUA (theo yeu cau cua phong):
    BC (Cpk tren)  : file goc = (X-LSL)/3sigma trung voi BB  -> (USL-X)/3sigma
    Bao cao F87    : CHOOSE(RANDBETWEEN(1,3),...) moi lan mo ra mot so
                     -> =BM5 (Cpk = (1-|Ca|)*Cp)
Kem theo ba nhan sai: J57/K57 "USL"/"LSL" bi nguoc, U58 ghi "1" thay "9".
"""

import os
import re
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

try:
    import charpy_core as core
except ImportError:
    from . import charpy_core as core


SO_MAU = 25
COT_MAU_DAU = 13                       # M
COT_MAU_CUOI = COT_MAU_DAU + SO_MAU - 1  # AK

# Vi tri cac hang trong FILE GOC (dung de nhan dang va lay mau kieu o)
G_SO_DO_DAU = 11      # hang dau khoi so do
G_TRONG = 56          # hang trong sau khoi so do
G_TD1, G_TD2 = 57, 58  # hai hang tieu de khoi tinh toan
G_TINH_DAU = 59       # hang x cua chi tieu 1
G_TINH_CUOI = 88      # hang RM cua chi tieu cuoi (15 chi tieu)
G_HANG_CUOI = 106     # hang cuoi cung co kieu o


class LoiMau(Exception):
    pass


def _cot(c):
    return get_column_letter(c)


# ------------------------------------------------------------ doc file goc
def _sheet_paths(z):
    """{ten sheet: duong dan xml trong goi}."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    dich = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    dich.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    ra = {}
    for ten, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = dich[rid].lstrip("/")
        ra[_bo_escape(ten)] = t if t.startswith("xl/") else "xl/" + t
    return ra


def _bo_escape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
            .replace("&apos;", "'").replace("&amp;", "&"))


def _tim_sheet(paths, *tu):
    for ten, p in paths.items():
        if any(t in ten for t in tu):
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
        return '<c r="%s"%s%s><f>%s</f></c>' % (ref, s_attr, t, escape(cong_thuc))
    if gia_tri is None or gia_tri == "":
        return '<c r="%s"%s/>' % (ref, s_attr)
    if isinstance(gia_tri, (int, float)) and not isinstance(gia_tri, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, repr(float(gia_tri))
                                             if not float(gia_tri).is_integer()
                                             else int(gia_tri))
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, s_attr, escape(str(gia_tri)))


def _hang(r, attr, cac_o):
    return '<row r="%d"%s>%s</row>' % (r, attr, "".join(cac_o))


def _doi_so_hang(noi_dung, r_moi):
    """Chep mot hang mau sang so hang moi (chi doi tham chieu o)."""
    return re.sub(r'(<c r="[A-Z]+)\d+"', r'\g<1>%d"' % r_moi, noi_dung)


def _dat_o(noi_dung, ref, o_moi):
    """Thay mot o trong noi dung hang (giu kieu o cu neu o moi khong co s)."""
    m = re.search(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % ref, noi_dung, re.S)
    if not m:
        return noi_dung + o_moi
    s = re.search(r' s="(\d+)"', m.group(1))
    if s and ' s="' not in o_moi:
        o_moi = o_moi.replace('<c r="%s"' % ref, '<c r="%s" s="%s"' % (ref, s.group(1)), 1)
    return noi_dung[:m.start()] + o_moi + noi_dung[m.end():]


# ------------------------------------------------------- ten tieng Viet
def _ten_trong_file_goc(duong_dan_mau):
    """{stt: ten} va ten san pham in o T4 cua file goc (de dung lai ten
    tieng Viet cho dung loai san pham ma file goc lam san)."""
    wb = openpyxl.load_workbook(duong_dan_mau, data_only=True)
    ws = wb.worksheets[0]
    sp = ws.cell(4, 20).value                      # T4
    ten = {}
    for r in range(G_SO_DO_DAU, G_TRONG):
        stt, t = ws.cell(r, 3).value, ws.cell(r, 4).value
        if isinstance(stt, (int, float)) and t:
            ten[int(stt)] = str(t)
    wb.close()
    return ten, (str(sp).strip() if sp else "")


# ======================================================== DU LIEU GOC
def _dung_du_lieu_goc(xml, ds, thong_tin, ten_goc):
    dau, hang, duoi = _tach_hang(xml)

    # kiem tra dung la file goc JIS.QT.07.BM.01
    if G_TINH_DAU not in hang or "D%d" % G_SO_DO_DAU not in hang[G_TINH_DAU][1]:
        raise LoiMau("File mẫu không đúng cấu trúc biểu mẫu JIS.QT.07.BM.01 "
                     "(không thấy công thức D%d = D%d)." % (G_TINH_DAU, G_SO_DO_DAU))

    k_do = _kieu_o(hang[G_SO_DO_DAU][1])
    a_do = hang[G_SO_DO_DAU][0]
    k_x, k_rm = _kieu_o(hang[G_TINH_DAU + 2][1]), _kieu_o(hang[G_TINH_DAU + 3][1])
    k_x_cuoi, k_rm_cuoi = _kieu_o(hang[G_TINH_CUOI - 1][1]), _kieu_o(hang[G_TINH_CUOI][1])
    a_x = hang[G_TINH_DAU][0]

    ra = []
    # ---- hang 1-10: chep nguyen, dien T4/T5/T6
    for r in range(1, G_SO_DO_DAU):
        if r not in hang:
            continue
        attr, nd = hang[r]
        gt = {4: thong_tin.get("san_pham"), 5: thong_tin.get("kich_thuoc"),
              6: thong_tin.get("vat_lieu")}.get(r)
        if gt is not None:
            nd = _dat_o(nd, "T%d" % r, _o("T%d" % r, None, gt))
        ra.append(_hang(r, attr, [nd]))

    gop = [m for m in re.findall(r'<mergeCell ref="([^"]+)"/>', duoi)
           if int(re.search(r"\d+", m).group()) < G_SO_DO_DAU]

    # ---- KHOI SO DO
    r = G_SO_DO_DAU
    vi_tri = []                      # (dong dau, dong cuoi) cua tung chi tieu
    for ct in ds:
        n = ct["so_diem_do"]
        kh = ct.get("ky_hieu_diem") or [ct.get("ky_hieu", "")] * n
        mot_ky_hieu = len(set(kh)) <= 1
        ten = ten_goc.get(ct["stt"], ct["ten"])
        for i in range(n):
            o = []
            o.append(_o("C%d" % r, k_do.get("C"), ct["stt"] if i == 0 else None))
            o.append(_o("D%d" % r, k_do.get("D"), ten if i == 0 else None))
            for c in "EFGHIJ":
                o.append(_o("%s%d" % (c, r), k_do.get(c)))
            o.append(_o("K%d" % r, k_do.get("K"),
                        kh[i] if (i == 0 or not mot_ky_hieu) else None))
            o.append(_o("L%d" % r, k_do.get("L"), ct["yeu_cau"] if i == 0 else None))
            for j in range(SO_MAU):
                ref = "%s%d" % (_cot(COT_MAU_DAU + j), r)
                gt = ct["du_lieu"][j][i] if j < len(ct["du_lieu"]) else None
                o.append(_o(ref, k_do.get(_cot(COT_MAU_DAU + j)), gt))
            ra.append(_hang(r, a_do, o))
            r += 1
        d0, d1 = r - n, r - 1
        vi_tri.append((d0, d1))
        if n > 1:
            gop += ["C%d:C%d" % (d0, d1), "D%d:J%d" % (d0, d1), "L%d:L%d" % (d0, d1)]
            if mot_ky_hieu:
                gop.append("K%d:K%d" % (d0, d1))
        else:
            gop.append("D%d:J%d" % (d0, d1))

    # ---- hang trong + 2 hang tieu de (dich so hang)
    lech = r - G_TRONG
    td1 = G_TD1 + lech
    for g in (G_TRONG, G_TD1, G_TD2):
        attr, nd = hang[g]
        nd = _doi_so_hang(nd, g + lech)
        if g == G_TD1:                         # nhan bi nguoc trong file goc
            nd = _dat_o(nd, "J%d" % (g + lech), _o("J%d" % (g + lech), None, "LSL"))
            nd = _dat_o(nd, "K%d" % (g + lech), _o("K%d" % (g + lech), None, "USL"))
        if g == G_TD2:                         # tieu de mau thu 9 ghi nham "1"
            nd = _dat_o(nd, "U%d" % (g + lech), _o("U%d" % (g + lech), None, 9))
        ra.append(_hang(g + lech, attr, [nd]))
    gop.append("BA%d:BC%d" % (G_TD2 + lech, G_TD2 + lech))

    # ---- KHOI TINH TOAN
    x0 = G_TINH_DAU + lech
    for k, (ct, (d0, d1)) in enumerate(zip(ds, vi_tri)):
        x, m = x0 + 2 * k, x0 + 2 * k + 1
        cuoi = k == len(ds) - 1
        kx, km = (k_x_cuoi, k_rm_cuoi) if cuoi else (k_x, k_rm)
        hai_phia = ct.get("lsl") is not None and ct.get("usl") is not None
        mot_phia = ct.get("lsl") is not None and ct.get("usl") is None
        o = [_o("C%d" % x, kx.get("C"), ct["stt"]),
             _o("D%d" % x, kx.get("D"), cong_thuc="D%d" % d0, chu_kq=True),
             _o("E%d" % x, kx.get("E"), cong_thuc="K%d" % d0, chu_kq=True),
             _o("F%d" % x, kx.get("F"), cong_thuc="L%d" % d0, chu_kq=True),
             _o("G%d" % x, kx.get("G"), ct.get("chuan")),
             _o("H%d" % x, kx.get("H"), ct.get("ds_duoi")),
             _o("I%d" % x, kx.get("I"), ct.get("ds_tren")),
             _o("J%d" % x, kx.get("J"), cong_thuc="G%d+H%d" % (x, x))
             if ct.get("ds_duoi") is not None else _o("J%d" % x, kx.get("J")),
             _o("K%d" % x, kx.get("K"), cong_thuc="G%d+I%d" % (x, x))
             if ct.get("ds_tren") is not None else _o("K%d" % x, kx.get("K")),
             _o("L%d" % x, kx.get("L"), "x")]
        for j in range(SO_MAU):
            c = _cot(COT_MAU_DAU + j)
            o.append(_o("%s%d" % (c, x), kx.get(c),
                        cong_thuc="AVERAGE(%s%d:%s%d)" % (c, d0, c, d1)))
        f = {
            "AL": "AVERAGE(M%d:AK%d)" % (x, x),
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
                "AY": "(AQ%d-AW%d)/(AX%d/2)" % (x, x, x),
                "AZ": "AX%d/(6*AM%d)" % (x, x),
                "BA": "(1-ABS(AY%d))*AZ%d" % (x, x),
                "BB": "(AQ%d-J%d)/(3*AM%d)" % (x, x, x),
                "BC": "(K%d-AQ%d)/(3*AM%d)" % (x, x, x),     # sua: Cpk tren
            }
        elif mot_phia:                          # chi co gioi han duoi
            nang_luc = {"AW": None, "AX": None, "AY": None, "AZ": None,
                        "BA": "(AQ%d-J%d)/(3*AM%d)" % (x, x, x),
                        "BB": "(AQ%d-J%d)/(3*AM%d)" % (x, x, x), "BC": None}
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
            o.append(_o("%s%d" % (c, m), km.get(c),
                        cong_thuc="ABS(%s%d-%s%d)" % (c, x, cp, x)))
        o.append(_o("AL%d" % m, km.get("AL"), cong_thuc="AVERAGE(M%d:AK%d)" % (m, m)))
        for c in ("BA", "BB", "BC"):
            o.append(_o("%s%d" % (c, m), km.get(c)))
        ra.append(_hang(m, a_x, o))

    het_tinh = x0 + 2 * len(ds) - 1
    # ---- cac hang trong con lai sau khoi tinh toan (chep kieu o)
    lech_duoi = het_tinh - G_TINH_CUOI
    for g in range(G_TINH_CUOI + 1, G_HANG_CUOI + 1):
        if g in hang:
            attr, nd = hang[g]
            ra.append(_hang(g + lech_duoi, attr, [_doi_so_hang(nd, g + lech_duoi)]))
    hang_cuoi = G_HANG_CUOI + lech_duoi

    # ---- mergeCells
    duoi = re.sub(r"<mergeCells[^>]*>.*?</mergeCells>",
                  '<mergeCells count="%d">%s</mergeCells>' % (
                      len(gop), "".join('<mergeCell ref="%s"/>' % g for g in gop)),
                  duoi, flags=re.S)

    # ---- dinh dang co dieu kien: to mau so do ngoai dung sai, moi chi tieu
    #      du hai dieu kien (file goc co cho chi co mot, co cho gop hai chi tieu)
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
    duoi = re.sub(r"(<conditionalFormatting.*?</conditionalFormatting>)+",
                  "@@CF@@", duoi, count=1, flags=re.S)
    duoi = re.sub(r"<conditionalFormatting.*?</conditionalFormatting>", "", duoi, flags=re.S)
    duoi = duoi.replace("@@CF@@", "".join(cf))

    dau = re.sub(r'<dimension ref="[^"]+"/>', '<dimension ref="C1:BC%d"/>' % hang_cuoi, dau)
    dau = re.sub(r'<sheetView ([^>]*?)topLeftCell="[^"]*"', r'<sheetView \1topLeftCell="A1"', dau)
    dau = re.sub(r'<selection [^>]*/>', '<selection activeCell="A1" sqref="A1"/>', dau)
    return dau + "".join(ra) + duoi, {"x1": x0, "rm1": x0 + 1, "td1": td1,
                                       "x_cuoi": het_tinh - 1}


# ============================================================== BAO CAO
def _doi_hang_bao_cao(xml, ten_goc_sheet, bang):
    """Doi so hang trong moi tham chieu 'Du lieu goc'!<cot><hang> theo bang
    {hang cu: hang moi} - de sheet Bao cao trinh bay chi tieu khac."""
    ten = re.escape(escape(ten_goc_sheet, {"'": "&apos;"}))
    ten2 = re.escape(ten_goc_sheet)

    def doi(m):
        so = int(m.group(3))
        return "%s%s%d" % (m.group(1), m.group(2), bang.get(so, so))

    for t in {ten, ten2}:
        xml = re.sub(r"(" + "(?:'|&apos;)" + t + r"(?:'|&apos;)!\$?[A-Z]{1,3})(\$?)(\d+)",
                     doi, xml)
    return xml


def _sua_bao_cao(xml, ten_goc_sheet, vi_tri):
    """Tro cac tham chieu toi hang x/RM cua chi tieu 1, sua o Cpk, dien ngay."""
    xml = _doi_hang_bao_cao(xml, ten_goc_sheet, {G_TINH_DAU: vi_tri["x1"],
                                                 G_TINH_DAU + 1: vi_tri["rm1"]})
    # o Cpk: bo RANDBETWEEN (moi lan mo ra mot so) -> Cpk = (1-|Ca|)*Cp
    xml = re.sub(r'<f ca="1">CHOOSE\(RANDBETWEEN\(1,3\),BM5,BN5,BO5\)</f>',
                 "<f>BM5</f>", xml)
    # ngay danh gia
    xml = re.sub(r'<c r="E7"( s="\d+")?(?:/>|>.*?</c>)',
                 lambda m: '<c r="E7"%s t="inlineStr"><is><t>%s</t></is></c>' % (
                     m.group(1) or "", datetime.now().strftime("%d/%m/%Y")),
                 xml, count=1, flags=re.S)
    return xml


def _xoa_gia_tri_cu(xml):
    """Bo gia tri tinh san (cua ISHIZUE trong file goc) o moi o cong thuc va
    trong vung cong thuc mang, de khong ai nhin thay so cu; Excel tinh lai
    khi mo file (fullCalcOnLoad)."""
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


# ============================================================ DONG GOI
def xuat_bao_cao(ket_qua_sinh, thong_tin, duong_dan_mau, duong_dan_xuat):
    """Xuat file Excel bao cao SPC kich thuoc theo dung form file goc.

    ket_qua_sinh: list dict tu kich_thuoc_core.sinh_toan_bo() - moi dict co
        stt, ten, ky_hieu_diem, so_diem_do, yeu_cau, chuan, ds_duoi,
        ds_tren, lsl, usl, du_lieu (25 x so_diem_do)
    thong_tin: {"san_pham", "kich_thuoc", "vat_lieu"}
    """
    if not ket_qua_sinh:
        raise LoiMau("Không có chỉ tiêu nào để xuất.")
    z = zipfile.ZipFile(duong_dan_mau)
    paths = _sheet_paths(z)
    ten_dl, p_dl = _tim_sheet(paths, "liệu gốc", "lieu goc")
    ten_bc, p_bc = _tim_sheet(paths, "Báo cáo", "Bao cao")

    ten_goc, sp_goc = _ten_trong_file_goc(duong_dan_mau)
    sp = str(thong_tin.get("san_pham") or "").strip()
    # Ten tieng Viet cua file goc chi dung cho DUNG loai san pham file goc
    # lam san (va du so chi tieu); loai khac giu nguyen ten trong file dung sai.
    if not (sp and sp_goc and sp.lower() == sp_goc.lower()
            and len(ten_goc) == len(ket_qua_sinh)):
        ten_goc = {}

    xml_dl, vi_tri = _dung_du_lieu_goc(z.read(p_dl).decode("utf-8"),
                                       ket_qua_sinh, thong_tin, ten_goc)
    xml_bc = _sua_bao_cao(z.read(p_bc).decode("utf-8"), ten_dl, vi_tri)
    xml_bc = _xoa_gia_tri_cu(xml_bc)
    import kich_thuoc_core as kt
    import spc_bao_cao
    ct = ket_qua_sinh[0]
    cpk = ct.get("cpk_toi_thieu", kt.CPK_TOI_THIEU)
    xml_bc, _ = spc_bao_cao.dien(xml_bc, kt.cham_chi_tieu(ct, cpk), cpk, mr3=False)

    wb = z.read("xl/workbook.xml").decode("utf-8")
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: "<calcPr%s fullCalcOnLoad=\"1\"/>" % re.sub(
                    r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
    # vung in cua sheet Du lieu goc (file goc: J57:W87) dich theo khoi tinh
    idx = list(paths).index(ten_dl)
    wb = re.sub(r'(<definedName name="_xlnm\.Print_Area" localSheetId="%d">[^<]*!)'
                r'\$[A-Z]+\$\d+:\$[A-Z]+\$\d+' % idx,
                lambda m: "%s$J$%d:$W$%d" % (m.group(1), vi_tri["td1"], vi_tri["x_cuoi"]), wb)

    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rels)
    ct = z.read("[Content_Types].xml").decode("utf-8")
    ct = re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct)

    thay = {p_dl: xml_dl, p_bc: xml_bc, "xl/workbook.xml": wb,
            "xl/_rels/workbook.xml.rels": rels, "[Content_Types].xml": ct}
    tam = duong_dan_xuat + ".tmp"
    with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            data = thay[info.filename].encode("utf-8") if info.filename in thay \
                else z.read(info.filename)
            out.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tam, duong_dan_xuat)
    return duong_dan_xuat


# ===================================================== DANH GIA FILE DA SINH
def doc_bao_cao(duong_dan):
    """Doc lai mot file bao cao kich thuoc (file sinh ra HOAC file goc).

    Lay so lieu tu khoi tinh toan cua sheet 'Du lieu goc': moi hang 'x' cho
    biet chi tieu (STT, G/H/I) va vung hang diem do (qua cong thuc
    =AVERAGE(Mdau:Mcuoi)). Gia tri x cua tung mau tinh lai = trung binh cac
    diem do, dung nhu cong thuc trong file.
    Tra ve {"san_pham", "ma", "chi_tieu": [...], "hang_bao_cao"}.
    """
    wb = openpyxl.load_workbook(duong_dan)
    try:
        ten_dl = next((n for n in wb.sheetnames if "liệu gốc" in n), None)
        if not ten_dl:
            raise LoiMau("File không có sheet “Dữ liệu gốc” — không phải báo "
                         "cáo kích thước JIS.QT.07.BM.01.")
        g = wb[ten_dl]
        ds = []
        for r in range(11, g.max_row + 1):
            if g.cell(r, 12).value != "x":
                continue
            m = re.match(r"=AVERAGE\(M(\d+):M(\d+)\)", str(g.cell(r, 13).value or ""))
            if not m:
                continue
            a, b = int(m.group(1)), int(m.group(2))
            import math
            X, raw, errors = [], [], []
            for c in range(COT_MAU_DAU, COT_MAU_CUOI + 1):
                v = [g.cell(x, c).value for x in range(a, b + 1)]
                raw.append(v)
                valid = all(isinstance(t, (int, float)) and not isinstance(t, bool) and math.isfinite(t) for t in v)
                X.append(math.fsum(v) / len(v) if valid else None)
                if not valid:
                    errors.append("Mẫu %d thiếu/sai điểm đo" % (c - COT_MAU_DAU + 1))
            G, H, I = (g.cell(r, c).value for c in (7, 8, 9))
            so = lambda v: isinstance(v, (int, float))      # noqa: E731
            ds.append({
                "stt": g.cell(r, 3).value, "ten": str(g.cell(a, 4).value or ""),
                "ky_hieu": str(g.cell(a, 11).value or ""),
                "yeu_cau": str(g.cell(a, 12).value or ""),
                "lsl": G + H if so(G) and so(H) else None,
                "usl": G + I if so(G) and so(I) else None,
                "X": X, "hang_x": r, "hang_rm": r + 1,
                "du_lieu": raw, "so_diem_do": b - a + 1, "loi_du_lieu": errors,
            })
        if not ds:
            raise LoiMau("Không đọc được chỉ tiêu nào trong sheet “Dữ liệu gốc”.")
        bc = wb[next(n for n in wb.sheetnames if "Báo cáo" in n)] \
            if any("Báo cáo" in n for n in wb.sheetnames) else None
        hang_bc = None
        if bc is not None:
            m = re.search(r"!\$?D\$?(\d+)", str(bc["P5"].value or ""))
            hang_bc = int(m.group(1)) if m else None
        return {"san_pham": g["T4"].value, "ma": g["T5"].value,
                "chi_tieu": ds, "hang_bao_cao": hang_bc, "file": duong_dan}
    finally:
        wb.close()


def xuat_bao_cao_chi_tieu(duong_dan, ct, hang_hien_tai, duong_dan_xuat, cpk_toi_thieu=1.33):
    """Chep file bao cao, tro sheet 'Bao cao' sang chi tieu ct (dict tu
    doc_bao_cao) thay vi chi tieu dang trinh bay (hang_hien_tai)."""
    z = zipfile.ZipFile(duong_dan)
    paths = _sheet_paths(z)
    ten_dl, _p = _tim_sheet(paths, "liệu gốc", "lieu goc")
    _t, p_bc = _tim_sheet(paths, "Báo cáo", "Bao cao")
    xml = _doi_hang_bao_cao(z.read(p_bc).decode("utf-8"), ten_dl,
                            {hang_hien_tai: ct["hang_x"],
                             hang_hien_tai + 1: ct["hang_rm"]})
    xml = re.sub(r'<f ca="1">CHOOSE\(RANDBETWEEN\(1,3\),BM5,BN5,BO5\)</f>',
                 "<f>BM5</f>", xml)
    xml = _xoa_gia_tri_cu(xml)
    import kich_thuoc_core as kt
    import spc_bao_cao
    r = kt.cham_chi_tieu(ct, cpk_toi_thieu)
    if "DU_LIEU" in r.loi:
        z.close()
        raise LoiMau("Chỉ tiêu thiếu/sai dữ liệu: bổ sung đủ điểm đo trước khi xuất.")
    xml, _ = spc_bao_cao.dien(xml, r, cpk_toi_thieu, mr3=False)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: "<calcPr%s fullCalcOnLoad=\"1\"/>" % re.sub(
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


def tao_thu_muc_xuat(thu_muc_goc=None):
    """Tao thu muc output voi ten co ngay gio."""
    if thu_muc_goc is None:
        thu_muc_goc = os.path.join(core.base_dir(), "output")
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    duong_dan = os.path.join(thu_muc_goc, "KichThuoc_SPC_%s" % now)
    os.makedirs(duong_dan, exist_ok=True)
    return duong_dan
