# -*- coding: utf-8 -*-
"""
XUAT BAO CAO SPC THU KEO THEO BIEU MAU CUA MAC VAT LIEU
========================================================
Bieu mau "JIS.QT.07.BM.01 Bao cao danh gia su on dinh chat luong - Keo -
SCW480 / SCW410": MOT file 5 sheet cho ca ba chi tieu

    Du lieu goc          25 mau x (ReH, Rm, A) + khoi tinh X-bar, sigma, Cpk
    Gioi han chay ReH    bieu do I-MR + bang danh gia, lay so tu hang 16
    Do ben keo Rm        nhu tren, hang 18
    Do giai dai A        nhu tren, hang 20
    Danh gia             bang tra cap Ca / Cp / Cpk

Hai bieu mau chi khac nhau o mac vat lieu (o F8) va ba nguong duoi (o C16,
C18, C20). Chon mac nao thi dung bieu mau cua mac do.

Sua thang XML thay vi dung openpyxl - cung ly do voi xuat_kich_thuoc:
openpyxl doc-roi-ghi lai lam mat checkbox va kieu dang bieu do. O day chi
viet lai cac o so lieu, moi phan con lai chep nguyen; Excel tinh lai toan bo
cong thuc khi mo file (fullCalcOnLoad).

Ba cho cua bieu mau duoc sua khi xuat:
    F87 (Cpk)   : bieu mau ghi CHOOSE(RANDBETWEEN(1,3),BE5,BF5,BG5) - moi lan
                  mo ra mot o khac nhau, co khi ra #VALUE!. Thu keo chi co
                  nguong duoi nen Cpk = (X-bar - LSL)/3sigma = o BF5.
    To/T/Ca/Cp  : khong co gioi han tren nen bieu mau hien #VALUE!; boc
                  IFERROR de hien "-".
    K9 (don vi) : ban SCW480 ghi "mm" o ca ba sheet -> MPa, MPa, %.
"""

import os
import re
import glob
import zipfile
import spc_luat
import spc_bao_cao
from xml.sax.saxutils import escape

try:
    import charpy_core as core
except ImportError:
    from . import charpy_core as core

SO_MAU = 25
COT_DAU = 6                                  # cot F = mau 1 ... AD = mau 25

# (hang so do, hang tinh, don vi) cua ba chi tieu trong sheet "Du lieu goc";
# sheet bieu do thu i (2, 3, 4 trong workbook) ung voi chi tieu thu i.
CHI_TIEU = [("ReH", 10, 16, "MPa"), ("Rm", 11, 18, "MPa"), ("A", 12, 20, "%")]


class LoiMau(Exception):
    """Loi giai thich duoc cho nguoi dung."""
    pass


def _cot(i):
    s = ""
    while i:
        i, d = divmod(i - 1, 26)
        s = chr(65 + d) + s
    return s


def file_mau(mac):
    """Bieu mau cua mac. Hai bieu mau cung cau truc, nen thieu file cua mac
    nay thi dung tam file cua mac kia - mac va nguong deu duoc ghi de."""
    tm = core.thu_muc_templates()
    p = os.path.join(tm, "TPL_SPC_Keo_%s.xlsx" % mac)
    if os.path.exists(p):
        return p
    g = sorted(glob.glob(os.path.join(tm, "TPL_SPC_Keo_*.xlsx")))
    if g:
        return g[0]
    raise LoiMau("Không tìm thấy biểu mẫu SPC thử kéo:\n%s" % p)


def _dat_o(xml, ref, ruot="", kieu=None):
    """Thay noi dung o `ref`, giu nguyen dinh dang (thuoc tinh s)."""
    def f(m):
        s = re.search(r' s="\d+"', m.group(0))
        return '<c r="%s"%s%s>%s</c>' % (ref, s.group(0) if s else "",
                                        ' t="%s"' % kieu if kieu else "", ruot)
    # thuoc tinh phai khop tung cap ten="gia tri": viet long [^>]* thi o
    # trong dang <c .../> bi hieu thanh o mo va nuot luon cac o dung sau
    moi, n = re.subn(r'<c r="%s"(?: [\w:]+="[^"]*")*\s*(?:/>|>.*?</c>)' % ref,
                     f, xml, count=1, flags=re.S)
    if not n:
        raise LoiMau("Biểu mẫu SPC thử kéo thiếu ô %s — file mẫu không đúng "
                     "cấu trúc." % ref)
    return moi


def _so(xml, ref, v):
    return _dat_o(xml, ref, "<v>%s</v>" % repr(spc_luat.so_huu_han(v, ref)))


def _chu(xml, ref, s):
    return _dat_o(xml, ref, '<is><t xml:space="preserve">%s</t></is>'
                  % escape(str(s)), "inlineStr")


def _cong_thuc(xml, ref, f):
    return _dat_o(xml, ref, "<f>%s</f>" % escape(f))


def _xoa_gia_tri_cu(xml):
    """Bo so tinh san cua file mau o moi o cong thuc (ke ca vung cong thuc
    mang) de khong ai doc phai so cu; Excel tinh lai khi mo file."""
    def o(m):
        dau, ruot = m.group(1), m.group(2)
        if "<f" not in ruot:
            return m.group(0)
        dau = re.sub(r' t="[^"]*"', "", dau)
        return "%s>%s</c>" % (dau, re.sub(r"<v>[^<]*</v>", "", ruot))
    xml = re.sub(r'(<c r="[A-Z]+\d+"[^>/]*)>(.*?)</c>', o, xml, flags=re.S)
    for a, ca, ra, cb, rb in re.findall(
            r'<f t="array" ref="(([A-Z]+)(\d+)):([A-Z]+)(\d+)"', xml):
        if ca != cb:
            continue
        for r in range(int(ra) + 1, int(rb) + 1):
            xml = re.sub(r'<c r="%s%d"((?: [a-z]+="[^"]*")*)><v>[^<]*</v></c>'
                         % (ca, r), r'<c r="%s%d"\1/>' % (ca, r), xml)
    return xml


def xuat(mac, gia_tri, nguong, cfg, duong_dan, nhan_chay="ReH"):
    """Xuat mot file bao cao SPC thu keo theo bieu mau cua mac `mac`.

    gia_tri  : {"ReH": [25 so], "Rm": [...], "A": [...]}
    nguong   : {"ReH": LSL, "Rm": LSL, "A": LSL}
    nhan_chay: ky hieu in o hang gioi han chay ("ReH", hoac "ReL" khi so
               lieu nguon chi co gioi han chay duoi)
    """
    for k, _h, _t, _u in CHI_TIEU:
        if len(gia_tri.get(k) or []) != SO_MAU:
            raise LoiMau("Biểu mẫu SPC thử kéo cần đúng %d mẫu cho chỉ tiêu "
                         "%s." % (SO_MAU, k))
    cfg = cfg or {}
    try:
        danh_gia = {k: spc_luat.cham(gia_tri[k], nguong.get(k), None,
                                    cfg.get("cpk_toi_thieu", 1.33)) for k, *_ in CHI_TIEU}
    except ValueError as e:
        raise LoiMau(str(e))

    z = zipfile.ZipFile(file_mau(mac))
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    dich = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
    dich.update((i, t) for t, i in
                re.findall(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels))
    sheet = ["xl/" + dich[i] for i in re.findall(r'<sheet [^>]*r:id="([^"]+)"', wb)]
    if len(sheet) < 4:
        z.close()
        raise LoiMau("Biểu mẫu SPC thử kéo phải có sheet Dữ liệu gốc và ba "
                     "sheet chỉ tiêu.")
    thay = {}

    # ---- Sheet "Du lieu goc": mac, so do, nguong ----
    x = z.read(sheet[0]).decode("utf-8")
    x = _chu(x, "F8", cfg.get("vat_lieu_spc") or mac)
    for k, h_do, h_tinh, _u in CHI_TIEU:
        for i, v in enumerate(gia_tri[k]):
            x = _so(x, "%s%d" % (_cot(COT_DAU + i), h_do), v)
        x = (_so(x, "C%d" % h_tinh, nguong[k]) if nguong.get(k) is not None
             else _chu(x, "C%d" % h_tinh, "-"))
        x = _chu(x, "D%d" % h_tinh, "-")
        # Tinh khong lam tron; number format chi quy dinh cach hien thi.
        for c, f in [("AF", "AE{mr}/1.128"), ("AG", "AE{r}+3*AF{r}"),
                     ("AH", "AE{r}+2*AF{r}"), ("AI", "AE{r}+AF{r}"),
                     ("AK", "AE{r}-AF{r}"), ("AL", "AE{r}-2*AF{r}"),
                     ("AM", "AE{r}-3*AF{r}"), ("AN", "3.267*AE{mr}"),
                     ("AU", 'IFERROR((AE{r}-C{r})/(3*AF{r}),"-")'),
                     ("AV", 'IFERROR((D{r}-AE{r})/(3*AF{r}),"-")')]:
            x = _cong_thuc(x, "%s%d" % (c, h_tinh), f.format(r=h_tinh, mr=h_tinh+1))
        for c, f in [("AP", "(C{r}+D{r})/2"), ("AQ", "D{r}-C{r}"),
                     ("AR", "(AJ{r}-AP{r})/(AQ{r}/2)"), ("AS", "AQ{r}/(6*AF{r})"),
                     ("AT", "(1-ABS(AR{r}))*AS{r}")]:
            x = _cong_thuc(x, "%s%d" % (c, h_tinh),
                           'IFERROR(%s,"-")' % f.format(r=h_tinh))
    if nhan_chay != "ReH":
        x = _chu(x, "E10", nhan_chay)
        x = _chu(x, "B16", nhan_chay)
    thay[sheet[0]] = _xoa_gia_tri_cu(x)

    # ---- Ba sheet chi tieu: thong tin dau trang, don vi, Cpk, chu ky ----
    for (k, _h, _t, don_vi), p in zip(CHI_TIEU, sheet[1:4]):
        x = z.read(p).decode("utf-8")
        for ref, khoa in [("E7", "ngay_danh_gia"), ("E8", "du_lieu_danh_gia"),
                          ("E10", "tieu_chuan_ap_dung"), ("E11", "ten_san_pham"),
                          ("D99", "nguoi_lap"), ("G99", "nguoi_kiem_tra"),
                          ("J99", "nguoi_phe_duyet")]:
            if cfg.get(khoa):
                x = _chu(x, ref, cfg[khoa])
        x = _chu(x, "K9", don_vi)
        x = _cong_thuc(x, "F87", "BF5")
        x = _xoa_gia_tri_cu(x)
        x, _ticks = spc_bao_cao.dien(x, danh_gia[k], cfg.get("cpk_toi_thieu", 1.33))
        thay[p] = x

    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                    r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
    ct = z.read("[Content_Types].xml").decode("utf-8")
    thay.update({
        "xl/workbook.xml": wb,
        "xl/_rels/workbook.xml.rels": re.sub(
            r'<Relationship [^>]*calcChain[^>]*/>', "", rels),
        "[Content_Types].xml": re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct)})

    tam = duong_dan + ".tmp"
    with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            out.writestr(info, thay[info.filename].encode("utf-8")
                         if info.filename in thay else z.read(info.filename),
                         compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tam, duong_dan)
    return duong_dan
