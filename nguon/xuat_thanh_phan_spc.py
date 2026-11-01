# -*- coding: utf-8 -*-
"""
XUAT FILE SPC THANH PHAN HOA (JIS.QT.07.BM.01 - Thanh phan hoa - SCWxxx)
=========================================================================
Sua thang XML (spc.SheetXML) de giu 16 bieu do, 4 checkbox, vmlDrawing.
Moi file mot nhom <= 25 mau.

Bo cuc file mau:
  'Du lieu goc'  F..AD hang 10-19 = 25 mau x (C, Si, Mn, P, S, Ni, Cr, Mo,
                 V, Cu); khoi tinh hang x = 23, 25, ..., 41 (RM = x+1),
                 C = LSL, D = USL, AE..AV chi so.
  Sheet C, Mn, Si, Ni, Cr, Mo, V, Cu: hang 5 P..BE <- 'Du lieu goc'!E..AT
                 hang x, hang 6 P..AP <- E..AE hang RM, hai bieu do MR (hang
                 10-33) va I (hang 10-34), o tick ky tu ☐/☑.

Loi cua file mau duoc SUA tren ban sao:
  1. Hang 6 (MR) moi sheet nguyen to tro 'Du lieu goc'!F24.. (RM cua C)
     -> tro dung hang RM cua nguyen to.
  2. Sheet Mo: BE5 tro AT39 (V); Mo, V: M10 tro C39 -> viet lai toan bo
     hang 5/6 va M10/M11 theo dung hang cua nguyen to.
  3. ROUND(sigma,2) / ROUND(gioi han,1) -> theo so chu so d (d+2 / d+1).
  4. Ca/Cp/Cpk #VALUE! khi chi co USL -> Cpk mot phia, Ca/Cp '-'.
  5. Cot AO (LCLMR) ghi so rac 0,0,0,1,2,3,4,5,6,8 -> duong LCLMR bieu do MR
     cua Ni/Cr/Mo/V/Cu nam o 3-8 % -> ghi 0.
  6. File mau luu vi tri cuon cua lan sua cuoi (Cr cuon sang cot I nen trong
     nhu bi day sang phai, 'Du lieu goc' cuon toi hang 19-22 che bang so lieu)
     -> moi sheet mo tu dau trang.
  7. Bieu do I co 6 duong +-1/2/3 sigma (AR5..AX5) nhung BO QUA AU5 = X-bar:
     khong co duong tam, mat lay duong -1 sigma lam tam. Chuoi so lieu lai ten
     'X-bar'. -> them cot AE10:AE34 = $AU$5 + chuoi 'X-bar' (net den), chuoi
     so lieu doi ten 'X'.

Muc 6 Danh gia nang luc: Ca/Cp/Cpk ghi vao bang BEN PHAI (O-W, ngoai vung in);
bang trong vung in (F-M) de trong de nguoi danh gia chep tay sau khi in.
"""

import re
import zipfile

import openpyxl

try:
    import spc
    import spc_bao_cao
    import thanh_phan_core as tp
except ImportError:
    from . import spc, spc_bao_cao
    from . import thanh_phan_core as tp

SheetXML, cot = spc.SheetXML, spc.cot
DL = "'Dữ liệu gốc'!"
TEN_DL = "Dữ liệu gốc"
SO_MAU = 25
COT_DAU = 6                                    # F = mau 1 tren 'Du lieu goc'
HANG_DL = {t: 10 + i for i, t in enumerate(tp.PHAN_TU_SPC)}
HANG_X = {t: 23 + 2 * i for i, t in enumerate(tp.PHAN_TU_SPC)}
SHEET_NT = ("C", "Mn", "Si", "Ni", "Cr", "Mo", "V", "Cu")
LECH = 11                                      # cot sheet nguyen to - 11 = cot Du lieu goc


class LoiMauSPC(Exception):
    pass


def kiem_tra_mau_spc(duong_dan_mau):
    """Bao loi som neu file mau khong dung bo cuc da biet."""
    wb = openpyxl.load_workbook(duong_dan_mau)
    try:
        if TEN_DL not in wb.sheetnames:
            raise LoiMauSPC("File mẫu SPC thiếu sheet '%s'." % TEN_DL)
        thieu = [t for t in SHEET_NT if t not in wb.sheetnames]
        if thieu:
            raise LoiMauSPC("File mẫu SPC thiếu sheet: %s." % ", ".join(thieu))
        dl = wb[TEN_DL]
        for t, r in HANG_DL.items():
            if str(dl.cell(r, 5).value).strip() != t:
                raise LoiMauSPC("File mẫu SPC: ô E%d phải là '%s' (đang là '%s')."
                                % (r, t, dl.cell(r, 5).value))
    finally:
        wb.close()


def _so_hoac_gach(sh, ref, v):
    if v is None:
        sh.chu(ref, "-")
    else:
        sh.dat(ref, gia_tri=v)


def _sua_du_lieu_goc(sh, mau, ds, mac):
    n = len(mau)
    nt = {x["ten"]: x for x in ds}
    sh.chu("F8", mac)
    for t in tp.PHAN_TU_SPC:
        r, x, rm, d = HANG_DL[t], HANG_X[t], HANG_X[t] + 1, nt[t]["so_le"]
        X = [m[t] for m in mau]
        for k in range(SO_MAU):
            c = cot(COT_DAU + k)
            if k < n:
                sh.dat("%s%d" % (c, r), gia_tri=X[k])
                sh.dat("%s%d" % (c, x), X[k], "%s%d" % (c, r))
                if k:
                    sh.dat("%s%d" % (c, rm), abs(X[k] - X[k - 1]),
                           "ABS(%s%d-%s%d)" % (c, x, cot(COT_DAU + k - 1), x))
            else:
                for rr in (r, x, rm):
                    sh.xoa("%s%d" % (c, rr))
        lsl, usl = nt[t]["lsl"], nt[t]["usl"]
        _so_hoac_gach(sh, "C%d" % x, lsl)
        _so_hoac_gach(sh, "D%d" % x, usl)
        c = tp.chi_so(X, lsl, usl, d)
        f = lambda s: s % {"x": x, "rm": rm}   # noqa: E731
        sh.dat("AE%d" % x, c["xb"], f("AVERAGE(F%(x)d:AD%(x)d)"))
        sh.dat("AE%d" % rm, c["mrb"], f("AVERAGE(F%(rm)d:AD%(rm)d)"))
        sh.dat("AF%d" % x, c["sigma"], f("ROUND(AE%(rm)d/1.128,") + "%d)" % (d + 2))
        for col, k, khoa in (("AG", 3, "u3"), ("AH", 2, "u2"), ("AI", 1, "u1"),
                             ("AK", -1, "l1"), ("AL", -2, "l2"), ("AM", -3, "l3")):
            sh.dat("%s%d" % (col, x), c[khoa],
                   "ROUND(AE%d%s%d*AF%d,%d)" % (x, "+" if k > 0 else "-", abs(k), x, d + 1))
        sh.dat("AJ%d" % x, c["xb"], "AE%d" % x)
        sh.dat("AN%d" % x, c["ucl_mr"], "ROUND(3.267*AE%d,%d)" % (rm, d + 2))
        sh.dat("AO%d" % x, gia_tri=0)          # LCLMR = 0 (mau ghi so rac 0..8)
        if lsl is not None and usl is not None:
            sh.dat("AP%d" % x, (lsl + usl) / 2.0, f("(C%(x)d+D%(x)d)/2"))
            sh.dat("AQ%d" % x, usl - lsl, f("D%(x)d-C%(x)d"))
            sh.dat("AR%d" % x, c["ca"], f("(AJ%(x)d-AP%(x)d)/(AQ%(x)d/2)"))
            sh.dat("AS%d" % x, c["cp"], f("AQ%(x)d/(6*AF%(x)d)"))
            sh.dat("AT%d" % x, c["cpk"], f("(1-ABS(AR%(x)d))*AS%(x)d"))
        else:
            for col in ("AP", "AQ", "AR", "AS"):
                sh.chu("%s%d" % (col, x), "-")
            if usl is not None:
                sh.dat("AT%d" % x, c["cpk"], f("(D%(x)d-AJ%(x)d)/(3*AF%(x)d)"))
            elif lsl is not None:
                sh.dat("AT%d" % x, c["cpk"], f("(AJ%(x)d-C%(x)d)/(3*AF%(x)d)"))
            else:
                sh.chu("AT%d" % x, "-")
        for col, gh, ct in (("AU", lsl, "(AJ%(x)d-C%(x)d)/(3*AF%(x)d)"),
                            ("AV", usl, "(D%(x)d-AJ%(x)d)/(3*AF%(x)d)")):
            if gh is None or not c["sigma"]:
                sh.chu("%s%d" % (col, x), "-")
            else:
                v = (c["xb"] - gh) if col == "AU" else (gh - c["xb"])
                sh.dat("%s%d" % (col, x), v / (3 * c["sigma"]), f(ct))


class KieuSo(object):
    """Them dinh dang so '0.000…' vao styles.xml: file mau dinh dang 1 chu so
    (lam cho kich thuoc), so thanh phan hoa 0,003 hien thanh '0.0'."""

    def __init__(self, xml):
        self.xml = xml
        self.cache = {}
        m = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", xml, re.S)
        self.xfs = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", m.group(1), re.S)
        self.moi_fmt, self.moi_xf = [], []
        ids = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', xml)]
        self.fmt = {c: int(i) for i, c in re.findall(
            r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', xml)}
        self.id_tiep = max(ids + [163]) + 1

    def kieu(self, s, so_le):
        """Style id moi = style s + dinh dang so_le chu so thap phan."""
        ma = "0." + "0" * so_le if so_le > 0 else "0"
        if (s, ma) in self.cache:
            return self.cache[(s, ma)]
        if ma not in self.fmt:
            self.fmt[ma] = self.id_tiep
            self.moi_fmt.append('<numFmt numFmtId="%d" formatCode="%s"/>' % (self.id_tiep, ma))
            self.id_tiep += 1
        xf = self.xfs[int(s)] if s is not None else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        xf = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % self.fmt[ma], xf, count=1)
        if "applyNumberFormat" in xf:
            xf = re.sub(r'applyNumberFormat="\d"', 'applyNumberFormat="1"', xf)
        else:
            xf = xf.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        moi = len(self.xfs) + len(self.moi_xf)
        self.moi_xf.append(xf)
        self.cache[(s, ma)] = str(moi)
        return str(moi)

    def ap(self, sh, ref, so_le):
        c = sh.o.get(ref)
        if c is not None:
            c.set("s", self.kieu(c.get("s"), so_le))

    def bytes(self):
        x = self.xml
        if self.moi_fmt:
            if "<numFmts" in x:
                x = re.sub(r'<numFmts count="\d+">(.*?)</numFmts>',
                           lambda m: '<numFmts count="%d">%s%s</numFmts>' % (
                               len(self.fmt), m.group(1), "".join(self.moi_fmt)), x, count=1, flags=re.S)
            else:
                x = x.replace("<fonts", '<numFmts count="%d">%s</numFmts><fonts' % (
                    len(self.moi_fmt), "".join(self.moi_fmt)), 1)
        if self.moi_xf:
            x = re.sub(r'<cellXfs count="\d+">(.*?)</cellXfs>',
                       lambda m: '<cellXfs count="%d">%s%s</cellXfs>' % (
                           len(self.xfs) + len(self.moi_xf), m.group(1), "".join(self.moi_xf)),
                       x, count=1, flags=re.S)
        return x.encode("utf-8")


def _dinh_dang_sheet_nt(sh, kieu, d):
    """So lieu B17:N17/B20:M20 d chu so; X-bar, gioi han d+1; MR-bar, sigma, UCLMR d+2."""
    for k in range(SO_MAU):
        hang = 17 if k < 13 else 20
        kieu.ap(sh, "%s%d" % (cot(2 + (k if k < 13 else k - 13)), hang), d)
    for ref, them in (("A28", 1), ("C28", 2), ("E28", 2), ("G28", 1), ("I28", 1), ("K28", 2)):
        kieu.ap(sh, ref, d + them)
    for c in "BDFHJLN":
        kieu.ap(sh, "%s33" % c, d + 1)


def _sua_sheet_nt(sh, ten, n, ngay):
    x, rm = HANG_X[ten], HANG_X[ten] + 1
    for col in range(16, 58):                                  # P5..BE5
        if sh.o.get("%s5" % cot(col)) is not None:
            sh.dat("%s5" % cot(col), cong_thuc="%s%s%d" % (DL, cot(col - LECH), x))
    for col in range(16, 43):                                  # P6..AP6
        if sh.o.get("%s6" % cot(col)) is not None:
            sh.dat("%s6" % cot(col), cong_thuc="%s%s%d" % (DL, cot(col - LECH), rm))
    for k in range(n, SO_MAU):                                 # mau thua
        sh.xoa("%s5" % cot(17 + k))
        sh.xoa("%s6" % cot(17 + k))
        d_stt, d_dl = (16, 17) if k < 13 else (19, 20)
        c = cot(2 + (k if k < 13 else k - 13))
        sh.xoa("%s%d" % (c, d_stt))
        sh.xoa("%s%d" % (c, d_dl))
    sh.dat("M10", cong_thuc="%sD%d" % (DL, x))
    sh.dat("M11", cong_thuc="%sC%d" % (DL, x))
    # Duong tam bieu do I (loi 7): cot AE = X-bar, chuoi so lieu ten 'X'
    sh.chu("X9", "X")
    spc._them_o(sh, "AE9", sh.o["AD9"].get("s") if "AD9" in sh.o else None)
    sh.chu("AE9", "X̄")
    for hang in range(10, 35):
        spc._them_o(sh, "AE%d" % hang, sh.o.get("AD%d" % hang, {}).get("s"))
        sh.dat("AE%d" % hang, cong_thuc="$AU$5")
    spc._them_o(sh, "E7", sh.o["B7"].get("s") if "B7" in sh.o else None)
    sh.chu("E7", ngay)


# Thong tin in tren sheet nguyen to (tab Danh gia SPC khai): khoa -> (o, o lay kieu)
O_THONG_TIN = (("du_lieu_danh_gia", "E8", "B8"), ("tieu_chuan_ap_dung", "E10", "B10"),
               ("ten_san_pham", "E11", "B11"), ("nguoi_lap", "D99", "G99"),
               ("nguoi_kiem_tra", "G99", "G99"), ("nguoi_phe_duyet", "J99", "G99"))


def _ghi_thong_tin(sh, thong_tin):
    """O nao khai thi ghi, o de trong giu chu cua file mau."""
    for khoa, ref, mau in O_THONG_TIN:
        v = str((thong_tin or {}).get(khoa) or "").strip()
        if v:
            spc._them_o(sh, ref, sh.o[mau].get("s") if mau in sh.o else None)
            sh.chu(ref, v)


_CHU_CAP = {85: {"A": "Rất tốt", "B": "Tốt", "C": "Kém", "D": "Không đạt"},
            86: {"A+": "Rất tốt", "A": "Tốt", "B": "Đạt", "C": "Cảnh báo", "D": "Không đạt"},
            87: {"A+": "Rất tốt", "A": "Đạt tốt", "B": "Đạt", "C": "Không đạt",
                 "D": "Không đạt nghiêm trọng"}}


def _bang_nang_luc(sh, r, on_dinh):
    """Muc 6: gia tri + cap + danh gia theo sheet 'Danh gia' cua bieu mau, ghi
    vao bang ben phai P/Q/R (ngoai vung in); bang F/G/H de trong chep tay."""
    dong = ((85, None if r.ca is None else abs(r.ca), spc.CAP_CA, True),
            (86, r.cp, spc.CAP_CP, False),
            (87, r.cpk, spc.CAP_CPK, False))
    for d, v, bang, nho in dong:
        s = sh.o["O%d" % d].get("s") if "O%d" % d in sh.o else None
        for c in "PQR":
            spc._them_o(sh, "%s%d" % (c, d), s)
        if v is None:
            sh.chu("P%d" % d, "-")
            sh.chu("Q%d" % d, "-")
            co_gh = r.lsl is not None or r.usl is not None
            sh.chu("R%d" % d, "Không tính (chỉ có giới hạn một phía)" if co_gh
                   else "Không có giới hạn kỹ thuật")
        else:
            cap = spc.xep_cap(v, bang, nho_hon=nho)
            sh.chu("P%d" % d, "%.2f" % v)
            sh.chu("Q%d" % d, cap)
            sh.chu("R%d" % d, _CHU_CAP[d][cap])
    nx, hd = spc.ket_luan_nang_luc(r, on_dinh)
    for ten, ref in spc.O_HANH_DONG.items():
        sh.chu(ref, "☑" if ten == hd else "☐")
    sh.chu("C93", "%s. Hành động đề xuất: %s. Kết luận tại lúc xuất - sửa số liệu/"
                  "ngưỡng trong Excel cần đánh giá và xuất lại." % (nx, hd))


def _ten_trong_bieu_do(x, ten):
    """Ten sheet nhu bieu do I ghi: C -> 'C' (Excel bao ngoac de khong lan voi cot C)."""
    return next((t for t in (ten, "'%s'" % ten) if "<c:f>%s!$X$10:$X$34</c:f>" % t in x), None)


def _them_duong_tam(x, ten):
    """Bieu do I cua sheet `ten`: them chuoi X-bar (AE10:AE34) nhan ban tu chuoi
    UCLI - net den lien, khong cham tron; doi ten cache chuoi so lieu thanh 'X'."""
    ten = _ten_trong_bieu_do(x, ten)
    if ten is None:
        return x
    ser = re.findall(r"<c:ser>.*?</c:ser>", x, re.S)
    goc = next((s for s in ser if "<c:f>%s!$Y$10:$Y$34</c:f>" % ten in s), None)
    if goc is None or "$AE$10:$AE$34" in x:
        return x
    so = max(int(v) for v in re.findall(r'<c:idx val="(\d+)"/>', x)) + 1
    s = re.sub(r'<c:idx val="\d+"/><c:order val="\d+"/>',
               '<c:idx val="%d"/><c:order val="%d"/>' % (so, so), goc, count=1)
    s = s.replace("$Y$9<", "$AE$9<").replace("$Y$10:$Y$34", "$AE$10:$AE$34")
    s = re.sub(r"<c:strCache>.*?</c:strCache>",
               '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>X̄</c:v></c:pt></c:strCache>',
               s, count=1, flags=re.S)
    s = re.sub(r"(<c:yVal><c:numRef><c:f>[^<]*</c:f>)<c:numCache>.*?</c:numCache>", r"\1", s, flags=re.S)
    s = re.sub(r"<c:spPr>.*?</c:spPr>",
               '<c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="000000"/>'
               '</a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>', s, count=1, flags=re.S)
    s = re.sub(r"<c:marker>.*?</c:marker>", '<c:marker><c:symbol val="none"/></c:marker>',
               s, count=1, flags=re.S)
    s = re.sub(r"<c:extLst>.*?</c:extLst>", "", s, flags=re.S)       # uniqueId cua chuoi goc
    x = x.replace(ser[-1], ser[-1] + s, 1)
    so_lieu = next((t for t in ser if "<c:f>%s!$X$10:$X$34</c:f>" % ten in t), None)
    if so_lieu:
        moi = re.sub(r"(<c:tx><c:strRef>.*?<c:v>)[^<]*(</c:v>)", r"\1X\2", so_lieu, count=1, flags=re.S)
        x = x.replace(so_lieu, moi, 1)
    return x


def _truc_mr(x, ten, d):
    """Bieu do MR cua sheet `ten`: truc doc hien d+1 chu so (dinh dang lay theo o
    so lieu d chu so + Excel chia buoc nua don vi -> nhan trung '0.0005, 0.0005')."""
    t = next((t for t in (ten, "'%s'" % ten) if "<c:f>%s!$R$10:$R$33</c:f>" % t in x), None)
    if t is None:
        return x

    def sua(m):
        v = m.group(0)
        if '<c:axPos val="l"/>' not in v:
            return v
        return re.sub(r'<c:numFmt formatCode="[^"]*" sourceLinked="\d"/>',
                      '<c:numFmt formatCode="0.%s" sourceLinked="0"/>' % ("0" * (d + 1)), v, count=1)
    return re.sub(r"<c:valAx>.*?</c:valAx>", sua, x, flags=re.S)


def _mo_tu_dau_trang(x):
    """Bo vi tri cuon / o dang chon con luu trong file mau. Giu kieu xem, zoom.
    Sheet co dong bang (frozen pane): cuon ve o dau tien sau phan dong bang."""
    def sua(m):
        sv = m.group(0)
        pn = re.search(r"<pane\b[^>]*/>", sv)
        if pn is None:
            sv = re.sub(r'(<sheetView\b[^>]*?)\s+topLeftCell="[^"]*"', r"\1", sv, count=1)
            return re.sub(r"<selection\b[^>]*/>", '<selection activeCell="A1" sqref="A1"/>', sv)
        if 'state="frozen"' not in pn.group(0):
            return sv
        so = lambda k: int(float((re.search(r'%s="([\d.]+)"' % k, pn.group(0)) or [0, 0])[1]))
        dau = "%s%d" % (cot(so("xSplit") + 1), so("ySplit") + 1)
        p = re.sub(r'\s+topLeftCell="[^"]*"', "", pn.group(0)).replace(
            "<pane", '<pane topLeftCell="%s"' % dau, 1)
        sv = sv.replace(pn.group(0), p)
        return re.sub(r'<selection\b(?=[^>]*pane="bottomRight")[^>]*/>',
                      '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>' % (dau, dau), sv)
    return re.sub(r"<sheetView\b[^>]*/>|<sheetView\b.*?</sheetView>", sua, x, flags=re.S)


def xuat_spc(duong_dan_mau, duong_dan_xuat, mau_nhom, ds, spc_nhom, mac,
             cpk_toi_thieu, ngay_danh_gia, thong_tin=None):
    """Xuat mot file SPC cho mot nhom 2..25 mau. Tra ve duong dan.
    thong_tin: {khoa O_THONG_TIN: chu} - None thi giu nguyen chu cua mau."""
    n = len(mau_nhom)
    if n < 2:
        raise LoiMauSPC("Nhóm chỉ có %d mẫu - I-MR cần ít nhất 2 điểm." % n)
    with zipfile.ZipFile(duong_dan_mau) as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        parts = {i.filename: z.read(i.filename) for i in infos}
    paths = spc._sheet_cua_mau(parts)

    sh = SheetXML(parts[paths[TEN_DL]])
    spc._bo_gia_tri_dem(sh)
    _sua_du_lieu_goc(sh, mau_nhom, ds, mac)
    parts[paths[TEN_DL]] = sh.bytes()

    kieu = KieuSo(parts["xl/styles.xml"].decode("utf-8"))
    so_le = {x["ten"]: x["so_le"] for x in ds}
    for ten in SHEET_NT:
        sh = SheetXML(parts[paths[ten]])
        spc._bo_gia_tri_dem(sh)
        r = spc_nhom[ten]
        _sua_sheet_nt(sh, ten, n, ngay_danh_gia)
        _ghi_thong_tin(sh, thong_tin)
        _dinh_dang_sheet_nt(sh, kieu, so_le[ten])
        xml, ticks = spc_bao_cao.dien(sh.bytes(), r, cpk_toi_thieu, mr3=False)
        sh = SheetXML(xml.encode("utf-8"))
        _bang_nang_luc(sh, r, ticks["B78"])
        parts[paths[ten]] = sh.bytes()
    parts["xl/styles.xml"] = kieu.bytes()
    for nm in list(parts):
        if re.match(r"xl/worksheets/sheet\d+\.xml$", nm):
            parts[nm] = _mo_tu_dau_trang(parts[nm].decode("utf-8")).encode("utf-8")
        elif re.match(r"xl/charts/chart\d+\.xml$", nm):
            s = parts[nm].decode("utf-8")
            for ten in SHEET_NT:
                if _ten_trong_bieu_do(s, ten):
                    s = _them_duong_tam(s, ten)
                s = _truc_mr(s, ten, so_le[ten])
            parts[nm] = s.encode("utf-8")

    # Thu hep vung du lieu + truc bieu do khi it hon 25 diem
    # (MR: hang 10-33, truc toi 24; I: hang 10-34, truc toi 25)
    if n < SO_MAU:
        for nm in list(parts):
            if re.match(r"xl/charts/chart\d+\.xml$", nm):
                s = parts[nm].decode("utf-8")
                s = re.sub(r"(\$[A-Z]{1,2}\$10:\$[A-Z]{1,2}\$)(33|34)\b",
                           lambda m: "%s%d" % (m.group(1), 9 + n - (m.group(2) == "33")), s)
                s = re.sub(r'<c:max val="(24|25)"/>',
                           lambda m: '<c:max val="%d"/>' % (n - (m.group(1) == "24")), s)
                parts[nm] = s.encode("utf-8")

    wb = parts["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r"<calcPr([^>]*?)\s*/>",
                lambda m: '<calcPr%s fullCalcOnLoad="1"/>' % re.sub(
                    r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)), wb)
    parts["xl/workbook.xml"] = wb.encode("utf-8")
    for p, pat in (("xl/_rels/workbook.xml.rels", r"<Relationship [^>]*calcChain[^>]*/>"),
                   ("[Content_Types].xml", r"<Override [^>]*calcChain[^>]*/>")):
        parts[p] = re.sub(pat, "", parts[p].decode("utf-8")).encode("utf-8")

    with zipfile.ZipFile(duong_dan_xuat, "w", zipfile.ZIP_DEFLATED) as o:
        for info in infos:
            if info.filename == "xl/calcChain.xml":
                continue
            o.writestr(info, parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
    return duong_dan_xuat
