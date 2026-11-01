# -*- coding: utf-8 -*-
"""
VE BIEU DO KEO TU SO LIEU CO SAN
=================================
Hai che do, khong che do nao sinh so:

  A. SO LIEU THO CUA MAY  (chay_do)
     Doc file luc - hanh trinh ma phan mem may thu xuat ra (CSV / TXT /
     XLSX), ve DUNG day diem do duoc, tu tim FeH / FeL / Fm tren duong cong
     roi tinh ReH / ReL / Rm theo duong kinh mau. A va Z chi tinh khi co
     Lu, du do tay (file kich thuoc, hoac go thang vao file ket qua).

  B. MINH HOA TU DIEM DAC TRUNG  (chay_minh_hoa)
     Co san FeL, ReL, Fm, A, Z -> noi suy mot duong cong di qua cac diem do
     (cung cong thuc voi phan_mem_ve_bieu_do_keo.html). Hinh dang duong cong
     la do cong thuc dat ra, KHONG phai do mau thu, nen moi trang luon in
     dong "DUONG CONG MINH HOA - KHONG PHAI DU LIEU DO". Dong nay khong co
     nut tat, va bo file nay khong noi sang Buoc 2 (gop bao cao Word).

Ket qua:
    output/Keo_Ve_<ngay gio>_<N>mau/       BieuDo_<N>.pdf + KetQua_<N>.xlsx
    output/Keo_MinhHoa_<ngay gio>_<N>mau/  BieuDo_MinhHoa_<N>.pdf + SoLieu_<N>.xlsx

Tim diem chay (che do A)
------------------------
    Fm   = luc lon nhat cua ca phep thu
    FeH  = dinh dau tien ma sau do luc tut it nhat NGUONG_TUT (0,5 %) va
           chua vuot lai ngay (vung chay rong it nhat 1 % hanh trinh toi Fm)
    FeL  = luc thap nhat trong vung chay, BO QUA lan tut dau tien ngay sau
           FeH (hieu ung chuyen tiep - ISO 6892-1 / JIS Z 2241). Tat duoc
           bang cfg["bo_tut_dau"] = False.
Mau khong co diem chay ro (duong cong tron) thi chi danh dau Fm; gioi han
chay quy uoc Rp0,2 can gian ke nen phan mem khong tinh.
"""

import os
import re
import math
import glob
import unicodedata
from datetime import datetime

import openpyxl
from openpyxl.styles import Font

import sinh_keo_core as keo_core
import ve_bieu_do_pdf as ve

LoiNghiepVu = keo_core.LoiNghiepVu

NGUONG_TUT = 0.005          # luc tut 0,5 % so voi dinh thi coi la bat dau chay
RONG_VUNG_CHAY = 0.01       # vung chay rong it nhat 1 % hanh trinh toi Fm
NGUONG_DUT = 0.30           # sau Fm, luc con >= 30 % Fm thi mau chua dut
SO_DIEM_VE_TOI_DA = 6000    # nhieu hon thi thu gon truoc khi ve
DUOI_FILE = (".csv", ".txt", ".dat", ".tsv", ".xlsx", ".xlsm")


def _bo_dau(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def _chuan(s):
    """Chu thuong, bo dau, bo khoang trang - de so tieu de cot."""
    return re.sub(r"\s+", "", _bo_dau(s).lower())


def _so(v):
    """So thuc tu mot o; None neu khong phai so. Nhan ca dau phay thap phan."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    s = str(v).strip().replace("\u00a0", "")
    if not s:
        return None
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return None
    return x if math.isfinite(x) else None


# ============================================================ 1. DOC FILE
def _doc_van_ban(path):
    raw = open(path, "rb").read()
    for mh in ("utf-8-sig", "utf-16", "cp1258", "latin-1"):
        try:
            txt = raw.decode(mh)
        except (UnicodeDecodeError, UnicodeError):
            continue
        if mh == "utf-16" and "\x00" in txt:
            continue
        break
    dong = [d for d in txt.splitlines() if d.strip()]
    mau = dong[:200]

    def da_so(k):
        return sum(1 for d in mau if k in d) > len(mau) / 2.0
    if da_so("\t"):
        tach = lambda d: d.split("\t")
    elif da_so(";"):
        tach = lambda d: d.split(";")
    elif da_so(",") and sum(1 for d in mau if len(d.split()) > 1) < len(mau) / 2.0:
        tach = lambda d: d.split(",")
    else:
        tach = lambda d: d.split()
    return [[o.strip() for o in tach(d)] for d in dong]


def doc_bang(path):
    """Doc mot file thanh [(ten sheet hoac '', [dong, ...])]."""
    duoi = os.path.splitext(path)[1].lower()
    if duoi == ".xls":
        raise LoiNghiepVu("File .xls (Excel cũ) chưa đọc được:\n%s\n\nMở file "
                          "bằng Excel rồi Save As sang .xlsx." % path)
    if duoi in (".xlsx", ".xlsm"):
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception as e:
            raise LoiNghiepVu("Không đọc được file Excel:\n%s\n\n%s" % (path, e))
        ra = [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)])
              for ws in wb.worksheets]
        wb.close()
        return ra
    try:
        return [("", _doc_van_ban(path))]
    except OSError as e:
        raise LoiNghiepVu("Không đọc được file:\n%s\n\n%s" % (path, e))


TU_BO = ("time", "thoigian", "stress", "ungsuat", "mpa", "n/mm", "strain",
         "biendang", "%", "(s)", "[s]", "sec")
TU_HANH_TRINH = ("position", "displacement", "extension", "stroke",
                 "crosshead", "travel", "elongation", "hanhtrinh", "chuyenvi",
                 "dogian", "giandai")
TU_LUC = ("load", "force", "taitrong", "luc")


def _loai_cot(tieu_de):
    t = _chuan(tieu_de)
    if not t:
        return None
    if any(k in t for k in TU_BO):
        return None
    if any(k in t for k in TU_LUC):
        return "luc"
    if any(k in t for k in TU_HANH_TRINH):
        return "ht"
    if re.search(r"(^|[^a-z])kn($|[^a-z])", t) or t in ("n", "(n)", "[n]"):
        return "luc"
    if re.search(r"(^|[^a-z])mm($|[^a-z])", t):
        return "ht"
    return None


def _he_so_luc(tieu_de, gia_tri, don_vi):
    """He so doi luc ve kN."""
    if don_vi == "kN":
        return 1.0
    if don_vi == "N":
        return 0.001
    t = _chuan(tieu_de)
    if "kn" in t:
        return 1.0
    if re.search(r"[\(\[_]n[\)\]]?$", t) or t == "n":
        return 0.001
    # khong ghi don vi: mau thu keo thong thuong khong vuot 2000 kN
    return 0.001 if max(gia_tri) > 2000.0 else 1.0


def _chi_so_cot(v):
    """'B' hoac '2' -> 1 (chi so tu 0). Trong -> None."""
    v = str(v or "").strip()
    if not v:
        return None
    if v.isdigit():
        if int(v) < 1:
            raise LoiNghiepVu("Số thứ tự cột phải từ 1 trở lên.")
        return int(v) - 1
    if re.match(r"^[A-Za-z]{1,2}$", v):
        n = 0
        for c in v.upper():
            n = n * 26 + ord(c) - 64
        return n - 1
    raise LoiNghiepVu("Ô chọn cột phải là chữ cái cột (B) hoặc số thứ tự (2). "
                      "Đang là: %s" % v)


def tach_mau(ten_goc, dong, cot_ht=None, cot_luc=None, don_vi="auto"):
    """Tach mot bang thanh danh sach mau {"ten", "diem": [(mm, kN)]}.

    Bang hep (mot cot hanh trinh, mot cot luc) ra mot mau; bang rong (mot cot
    hanh trinh, nhieu cot luc) ra moi cot luc mot mau, ten lay tu tieu de.
    Tra ve [] neu bang khong co so lieu (sheet ghi chu...).
    """
    def so_o(r):
        return sum(1 for o in r if _so(o) is not None)

    dau = None
    for i in range(len(dong) - 1):
        if so_o(dong[i]) >= 2 and so_o(dong[i + 1]) >= 2:
            dau = i
            break
    if dau is None:
        return []
    so_cot = max(len(r) for r in dong[dau:dau + 50])

    tieu_de = [""] * so_cot
    for r in dong[max(0, dau - 3):dau]:
        for j, o in enumerate(r[:so_cot]):
            if o not in (None, "") and _so(o) is None:
                tieu_de[j] = (tieu_de[j] + " " + str(o)).strip()

    ix, iy = _chi_so_cot(cot_ht), _chi_so_cot(cot_luc)
    if ix is not None and iy is not None:
        ds_luc = [iy]
    else:
        loai = [_loai_cot(t) for t in tieu_de]
        ds_ht = [j for j, l in enumerate(loai) if l == "ht"]
        ds_luc = [j for j, l in enumerate(loai) if l == "luc"]
        if ix is None and iy is None and not ds_ht and not ds_luc \
                and not any(tieu_de) and so_cot == 2:
            ds_ht, ds_luc = [0], [1]          # file hai cot khong tieu de
        if ix is None and len(ds_ht) > 1:
            # nhieu cot hanh trinh (FeH_Position, FeL_Position...) la bang
            # tong hop ket qua, khong phai day diem cua mot duong cong
            raise LoiNghiepVu(
                "'%s' có %d cột hành trình nên không phải bảng số liệu của "
                "một đường cong. Nếu đúng là số liệu máy, hãy điền ô 'Cột "
                "hành trình' và 'Cột lực'." % (ten_goc, len(ds_ht)))
        if ix is None:
            ix = ds_ht[0] if ds_ht else None
        if iy is not None:
            ds_luc = [iy]
        if ix is None or not ds_luc:
            raise LoiNghiepVu(
                "Không tự nhận ra cột hành trình và cột lực trong '%s'.\n\n"
                "Tiêu đề các cột đang có: %s\n\nHãy điền ô 'Cột hành trình' "
                "và 'Cột lực' (chữ cái cột như B, hoặc số thứ tự như 2)."
                % (ten_goc, ", ".join("%d=%s" % (j + 1, t or "(trống)")
                                      for j, t in enumerate(tieu_de))))
    if max([ix] + ds_luc) >= so_cot:
        raise LoiNghiepVu("'%s' chỉ có %d cột, không có cột đã chọn."
                          % (ten_goc, so_cot))

    ra = []
    for j in ds_luc:
        diem = []
        for r in dong[dau:]:
            x = _so(r[ix]) if ix < len(r) else None
            y = _so(r[j]) if j < len(r) else None
            if x is not None and y is not None:
                diem.append((x, y))
        if len(diem) < 10:
            continue
        k = _he_so_luc(tieu_de[j], [p[1] for p in diem], don_vi)
        if k != 1.0:
            diem = [(x, y * k) for x, y in diem]
        ten = ten_goc
        if len(ds_luc) > 1:
            ten = re.sub(r"[\s_\-]*(load|force|luc)?[\s_\-]*\(?(kn|n)?\)?$", "",
                         tieu_de[j], flags=re.I).strip() or "Cot%d" % (j + 1)
        ra.append({"ten": ten, "diem": diem})
    return ra


def doc_file_may(path, cot_ht=None, cot_luc=None, don_vi="auto"):
    """Moi mau trong mot file so lieu may -> {"ten", "diem", "nguon"}."""
    goc = os.path.splitext(os.path.basename(path))[0]
    bang = doc_bang(path)
    ra = []
    for ten_sheet, dong in bang:
        try:
            ds = tach_mau(goc, dong, cot_ht, cot_luc, don_vi)
        except LoiNghiepVu:
            if len(bang) > 1:       # workbook nhieu sheet: bo qua sheet la
                continue
            raise
        for m in ds:
            if any(c["ten"] == m["ten"] and c["diem"] == m["diem"] for c in ra):
                continue            # sheet chep lai mot mau da co
            m["nguon"] = os.path.basename(path) + \
                (" [%s]" % ten_sheet if ten_sheet and len(bang) > 1 else "")
            ra.append(m)
    # nhieu sheet moi sheet mot mau: ten file trung nhau -> them ten sheet
    if len(ra) > 1 and len(set(m["ten"] for m in ra)) < len(ra):
        for m in ra:
            s = re.search(r"\[(.*)\]$", m["nguon"])
            if s:
                m["ten"] = s.group(1)
    return ra


def liet_ke_file(duong_dan):
    """Thu muc -> moi file so lieu ben trong; file -> chinh no."""
    ra = []
    for p in duong_dan:
        if os.path.isdir(p):
            for f in sorted(glob.glob(os.path.join(p, "*"))):
                if os.path.splitext(f)[1].lower() in DUOI_FILE \
                        and not os.path.basename(f).startswith("~$"):
                    ra.append(f)
        elif p:
            ra.append(p)
    return ra


def doc_kich_thuoc(path):
    """File kich thuoc do tay -> {ten mau (chu thuong): {d, du, Lo, Lu}}.

    Nhan dang cot theo tieu de: Mau/NO/Ma, d, du, Lo, Lu (co hay khong co
    '(mm)' deu duoc) - dung dinh dang sheet Co_ly_tinh.
    """
    def khoa(t):
        t = re.sub(r"\(.*?\)|_?mm", "", _chuan(t))
        t = re.sub(r"[^a-z0-9]", "", t)
        if t in ("mau", "no", "ma", "mamau", "sample", "stt", "malo"):
            return "ten"
        return {"d": "d", "d0": "d", "du": "du", "lo": "Lo", "l0": "Lo",
                "lu": "Lu"}.get(t)

    for _ten, dong in doc_bang(path):
        for i, r in enumerate(dong[:15]):
            vt = {}
            for j, o in enumerate(r):
                k = khoa(o) if o not in (None, "") else None
                if k and k not in vt:
                    vt[k] = j
            if "ten" not in vt or "d" not in vt:
                continue
            ra = {}
            for r2 in dong[i + 1:]:
                ten = r2[vt["ten"]] if vt["ten"] < len(r2) else None
                if ten in (None, ""):
                    continue
                kt = {k: (_so(r2[j]) if j < len(r2) else None)
                      for k, j in vt.items() if k != "ten"}
                if kt.get("d"):
                    if isinstance(ten, float) and ten == int(ten):
                        ten = int(ten)
                    ra[str(ten).strip().lower()] = kt
            if ra:
                return ra
    raise LoiNghiepVu("File kích thước không có bảng nào gồm cột mã mẫu "
                      "(Mau / NO) và cột d:\n%s" % path)


# ====================================================== 2. PHAN TICH (A)
def _day_vung_chay(y, a, b, bo_tut_dau):
    """Chi so diem FeL trong doan [a, b] (hai dau deu thuoc vung chay)."""
    thap = min(range(a, b + 1), key=lambda i: y[i])
    if not bo_tut_dau or b - a < 3:
        return thap
    # Cac day "that": sau day luc nay len lai it nhat 10 % do sau vung chay.
    nay = 0.10 * (y[a - 1] - y[thap])            # y[a - 1] la FeH
    if nay <= 0:
        return thap
    day, i_min, i_max, xuong = [], a, a, True
    for i in range(a, b + 1):
        if xuong:
            if y[i] < y[i_min]:
                i_min = i
            elif y[i] - y[i_min] >= nay:
                day.append(i_min)
                xuong, i_max = False, i
        else:
            if y[i] > y[i_max]:
                i_max = i
            elif y[i_max] - y[i] >= nay:
                xuong, i_min = True, i
    if xuong:
        day.append(i_min)             # day cuoi, ngay truoc khi hoa ben
    if len(day) >= 2:
        return min(day[1:], key=lambda i: y[i])
    return thap


def phan_tich(diem, bo_tut_dau=True):
    """Tim Fm, diem dut va (neu co) FeH, FeL tren day diem do (mm, kN)."""
    n = len(diem)
    if n < 10:
        raise LoiNghiepVu("Chỉ có %d điểm đo - quá ít để vẽ." % n)
    x = [p[0] for p in diem]
    y = [p[1] for p in diem]
    m = max(range(n), key=lambda i: y[i])
    fm = y[m]
    if fm <= 0:
        raise LoiNghiepVu("Lực lớn nhất không dương - kiểm tra cột lực.")

    b = m
    for i in range(n - 1, m - 1, -1):
        if y[i] >= NGUONG_DUT * fm:
            b = i
            break
    kq = {"fm": fm, "fm_pos": x[m], "dut_pos": x[b], "dut_tren": y[b],
          "dut_duoi": y[b + 1] if b + 1 < n else y[b],
          "feh": None, "feh_pos": None, "fel": None, "fel_pos": None,
          "so_diem": n}

    i = next(k for k in range(n) if y[k] >= 0.2 * fm)
    dinh = i
    while i < m:
        if y[i] > y[dinh]:
            dinh = i
        elif y[i] < y[dinh] * (1.0 - NGUONG_TUT):
            j = i
            while j < m and y[j] <= y[dinh]:
                j += 1
            if abs(x[j] - x[dinh]) >= RONG_VUNG_CHAY * abs(x[m] - x[0]):
                c = _day_vung_chay(y, dinh + 1, j - 1, bo_tut_dau)
                kq.update(feh=y[dinh], feh_pos=x[dinh], fel=y[c], fel_pos=x[c])
                break
            dinh = i = j          # chi la nhieu, di tiep
            continue
        i += 1
    return kq


def tinh_co_ly(kq, kt):
    """ReH/ReL/Rm/A/Z tu ket qua phan tich va kich thuoc {d, du, Lo, Lu}."""
    d = kt.get("d")
    ra = {"S0": None, "ReH": None, "ReL": None, "Rm": None, "A": None, "Z": None}
    if d:
        s0 = math.pi * d * d / 4.0
        ra["S0"] = s0
        ra["Rm"] = kq["fm"] * 1000.0 / s0
        if kq["feh"] is not None:
            ra["ReH"] = kq["feh"] * 1000.0 / s0
            ra["ReL"] = kq["fel"] * 1000.0 / s0
        if kt.get("du"):
            ra["Z"] = (1.0 - (kt["du"] / d) ** 2) * 100.0
    if kt.get("Lo") and kt.get("Lu"):
        ra["A"] = (kt["Lu"] - kt["Lo"]) / kt["Lo"] * 100.0
    return ra


# ================================================ 3. DUONG CONG MINH HOA (B)
COT_DAC_TRUNG = ["lot", "fel", "rel", "fm", "rm", "a", "z"]
TU_KHOA_DT = {
    "lot": ["malo", "masanpham", "malosx", "lotno", "batch", "madoi", "mamau"],
    "fel": ["fel"], "rel": ["rel"], "fm": ["fm"], "rm": ["rm"],
    "a": ["giandai", "elongation", "dogian"],
    "z": ["that", "reduction", "cothat"],
}


def tach_dong_dan(van_ban):
    """Cac dong dan tu Excel, thu tu cot: Ma lo, FeL(N), ReL, Fm(N), Rm, A, Z."""
    ra = []
    for dong in van_ban.splitlines():
        if not dong.strip():
            continue
        c = [s.strip() for s in re.split(r"\t|;", dong)]
        if len(c) < 7:
            c = dong.split()
        if len(c) >= 8:                     # co cot STT o dau
            c = c[len(c) - 7:]
        so = [_so(v) for v in c[1:]] if len(c) == 7 else [None]
        if len(c) < 7 or any(v is None for v in (so[0], so[1], so[2], so[4], so[5])):
            if not ra and (len(c) < 7 or all(v is None for v in so)):
                continue                    # dong tieu de
            raise LoiNghiepVu("Dòng chưa đúng 7 cột (Mã lô, FeL, ReL, Fm, Rm, "
                              "A, Z):\n%s" % dong[:80])
        ra.append(dict(zip(COT_DAC_TRUNG, [c[0]] + so)))
    if not ra:
        raise LoiNghiepVu("Chưa có dòng số liệu nào. Dán các dòng từ Excel "
                          "theo thứ tự: Mã lô, FeL, ReL, Fm, Rm, A, Z.")
    return ra


def doc_bang_dac_trung(path):
    """Doc file Excel/CSV co cac cot FeL, ReL, Fm, Rm, A, Z -> van ban 7 cot."""
    for _ten, dong in doc_bang(path):
        tot, vt_tot, i_tot = 0, None, 0
        for i, r in enumerate(dong[:15]):
            o = [_chuan(x) for x in r]
            vt, dung = {}, set()
            for k in COT_DAC_TRUNG:
                for j, t in enumerate(o):
                    if j not in dung and any(w in t for w in TU_KHOA_DT[k]):
                        vt[k] = j
                        dung.add(j)
                        break
            if len(vt) > tot:
                tot, vt_tot, i_tot = len(vt), vt, i
        if not vt_tot or "fel" not in vt_tot or "fm" not in vt_tot:
            continue
        ra = []
        for r in dong[i_tot + 1:]:
            def o(k):
                j = vt_tot.get(k)
                return r[j] if j is not None and j < len(r) else None
            if not (_so(o("fel")) and _so(o("fm"))):
                continue
            ma = o("lot")
            if isinstance(ma, float) and ma == int(ma):
                ma = int(ma)
            ra.append("\t".join(
                [str(ma) if ma not in (None, "") else "Mau%d" % (len(ra) + 1)]
                + ["%g" % _so(o(k)) if _so(o(k)) is not None else ""
                   for k in COT_DAC_TRUNG[1:]]))
        if ra:
            return "\n".join(ra)
    raise LoiNghiepVu("Không tìm thấy bảng có cột FeL và Fm trong file:\n%s"
                      % path)


def duong_cong_minh_hoa(d):
    """Noi suy duong cong Load(kN) - Strain tu FeL(N), ReL, Fm(N), A, Z."""
    fel, fm, rel = d["fel"] / 1000.0, d["fm"] / 1000.0, d["rel"]
    a, z = d["a"], d["z"]
    if fm <= fel:
        raise LoiNghiepVu("Mẫu %s: Fm phải lớn hơn FeL." % d["lot"])
    if not (1 < a <= 80):
        raise LoiNghiepVu("Mẫu %s: A nên nằm trong khoảng 1 đến 80 %%." % d["lot"])
    ey = min(rel / 210000.0, 0.004)
    ep = max(0.01, ey * 4)
    ef = a / 100.0
    em = 0.64 * ef
    ff = fm * (1 - 0.0045 * min(max(z, 0), 90))
    fup = fel * 1.03
    P = [(ey * i / 20.0, fup * i / 20.0) for i in range(21)]
    for i in range(1, 41):
        e = ey + (ep - ey) * i / 40.0
        P.append((e, fel + (fup - fel) * math.exp(-(e - ey) / 0.0006)
                  + 0.15 * math.sin(12 * i / 40.0) * (1 - i / 40.0)))

    def herm(t, p0, p1, m0, m1):
        return (p0 * (2 * t ** 3 - 3 * t ** 2 + 1) + m0 * (t ** 3 - 2 * t ** 2 + t)
                + p1 * (-2 * t ** 3 + 3 * t ** 2) + m1 * (t ** 3 - t ** 2))
    m0, m1 = 2.8 * (fm - fel), -3.0 * (fm - ff)
    for i in range(1, 301):
        t = i / 300.0
        P.append((ep + (em - ep) * t, herm(t, fel, fm, m0, 0)))
    for i in range(1, 201):
        t = i / 200.0
        P.append((em + (ef - em) * t, herm(t, fm, ff, 0, m1)))
    return {"P": P, "ey": ey, "fel": fel, "fup": fup, "em": em, "fm": fm,
            "ef": ef, "ff": ff}


# ============================================================ 4. VE TRANG PDF
RONG, CAO = ve.RONG_TRANG, ve.CAO_TRANG
TRAI, PHAI, DUOI, TREN = 62.0, 738.0, 62.0, 598.0
DO_DAM = (0.78, 0.0, 0.0)


def _latin(s):
    return _bo_dau(s).encode("latin-1", "replace").decode("latin-1")


def _chu(s, x, y, co=11, font="F1", mau=(0, 0, 0)):
    return ("%.3f %.3f %.3f rg BT /%s %d Tf 1 0 0 1 %.2f %.2f Tm (%s) Tj ET\n"
            % (mau[0], mau[1], mau[2], font, co, x, y, ve._thoat(_latin(s))))


def _duong(x1, y1, x2, y2):
    return "n %.3f %.3f m %.3f %.3f l S\n" % (x1, y1, x2, y2)


def buoc_dep(v, so_vach=12):
    """Buoc luoi tron (1, 2, 5 x 10^k) de truc co khoang so_vach vach."""
    tho = max(v, 1e-12) / float(so_vach)
    k = 10.0 ** math.floor(math.log10(tho))
    for m in (1, 2, 5, 10):
        if m * k >= tho * (1 - 1e-9):
            return m * k


def thu_gon(diem, toi_da=SO_DIEM_VE_TOI_DA):
    """Giam so diem ve ma giu nguyen dinh/day: moi nhom lay min va max."""
    n = len(diem)
    if n <= toi_da:
        return diem
    nhom = int(math.ceil(n / (toi_da / 2.0)))
    ra = []
    for i in range(0, n, nhom):
        khuc = range(i, min(i + nhom, n))
        a = min(khuc, key=lambda k: diem[k][1])
        b = max(khuc, key=lambda k: diem[k][1])
        ra.extend(diem[k] for k in sorted(set((a, b))))
    return ra


def ve_trang(diem, nhan_x, nhan_y, moc, tieu_de, chan=(), canh_bao=None,
             chu_chim=None):
    """Content stream mot trang bieu do.

    moc: [(ten, x, y, (r, g, b), dx, dy, [dong chu])] - cham tron + chu thich,
    dx/dy la do lech cua dong chu dau tien so voi cham (pt, y huong len).
    """
    x_lon = max(max(p[0] for p in diem), 1e-9)
    y_lon = max(max(p[1] for p in diem), 1e-9)
    bx, by = buoc_dep(x_lon), buoc_dep(y_lon * 1.08)
    x_max = (math.floor(x_lon / bx + 1e-9) + 1) * bx
    y_max = (math.floor(y_lon * 1.08 / by + 1e-9) + 1) * by
    sx, sy = (PHAI - TRAI) / x_max, (TREN - DUOI) / y_max

    def px(v):
        return TRAI + v * sx

    def py(v):
        return DUOI + max(v, 0.0) * sy

    o = ["1 1 1 rg n 0 0 %.2f %.2f re f\n" % (RONG, CAO)]
    o.append("%.4f %.4f %.4f RG 0.8 w 0 J\n" % (ve.XAM, ve.XAM, ve.XAM))
    nx, ny = int(round(x_max / bx)), int(round(y_max / by))
    for i in range(1, nx + 1):
        o.append(_duong(px(i * bx), DUOI, px(i * bx), TREN))
    for i in range(1, ny + 1):
        o.append(_duong(TRAI, py(i * by), PHAI, py(i * by)))
    o.append("0 0 0 RG 1.6 w\n")
    o.append(_duong(TRAI, DUOI, PHAI + 10, DUOI))
    o.append(_duong(TRAI, DUOI, TRAI, TREN + 10))
    for i in range(nx + 1):
        s = "%g" % round(i * bx, 10)
        o.append(_chu(s, px(i * bx) - ve._be_rong(s, 9) / 2.0, DUOI - 14, 9))
    for i in range(ny + 1):
        s = "%g" % round(i * by, 10)
        o.append(_chu(s, TRAI - 6 - ve._be_rong(s, 9), py(i * by) - 3, 9))
    o.append(_chu(nhan_y, TRAI + 8, TREN + 4, 11))
    o.append(_chu(nhan_x, PHAI - ve._be_rong(nhan_x, 11), DUOI + 8, 11))

    if chu_chim:
        g = math.radians(24)
        o.append("1 0.80 0.80 rg BT /F2 54 Tf %.4f %.4f %.4f %.4f %.2f %.2f Tm "
                 "(%s) Tj ET\n" % (math.cos(g), math.sin(g), -math.sin(g),
                                   math.cos(g), TRAI + 70, DUOI + 150,
                                   ve._thoat(_latin(chu_chim))))

    for _ten, x, y, mau, _dx, _dy, _dong in moc:
        o.append("%.4f %.4f %.4f rg\n" % mau)
        o.append(ve._hinh_tron(px(x), py(y), 4.5))

    d = thu_gon(diem)
    o.append("0 0 0 RG 1.1 w 1 j 1 J\n")
    o.append("n %.3f %.3f m " % (px(d[0][0]), py(d[0][1])))
    for x, y in d[1:]:
        o.append("%.3f %.3f l " % (px(x), py(y)))
    o.append("S\n")

    for _ten, x, y, _mau, dx, dy, dong in moc:
        tx = min(max(px(x) + dx, TRAI + 4), PHAI - 118)
        ty = min(max(py(y) + dy, DUOI + 6 + 13 * (len(dong) - 1)), TREN - 12)
        o.append("0 0 0 RG 0.7 w\n")
        o.append(_duong(px(x), py(y), tx - 2, ty + 4))
        for i, s in enumerate(dong):
            o.append(_chu(s, tx, ty - 13 * i, 10))

    o.append(_chu(tieu_de, TRAI, CAO - 24, 13, "F2"))
    if canh_bao:
        o.append(_chu(canh_bao, TRAI, CAO - 42, 11, "F2", DO_DAM))
    for i, s in enumerate(chan):
        o.append(_chu(s, TRAI, 26 - 12 * i, 8, mau=(0.3, 0.3, 0.3)))
    return "".join(o).encode("latin-1")


class TrangDo(object):
    """Trang bieu do ve tu so lieu tho cua may."""

    def __init__(self, mau):
        self.mau = mau

    def noi_dung(self):
        m, k = self.mau, self.mau["kq"]
        moc = []
        if k["feh"] is not None:
            moc.append(("FeH", k["feh_pos"], k["feh"], ve.MAU_FEH, 18, 58,
                        ["FeH:", "Position:%.3f mm" % k["feh_pos"],
                         "Load:%.3f kN" % k["feh"]]))
            moc.append(("FeL", k["fel_pos"], k["fel"], ve.MAU_FEL, 44, -34,
                        ["FeL:", "Position:%.3f mm" % k["fel_pos"],
                         "Load:%.3f kN" % k["fel"]]))
        moc.append(("Fm", k["fm_pos"], k["fm"], ve.MAU_FM, -46, -30,
                    ["Fm:", "Position:%.3f mm" % k["fm_pos"],
                     "Load:%.3f kN" % k["fm"]]))
        chan = ["Nguon so lieu: %s  -  %d diem do  -  ve luc %s"
                % (m["nguon"], k["so_diem"], m.get("luc_ve", ""))]
        if k["feh"] is None:
            chan.append("Khong thay diem chay ro tren duong cong - chi danh "
                        "dau Fm.")
        return ve_trang(m["diem"], "Position(mm)", "Load(kN)", moc,
                        "Mau: %s" % m["ten"], chan)


def _so_bien_dang(v):
    return "%.5f" % v if v < 0.001 else ("%.4f" % v if v < 0.01 else "%.3f" % v)


class TrangMinhHoa(object):
    """Trang duong cong minh hoa - luon in dong canh bao."""

    CANH_BAO = ("DUONG CONG MINH HOA - noi suy tu FeL, ReL, Fm, A, Z - "
                "KHONG PHAI DU LIEU DO")

    def __init__(self, dac_trung):
        self.d = dac_trung
        self.c = duong_cong_minh_hoa(dac_trung)

    def noi_dung(self):
        c = self.c
        moc = [
            ("FeH", c["ey"], c["fup"], (0.75, 0.13, 0.75), 40, 80,
             ["FeH:", "Strain:%s /" % _so_bien_dang(c["ey"]),
              "Load:%.3f kN" % c["fup"]]),
            ("FeL", c["ey"], c["fel"], (0.13, 0.75, 0.13), 40, -50,
             ["FeL:", "Strain:%s /" % _so_bien_dang(c["ey"]),
              "Load:%.3f kN" % c["fel"]]),
            ("Fm", c["em"], c["fm"], ve.MAU_FM, -8, -28,
             ["Fm:", "Strain:%s /" % _so_bien_dang(c["em"]),
              "Load:%.3f kN" % c["fm"]]),
        ]
        chan = ["Hinh dang duong cong do cong thuc noi suy dat ra, khong "
                "phai do mau thu. Khong dung thay bieu do cua may thu."]
        return ve_trang(c["P"], "Strain(/)", "Load(kN)", moc,
                        "Mau: %s" % self.d["lot"], chan,
                        canh_bao=self.CANH_BAO, chu_chim="MINH HOA")


# ================================================================ 5. XUAT
def _thu_muc_xuat(tien_to, n, thu_muc=None):
    goc = thu_muc or os.path.join(keo_core.base_dir(), "output")
    p = os.path.join(goc, "%s_%s_%dmau"
                     % (tien_to, datetime.now().strftime("%Y-%m-%d_%H-%M-%S"), n))
    os.makedirs(p, exist_ok=True)
    return p


def _dinh_dang(ws, rong):
    for o in ws[1]:
        o.font = Font(name="Arial", bold=True)
    for r in ws.iter_rows(min_row=2):
        for o in r:
            o.font = Font(name="Arial")
    for cot, w in rong.items():
        ws.column_dimensions[cot].width = w
    ws.freeze_panes = "B2"


COT_KET_QUA = ["Mau", "FeH_Position_mm", "FeH_Load_kN", "FeL_Position_mm",
               "FeL_Load_kN", "Fm_Position_mm", "Fm_Load_kN",
               "Break_Position_mm", "d (mm)", "du (mm)", "Lo (mm)", "Lu (mm)",
               "S0 (mm2)", "ReH (MPa)", "ReL (MPa)", "Rm (MPa)", "A (%)",
               "Z (%)", "So diem do", "Nguon so lieu", "Ghi chu"]


def xuat_excel_do(duong_dan, ds):
    """File ket qua che do A. Cac cot S0..Z la CONG THUC tinh tu luc va kich
    thuoc cung dong, nen go them d, du, Lu do tay vao la ra ket qua."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ket_qua"
    ws.append(COT_KET_QUA)
    for i, m in enumerate(ds):
        r, k, kt = i + 2, m["kq"], m["kt"]

        def tron(v, n=3):
            return None if v is None else round(v, n)
        ws.append([
            m["ten"], tron(k["feh_pos"]), tron(k["feh"]), tron(k["fel_pos"]),
            tron(k["fel"]), tron(k["fm_pos"]), tron(k["fm"]),
            tron(k["dut_pos"]), kt.get("d"), kt.get("du"), kt.get("Lo"),
            kt.get("Lu"),
            '=IF(I{r}="","",PI()*I{r}^2/4)'.format(r=r),
            '=IF(OR(M{r}="",C{r}=""),"",C{r}*1000/M{r})'.format(r=r),
            '=IF(OR(M{r}="",E{r}=""),"",E{r}*1000/M{r})'.format(r=r),
            '=IF(M{r}="","",G{r}*1000/M{r})'.format(r=r),
            '=IF(OR(K{r}="",L{r}=""),"",(L{r}-K{r})/K{r}*100)'.format(r=r),
            '=IF(OR(I{r}="",J{r}=""),"",(1-(J{r}/I{r})^2)*100)'.format(r=r),
            k["so_diem"], m["nguon"], "; ".join(m["ghi_chu"])])
        for cot, dd in (("M", "0.00"), ("N", "0.0"), ("O", "0.0"),
                        ("P", "0.0"), ("Q", "0.0"), ("R", "0.0")):
            ws["%s%d" % (cot, r)].number_format = dd
    rong = {c: 15 for c in "BCDEFGHMNOPQRS"}
    rong.update({"A": 18, "I": 9, "J": 9, "K": 9, "L": 9, "T": 34, "U": 60})
    _dinh_dang(ws, rong)

    ws2 = wb.create_sheet("Huong_dan")
    for d in [
        "So lieu trong sheet Ket_qua doc tu file luc - hanh trinh cua may thu.",
        "FeH, FeL, Fm, diem dut: phan mem tim tren duong cong do duoc "
        "(don vi kN, mm).",
        "d, du, Lo, Lu: so do tay. O nao con trong thi go thang vao day.",
        "S0, ReH, ReL, Rm, A, Z la cong thuc - tu tinh lai khi sua d, du, Lo, Lu:",
        "    S0 = pi*d^2/4      Re = Fe*1000/S0      Rm = Fm*1000/S0",
        "    A = (Lu - Lo)/Lo*100      Z = (1 - (du/d)^2)*100   (JIS Z 2241)",
        "FeL bo qua lan tut dau tien ngay sau FeH (hieu ung chuyen tiep) "
        "neu o tuy chon tuong ung dang bat.",
        "Cot Ghi chu neu ro mau nao dung d danh nghia, mau nao khong thay "
        "diem chay ro - hay doi chieu voi bieu do truoc khi dung.",
    ]:
        ws2.append([d])
    for r in ws2.iter_rows():
        r[0].font = Font(name="Arial")
    ws2.column_dimensions["A"].width = 100
    wb.save(duong_dan)
    return duong_dan


def chay_do(cfg, tien_do=None):
    """Che do A. cfg: duong_dan [file/thu muc], cot_ht, cot_luc, don_vi,
    file_kich_thuoc, d0, Lo, bo_tut_dau, thu_muc_xuat."""
    ds_file = liet_ke_file(cfg.get("duong_dan") or [])
    if not ds_file:
        raise LoiNghiepVu("Chưa chọn file số liệu máy nào.")
    kich_thuoc = doc_kich_thuoc(cfg["file_kich_thuoc"]) \
        if cfg.get("file_kich_thuoc") else {}

    ds, bo_qua = [], []
    for i, f in enumerate(ds_file):
        try:
            mau_file = doc_file_may(f, cfg.get("cot_ht"), cfg.get("cot_luc"),
                                    cfg.get("don_vi") or "auto")
        except LoiNghiepVu as e:
            if len(ds_file) == 1:
                raise
            bo_qua.append("%s: %s" % (os.path.basename(f),
                                      str(e).splitlines()[0]))
            continue
        if not mau_file:
            bo_qua.append("%s: không có cột số liệu" % os.path.basename(f))
        ds.extend(mau_file)
        if tien_do:
            tien_do(i + 1, len(ds_file))
    if not ds:
        raise LoiNghiepVu("Không đọc được mẫu nào.\n\n" + "\n".join(bo_qua))

    luc_ve = datetime.now().strftime("%Y-%m-%d %H:%M")
    da_dung = set()
    for m in ds:
        m["kq"] = phan_tich(m["diem"], cfg.get("bo_tut_dau", True))
        m["luc_ve"] = luc_ve
        m["ghi_chu"] = []
        kt = dict(kich_thuoc.get(m["ten"].strip().lower()) or {})
        if kt:
            da_dung.add(m["ten"].strip().lower())
        if not kt.get("d") and cfg.get("d0"):
            kt["d"] = float(cfg["d0"])
            m["ghi_chu"].append("d = duong kinh DANH NGHIA, chua phai so do")
        if not kt.get("Lo") and cfg.get("Lo"):
            kt["Lo"] = float(cfg["Lo"])
        if not kt.get("Lu"):
            m["ghi_chu"].append("chua co Lu -> chua tinh A")
        if not kt.get("du"):
            m["ghi_chu"].append("chua co du -> chua tinh Z")
        if m["kq"]["feh"] is None:
            m["ghi_chu"].append("khong thay diem chay ro")
        m["kt"] = kt
        m["co_ly"] = tinh_co_ly(m["kq"], kt)

    n = len(ds)
    thu_muc = _thu_muc_xuat("Keo_Ve", n, cfg.get("thu_muc_xuat"))
    pdf = ve.viet_pdf(os.path.join(thu_muc, "BieuDo_%d.pdf" % n),
                      [TrangDo(m) for m in ds])
    xlsx = xuat_excel_do(os.path.join(thu_muc, "KetQua_%d.xlsx" % n), ds)
    return {"thu_muc": thu_muc, "pdf": pdf, "xlsx": xlsx, "mau": ds,
            "bo_qua": bo_qua,
            "kt_khong_khop": sorted(set(kich_thuoc) - da_dung)}


def bao_cao_do(kq):
    def s(v, dd="%.1f"):
        return "-" if v is None else dd % v
    o = ["ĐÃ VẼ XONG %d BIỂU ĐỒ từ số liệu máy." % len(kq["mau"]),
         "  PDF  : %s" % kq["pdf"], "  Excel: %s" % kq["xlsx"], "",
         "%-16s %8s %8s %8s %7s %7s %7s %6s %6s  %s"
         % ("Mẫu", "FeH kN", "FeL kN", "Fm kN", "ReH", "ReL", "Rm", "A %",
            "Z %", "Ghi chú")]
    for m in kq["mau"]:
        k, c = m["kq"], m["co_ly"]
        o.append("%-16s %8s %8s %8s %7s %7s %7s %6s %6s  %s"
                 % (m["ten"][:16], s(k["feh"], "%.3f"), s(k["fel"], "%.3f"),
                    s(k["fm"], "%.3f"), s(c["ReH"]), s(c["ReL"]), s(c["Rm"]),
                    s(c["A"]), s(c["Z"]), "; ".join(m["ghi_chu"])))
    if kq["bo_qua"]:
        o += ["", "CẢNH BÁO - bỏ qua %d file:" % len(kq["bo_qua"])] + \
             ["  " + x for x in kq["bo_qua"]]
    if kq["kt_khong_khop"]:
        o += ["", "CẢNH BÁO - file kích thước có mã mẫu không khớp tên mẫu "
                  "nào: " + ", ".join(kq["kt_khong_khop"][:20])]
    o += ["", "FeH/FeL do phần mềm tìm trên đường cong - hãy xem lại vị trí "
              "chấm trên từng biểu đồ trước khi dùng số."]
    return "\n".join(o)


def chay_minh_hoa(ds, thu_muc_xuat=None):
    """Che do B. ds: danh sach dict cua tach_dong_dan()."""
    if not ds:
        raise LoiNghiepVu("Chưa có dòng số liệu nào.")
    trang = [TrangMinhHoa(d) for d in ds]
    n = len(ds)
    thu_muc = _thu_muc_xuat("Keo_MinhHoa", n, thu_muc_xuat)
    pdf = ve.viet_pdf(os.path.join(thu_muc, "BieuDo_MinhHoa_%d.pdf" % n), trang)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "So_lieu"
    ws.append(["Ma lo", "FeL (N)", "ReL (N/mm2)", "Fm (N)", "Rm (N/mm2)",
               "A (%)", "Z (%)"])
    for d in ds:
        ws.append([d[k] for k in COT_DAC_TRUNG])
    ws.append([])
    ws.append(["So lieu do nguoi dung nhap. Bieu do kem theo la duong cong "
               "MINH HOA noi suy tu cac so nay, khong phai du lieu do."])
    _dinh_dang(ws, {c: 14 for c in "ABCDEFG"})
    xlsx = os.path.join(thu_muc, "SoLieu_%d.xlsx" % n)
    wb.save(xlsx)
    return {"thu_muc": thu_muc, "pdf": pdf, "xlsx": xlsx, "mau": ds}
