# -*- coding: utf-8 -*-
"""
XUAT WORD "KET QUA IN" - BAN IN MAY QUANG PHO
==============================================
Nhan ban file mau Ket qua in.docx: chi thay chu trong word/document.xml va
word/footer*.xml (logo, font, duong ke giu nguyen - cach lam cua Charpy).

Bang: moi dong mot nguyen to, cot 0 = Name, 2 = Average, 3 = Rel. Std. Dev.,
4..6 = 3 lan do, 7 = trong. Dong duoc nhan dang THEO CHU o cot Name; dong
nguyen to khong co trong du lieu chuan (Zr) bi bo.
"""

import re
import zipfile
from xml.sax.saxutils import escape, unescape

_TR = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
_TC = re.compile(r"<w:tc>.*?</w:tc>", re.S)
_T = re.compile(r"<w:t(?: [^>]*)?>[^<]*</w:t>")


class LoiMauWord(Exception):
    pass


def _chu(xml):
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))


def _ghi_o(tc, chu):
    """Thay chu trong o: giu run dau, xoa cac w:t con lai."""
    ts = list(_T.finditer(tc))
    if not ts:
        raise LoiMauWord("Ô trong bảng Word mẫu không có chỗ ghi chữ.")
    for m in reversed(ts[1:]):
        tc = tc[:m.start()] + tc[m.end():]
    m = ts[0]
    return (tc[:m.start()] + '<w:t xml:space="preserve">%s</w:t>' % escape(chu)
            + tc[m.end():])


def _so(v, so_le):
    if so_le is None or isinstance(v, str):
        return str(v)
    return "%.*f" % (so_le, v)


def _doc_so(s):
    try:
        return float(s)
    except ValueError:
        return s


def dinh_dang_gio(t):
    """'dd/mm/yyyy   h:mm AM|PM' nhu chan trang ban in may."""
    h = t.hour % 12 or 12
    return "%s   %d:%02d %s" % (t.strftime("%d/%m/%Y"), h, t.minute,
                                "AM" if t.hour < 12 else "PM")


class MauWord(object):
    """Nap file mau mot lan, xuat nhieu file (1000 file khong doc lai mau)."""

    def __init__(self, duong_dan_mau):
        with zipfile.ZipFile(duong_dan_mau) as z:
            self.infos = [i for i in z.infolist() if not i.is_dir()]
            self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        if "word/document.xml" not in self.parts:
            raise LoiMauWord("File Word mẫu không hợp lệ (thiếu document.xml).")
        self.doc = self.parts["word/document.xml"].decode("utf-8")
        self.footer = [n for n in self.parts if re.match(r"word/footer\d+\.xml$", n)]

    def ten_nguyen_to(self):
        ra = []
        for h in _TR.findall(self.doc)[1:]:
            cs = _TC.findall(h)
            if cs:
                ra.append(_chu(cs[0]).strip())
        return ra

    def kiem_tra(self, ds_ten):
        co = set(self.ten_nguyen_to())
        thieu = [t for t in ds_ten if t not in co]
        if thieu:
            raise LoiMauWord("File Word mẫu thiếu dòng nguyên tố: %s." % ", ".join(thieu))
        for h in _TR.findall(self.doc)[1:]:
            cs = _TC.findall(h)
            if cs and _chu(cs[0]).strip() in ds_ten:
                if len(cs) < 7 or not all(_T.search(c) for c in cs[2:7]):
                    raise LoiMauWord("Dòng %s trong Word mẫu thiếu cột số liệu (cần 7 cột)."
                                     % _chu(cs[0]).strip())
        if "Element Concentrations" not in self.doc:
            raise LoiMauWord("Không thấy tiêu đề '... Element Concentrations' trong Word mẫu.")
        if not any(b"User:" in self.parts[f] for f in self.footer):
            raise LoiMauWord("Chân trang Word mẫu không có 'User: ... Project: ...'.")

    def doc_so_lieu(self):
        """Doc nguoc bang so lieu cua mot ban in da xuat:
        {ten: {"lan": [3 lan], "tb", "rsd", "so_le": d|None}} - cung dang voi
        thanh_phan_core.ba_lan_do_mau. O chu (<0.0020) giu chuoi, so_le None."""
        ra = {}
        for h in _TR.findall(self.doc):
            cs = _TC.findall(h)
            if len(cs) < 7:
                continue
            o = [unescape(_chu(c)).strip() for c in cs]
            if o[0] in ("", "Name"):
                continue
            tb = _doc_so(o[2])
            try:
                rsd = float(o[3])
            except ValueError:
                rsd = 0.0
            ra[o[0]] = {"lan": [_doc_so(x) for x in o[4:7]], "tb": tb, "rsd": rsd,
                        "so_le": None if isinstance(tb, str) else
                        (len(o[2].split(".")[1]) if "." in o[2] else 0)}
        return ra

    def ghi_so_lieu(self, duong_dan_xuat, so_lieu):
        """Chi viet lai bang so lieu (giu tieu de, chan trang, dong khong co
        trong so_lieu) - dung khi sua so cua ban in da co."""
        thay = self._bang(so_lieu, bo_dong_la=False).encode("utf-8")
        with zipfile.ZipFile(duong_dan_xuat, "w", zipfile.ZIP_DEFLATED) as o:
            for info in self.infos:
                o.writestr(info, thay if info.filename == "word/document.xml"
                           else self.parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
        return duong_dan_xuat

    def _bang(self, so_lieu, bo_dong_la=True):
        def sua(mh):
            h = mh.group(0)
            cs = list(_TC.finditer(h))
            if not cs:
                return h
            ten = _chu(cs[0].group(0)).strip()
            if ten == "Name":
                return h
            if ten not in so_lieu:
                return "" if bo_dong_la else h          # vd Zr
            s = so_lieu[ten]
            gia = {2: _so(s["tb"], s["so_le"]), 3: "%.2f" % s["rsd"]}
            for i, x in enumerate(s["lan"]):
                gia[4 + i] = _so(x, s["so_le"])
            ra, cuoi = [], 0
            for i, m in enumerate(cs):
                ra.append(h[cuoi:m.start()])
                ra.append(_ghi_o(m.group(0), gia[i]) if i in gia else m.group(0))
                cuoi = m.end()
            ra.append(h[cuoi:])
            return "".join(ra)
        return _TR.sub(sua, self.doc)

    def xuat(self, duong_dan_xuat, so_lieu, user, project, thoi_gian, tieu_de):
        doc = self._bang(so_lieu)
        doc = re.sub(r"(<w:t(?: [^>]*)?>)[^<]*Element Concentrations[^<]*(</w:t>)",
                     lambda m: m.group(1) + escape(tieu_de) + m.group(2), doc, count=1)
        thay = {"word/document.xml": doc.encode("utf-8")}
        for f in self.footer:
            x = self.parts[f].decode("utf-8")
            x = re.sub(r"<w:t(?: [^>]*)?>User:[^<]*</w:t>",
                       lambda m: '<w:t xml:space="preserve">%s</w:t>'
                       % escape("User: %s - Project: %s" % (user, project)), x, count=1)
            x = re.sub(r"<w:t(?: [^>]*)?>\d{1,2}/\d{1,2}/\d{4}[^<]*</w:t>",
                       lambda m: '<w:t xml:space="preserve">%s</w:t>'
                       % dinh_dang_gio(thoi_gian), x, count=1)
            thay[f] = x.encode("utf-8")
        with zipfile.ZipFile(duong_dan_xuat, "w", zipfile.ZIP_DEFLATED) as o:
            for info in self.infos:
                o.writestr(info, thay.get(info.filename, self.parts[info.filename]),
                           compress_type=zipfile.ZIP_DEFLATED)
        return duong_dan_xuat
