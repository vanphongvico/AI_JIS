# -*- coding: utf-8 -*-
"""
BO NAP MODULE NOI BO - KHONG SUA FILE NAY KHI KHONG CAN THIET
================================================================
Cho phep mot so module logic (charpy_core, spc, tensile_core, tab_keo)
ton tai duoi dang FILE MA HOA (*.dat) thay vi *.py binh thuong.

Cach hoat dong:
  - Khi noi "import charpy_core", Python luon thu bo tim mac dinh
    truoc tien: neu thay charpy_core.py THAT trong thu muc nguon thi
    dung ngay file do (che do lam viec cua nguoi phat trien - khong
    anh huong gi ca).
  - Chi khi KHONG co file .py that (ban ban giao) thi bo tim o duoi
    file nay moi duoc goi den: doc charpy_core.dat, giai ma bang khoa
    trong thu muc _khoa/, roi nap thang vao bo nho - KHONG BAO GIO ghi
    ra dia mot file .py ro rang nao.

Nguoi dung cuoi chi can chay CHAY_PHAN_MEM.bat nhu binh thuong, hoan
toan khong biet va khong bi anh huong boi co che nay.

File nay VA thu muc _khoa/ luon phai o dang chu (khong ma hoa) - day
la phan "chia khoa" bat buoc phai co de tu giai ma khi chay offline.
"""

import os
import sys
import importlib.abc
import importlib.util

_THU_MUC_NGUON = os.path.dirname(os.path.abspath(__file__))
_THU_MUC_KHOA = os.path.join(_THU_MUC_NGUON, "_khoa")
_TEN_FILE_KHOA = os.path.join(_THU_MUC_KHOA, "he_thong.bin")

# Danh sach module duoc phep ton tai o dang ma hoa. Them ten vao day
# neu sau nay muon bao ve them file khac (nho cap nhat dong_goi.py
# tuong ung ben thu muc _CONG_CU_DEV_MA_HOA).
MODULE_BAO_VE = ("charpy_core", "spc", "tensile_core", "tab_keo")

# Lop "xao tron" NHE cho file khoa nam tren dia - day KHONG PHAI lop
# bao mat chinh (lop bao mat chinh la ma hoa Fernet/AES cua thu vien
# cryptography). Muc dich duy nhat: file khoa khong hien ra nhu mot
# chuoi base64 quen mat khi ai do lo mo bang trinh soan thao van ban.
_PEPPER = bytes((0x8F, 0x3A, 0xC1, 0x5D, 0x77, 0x02, 0xE9, 0x44,
                 0x1B, 0xAF, 0x66, 0x90, 0x2C, 0xD8, 0x5E, 0x11))


def xao_tron(du_lieu):
    """Doi xung: goi 2 lan lien tiep tra ve dung du lieu ban dau."""
    return bytes(b ^ _PEPPER[i % len(_PEPPER)] for i, b in enumerate(du_lieu))


class LoiGiaiMa(Exception):
    """Loi khi khong giai ma duoc mot module - thong bao de nguoi dung
    (hoac nguoi phat trien) hieu ngay nguyen nhan."""
    pass


def duong_dan_khoa():
    return _TEN_FILE_KHOA


def doc_khoa():
    if not os.path.exists(_TEN_FILE_KHOA):
        raise LoiGiaiMa(
            "Thiếu file khóa hệ thống:\n%s\n\n"
            "Bộ phần mềm này chưa được cấp phát đầy đủ. Hãy chép lại "
            "nguyên thư mục 'BaoCaoCoTinh' đã bàn giao, hoặc liên hệ "
            "người phát triển để được cấp lại." % _TEN_FILE_KHOA)
    with open(_TEN_FILE_KHOA, "rb") as f:
        tho = f.read()
    if len(tho) < 8:
        raise LoiGiaiMa("File khóa hệ thống rỗng hoặc bị hỏng:\n%s"
                        % _TEN_FILE_KHOA)
    return xao_tron(tho)


def _giai_ma_noi_dung(ten_module):
    """Tra ve chuoi source code da giai ma, hoac None neu khong co
    file .dat tuong ung (nghia la nen dung file .py that neu co)."""
    duong_dan = os.path.join(_THU_MUC_NGUON, ten_module + ".dat")
    if not os.path.exists(duong_dan):
        return None

    try:
        from cryptography.fernet import Fernet, InvalidToken
    except ImportError:
        raise LoiGiaiMa(
            "Thiếu thư viện 'cryptography' cần để mở khóa phần mềm.\n\n"
            "Chạy lại CHAY_PHAN_MEM.bat để phần mềm tự cài, hoặc mở "
            "Command Prompt gõ:\n    python -m pip install cryptography")

    khoa = doc_khoa()
    try:
        cipher = Fernet(khoa)
    except Exception:
        raise LoiGiaiMa(
            "File khóa hệ thống sai định dạng, không dùng được:\n%s"
            % _TEN_FILE_KHOA)

    with open(duong_dan, "rb") as f:
        du_lieu_ma_hoa = f.read()
    try:
        noi_dung = cipher.decrypt(du_lieu_ma_hoa)
    except InvalidToken:
        raise LoiGiaiMa(
            "Không mở khóa được module '%s.dat'.\n\n"
            "File khóa hệ thống không khớp với dữ liệu đã mã hóa - có "
            "thể bộ cài đã bị chép thiếu, chỉnh sửa, hoặc trộn lẫn từ "
            "hai lần bàn giao khác nhau. Chép lại nguyên thư mục gốc "
            "để khắc phục." % ten_module)
    return noi_dung.decode("utf-8")


class _BoTaiModuleMaHoa(importlib.abc.Loader):
    def __init__(self, ten_module, nguon_ma):
        self._ten = ten_module
        self._nguon_ma = nguon_ma

    def create_module(self, spec):
        return None                       # dung cach tao module mac dinh

    def exec_module(self, module):
        # Gia lap duong dan nhu file that de moi doan code trong module
        # dung os.path.dirname(os.path.abspath(__file__)) van tinh ra
        # dung thu muc "nguon" - giu nguyen logic doc template/JSON.
        duong_dan_gia = os.path.join(_THU_MUC_NGUON, self._ten + ".py")
        module.__file__ = duong_dan_gia
        ma = compile(self._nguon_ma, duong_dan_gia, "exec")
        exec(ma, module.__dict__)


class _BoTimModuleMaHoa(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path, target=None):
        if fullname not in MODULE_BAO_VE:
            return None
        noi_dung = _giai_ma_noi_dung(fullname)
        if noi_dung is None:
            return None                   # khong co .dat -> nhuong lai
        loader = _BoTaiModuleMaHoa(fullname, noi_dung)
        return importlib.util.spec_from_loader(
            fullname, loader, origin=fullname + ".dat")


_DA_KICH_HOAT = False


def kich_hoat():
    """Goi mot lan, cang som cang tot, TRUOC khi import cac module duoc
    bao ve (charpy_core, spc, tensile_core, tab_keo). An toan khi goi
    nhieu lan - lan sau khong lam gi ca.

    Bo tim o day luon duoc them vao CUOI sys.meta_path, nen bo tim mac
    dinh cua Python (doc file .py that tren dia) luon duoc thu TRUOC.
    Ket qua: co file .py that -> dung file .py; khong co -> tu dong
    giai ma file .dat tuong ung. Nguoi lam viec voi file .py khong bi
    anh huong gi."""
    global _DA_KICH_HOAT
    if _DA_KICH_HOAT:
        return
    sys.meta_path.append(_BoTimModuleMaHoa())
    _DA_KICH_HOAT = True
