# -*- coding: utf-8 -*-
"""
XUAT THANH PHAN HOA - DIEU PHOI
===============================
chay(cfg): doc chuan -> sinh N mau -> xuat
    TongHop_ThanhPhan.xlsx
    KetQuaIn/KetQuaIn_<stt>.docx                         (N file)
    DanhGia_SPC_ThanhPhan_<mac>_<kk>_mau<a>-<b>.xlsx     (moi nhom 25)
vao <thu_muc_goc>/ThanhPhan_<mac>_<ngay gio>_<N>mau/.
"""

import os
import random
import re
from copy import copy
from datetime import datetime, timedelta

import openpyxl

try:
    import thanh_phan_core as tp
    import xuat_thanh_phan_word as xw
    import xuat_thanh_phan_spc as xs
except ImportError:
    from . import thanh_phan_core as tp
    from . import xuat_thanh_phan_word as xw
    from . import xuat_thanh_phan_spc as xs


def ten_mau_spc(mac):
    return ("JIS.QT.07.BM.01 Báo cáo đánh giá sự ổn định chất lượng - "
            "Thành phân hoá - %s.xlsx" % mac)


def xuat_tong_hop(duong_dan_mau, duong_dan_xuat, mau, ds, stt_dau=1):
    """Nhan ban file mau Tong hop: dong tieu de co 'STT' o cot A, cot theo
    ten nguyen to; Ceq giu cong thuc cua dong mau (doi so hang)."""
    wb = openpyxl.load_workbook(duong_dan_mau)
    ws = wb.worksheets[0]
    r0 = next((r for r in range(1, ws.max_row + 1)
               if str(ws.cell(r, 1).value).strip().upper() == "STT"), None)
    if r0 is None:
        raise tp.LoiDuLieu("File mẫu Tổng hợp không có ô 'STT' ở cột A.")
    cot = {str(ws.cell(r0, c).value).strip(): c for c in range(2, ws.max_column + 1)
           if ws.cell(r0, c).value is not None}
    thieu = [n["ten"] for n in ds if n["ten"] not in cot]
    if thieu:
        raise tp.LoiDuLieu("File mẫu Tổng hợp thiếu cột: %s." % ", ".join(thieu))
    mau_dong = r0 + 1
    kieu = {c: copy(ws.cell(mau_dong, c)._style) for c in range(1, ws.max_column + 1)}
    ceq = ws.cell(mau_dong, cot["Ceq"]).value if "Ceq" in cot else None
    for r in range(mau_dong, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    for i, m in enumerate(mau):
        r = mau_dong + i
        for c, s in kieu.items():
            ws.cell(r, c)._style = copy(s)
        ws.cell(r, 1).value = stt_dau + i
        for n in ds:
            o = ws.cell(r, cot[n["ten"]])
            o.value = m[n["ten"]]
            if n["so_le"] is not None:
                o.number_format = "0." + "0" * n["so_le"]
        if isinstance(ceq, str) and ceq.startswith("="):
            ws.cell(r, cot["Ceq"]).value = re.sub(
                r"\b([A-Z]{1,2})%d\b" % mau_dong, lambda k: "%s%d" % (k.group(1), r), ceq)
    wb.save(duong_dan_xuat)
    return duong_dan_xuat


def chay(cfg, log=None, dung=None):
    """Chay toan bo. cfg: file_chuan, mac, nguyen_to (list|None = doc tu file),
    file_mau_spc, file_mau_word, file_mau_tong_hop, so_mau, stt_dau, seed,
    cpk, user, project, thoi_gian (datetime), tieu_de, thu_muc_goc.
    log(chuoi, loai): loai = normal | accent | success | error.
    dung(): True thi dung (raise LoiDuLieu)."""
    log = log or (lambda s, loai="normal": None)
    ch = tp.doc_chuan(cfg["file_chuan"])
    mac = cfg["mac"]
    if mac not in ch:
        raise tp.LoiDuLieu("Không có mác %s trong file dữ liệu chuẩn (có: %s)."
                           % (mac, ", ".join(ch)))
    ds = cfg.get("nguyen_to") or ch[mac]["nguyen_to"]
    ceq = ch[mac]["ceq"]
    tp.kiem_tra_chuan(ds)
    w = xw.MauWord(cfg["file_mau_word"])
    w.kiem_tra([n["ten"] for n in ds])
    xs.kiem_tra_mau_spc(cfg["file_mau_spc"])

    seed = cfg.get("seed")
    n_mau, stt_dau = int(cfg["so_mau"]), int(cfg.get("stt_dau") or 1)
    cpk = float(cfg.get("cpk") or tp.CPK_TOI_THIEU)
    log("Sinh %d mẫu %s…" % (n_mau, mac), "accent")

    def cb(xong, tong, spc_nhom):
        if not spc_nhom:
            log("  Mẫu %d nằm riêng một nhóm - không chấm SPC (I-MR cần ≥ 2 điểm)."
                % (stt_dau + xong - 1), "error")
            return
        loi = ["%s: %s" % (t, ",".join(r.loi)) for t, r in spc_nhom.items() if r.loi]
        log("  Nhóm đến mẫu %d/%d: %s" % (xong, tong, "ĐẠT SPC" if not loi
                                          else "CHƯA ĐẠT - " + "; ".join(loi)),
            "success" if not loi else "error")
    kq = tp.sinh_toan_bo(ds, n_mau, cpk, seed, ceq, cb, dung)
    if not kq["dat"]:
        # Spec 4.3: het so lan thu ma chua dat -> bao ro, KHONG xuat file
        loi = ["mẫu %d-%d / %s: %s" % (stt_dau + g["dau"], stt_dau + g["cuoi"] - 1, t,
                                       ", ".join(r.mo_ta_loi()))
               for g in kq["nhom"] for t, r in g["spc"].items() if r.loi]
        raise tp.LoiDuLieu("Không sinh được bộ số liệu đạt SPC sau %d lần thử - chưa xuất "
                           "file nào.\n%s\nHãy nới khoảng Min/Max hoặc giảm Cpk tối thiểu."
                           % (tp.SO_LAN_SINH_LAI, "\n".join(loi[:10])))

    tm = os.path.join(cfg["thu_muc_goc"], "ThanhPhan_%s_%s_%dmau" % (
        mac, datetime.now().strftime("%Y-%m-%d_%H-%M-%S"), n_mau))
    os.makedirs(os.path.join(tm, "KetQuaIn"), exist_ok=True)
    ra = {"thu_muc": tm, "word": [], "spc": [], "canh_bao": [], "ket_qua": kq}

    ra["tong_hop"] = xuat_tong_hop(cfg["file_mau_tong_hop"],
                                   os.path.join(tm, "TongHop_ThanhPhan.xlsx"),
                                   kq["mau"], ds, stt_dau)
    log("Đã xuất Tổng hợp.", "normal")

    rng = random.Random(None if seed is None else "%s-in" % seed)
    t = cfg["thoi_gian"]
    project = int(cfg["project"])
    for i, m in enumerate(kq["mau"]):
        if dung and dung():
            raise tp.LoiDuLieu("Đã dừng theo yêu cầu.")
        if i:
            t = t + timedelta(minutes=rng.randint(3, 8))
        p = os.path.join(tm, "KetQuaIn", "KetQuaIn_%d.docx" % (stt_dau + i))
        w.xuat(p, tp.ba_lan_do_mau(m, ds, rng), cfg["user"], project + i, t, cfg["tieu_de"])
        ra["word"].append(p)
        if (i + 1) % 50 == 0:
            log("  Kết quả in %d/%d" % (i + 1, n_mau), "normal")
    log("Đã xuất %d file Kết quả in." % n_mau, "normal")

    ngay = cfg["thoi_gian"].strftime("%d/%m/%Y")
    for k, g in enumerate(kq["nhom"]):
        a, b = stt_dau + g["dau"], stt_dau + g["cuoi"] - 1
        if g["cuoi"] - g["dau"] < 2:
            ra["canh_bao"].append("Mẫu %d nằm riêng một nhóm - I-MR cần ≥ 2 điểm nên "
                                  "không xuất file SPC cho mẫu này." % a)
            continue
        p = os.path.join(tm, "DanhGia_SPC_ThanhPhan_%s_%02d_mau%d-%d.xlsx" % (mac, k + 1, a, b))
        xs.xuat_spc(cfg["file_mau_spc"], p, kq["mau"][g["dau"]:g["cuoi"]], ds, g["spc"],
                    mac, cpk, ngay)
        ra["spc"].append(p)
    log("Đã xuất %d file đánh giá SPC." % len(ra["spc"]), "normal")
    for c in ra["canh_bao"]:
        log("⚠ " + c, "error")
    log("HOÀN TẤT: %s" % tm, "success")
    return ra
