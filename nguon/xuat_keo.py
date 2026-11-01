# -*- coding: utf-8 -*-
"""
XUAT BO FILE DAU RA CHO TINH NANG SINH BIEU DO KEO
===================================================
Ghep sinh_keo_core (so lieu) voi ve_bieu_do_pdf (hinh) thanh bo file:

    Keo_Sinh_<ngay gio>_<N>mau/
        BieuDo_<N>.pdf              N trang, moi trang mot bieu do
        SoLieu_<N>.xlsx             sheet Tong_hop dung dinh dang Tab 2 doc
        DanhGia_DoOnDinh_SPC.xlsx   bo 25 mau dau (tuy chon)
        BaoCao_Phu_SPC.txt          bang phu moi buoc nhay

Diem quan trong: PDF va Excel sinh ra tu CUNG MOT bo tham so, nen ba con so
FeH/FeL/Fm tren bieu do va trong bang luon trung khop tuyet doi. Bo file mau
cu ('1000_bieu_do_keo_theo_mau.pdf' + '1000_bieu_do_keo_du_lieu.xlsx') sinh
tu hai lan chay khac nhau nen lech - do la loi ma bo sinh nay xoa bo.
"""

import os
import random
import tempfile
import shutil
from datetime import datetime

import openpyxl

import sinh_keo_core as core
import ve_bieu_do_pdf as ve
import spc_luat

# Bay cot dau giu dung dinh dang file mau cu de Tab 2 doc duoc khong can sua.
# Hai cot d_mm/du_mm la phan them: co chung thi Tab 2 dung dung duong kinh da
# sinh thay vi boc lai ngau nhien, nho vay bao cao Word, sheet Co_ly_tinh va
# bieu mau BM32 cung ke mot cau chuyen.
COT_TONG_HOP = ["Mau", "FeH_Position_mm", "FeH_Load_kN",
                "FeL_Position_mm", "FeL_Load_kN",
                "Fm_Position_mm", "Fm_Load_kN",
                "Break_Position_mm", "Break_Top_Load_kN",
                "Break_Bottom_Load_kN", "d_mm", "du_mm"]

COT_CO_LY = ["NO", "d (mm)", "du (mm)", "Lo (mm)", "Lu (mm)", "Le (mm)",
             "FeL (N)", "ReL (MPa)", "FeH (N)", "ReH (MPa)",
             "Fm (N)", "Rm (MPa)", "A (%)", "Z (%)", "Ket_luan"]

BUOC_DU_LIEU = 0.5        # luoi lay mau duong cong trong sheet Du_lieu


def _sang_n(kn):
    """Luc kN -> N so nguyen: Co_ly_tinh ghi N nhu file Tong hop keo."""
    return int(round(kn * 1000))


def _duong_cong_thua(t, buoc=BUOC_DU_LIEU, het=30.0):
    """Lay mau duong cong tren luoi deu, cat o diem dut (giong file mau)."""
    d = ve.duong_cong(t)
    ra = []
    x = 0.0
    i = 0
    while x <= het + 1e-9:
        if x > t["dut_pos"]:
            ra.append((round(x, 3), None))
        else:
            while i + 1 < len(d) and d[i + 1][0] < x:
                i += 1
            if i + 1 < len(d) and d[i + 1][0] != d[i][0]:
                u = (x - d[i][0]) / (d[i + 1][0] - d[i][0])
                y = d[i][1] + u * (d[i + 1][1] - d[i][1])
            else:
                y = d[i][1]
            ra.append((round(x, 3), round(max(y, 0.0), 3)))
        x += buoc
    return ra


def xuat_excel(duong_dan, mau, thamso, cfg, kem_duong_cong=True):
    """Ghi file Excel so lieu. mau/thamso: hai list song song, cung thu tu."""
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Tong_hop"
    ws.append(COT_TONG_HOP)
    for m, t in zip(mau, thamso):
        ws.append([m["ma_mau"], t["feh_pos"], t["feh"], t["fel_pos"], t["fel"],
                   t["fm_pos"], t["fm"], t["dut_pos"], t["dut_tren"],
                   t["dut_duoi"], m["d"], m["du"]])

    ws2 = wb.create_sheet("Co_ly_tinh")
    ws2.append(COT_CO_LY)
    ws2.append([None, "mm", "mm", "mm", "mm", "mm", "N", "MPa", "N", "MPa",
                "N", "MPa", "%", "%", None])
    for m in mau:
        ws2.append([m["ma_mau"], m["d"], m["du"], m["Lo"], m["Lu"], m["Le"],
                    _sang_n(m["FeL"]), m["ReL"], _sang_n(m["FeH"]), m["ReH"],
                    _sang_n(m["Fm"]), m["Rm"], m["A"], m["Z"], "OK"])
    # ReL/ReH/Rm/A/Z co 1 chu so le - hien du ca so 572.0 (khong thanh 572)
    for cot in ("H", "J", "L", "M", "N"):
        for o in ws2[cot][2:]:
            o.number_format = "0.0"

    if kem_duong_cong:
        ws3 = wb.create_sheet("Du_lieu")
        # lay mau het truc X dai nhat (truc tu co gian, co the qua 30 mm)
        het = max([30.0] + [ve.pham_vi_truc(t)[0] for t in thamso])
        cong = [_duong_cong_thua(t, het=het) for t in thamso]
        ws3.append(["Position_mm"] + ["%s_Load_kN" % m["ma_mau"] for m in mau])
        for r in range(len(cong[0])):
            ws3.append([cong[0][r][0]] + [c[r][1] for c in cong])

        ws4 = wb.create_sheet("Bieu_do_mau")
        k = min(3, len(mau))
        ws4.append(["Position_mm"] + ["%s_Load_kN" % mau[i]["ma_mau"]
                                      for i in range(k)])
        for r in range(len(cong[0])):
            ws4.append([cong[0][r][0]] + [cong[i][r][1] for i in range(k)])

    ws5 = wb.create_sheet("Huong_dan")
    for d in [
        "File so lieu mo phong %d bieu do keo." % len(mau),
        "Don vi: Load = kN, Position = mm (rieng FeL/FeH/Fm o Co_ly_tinh = N).",
        "Sheet Tong_hop  : thong so dac trung, dung dinh dang Tab 2 doc duoc.",
        "Sheet Co_ly_tinh: thong so co ly tinh da suy ra, dung cho bieu mau BM32.",
        "Sheet Du_lieu   : duong cong lay mau moi %g mm, moi cot mot mau."
        % BUOC_DU_LIEU,
        "Bo bieu do xuat rieng dang PDF de tranh file Excel qua nang.",
        "",
        "Do that Z = (S0 - Su)/S0 x 100, S0 = pi*d^2/4, Su = pi*du^2/4 (JIS Z 2241)",
    ]:
        ws5.append([d])

    wb.save(duong_dan)
    return duong_dan


def bao_cao_phu(ds_gia_tri, cfg, con_hong, vong, giay, dung_lai=None):
    """Bang phu SPC dang chu.

    Phan quan trong nhat cua ca bao cao: noi ro cach lay mau nao duoc bao
    dam va cach nao khong. Khong lam ro cho nay thi nguoi dung de tuong la
    "lay kieu gi cung dat", trong khi thuc te chi mot so cach duoc bao dam.
    """
    tu_dau = sorted(set(int(x) for x in core._lay(cfg, "buoc_nhay") or ()))
    moi_diem = sorted(set(int(x) for x in (cfg.get("buoc_moi_diem") or ())))
    L = []
    L.append("=" * 72)
    L.append("  BANG PHU DANH GIA DO ON DINH SPC")
    L.append("=" * 72)
    L.append("  Bo luat       : 8 muc I-chart + 3 muc MR-chart + Cpk")
    L.append("  Cpk toi thieu : %g" % core._lay(cfg, "cpk_toi_thieu"))
    L.append("  Toi uu thu tu : %d vong, %.1f giay" % (vong, giay))
    L.append("")
    L.append("  DA BAO DAM:")
    if moi_diem:
        for k in moi_diem:
            L.append("    - Buoc nhay %d, BAT DAU TU BAT KY DAU trong lo%s"
                     % (k, "  (tuc: 25 mau lien tiep bat ky)" if k == 1 else ""))
    if tu_dau:
        L.append("    - Buoc nhay %s, nhung CHI khi bat dau tu dau lo"
                 % ", ".join(str(k) for k in tu_dau))
    if con_hong:
        L.append("    !! CON %d KHOI CHUA DAT - het thoi gian truoc khi xong."
                 % con_hong)
        L.append("       Tang o 'Thoi gian toi da' hoac bot buoc nhay.")
    L.append("")
    L.append("  KHONG BAO DAM: moi cach lay khac. Bang duoi cho thay thuc te.")
    L.append("-" * 72)
    L.append("  buoc      bat dau tu dau lo        bat dau tu giua lo")
    for k, d1, t1, d2, t2 in core.quet_phu(ds_gia_tri, cfg, dung_lai=dung_lai):
        c1 = ("%d/%d %s" % (d1, t1, "DAT" if d1 == t1 else "")) if t1 else "-"
        c2 = ("%d/%d (%.0f%%)" % (d2, t2, 100.0 * d2 / t2)) if t2 else "-"
        sao = ""
        if k in moi_diem:
            sao = "  <- da bao dam moi diem bat dau"
        elif k in tu_dau:
            sao = "  <- da bao dam tu dau lo"
        L.append("  %4d      %-22s  %-18s%s" % (k, c1, c2, sao))
    L.append("=" * 72)
    L.append("  LUU Y: khong the bao dam MOI cach lay 25 gia tri. Bo luat SPC")
    L.append("  duoc thiet ke de mot qua trinh on dinh that van thinh thoang")
    L.append("  cham phai; voi hang chuc nghin cua so kha di, viec khong cua")
    L.append("  so nao cham la dieu khong the xay ra.")
    L.append("=" * 72)
    return "\n".join(L)


DON_VI_COT = {"ReL": "N/mm2", "ReH": "N/mm2", "Rm": "N/mm2",
              "A": "%", "Z": "%"}
NHAN_COT = {"ReL": "Gioi han chay duoi ReL", "ReH": "Gioi han chay tren ReH",
            "Rm": "Do ben keo Rm", "A": "Do gian dai tuong doi A",
            "Z": "Do that tuong doi Z"}


def xuat_spc_25(gia_tri, cfg, thu_muc, bat_dau=0, cot=None):
    """Xuat file danh gia do on dinh SPC cho 25 mau, theo mau TPL_SPC.xlsx.

    Mac dinh xuat cho cot ReL - cot co nguong nghiem thu ro rang nhat. Khai
    cfg['cot_spc'] de doi sang cot khac, hoac list de xuat nhieu file.
    """
    import spc as _spc
    cot = cot or cfg.get("cot_spc") or ["ReL"]
    if isinstance(cot, str):
        cot = [cot]
    lsl = dict(zip(core.COT_SPC, core._lsl_cot(cfg)))
    ra = []
    for ten in cot:
        if ten not in core.COT_SPC:
            continue
        j = core.COT_SPC.index(ten)
        X = [gia_tri[i][j] for i in range(bat_dau, bat_dau + 25)]
        c = dict(cfg)
        c["lsl"] = lsl.get(ten)
        c["usl"] = None
        c["chi_tieu_danh_gia"] = NHAN_COT.get(ten, ten)
        c["don_vi"] = DON_VI_COT.get(ten, "")
        p = os.path.join(thu_muc, "DanhGia_DoOnDinh_SPC_%s_mau%d-%d.xlsx"
                         % (ten, bat_dau + 1, bat_dau + 25))
        ra.append(_spc.sinh_bao_cao_spc(X, c, p))
    return ra


def chay(cfg, thu_muc_xuat=None, tien_do=None, dung_lai=None):
    """Chi dua bo file hoan chinh ra output; huy/loi thi bo ban tam."""
    core.kiem_tra_dung(dung_lai)
    c = core.kiem_tra_cfg(cfg)
    root = os.path.abspath(thu_muc_xuat or os.path.join(core.base_dir(), "output"))
    os.makedirs(root, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".keo_dang_sinh_", dir=root) as stage:
        result = _chay(c, stage, tien_do, dung_lai)
        core.kiem_tra_dung(dung_lai)
        old = result["thu_muc"]
        target = os.path.join(root, os.path.basename(old))
        if os.path.exists(target):
            target += "_" + datetime.now().strftime("%f")
        os.replace(old, target)
        for key, v in list(result.items()):
            if isinstance(v, str) and v.startswith(old):
                result[key] = target + v[len(old):]
        result["files_spc"] = [target + p[len(old):] for p in result["files_spc"]]
        return result


def _chay(cfg, thu_muc_xuat=None, tien_do=None, dung_lai=None):
    """Sinh toan bo bo file. Tra ve dict ket qua."""
    c = dict(core.CFG_MAC_DINH)
    c.update(cfg or {})
    t0 = datetime.now()
    rnd = random.Random(c.get("seed") if c.get("seed") not in (None, "") else None)

    # 1. sinh gia tri cho nam cot chiu SPC
    gia_tri = core.sinh_gia_tri(c, rnd, dung_lai)
    n = len(gia_tri)

    # 2. hoan vi thu tu cho dat SPC
    def td(con, giay):
        if tien_do:
            tien_do("spc", con, giay)
    con_hong, vong, giay_spc = core.toi_uu_thu_tu(gia_tri, c, rnd, td, dung_lai)
    core.kiem_tra_dung(dung_lai)

    # 3. suy ra kich thuoc va tham so bieu do
    mau, thamso = [], []
    for i, hang in enumerate(gia_tri):
        core.kiem_tra_dung(dung_lai)
        m = core.suy_ra(hang, c, rnd)
        m["ma_mau"] = core.ma_mau(c, i)
        m["stt"] = i + 1
        mau.append(m)
        thamso.append(core.tham_so_duong_cong(m, c, rnd))

    # 4. ghi file
    goc = thu_muc_xuat or os.path.join(core.base_dir(), "output")
    ten = "Keo_Sinh_%s_%dmau" % (t0.strftime("%Y-%m-%d_%H-%M-%S"), n)
    out = os.path.join(goc, ten)
    os.makedirs(out, exist_ok=True)

    f_pdf = os.path.join(out, "BieuDo_%d.pdf" % n)
    trang = [ve.Trang(t, "%04d" % (i + 1), "1")
             for i, t in enumerate(thamso)]

    def td2(i, tong):
        core.kiem_tra_dung(dung_lai)
        if tien_do:
            tien_do("pdf", i, tong)
    ve.viet_pdf(f_pdf, trang, tien_do=td2)
    core.kiem_tra_dung(dung_lai)

    f_xls = os.path.join(out, "SoLieu_%d.xlsx" % n)
    xuat_excel(f_xls, mau, thamso, c,
               kem_duong_cong=bool(c.get("kem_duong_cong", True)))
    core.kiem_tra_dung(dung_lai)

    # 5. file SPC mau cho 25 mau dau (mac dinh co)
    files_spc = []
    if c.get("xuat_spc", True) and n >= 25:
        files_spc = xuat_spc_25(gia_tri, c, out)

    f_txt = os.path.join(out, "BaoCao_Phu_SPC.txt")
    bang = bao_cao_phu(gia_tri, c, con_hong, vong, giay_spc, dung_lai)
    with open(f_txt, "w", encoding="utf-8") as f:
        f.write(bang)

    # Ghi lai pham vi bao dam va cac nguong, de sau nay muon sua du lieu
    # trong thu muc thi con biet ma GIU nhung gi da bao dam.
    import json
    with open(os.path.join(out, "_thongtin_sinh.json"), "w",
              encoding="utf-8") as f:
        json.dump({
            "loai": "keo", "so_mau": n, "seed": c.get("seed"),
            "buoc_nhay": list(c.get("buoc_nhay") or []),
            "buoc_moi_diem": list(c.get("buoc_moi_diem") or []),
            "day_rieng": [list(x) for x in (c.get("day_rieng") or [])],
            "cpk_toi_thieu": c.get("cpk_toi_thieu"),
            "nguong": dict(zip(core.COT_SPC, core._lsl_cot(c))),
            "d0": c.get("d0"), "dung_sai_d": c.get("dung_sai_d"),
            "Lo": c.get("Lo"), "Le": c.get("Le"),
            "he_so_gian_dai": c.get("he_so_gian_dai"),
            "ung_suat_dut": c.get("ung_suat_dut"),
        }, f, ensure_ascii=False, indent=2)

    return {"thu_muc": out, "file_pdf": f_pdf, "file_excel": f_xls,
            "file_bang_phu": f_txt, "files_spc": files_spc,
            "mau": mau, "thamso": thamso,
            "gia_tri": gia_tri, "con_hong": con_hong, "vong": vong,
            "giay_spc": giay_spc, "bang_phu": bang, "cfg": c,
            "giay": (datetime.now() - t0).total_seconds()}


def bao_cao(kq):
    m = kq["mau"]
    L = ["DA SINH XONG %s MAU THU KEO" % "{:,}".format(len(m)), ""]
    L.append("Thu muc  : %s" % kq["thu_muc"])
    L.append("Bieu do  : %s" % os.path.basename(kq["file_pdf"]))
    L.append("So lieu  : %s" % os.path.basename(kq["file_excel"]))
    for f in kq.get("files_spc") or []:
        L.append("File SPC : %s" % os.path.basename(f))
    L.append("Thoi gian: %.1f giay  (toi uu SPC %.1f giay, %d vong)"
             % (kq["giay"], kq["giay_spc"], kq["vong"]))
    L.append("")
    for ten, khoa in [("ReL", "ReL"), ("ReH", "ReH"), ("Rm", "Rm"),
                      ("A", "A"), ("Z", "Z")]:
        v = [x[khoa] for x in m]
        L.append("  %-4s %8.1f - %-8.1f" % (ten, min(v), max(v)))
    L.append("")
    if kq["con_hong"]:
        L.append("LUU Y: con %d khoi trong danh sach buoc nhay da khai CHUA "
                 "dat. Xem BaoCao_Phu_SPC.txt, hoac tang o 'Thoi gian toi da'."
                 % kq["con_hong"])
    else:
        L.append("Moi buoc nhay da khai: DAT 100%. Chi tiet o BaoCao_Phu_SPC.txt")
    return "\n".join(L)
