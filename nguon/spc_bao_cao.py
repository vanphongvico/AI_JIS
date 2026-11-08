# -*- coding: utf-8 -*-
"""Điền biểu mẫu báo cáo SPC: để trống toàn bộ các ô tick ☐ và ô nhận xét trong vùng in."""
import os
import re
import posixpath
import zipfile
from lxml import etree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EX = "urn:schemas-microsoft-com:office:excel"

O_TICK_CHUAN = (
    "I50", "K50", "I51", "K51",
    "K68", "M68", "K69", "M69", "K70", "M70", "K71", "M71",
    "K72", "M72", "K73", "M73", "K74", "M74", "K75", "M75",
    "B78", "B79", "D89", "H89", "F95", "K95", "F96", "K96",
)


def dien(xml, r, cpk=1.33, chung=False, mr3=True):
    from spc import SheetXML
    sh = SheetXML(xml.encode("utf-8") if isinstance(xml, str) else xml)
    ticks = {}
    for i in range(1, 9):
        row = (64 if chung else 67) + i
        ticks["K%d" % row] = False
        ticks["M%d" % row] = False
    for i in range(1, 4 if chung else 3):
        row = (46 if chung else 49) + i
        ticks[("H" if chung else "I") + str(row)] = False
        ticks[("J" if chung else "K") + str(row)] = False
    ticks.update({"B75" if chung else "B78": False,
                  "B76" if chung else "B79": False,
                  "D89": False, "H89": False,
                  "F95": False, "K95": False, "F96": False, "K96": False})
    if not chung:
        for ref in ticks:
            sh.chu(ref, "☐")
    sh.chu("C70" if chung else "B73",
           "14 điểm liên tiếp luân phiên tăng và giảm\n連続する14点が交互に増加・減少している。")
    if chung:
        sh.chu("D48", "6 điểm MR liên tiếp tăng hoặc giảm\n連続する6点が増加または減少している。")
    sh.xoa("B77" if chung else "E80")
    sh.xoa("C93")
    return sh.bytes().decode("utf-8"), ticks


def checkbox(parts, sheet_path, ticks):
    """Cập nhật cả VML (Excel cũ) và ctrlProps (Excel mới)."""
    from spc import cot
    states = {}
    for path in list(parts):
        if not path.endswith(".vml"):
            continue
        tree = ET.fromstring(parts[path])
        for shape in tree.findall("{urn:schemas-microsoft-com:vml}shape"):
            data = shape.find("{%s}ClientData" % EX)
            if data is None or data.get("ObjectType", "").lower() != "checkbox":
                continue
            anchor = data.find("{%s}Anchor" % EX)
            vals = [int(x.strip()) for x in anchor.text.split(",")]
            ref = cot(vals[0] + 1) + str(vals[2] + 1)
            if ref not in ticks:
                continue
            state = bool(ticks[ref])
            states[shape.get("id").split("s")[-1]] = state
            checked = data.find("{%s}Checked" % EX)
            if checked is not None:
                data.remove(checked)
            if state:
                ET.SubElement(data, "{%s}Checked" % EX).text = "1"
        parts[path] = ET.tostring(tree)
    relpath = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    if relpath not in parts:
        return
    rels = ET.fromstring(parts[relpath])
    targets = {e.get("Id"): posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), e.get("Target"))) for e in rels}
    tree = ET.fromstring(parts[sheet_path])
    for e in tree.iter():
        if ET.QName(e).localname != "control" or e.get("shapeId") not in states:
            continue
        rid = e.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = targets.get(rid)
        if target in parts:
            control = ET.fromstring(parts[target])
            if states[e.get("shapeId")]:
                control.set("checked", "Checked")
            elif "checked" in control.attrib:
                del control.attrib["checked"]
            parts[target] = ET.tostring(control, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)


def bo_tick_vml_ctrl_bytes(filename, raw_bytes):
    """Bỏ trạng thái checked trên mọi Form Control Checkbox (.vml và ctrlProps/*.xml)."""
    if filename.endswith(".vml"):
        s = raw_bytes.decode("utf-8", errors="ignore")
        s = re.sub(r"\s*<[A-Za-z0-9_]*:?Checked>[^<]*</[A-Za-z0-9_]*:?Checked>", "", s)
        s = re.sub(r"\s*<[A-Za-z0-9_]*:?Checked\s*/>", "", s)
        return s.encode("utf-8")
    if "ctrlProps/" in filename and filename.endswith(".xml"):
        s = raw_bytes.decode("utf-8", errors="ignore")
        s = re.sub(r'\s+checked="[^"]*"', "", s)
        return s.encode("utf-8")
    return raw_bytes


def lam_sach_tick_file(duong_dan):
    """Đảm bảo các sheet báo cáo SPC (vd từ spc.sinh_bao_cao_spc của Va đập)
    để trống hoàn toàn các ô tick ☐ (kể cả mục 7 F95:K96) và ô nhận xét E80/B77/C93."""
    if not duong_dan or not os.path.isfile(duong_dan):
        return duong_dan
    import spc
    with zipfile.ZipFile(duong_dan, "r") as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        parts = {i.filename: z.read(i.filename) for i in infos}
    paths = spc._sheet_cua_mau(parts)
    da_sua = False
    for ten, p_sh in paths.items():
        low = ten.strip().lower()
        if "liệu gốc" in low or "lieu goc" in low or "đánh giá" in low or "danh gia" in low:
            continue
        raw_txt = parts[p_sh].decode("utf-8", errors="ignore")
        sh = spc.SheetXML(parts[p_sh])
        chung = ("B75" in sh.o and "B78" not in sh.o)
        if not chung:
            for ref in O_TICK_CHUAN:
                if ref in sh.o or "☑" in raw_txt or "☒" in raw_txt:
                    sh.chu(ref, "☐")
        sh.xoa("B77" if chung else "E80")
        sh.xoa("C93")
        parts[p_sh] = sh.bytes()
        if chung:
            ticks_all_false = {
                ref: False for ref in (
                    ["H47", "J47", "H48", "J48", "H49", "J49", "B75", "B76",
                     "D89", "H89", "F95", "K95", "F96", "K96"]
                    + ["K%d" % (64 + i) for i in range(1, 9)]
                    + ["M%d" % (64 + i) for i in range(1, 9)]
                )
            }
            checkbox(parts, p_sh, ticks_all_false)
        da_sua = True

    for nm in list(parts):
        if nm.endswith(".vml") or ("ctrlProps/" in nm and nm.endswith(".xml")):
            moi = bo_tick_vml_ctrl_bytes(nm, parts[nm])
            if moi != parts[nm]:
                parts[nm] = moi
                da_sua = True

    if da_sua:
        tam = duong_dan + ".tmp"
        with zipfile.ZipFile(tam, "w", zipfile.ZIP_DEFLATED) as o:
            for info in infos:
                o.writestr(info, parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tam, duong_dan)
    return duong_dan


def lam_sach_tick_thu_muc(thu_muc):
    """Làm sạch các file báo cáo SPC (.xlsx) vừa sinh trong thư mục (vd bước ① Va đập)."""
    if not thu_muc or not os.path.isdir(thu_muc):
        return
    for fn in sorted(os.listdir(thu_muc)):
        low = fn.lower()
        if low.endswith(".xlsx") and not fn.startswith("~$") and "spc" in low:
            try:
                lam_sach_tick_file(os.path.join(thu_muc, fn))
            except Exception:
                pass
