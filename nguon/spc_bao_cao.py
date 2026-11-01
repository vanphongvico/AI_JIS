# -*- coding: utf-8 -*-
"""Điền kết quả chấm tại thời điểm xuất, giữ biểu đồ và checkbox của mẫu."""
import posixpath
from lxml import etree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EX = "urn:schemas-microsoft-com:office:excel"


def dien(xml, r, cpk=1.33, chung=False, mr3=True):
    from spc import SheetXML
    sh = SheetXML(xml.encode("utf-8") if isinstance(xml, str) else xml)
    ticks = {}
    for i in range(1, 9):
        row = (64 if chung else 67) + i
        ticks["K%d" % row] = "I%d" % i not in r.loi
        ticks["M%d" % row] = "I%d" % i in r.loi
    for i in range(1, 4 if chung else 3):
        row = (46 if chung else 49) + i
        ticks[("H" if chung else "I") + str(row)] = "MR%d" % i not in r.loi
        ticks[("J" if chung else "K") + str(row)] = "MR%d" % i in r.loi
    on_dinh = not any(x.startswith(("I", "MR")) or x in ("SIGMA", "DU_LIEU") for x in r.loi)
    ticks.update({"B75" if chung else "B78": on_dinh,
                  "B76" if chung else "B79": not on_dinh,
                  "D89": r.dat, "H89": not r.dat})
    # Các hành động khắc phục cần do người có trách nhiệm lựa chọn.
    ticks.update({"F95": False, "K95": False, "F96": False, "K96": False})
    if not chung:
        for ref, yes in ticks.items():
            sh.chu(ref, "☑" if yes else "☐")
    sh.chu("C70" if chung else "B73", "14 điểm liên tiếp luân phiên tăng và giảm\n連続する14点が交互に増加・減少している。")
    if chung:
        sh.chu("D48", "6 điểm MR liên tiếp tăng hoặc giảm\n連続する6点が増加または減少している。")
    ket = "ĐẠT" if r.dat else "KHÔNG ĐẠT: " + ", ".join(r.loi)
    ng = " | Không xét Cpk (không có ngưỡng)" if r.lsl is None and r.usl is None else " | Cpk tối thiểu: %g" % cpk
    if "SPEC" in r.loi:
        ket += " (ngoài dung sai)"
    sh.chu("B77" if chung else "E80", ket + ng)
    mr = " | MR3: " + ("Có" if "MR3" in r.loi else "Không") if not chung and mr3 else ""
    sh.chu("C93", "Kết luận tại lúc xuất" + mr +
           ". Sửa số liệu/ngưỡng trong Excel: cần đánh giá và xuất lại.")
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
            # Excel chỉ chấp nhận checked="Checked"; ô không tick thì bỏ hẳn
            # thuộc tính - ghi "Unchecked" làm Excel báo lỗi phải sửa file.
            if states[e.get("shapeId")]:
                control.set("checked", "Checked")
            elif "checked" in control.attrib:
                del control.attrib["checked"]
            parts[target] = ET.tostring(control, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
