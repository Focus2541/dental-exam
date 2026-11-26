"""
jhcis_link_existing.py  (อ่าน/เขียนแค่ไฟล์ Excel ไม่แตะฐานข้อมูล)
เด็กในแท็บ "4 ไม่มีใน JHCIS" ที่จริงแล้ว "มีอยู่แล้ว" ใน JHCIS แต่เลขบัตรไม่ตรง
(เช่น ในไฟล์เป็นเลข G แต่ใน JHCIS เป็นเลขที่ระบบสร้าง 13990-00-HN)
-> จับคู่ด้วย ชื่อ + สกุล + วันเกิด แล้วสร้างไฟล์ให้ jhcis_student_fill.py ใส่ชั้นให้คนเดิม (ไม่สร้างคนใหม่ซ้ำ)

ใช้:  python jhcis_link_existing.py "ผลเทียบ_....xlsx"
ได้:  จับคู่_....xlsx
   แผ่น "3 เพิ่มเข้าโรงเรียน" = ชื่อ+สกุล+วันเกิดตรงกับคนใน JHCIS 1 คนพอดี  -> ใช้กับ jhcis_student_fill.py ได้เลย
   แผ่น "ให้คนตัดสิน"        = วันเกิดไม่ตรง หรือมีหลายคน  -> ดูเอง แก้ในแผ่นนี้แล้วย้ายไปแผ่น 3 ได้
"""
import re
import jhcis_roster_compare as rc
import sys
import os

CLS = {"อ": "อนุบาล", "ป": "ประถม", "ม": "มัธยม"}


def full_cls(c):
    m = re.match(r"\s*([อปม])\.?\s*(\d)", str(c or ""))
    return f"{CLS[m.group(1)]}{m.group(2)}" if m else ""


def main(path):
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font
    wb = load_workbook(path, data_only=True)
    school = wb["สรุป"]["B1"].value
    ws = wb["4 ไม่มีใน JHCIS"]
    h = [c.value for c in ws[1]]
    hint_col = [x for x in h if x and str(x).startswith("คนใน JHCIS")][0]
    head3 = ["ชั้น(ไฟล์)", "เลขประจำตัว", "คำนำหน้า", "ชื่อ", "นามสกุล", "วันเกิด(ค.ศ.)", "pid", "typelive",
             "dischargetype", "สถานะใน JHCIS", "ชั้นที่ต้องใส่", "หมายเหตุ", "ห้อง"]
    auto, review = [], []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(h, r))
        hint = str(d.get(hint_col) or "")
        if not hint:
            continue
        cands = []
        for part in hint.split(";"):
            m = re.search(r"pid (\d+) (.+?) เกิด (\S+) เลข (\S+)", part.strip())
            if m:
                cands.append({"pid": m.group(1), "name": m.group(2).strip(), "birth": m.group(3), "id": m.group(4)})
        want = f"{d['ชื่อ']} {d['นามสกุล']}".strip()
        bd = str(d.get("วันเกิด(ค.ศ.)") or "")
        exact = [c for c in cands if c["name"] == want and c["birth"] == bd]
        cls = full_cls(d["ชั้น(ไฟล์)"])
        if len(exact) == 1 and cls:
            c = exact[0]
            auto.append([d["ชั้น(ไฟล์)"], c["id"], d["คำนำหน้า"], d["ชื่อ"], d["นามสกุล"], bd, c["pid"], "", "",
                         f"จับคู่จากเลข {d['เลขประจำตัว']}", cls, "", d.get("ห้อง") or ""])
        else:
            why = ("ไม่รู้รหัสชั้น" if not cls else "พบหลายคน" if len(exact) > 1 else
                   "ชื่อ/สกุลสะกดต่าง" if all(c["name"] != want for c in cands) else "วันเกิดไม่ตรง")
            review.append([d["ชั้น(ไฟล์)"], d["เลขประจำตัว"], d["คำนำหน้า"], d["ชื่อ"], d["นามสกุล"], bd, why, hint])
    out = Workbook()
    s = out.active; s.title = "สรุป"
    s["A1"], s["B1"] = "โรงเรียน", school
    s["A2"], s["B2"] = "จับคู่ได้ (แผ่น 3)", len(auto)
    s["A3"], s["B3"] = "ให้คนตัดสิน", len(review)
    s2 = out.create_sheet("2 แก้ชั้น"); s2.append(head3[:10] + ["ต้องแก้เป็น", "หมายเหตุ", "ห้อง"])
    s3 = out.create_sheet("3 เพิ่มเข้าโรงเรียน"); s3.append(head3)
    for row in auto:
        s3.append(row)
    s4 = out.create_sheet("ให้คนตัดสิน")
    s4.append(["ชั้น(ไฟล์)", "เลขในไฟล์", "คำนำหน้า", "ชื่อ", "นามสกุล", "วันเกิดในไฟล์(ค.ศ.)", "เหตุผล", "คนใน JHCIS ที่ชื่อตรง"])
    for row in review:
        s4.append(row)
    for sh in out.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True, name="Arial")
    s4.column_dimensions["H"].width = 90
    o = os.path.join(os.path.dirname(os.path.abspath(path)),
                     "จับคู่_" + os.path.basename(path).replace("ผลเทียบ_", ""))
    out.save(o)
    print(f"{school}: จับคู่ได้ {len(auto)} คน | ให้คนตัดสิน {len(review)} คน -> {o}")
    for row in review:
        print(f"  ตัดสิน: {row[0]} {row[3]} {row[4]} เกิด {row[5]}  ({row[6]})")


def cli():
    if len(sys.argv) < 2:
        sys.exit('ใช้: python jhcis_link_existing.py "ผลเทียบ_....xlsx"')
    main(sys.argv[1])


if __name__ == "__main__":
    cli()
