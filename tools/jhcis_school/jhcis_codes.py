# -*- coding: utf-8 -*-
"""
jhcis_codes.py  ** อ่านอย่างเดียว **
ดึงตารางรหัสของ JHCIS (สัญชาติ เชื้อชาติ ศาสนา คำนำหน้า) และค่าที่พบในไฟล์รายชื่อ
มาไว้ในไฟล์ "รหัส_JHCIS.xlsx" เพื่อจับคู่ค่าในไฟล์โรงเรียนกับรหัสของ JHCIS

ใช้:  python jhcis_codes.py ไฟล์รายชื่อ1.xlsx [ไฟล์รายชื่อ2.xlsx ...]
ไฟล์ผลลัพธ์มีแค่ตารางรหัสและรายการค่าที่ไม่ซ้ำ (ไม่มีชื่อ/เลขบัตรเด็ก) ส่งรูปให้ Claude ดูได้
"""
import os, re, sys, getpass, collections
import jhcis_roster_compare as rc

KEYS = ("nation", "race", "relig", "prename", "title", "prefix")


def main():
    pw = rc.ask_pw("รหัสผ่าน root ของ MySQL ในเครื่องนี้: ")
    print("ฐานข้อมูลของเครื่อง:", rc.q("SELECT @@hostname AS h", pw)[0]["h"])
    tabs = [list(r.values())[0] for r in rc.q("SHOW TABLES", pw)]
    pick = [t for t in tabs if t.lower().startswith("c") and any(k in t.lower() for k in KEYS)]
    print("ตารางรหัสที่พบ:", ", ".join(pick) or "-")

    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "ค่าในไฟล์รายชื่อ"
    ws.append(["คอลัมน์", "ค่า", "จำนวนคน"])
    for c in ws[1]:
        c.font = Font(bold=True)
    want = {"สัญชาติ": "สัญชาติ", "เชื้อชาติ": "เชื้อชาติ", "ศาสนา": "ศาสนา", "คำนำหน้าชื่อ": "คำนำหน้า", "คำนำหน้า": "คำนำหน้า"}
    from openpyxl import load_workbook
    for path in sys.argv[1:]:
        wbr = load_workbook(path, read_only=True, data_only=True)
        cnt = collections.Counter()
        for sh in wbr.worksheets:
            rows = list(sh.iter_rows(values_only=True))
            for hi, r in enumerate(rows[:15]):
                h = [str(x or "").strip() for x in r]
                if any(n in h for n in rc.ID_HEADS) and "ชื่อ" in h:
                    for name, label in want.items():
                        if name in h:
                            i = h.index(name)
                            for x in rows[hi + 1:]:
                                if x and len(x) > i and x[i]:
                                    cnt[(label, str(x[i]).strip())] += 1
                    break
            if cnt:
                break
        ws.append([os.path.basename(path)])
        for (label, v), n in sorted(cnt.items()):
            ws.append([label, v, n])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 36

    for t in pick:
        rows = rc.q(f"SELECT * FROM `{t}`", pw)
        w = wb.create_sheet(t[:31])
        if rows:
            w.append(list(rows[0].keys()))
            for c in w[1]:
                c.font = Font(bold=True)
            for r in rows:
                w.append(list(r.values()))
            for i in range(1, len(rows[0]) + 1):
                w.column_dimensions[w.cell(1, i).column_letter].width = 22
        print(f"  {t}: {len(rows)} รหัส")
    out = os.path.join(rc.home(), "รหัส_JHCIS.xlsx")
    wb.save(out)
    print("บันทึกที่:", out)


def cli():
    main()


if __name__ == "__main__":
    cli()
