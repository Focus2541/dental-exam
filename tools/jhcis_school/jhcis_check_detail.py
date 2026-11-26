"""
jhcis_check_detail.py  (อ่านอย่างเดียว)
ทำรายชื่อนักเรียนปีนี้ที่ "จะไม่ขึ้นรายงาน/ขึ้นผิด" ให้แก้ทีละคน
  - ไม่มีวันเกิด (เอาวันเกิดจากไฟล์รายชื่อมาให้ ถ้ามี)
  - สถานะจำหน่าย (dischargetype) ไม่ใช่ 9 = ยังไม่จำหน่าย

ใช้:  python jhcis_check_detail.py "รายชื่อ รร สันป่าสัก เทอม 2 ปี 69.xlsx" "รร บวกหมื้อ เทอม 2 ปี 69.xlsx" "รายชื่อ รร วัดห้วยไร่ เทอม 2 ปี 69.xlsx"
ได้:  ต้องแก้ในJHCIS_<วันที่>.xlsx
"""
import os
import sys
import time
import getpass
import jhcis_roster_compare as rc


def cli():

    srv = None
    try:
        for ln in open(r"C:\Program Files\JHCIS\database.properties", encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                srv = ln.split("=", 1)[1].strip()
    except OSError:
        pass
    if srv and srv.lower() not in ("localhost", "127.0.0.1"):
        rc.HOST = srv
    print(f"ตรวจฐาน: {rc.HOST}")

    birth_of = {}                                                 # เลขบัตร -> วันเกิดจากไฟล์ (ค.ศ.)
    for f in sys.argv[1:]:
        try:
            for r in rc.read_roster(f):
                if r.get("bd"):
                    birth_of[rc.norm_id(r["id"])] = r["bd"]
        except Exception as e:
            print(f"  อ่าน {f} ไม่ได้: {e}")
    print(f"วันเกิดจากไฟล์รายชื่อ: {len(birth_of)} คน")

    pw = rc.ask_pw(f"รหัสผ่าน root ของ MySQL ที่ {rc.HOST}: ")
    year = time.localtime().tm_year + 543 - (0 if time.localtime().tm_mon >= 5 else 1)
    rows = rc.q("SELECT v.schoolname, s.classeducate, s.classroom, p.pid, p.idcard, p.prename, p.fname, p.lname, p.birth, "
                "p.typelive, p.dischargetype, p.dischargedate FROM personstudent s "
                "JOIN person p ON p.pid=s.pid AND p.pcucodeperson=s.pcucodeperson "
                "JOIN villageschool v ON v.villcode=s.villcode AND v.schoolno=s.schoolno "
                f"WHERE s.yeareducate='{year}' AND (p.birth IS NULL OR p.dischargetype IS NULL OR p.dischargetype<>'9') "
                "ORDER BY v.schoolname, s.classeducate+0, p.fname", pw)
    names = {r["dischargetype"]: r for r in rc.q("SELECT * FROM cdischarge", pw)} if rc.q("SHOW TABLES LIKE 'cdischarge'", pw) else {}

    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook(); ws = wb.active; ws.title = "ต้องแก้"
    ws.append(["โรงเรียน", "ชั้นรหัส", "ห้อง", "HN", "เลขบัตร", "ชื่อ", "ปัญหา", "ค่าตอนนี้", "แก้เป็น (จากไฟล์)"])
    n_b = n_d = 0
    for r in rows:
        name = f"{r['prename'] or ''}{r['fname']} {r['lname']}"
        if not r["birth"]:
            n_b += 1
            bd = birth_of.get(rc.norm_id(r["idcard"]), "")
            if bd:
                y, m, d = bd.split("-")
                bd = f"{d}/{m}/{int(y) + 543}"
            ws.append([r["schoolname"], r["classeducate"], r["classroom"], r["pid"], r["idcard"], name,
                       "ไม่มีวันเกิด", "", bd or "(ไม่มีในไฟล์)"])
        if str(r["dischargetype"] or "") != "9":
            n_d += 1
            dn = names.get(r["dischargetype"], {})
            label = next((v for k, v in dn.items() if k != "dischargetype"), "") if dn else ""
            ws.append([r["schoolname"], r["classeducate"], r["classroom"], r["pid"], r["idcard"], name,
                       "สถานะจำหน่าย", f"{r['dischargetype']} {label} {r['dischargedate'] or ''}".strip(), "9 ยังไม่จำหน่าย (ถ้ายังเรียนอยู่จริง)"])
    for c in ws[1]:
        c.font = Font(bold=True, name="Arial")
    for col, w in zip("ABCDEFGHI", (14, 8, 6, 8, 16, 30, 14, 26, 26)):
        ws.column_dimensions[col].width = w
    o = os.path.join(rc.home(), f"ต้องแก้ในJHCIS_{time.strftime('%Y%m%d_%H%M')}.xlsx")
    wb.save(o)
    print(f"ไม่มีวันเกิด {n_b} คน | สถานะจำหน่ายไม่ใช่ 9: {n_d} คน -> {o}")



if __name__ == "__main__":
    cli()
