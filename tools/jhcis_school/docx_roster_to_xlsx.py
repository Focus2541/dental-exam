"""
docx_roster_to_xlsx.py  (v0.1)
แปลงรายชื่อนักเรียนแบบ Word (แบบ รร.วัดห้วยไร่: หัวข้อ "รายชื่อนักเรียนชั้น..." ตามด้วยตาราง
ที่ / เลขบัตรประจำตัวประชาชน / ว/ด/ป ที่เกิด / รหัส / ชื่อ - นามสกุล / หมายเหตุ)
เป็นไฟล์ Excel ที่ jhcis_roster_compare.py อ่านได้

ใช้:  python docx_roster_to_xlsx.py "รายชื่อ.docx"
ได้:  รายชื่อ.xlsx  (วางไว้ข้างไฟล์เดิม)

สัญชาติ: ไฟล์ Word ไม่มี จึงเดาจากเลขบัตร (ขึ้นต้น 1-5 = ไทย, อื่น ๆ = ไม่ได้สัญชาติไทย)
และเขียนไว้ในคอลัมน์ "ที่มาสัญชาติ" ให้ตรวจ/แก้ใน Excel ก่อนนำไปเทียบ
"""
import re
import jhcis_roster_compare as rc
import sys
import os

MONTHS = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4, "พ.ค.": 5, "มิ.ย.": 6,
          "ก.ค.": 7, "ส.ค.": 8, "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}
TITLES = [("เด็กชาย", "เด็กชาย"), ("เด็กหญิง", "เด็กหญิง"), ("ด.ช.", "เด็กชาย"), ("ด.ญ.", "เด็กหญิง"),
          ("นางสาว", "นางสาว"), ("น.ส.", "นางสาว"), ("นาย", "นาย")]
CLASS_RE = [(r"เตรียมอนุบาล", "เตรียมอนุบาล"), (r"อนุบาล\s*(\d)", "อ.{}"),
            (r"ประถมศึกษาปีที่\s*(\d)", "ป.{}"), (r"มัธยมศึกษาปีที่\s*(\d)", "ม.{}")]


def thai_date(s):
    """'28 พ.ค. 66' / '29  มิ.ย. 64' / '3 ก.พ. 2562' -> '28/05/2566'"""
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    s = re.sub(r"\s*\(.*?\)\s*$", "", s)                 # '14 มี.ค. 63(61)' -> ใช้วันแรก (วงเล็บแจ้งไว้ในคอลัมน์ตรวจ)
    s = re.sub(r"(\d)[^\d]+$", r"\1", s)                   # '63ฌ' -> '63'
    m = re.match(r"(\d{1,2}) ?([ก-๙\.]+?\.) ?(\d{2,4})$", s.replace(". ", "."))
    if not m:
        m = re.match(r"(\d{1,2}) ([ก-๙\.]+) (\d{2,4})$", s)
    if not m:
        return ""
    d, mon, y = m.groups()
    mon = mon if mon.endswith(".") else mon + "."
    if mon not in MONTHS:
        return ""
    y = int(y)
    y = 2500 + y if y < 100 else y
    return f"{int(d):02d}/{MONTHS[mon]:02d}/{y}"


def split_name(s):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    pre = ""
    for k, v in TITLES:
        if s.startswith(k):
            pre, s = v, s[len(k):].strip()
            break
    parts = s.split(" ", 1)
    fn = parts[0] if parts else ""
    ln = parts[1].strip() if len(parts) > 1 and parts[1].strip() else "-"
    return pre, fn, ln


def cls_of(text):
    for pat, fmt in CLASS_RE:
        m = re.search(pat, text)
        if m:
            return fmt.format(*m.groups()) if m.groups() else fmt
    return None


def main(path):
    import docx
    from docx.oxml.ns import qn
    from openpyxl import Workbook
    from openpyxl.styles import Font
    d = docx.Document(path)
    rows, cur, ti, school = [], None, 0, ""
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            t = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
            if "รายชื่อนักเรียนชั้น" in t:
                cur = cls_of(t)
            elif t.startswith("โรงเรียน") and not school:
                school = t.replace("โรงเรียน", "").split(" อ.")[0].strip()
        elif el.tag == qn("w:tbl"):
            tb = d.tables[ti]; ti += 1
            head = [c.text.strip() for c in tb.rows[0].cells]
            if not cur or not any("เลขบัตร" in h for h in head):
                continue                                       # ตารางสรุปจำนวน
            for r in tb.rows[1:]:
                c = [x.text.strip() for x in r.cells]
                if len(c) < 5 or not c[4]:
                    continue
                pid = re.sub(r"[\s\-]", "", c[1]).upper()
                pre, fn, ln = split_name(c[4])
                thai = bool(re.fullmatch(r"[1-5]\d{12}", pid))
                note = "วันเกิดมีวงเล็บ/ตัวเกิน ให้ตรวจ" if re.search(r"\(|[^\d\s\.ก-ฮะ-ฺเ-ๅ]", c[2]) or re.search(r"\d[ก-ฮ]$", c[2]) else ""
                rows.append([pid, cur, pre, fn, ln, thai_date(c[2]), c[2],
                             "ไทย" if thai else "ไม่ได้สัญชาติไทย",
                             "เดาจากเลขบัตร" if pid else "ไม่มีเลขบัตร - เดาไม่ได้", c[3], c[5] if len(c) > 5 else "", note])
    wb = Workbook(); ws = wb.active; ws.title = "รายชื่อรวม"
    ws.append(["เลขประจำตัวประชาชน", "ชั้น", "คำนำหน้าชื่อ", "ชื่อ", "นามสกุล", "วันเกิด", "วันเกิด(ในไฟล์ Word)",
               "สัญชาติ", "ที่มาสัญชาติ", "รหัสนักเรียน", "หมายเหตุ", "ชื่อโรงเรียน", "ต้องตรวจ"])
    for r in rows:
        ws.append(r[:11] + [school, r[11]])
    for c in ws[1]:
        c.font = Font(bold=True, name="Arial")
    for col, w in zip("ABCDEFGHIJKLM", (17, 12, 11, 16, 16, 12, 14, 18, 22, 11, 14, 14, 28)):
        ws.column_dimensions[col].width = w
    out = os.path.splitext(path)[0] + ".xlsx"
    wb.save(out)
    import collections
    print(f"โรงเรียน {school}: {len(rows)} คน -> {out}")
    print("  " + "  ".join(f"{k}={v}" for k, v in collections.Counter(r[1] for r in rows).items()))
    chk = [r for r in rows if r[11]]
    if chk:
        print(f"  วันเกิดต้องตรวจ {len(chk)} คน (ใช้วันแรกไว้ก่อน): " + ", ".join(f"{r[3]}({r[6]}->{r[5]})" for r in chk))
    bad = [r for r in rows if not r[5]]
    nid = [r for r in rows if not re.fullmatch(r"(\d{13}|G\d{12})", r[0])]
    if bad:
        print(f"  อ่านวันเกิดไม่ได้ {len(bad)} คน: " + ", ".join(f"{r[3]}({r[6]})" for r in bad))
    if nid:
        print(f"  ไม่มีเลขบัตร/เลขไม่ครบ {len(nid)} คน (จะถูกข้าม ต้องลงด้วยมือ): " + ", ".join(r[3] for r in nid))


def cli():
    if len(sys.argv) < 2:
        sys.exit('ใช้: python docx_roster_to_xlsx.py "รายชื่อ.docx"')
    main(sys.argv[1])


if __name__ == "__main__":
    cli()
