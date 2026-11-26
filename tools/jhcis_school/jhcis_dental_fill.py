# -*- coding: utf-8 -*-
"""
jhcis_dental_fill.py  (รุ่นทดลอง v0.7.1 ตำแหน่งนับจากมุมหน้าต่าง ทั้งตารางรายชื่อและฟอร์ม ย้ายหน้าต่างได้ + โหมด check)
กรอกฟอร์ม "การตรวจสุขภาพช่องปาก" ของ JHCIS แทนคน โดยอ่านไฟล์ Excel ที่ส่งออกจากแอป dental-exam (แผ่นงานละชั้น)

ติดตั้งครั้งแรก:   pip install pyautogui openpyxl
ขั้นที่ 1 จดตำแหน่งช่อง:   python jhcis_dental_fill.py calibrate
ถ้าจดไว้แล้วด้วยรุ่น v0.6 จดใหม่เฉพาะปุ่ม Dental 2 แถว: python jhcis_dental_fill.py calibrate-rows
ขั้นที่ 1.5 เช็กตำแหน่ง (เปิดฟอร์ม Dental ค้างไว้ เมาส์จะชี้ทีละช่อง ไม่คลิก): python jhcis_dental_fill.py check
ขั้นที่ 2 ซ้อม (ไม่กดบันทึก): python jhcis_dental_fill.py run ไฟล์.xlsx ป1 --rows 1-3
ขั้นที่ 3 บันทึกจริง:        python jhcis_dental_fill.py run ไฟล์.xlsx ป1 --rows 1-3 --save

ข้อบังคับก่อนรัน
- ทดลองกับ JHCIS ชุดทดสอบก่อนเสมอ
- เปลี่ยนแป้นพิมพ์เป็นภาษาอังกฤษ (EN) ไม่เช่นนั้นตัวเลขจะกลายเป็นอักษรไทย
- เปิดหน้า "ตรวจทันตกรรม" ของห้องนั้นค้างไว้ เลื่อนตารางไว้บนสุด อย่าย้ายหน้าต่างหลังจดตำแหน่ง
- หยุดฉุกเฉิน: ลากเมาส์ไปมุมซ้ายบนสุดของจอ
- รุ่นนี้ทำได้เฉพาะแถวที่มองเห็นในตารางโดยไม่ต้องเลื่อน
"""
import json, sys, time, argparse, os
import jhcis_roster_compare as rc

CAL = os.path.join(rc.home(), "jhcis_dental_points.json")

# (ชื่อจุด, คำอธิบายที่แสดงตอนจดตำแหน่ง)
POINTS = [
    ("row1_dental", "ปุ่ม Dental ของนักเรียน 'แถวที่ 1' ในตาราง"),
    ("row2_dental", "ปุ่ม Dental ของนักเรียน 'แถวที่ 2' ในตาราง"),
    # --- จากนี้ให้กดปุ่ม Dental เปิดฟอร์มของเด็กคนใดก็ได้ค้างไว้ ---
    ("gum", "ช่อง สุขภาพเหงือก (ตรงกลางช่อง ไม่ใช่ลูกศร)"),
    ("tartar", "ช่อง หินน้ำลายหรือไม่"),
    ("need_fluoride", "ช่อง จำเป็นต้องเคลือบฟลูออไรด์"),
    ("need_scaling", "ช่อง จำเป็นต้องขูดหินน้ำลาย"),
    ("d_have", "ช่อง จำนวนฟันน้ำนมที่มี"),
    ("d_decay", "ช่อง ฟันน้ำนมผุที่ไม่ได้อุด"),
    ("d_filled", "ช่อง ฟันน้ำนมผุที่ได้รับการอุด"),
    ("d_lost", "ช่อง ฟันน้ำนมผุที่ถอนหรือหลุด"),
    ("d_needfill", "ช่อง ฟันน้ำนมต้องรักษา-อุด"),
    ("p_have", "ช่อง จำนวนฟันแท้ที่มี"),
    ("p_decay", "ช่อง ฟันแท้ผุที่ไม่ได้อุด"),
    ("p_filled", "ช่อง ฟันแท้ผุที่ได้รับการอุด"),
    ("p_lost", "ช่อง ฟันแท้ผุที่ถอนหรือหลุด"),
    ("p_needfill", "ช่อง ฟันแท้ที่ต้องรักษา-อุด"),
    ("p_sealant", "ช่อง ฟันแท้ที่ต้องเคลือบหลุมร่องฟัน"),
    ("p_rct", "ช่อง ฟันแท้ที่ต้องถอน/รักษาคลองรากฟัน"),
    ("d_rct", "ช่อง ฟันน้ำนมที่ต้องถอน/รักษาคลองรากฟัน"),
    ("occl", "ช่อง จำนวนคู่สบฟันแท้กับฟันแท้"),
    ("cpi_all_normal", "ปุ่ม 'ปกติทั้งหมด' (สภาวะปริทันต์)"),
    ("save", "ปุ่ม 'บันทึกผลการตรวจสุขภาพช่องปาก'"),
    ("close", "ปุ่ม X ปิดฟอร์ม (มุมขวาบนของฟอร์ม)"),
]

# คอลัมน์ในไฟล์ Excel ที่ส่งออกจากแอป dental-exam (A=0) แถวข้อมูลเริ่มแถวที่ 5
# F-I ฟันน้ำนม ทั้งหมด/ผุ/อุด/ถอน, J-M ฟันแท้, N เหงือกอักเสบ, O หินปูน,
# P อุดน้ำนม, Q ถอนน้ำนม, R อุดแท้, S ถอนแท้, T Sealant
NUMS = [(5, "d_have"), (6, "d_decay"), (7, "d_filled"), (8, "d_lost"), (15, "d_needfill"),
        (9, "p_have"), (10, "p_decay"), (11, "p_filled"), (12, "p_lost"), (17, "p_needfill"),
        (19, "p_sealant"), (18, "p_rct"), (16, "d_rct")]
GUM_YES, TARTAR_YES = "1", "1"          # รหัสใน JHCIS เมื่อเหงือกอักเสบ / มีหินปูน
FLUORIDE = "1"                          # จำเป็นต้องเคลือบฟลูออไรด์: 1 = จำเป็น
SCALING_YES, SCALING_NO = "1", "2"      # ต้องขูดหินน้ำลาย: มีหินปูน -> จำเป็น อัตโนมัติ


def calibrate():
    import pyautogui
    pts = {}
    print("จดตำแหน่ง: เอาเมาส์ไปชี้ที่จุดที่บอก (ไม่ต้องคลิก) แล้วกลับมากด Enter ที่หน้าต่างนี้")
    print("ใช้ Alt+Tab สลับมาหน้าต่างนี้ได้ เมาส์จะยังชี้ที่เดิม\n")
    if not table_origin():
        sys.exit(f"ไม่พบหน้าต่างตาราง '{TABLE_KEY}' เปิดหน้าตรวจทันตกรรมของห้องก่อนแล้วเริ่มใหม่")
    for i, (key, desc) in enumerate(POINTS):
        if i == 2:
            input(">>> ตอนนี้ให้กดปุ่ม Dental เปิดฟอร์มของเด็ก 1 คนค้างไว้ แล้วกด Enter")
            if not form_origin():
                sys.exit(f"ไม่พบหน้าต่าง '{FORM_KEY}' เปิดฟอร์มก่อนแล้วเริ่มใหม่")
        input(f"[{i+1}/{len(POINTS)}] ชี้ที่: {desc}  แล้วกด Enter ")
        x, y = pyautogui.position()
        if i >= 2:                      # ช่องในฟอร์ม: เก็บระยะจากมุมซ้ายบนของฟอร์ม
            ox, oy = form_origin()
            pts[key] = [x - ox, y - oy]
            print(f"      = ({x}, {y})  ห่างจากมุมฟอร์ม ({x-ox}, {y-oy})")
        else:                           # ปุ่ม Dental: เก็บระยะจากมุมหน้าต่างตาราง
            ox, oy = table_origin()
            pts[key] = [x - ox, y - oy]
            print(f"      = ({x}, {y})  ห่างจากมุมตาราง ({x-ox}, {y-oy})")
    pts["_form_relative"] = True
    pts["_rows_relative"] = True
    with open(CAL, "w", encoding="utf-8") as f:
        json.dump(pts, f, indent=1)
    print("\nบันทึกตำแหน่งแล้วที่", CAL)


def num(v):
    if v is None or str(v).strip() == "":
        return "0"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = str(v).strip()
    if not t.isdigit():
        raise ValueError(f"ค่า '{t}' ไม่ใช่ตัวเลข")
    return t


def mark(v):
    return v is not None and str(v).strip() not in ("", "0")


def load_rows(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"ไม่พบแผ่นงาน '{sheet}' ในไฟล์ มีแผ่นงาน: {', '.join(wb.sheetnames)}")
    ws = wb[sheet]
    if "แท้" not in str(ws.cell(4, 18).value or ""):
        sys.exit("ไฟล์นี้เป็นแบบฟอร์มรุ่นเก่า (Tx. plan 3 ช่อง) ให้ส่งออกจากแอปรุ่นใหม่ที่แยกน้ำนม/แท้แล้ว")
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        r = list(r) + [None] * 20
        if not str(r[3] or "").strip():
            continue
        name = f"{r[2] or ''}{r[3]} {r[4] or ''}".strip()
        exam = any(str(x).strip() != "" for x in r[5:13] if x is not None)
        d = {"name": name, "fn": str(r[3]).strip(), "exam": exam, "err": None, "combos": [], "nums": []}
        try:
            tartar = mark(r[14])
            d["combos"] = [("gum", GUM_YES if mark(r[13]) else "0"), ("tartar", TARTAR_YES if tartar else "0"),
                           ("need_fluoride", FLUORIDE), ("need_scaling", SCALING_YES if tartar else SCALING_NO)]
            d["nums"] = [(k, num(r[c])) for c, k in NUMS] + [("occl", "0")]
        except ValueError as e:
            d["err"] = str(e)
        out.append(d)
    return out


FORM_KEY = "การตรวจสุขภาพช่องปาก"
TABLE_KEY = "ตรวจทันตกรรม"   # หน้าต่างตารางรายชื่อ "บันทึกผลงานอนามัยโรงเรียน ตรวจทันตกรรม วันที่ : ..."


def active_title():
    import pygetwindow
    w = pygetwindow.getActiveWindow()
    return (w.title if w else "") or ""


def wait_form(timeout=4.0):
    """รอจนหน้าต่างที่อยู่ด้านหน้าเป็นฟอร์มตรวจช่องปาก คืนชื่อหน้าต่าง หรือ None ถ้าไม่ขึ้น"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        t = active_title()
        if FORM_KEY in t:
            return t
        time.sleep(0.2)
    return None


def form_origin():
    """มุมซ้ายบนของฟอร์มตรวจช่องปาก (หน้าจอจริง) หรือ None ถ้าไม่ได้เปิดอยู่"""
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if FORM_KEY in (w.title or "")]
    return (ws[0].left, ws[0].top) if ws else None


def table_origin():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if TABLE_KEY in (w.title or "")]
    return (ws[0].left, ws[0].top) if ws else None


def calibrate_rows():
    """จดใหม่เฉพาะปุ่ม Dental แถว 1-2 (นับจากมุมหน้าต่างตาราง) เก็บตำแหน่งช่องในฟอร์มไว้เหมือนเดิม"""
    import pyautogui
    if not os.path.exists(CAL):
        sys.exit("ยังไม่มีไฟล์ตำแหน่ง ให้รัน calibrate")
    P = json.load(open(CAL, encoding="utf-8"))
    if not P.get("_form_relative"):
        sys.exit("ไฟล์ตำแหน่งเป็นรุ่นเก่ามาก ให้รัน calibrate ใหม่ทั้งหมด")
    if not table_origin():
        sys.exit(f"ไม่พบหน้าต่างตาราง '{TABLE_KEY}' เปิดหน้าตรวจทันตกรรมของห้องก่อน")
    for key, desc in POINTS[:2]:
        input(f"ชี้ที่: {desc}  แล้วกด Enter ")
        x, y = pyautogui.position()
        ox, oy = table_origin()
        P[key] = [x - ox, y - oy]
        print(f"      = ({x}, {y})  ห่างจากมุมตาราง ({x-ox}, {y-oy})")
    P["_rows_relative"] = True
    json.dump(P, open(CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว")


def load_points():
    if not os.path.exists(CAL):
        sys.exit("ยังไม่ได้จดตำแหน่ง ให้รันคำสั่ง calibrate ก่อน")
    P = json.load(open(CAL, encoding="utf-8"))
    if not P.get("_form_relative"):
        sys.exit("ไฟล์ตำแหน่งเป็นรุ่นเก่า (เก็บพิกัดตายตัว) ให้รัน calibrate ใหม่อีกครั้ง")
    if not P.get("_rows_relative"):
        sys.exit("ตำแหน่งปุ่ม Dental ยังเป็นแบบเก่า ให้รัน calibrate-rows (จดใหม่แค่ 2 จุด)")
    return P


def fp(P, key):
    """ตำแหน่งจริงบนจอของช่องในฟอร์ม = มุมฟอร์มตอนนี้ + ระยะที่จดไว้"""
    o = form_origin()
    if not o:
        raise RuntimeError("ฟอร์มปิดไปแล้ว")
    return (o[0] + P[key][0], o[1] + P[key][1])


def check():
    import pyautogui
    P = load_points()
    if not form_origin():
        sys.exit(f"เปิดฟอร์ม '{FORM_KEY}' ของเด็ก 1 คนค้างไว้ก่อน แล้วรัน check ใหม่")
    print("เมาส์จะไปชี้ทีละช่อง (ไม่คลิก) ดูว่าปลายเมาส์อยู่ในช่องทุกครั้ง ถ้าไม่ตรง ให้ calibrate ใหม่")
    input("กด Enter เพื่อเริ่ม แล้วอย่าขยับเมาส์")
    for key, desc in POINTS[2:]:
        x, y = fp(P, key)
        pyautogui.moveTo(x, y, duration=0.3)
        print(f"  {desc}")
        time.sleep(1.2)
    print("เสร็จ")


def focus_form():
    """ดึงฟอร์มตรวจช่องปากกลับมาไว้หน้าสุด ก่อนคลิกปุ่มบันทึก/ปิด
    (หลังหน้าต่างถามยืนยันปิด หน้าต่างอื่นอาจขึ้นมาบังตำแหน่งปุ่ม)"""
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if FORM_KEY in (w.title or "")]
    if not ws:
        return False
    try:
        ws[0].activate()
    except Exception:
        pass
    time.sleep(0.4)
    return FORM_KEY in active_title()


def form_closed(timeout=3.0):
    t0 = time.time()
    while time.time() - t0 < timeout:
        if FORM_KEY not in active_title():
            return True
        time.sleep(0.2)
    return False


def ask(msg, buttons):
    """หน้าต่างถามแบบมีปุ่ม ลอยอยู่บนสุด กดด้วยเมาส์ ไม่ต้องสลับไปหน้าต่างดำ
    (ป้องกันกด Enter แล้วไปโดน JHCIS แทน) ปิดหน้าต่างโดยไม่เลือก = หยุด"""
    import pyautogui
    print(msg)
    return pyautogui.confirm(text=msg, title="กรอก JHCIS อัตโนมัติ", buttons=buttons)


def parse_range(s, n):
    if not s:
        return list(range(1, n + 1))
    a, _, b = s.partition("-")
    return list(range(int(a), int(b or a) + 1))


def run(path, sheet, rows_arg, do_save, delay):
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = delay
    P = load_points()
    rows = load_rows(path, sheet)
    want = [n for n in parse_range(rows_arg, len(rows)) if 1 <= n <= len(rows)]
    if not want:
        sys.exit("ไม่มีแถวให้ทำ")
    row_h = P["row2_dental"][1] - P["row1_dental"][1]
    mode = "บันทึกจริง" if do_save else "ซ้อม (ไม่กดบันทึก)"
    print(f"โหมด: {mode} | แผ่นงาน {sheet} มี {len(rows)} คน | จะทำแถว {want[0]}-{want[-1]}")
    input("ตรวจว่าแป้นพิมพ์เป็น EN และตารางอยู่บนสุด แล้วกด Enter เพื่อเริ่ม (มีเวลา 3 วินาทีสลับไป JHCIS)")
    time.sleep(3)
    log = []
    for n in want:
        r = rows[n - 1]
        name = r["name"]
        if not r["exam"]:
            print(f"แถว {n}: {name} -> ยังไม่มีผลตรวจ ข้าม")
            log.append((n, name, "ข้าม-ไม่มีผลตรวจ")); continue
        if r["err"]:
            print(f"แถว {n}: {name} -> ข้อมูลผิด: {r['err']} ข้าม")
            log.append((n, name, "ข้าม-" + r["err"])); continue
        to = table_origin()
        if not to:
            print(f"ไม่พบหน้าต่างตาราง '{TABLE_KEY}' -> หยุด")
            log.append((n, name, "หยุด-ไม่พบตาราง")); break
        x, y = to[0] + P["row1_dental"][0], to[1] + P["row1_dental"][1]
        pyautogui.click(x, y + row_h * (n - 1))
        title = wait_form()
        if not title:
            print(f"แถว {n}: หลังคลิกปุ่ม Dental ไม่พบฟอร์ม '{FORM_KEY}' -> หยุดทันที ไม่พิมพ์อะไรลงไป")
            print("   หน้าต่างที่อยู่ด้านหน้าตอนนี้คือ:", active_title() or "(ไม่ทราบ)")
            log.append((n, name, "หยุด-ไม่พบฟอร์ม")); break
        who = title.split("ของ :", 1)[-1].strip(" :")
        if r["fn"] and r["fn"] not in who:
            ans = ask(f"แถว {n}: ชื่อไม่ตรง!\n\nในไฟล์ Excel: {name}\nฟอร์มที่เปิด: {who}",
                      ["ปิดฟอร์มแล้วหยุด", "กรอกต่อ (ซ้อมเท่านั้น)"])
            if ans != "กรอกต่อ (ซ้อมเท่านั้น)":
                pyautogui.click(*fp(P, "close"))
                log.append((n, name, f"หยุด-ชื่อไม่ตรง ({who})")); break
        for k, v in r["combos"]:      # ช่องตัวเลือก: คลิก พิมพ์เลขรหัส แล้ว Enter
            pyautogui.click(*fp(P, k)); pyautogui.write(v); pyautogui.press("enter")
        for k, v in r["nums"]:        # ช่องตัวเลข: คลิก ล้างค่าเดิม แล้วพิมพ์
            pyautogui.click(*fp(P, k)); pyautogui.hotkey("ctrl", "a"); pyautogui.write(v)
        pyautogui.click(*fp(P, "cpi_all_normal"))
        ok_btn = "บันทึก" if do_save else "ปิดโดยไม่บันทึก (ซ้อม)"
        ans = ask(f"แถว {n} / {want[-1]}: {name}\n\nกรอกเสร็จแล้ว ดูค่าในฟอร์ม JHCIS ให้ถูกต้องก่อนกด",
                  [ok_btn, "ข้ามคนนี้", "หยุด"])
        time.sleep(0.5)
        if not focus_form():
            print(f"แถว {n}: ดึงฟอร์มขึ้นหน้าสุดไม่ได้ -> หยุด ไม่กดบันทึก ตรวจหน้าจอด้วยตัวเอง")
            log.append((n, name, "หยุด-ฟอร์มไม่อยู่หน้าสุด")); break
        if ans in (None, "หยุด"):
            pyautogui.click(*fp(P, "close"))
            log.append((n, name, "หยุดโดยผู้ใช้")); break
        if do_save and ans == ok_btn:
            pyautogui.click(*fp(P, "save")); time.sleep(1.5)
            log.append((n, name, "กดบันทึก"))
        else:
            log.append((n, name, f"ไม่บันทึก (ปุ่มที่กด: {ans})"))
        if FORM_KEY in active_title():
            pyautogui.click(*fp(P, "close"))
        if not form_closed():
            print(f"แถว {n}: ปิดฟอร์มไม่สำเร็จ -> หยุด ตรวจหน้าจอด้วยตัวเอง")
            break
        time.sleep(0.5)
    with open(os.path.join(rc.home(), "jhcis_dental_log.txt"), "a", encoding="utf-8") as f:
        for n, name, st in log:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{os.path.basename(path)}\t{sheet}\t{n}\t{name}\t{st}\n")
    print("เสร็จ ดูบันทึกที่ jhcis_dental_log.txt")


def cli():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("calibrate")
    sub.add_parser("check")
    sub.add_parser("calibrate-rows")
    r = sub.add_parser("run")
    r.add_argument("xlsx")
    r.add_argument("sheet", help="ชื่อแผ่นงาน (ชั้น) เช่น ป1")
    r.add_argument("--rows", help="เช่น 1-3")
    r.add_argument("--save", action="store_true", help="กดปุ่มบันทึกจริง")
    r.add_argument("--delay", type=float, default=0.25)
    a = ap.parse_args()
    if a.cmd == "calibrate":
        calibrate()
    elif a.cmd == "check":
        check()
    elif a.cmd == "calibrate-rows":
        calibrate_rows()
    else:
        run(a.xlsx, a.sheet, a.rows, a.save, a.delay)


if __name__ == "__main__":
    cli()
