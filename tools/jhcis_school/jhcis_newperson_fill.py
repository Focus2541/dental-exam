# -*- coding: utf-8 -*-
"""
jhcis_newperson_fill.py  (รุ่นทดลอง v0.8.9 เช็กว่าช่องข้อความมีตัวอักษรจริง เลือกหมู่บ้าน/บ้านนอกเขตเองทุกคน คำนำหน้า/เพศ/โรงเรียน/ชั้น เลือกด้วยเมาส์)
เพิ่มเด็กที่ยังไม่มีใน JHCIS (แผ่น "4 ไม่มีใน JHCIS" ของไฟล์ผลเทียบ) เป็นคนนอกเขต แล้วใส่โรงเรียน/ชั้นให้
ทำตามขั้นตอนเดียวกับที่ทำด้วยมือ: เพิ่มคนใหม่ -> ข้อมูลบุคคล -> แท็บนักเรียน -> SAVE -> แผ่นดิสก์

ต้องมี jhcis_roster_compare.py (v0.3) ในโฟลเดอร์เดียวกัน   ติดตั้งเพิ่ม: pip install pyperclip

1) จดตำแหน่ง (ครั้งเดียว):  python jhcis_newperson_fill.py calibrate
   จดเพิ่มเฉพาะส่วนรายการ (ถ้าเคยจดด้วยรุ่น v0.1-v0.2 แล้ว):  python jhcis_newperson_fill.py calibrate-lists
2) ซ้อม ไม่บันทึก:          python jhcis_newperson_fill.py run "ผลเทียบ_xxx.xlsx" --school-pos N --limit 1
3) บันทึกจริง:              เติม --save          (ไม่ถามทีละคน: เติม --auto)

ก่อนรัน: เลือก หมู่บ้าน "00-หมู่บ้านนอกเขต หมู่ 0" และ บ้านเลขที่ "0" ด้านบนของ JHCIS ไว้ก่อน
         แป้นพิมพ์เป็น ENG   หยุดฉุกเฉิน: ลากเมาส์ไปมุมซ้ายบนสุดของจอ

ความปลอดภัย
- ปกติทำงานเฉพาะเมื่อ JHCIS ต่อกับฐานทดสอบในเครื่อง (SERVER=localhost) ฐานจริงต้องเติม --real --save
- ข้อควรรู้: JHCIS สร้างคนในฐานข้อมูลทันทีที่กด 'เพิ่มคนใหม่' (ยังไม่ต้องกดบันทึก)
  โหมดซ้อมจึงสร้างคนค้างไว้ ใช้ได้เฉพาะฐานทดสอบ
- ก่อนสร้างแต่ละคน เช็กว่าเลขบัตรยังไม่มีในฐานข้อมูล ถ้ามีแล้วข้าม (รันซ้ำได้ ไม่สร้างคนซ้ำ)
- ข้ามให้คนทำเอง: เลขขึ้นต้นด้วย G / เลขไม่ครบ 13 หลัก / ไม่มีวันเกิด / มีคนชื่อ+วันเกิดตรงใน JHCIS / รหัสสัญชาติว่าง
- หลังบันทึกแต่ละคน อ่านฐานข้อมูลตรวจ: มีเลขบัตรนี้ 1 คน, ชื่อ สกุล วันเกิด ประเภทการอยู่ 4,
  สัญชาติ เชื้อชาติ ศาสนา คำนำหน้า เพศ, โรงเรียน ชั้น ห้อง ปี  ถ้าไม่ตรงหยุดทันที
"""
import os, re, sys, json, time, argparse, getpass
import jhcis_roster_compare as rc

HERE = rc.home()
CAL = os.path.join(HERE, "jhcis_newperson_points.json")
PROP = r"C:\Program Files\JHCIS\database.properties"
WIN_KEY = "JHCIS V"

# ลำดับในรายการของ JHCIS (นับจากบน เริ่ม 1) ดูจากหน้าจอจริง 17 พ.ย. 2569
TITLE_POS = {"เด็กชาย": (2, "001", 1), "ด.ช.": (2, "001", 1), "เด็กหญิง": (3, "002", 2), "ด.ญ.": (3, "002", 2),
             "นาย": (4, "003", 1), "นางสาว": (5, "004", 2), "น.ส.": (5, "004", 2)}      # ลำดับ, รหัส, เพศ
NATION_POS = {"099": (2, "099"), "048": (3, "048"), "999": (8, "000")}   # มีแถวว่างบนสุด   # 999 ในไฟล์ผลเทียบ -> เลือก "000-ไม่ระบุ"
RELIGION_POS = {"พุทธ": (3, "01"), "อิสลาม": (4, "02"), "คริสต์": (5, "03"), "อื่นๆ": (12, "99"), "อื่น ๆ": (12, "99")}

POINTS = [
    ("new", "ปุ่ม 'เพิ่มคนใหม่ [User keys]' มุมซ้ายล่าง"),
    ("tab_person", "แท็บ 'ข้อมูลบุคคล'"),
    ("title", "ช่อง 'คำนำหน้า' (กลางช่อง ไม่ใช่ลูกศร)"),
    ("fname", "ช่อง 'ชื่อ'"),
    ("lname", "ช่อง 'สกุล'"),
    ("idcard", "ช่อง 'เลขบัตร ปชช.' ในข้อมูลบุคคล (ไม่ใช่ช่องค้นหาด้านบน)"),
    ("birth", "ช่อง 'ว/ด/ป เกิด'"),
    ("nation", "ช่อง 'สัญชาติ' (กลางช่อง)"),
    ("race", "ช่อง 'เชื้อชาติ' (กลางช่อง)"),
    ("religion", "ช่อง 'ศาสนา' (กลางช่อง)"),
    ("tab_student", "แท็บ 'ตำแหน่งในชุมชน & นักเรียน & ตำแหน่งทางศาสนา'"),
    ("school", "ช่อง 'โรงเรียน' (กลางช่อง)"),
    ("cls", "ช่อง 'ชั้นเรียน' (กลางช่อง)"),
    ("room", "ช่อง 'ห้อง'"),
    ("year", "ช่อง 'ปีการศึกษา'"),
    ("save", "ปุ่ม 'บันทึก ( SAVE )' ของสถานะนักเรียน (ปุ่มบน)"),
    ("disk", "ปุ่มรูปแผ่นดิสก์ มุมซ้ายล่าง"),
]


def origin():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if (w.title or "").startswith(WIN_KEY)]
    return (ws[0].left, ws[0].top) if ws else None


def calibrate():
    import pyautogui
    if not origin():
        sys.exit("ไม่พบหน้าต่าง JHCIS")
    print("ชี้เมาส์ (ไม่คลิก) แล้วกด Enter ที่หน้าต่างนี้   **ห้ามกดบันทึกใน JHCIS ระหว่างจดตำแหน่ง**")
    pts = {}
    for i, (k, d) in enumerate(POINTS):
        if k == "tab_person":
            input(">>> กด 'เพิ่มคนใหม่' ใน JHCIS ให้ได้ฟอร์มเปล่า แล้วกลับมากด Enter ที่นี่")
        if k == "school":
            input(">>> คลิกแท็บนักเรียนใน JHCIS ค้างไว้ แล้วกลับมากด Enter ที่นี่")
        input(f"[{i+1}/{len(POINTS)}] ชี้ที่: {d}  แล้วกด Enter ")
        x, y = pyautogui.position()
        ox, oy = origin()
        pts[k] = [x - ox, y - oy]
        print(f"      = ({x}, {y})")
    json.dump(pts, open(CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกตำแหน่งหลักแล้ว  ต่อด้วยส่วนรายการ:")
    calibrate_lists()


LIST_POINTS = [
    ("gender", "ช่อง 'เพศ' ในแท็บข้อมูลบุคคล (กลางช่อง)", "person"),
    ("li_a", "เปิดรายการ 'คำนำหน้า' ค้างไว้ แล้วชี้ที่แถว 'ด.ช.(001)'", "person"),
    ("li_b", "เปิดรายการ 'คำนำหน้า' อีกครั้ง แล้วชี้ที่แถว 'นาย(003)'", "person"),
    ("cls_last", "แท็บนักเรียน (เลือกโรงเรียนบ้านบวกหมื้อไว้) เปิดรายการ 'ชั้นเรียน' ห้ามเลื่อน แล้วชี้ที่แถวที่ 8 = แถวล่างสุดที่มองเห็นในรายการ", "student"),
    ("vill", "ช่อง 'หมู่บ้าน' มุมขวาบนของ JHCIS (กลางช่อง)", "top"),
    ("house", "ช่อง 'บ้านเลขที่' ใต้ช่องหมู่บ้าน (กลางช่อง)", "top"),
    ("hnbox", "ตัวเลข HN ในช่องสีเขียว 'HN:' ของแท็บ 'ข้อมูลบุคคล' (เปิดแท็บข้อมูลบุคคลไว้ก่อน) แค่ชี้ ห้ามคลิก", "top"),
    ("del", "ปุ่ม 'ลบ' แถวปุ่มล่างสุดของ JHCIS (ระหว่าง 'รับบริการ' กับ 'ตาย../ไม่ตาย') แค่ชี้ ห้ามคลิก", "top"),
    ("cls_down", "เปิดรายการ 'ชั้นเรียน' อีกครั้ง แล้วชี้ที่ปุ่ม ▼ เล็ก ๆ ที่มุมขวาล่างของ 'รายการ' (ปลายแถบเลื่อน ไม่ใช่ลูกศรของช่อง)", "student"),
]


def calibrate_lists(only=None):
    import pyautogui
    if not os.path.exists(CAL):
        sys.exit("ยังไม่มีไฟล์ตำแหน่ง ให้รัน calibrate ก่อน")
    P = json.load(open(CAL, encoding="utf-8"))
    print("วิธีจด: อ่านคำสั่ง แล้วกด Enter ที่หน้าต่างนี้ จากนั้นมีเวลา 8 วินาที ไปทำใน JHCIS แล้ว 'วางเมาส์นิ่ง' บนจุดที่บอก")
    print("สคริปต์จะจับตำแหน่งเองเมื่อครบเวลา (ไม่ต้องกด Enter ตอนรายการเปิด)   ห้ามบันทึกใน JHCIS\n")
    pts = [p for p in LIST_POINTS if not only or p[0] in only]
    for i, (k, d, _) in enumerate(pts):
        input(f"[{i+1}/{len(pts)}] {d}\n      พร้อมแล้วกด Enter ")
        for sec in range(8, 0, -1):
            print(f"      จับตำแหน่งใน {sec} ...", end="\r", flush=True)
            time.sleep(1)
        x, y = pyautogui.position()
        ox, oy = origin()
        P[k] = [x - ox, y - oy]
        print(f"      = ({x}, {y})")
    if "title" in P and "li_a" in P and "li_b" in P and "cls_last" in P:
        row, first, vis = list_geom(P)
        print(f"ตรวจ: สูงแถว {row:.1f} px, มองเห็น {vis} แถว (ควรเป็น 8), ลูกศรลงอยู่ใต้ช่องชั้นเรียน {P['cls_down'][1] - P['cls'][1]} px (ต้องเป็นบวก)")
        if vis != 8 or P["cls_down"][1] <= P["cls"][1]:
            print("*** ค่าไม่ตรงที่ควรเป็น ให้จดใหม่")
    json.dump(P, open(CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว  ปิดรายการ (Esc)")


def list_geom(P):
    """ระยะแถวในรายการ: row = ความสูงแถว, first = ระยะจากกลางช่องถึงกลางแถวแรก (คำนวณจากรายการคำนำหน้า)"""
    row = (P["li_b"][1] - P["li_a"][1]) / 2.0
    first = P["li_a"][1] - P["title"][1] - row          # ด.ช. เป็นแถวที่ 2
    vis = int(round((P["cls_last"][1] - P["cls"][1] - first) / row)) + 1
    return row, first, vis


def jhcis_server():
    try:
        for ln in open(PROP, encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                return ln.split("=", 1)[1].strip()
    except OSError:
        return None


def be_date(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}/{int(y) + 543}"


def read_rows(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    school = wb["สรุป"]["B1"].value
    ws = wb["4 ไม่มีใน JHCIS"]
    h = [c.value for c in ws[1]]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(h, r))
        t = {"cls": str(d.get("ชั้น(ไฟล์)") or ""), "id": str(d.get("เลขประจำตัว") or "").strip(),
             "pre": str(d.get("คำนำหน้า") or "").strip(), "fn": str(d.get("ชื่อ") or "").strip(),
             "ln": str(d.get("นามสกุล") or "").strip() or "-", "bd": str(d.get("วันเกิด(ค.ศ.)") or ""),
             "room": str(d.get("ห้อง") or "").strip(), "nat": str(d.get("รหัสสัญชาติ") or ""),
             "race": str(d.get("รหัสเชื้อชาติ") or ""), "rel": str(d.get("ศาสนา") or "").strip(),
             "hint": str(d.get("คนใน JHCIS ที่ชื่อตรง (ให้คนตัดสิน)") or d.get("คนใน JHCIS ที่ชื่อ+วันเกิดตรง (ให้คนตัดสิน)") or "").strip()}
        t["name"] = f"{t['pre']}{t['fn']} {t['ln']}"
        why = []
        if not re.fullmatch(r"\d{13}|[GN]\d{12}", t["id"]):     # เลข G/N (เด็กไม่มีบัตร) ช่องเลขบัตรของ JHCIS รับได้
            why.append("เลขบัตรไม่ใช่ตัวเลข 13 หลัก หรือ G/N+12 หลัก")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", t["bd"]):
            why.append("ไม่มีวันเกิด")
        if t["hint"]:
            why.append("มีคนชื่อ+วันเกิดตรงใน JHCIS")
        if t["pre"] not in TITLE_POS:
            why.append(f"คำนำหน้า '{t['pre']}'")
        if t["nat"] not in NATION_POS or t["race"] not in NATION_POS:
            why.append("รหัสสัญชาติ/เชื้อชาติว่าง")
        if t["rel"] and t["rel"] not in RELIGION_POS:
            why.append(f"ศาสนา '{t['rel']}'")
        t["skip"] = ", ".join(why)
        out.append(t)
    return school, out


def leftovers(a):
    """อ่านอย่างเดียว: หาคนที่ถูกสร้างค้างไว้ (นอกเขต typelive 4 ที่ชื่อ/เลขบัตรว่าง หรือยังไม่มีสถานะนักเรียน)"""
    srv = jhcis_server() or "localhost"
    local = srv.lower() in ("localhost", "127.0.0.1")
    rc.HOST = "127.0.0.1" if local else srv
    rc.USER = a.user
    print(f"JHCIS ต่อกับ: {srv} -> " + ("ฐานทดสอบในเครื่อง" if local else "*** ฐานข้อมูลจริง (อ่านอย่างเดียว) ***"))
    pw = rc.ask_pw(f"รหัสผ่าน {a.user} ของ MySQL ที่ {rc.HOST}: ")
    rs = rc.q("SELECT p.pid,p.fname,p.lname,p.idcard,s.schoolno,s.classeducate,s.classroom,s.yeareducate "
              "FROM person p LEFT JOIN personstudent s ON s.pid=p.pid AND s.pcucodeperson=p.pcucodeperson "
              "WHERE p.typelive='4' AND (p.fname IS NULL OR TRIM(p.fname)='' OR p.idcard IS NULL OR TRIM(p.idcard)='' OR p.birth IS NULL "
              "OR s.pid IS NULL OR s.classroom IS NULL OR TRIM(s.classroom)='') ORDER BY p.pid DESC LIMIT 40", pw)
    if not rs:
        print("ไม่พบคนค้าง ✓")
        return
    print(f"พบ {len(rs)} คน (ใหม่สุดก่อน) ตรวจทีละ HN ใน JHCIS: ถ้าว่างหรือเป็นคนทดสอบ -> ลบ / ถ้าข้อมูลครบแต่ไม่มีชั้น -> ใส่ชั้นแล้วบันทึก")
    for r in rs:
        st = f"รร {r['schoolno']} ชั้น {r['classeducate']} ห้อง {r['classroom'] or '-'} ปี {r['yeareducate']}" if r["schoolno"] else "ไม่มีสถานะนักเรียน"
        print(f"  HN {r['pid']:>6}  {(r['fname'] or '(ไม่มีชื่อ)')} {r['lname'] or ''}  บัตร {r['idcard'] or '-'}  | {st}")
    print("หมายเหตุ: คนนอกเขตที่ไม่ใช่นักเรียน (เช่น คนไข้ทั่วไป) ก็อาจติดรายการนี้ได้ ห้ามลบถ้าไม่แน่ใจ")


def run(a):
    import pyautogui, pyperclip, pygetwindow
    print("jhcis_newperson_fill รุ่น v0.8.9")
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = a.delay
    if not os.path.exists(CAL):
        sys.exit("ยังไม่ได้จดตำแหน่ง ให้รัน calibrate ก่อน")
    P = json.load(open(CAL, encoding="utf-8"))
    miss = [k for k, _, _ in LIST_POINTS if k not in P]
    if miss:
        sys.exit("ยังไม่ได้จดตำแหน่งบางจุด: " + ", ".join(miss) +
                 "\n  ถ้าขาดแค่ vill/house ให้รัน  python jhcis_newperson_fill.py calibrate-vill")
    ROW, FIRST, VIS = list_geom(P)
    if not (8 <= ROW <= 60 and VIS >= 3):
        sys.exit(f"ตำแหน่งรายการดูผิดปกติ (สูงแถว {ROW:.1f}, มองเห็น {VIS} แถว) ให้รัน calibrate-lists ใหม่")
    srv = jhcis_server()
    if srv is None:
        sys.exit("อ่านไฟล์ตั้งค่า JHCIS ไม่ได้: " + PROP)
    local = srv.lower() in ("localhost", "127.0.0.1")
    if not local and not a.real:
        sys.exit(f"JHCIS ต่อกับ {srv} (ฐานจริง) สคริปต์จะไม่ทำงาน  สลับเป็น localhost หรือเติม --real เมื่อพร้อมใช้จริง")
    if not local and not a.save:
        sys.exit("ห้ามซ้อมบนฐานจริง: JHCIS สร้างคนทันทีที่กด 'เพิ่มคนใหม่' โหมดซ้อมจึงทิ้งคนค้างไว้ในฐานจริง")
    rc.HOST = "127.0.0.1" if local else srv
    rc.USER = a.user
    print("=" * 60 + f"\nJHCIS ต่อกับ: {srv} -> " + ("ฐานทดสอบในเครื่อง" if local else "*** ฐานข้อมูลจริง ***") + "\n" + "=" * 60)
    pw = rc.ask_pw(f"รหัสผ่าน {a.user} ของ MySQL ที่ {rc.HOST}: ")
    stale = rc.q("SELECT pid FROM person WHERE typelive='4' AND (fname IS NULL OR TRIM(fname)='' OR idcard IS NULL OR TRIM(idcard)='' OR birth IS NULL) "
                 "AND pid > (SELECT MAX(pid) - 300 FROM person p2) ORDER BY pid", pw)
    if stale and not a.ignore_leftovers:
        sys.exit("มีคนนอกเขตที่สร้างค้างไว้ (ไม่มีชื่อ/เลขบัตร/วันเกิด): HN " + ", ".join(r["pid"] for r in stale) +
                 "\nลบใน JHCIS ก่อน (ดูรายละเอียด: python jhcis_newperson_fill.py leftovers) แล้วค่อยรัน")

    school, rows = read_rows(a.xlsx)
    sc = [s for s in rc.q("SELECT villcode,schoolno,schoolname FROM villageschool", pw) if s["schoolname"] == school]
    if len(sc) != 1:
        sys.exit(f"หาโรงเรียน '{school}' ในฐานข้อมูลไม่เจอแบบชัดเจน")
    vill, sno = sc[0]["villcode"], sc[0]["schoolno"]
    crow = rc.q("SELECT * FROM cschoolclass", pw)
    ck = list(crow[0].keys())
    cls_code = {rc.norm_class(r[ck[1]]): r[ck[0]] for r in crow if rc.norm_class(r[ck[1]])}
    pcols = {r["Field"] for r in rc.q("SHOW COLUMNS FROM person", pw)}
    col = lambda *names: next((n for n in names if n in pcols), None)
    C = {"nation": col("nation", "nationality"), "race": col("origin", "race"), "religion": col("religion"),
         "prename": col("prename", "title"), "sex": col("sex")}
    print("คอลัมน์ที่ใช้ตรวจ:", ", ".join(f"{k}={v}" for k, v in C.items() if v), "| ไม่พบ:", ", ".join(k for k, v in C.items() if not v) or "-")

    def at(k):
        o = origin()
        if not o:
            raise RuntimeError("ไม่พบหน้าต่าง JHCIS")
        return (o[0] + P[k][0], o[1] + P[k][1])

    def pick(k, pos):
        pyautogui.click(*at(k)); time.sleep(0.3)
        # กด down 1 ครั้งก่อน: ช่องที่ว่างอยู่ (ไม่มีค่าเลือก) จะได้มีค่าเลือกก่อน แล้วค่อยขึ้นไปบนสุด
        pyautogui.press("down"); pyautogui.press("up", presses=500)
        pyautogui.press("down", presses=pos - 1); pyautogui.press("enter")
        time.sleep(0.2)

    def mpick(k, pos, home=False):
        """เลือกด้วยเมาส์: เปิดรายการ แล้วคลิกแถวที่ pos (ใช้กับช่องที่ปุ่มลูกศรใช้ไม่ได้ และยังว่างอยู่)"""
        x, y = at(k)
        pyautogui.press("esc"); time.sleep(0.2)            # ปิดรายการที่อาจเปิดค้าง (ไม่งั้นคลิกแล้วรายการปิดแทนเปิด)
        pyautogui.click(x, y); time.sleep(0.5)
        if home:                                           # ช่องมีค่าแล้ว รายการจะเปิดที่ค่านั้น กลับขึ้นบนสุดก่อน
            pyautogui.press("home"); time.sleep(0.3)
        if pos > VIS:
            dx, dy = at("cls_down")
            for _ in range(pos - VIS):
                pyautogui.click(dx, dy); time.sleep(0.15)
            ty = y + FIRST + (VIS - 1) * ROW
        else:
            ty = y + FIRST + (pos - 1) * ROW
        pyautogui.click(x, int(round(ty))); time.sleep(0.3)

    def has_text(k):
        """ดูภาพหน้าจอตรงกลางช่อง: มีจุดสีเข้ม (ตัวอักษร) มากพอหรือไม่"""
        x, y = at(k)
        img = pyautogui.screenshot(region=(int(x) - 45, int(y) - 7, 90, 14)).convert("RGB")
        b = img.tobytes()                                      # นับเฉพาะจุดดำ/เทาเข้ม (ตัวอักษร) ไม่นับกรอบสีแดงของช่อง
        n = 0
        for i in range(0, len(b), 3):                          # ตัวอักษรดำ/น้ำเงิน นับ  กรอบแดง ไม่นับ
            r, g, bl = b[i], b[i + 1], b[i + 2]
            if 0.299 * r + 0.587 * g + 0.114 * bl < 110 and not (r > g + 30 and r > bl + 30):
                n += 1
        return n >= 25

    typed_shot = {}

    def owner_box():
        """ช่อง 'ข้อมูลของ .. (HN:xxxxx)' ด้านบน (เห็นทุกแท็บ) คำนวณจากตำแหน่งแท็บข้อมูลบุคคล"""
        try:
            x, y = at("tab_person")
            return pyautogui.screenshot(region=(int(x) + 240, int(y) - 54, 240, 18)).convert("L")
        except Exception:
            return None

    def box_changed(a_img, b_img):
        return sum(1 for p, q in zip(a_img.tobytes(), b_img.tobytes()) if abs(p - q) > 60) >= 15

    def shot(k):
        x, y = at(k)
        return pyautogui.screenshot(region=(int(x) - 45, int(y) - 7, 90, 14)).convert("L")

    def differs(a_img, b_img, need=40):
        return sum(1 for p, q in zip(a_img.tobytes(), b_img.tobytes()) if abs(p - q) > 60) >= need

    def warns():
        """ปิดคำเตือนเรื่องเลขบัตรที่เด้งระหว่างกรอก: Invalid CID (OK) / Confirm (Party) Mod11 (CANCEL) คืน True ถ้าปิดไป"""
        hit = False
        for _ in range(4):
            w = pygetwindow.getActiveWindow(); ttl = (w.title or "") if w else ""
            if ttl.startswith("Invalid CID") or ttl.startswith("!This is THAI"):   # แค่แจ้งเตือนเลขบัตร มีปุ่ม OK ปุ่มเดียว
                pyautogui.press("enter"); time.sleep(0.8); hit = True; continue
            if party():
                hit = True; continue
            break
        return hit

    def party():
        """ปิดหน้าต่างเตือน 'Confirm (Party!!!)....Duplicate CID' (เลขบัตรไม่ผ่าน Mod11) ด้วย CANCEL [เข้าใจแล้ว]"""
        w = pygetwindow.getActiveWindow()
        if w and (w.title or "").startswith("Confirm (Party"):
            # ต้องกด CANCEL [เข้าใจแล้ว] (ปุ่มขวา) เท่านั้น: OK/Esc = อ่านทวน -> กลับไปช่องเลขบัตร แล้วเตือนซ้ำไม่จบ
            pyautogui.click(int(w.left + w.width * 0.66), int(w.top + w.height * 0.89)); time.sleep(0.8)
            return True
        return False

    def typ(k, text, thai=False, check=True):
        base = shot(k) if check else None                      # ภาพช่องก่อนพิมพ์ (ช่องเลขบัตรมีขีด _-____- อยู่แล้ว นับจุดดำอย่างเดียวพลาดได้)
        for attempt in range(3):
            warns()
            pyautogui.click(*at(k)); time.sleep(0.3)
            warns()
            pyautogui.click(*at(k)); time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a"); time.sleep(0.1)
            if thai:
                pyperclip.copy(text); time.sleep(0.2)
                pyautogui.hotkey("ctrl", "v"); time.sleep(0.3)
            else:
                pyautogui.write(text, interval=0.03)
            time.sleep(0.2)
            pyautogui.press("tab"); time.sleep(0.4)
            warns()
            if k == "idcard" and not text.isdigit():           # เลข G: JHCIS เตือน Mod11 ช้า ๆ หลังออกจากช่อง -> รอแล้วกด CANCEL
                for _ in range(6):
                    if warns():
                        break
                    time.sleep(0.4)
            if not check or has_text(k) or differs(base, shot(k), 40 if len(text.strip()) >= 2 else 6):
                if check:
                    typed_shot[k] = shot(k)
                return True
        return False

    def num(v):
        try:
            return int(str(v).strip())
        except (TypeError, ValueError):
            return None

    def check(t):
        sel = "p.pid,p.fname,p.lname,p.birth,p.typelive" + "".join(f",p.{v} AS {k}" for k, v in C.items() if v) + \
              ",s.villcode,s.schoolno,s.classeducate,s.classroom,s.yeareducate"
        rs = rc.q(f"SELECT {sel} FROM person p LEFT JOIN personstudent s ON s.pid=p.pid AND s.pcucodeperson=p.pcucodeperson "
                  f"WHERE p.idcard='{t['id']}'", pw)
        if len(rs) != 1:
            return f"พบเลขบัตรนี้ {len(rs)} แถว (ต้องเป็น 1)"
        r = rs[0]
        tp, tcode, tsex = TITLE_POS[t["pre"]]
        exp = {"fname": t["fn"], "lname": t["ln"], "birth": t["bd"], "typelive": "4",
               "villcode": vill, "schoolno": sno, "classeducate": t["code"], "classroom": t["room"], "yeareducate": str(a.year)}
        cut = lambda k, x: x[:10] if k == "birth" else x.strip()   # ตัด 10 ตัวเฉพาะวันเกิด (มีเวลาต่อท้าย) ชื่อยาวต้องเทียบทั้งคำ
        bad = [f"{k}={r.get(k)} (ต้องเป็น {v})" for k, v in exp.items() if cut(k, str(r.get(k) or "")) != v]
        for k, v in (("nation", NATION_POS[t["nat"]][1]), ("race", NATION_POS[t["race"]][1]),
                     ("religion", RELIGION_POS[t["rel"]][1] if t["rel"] else None), ("prename", tcode), ("sex", tsex)):
            if C[k] and v is not None and num(r.get(k)) != num(v):
                bad.append(f"{k}={r.get(k)} (ต้องเป็น {v})")
        return "; ".join(bad)

    todo = [t for t in rows if not t["skip"]]
    if a.sex:
        want = 2 if a.sex.upper().startswith("F") else 1
        todo = [t for t in todo if TITLE_POS[t["pre"]][2] == want]
        print(f"เลือกเฉพาะเพศ{'หญิง' if want == 2 else 'ชาย'}: {len(todo)} คน")
    for t in todo:
        t["code"] = cls_code.get(rc.norm_class(t["cls"]))
        t["room"] = t["room"] or str(a.room)
    print(f"โรงเรียน: {school} | ในแผ่นงาน {len(rows)} คน | ทำได้ {len(todo)} คน | ข้ามให้คนทำ {len(rows) - len(todo)} คน")
    print("โหมด: " + ("บันทึกจริง" if a.save else "ซ้อม ไม่บันทึก"))
    input("เลือกหมู่บ้าน 00-หมู่บ้านนอกเขต และบ้านเลขที่ 0 ไว้แล้ว / แป้นพิมพ์ ENG แล้วกด Enter (มีเวลา 3 วินาที)")
    time.sleep(3)
    log, done = [(t, "ข้าม-" + t["skip"]) for t in rows if t["skip"]], 0
    retry, reuse, fixes = [], {}, []
    cur, skipped, deleted = {}, set(), []                      # cur = HN ที่สคริปต์สร้างแล้วยังไม่บันทึก                                      # reuse[id] = HN ว่างที่สคริปต์สร้างเองแล้ว JHCIS ล้างฟอร์ม -> กรอกซ้ำใน HN เดิม

    def todo_iter():
        for x in todo:
            yield x
            while retry:
                yield retry.pop()

    def delete_pending():
        """ลบ HN ว่างที่สคริปต์เพิ่งสร้าง (ยังไม่บันทึก) ด้วยปุ่ม 'ลบ' ของ JHCIS  คืน None = สำเร็จ, ข้อความ = ทำไม่ได้"""
        pid = cur["pid"]
        r = rc.q(f"SELECT pid,fname,idcard,typelive FROM person WHERE pid={int(pid)}", pw)
        if not r:
            return None                                        # ไม่มีในฐานแล้ว
        r = r[0]
        idc = (r["idcard"] or "").strip()
        if str(r["typelive"]) != "4" or (idc and idc != cur["t"]["id"]):   # ว่าง หรือเป็นเด็กคนนี้ที่สคริปต์เพิ่งพิมพ์ (JHCIS บันทึกข้อมูลบุคคลให้เองตอนเปลี่ยนแท็บ)
            return f"HN {pid} เป็นคนอื่น (เลขบัตร {idc or '-'}) หรือไม่ใช่คนนอกเขต ไม่ลบ"
        if rc.q(f"SELECT pid FROM personstudent WHERE pid={int(pid)} AND classroom IS NOT NULL AND TRIM(classroom)<>''", pw):
            return f"HN {pid} มีสถานะนักเรียนบันทึกแล้ว ไม่ลบ"
        if maxpid() != int(pid):
            return f"มีคนใหม่ถูกสร้างหลัง HN {pid} (อาจมีคนอื่นใช้ JHCIS อยู่) ไม่ลบ"
        if "del" not in P:
            return "ยังไม่ได้จดตำแหน่งปุ่ม 'ลบ' (รัน calibrate-del)"
        bad = clear_popups()
        if bad:
            return f"มีหน้าต่าง '{bad}' ค้างอยู่"
        pyautogui.click(*at("tab_person")); time.sleep(0.8)
        if cur.get("shot") is None or box_changed(cur["shot"], owner_box()):
            return f"คนบนจอไม่ใช่ HN {pid} ที่สร้างไว้ ไม่ลบ"
        n0 = int(rc.q("SELECT COUNT(*) AS n FROM person", pw)[0]["n"])
        pyautogui.click(*at("del")); time.sleep(1.2)
        for _ in range(5):                                     # หน้าต่างยืนยันการลบ
            w = pygetwindow.getActiveWindow(); ttl = (w.title or "") if w else ""
            if ttl.startswith(WIN_KEY):
                break
            if ttl in ("Message", "Select an Option", "Confirm", "Input") or ttl.startswith(("Confirm", "ยืนยัน", "คำเตือน", "ลบ")):
                print(f"    (หน้าต่าง '{ttl}' -> OK)"); pyautogui.press("enter"); time.sleep(1.2); continue
            return f"หลังกดลบ มีหน้าต่างที่ไม่รู้จัก '{ttl}' (ยังไม่ได้กด) ส่งรูปให้ดู"
        time.sleep(a.wait)
        gone = not rc.q(f"SELECT pid FROM person WHERE pid={int(pid)}", pw)
        n1 = int(rc.q("SELECT COUNT(*) AS n FROM person", pw)[0]["n"])
        if not gone and n1 < n0:
            return f"!!! อันตราย: HN {pid} ยังอยู่ แต่จำนวนคนลดลง = ลบคนอื่นไป ตรวจทันที (มีไฟล์สำรอง)"
        if not gone:
            return f"กดลบแล้ว HN {pid} ยังอยู่ในฐาน"
        return None

    SKIPF = os.path.join(HERE, "jhcis_newperson_skip.txt")   # เด็กที่เคยมีปัญหา -> ข้ามทุกรอบ ให้คนทำมือ
    try:
        old_skip = {ln.split("\t")[0].strip() for ln in open(SKIPF, encoding="utf-8") if ln.strip()}
    except OSError:
        old_skip = set()

    def mark_skip(t, why):
        with open(SKIPF, "a", encoding="utf-8") as f:
            f.write(f"{t['id']}\t{t['name']}\t{t['cls']}\t{why}\n")

    while True:
        for t in todo_iter():
            if t["id"] in skipped:
                continue
            if t["id"] in old_skip:
                log.append((t, "ข้าม-เคยมีปัญหา (อยู่ใน jhcis_newperson_skip.txt) ทำมือ")); continue
            if a.limit and done >= a.limit and t["id"] not in reuse:
                break
            if t["code"] is None:
                log.append((t, "ข้าม-ไม่รู้รหัสชั้น")); continue
            if rc.q(f"SELECT pid FROM person WHERE idcard='{t['id']}'", pw):
                log.append((t, "มีในฐานข้อมูลแล้ว-ข้าม")); continue
            # ส่งชื่อไทยเป็นเลขฐาน 16 (ถ้าส่งเป็นตัวอักษรไทยตรง ๆ Windows แปลงรหัสเพี้ยน หาไม่เจอ -> สร้างคนซ้ำ)
            hx = lambda v: "_utf8 0x" + v.strip().encode("utf-8").hex()
            nolast = t["ln"].strip() in ("", "-")                  # สกุล "-": ชื่อซ้ำกันเยอะ ใช้ ชื่อ+วันเกิด แทน
            same = rc.q(f"SELECT pid,idcard,birth FROM person WHERE TRIM(fname)={hx(t['fn'])} AND "
                        + (f"DATE(birth)='{t['bd']}'" if nolast else f"TRIM(lname)={hx(t['ln'])}"), pw)
            if same:                                               # ชื่อ+สกุลซ้ำ = อาจเป็นคนเดิมที่เลขบัตร/วันเกิดไม่ตรง ห้ามสร้างซ้ำ
                log.append((t, ("ข้าม-มีชื่อ+วันเกิดนี้ใน JHCIS แล้ว: " if nolast else "ข้าม-มีชื่อ+สกุลนี้ใน JHCIS แล้ว: ") + ", ".join(f"HN {x['pid']}" for x in same))); continue
            if t["id"] not in reuse:
                done += 1
            import pygetwindow

            def conflict():
                """ถ้ามีหน้าต่าง 'Data Conflict' (เพศขัดกับคำนำหน้า) กด OK แล้วคืน True"""
                w = pygetwindow.getActiveWindow()
                if w and (w.title or "").startswith("Data Conflict"):
                    pyautogui.press("enter"); time.sleep(0.8)
                    return True
                return False

            def clear_popups(max_n=4):
                """ปิดหน้าต่างของ JHCIS ที่รู้จัก (Message = แจ้งเตือนมีปุ่ม OK, Confirm insert new person = ยืนยันเพิ่มคน)
                คืนชื่อหน้าต่างที่ไม่รู้จัก ถ้ามี"""
                w = pygetwindow.getActiveWindow()
                ttl = (w.title or "") if w else ""
                if not (ttl == "Message" or ttl.startswith(("Confirm insert new person", "Invalid CID", "Confirm (Party", "!This is THAI")) or ttl.startswith(WIN_KEY)):
                    mains = [x for x in pygetwindow.getAllWindows() if (x.title or "").startswith(WIN_KEY)]
                    if not mains:
                        return "ไม่พบหน้าต่าง JHCIS"
                    try:
                        mains[0].activate()                         # ดึง JHCIS ขึ้นหน้า (ถ้ามีข้อความค้าง จะขึ้นมาด้วย)
                    except Exception:
                        pass
                    time.sleep(0.8)
                for _ in range(max_n):
                    w = pygetwindow.getActiveWindow()
                    ttl = (w.title or "") if w else ""
                    if ttl == "Message" or ttl.startswith("Confirm insert new person"):
                        pyautogui.press("enter"); time.sleep(0.9)
                        continue
                    if ttl.startswith("Invalid CID") or ttl.startswith("!This is THAI"):   # แจ้งเตือนเลขบัตร -> OK
                        pyautogui.press("enter"); time.sleep(0.9)
                        continue
                    if ttl.startswith("Confirm (Party"):          # เตือนเลข G ไม่ผ่าน Mod11 -> CANCEL [เข้าใจแล้ว]
                        party(); continue
                    return None if (not ttl or ttl.startswith(WIN_KEY)) else ttl
                return "หน้าต่างแจ้งเตือนขึ้นซ้ำหลายครั้ง"

            bad = clear_popups()                                   # ข้อความค้างจากรอบก่อน
            if bad:
                print(f"\n*** หยุด: มีหน้าต่าง '{bad}' เปิดค้างอยู่"); log.append((t, f"หยุด-มีหน้าต่าง {bad}")); break
            maxpid = lambda: int((rc.q("SELECT MAX(pid) AS m FROM person", pw) or [{"m": 0}])[0]["m"] or 0)
            if t["id"] in reuse:                                   # กรอกซ้ำใน HN ว่างเดิม: ต้องเป็น HN ล่าสุด และยังว่างจริงในฐานข้อมูล
                rp = reuse[t["id"]]
                ok_r = maxpid() == rp and rc.q(f"SELECT pid FROM person WHERE pid={rp} AND (fname IS NULL OR TRIM(fname)='')", pw)
                if not ok_r:
                    print(f"\n*** หยุด: จะกรอกซ้ำใน HN {rp} แต่ไม่ใช่ HN ล่าสุดหรือไม่ว่างแล้ว ตรวจหน้าจอ JHCIS")
                    log.append((t, f"หยุด-กรอกซ้ำ HN {rp} ไม่ได้")); break
                print(f"  ↻ กรอกซ้ำ {t['name']} ใน HN {rp} (JHCIS ล้างฟอร์มรอบที่แล้ว)")
            else:
                mpick("vill", a.vill_pos); time.sleep(a.wait)          # 00-หมู่บ้านนอกเขต หมู่ที่ 0
                mpick("house", a.house_pos); time.sleep(a.wait)        # บ้านเลขที่ 0
                before_new = maxpid()                                  # ใช้ฐานข้อมูลดูว่ากด 'เพิ่มคนใหม่' แล้วมีคนใหม่เกิดจริง (แม่นกว่าดูภาพ)
                pyautogui.click(*at("new")); time.sleep(a.wait)
                bad = clear_popups()                                   # เตือนยังไม่ระบุชื่อ / ยืนยันเพิ่มคน ขึ้นลำดับไหนก็ได้
                if not bad:                                            # บางครั้งหน้าต่างยืนยันขึ้นช้า แล้วล้างฟอร์มทีหลัง -> รอแล้วเก็บอีกรอบ
                    time.sleep(0.5); bad = clear_popups()
                if bad:
                    print(f"\n*** หยุด: หลังกด 'เพิ่มคนใหม่' มีหน้าต่าง '{bad}'  ตรวจหน้าจอ JHCIS")
                    log.append((t, f"หยุด-มีหน้าต่าง {bad} หลังเพิ่มคนใหม่")); break
                if True:                                               # JHCIS บางครั้งสร้างคนใหม่ช้า -> รอดูสูงสุด ~8 วินาที
                    for _ in range(16):
                        if maxpid() > before_new:
                            break
                        time.sleep(0.5); bad = clear_popups()
                        if bad:
                            break
                if bad:
                    print(f"\n*** หยุด: หลังกด 'เพิ่มคนใหม่' มีหน้าต่าง '{bad}'  ตรวจหน้าจอ JHCIS")
                    log.append((t, f"หยุด-มีหน้าต่าง {bad} หลังเพิ่มคนใหม่")); break
                time.sleep(1.0); bad = bad or clear_popups()           # ให้หน้าจอขึ้นคนใหม่ให้เสร็จก่อนเริ่มกรอก
                if maxpid() <= before_new:
                    print("\n*** หยุด: กด 'เพิ่มคนใหม่' แล้วไม่มี HN ใหม่ในฐานข้อมูล = มีคนค้าง (ยังไม่บันทึก) อยู่บนจอ"
                          "\n    ห้ามบันทึกคนนั้น ให้กด 'ลบ' หรือกรอกให้ครบแล้วบันทึก จากนั้นรันใหม่")
                    log.append((t, "หยุด-เพิ่มคนใหม่ไม่ได้ (มีคนค้าง)")); break
            want = reuse.get(t["id"]) or maxpid()
            # ด่านสำคัญ: หน้าจอต้องเป็นคนใหม่ที่ "ว่าง" จริง (ชื่อ/วันเกิดว่าง) กันกรอกทับคนเดิมที่เปิดค้างบนจอ (เช่น HN 17168)
            pyautogui.click(*at("tab_person")); time.sleep(0.6); warns()
            def inside_text(k):                                # ดูเฉพาะ "ในช่อง" (ขวาของจุดที่จด) ไม่ให้ป้ายชื่อช่องทางซ้ายมาปน
                x, y = at(k)
                im = pyautogui.screenshot(region=(int(x) + 6, int(y) - 7, 80, 14)).convert("RGB").tobytes()
                return sum(1 for i in range(0, len(im), 3)
                           if 0.299 * im[i] + 0.587 * im[i + 1] + 0.114 * im[i + 2] < 110
                           and not (im[i] > im[i + 1] + 30 and im[i] > im[i + 2] + 30)) >= 25
            filled = [n for k, n in (("fname", "ชื่อ"), ("lname", "สกุล"), ("birth", "วันเกิด")) if inside_text(k)]
            if filled:
                for k in ("fname", "lname", "birth"):           # เก็บภาพช่องไว้ให้ตรวจว่าเตือนจริงหรือเตือนผิด
                    x, y = at(k)
                    pyautogui.screenshot(os.path.join(HERE, f"debug_{k}.png"), region=(int(x) + 6, int(y) - 7, 80, 14))
                print(f"\n*** หยุด: กด 'เพิ่มคนใหม่' แล้ว แต่บนจอช่อง {', '.join(filled)} มีข้อมูลอยู่ = ไม่ใช่คนใหม่ (อาจเป็นคนเดิมที่เปิดค้าง)"
                      f"\n    ไม่กรอกอะไร ห้ามบันทึก  ให้ลบ HN {want} ที่ว่างอยู่ด้วยมือ (ดู HN สีเขียวให้เป็น {want}) แล้วรันใหม่")
                log.append((t, f"หยุด-หน้าจอไม่ใช่คนใหม่ ({', '.join(filled)} มีข้อมูล)")); break
            cur.clear(); cur.update(pid=want, t=t, shot=owner_box())
            pyautogui.click(*at("tab_person")); time.sleep(0.6)
            # คำนำหน้าก่อน (เลือกครั้งเดียว เพราะรายการยาว ถ้ามีค่าแล้วรายการจะเปิดตรงค่านั้น ตำแหน่งจะเพี้ยน)
            # JHCIS เด้ง Data Conflict เมื่อเพศกับคำนำหน้าไม่ตรง (บางทีเด้งตอนย้ายไปช่องอื่น) -> กด OK แล้วเลือกเพศใหม่
            gpos = 1 + TITLE_POS[t["pre"]][2]                      # รายการเพศ: ว่าง, 1 ชาย, 2 หญิง
            tpos = TITLE_POS[t["pre"]][0]
            # ตั้งเพศเป็น "ว่าง" ก่อนเสมอ (เพศว่างไม่ขัดกับคำนำหน้าใด) แล้วค่อยเลือกคำนำหน้า -> เพศ
            mpick("gender", 1, home=True); time.sleep(0.6); conflict()
            mpick("title", tpos, home=True); time.sleep(0.6)
            if conflict():                                         # คำนำหน้าใหม่ขัดกับเพศที่ติดมาจากคนก่อน (JHCIS คืนค่าเดิม)
                # เปลี่ยนเพศตรง ๆ ก็ขัดกับคำนำหน้าเดิมอีก (วนไม่จบ) -> ตั้งเพศเป็น "ว่าง" ก่อน แล้วเลือกคำนำหน้าใหม่
                mpick("gender", 1, home=True); time.sleep(0.6); conflict()
                mpick("title", tpos, home=True); time.sleep(0.6)
                if conflict():
                    print(f"\n*** หยุด: เลือกคำนำหน้าไม่ได้ (Data Conflict) HN นี้ถูกสร้างแล้ว ให้แก้ด้วยมือ หรือกด 'ลบ'")
                    log.append((t, "หยุด-เพศขัดกับคำนำหน้า")); break
            okg = False
            for _ in range(3):
                mpick("gender", gpos); time.sleep(0.6); c1 = conflict()
                pyautogui.click(*at("fname")); time.sleep(0.6); c2 = conflict()
                if not (c1 or c2):
                    okg = True; break
            if not okg:
                print(f"\n*** หยุด: เพศกับคำนำหน้าขัดกัน (Data Conflict) HN นี้ถูกสร้างแล้ว ให้แก้ด้วยมือ หรือกด 'ลบ'")
                log.append((t, "หยุด-เพศขัดกับคำนำหน้า")); break
            time.sleep(0.3)                                        # รอรายการปิดก่อนพิมพ์
            empty, warn = [], False
            for k, v, th, n in (("fname", t["fn"], True, "ชื่อ"), ("lname", t["ln"], True, "สกุล"),
                                ("idcard", t["id"], False, "เลขบัตร"), ("birth", be_date(t["bd"]), False, "วันเกิด")):
                if not typ(k, v, thai=th, check=len(v.strip()) >= 2):   # "-" ตัวเล็กเกิน ภาพจับไม่ได้ -> ไปตรวจในฐานหลังบันทึกแทน
                    empty.append(n)
                w = pygetwindow.getActiveWindow()
                if party():                                       # เตือนเลข G ไม่ผ่าน Mod11 -> CANCEL [เข้าใจแล้ว]
                    w = pygetwindow.getActiveWindow()
                if w and (w.title or "") == "Message":            # หลังพิมพ์ชื่อ/สกุล JHCIS เตือน "ชื่อ-สกุลซ้ำ"
                    pyautogui.press("esc"); time.sleep(0.8)       # Esc ไม่ใช่ Enter: Enter อาจพาไปเปิดคนเดิม (เคส HN 17168)
                    if k in ("fname", "lname"):                    # ชื่อซ้ำ = หยุดคนนี้ (โหมด auto จะลบ HN ว่างแล้วข้ามให้คนทำ)
                        warn = True; break
                if cur.get("shot") is not None and box_changed(cur["shot"], owner_box()):
                    print(f"\n*** หยุด: หลังกรอกช่อง {n} คนบนจอเปลี่ยนไปจาก HN {cur['pid']} (JHCIS เปิดคนอื่นขึ้นมา) ไม่กรอกต่อ ห้ามบันทึก")
                    log.append((t, "หยุด-คนบนจอเปลี่ยนระหว่างกรอก")); cur.clear(); warn = None; break
            if warn is None:
                break
            if warn:
                print(f"\n*** หยุด: JHCIS เตือนหลังพิมพ์ชื่อ-สกุล (น่าจะเป็น 'ชื่อ-สกุลซ้ำ') HN นี้ถูกสร้างแล้ว"
                      "\n    ตรวจใน JHCIS ว่ามีคนนี้อยู่แล้วไหม ถ้ามี ให้กด 'ลบ' HN ที่เพิ่งสร้าง")
                log.append((t, "หยุด-JHCIS เตือนชื่อ-สกุลซ้ำ")); break
            if empty:
                print(f"\n*** หยุด: กรอกช่อง {', '.join(empty)} ไม่ติด (ลอง 3 ครั้ง)  HN นี้ถูกสร้างแล้ว ให้กรอกต่อด้วยมือ หรือกด 'ลบ'")
                log.append((t, "หยุด-กรอกไม่ติด: " + ", ".join(empty))); break
            pick("nation", NATION_POS[t["nat"]][0])
            pick("race", NATION_POS[t["race"]][0])
            if t["rel"]:
                pick("religion", RELIGION_POS[t["rel"]][0])
            time.sleep(0.5)
            lost = [n for k, n in (("fname", "ชื่อ"), ("idcard", "เลขบัตร"), ("birth", "วันเกิด"))
                    if not has_text(k) and (k not in typed_shot or differs(typed_shot[k], shot(k)))]   # หายจริง = ภาพเปลี่ยนไปจากตอนพิมพ์เสร็จ
            if lost:                                               # เช่น JHCIS ล้างฟอร์มเองทีหลัง (ยืนยันเพิ่มคนขึ้นช้า)
                print(f"\n*** หยุด: ก่อนไปแท็บนักเรียน ช่อง {', '.join(lost)} ว่าง (เคยกรอกแล้ว) HN นี้ถูกสร้างแล้ว ให้ตรวจหน้าจอ")
                log.append((t, "หยุด-ช่องว่างก่อนแท็บนักเรียน: " + ", ".join(lost))); break
            pyautogui.click(*at("tab_student")); time.sleep(0.8)
            import pygetwindow
            for _ in range(3):                                 # เตือนเลข G (Mod11) ตอนเปลี่ยนแท็บ -> CANCEL แล้วเปิดแท็บนักเรียนอีกครั้ง
                if not warns():
                    break
                pyautogui.click(*at("tab_student")); time.sleep(0.8)
            w = pygetwindow.getActiveWindow()
            if w and not (w.title or "").startswith(WIN_KEY):
                if w.title == "Message" and t["id"] not in reuse:  # "ยังไม่ระบุชื่อ" = JHCIS ล้างฟอร์มหลังกรอก -> กรอกซ้ำ 1 รอบใน HN เดิม
                    pyautogui.press("enter"); time.sleep(1.0); clear_popups()
                    reuse[t["id"]] = maxpid(); retry.append(t)
                    log.append((t, f"กรอกซ้ำ-JHCIS ล้างฟอร์ม HN {reuse[t['id']]}")); continue
                print(f"\n*** หยุด: หลังเปิดแท็บนักเรียน มีหน้าต่าง '{w.title}' เด้งขึ้นมา (เช่น ยังไม่ปรากฏข้อมูลประชากร)")
                print(f"    ({t['name']}) HN ที่เพิ่งสร้างยังว่าง ให้กด OK แล้วกด 'ลบ' HN นั้น จากนั้นรันใหม่")
                log.append((t, f"หยุด-มีหน้าต่าง {w.title}")); break
            # ใส่ห้องกับปีก่อน แล้วค่อยเลือกโรงเรียน/ชั้น: ถ้าเลือกชั้นตอนห้องยังว่าง JHCIS เด้ง
            # "โปรดระบุโรงเรียน ชั้น และห้องเรียนให้ครบถ้วน" มาขัดตอนพิมพ์ห้อง
            def setf(k, v):                                        # ช่องห้อง/ปี: ลบค่าเดิมทีละตัว (Ctrl+A ใช้ไม่ได้) แล้วพิมพ์ใหม่
                pyautogui.click(*at(k)); time.sleep(0.3)
                pyautogui.press("end"); pyautogui.press("backspace", presses=8, interval=0.01)
                pyautogui.write(v, interval=0.03); pyautogui.press("tab"); time.sleep(0.3)
            setf("room", t["room"]); setf("year", str(a.year))
            if clear_popups():
                print("\n*** หยุด: หลังใส่ห้อง/ปี มีหน้าต่างแปลก ๆ ค้างอยู่ ตรวจหน้าจอ JHCIS")
                log.append((t, "หยุด-หน้าต่างหลังใส่ห้อง/ปี")); break
            mpick("school", a.school_pos)
            time.sleep(0.4)
            mpick("cls", a.class1_pos + int(t["code"]) - 1, home=True)
            time.sleep(0.8)
            clear_popups()                                         # เช่น "โปรดระบุโรงเรียน ชั้น และห้อง..." -> OK
            # เปลี่ยนชั้นแล้ว JHCIS อาจล้างห้อง/ปี -> ใส่ซ้ำทุกครั้ง (ลบค่าเดิมก่อน จึงไม่ต่อท้ายเป็น 11 / 25692569)
            setf("room", t["room"]); setf("year", str(a.year))
            time.sleep(0.4); clear_popups()
            w = pygetwindow.getActiveWindow()
            if (w and not (w.title or "").startswith(WIN_KEY)) or not has_text("year"):
                print(f"\n*** หยุด: {t['name']} หลังใส่ห้อง/ปี มีหน้าต่าง '{w.title if w else ''}' หรือช่องปีว่าง ไม่บันทึก"
                      f"\n    HN นี้ถูกสร้างแล้ว: ใส่ห้อง {t['room']} ปี {a.year} แล้วกด SAVE + แผ่นดิสก์เอง หรือกด 'ลบ'")
                log.append((t, "หยุด-ห้อง/ปีไม่ติด")); break
            okb = "บันทึก" if a.save else "ไม่บันทึก ไปคนถัดไป (ซ้อม)"
            # ตรวจด้วยฐานข้อมูลว่าคนที่เพิ่งสร้างอยู่บ้านนอกเขต (JHCIS สร้างคนทันทีที่กด เพิ่มคนใหม่)
            ans = okb if (a.auto and a.save) else pyautogui.confirm(
                text=f"{t['name']}\nเลข {t['id']}  เกิด {be_date(t['bd'])}\n"
                     f"สัญชาติ {NATION_POS[t['nat']][1]} เชื้อชาติ {NATION_POS[t['race']][1]} ศาสนา {t['rel'] or '(ไม่แก้)'}\n"
                     f"{school} / {t['cls']} / ห้อง {t['room']} / ปี {a.year}\n\nดูทั้ง 2 แท็บใน JHCIS ว่าถูกต้อง",
                title="เพิ่มคนใหม่ JHCIS", buttons=[okb, "ข้ามคนนี้", "หยุด"])
            if ans in (None, "หยุด"):
                log.append((t, "หยุดโดยผู้ใช้-คนนี้ถูกสร้างค้างไว้ใน JHCIS ต้องแก้ด้วยมือ")); break
            if not (a.save and ans == okb):
                log.append((t, f"ไม่บันทึก ({ans})-คนนี้ถูกสร้างค้างไว้ใน JHCIS")); break
            time.sleep(0.4)
            if cur.get("shot") is not None and box_changed(cur["shot"], owner_box()):
                print(f"\n*** หยุด: ก่อนบันทึก คนบนจอไม่ใช่ HN {cur['pid']} ที่สร้างไว้ ไม่บันทึก")
                log.append((t, "หยุด-คนบนจอเปลี่ยนก่อนบันทึก")); cur.clear(); break
            pyautogui.click(*at("save")); time.sleep(1.0)
            pyautogui.click(*at("disk")); time.sleep(a.wait + 1)
            cur.clear()                                            # บันทึกแล้ว ไม่ใช่คนว่างค้าง
            bad = check(t)
            if bad:                                                # บันทึกไปแล้ว (คนนี้ไม่ใช่คนว่างค้าง) -> จดไว้แก้มือ แล้วทำคนต่อไป
                log.append((t, "ต้องแก้มือ-ตรวจฐานข้อมูลไม่ผ่าน: " + bad))
                fixes.append((t, bad))
                print(f"  ✗ {t['name']} ({t['id']}) บันทึกแล้วแต่ไม่ตรง: {bad}  -> จดไว้แก้มือ")
                if len(fixes) >= 3:                                # ผิดซ้ำหลายคน = น่าจะผิดเป็นระบบ หยุดก่อน
                    print("\n*** หยุด: ตรวจฐานข้อมูลไม่ผ่าน 3 คนแล้ว น่าจะมีอะไรผิดเป็นระบบ ตรวจหน้าจอ JHCIS ก่อนรันต่อ")
                    break
                continue
            log.append((t, "เพิ่มแล้ว-ตรวจฐานข้อมูลผ่าน")); print(f"  ✓ {t['name']}")
        if cur and a.save and not a.auto:                     # โหมดถามทีละคน: จดไว้ให้รอบหน้าข้ามคนนี้ (ต้องลบ HN ค้างเอง)
            mark_skip(cur["t"], f"หยุด HN {cur['pid']} ค้าง")
            print(f"    (จดชื่อ {cur['t']['name']} ไว้แล้ว รอบหน้าจะข้ามคนนี้ ให้ทำมือ)")
        if not cur or not a.save or not a.auto:              # ลบเองเฉพาะโหมด --auto
            break
        ft = cur["t"]
        why = delete_pending()
        if why:
            print(f"\n*** หยุด: ลบ HN ค้างเองไม่ได้: {why}\n    ให้ลบ HN {cur['pid']} ด้วยมือ แล้วรันใหม่")
            break
        print(f"  🗑 ลบ HN {cur['pid']} ที่ค้าง ({ft['name']}) แล้ว -> ข้ามคนนี้ไปทำคนต่อไป")
        mark_skip(ft, "ลบ HN ค้างแล้ว")
        deleted.append(ft); skipped.add(ft["id"]); cur.clear()
        log.append((ft, "ข้าม-ลบ HN ค้างแล้ว ต้องเพิ่มด้วยมือหรือรันใหม่"))
        if len(deleted) >= 5:
            print("\n*** หยุด: ต้องลบ HN ค้างครบ 5 ครั้งแล้ว น่าจะมีอะไรผิดเป็นระบบ ตรวจหน้าจอ JHCIS")
            break
    if deleted:
        print("\nข้ามไป (ลบ HN ค้างแล้ว) รันคำสั่งเดิมซ้ำทีหลังได้ หรือเพิ่มด้วยมือ:")
        for t in deleted:
            print(f"  {t['id']}  {t['name']}  -> {t['cls']} ห้อง {t['room']}")
    if fixes:
        print("\nต้องแก้มือใน JHCIS (ค้นเลขบัตร แล้วแก้ช่องที่บอก แล้วบันทึก):")
        for t, b in fixes:
            print(f"  {t['id']}  {t['name']}  -> {t['cls']} ห้อง {t['room']} ปี {a.year} | {b}")
    with open(os.path.join(HERE, "jhcis_newperson_log.txt"), "a", encoding="utf-8") as f:
        for t, s in log:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{srv}\t{school}\t{t['id']}\t{t['name']}\t{t['cls']}\t{s}\n")
    from collections import Counter
    print("\nสรุป:", dict(Counter(s.split(":")[0].split(" (")[0] if not s.startswith("ข้าม-") else "ข้าม(ให้คนทำ)" for _, s in log)))
    print("รายละเอียดใน jhcis_newperson_log.txt")


def cli():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("calibrate")
    sub.add_parser("calibrate-lists")
    sub.add_parser("calibrate-cls")
    sub.add_parser("calibrate-vill")
    sub.add_parser("calibrate-del")
    sub.add_parser("calibrate-hn")
    lo = sub.add_parser("leftovers")
    lo.add_argument("--user", default="root")
    r = sub.add_parser("run")
    r.add_argument("xlsx")
    r.add_argument("--school-pos", type=int, required=True)
    r.add_argument("--class1-pos", type=int, default=3, help="ลำดับของ 'อนุบาล1' ในรายการชั้นเรียน (ว่าง, ก่อนอนุบาล, อนุบาล1 = 3)")
    r.add_argument("--vill-pos", type=int, default=2, help="ลำดับของ '00-หมู่บ้านนอกเขต' ในรายการหมู่บ้าน")
    r.add_argument("--house-pos", type=int, default=2, help="ลำดับของบ้านเลขที่ '0' ในรายการบ้านเลขที่")
    r.add_argument("--room", type=int, default=1)
    r.add_argument("--year", type=int, default=None, help="ปีการศึกษา พ.ศ. (ไม่ใส่ = คำนวณจากวันที่: เปิดเทอม พ.ค.)")
    r.add_argument("--limit", type=int, default=0)
    r.add_argument("--sex", default="", help="ทำเฉพาะเพศ: M = ชาย, F = หญิง (ใช้ตอนซ้อม)")
    r.add_argument("--save", action="store_true")
    r.add_argument("--auto", action="store_true")
    r.add_argument("--real", action="store_true")
    r.add_argument("--ignore-leftovers", action="store_true", help="รันได้แม้มีคนค้าง (ไม่แนะนำ)")
    r.add_argument("--user", default="root")
    r.add_argument("--delay", type=float, default=0.15)
    r.add_argument("--wait", type=float, default=1.5)
    a = ap.parse_args()
    if getattr(a, "year", 0) is None:
        import datetime as _dt
        _d = _dt.date.today()
        a.year = _d.year + 543 - (0 if _d.month >= 5 else 1)   # ปีการศึกษาไทยเริ่มเดือนพฤษภาคม
        print(f"ปีการศึกษา: {a.year} (คำนวณจากวันที่ {_d:%d/%m/%Y}  ถ้าไม่ถูกให้ใส่ --year เอง)")
    {"calibrate": calibrate, "calibrate-lists": calibrate_lists,
     "calibrate-cls": lambda: calibrate_lists(("cls_last", "cls_down")),
     "calibrate-vill": lambda: calibrate_lists(("vill", "house")),
     "calibrate-del": lambda: calibrate_lists(("del",)),
     "calibrate-hn": lambda: calibrate_lists(("hnbox",)),
     "leftovers": lambda: leftovers(a)}.get(a.cmd, lambda: run(a))()


if __name__ == "__main__":
    cli()
