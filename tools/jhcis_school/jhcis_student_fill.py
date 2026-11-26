# -*- coding: utf-8 -*-
"""
jhcis_student_fill.py  (รุ่นทดลอง v0.7.3 เลือกรายการด้วยเมาส์ ไม่กด เพิ่มคนใหม่)
ใส่/แก้ โรงเรียน-ชั้น-ห้อง-ปีการศึกษา ของนักเรียนใน JHCIS แทนคน (แท็บ "ตำแหน่งในชุมชน & นักเรียน")
อ่านรายชื่อจากไฟล์ "ผลเทียบ_....xlsx" แผ่น "2 แก้ชั้น" และ "3 เพิ่มเข้าโรงเรียน"

ต้องมีไฟล์ jhcis_roster_compare.py อยู่ในโฟลเดอร์เดียวกัน

1) จดตำแหน่ง (ครั้งเดียว):  python jhcis_student_fill.py calibrate
2) ซ้อม ไม่กดบันทึก:        python jhcis_student_fill.py run "ผลเทียบ_xxx.xlsx" --school-pos N --limit 2
3) บันทึกจริง:              เติม --save      (ทำต่อเนื่องไม่ถามทีละคน: เติม --auto)

--school-pos N   โรงเรียนนี้อยู่ลำดับที่เท่าไรในรายการ "โรงเรียน" ของ JHCIS (นับจากบน เริ่มที่ 1)
--class1-pos N   "อนุบาล1" อยู่ลำดับที่เท่าไรในรายการ "ชั้นเรียน" (ว่าง, ก่อนอนุบาล, อนุบาล1 = 3)
ใช้ระยะแถวของรายการจากไฟล์ jhcis_newperson_points.json (จดไว้แล้วตอนทำสคริปต์เพิ่มคนใหม่)
--room N         ห้อง (ปกติ 1)      --year 2569   ปีการศึกษา

ความปลอดภัย
- ก่อนทำแต่ละคน อ่านฐานข้อมูลดูก่อน ถ้าถูกอยู่แล้วจะข้าม (รันซ้ำได้)
- หลังบันทึกแต่ละคน อ่านฐานข้อมูลตรวจว่า โรงเรียน/ชั้น/ห้อง/ปี ตรงตามเป้า ถ้าไม่ตรงหยุดทันที
- ปกติทำได้เฉพาะเมื่อ JHCIS ต่อกับฐานทดสอบในเครื่อง (SERVER=localhost) ฐานจริงต้องเติม --real
- แป้นพิมพ์ต้องเป็น ENG  หยุดฉุกเฉิน: ลากเมาส์ไปมุมซ้ายบนสุดของจอ
"""
import os, re, sys, json, time, argparse, getpass
import jhcis_roster_compare as rc

HERE = rc.home()
CAL = os.path.join(HERE, "jhcis_student_points.json")
PROP = r"C:\Program Files\JHCIS\database.properties"
WIN_KEY = "JHCIS V"
POINTS = [
    ("id_field", "ช่อง 'บัตร ปชช.' ที่ใช้พิมพ์เลขบัตรค้นหา (ด้านบนซ้าย)"),
    ("first_row", "แถวแรกของผลการค้นหา (ถ้ากด Enter แล้ว JHCIS เปิดข้อมูลคนนั้นเลย ไม่มีรายการ ให้พิมพ์ s แล้ว Enter เพื่อข้าม)"),
    ("tab", "แท็บ 'ตำแหน่งในชุมชน & นักเรียน & ตำแหน่งทางศาสนา'"),
    ("school", "ช่อง 'โรงเรียน' (กลางช่อง ไม่ใช่ลูกศร)"),
    ("cls", "ช่อง 'ชั้นเรียน' (กลางช่อง ไม่ใช่ลูกศร)"),
    ("room", "ช่อง 'ห้อง'"),
    ("year", "ช่อง 'ปีการศึกษา'"),
    ("save", "ปุ่ม 'บันทึก ( SAVE )' ของสถานะนักเรียน (ปุ่มบน ไม่ใช่ของตำแหน่งทางศาสนา)"),
    ("disk", "ปุ่มรูปแผ่นดิสก์ มุมซ้ายล่าง (save ครั้งที่ 2)"),
]


def origin():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if (w.title or "").startswith(WIN_KEY)]
    return (ws[0].left, ws[0].top) if ws else None


def calibrate():
    import pyautogui
    if not origin():
        sys.exit("ไม่พบหน้าต่าง JHCIS เปิดหน้าข้อมูลบุคคลไว้ก่อน")
    print("ชี้เมาส์ (ไม่คลิก) แล้วกด Enter ที่หน้าต่างนี้  **ห้ามกดบันทึกใน JHCIS ระหว่างจดตำแหน่ง**")
    pts = {}
    for i, (k, d) in enumerate(POINTS):
        if i == 1:
            input(">>> พิมพ์เลขบัตรของเด็กที่ 'มีอยู่แล้ว' ใน JHCIS ในช่องค้นหา กด Enter ใน JHCIS แล้วกลับมากด Enter ที่นี่")
        if i == 3:
            input(">>> ให้ข้อมูลของเด็กคนนั้นเปิดอยู่ (ถ้ามีรายการให้คลิกเลือกก่อน) แล้วเปิดแท็บนักเรียนค้างไว้ แล้วกด Enter ที่นี่")
        ans = input(f"[{i+1}/{len(POINTS)}] ชี้ที่: {d}  แล้วกด Enter ").strip().lower()
        if k == "first_row" and ans == "s":
            print("      = ข้าม (ค้นแล้วเปิดข้อมูลทันที)")
            continue
        x, y = pyautogui.position()
        ox, oy = origin()
        pts[k] = [x - ox, y - oy]
        print(f"      = ({x}, {y})")
    json.dump(pts, open(CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกตำแหน่งแล้วที่", CAL)


def calibrate_hn():
    """จดตำแหน่งช่อง HN เพิ่ม (ใช้ค้นเด็กที่เลขบัตรมีตัวอักษร เช่น G...) ไม่แตะตำแหน่งอื่น"""
    import pyautogui
    if not os.path.exists(CAL):
        sys.exit("ยังไม่มีไฟล์ตำแหน่ง ให้รัน calibrate ก่อน")
    if not origin():
        sys.exit("ไม่พบหน้าต่าง JHCIS")
    pts = json.load(open(CAL, encoding="utf-8"))
    print("**ห้ามกดบันทึกใน JHCIS ระหว่างนี้**")
    input("[1/2] ชี้เมาส์ (ไม่คลิก) ที่ช่อง 'HN:' ด้านบน (ข้างช่อง ไฟล์รูปภาพ) แล้วกด Enter ที่นี่ ")
    x, y = pyautogui.position(); ox, oy = origin()
    pts["hn_field"] = [x - ox, y - oy]
    print(f"      = ({x}, {y})")
    print(">>> ลองใน JHCIS: คลิกช่อง HN พิมพ์ HN ของเด็กที่มีอยู่ (เช่น 16133) แล้วกด Enter ใน JHCIS")
    ans = input("[2/2] JHCIS ขึ้นเป็น 'รายการ' ในแท็บรายชื่อ ปชก. เหมือนตอนค้นเลขบัตร ใช่ไหม? พิมพ์ y = มีรายการ / n = เปิดข้อมูลคนนั้นเลย แล้ว Enter ").strip().lower()
    pts["hn_list"] = ans.startswith("y")
    json.dump(pts, open(CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว: ค้นด้วย HN", "แล้วดับเบิลคลิกแถวแรก" if pts["hn_list"] else "แล้วเปิดข้อมูลทันที")


def jhcis_server():
    try:
        for ln in open(PROP, encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                return ln.split("=", 1)[1].strip()
    except OSError:
        return None


def read_targets(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    school = wb["สรุป"]["B1"].value
    out = []
    for sheet, col in (("2 แก้ชั้น", "ต้องแก้เป็น"), ("3 เพิ่มเข้าโรงเรียน", "ชั้นที่ต้องใส่")):
        ws = wb[sheet]
        h = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            d = dict(zip(h, r))
            note = str(d.get("หมายเหตุ") or "")
            out.append({"id": str(d["เลขประจำตัว"]), "pid": str(d["pid"]), "cls": d.get(col) or "",
                        "name": f"{d['คำนำหน้า']}{d['ชื่อ']} {d['นามสกุล']}", "room": str(d.get("ห้อง") or ""),
                        "skip": "ตรวจด้วยคน" if "มากกว่า 1 คน" in note else ("ไม่พบชั้น" if not d.get(col) else "")})
    return school, out


def ask(msg, buttons):
    import pyautogui
    print(msg.replace("\n\n", "\n"))
    return pyautogui.confirm(text=msg, title="ใส่ชั้นเรียน JHCIS อัตโนมัติ", buttons=buttons)


def run(a):
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = a.delay
    if not os.path.exists(CAL):
        sys.exit("ยังไม่ได้จดตำแหน่ง ให้รัน calibrate ก่อน")
    P = json.load(open(CAL, encoding="utf-8"))
    NP = os.path.join(HERE, "jhcis_newperson_points.json")
    if not os.path.exists(NP):
        sys.exit("ไม่พบ jhcis_newperson_points.json (ระยะแถวของรายการ) ให้จดตำแหน่งของสคริปต์เพิ่มคนใหม่ก่อน")
    Q = json.load(open(NP, encoding="utf-8"))
    ROW = (Q["li_b"][1] - Q["li_a"][1]) / 2.0
    FIRST = Q["li_a"][1] - Q["title"][1] - ROW
    VIS = int(round((Q["cls_last"][1] - Q["cls"][1] - FIRST) / ROW)) + 1
    DOWN = (Q["cls_down"][0] - Q["cls"][0], Q["cls_down"][1] - Q["cls"][1])   # ลูกศรลงของรายการ เทียบกับช่องชั้นเรียน
    print(f"ระยะรายการ: สูงแถว {ROW:.1f} px, มองเห็น {VIS} แถว")
    srv = jhcis_server()
    if srv is None:
        sys.exit("อ่านไฟล์ตั้งค่า JHCIS ไม่ได้: " + PROP)
    local = srv.lower() in ("localhost", "127.0.0.1")
    if not local and not a.real:
        sys.exit(f"JHCIS ตอนนี้ต่อกับ {srv} (ฐานจริง) สคริปต์จะไม่ทำงาน\nสลับเป็น SERVER=localhost เพื่อทดสอบ หรือเติม --real เมื่อพร้อมใช้จริง")
    rc.HOST = "127.0.0.1" if local else srv
    rc.USER = a.user
    print("=" * 60 + f"\nJHCIS ต่อกับ: {srv}   -> " + ("ฐานทดสอบในเครื่อง" if local else "*** ฐานข้อมูลจริง ***") + "\n" + "=" * 60)
    pw = rc.ask_pw(f"รหัสผ่าน {a.user} ของ MySQL ที่ {rc.HOST}: ")

    school, T = read_targets(a.xlsx)
    sc = [s for s in rc.q("SELECT villcode,schoolno,schoolname FROM villageschool", pw) if s["schoolname"] == school]
    if len(sc) != 1:
        sys.exit(f"หาโรงเรียน '{school}' ในฐานข้อมูลไม่เจอแบบชัดเจน")
    vill, sno = sc[0]["villcode"], sc[0]["schoolno"]
    crow = rc.q("SELECT * FROM cschoolclass", pw)
    ck = list(crow[0].keys())
    codes = sorted((r[ck[0]] for r in crow), key=int)
    code_of = {r[ck[1]]: r[ck[0]] for r in crow}

    def state(pid):
        r = rc.q(f"SELECT villcode,schoolno,classeducate,classroom,yeareducate FROM personstudent WHERE pid={int(pid)}", pw)
        return r[0] if r else None

    def is_ok(st, t):
        return bool(st) and (st["villcode"], st["schoolno"], st["classeducate"], str(st["classroom"] or ""), str(st["yeareducate"] or "")) == \
            (vill, sno, t["code"], t["room"], str(a.year))

    cls_ref = {}

    def cls_shot():
        x, y = at("cls")
        return pyautogui.screenshot(region=(int(x) - 200, int(y) - 8, 180, 16)).convert("L")

    def cls_diff(a_img, b_img):
        return sum(1 for p, q in zip(a_img.tobytes(), b_img.tobytes()) if abs(p - q) > 60) >= 25

    def pick(key, pos):
        """เลือกด้วยเมาส์: เปิดรายการ เลื่อนกลับบนสุด แล้วคลิกแถวที่ pos"""
        x, y = at(key)
        pyautogui.click(x, y); time.sleep(0.5)
        pyautogui.press("home"); time.sleep(0.3)            # ถ้าช่องมีค่าอยู่แล้ว รายการจะเปิดที่ค่านั้น ให้กลับขึ้นบนสุดก่อน
        if pos > VIS:
            for _ in range(pos - VIS):
                pyautogui.click(x + DOWN[0], y + DOWN[1]); time.sleep(0.15)
            ty = y + FIRST + (VIS - 1) * ROW
        else:
            ty = y + FIRST + (pos - 1) * ROW
        pyautogui.click(x, int(round(ty))); time.sleep(0.4)

    import pygetwindow

    def clear_popups(max_n=4):
        for _ in range(max_n):
            w = pygetwindow.getActiveWindow()
            ttl = (w.title or "") if w else ""
            if ttl == "Message" or ttl.startswith("สิทธิ"):     # สิทธิขัดแย้งฯ = แค่เตือนเรื่องสิทธิ กด OK ได้
                pyautogui.press("enter"); time.sleep(0.9); continue
            return None if (not ttl or ttl.startswith(WIN_KEY) or "PowerShell" in ttl) else ttl
        return "หน้าต่างแจ้งเตือนขึ้นซ้ำหลายครั้ง"

    def at(key):
        o = origin()
        if not o:
            raise RuntimeError("ไม่พบหน้าต่าง JHCIS")
        return (o[0] + P[key][0], o[1] + P[key][1])

    todo = [t for t in T if not t["skip"]]
    for t in todo:
        t["code"] = code_of.get(t["cls"]); t["room"] = t["room"] or str(a.room)
    print(f"โรงเรียน: {school} | ต้องทำ {len(todo)} คน | ข้ามให้คนตรวจ {len(T)-len(todo)} คน | โหมด: " + ("บันทึกจริง" if a.save else "ซ้อม ไม่กดบันทึก"))
    input("แป้นพิมพ์เป็น ENG และเปิดหน้าข้อมูลบุคคลของ JHCIS ไว้ แล้วกด Enter (มีเวลา 3 วินาที)")
    time.sleep(3)
    def owner_shot():
        """ภาพช่อง 'ข้อมูลของ' (ใต้ช่องบัตร ปชช.) ใช้ดูว่าค้นแล้วเปลี่ยนคนจริง"""
        try:
            ix, iy = at("id_field")
            return pyautogui.screenshot(region=(ix + 15, iy + 26, 230, 18)).convert("L")
        except Exception:
            return None

    def changed(a_img, b_img):
        """ชื่อในช่อง 'ข้อมูลของ' เปลี่ยนจริงไหม: นับจุดที่ต่างกันชัด ๆ (เคอร์เซอร์/ขอบช่องกระพริบ ไม่นับ)"""
        if a_img is None or b_img is None:
            return True
        n = sum(1 for x, y in zip(a_img.tobytes(), b_img.tobytes()) if abs(x - y) > 60)
        return n >= 40

    def activate():
        w = pygetwindow.getActiveWindow()
        if not (w and (w.title or "").startswith(WIN_KEY)):
            ms = [x for x in pygetwindow.getAllWindows() if (x.title or "").startswith(WIN_KEY)]
            if ms:
                try:
                    ms[0].activate()
                except Exception:
                    pass
                time.sleep(0.8)

    log, done = [], 0
    pre = {}                                                   # อ่านสถานะทุกคนในครั้งเดียว (เร็วกว่าถามฐานข้อมูลทีละคน)
    ids = sorted({int(t["pid"]) for t in todo if str(t["pid"]).isdigit()})
    for i in range(0, len(ids), 400):
        for r in rc.q("SELECT pid,villcode,schoolno,classeducate,classroom,yeareducate FROM personstudent WHERE pid IN ("
                      + ",".join(map(str, ids[i:i + 400])) + ")", pw):
            pre[r["pid"]] = r
    for t in todo:
        if a.limit and done >= a.limit:
            break
        if t["code"] is None:
            log.append((t, "ข้าม-ไม่รู้รหัสชั้น")); continue
        if is_ok(pre.get(str(t["pid"])), t):
            log.append((t, "ถูกอยู่แล้ว-ข้าม")); continue
        by_hn = not re.fullmatch(r"\d{13}", str(t["id"]))       # เลข G... ช่องค้นบัตร ปชช. ของ JHCIS รับเฉพาะตัวเลข (ตัว G หาย)
        if by_hn and "hn_field" not in P:
            log.append((t, "ข้าม-เลขบัตรมีตัวอักษร ทำมือ (ค้นด้วย HN)")); continue
        done += 1
        activate()
        before = owner_shot()
        found, stop = False, None
        for attempt in range(3):                               # ค้นไม่ติด (JHCIS ช้า) -> ลองใหม่ รอนานขึ้นทีละรอบ
            ex = attempt * 1.5
            activate()
            ix, iy = at("hn_field" if by_hn else "id_field")
            pyautogui.click(ix, iy); time.sleep(0.3 + ex / 3)
            pyautogui.press("end"); pyautogui.press("backspace", presses=16, interval=0.01)   # ล้างค่าเดิมให้หมด
            pyautogui.press("delete", presses=4, interval=0.01)
            pyautogui.write(t["pid"] if by_hn else t["id"], interval=0.03); pyautogui.press("enter")
            time.sleep(a.wait + 0.5 + ex)
            bad = clear_popups()
            if bad:
                stop = f"หลังค้นเลขบัตร มีหน้าต่าง '{bad}'"; break
            if "first_row" in P and (not by_hn or P.get("hn_list")):
                pyautogui.doubleClick(*at("first_row")); time.sleep(a.wait + ex)
            bad = clear_popups()                               # เปิดข้อมูลคนแล้วอาจมีเตือน (เช่น สิทธิขัดแย้งฯ)
            if bad:
                stop = f"หลังเปิดข้อมูล มีหน้าต่าง '{bad}'"; break
            if by_hn:                                          # ล้างช่อง HN ไม่ให้ค้างไปรบกวนคนถัดไป
                pyautogui.click(*at("hn_field")); time.sleep(0.2)
                pyautogui.press("end"); pyautogui.press("backspace", presses=12, interval=0.02)
            if changed(before, owner_shot()):
                found = True; break
            print(f"  ... ค้น {t['name']} ไม่ติด ลองใหม่ ({attempt + 2}/3)")
        if stop:
            print(f"\n*** หยุด: {stop}"); log.append((t, "หยุด-" + stop)); break
        if not found:                                          # ยังไม่ได้แตะข้อมูลใคร -> ข้ามคนนี้ไปทำคนถัดไป (ไม่หยุดทั้งรอบ)
            print(f"  ✗ ข้าม {t['name']}: ค้น 3 ครั้งแล้วชื่อในช่อง 'ข้อมูลของ' ไม่เปลี่ยน (ไม่ได้บันทึก)")
            log.append((t, "ข้าม-ค้นไม่ติด 3 ครั้ง")); continue
        pyautogui.click(*at("tab")); time.sleep(0.8)
        pick("school", a.school_pos); time.sleep(0.4)
        pick("cls", a.class1_pos + int(t["code"]) - 1)
        ref = cls_ref.get(t["code"])                           # ภาพช่องชั้นเรียนของคนก่อนหน้าที่ชั้นเดียวกันและบันทึกผ่านแล้ว
        if ref is not None:
            time.sleep(0.3)
            if cls_diff(ref, cls_shot()):                      # ชั้นที่เลือกได้ไม่เหมือนที่ควรเป็น -> เลือกใหม่ 1 ครั้ง
                print(f"  ... ชั้นที่เลือกได้ไม่ตรง ({t['name']}) เลือกใหม่")
                pyautogui.press("esc"); time.sleep(0.3)
                pick("cls", a.class1_pos + int(t["code"]) - 1); time.sleep(0.3)
                if cls_diff(ref, cls_shot()):
                    print(f"  ✗ ข้าม {t['name']}: เลือกชั้นไม่ได้ตรง (ไม่ได้บันทึก)")
                    pyautogui.press("esc")
                    log.append((t, "ข้าม-บันทึกไม่ติด เลือกชั้นไม่ตรง")); continue
        for k, v in (("room", t["room"]), ("year", str(a.year))):
            pyautogui.click(*at(k)); time.sleep(0.2)
            pyautogui.press("end"); pyautogui.press("backspace", presses=12, interval=0.02)   # ล้างค่าเดิม (Ctrl+A ใช้ไม่ได้ในช่องนี้)
            pyautogui.press("delete", presses=12, interval=0.02)
            pyautogui.write(v, interval=0.03); pyautogui.press("tab"); time.sleep(0.2)
        okb = "บันทึก" if a.save else "ไม่บันทึก ไปคนถัดไป (ซ้อม)"
        ans = okb if (a.auto and a.save) else ask(
            f"{t['name']}\nเลข {t['id']}\n\nต้องเป็น: {school} / {t['cls']} / ห้อง {t['room']} / ปี {a.year}\n\n"
            "ดูใน JHCIS ว่าเป็นเด็กคนนี้ และค่าที่เลือกถูกต้อง", [okb, "ข้ามคนนี้", "หยุด"])
        if ans in (None, "หยุด"):
            log.append((t, "หยุดโดยผู้ใช้")); break
        if not (a.save and ans == okb):
            log.append((t, f"ไม่บันทึก ({ans})")); continue      # ไม่กด เพิ่มคนใหม่ (จะสร้างคนเปล่า) คนถัดไปค้นทับได้เลย
        time.sleep(0.4)
        activate()                                             # หลังกดปุ่มในหน้าต่างยืนยัน หน้าต่างดำอาจอยู่ข้างหน้า
        bad = clear_popups()
        w = pygetwindow.getActiveWindow()
        if bad or not (w and (w.title or "").startswith(WIN_KEY)):
            print(f"\n*** หยุด: ก่อนกดบันทึก มีหน้าต่าง '{bad or (w.title if w else '')}' ไม่บันทึก")
            log.append((t, "หยุด-มีหน้าต่างก่อนบันทึก")); break
        pyautogui.click(*at("save")); time.sleep(1.0)
        pyautogui.click(*at("disk")); time.sleep(a.wait + 0.5)
        clear_popups()
        st = state(t["pid"])
        if is_ok(st, t):
            log.append((t, "บันทึกแล้ว-ตรวจฐานข้อมูลผ่าน")); print(f"  ✓ {t['name']} -> {t['cls']}")
            if t["code"] not in cls_ref:
                cls_ref[t["code"]] = cls_shot()
        elif st and pre.get(str(t["pid"])) and all(str(st.get(k)) == str(pre[str(t["pid"])].get(k))
                 for k in ("villcode", "schoolno", "classeducate", "classroom", "yeareducate")):                    # ค่าในฐานเหมือนก่อนเริ่มทุกช่อง = บันทึกไม่ติด (เช่น หน้าต่างเตือนเด้งมาขัด) ไม่มีอะไรเสีย
            log.append((t, "ข้าม-บันทึกไม่ติด (ค่าในฐานไม่เปลี่ยน)"))
            print(f"  ✗ ข้าม {t['name']}: บันทึกไม่ติด ค่าในฐานยังเป็นค่าเดิม (รันซ้ำทีหลังได้)")
            pyautogui.press("esc"); time.sleep(0.5); clear_popups()
        else:
            log.append((t, f"ตรวจฐานข้อมูลไม่ผ่าน: {st}"))
            print(f"\n*** หยุด: {t['name']} (pid {t['pid']}) บันทึกแล้วแต่ค่าในฐานข้อมูลไม่ตรงเป้า\n    ที่ได้: {st}\n    ตรวจใน JHCIS ด้วยตัวเองว่าบันทึกไปที่ใคร/ค่าอะไร ก่อนรันต่อ")
            break
    with open(os.path.join(HERE, "jhcis_student_log.txt"), "a", encoding="utf-8") as f:
        for t, s in log:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{srv}\t{school}\t{t['pid']}\t{t['name']}\t{t['cls']}\t{s}\n")
    from collections import Counter
    manual = [t for t, x in log if x.startswith("ข้าม-เลขบัตรมีตัวอักษร") or x.startswith("ข้าม-ค้นไม่ติด") or x.startswith("ข้าม-บันทึกไม่ติด")]
    if manual:
        print("\nต้องทำมือ หรือรันคำสั่งเดิมซ้ำ (ค้นด้วย HN ใน JHCIS ได้):")
        for t in manual:
            print(f"  HN {t['pid']}  {t['name']}  -> {t['cls']} ห้อง {t['room']} ปี {a.year}")
    print("\nสรุป:", dict(Counter(s.split(":")[0].split(" (")[0] for _, s in log)), "\nรายละเอียดใน jhcis_student_log.txt")


def cli():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("calibrate")
    sub.add_parser("calibrate-hn")
    r = sub.add_parser("run")
    r.add_argument("xlsx")
    r.add_argument("--school-pos", type=int, required=True)
    r.add_argument("--class1-pos", type=int, default=3)
    r.add_argument("--room", type=int, default=1)
    r.add_argument("--year", type=int, default=None, help="ปีการศึกษา พ.ศ. (ไม่ใส่ = คำนวณจากวันที่: เปิดเทอม พ.ค.)")
    r.add_argument("--limit", type=int, default=0)
    r.add_argument("--save", action="store_true")
    r.add_argument("--auto", action="store_true")
    r.add_argument("--real", action="store_true")
    r.add_argument("--user", default="root")
    r.add_argument("--delay", type=float, default=0.2)
    r.add_argument("--wait", type=float, default=1.5)
    a = ap.parse_args()
    if getattr(a, "year", 0) is None:
        import datetime as _dt
        _d = _dt.date.today()
        a.year = _d.year + 543 - (0 if _d.month >= 5 else 1)   # ปีการศึกษาไทยเริ่มเดือนพฤษภาคม
        print(f"ปีการศึกษา: {a.year} (คำนวณจากวันที่ {_d:%d/%m/%Y}  ถ้าไม่ถูกให้ใส่ --year เอง)")
    {"calibrate": calibrate, "calibrate-hn": calibrate_hn}.get(a.cmd, lambda: run(a))()


if __name__ == "__main__":
    cli()
