# -*- coding: utf-8 -*-
"""
jhcis_dental_room.py  v1.1  ลงผลตรวจฟันทั้งห้องในหน้า "บันทึกผลงานอนามัยโรงเรียน ตรวจทันตกรรม" ของ JHCIS
อ่านผลจากไฟล์ "ลงJHCIS_โรงเรียน.csv" (แอปกรอกผลตรวจฟัน > ⋯ เพิ่มเติม > ไฟล์สำหรับลง JHCIS)

ต่างจาก jhcis_dental_fill.py เดิม:
- จับคู่ด้วย "ชื่อในฟอร์มที่เปิด" ไม่ใช่ลำดับแถว  ลำดับใน JHCIS กับในไฟล์ไม่ต้องตรงกัน
- เลื่อนตารางลงเองจนครบทั้งห้อง
- ข้ามเด็กที่ไม่มา / ไม่มีในไฟล์ แล้วสรุปรายชื่อให้ตอนจบ

ใช้ตำแหน่งช่องที่จดไว้แล้วใน jhcis_dental_points.json และต้องมี jhcis_dental_fill.py อยู่โฟลเดอร์เดียวกัน

วิธีใช้ (เปิดหน้า ตรวจทันตกรรม ของห้องนั้นค้างไว้ เลื่อนตารางไว้บนสุด แป้นพิมพ์ EN)
  ง่ายสุด: ดับเบิลคลิก run_room.bat แล้วพิมพ์ห้อง (a2 = อ2, p1 = ป1)
  ซ้อม ไม่บันทึก:     python jhcis_dental_room.py a2
  บันทึกจริง:         python jhcis_dental_room.py a2 --save
  ไม่ถามทุกคน:        เติม --auto  (ถาม 3 คนแรกให้ตรวจก่อน แล้วทำต่อเอง)
  หยุดฉุกเฉิน: ลากเมาส์ไปมุมซ้ายบนสุดของจอ
"""
import argparse
import csv
import os
import re
import sys
import time

import jhcis_dental_fill as jf

try:                                              # ในโปรแกรมสำเร็จรูป เก็บบันทึกไว้ใน "ข้อมูลเครื่อง_..." ข้าง .exe
    import jhcis_roster_compare as _rc
    HERE = _rc.home()
except Exception:
    HERE = os.path.dirname(os.path.abspath(__file__))
PRE = re.compile(r"^(ด\.?\s*ช\.?|ด\.?\s*ญ\.?|เด็กชาย|เด็กหญิง|นางสาว|น\.?\s*ส\.?|นาย|นาง)")
NUMS = [("d_have", "DTEETH"), ("d_decay", "DCARIES"), ("d_filled", "DFILLING"), ("d_lost", "DEXTRACT"),
        ("d_needfill", "NEED_DFILLING"), ("p_have", "PTEETH"), ("p_decay", "PCARIES"), ("p_filled", "PFILLING"),
        ("p_lost", "PEXTRACT"), ("p_needfill", "NEED_PFILLING"), ("p_sealant", "NEED_SEALANT"),
        ("p_rct", "NEED_PEXTRACT"), ("d_rct", "NEED_DEXTRACT")]


def norm(s):
    s = re.sub(r"\s+", "", str(s or ""))
    s = PRE.sub("", s)
    return s.replace("-", "")


def n0(v):
    v = str(v or "").strip()
    if v == "":
        return "0"
    if not v.isdigit():
        raise ValueError(f"ค่า '{v}' ไม่ใช่ตัวเลข")
    return v


def load(path, room):
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    if not rows or "CID" not in rows[0]:
        sys.exit("ไฟล์นี้ไม่ใช่ไฟล์ ลงJHCIS_*.csv จากแอปกรอกผล")
    rooms = sorted({r["ROOM"] for r in rows})
    if not any(r["ROOM"] == room for r in rows):
        sys.exit(f"ไม่พบห้อง '{room}' ในไฟล์ มีห้อง: {', '.join(rooms)}")
    kids = []
    for r in rows:                                # โหลดทั้งโรงเรียน เผื่อ JHCIS ยังใส่เด็กไว้ผิดชั้น
        k = {"name": f"{r['PRENAME']}{r['NAME']} {r['LNAME']}".strip(), "full": norm(r["NAME"] + r["LNAME"]),
             "room": r["ROOM"], "ln": norm(r["LNAME"]),
             "fn": norm(r["NAME"]), "absent": r["ABSENT"] == "1", "err": None, "done": False}
        try:
            k["nums"] = [(key, n0(r[col])) for key, col in NUMS] + [("occl", "0")]
            tartar = n0(r["TARTAR"]) != "0"
            k["combos"] = [("gum", jf.GUM_YES if n0(r["GUM"]) != "0" else "0"),
                           ("tartar", jf.TARTAR_YES if tartar else "0"),
                           ("need_fluoride", jf.FLUORIDE),
                           ("need_scaling", jf.SCALING_YES if tartar else jf.SCALING_NO)]
        except ValueError as e:
            k["err"] = str(e)
        kids.append(k)
    return kids


def match(kids, who, room):
    """คืน (เด็ก, วิธีจับคู่)  ลำดับ: ชื่อ-สกุลตรงในห้อง > ตรงทั้งโรงเรียน > ชื่ออย่างเดียวในห้อง > สะกดต่างนิดหน่อย"""
    from difflib import SequenceMatcher as SM
    w = norm(who)
    parts = re.sub(r"\s+", " ", PRE.sub("", str(who).strip())).strip().split(" ")
    wfn = norm(parts[0])
    wln = norm(parts[-1]) if len(parts) > 1 else ""
    same = [k for k in kids if k["room"] == room]
    for pool, other in ((same, False), (kids, True)):
        hit = [k for k in pool if k["full"] and k["full"] == w]
        if len(hit) == 1:
            return hit[0], (f"ชั้นในไฟล์ {hit[0]['room']}" if other else "")
        if len(hit) > 1:
            return "many", ""
    hit = [k for k in same if k["fn"] and k["fn"] == wfn and (not k["ln"] or not wln)]
    if len(hit) == 1:
        return hit[0], ""
    # สะกดต่างกันนิดหน่อย เช่น นัธวัฒน์ / นัธทวัฒน์: นามสกุลต้องตรง และชื่อคล้ายกันมาก
    hit = [k for k in kids if wln and k["ln"] == wln and SM(None, k["fn"], wfn).ratio() >= 0.8]
    if not hit:
        hit = [k for k in same if not wln and not k["ln"] and SM(None, k["fn"], wfn).ratio() >= 0.85]
    if len(hit) == 1:
        return hit[0], f"สะกดต่าง ในไฟล์: {hit[0]['name']} ({hit[0]['room']})"
    return (None if not hit else "many"), ""


SB_DEFAULT = {"sb_up": [1246, 125], "sb_down": [1246, 646]}   # ปุ่มลูกศรแถบเลื่อน วัดจากมุมหน้าต่างตาราง


def console(show):
    """ย่อ/คืนหน้าต่างดำของโปรแกรมนี้ (Windows)"""
    try:
        import ctypes
        h = ctypes.windll.kernel32.GetConsoleWindow()
        if h:
            ctypes.windll.user32.ShowWindow(h, 9 if show else 6)   # 9 = คืน, 6 = ย่อ
    except Exception:
        pass


def table_front():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if jf.TABLE_KEY in (w.title or "")]
    if ws:
        try:
            ws[0].activate()
        except Exception:
            pass
        time.sleep(0.3)


def scroll_rows(P, n):
    """n > 0 เลื่อนลง n แถว, n < 0 เลื่อนขึ้น"""
    import pyautogui
    to = jf.table_origin()
    if not to:
        return
    table_front()
    key = "sb_down" if n > 0 else "sb_up"
    up, dn = P.get("sb_up"), P.get("sb_down")
    ok = up and dn and dn[1] - up[1] > 200 and abs(dn[0] - up[0]) < 30   # จดผิด (จุดเดียวกัน) ให้ใช้ค่าตั้งต้น
    x, y = (P[key] if ok else SB_DEFAULT[key])
    old = pyautogui.PAUSE
    pyautogui.PAUSE = 0.05
    for _ in range(abs(n)):
        pyautogui.click(to[0] + x, to[1] + y)
    pyautogui.PAUSE = old
    time.sleep(0.5)


def calib_scroll():
    """จดตำแหน่งปุ่มลูกศร ขึ้น/ลง ของแถบเลื่อนตาราง (ครั้งเดียว)"""
    import json, pyautogui
    if not jf.table_origin():
        sys.exit("เปิดหน้า ตรวจทันตกรรม ค้างไว้ก่อน")
    P = json.load(open(jf.CAL, encoding="utf-8"))
    for key, desc in (("sb_up", "ปุ่มลูกศรขึ้น ▲ บนสุดของแถบเลื่อนตาราง"), ("sb_down", "ปุ่มลูกศรลง ▼ ล่างสุดของแถบเลื่อนตาราง")):
        input(f"ชี้เมาส์ที่ {desc} (ไม่ต้องคลิก) แล้วกลับมากด Enter ")
        x, y = pyautogui.position(); ox, oy = jf.table_origin()
        P[key] = [x - ox, y - oy]; print("   =", P[key])
    if P["sb_down"][1] - P["sb_up"][1] < 200:
        P.pop("sb_up"); P.pop("sb_down")
        json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)
        sys.exit("สองจุดอยู่ใกล้กันเกินไป (เมาส์ไม่ได้ย้ายไปชี้ที่ลูกศร) จะใช้ตำแหน่งตั้งต้นแทน ลองจดใหม่ได้")
    json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว")


def close_form(P):
    import pyautogui
    if jf.FORM_KEY in jf.active_title() or jf.focus_form():
        pyautogui.click(*jf.fp(P, "close"))
    return jf.form_closed()


def run(path, room, do_save, auto, visible, delay):
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = delay
    P = jf.load_points()
    kids = load(path, room)
    row_h = P["row2_dental"][1] - P["row1_dental"][1]
    mine = [k for k in kids if k["room"] == room]
    todo = [k for k in mine if not k["absent"]]
    print(f"ห้อง {room}: ในไฟล์ {len(mine)} คน | ต้องลง {len(todo)} | ไม่มา {len(mine) - len(todo)}")
    print("โหมด:", "บันทึกจริง" if do_save else "ซ้อม (ไม่กดบันทึก)", "| ถามทุกคน" if not auto else "| ถาม 3 คนแรก")
    input("เปิดหน้า ตรวจทันตกรรม ของห้องนี้ เลื่อนตารางไว้บนสุด แป้นพิมพ์ EN แล้วกด Enter (มีเวลา 3 วินาทีสลับไป JHCIS)")
    time.sleep(3)
    console(False)                                # ย่อหน้าต่างดำลง ไม่ให้บังแถบเลื่อน/ปุ่ม Dental
    seen, log, extra, asked = set(), [], [], 0
    import glob
    already = set()                               # คนที่บันทึกไปแล้วในรอบก่อน ๆ ของห้องนี้ ไม่บันทึกซ้ำ
    for f in glob.glob(os.path.join(HERE, f"ลงJHCIS_ผล_{room}_*.txt")):
        for line in open(f, encoding="utf-8"):
            nm, _, st = line.rstrip("\n").partition("\t")
            if st.startswith("บันทึกแล้ว"):
                already.add(norm(nm))
    if already:
        print(f"บันทึกไปแล้วในรอบก่อน {len(already)} คน จะข้ามคนเหล่านี้")
    stop, page = False, 0
    scroll_rows(P, -60)                           # เลื่อนตารางขึ้นบนสุดก่อนเริ่ม
    while not stop:
        new_here = 0
        for slot in range(visible):
            to = jf.table_origin()
            if not to:
                print("ไม่พบหน้าต่างตาราง ตรวจทันตกรรม -> หยุด"); stop = True; break
            pyautogui.click(to[0] + P["row1_dental"][0], to[1] + P["row1_dental"][1] + row_h * slot)
            title = jf.wait_form(3.0)
            if not title:                         # คลิกแถวว่าง = สุดรายชื่อแล้ว
                if slot == 0 and page == 0:
                    print("คลิกแถวแรกแล้วไม่มีฟอร์มขึ้น -> หยุด ตรวจตำแหน่งด้วย python jhcis_dental_fill.py check")
                stop = True; break
            who = title.split("ของ :", 1)[-1].strip(" :")
            key = norm(who)
            if key in seen:                       # แถวที่ทำไปแล้วก่อนเลื่อนตาราง
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); stop = True; break
                continue
            seen.add(key); new_here += 1
            if key in already:
                log.append((who, "บันทึกแล้ว (รอบก่อน)")); print(f"= {who}: บันทึกไว้แล้วรอบก่อน ข้าม")
                for k0 in kids:
                    if match(kids, who, room)[0] is k0: k0["done"] = True
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); stop = True; break
                continue
            k, how = match(kids, who, room)
            if how:
                print(f"  * {who}: จับคู่แบบ {how}")
            if k is None or k == "many":
                extra.append(who); log.append((who, "ไม่มีในไฟล์" if k is None else "ชื่อซ้ำในไฟล์ ไม่ลง"))
                print(f"- {who}: {'ไม่มีในไฟล์ห้องนี้' if k is None else 'ชื่อซ้ำในไฟล์'} ข้าม")
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); stop = True; break
                continue
            k["done"] = True
            if k["absent"] or k["err"]:
                st = "ไม่มาตรวจ ข้าม" if k["absent"] else "ข้อมูลผิด: " + k["err"]
                log.append((who, st)); print(f"- {who}: {st}")
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); stop = True; break
                continue
            for f, v in k["combos"]:
                pyautogui.click(*jf.fp(P, f)); pyautogui.write(v); pyautogui.press("enter")
            for f, v in k["nums"]:
                pyautogui.click(*jf.fp(P, f)); pyautogui.hotkey("ctrl", "a"); pyautogui.write(v)
            pyautogui.click(*jf.fp(P, "cpi_all_normal"))
            ok = "บันทึก" if do_save else "ปิดโดยไม่บันทึก (ซ้อม)"
            ans = ok
            if not auto or asked < 3:
                asked += 1
                ans = jf.ask(f"{who}\n\nกรอกเสร็จแล้ว ดูค่าในฟอร์ม JHCIS ให้ถูกต้องก่อนกด", [ok, "ข้ามคนนี้", "หยุด"])
                time.sleep(0.4)
            if not jf.focus_form():
                print(f"{who}: ดึงฟอร์มขึ้นหน้าสุดไม่ได้ -> หยุด ไม่กดบันทึก"); log.append((who, "หยุด")); stop = True; break
            if ans in (None, "หยุด"):
                close_form(P); log.append((who, "หยุดโดยผู้ใช้")); stop = True; break
            if do_save and ans == ok:
                pyautogui.click(*jf.fp(P, "save")); time.sleep(1.5); log.append((who, "บันทึกแล้ว" + (f" ({how})" if how else "")))
                print(f"✓ {who}")
            else:
                log.append((who, f"ไม่บันทึก ({ans})")); print(f"  {who}: ไม่บันทึก")
            if jf.FORM_KEY in jf.active_title() and not close_form(P):
                print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); stop = True; break
            time.sleep(0.3)
        if stop or new_here == 0:
            break
        # เลื่อนตารางลงทีละแถวด้วยปุ่มลูกศรของแถบเลื่อน (ล้อเมาส์ใช้กับ JHCIS ไม่ได้)
        # ทับแถวเก่า 2 แถว คนที่เปิดดูแล้วจะข้ามเอง
        scroll_rows(P, visible - 2)
        page += 1
    console(True)
    missing = [k["name"] for k in kids if k["room"] == room and not k["done"] and not k["absent"]]
    out = os.path.join(HERE, f"ลงJHCIS_ผล_{room}_{time.strftime('%Y%m%d_%H%M')}.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"ไฟล์ {os.path.basename(path)} ห้อง {room} โหมด {'บันทึกจริง' if do_save else 'ซ้อม'}\n\n")
        for who, st in log:
            f.write(f"{who}\t{st}\n")
        f.write(f"\nมีในไฟล์ แต่ไม่เจอในตาราง JHCIS ({len(missing)}):\n" + "\n".join(missing) + "\n")
    saved = sum(1 for _, st in log if st.startswith("บันทึกแล้ว") and "รอบก่อน" not in st)
    print("\n" + "=" * 60)
    print(f" ห้อง {room}: บันทึก {saved} คน | ไม่มีในไฟล์ {len(extra)} | ไม่เจอในตาราง {len(missing)}")
    if missing:
        print(" ไม่เจอในตาราง JHCIS (ต้องลงเอง หรือเช็กชั้นเรียน):", ", ".join(missing))
    print(" สรุปอยู่ที่:", out)
    print("=" * 60)


def room_name(r):
    """พิมพ์ห้องเป็นอังกฤษได้: a2/o2 = อ2, p1 = ป1, m1 = ม1 (ไม่ต้องสลับแป้นภาษา)"""
    m = re.match(r"^\s*([aopmAOPM])\.?\s*(\d+)\s*$", r or "")
    if not m:
        return (r or "").strip()
    return {"a": "อ", "o": "อ", "p": "ป", "m": "ม"}[m.group(1).lower()] + m.group(2)


def newest_csv():
    import glob
    fs = glob.glob(os.path.join(HERE, "ลงJHCIS_*.csv")) + glob.glob(os.path.join(os.getcwd(), "ลงJHCIS_*.csv"))
    if not fs:
        sys.exit("ไม่พบไฟล์ ลงJHCIS_*.csv ในโฟลเดอร์นี้ ดาวน์โหลดจากแอปกรอกผลแล้ววางไว้ที่นี่")
    return max(fs, key=os.path.getmtime)


def cli(pause=False):
    ap = argparse.ArgumentParser()
    ap.add_argument("room", help="ห้อง เช่น อ2 ป1 หรือพิมพ์ a2 p1 (หรือ calib = จดตำแหน่งแถบเลื่อน)")
    ap.add_argument("--csv", help="ไฟล์ ลงJHCIS_*.csv (ไม่ใส่ = ใช้ไฟล์ล่าสุดในโฟลเดอร์นี้)")
    ap.add_argument("--save", action="store_true")
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--visible", type=int, default=17, help="จำนวนแถวที่เห็นในตารางโดยไม่เลื่อน")
    ap.add_argument("--delay", type=float, default=0.25)
    try:
        if len(sys.argv) > 1 and sys.argv[1] == "calib":
            calib_scroll()
        else:
            a = ap.parse_args()
            path = a.csv or newest_csv()
            print("ไฟล์:", os.path.basename(path))
            run(path, room_name(a.room), a.save, a.auto, a.visible, a.delay)
    except BaseException as e:            # รวม sys.exit ด้วย จะได้เห็นข้อความก่อนหน้าต่างปิด
        import traceback
        console(True)
        print("\nหยุด:", e)
        try:
            with open(os.path.join(HERE, "run_room_error.txt"), "w", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n" + traceback.format_exc())
        except Exception:
            pass
    if pause:
        input("\nกด Enter เพื่อปิด")


if __name__ == "__main__":
    cli(pause=True)
