# -*- coding: utf-8 -*-
"""
jhcis_dental_room.py  v1.2  ลงผลตรวจฟันทั้งห้องในหน้า "บันทึกผลงานอนามัยโรงเรียน ตรวจทันตกรรม" ของ JHCIS
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
import json
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
    if room is not None and not any(r["ROOM"] == room for r in rows):
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
    same = [k for k in kids if k["room"] == room] if room else kids
    for pool, other in ((same, False), (kids, True)):
        hit = [k for k in pool if k["full"] and k["full"] == w]
        if len(hit) == 1:
            return hit[0], (f"ชั้นในไฟล์ {hit[0]['room']}" if other and room else "")
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


REF = os.path.join(HERE, "dental_provider_ok.png")      # ภาพช่องผู้ให้บริการในฟอร์ม ที่คนยืนยันแล้วตอนลงคนแรก
ITEM = os.path.join(HERE, "dental_provider_item.png")   # ภาพชื่อทันตแพทย์ในรายการ (ฟอร์ม)
TREF_OK = os.path.join(HERE, "dental_tprov_ok.png")     # ภาพช่องในตารางที่เป็นชื่อที่ต้องการ
TREF_DEF = os.path.join(HERE, "dental_tprov_def.png")   # ภาพช่องในตารางที่เป็นชื่อตั้งต้น (ชื่อคนล็อกอิน)
TITEM = os.path.join(HERE, "dental_tprov_item.png")     # ภาพชื่อทันตแพทย์ในรายการ (ตาราง)
IW, IH = 150, 16                                        # ขนาดภาพชื่อในรายการ (ฟอร์ม)
TIW = 54                                                # รายการในตารางแคบ เท่าความกว้างคอลัมน์


def ink(img):
    """เหลือแต่ตัวหนังสือ (ขาวดำ) ไม่ขึ้นกับสีพื้น/แถบเลือก"""
    return img.convert("L").point(lambda v: 255 if v < 110 else 0)


def tsame(a, b, lim=0.15):
    return score(a, b) < lim


def score(a, b):
    from PIL import ImageChops
    if a.size != b.size:
        return 9
    diff = ImageChops.difference(a, b).histogram()[255]
    n = max(ImageChops.lighter(a, b).histogram()[255], 1)
    return diff / n


def grab(x, y, w, h):
    import pyautogui
    return ink(pyautogui.screenshot(region=(int(x), int(y), int(w), int(h))))


def find_item(ax, ay, dx, tmpl):
    """หาแถวในรายการที่เปิดอยู่ (ใต้หรือเหนือจุด ax,ay) ที่ตัวหนังสือตรงกับภาพ tmpl คืน y บนจอ หรือ None"""
    from PIL import Image, ImageChops
    w, h = tmpl.size
    # เส้นตั้ง (ขอบรายการ/แถบเลื่อน) มีทุกแถว ใช้แยกชื่อไม่ได้ ไม่นับ
    keep = Image.new("L", (w, h), 255)
    for c in range(w):
        if sum(1 for r in range(h) if tmpl.getpixel((c, r))) >= h * 0.75:
            keep.paste(0, (c, 0, c + 1, h))
    t = ImageChops.darker(tmpl, keep)
    x0, top = ax + dx - 6, ay - 330
    strip = grab(x0, top, w, 660)
    best, by = 0.2, None
    for y in range(0, 660 - h):
        if abs(top + y + h // 2 - ay) < 14:               # ข้ามตัวช่องเอง (แสดงชื่อปัจจุบัน)
            continue
        sc = score(ImageChops.darker(strip.crop((0, y, w, y + h)), keep), t)
        if sc < best:
            best, by = sc, top + y + h // 2
    return by


def grab_item(x, y, w):
    """จับภาพชื่อในรายการรอบจุดที่ชี้ คืน None ถ้าแทบไม่มีตัวหนังสือ (รายการไม่ได้เปิด/ชี้ช่องว่าง)"""
    img = grab(x - 6, y - IH // 2, w, IH)
    mid = img.crop((0, 3, w, IH - 3)).histogram()[255]
    return img if mid >= 20 else None


def pick(ax, ay, dx, tmpl_path):
    """คลิกช่องให้รายการเปิด แล้วคลิกชื่อที่ตรงภาพ (เลื่อนรายการด้วยลูกศรถ้ายังไม่เห็น) คืน True ถ้าเจอ"""
    import pyautogui
    from PIL import Image
    tmpl = Image.open(tmpl_path).convert("L")
    pyautogui.click(ax, ay); time.sleep(0.6)
    for i in range(14):
        y = find_item(ax, ay, dx, tmpl)
        if y is not None:
            pyautogui.click(ax + dx + 30, y); time.sleep(0.5)
            return True
        old = pyautogui.PAUSE; pyautogui.PAUSE = 0.03
        for _ in range(4):                               # เลื่อนรายการลงทีละ 4 ชื่อ แล้วหาใหม่
            pyautogui.press("down" if i < 9 else "up")
        pyautogui.PAUSE = old
        time.sleep(0.3)
    pyautogui.press("escape"); time.sleep(0.3)
    return False


def countdown(sec, text):
    print(text)
    for i in range(sec, 0, -1):
        print(f"   {i} ...", end="\r", flush=True); time.sleep(1)
    print(" " * 20)


def prov_box(P):
    x, y = jf.fp(P, "provider")
    return (int(x - 200), int(y - 11), 400, 22)       # แถบชื่อผู้ให้บริการ กว้างพอเห็นชื่อทั้งชื่อ


def prov_img(P):
    import pyautogui
    return pyautogui.screenshot(region=prov_box(P)).convert("RGB")


def same(a, b):
    from PIL import ImageChops, ImageStat
    return a.size == b.size and sum(ImageStat.Stat(ImageChops.difference(a, b)).mean) / 3 < 12


def prov_settled(P):
    """ภาพช่องผู้ให้บริการหลังย้ายโฟกัสออก (JHCIS แสดงชื่อใหม่ตอนช่องเสียโฟกัส)"""
    import pyautogui
    pyautogui.press("tab"); time.sleep(0.4)
    return ink(prov_img(P))


def set_provider(P):
    """ช่องผู้ให้บริการในฟอร์ม  คืน True = ตรงชื่อที่จำไว้แล้ว, "manual" = ให้คนเลือกเอง
    JHCIS มักจำชื่อที่เลือกครั้งล่าสุดไว้ให้ฟอร์มถัดไป เลือกครั้งแรกแล้วคนต่อไปมักตรงเลย"""
    from PIL import Image
    if "provider" not in P:
        return True
    if not os.path.exists(REF):
        return "manual"
    ok = Image.open(REF).convert("L")
    if tsame(ink(prov_img(P)), ok):                     # เปิดมาเป็นชื่อนี้อยู่แล้ว
        return True
    if "prov_item_dx" in P and os.path.exists(ITEM):
        x, y = jf.fp(P, "provider")
        if pick(x, y, P["prov_item_dx"], ITEM) and tsame(prov_settled(P), ok):
            return True
    return "manual"


def calib_provider():
    """จดผู้ให้บริการ: ชี้ช่อง แล้วเปิดรายการเอง ชี้ชื่อทันตแพทย์ค้างไว้ โปรแกรมจำภาพชื่อ"""
    import pyautogui
    if not jf.form_origin():
        sys.exit("เปิดฟอร์ม Dental ของเด็ก 1 คนค้างไว้ก่อน (ฟอร์ม การตรวจสุขภาพช่องปาก)")
    P = json.load(open(jf.CAL, encoding="utf-8"))
    ox, oy = jf.form_origin()
    input("1) ชี้เมาส์ที่ช่อง 'ผู้ให้บริการ' ในฟอร์ม (กลางช่อง ไม่ต้องคลิก) แล้วกลับมากด Enter ")
    x, y = pyautogui.position(); P["provider"] = [x - ox, y - oy]
    bx, by = x, y
    for _ in range(3):
        countdown(15, "2) ภายใน 15 วินาที: คลิกช่องผู้ให้บริการใน JHCIS ให้รายการเปิด เลื่อนจนเห็น 'ท.พ.อำนาจ'\n"
                      "   แล้วชี้เมาส์ค้างไว้ที่ 'กลางตัวหนังสือ' ต้นชื่อ (ตัว ท) ไม่ต้องคลิก รายการต้องเปิดค้างอยู่")
        x, y = pyautogui.position()
        img = grab_item(x, y, IW) if abs(y - by) >= 14 else None
        if img:
            break
        print("   ไม่เห็นตัวหนังสือตรงที่ชี้ (รายการปิดไปแล้ว หรือชี้ที่ตัวช่อง) ลองใหม่")
    else:
        sys.exit("ยังไม่ได้บันทึก")
    img.save(ITEM)
    P["prov_item_dx"] = x - bx
    for k_ in ("provider_steps", "provider_item"):
        P.pop(k_, None)
    json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)
    if os.path.exists(REF):
        os.remove(REF)                                # ภาพช่องจะจดใหม่ตอนลงผลคนแรก (ให้คนยืนยัน)
    print("บันทึกแล้ว  คลิกชื่อในรายการหรือกด Esc ใน JHCIS แล้วปิดฟอร์มนี้โดยไม่บันทึกได้เลย")


# ---------- คอลัมน์ 'ผู้ให้บริการ' ในตาราง ตรวจทันตกรรม (คนละค่ากับช่องในฟอร์ม Dental) ----------
def tpos(P, slot):
    to = jf.table_origin()
    row_h = P["row2_dental"][1] - P["row1_dental"][1]
    return to[0] + P["tprov"][0], to[1] + P["row1_dental"][1] + row_h * slot


def tcell(P, slot):
    """ภาพตัวหนังสือต้นชื่อในช่องผู้ให้บริการของแถว slot (ไม่เอาปุ่มลูกศรด้านขวาของช่อง)"""
    x, y = tpos(P, slot)
    return grab(x - 6, y - 7, 34, 14)


def set_table_provider(P, slot):
    """คืน "skip" = ไม่ได้ตั้งค่าไว้, "already" = เป็นชื่อที่ต้องการอยู่แล้ว, "other" = เป็นชื่ออื่น ไม่แตะ,
    True = เปลี่ยนแล้วตรงภาพ, False = เปลี่ยนแล้วไม่ตรง/หาไม่เจอ"""
    from PIL import Image
    if "tprov" not in P or "tprov_item_dx" not in P or not all(map(os.path.exists, (TREF_OK, TREF_DEF, TITEM))):
        return "skip"
    table_front()
    ok, df = Image.open(TREF_OK).convert("L"), Image.open(TREF_DEF).convert("L")
    now = tcell(P, slot)
    if tsame(now, ok):
        return "already"
    if not tsame(now, df):
        return "other"
    x, y = tpos(P, slot)
    if not pick(x, y, P["tprov_item_dx"], TITEM):
        return False
    time.sleep(0.3)
    return tsame(tcell(P, slot), ok)


def calib_tprov():
    """จดคอลัมน์ผู้ให้บริการในตาราง: ชี้แถวที่ยังเป็นชื่อตั้งต้น แล้วเปิดรายการเอง ชี้ชื่อทันตแพทย์ค้างไว้"""
    import pyautogui
    if not jf.table_origin():
        sys.exit("เปิดหน้า ตรวจทันตกรรม ค้างไว้ก่อน")
    P = json.load(open(jf.CAL, encoding="utf-8"))
    for k_ in ("tprov", "tprov_steps", "tprov_item_dx"):
        P.pop(k_, None)
    print("ย้ายหน้าต่างดำนี้ไปไว้ด้านซ้ายล่าง ไม่ให้บังคอลัมน์ 'ผู้ให้บริการ' ในตาราง")
    input("1) ชี้เมาส์ที่ 'ตัวแรกของชื่อ' ในช่องผู้ให้บริการ ของแถวที่ยังเป็นชื่อตั้งต้น (ตัว น ของ นายไพโรจน์) ไม่ต้องคลิก แล้วกลับมากด Enter ")
    x, y = pyautogui.position(); ox, oy = jf.table_origin()
    row_h = P["row2_dental"][1] - P["row1_dental"][1]
    slot = round((y - oy - P["row1_dental"][1]) / row_h)
    if not (0 <= slot < 40):
        sys.exit("เมาส์ไม่ได้อยู่ในแถวของตาราง ลองใหม่")
    P["tprov"] = [x - ox, 0]
    tcell(P, slot).save(TREF_DEF)
    cx, cy = tpos(P, slot)
    for _ in range(3):
        countdown(15, f"2) ภายใน 15 วินาที: คลิกช่องเดิมนี้ (แถวที่ {slot + 1}) ใน JHCIS ให้รายการเปิด เลื่อนจนเห็น 'ท.พ.อำนาจ'\n"
                      "   แล้วชี้เมาส์ค้างไว้ที่ 'กลางตัวหนังสือ' ต้นชื่อ (ตัว ท) ไม่ต้องคลิก รายการต้องเปิดค้างอยู่")
        x, y = pyautogui.position()
        img = grab_item(x, y, TIW) if abs(y - cy) >= 14 else None
        if img:
            break
        print("   ไม่เห็นตัวหนังสือตรงที่ชี้ (รายการปิดไปแล้ว หรือชี้ที่ตัวช่อง) ลองใหม่")
    else:
        sys.exit("ยังไม่ได้บันทึก")
    img.save(TITEM)
    P["tprov_item_dx"] = x - cx
    pyautogui.click(x + 30, y); time.sleep(0.8)        # เลือกชื่อนี้ให้แถวนี้เลย
    table_front(); time.sleep(0.3)
    img = tcell(P, slot)
    if tsame(img, Image_open(TREF_DEF)):
        sys.exit("ช่องยังเป็นชื่อตั้งต้น (เลือกไม่สำเร็จ หรือหน้าต่างดำบังอยู่) ยังไม่ได้บันทึก ลองใหม่")
    img.save(TREF_OK)
    json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว  ดูช่องในตารางว่าแถวนี้เป็น ท.พ.อำนาจ แล้ว")


def ink_open(path):
    from PIL import Image
    return Image.open(path).convert("L")


def Image_open(path):
    from PIL import Image
    return Image.open(path).convert("L")


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


class St:
    """สถานะร่วมระหว่างห้อง (โหมดทุกห้องใช้ต่อกัน)"""
    def __init__(self, do_save, auto, redo, visible):
        self.do_save, self.auto, self.redo, self.visible = do_save, auto, redo, visible
        self.log, self.extra, self.asked, self.stop = [], [], 0, False
        self.already = set()
        import glob
        for f in glob.glob(os.path.join(HERE, "ลงJHCIS_ผล_*.txt")):   # คนที่บันทึกไปแล้วรอบก่อน ๆ ไม่บันทึกซ้ำ
            for line in open(f, encoding="utf-8"):
                nm, _, st = line.rstrip("\n").partition("\t")
                if st.startswith("บันทึกแล้ว"):
                    self.already.add(norm(nm))


def table_prov(P, slot, who, S):
    """เปลี่ยนคอลัมน์ผู้ให้บริการในตารางของแถวนี้  คืน False = ให้หยุด"""
    r = set_table_provider(P, slot)
    if r is False:
        ans = jf.ask(f"{who}\n\nคอลัมน์ 'ผู้ให้บริการ' ในตารางของแถวนี้ ไม่ใช่ชื่อที่จดไว้\nแก้เองในตารางแล้วทำต่อ หรือหยุด",
                     ["แก้เองแล้ว ทำต่อ", "หยุด"])
        if ans != "แก้เองแล้ว ทำต่อ":
            S.log.append((who, "หยุด-ผู้ให้บริการในตารางไม่ตรง")); S.stop = True
            return False
    elif r == "other":
        S.log.append((who, "ตาราง: ผู้ให้บริการเป็นชื่ออื่น ไม่ได้เปลี่ยน")); print(f"  {who}: คอลัมน์ผู้ให้บริการในตารางเป็นชื่ออื่น ไม่แตะ")
    return True


def process_table(P, kids, room, S):
    """ทำทุกแถวในหน้าต่าง ตรวจทันตกรรม ที่เปิดอยู่ คืนจำนวนแถวที่เปิดได้ (0 = ห้องว่าง)"""
    import pyautogui
    row_h = P["row2_dental"][1] - P["row1_dental"][1]
    seen, page, opened = set(), 0, 0
    scroll_rows(P, -60)                           # เลื่อนตารางขึ้นบนสุดก่อนเริ่ม
    while not S.stop:
        new_here = 0
        for slot in range(S.visible):
            to = jf.table_origin()
            if not to:
                print("ไม่พบหน้าต่างตาราง ตรวจทันตกรรม -> หยุด"); S.stop = True; break
            pyautogui.click(to[0] + P["row1_dental"][0], to[1] + P["row1_dental"][1] + row_h * slot)
            title = jf.wait_form(3.0)
            if not title:                         # คลิกแถวว่าง = สุดรายชื่อแล้ว
                return opened
            opened += 1
            who = title.split("ของ :", 1)[-1].strip(" :")
            key = norm(who)
            if key in seen:                       # แถวที่ทำไปแล้วก่อนเลื่อนตาราง
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True
                if S.stop: break
                continue
            seen.add(key); new_here += 1
            k, how = match(kids, who, room)
            if key in S.already and not S.redo:
                S.log.append((who, "บันทึกแล้ว (รอบก่อน)")); print(f"= {who}: บันทึกไว้แล้วรอบก่อน ข้าม")
                if isinstance(k, dict): k["done"] = True
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True; break
                if S.do_save and not table_prov(P, slot, who, S):
                    break
                continue
            if how:
                print(f"  * {who}: จับคู่แบบ {how}")
            if k is None or k == "many":
                S.extra.append(who); S.log.append((who, "ไม่มีในไฟล์" if k is None else "ชื่อซ้ำในไฟล์ ไม่ลง"))
                print(f"- {who}: {'ไม่มีในไฟล์' if k is None else 'ชื่อซ้ำในไฟล์'} ข้าม")
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True; break
                continue
            if k["done"]:                         # ลงไปแล้วในรอบนี้ (เด็กอยู่ 2 ชั้นใน JHCIS)
                S.log.append((who, "ลงแล้วในรอบนี้ ข้าม"))
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True; break
                continue
            k["done"] = True
            if k["absent"] or k["err"]:
                st = "ไม่มาตรวจ ข้าม" if k["absent"] else "ข้อมูลผิด: " + k["err"]
                S.log.append((who, st)); print(f"- {who}: {st}")
                if not close_form(P): print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True; break
                continue
            for f, v in k["combos"]:
                pyautogui.click(*jf.fp(P, f)); pyautogui.write(v); pyautogui.press("enter")
            for f, v in k["nums"]:
                pyautogui.click(*jf.fp(P, f)); pyautogui.hotkey("ctrl", "a"); pyautogui.write(v)
            pyautogui.click(*jf.fp(P, "cpi_all_normal"))
            sp = set_provider(P)
            if sp == "manual":                    # ให้คนเลือกเองในฟอร์ม (คนแรก: จำภาพชื่อนี้ไว้เทียบคนต่อไป)
                first = not os.path.exists(REF)
                ans = jf.ask(f"{who}\n\n" + ("คนแรก: " if first else "ช่องผู้ให้บริการไม่ใช่ชื่อที่จำไว้\n") +
                             "เลือก 'ผู้ให้บริการ' ในฟอร์ม JHCIS เป็น ท.พ.อำนาจ เอง\nแล้วกด 'เลือกแล้ว ทำต่อ'",
                             ["เลือกแล้ว ทำต่อ", "หยุด"])
                time.sleep(0.4)
                if ans != "เลือกแล้ว ทำต่อ" or not jf.focus_form():
                    close_form(P); S.log.append((who, "หยุด-ผู้ให้บริการ")); S.stop = True; break
                img = prov_settled(P)
                if first:
                    img.save(REF)
                elif not tsame(img, ink_open(REF)):
                    ans = jf.ask(f"{who}\n\nช่องผู้ให้บริการยังไม่ตรงกับชื่อที่จำไว้ ดูในฟอร์มอีกครั้ง",
                                 ["ถูกแล้ว ทำต่อ", "หยุด"])
                    if ans != "ถูกแล้ว ทำต่อ":
                        close_form(P); S.log.append((who, "หยุด-ผู้ให้บริการ")); S.stop = True; break
                    jf.focus_form()
            ok = "บันทึก" if S.do_save else "ปิดโดยไม่บันทึก (ซ้อม)"
            ans = ok
            if not S.auto or S.asked < 3:
                S.asked += 1
                ans = jf.ask(f"{who}\n\nกรอกเสร็จแล้ว ดูค่าในฟอร์ม JHCIS (รวมช่องผู้ให้บริการ) ให้ถูกต้องก่อนกด", [ok, "ข้ามคนนี้", "หยุด"])
                time.sleep(0.4)
            if not jf.focus_form():
                print(f"{who}: ดึงฟอร์มขึ้นหน้าสุดไม่ได้ -> หยุด ไม่กดบันทึก"); S.log.append((who, "หยุด")); S.stop = True; break
            if ans in (None, "หยุด"):
                close_form(P); S.log.append((who, "หยุดโดยผู้ใช้")); S.stop = True; break
            saved = S.do_save and ans == ok
            if saved:
                pyautogui.click(*jf.fp(P, "save")); time.sleep(1.5)
                S.log.append((who, "บันทึกแล้ว" + (f" ({how})" if how else ""))); print(f"✓ {who}")
            else:
                S.log.append((who, f"ไม่บันทึก ({ans})")); print(f"  {who}: ไม่บันทึก")
            if jf.FORM_KEY in jf.active_title() and not close_form(P):
                print("ปิดฟอร์มไม่สำเร็จ -> หยุด"); S.stop = True; break
            if saved and not table_prov(P, slot, who, S):
                break
            time.sleep(0.3)
        if S.stop or new_here == 0:
            break
        # เลื่อนตารางลงทีละแถวด้วยปุ่มลูกศรของแถบเลื่อน (ล้อเมาส์ใช้กับ JHCIS ไม่ได้)
        scroll_rows(P, S.visible - 2)
        page += 1
    return opened


def report(path, kids, room, S):
    mine = [k for k in kids if room is None or k["room"] == room]
    missing = [k for k in mine if not k["done"] and not k["absent"]]
    label = room or "ทุกห้อง"
    out = os.path.join(HERE, f"ลงJHCIS_ผล_{label}_{time.strftime('%Y%m%d_%H%M')}.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"ไฟล์ {os.path.basename(path)} ห้อง {label} โหมด {'บันทึกจริง' if S.do_save else 'ซ้อม'}\n\n")
        for who, st in S.log:
            f.write(f"{who}\t{st}\n")
        f.write(f"\nมีในไฟล์ แต่ไม่เจอในตาราง JHCIS ({len(missing)}):\n" + "\n".join(f"{k['name']} ({k['room']})" for k in missing) + "\n")
        ab = [k for k in mine if k["absent"]]
        f.write(f"\nไม่มาตรวจ ({len(ab)}) ควรเปลี่ยนเป็น 'ไม่ตรวจ' ในตาราง JHCIS:\n" + "\n".join(f"{k['name']} ({k['room']})" for k in ab) + "\n")
    saved = sum(1 for _, st in S.log if st.startswith("บันทึกแล้ว") and "รอบก่อน" not in st)
    print("\n" + "=" * 60)
    print(f" {label}: บันทึก {saved} คน | ไม่มีในไฟล์ {len(S.extra)} | ไม่เจอในตาราง {len(missing)}")
    if missing:
        print(" ไม่เจอในตาราง JHCIS (ต้องลงเอง หรือเช็กชั้นเรียน):", ", ".join(f"{k['name']} ({k['room']})" for k in missing))
    print(" สรุปอยู่ที่:", out)
    print("=" * 60)


def run(path, room, do_save, auto, visible, delay, redo=False):
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = delay
    P = jf.load_points()
    kids = load(path, room)
    mine = [k for k in kids if k["room"] == room]
    todo = [k for k in mine if not k["absent"]]
    print(f"ห้อง {room}: ในไฟล์ {len(mine)} คน | ต้องลง {len(todo)} | ไม่มา {len(mine) - len(todo)}")
    print("โหมด:", "บันทึกจริง" if do_save else "ซ้อม (ไม่กดบันทึก)", "| ถามทุกคน" if not auto else "| ถาม 3 คนแรก")
    input("เปิดหน้า ตรวจทันตกรรม ของห้องนี้ แป้นพิมพ์ EN แล้วกด Enter (มีเวลา 3 วินาทีสลับไป JHCIS)")
    time.sleep(3)
    console(False)                                # ย่อหน้าต่างดำลง ไม่ให้บังแถบเลื่อน/ปุ่ม Dental
    if not jf.table_origin():                     # ยังไม่ได้กดตรวจทันตกรรม: กดให้ถ้าจดตำแหน่งไว้แล้ว
        if "sch_dent" in P and school_origin() and open_table(P):
            print("เปิดหน้า ตรวจทันตกรรม ให้แล้ว")
        else:
            console(True)
            sys.exit("ยังไม่ได้เปิดหน้า ตรวจทันตกรรม: ในหน้าอนามัยโรงเรียน กดปุ่ม 'ตรวจทันตกรรม' ก่อน แล้วรันใหม่")
    S = St(do_save, auto, redo, visible)
    if S.already and not redo:
        print(f"บันทึกไปแล้วในรอบก่อน {len(S.already)} คน จะข้ามคนเหล่านี้")
    if process_table(P, kids, room, S) == 0 and not S.stop:
        print("คลิกแถวแรกแล้วไม่มีฟอร์มขึ้น: ตารางว่าง หรือตำแหน่งปุ่ม Dental ไม่ตรง (เช็กด้วย python jhcis_dental_fill.py check)")
    console(True)
    report(path, kids, room, S)


# ---------- โหมดทุกห้อง: ปิดตาราง -> เลื่อนชั้นในหน้าอนามัยโรงเรียน -> เปิดตรวจทันตกรรม ----------
SCHOOL_KEY = "อนามัยโรงเรียน(นักเรียน)"


def school_origin():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if SCHOOL_KEY in (w.title or "") and jf.TABLE_KEY not in (w.title or "")]
    return (ws[0].left, ws[0].top) if ws else None


def school_front():
    import pygetwindow
    ws = [w for w in pygetwindow.getAllWindows() if SCHOOL_KEY in (w.title or "") and jf.TABLE_KEY not in (w.title or "")]
    if ws:
        try:
            ws[0].activate()
        except Exception:
            pass
        time.sleep(0.4)


def cls_img(P):
    import pyautogui
    o = school_origin(); x, y = o[0] + P["sch_cls"][0], o[1] + P["sch_cls"][1]
    return pyautogui.screenshot(region=(int(x - 70), int(y - 10), 140, 20)).convert("RGB")


def close_table():
    """ปิดหน้าต่าง ตรวจทันตกรรม ด้วย Alt+F4 (ไม่คลิกใกล้ปุ่ม 'ลบข้อมูลการตรวจทั้งหมด')"""
    import pyautogui
    table_front()
    if jf.TABLE_KEY not in jf.active_title():
        return False
    pyautogui.hotkey("alt", "f4")
    t0 = time.time()
    while time.time() - t0 < 4:
        if not jf.table_origin():
            return True
        time.sleep(0.3)
    return False


def next_class(P):
    """เลือกชั้นถัดไปในรายการชั้นเรียน คืน False ถ้าสุดรายการแล้ว"""
    import pyautogui
    school_front()
    o = school_origin()
    if not o:
        return False
    before = cls_img(P)
    pyautogui.click(o[0] + P["sch_cls"][0], o[1] + P["sch_cls"][1]); time.sleep(0.4)
    pyautogui.press("down"); pyautogui.press("enter"); time.sleep(1.2)
    return not same(before, cls_img(P))


# หน้าต่างเด้งหลังกด 'ตรวจทันตกรรม' (เห็นจากวิดีโอ 26 พ.ย. 2569)
DLG_INFO = ("Please record", "1Hxxx")             # แจ้งเตือน มีปุ่ม OK ปุ่มเดียว
DLG_DATE = "Confirm StudentHealth"                # ถามวันออกอนามัย: ซ้าย=เปลี่ยนวัน (ห้ามกด) ขวา=ยืนยันวันที่
DLG_YES = "Select an Option"                      # ยืนยันการบันทึกตรวจทันตกรรม Yes/No/Cancel


def find_win(*keys):
    import pygetwindow
    for w in pygetwindow.getAllWindows():
        t = w.title or ""
        if any(k in t for k in keys) and w.width > 50:
            return w
    return None


def win_front(w):
    """ทำให้หน้าต่าง w อยู่หน้าสุด คืน True ถ้าเป็นหน้าต่างที่ใช้งานอยู่จริง"""
    import pyautogui
    t = (w.title or "")[:15]
    try:
        w.activate()
    except Exception:
        pass
    time.sleep(0.3)
    if t in jf.active_title():
        return True
    pyautogui.click(w.left + w.width // 2, w.top + 12)   # คลิกแถบชื่อหน้าต่าง
    time.sleep(0.3)
    return t in jf.active_title()


def gone(key, sec=4):
    t0 = time.time()
    while time.time() - t0 < sec:
        if not find_win(key):
            return True
        time.sleep(0.2)
    return False


def save_point(key, val):
    P = json.load(open(jf.CAL, encoding="utf-8"))
    P[key] = val
    json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)


def right_button(img):
    """หาปุ่มขวาในแถวปุ่มของหน้าต่าง 2 ปุ่ม จากเส้นขอบปุ่ม คืน (x, y) ในภาพ หรือ None
    ต้องเจอปุ่มซ้ายที่จบก่อนกลางหน้าต่าง และปุ่มขวาที่เริ่มหลังกลางหน้าต่าง ในแถวเดียวกัน"""
    W, H = img.size
    px = img.load()
    cx = W // 2
    rows = []
    for y in range(max(0, H - 110), H):
        runs, s = [], None
        for x in range(5, W - 5):
            d = sum(px[x, y][:3]) < 600
            if d and s is None:
                s = x
            if (not d or x == W - 6) and s is not None:
                if x - s >= 120:
                    runs.append((s, x))
                s = None
        L = [r for r in runs if cx - 60 < r[1] <= cx + 2]
        R = [r for r in runs if cx - 2 <= r[0] < cx + 60]
        if L and R:
            rows.append((y, R[0]))
    if len(rows) < 2:
        return None
    top, bot = rows[0], rows[-1]
    if not (15 <= bot[0] - top[0] <= 70) or abs(top[1][0] - bot[1][0]) > 4:
        return None
    (x0, x1), y0, y1 = top[1], top[0], bot[0]
    return ((x0 + x1) // 2, (y0 + y1) // 2)


def msg(text):
    """กล่องข้อความบนสุด (ไม่ต้องเปิดหน้าต่างดำที่บังจอ JHCIS)"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, "ลงผลตรวจฟัน JHCIS", 0x40030)   # OK, เตือน, บนสุด
    except Exception:
        console(True); print("\n" + text)


def confirm_date(P, w):
    """กดปุ่มขวา 'ยืนยันวันออกอนามัย นร. วันที่ ...' หาตำแหน่งปุ่มจากภาพทุกครั้ง  ไม่เจอ = ไม่กด"""
    import pyautogui
    if not win_front(w):
        return False
    time.sleep(0.3)
    w = find_win(DLG_DATE) or w
    img = pyautogui.screenshot(region=(w.left, w.top, w.width, w.height)).convert("RGB")
    pt = right_button(img)
    if not pt:
        try:
            img.save(os.path.join(HERE, "dlg_date_notfound.png"))
        except Exception:
            pass
        return False
    pyautogui.click(w.left + pt[0], w.top + pt[1])
    return gone(DLG_DATE)


def pass_dialogs(P, sec=20):
    """ตอบหน้าต่างเด้งจนกว่าตาราง ตรวจทันตกรรม จะขึ้น  เจอหน้าต่างที่ไม่รู้จัก = ไม่กดอะไร"""
    import pyautogui
    t0 = time.time()
    while time.time() - t0 < sec:
        if jf.table_origin():
            time.sleep(1.0); return True
        w = find_win(*DLG_INFO)
        if w:
            if win_front(w):
                pyautogui.press("enter")          # ปุ่ม OK ปุ่มเดียว
                gone(DLG_INFO[0]); t0 = time.time()
            continue
        w = find_win(DLG_DATE)
        if w:
            if not confirm_date(P, w):
                return False
            t0 = time.time(); continue
        w = find_win(DLG_YES)
        if w:
            if win_front(w):
                pyautogui.press("enter")          # Yes เป็นปุ่มตั้งต้น
                gone(DLG_YES); t0 = time.time()
            continue
        time.sleep(0.3)
    return False


def open_table(P):
    import pyautogui
    if not (find_win(*DLG_INFO) or find_win(DLG_DATE) or find_win(DLG_YES)):   # กดเองไปแล้ว ค้างที่หน้าต่างเด้ง
        school_front()
        o = school_origin()
        pyautogui.click(o[0] + P["sch_dent"][0], o[1] + P["sch_dent"][1])
    if pass_dialogs(P):
        return True
    # ไปไม่ถึงตาราง: ให้คนทำต่อเอง (ไม่ต้องกด Enter ในหน้าต่างดำ กัน Enter หลุดไปกดปุ่มใน JHCIS) แล้วโปรแกรมทำงานต่อ
    msg("ยังไปไม่ถึงตาราง 'ตรวจทันตกรรม'\n\nกด OK ที่กล่องนี้ แล้วตอบหน้าต่างใน JHCIS เองจนตารางขึ้น\n"
        "(หน้าต่างถามวันที่: กดปุ่มขวา 'ยืนยันวันออกอนามัย')\n\nโปรแกรมจะรอ 3 นาที แล้วทำต่อเอง")
    t0 = time.time()
    while time.time() - t0 < 180:
        if jf.table_origin():
            time.sleep(1.5); return True
        time.sleep(0.5)
    return False


def calib_school():
    import pyautogui
    if not school_origin():
        sys.exit("เปิดหน้า 'บันทึกข้อมูลงานอนามัยโรงเรียน(นักเรียน)' ค้างไว้ก่อน")
    P = json.load(open(jf.CAL, encoding="utf-8"))
    for key, desc in (("sch_cls", "ช่อง 'ชั้นเรียน' (กลางช่อง ไม่ใช่ลูกศร)"),
                      ("sch_dent", "ปุ่ม 'ตรวจทันตกรรม'")):
        input(f"ชี้เมาส์ที่ {desc} ในหน้าอนามัยโรงเรียน (ไม่ต้องคลิก) แล้วกลับมากด Enter ")
        x, y = pyautogui.position(); ox, oy = school_origin()
        P[key] = [x - ox, y - oy]; print("   =", P[key])
    if abs(P["sch_cls"][0] - P["sch_dent"][0]) + abs(P["sch_cls"][1] - P["sch_dent"][1]) < 30:
        sys.exit("สองจุดอยู่ใกล้กันเกินไป (เมาส์ไม่ได้ย้าย) ลองใหม่")
    json.dump(P, open(jf.CAL, "w", encoding="utf-8"), indent=1)
    print("บันทึกแล้ว")


def run_all(path, classes, do_save, auto, visible, delay, redo=False):
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = delay
    P = jf.load_points()
    if "sch_cls" not in P:
        sys.exit("ยังไม่ได้จดตำแหน่งหน้าอนามัยโรงเรียน: พิมพ์ห้องเป็น school ก่อน (ครั้งเดียว)")
    kids = load(path, None)
    rooms = sorted({k["room"] for k in kids}, key=lambda r: ("อปม".find(r[:1]) if r[:1] in "อปม" else 9, r))
    classes = classes or len(rooms)
    print(f"ในไฟล์ {len(kids)} คน {len(rooms)} ห้อง: {', '.join(rooms)}")
    print(f"จะทำ {classes} ชั้น เริ่มจากชั้นที่เลือกไว้ในหน้าอนามัยโรงเรียนตอนนี้ แล้วเลื่อนลงทีละชั้น")
    print("โหมด:", "บันทึกจริง" if do_save else "ซ้อม (ไม่กดบันทึก)", "| ถามทุกคน" if not auto else "| ถาม 3 คนแรก")
    input("ในหน้าอนามัยโรงเรียน เลือกชั้นแรก (เช่น อนุบาล2) ไว้ แป้นพิมพ์ EN แล้วกด Enter (มีเวลา 3 วินาที)")
    time.sleep(3)
    console(False)
    S = St(do_save, auto, redo, visible)
    for c in range(classes):
        if c > 0:
            if jf.table_origin() and not close_table():
                print("ปิดหน้าต่างตรวจทันตกรรมไม่ได้ -> หยุด"); break
            if not next_class(P):
                print("เลื่อนชั้นต่อไม่ได้ (สุดรายการชั้นแล้ว) -> จบ"); break
        if not jf.table_origin() and not open_table(P):
            print("กดตรวจทันตกรรมแล้วไม่มีหน้าต่างขึ้น -> หยุด"); break
        print(f"\n----- ชั้นที่ {c + 1}/{classes} -----")
        hint = rooms[c] if c < len(rooms) else None   # ชั้นใน JHCIS น่าจะตรงกับห้องลำดับเดียวกันในไฟล์ ใช้ช่วยแยกชื่อซ้ำ
        n = process_table(P, kids, hint, S)
        if n == 0:
            print("  ชั้นนี้ไม่มีรายชื่อในตาราง")
        if S.stop:
            break
    console(True)
    report(path, kids, None, S)


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
    ap.add_argument("--redo", action="store_true", help="บันทึกซ้ำคนที่ลงไปแล้ว (เช่น แก้ผู้ให้บริการ)")
    ap.add_argument("--classes", type=int, default=0, help="โหมด all: จำนวนชั้นที่จะทำ (ไม่ใส่ = เท่าจำนวนห้องในไฟล์)")
    # พิมพ์รวมในช่องห้องได้ เช่น "a2 + redo" หรือ "a2 save"
    args = [x for x in sys.argv[1:] if x not in ("+", ",")]
    for w, f in (("save", ["--save", "--auto"]), ("redo", ["--save", "--auto", "--redo"])):
        if w in [x.lower() for x in args[1:]]:
            args = [x for x in args if x.lower() != w] + f
    sys.argv = sys.argv[:1] + list(dict.fromkeys(args))
    try:
        if len(sys.argv) > 1 and sys.argv[1] == "calib":
            calib_scroll()
        elif len(sys.argv) > 1 and sys.argv[1] in ("calib-provider", "provider"):
            calib_provider()
        elif len(sys.argv) > 1 and sys.argv[1] in ("school", "calib-school"):
            calib_school()
        elif len(sys.argv) > 1 and sys.argv[1] in ("tprov", "table-provider"):
            calib_tprov()
        else:
            a = ap.parse_args()
            path = a.csv or newest_csv()
            print("ไฟล์:", os.path.basename(path))
            P0 = json.load(open(jf.CAL, encoding="utf-8"))
            if "provider" not in P0:
                print("** ยังไม่ได้จดผู้ให้บริการ: ช่องผู้ให้บริการจะเป็นชื่อตั้งต้นของ JHCIS (จดด้วยคำสั่ง provider)")
            if "tprov" in P0:
                from PIL import Image
                bad = "tprov_item_dx" not in P0 or not all(map(os.path.exists, (TITEM, TREF_OK, TREF_DEF)))
                if not bad:
                    bad = Image.open(TREF_OK).size != (34, 14) or Image.open(TITEM).size[0] != TIW
                if bad:
                    sys.exit("ต้องจดผู้ให้บริการในตารางใหม่ก่อน (พิมพ์ห้องเป็น tprov) ครั้งเดียว")
            if a.room.lower() == "all":
                run_all(path, a.classes, a.save, a.auto, a.visible, a.delay, a.redo)
            else:
                run(path, room_name(a.room), a.save, a.auto, a.visible, a.delay, a.redo)
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
