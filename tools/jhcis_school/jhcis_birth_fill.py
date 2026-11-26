"""
jhcis_birth_fill.py  (รุ่น v0.2)
ใส่วันเกิดให้นักเรียนที่ใน JHCIS ยังไม่มีวันเกิด (จากไฟล์ ต้องแก้ในJHCIS_....xlsx ที่ jhcis_check_detail.py ทำให้)
ใช้ตำแหน่งที่จดไว้แล้ว: jhcis_student_points.json (ช่อง HN, แถวแรก) + jhcis_newperson_points.json (แท็บข้อมูลบุคคล, ช่องวันเกิด, แผ่นดิสก์)

ด่านกันพลาด (หยุดทันที ไม่บันทึก):
  - ในฐานข้อมูลคนนี้ต้อง "ยังไม่มีวันเกิด" ก่อนเริ่ม (ถ้ามีแล้ว = ข้าม)
  - ค้น HN แล้วชื่อในช่อง 'ข้อมูลของ' ต้องเปลี่ยน
  - ช่องวันเกิดบนจอต้อง "ว่าง" ก่อนพิมพ์ (ถ้ามีค่า = เปิดผิดคน)
  - หลังบันทึก: วันเกิดในฐานต้องตรงไฟล์ และ ชื่อ/สกุล/เลขบัตร ต้องเหมือนเดิม

ใช้:  python jhcis_birth_fill.py "ต้องแก้ในJHCIS_....xlsx" --limit 3 --save --real
      python jhcis_birth_fill.py "ต้องแก้ในJHCIS_....xlsx" --save --auto --real
"""
import os, re, sys, json, time, argparse, getpass
import jhcis_roster_compare as rc

HERE = rc.home()
PROP = r"C:\Program Files\JHCIS\database.properties"
WIN_KEY = "JHCIS V"


def server():
    try:
        for ln in open(PROP, encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                return ln.split("=", 1)[1].strip()
    except OSError:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--save", action="store_true")
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--real", action="store_true")
    ap.add_argument("--wait", type=float, default=1.5)
    a = ap.parse_args()
    import pyautogui, pygetwindow
    pyautogui.FAILSAFE = True
    print("jhcis_birth_fill รุ่น v0.2")

    S = json.load(open(os.path.join(HERE, "jhcis_student_points.json"), encoding="utf-8"))
    N = json.load(open(os.path.join(HERE, "jhcis_newperson_points.json"), encoding="utf-8"))
    if "hn_field" not in S:
        sys.exit("ยังไม่ได้จดช่อง HN: python jhcis_student_fill.py calibrate-hn")
    P = {"hn": S["hn_field"], "first": S.get("first_row"), "id": S["id_field"],
         "tab_person": N["tab_person"], "birth": N["birth"], "disk": N["disk"]}

    srv = server() or "localhost"
    local = srv.lower() in ("localhost", "127.0.0.1")
    rc.HOST = "127.0.0.1" if local else srv
    print("=" * 50 + f"\nJHCIS ต่อกับ: {srv} -> " + ("ฐานทดสอบในเครื่อง" if local else "*** ฐานข้อมูลจริง ***") + "\n" + "=" * 50)
    if not local and not a.real:
        sys.exit("ต่อฐานจริงอยู่ ต้องใส่ --real")
    pw = rc.ask_pw(f"รหัสผ่าน root ของ MySQL ที่ {rc.HOST}: ")

    from openpyxl import load_workbook
    ws = load_workbook(a.xlsx, data_only=True).active
    h = [c.value for c in ws[1]]
    todo = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(h, r))
        m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", str(d.get("แก้เป็น (จากไฟล์)") or ""))
        if d.get("ปัญหา") == "ไม่มีวันเกิด" and m:
            dd, mm, bb = m.groups()
            todo.append({"pid": str(d["HN"]), "name": d["ชื่อ"], "be": f"{dd}/{mm}/{bb}",
                         "iso": f"{int(bb) - 543}-{mm}-{dd}"})
    print(f"ต้องใส่วันเกิด {len(todo)} คน | โหมด: " + ("บันทึกจริง" if a.save else "ซ้อม (ไม่บันทึก)"))

    def at(k):
        ws_ = [w for w in pygetwindow.getAllWindows() if (w.title or "").startswith(WIN_KEY)]
        if not ws_:
            raise RuntimeError("ไม่พบหน้าต่าง JHCIS")
        return (ws_[0].left + P[k][0], ws_[0].top + P[k][1])

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

    def popups():
        """ปิดเตือนที่รู้จัก คืนชื่อหน้าต่างแปลก ๆ ถ้ามี"""
        for _ in range(4):
            w = pygetwindow.getActiveWindow(); t = (w.title or "") if w else ""
            if t == "Message" or t.startswith(("สิทธิ", "Invalid CID", "!This is THAI")):
                pyautogui.press("enter"); time.sleep(0.8); continue
            if t.startswith("Confirm (Party"):                 # CANCEL [เข้าใจแล้ว]
                pyautogui.click(int(w.left + w.width * 0.66), int(w.top + w.height * 0.89)); time.sleep(0.8); continue
            return None if (not t or t.startswith(WIN_KEY) or "PowerShell" in t) else t
        return "หน้าต่างเตือนขึ้นซ้ำ"

    def owner():
        x, y = at("id")
        return pyautogui.screenshot(region=(int(x) + 15, int(y) + 26, 230, 18)).convert("L")

    def diff(a_, b_, need=40):
        return sum(1 for p, q in zip(a_.tobytes(), b_.tobytes()) if abs(p - q) > 60) >= need

    def field_empty(k):
        x, y = at(k)
        im = pyautogui.screenshot(region=(int(x) + 6, int(y) - 7, 80, 14)).convert("RGB").tobytes()
        n = sum(1 for i in range(0, len(im), 3)
                if 0.299 * im[i] + 0.587 * im[i + 1] + 0.114 * im[i + 2] < 110
                and not (im[i] > im[i + 1] + 30 and im[i] > im[i + 2] + 30))
        return n < 25

    def person(pid):
        r = rc.q(f"SELECT fname,lname,idcard,birth FROM person WHERE pid={int(pid)}", pw)
        return r[0] if r else None

    input("แป้นพิมพ์ ENG / เปิด JHCIS หน้าข้อมูลบุคคลไว้ แล้วกด Enter (มีเวลา 3 วินาที)")
    time.sleep(3)
    log, done = [], 0
    for t in todo:
        if a.limit and done >= a.limit:
            break
        b0 = person(t["pid"])
        if not b0:
            log.append((t, "ข้าม-ไม่พบ HN")); continue
        if b0["birth"]:
            log.append((t, "ข้าม-มีวันเกิดแล้ว")); continue
        done += 1
        activate()
        before = owner()
        ok = False
        for att in range(3):
            activate()
            x, y = at("hn")
            pyautogui.click(x, y); time.sleep(0.3)
            pyautogui.press("end"); pyautogui.press("backspace", presses=12, interval=0.01)
            pyautogui.write(t["pid"], interval=0.03); pyautogui.press("enter")
            time.sleep(a.wait + 0.5 + att * 1.5)
            bad = popups()
            if bad:
                break
            if P["first"] and S.get("hn_list"):
                pyautogui.doubleClick(*at("first")); time.sleep(a.wait + att * 1.5)
            bad = popups()
            if bad:
                break
            pyautogui.click(*at("hn")); time.sleep(0.2)
            pyautogui.press("end"); pyautogui.press("backspace", presses=12, interval=0.01)
            if diff(before, owner()):
                ok = True; break
        if bad:
            print(f"\n*** หยุด: มีหน้าต่าง '{bad}'"); log.append((t, f"หยุด-หน้าต่าง {bad}")); break
        if not ok:
            print(f"  ✗ ข้าม {t['name']}: ค้น HN {t['pid']} ไม่ติด"); log.append((t, "ข้าม-ค้นไม่ติด")); continue
        pyautogui.click(*at("tab_person")); time.sleep(0.7)
        if popups():
            print("\n*** หยุด: มีหน้าต่างหลังเปิดแท็บข้อมูลบุคคล"); log.append((t, "หยุด-หน้าต่าง")); break
        if not field_empty("birth"):
            print(f"\n*** หยุด: ช่องวันเกิดบนจอมีค่าอยู่ แต่ในฐาน HN {t['pid']} ไม่มีวันเกิด = เปิดผิดคน ไม่บันทึก")
            log.append((t, "หยุด-ช่องวันเกิดบนจอไม่ว่าง")); break
        x, y = at("birth")
        shot = lambda: pyautogui.screenshot(region=(int(x) + 6, int(y) - 7, 80, 14)).convert("L")
        base = shot(); typed = False; bad = None
        for att in range(3):                                   # พิมพ์ไม่ติด -> ลองใหม่ ช้าลง
            pyautogui.click(x, y); time.sleep(0.4 + att * 0.4); pyautogui.click(x, y); time.sleep(0.3)
            pyautogui.hotkey("ctrl", "a"); pyautogui.write(t["be"], interval=0.05 + att * 0.03)
            pyautogui.press("tab"); time.sleep(0.8)
            bad = popups()
            if bad:
                break
            if not field_empty("birth") or diff(base, shot(), 8):
                typed = True; break
        if bad:
            print("\n*** หยุด: มีหน้าต่างหลังพิมพ์วันเกิด ไม่บันทึก"); log.append((t, "หยุด-หน้าต่างหลังพิมพ์")); break
        if not typed:
            print(f"  ✗ ข้าม {t['name']}: พิมพ์วันเกิดไม่ติด (ไม่บันทึก)"); log.append((t, "ข้าม-พิมพ์ไม่ติด")); continue
        if not a.save:
            log.append((t, "ซ้อม-ไม่บันทึก")); print(f"  (ซ้อม) {t['name']} -> {t['be']}"); continue
        if not a.auto:
            ans = pyautogui.confirm(text=f"HN {t['pid']} {t['name']}\nวันเกิด {t['be']}\n\nดูใน JHCIS ว่าเป็นคนนี้ แล้วกด บันทึก",
                                    title="ใส่วันเกิด JHCIS", buttons=["บันทึก", "ข้ามคนนี้", "หยุด"])
            if ans != "บันทึก":
                pyautogui.press("esc")
                log.append((t, f"ไม่บันทึก ({ans})"))
                if ans in (None, "หยุด"):
                    break
                continue
            activate()
        pyautogui.click(*at("disk")); time.sleep(a.wait + 1)
        popups()
        b1 = person(t["pid"])
        same = b1 and all(str(b1[k] or "") == str(b0[k] or "") for k in ("fname", "lname", "idcard"))
        if same and str(b1["birth"] or "")[:10] == t["iso"]:
            log.append((t, "บันทึกแล้ว-ตรวจฐานข้อมูลผ่าน")); print(f"  ✓ {t['name']} -> {t['be']}")
        elif same and not b1["birth"]:
            log.append((t, "ข้าม-บันทึกไม่ติด")); print(f"  ✗ ข้าม {t['name']}: บันทึกไม่ติด (ค่าในฐานยังว่าง)")
        else:
            print(f"\n*** หยุด: HN {t['pid']} ตรวจฐานไม่ผ่าน ได้ {b1} (ก่อน {b0}) ตรวจใน JHCIS ด้วยมือ")
            log.append((t, "หยุด-ตรวจฐานไม่ผ่าน")); break
    from collections import Counter
    with open(os.path.join(HERE, "jhcis_birth_log.txt"), "a", encoding="utf-8") as f:
        for t, s in log:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{srv}\t{t['pid']}\t{t['name']}\t{t['be']}\t{s}\n")
    left = [t for t, s in log if s.startswith("ข้าม-") and s not in ("ข้าม-มีวันเกิดแล้ว",)]
    if left:
        print("\nยังไม่ได้ใส่ (รันซ้ำ หรือทำมือ):")
        for t in left:
            print(f"  HN {t['pid']}  {t['name']}  วันเกิด {t['be']}")
    print("\nสรุป:", dict(Counter(s.split(":")[0] for _, s in log)))


def cli():
    main()


if __name__ == "__main__":
    cli()
