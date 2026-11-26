# -*- coding: utf-8 -*-
"""
JHCIS อนามัยโรงเรียน (โปรแกรมสำเร็จรูป)  รุ่น 1.2
หน้าต่างเดียว กดทีละขั้น: เลือกรายชื่อ -> เทียบกับ JHCIS -> ลงรายชื่อ -> ลงผลตรวจฟัน -> ตรวจความพร้อมรายงาน
งานที่คลิกแทนคนจะเปิดหน้าต่างดำแยก (ใช้สคริปต์ชุดเดิมที่ทดสอบแล้ว) ระหว่างทำงานห้ามจับเมาส์
หยุดฉุกเฉิน: ลากเมาส์ไปมุมซ้ายบนสุดของจอ
"""
import os
import sys
import json
import glob
import threading
import subprocess

APP_VER = "1.2"
TOOLS = ["jhcis_roster_compare", "jhcis_student_fill", "jhcis_newperson_fill", "jhcis_dental_fill", "jhcis_dental_room",
         "jhcis_link_existing", "docx_roster_to_xlsx", "jhcis_birth_fill", "jhcis_check_report",
         "jhcis_check_detail", "jhcis_codes", "jhcis_compare_person"]


# ---------------- โหมดรันเครื่องมือในหน้าต่างดำ ----------------
def run_tool(name, args):
    import importlib
    if name not in TOOLS:
        print("ไม่รู้จักเครื่องมือ", name); return
    sys.argv = [name + ".py"] + args
    try:
        importlib.import_module(name).cli()
    except SystemExit as e:
        if e.code not in (None, 0):
            print("\n" + str(e.code))
    except Exception:
        import traceback
        traceback.print_exc()
        print("\n*** เกิดข้อผิดพลาด ถ่ายรูปหน้าต่างนี้ส่งให้ผู้ดูแล")
    finally:
        try:
            input("\n===== เสร็จแล้ว กด Enter เพื่อปิดหน้าต่างนี้ =====")
        except EOFError:
            pass


# ---------------- หน้าต่างหลัก ----------------
def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import jhcis_roster_compare as rc

    if getattr(sys, "frozen", False) and os.name == "nt":           # ซ่อนหน้าต่างดำของตัวโปรแกรมหลัก
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except Exception:
            pass

    HOME = rc.home()
    CFG = os.path.join(HOME, "app_settings.json")
    try:
        cfg = json.load(open(CFG, encoding="utf-8"))
    except Exception:
        cfg = {}

    def save_cfg():
        json.dump(cfg, open(CFG, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    srv = rc.jhcis_server() or "localhost"
    local = srv.lower() in ("localhost", "127.0.0.1")

    root = tk.Tk()
    root.title(f"JHCIS อนามัยโรงเรียน  รุ่น {APP_VER}")
    root.geometry("860x760")
    F = ("Tahoma", 11)
    FB = ("Tahoma", 11, "bold")
    root.option_add("*Font", F)

    busy = tk.StringVar(value="พร้อม")
    pw = tk.StringVar()
    roster = tk.StringVar(value=cfg.get("roster", ""))
    infer = tk.BooleanVar(value=cfg.get("infer", False))
    spos = tk.StringVar(value="")
    dfile = tk.StringVar(value=cfg.get("dental_csv", ""))
    dsheet = tk.StringVar()
    summary = tk.StringVar(value="(ยังไม่ได้เทียบ)")

    # ---- ตัวช่วย ----
    def cmd_for(name, args):
        if getattr(sys, "frozen", False):
            return [sys.executable, "--tool", name] + args
        return [sys.executable, os.path.abspath(__file__), "--tool", name] + args

    def need_pw():
        if not pw.get():
            messagebox.showwarning("รหัสผ่าน", "ใส่รหัสผ่าน MySQL (root) ด้านบนก่อน"); return False
        return True

    def run_chain(steps, done=None, title=""):
        """steps = [(tool, [args]), ...] รันต่อกันทีละหน้าต่าง"""
        def work():
            busy.set(f"กำลังทำ: {title} ... (ดูหน้าต่างดำ ห้ามจับเมาส์ระหว่างคลิกอัตโนมัติ)")
            env = dict(os.environ, JHCIS_PW=pw.get(), PYTHONIOENCODING="utf-8")
            for name, args in steps:
                kw = {"creationflags": subprocess.CREATE_NEW_CONSOLE} if os.name == "nt" else {}
                p = subprocess.Popen(cmd_for(name, args), env=env, cwd=HOME, **kw)
                p.wait()
            busy.set("พร้อม")
            if done:
                root.after(0, done)
        threading.Thread(target=work, daemon=True).start()

    def real_flags():
        return [] if local else ["--real"]

    def confirm_real(what):
        if local:
            return True
        return messagebox.askyesno("ฐานข้อมูลจริง",
                                   f"{what}\nจะบันทึกลง *ฐานข้อมูลจริง* ({srv})\n\n"
                                   "• สำรองฐานข้อมูลแล้วหรือยัง?\n• ไม่มีใครใช้ JHCIS อยู่ (นอกเวลาทำการ)?\n\nทำต่อ?")

    def paths():
        r = roster.get()
        if not r:
            return None, None, None
        base = os.path.splitext(r)[0]
        xlsx = base + ".xlsx" if r.lower().endswith(".docx") else r
        d, f = os.path.split(xlsx)
        res = os.path.join(d, "ผลเทียบ_" + os.path.splitext(f)[0] + ".xlsx")
        link = os.path.join(d, "จับคู่_" + os.path.splitext(f)[0] + ".xlsx")
        return xlsx, res, link

    def points_ok(fn, keys=()):
        try:
            P = json.load(open(os.path.join(HOME, fn), encoding="utf-8"))
            return all(k in P for k in keys)
        except Exception:
            return False

    # ---- ส่วนบน: เครื่องนี้ ----
    top = ttk.LabelFrame(root, text=" เครื่องนี้ ")
    top.pack(fill="x", padx=10, pady=6)
    ttk.Label(top, text=f"ชื่อเครื่อง: {os.environ.get('COMPUTERNAME', '-')}").grid(row=0, column=0, sticky="w", padx=6)
    tk.Label(top, text=f"JHCIS ต่อกับ: {srv}  ->  " + ("ฐานทดสอบในเครื่อง" if local else "ฐานข้อมูลจริง"),
             fg=("#1a7f37" if local else "#c62828"), font=FB).grid(row=0, column=1, sticky="w", padx=6)
    ttk.Label(top, text="รหัสผ่าน MySQL (root):").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ttk.Entry(top, textvariable=pw, show="*", width=18).grid(row=1, column=1, sticky="w", padx=6)

    def test_conn():
        if not need_pw():
            return
        rc.HOST = "127.0.0.1" if local else srv
        try:
            h = rc.q("SELECT @@hostname AS h", pw.get())[0]["h"]
            n = rc.q("SELECT COUNT(*) AS n FROM person", pw.get())[0]["n"]
            messagebox.showinfo("ต่อได้", f"ต่อฐานข้อมูลได้\nเครื่อง {h}\nมีประชากร {n} คน")
        except SystemExit as e:
            messagebox.showerror("ต่อไม่ได้", str(e))
    ttk.Button(top, text="ทดสอบการต่อ", command=test_conn).grid(row=1, column=1, sticky="e", padx=6)
    ttk.Button(top, text="เปิดโฟลเดอร์ข้อมูลเครื่องนี้", command=lambda: os.startfile(HOME)).grid(row=0, column=2, rowspan=2, padx=6)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=10, pady=4)

    # ---- แท็บ 0: ตั้งค่าเครื่อง ----
    t0 = ttk.Frame(nb); nb.add(t0, text=" 0 ตั้งค่าเครื่องนี้ (ครั้งแรก) ")
    ttk.Label(t0, text="จอแต่ละเครื่องไม่เท่ากัน ต้องจดตำแหน่งปุ่มของ JHCIS เครื่องละ 1 ครั้ง\n"
                       "เปิด JHCIS ค้างไว้ แล้วทำตามคำสั่งในหน้าต่างดำ  (ชี้เมาส์ ห้ามคลิก ห้ามกดบันทึกใน JHCIS)",
              foreground="#555").pack(anchor="w", padx=8, pady=6)
    st0 = tk.StringVar()

    def refresh0():
        a = points_ok("jhcis_student_points.json", ("id_field", "hn_field"))
        b = points_ok("jhcis_newperson_points.json", ("new", "cls_down", "del"))
        c = points_ok("jhcis_dental_points.json", ("row1_dental", "save"))
        st0.set(f"ใส่ชั้นเรียน: {'✓ จดแล้ว' if a else '✗ ยังไม่จด'}     "
                f"เพิ่มเด็กใหม่: {'✓ จดแล้ว' if b else '✗ ยังไม่จด'}     "
                f"ผลตรวจฟัน: {'✓ จดแล้ว' if c else '✗ ยังไม่จด'}")
    ttk.Label(t0, textvariable=st0, font=FB).pack(anchor="w", padx=8, pady=6)
    ttk.Button(t0, text="จดตำแหน่ง: ใส่ชั้นเรียน (9 จุด + ช่อง HN)",
               command=lambda: run_chain([("jhcis_student_fill", ["calibrate"]), ("jhcis_student_fill", ["calibrate-hn"])],
                                         refresh0, "จดตำแหน่งใส่ชั้น")).pack(fill="x", padx=8, pady=3)
    ttk.Button(t0, text="จดตำแหน่ง: เพิ่มเด็กใหม่ (รวมปุ่มลบ และช่อง HN)",
               command=lambda: run_chain([("jhcis_newperson_fill", ["calibrate"]), ("jhcis_newperson_fill", ["calibrate-del"]),
                                          ("jhcis_newperson_fill", ["calibrate-hn"])],
                                         refresh0, "จดตำแหน่งเพิ่มเด็กใหม่")).pack(fill="x", padx=8, pady=3)
    ttk.Button(t0, text="จดตำแหน่ง: ฟอร์มผลตรวจฟัน",
               command=lambda: run_chain([("jhcis_dental_fill", ["calibrate"])], refresh0, "จดตำแหน่งผลตรวจฟัน")).pack(fill="x", padx=8, pady=3)
    ttk.Button(t0, text="จดตำแหน่ง: ปุ่มลูกศร ▲▼ แถบเลื่อนตาราง 'ตรวจทันตกรรม'",
               command=lambda: run_chain([("jhcis_dental_room", ["calib"])], refresh0, "จดตำแหน่งแถบเลื่อน")).pack(fill="x", padx=8, pady=3)
    ttk.Button(t0, text="จดผู้ให้บริการ (เปิดฟอร์ม Dental ค้างไว้ แล้วชี้ช่องผู้ให้บริการ กับชื่อในรายการ)",
               command=lambda: run_chain([("jhcis_dental_room", ["provider"])], refresh0, "จดผู้ให้บริการ")).pack(fill="x", padx=8, pady=3)
    ttk.Button(t0, text="เช็กตำแหน่งฟอร์มผลตรวจฟัน (เมาส์ชี้ทีละช่อง ไม่คลิก)",
               command=lambda: run_chain([("jhcis_dental_fill", ["check"])], None, "เช็กตำแหน่ง")).pack(fill="x", padx=8, pady=3)
    refresh0()

    # ---- แท็บ 1: รายชื่อ + เทียบ ----
    t1 = ttk.Frame(nb); nb.add(t1, text=" 1 รายชื่อ & เทียบ ")
    ttk.Label(t1, text="ไฟล์รายชื่อนักเรียนจากโรงเรียน (Excel หรือ Word) 1 ไฟล์ = 1 โรงเรียน").pack(anchor="w", padx=8, pady=(8, 2))
    fr = ttk.Frame(t1); fr.pack(fill="x", padx=8)
    ttk.Entry(fr, textvariable=roster).pack(side="left", fill="x", expand=True)

    def pick_roster():
        f = filedialog.askopenfilename(filetypes=[("รายชื่อ", "*.xlsx *.docx")])
        if f:
            roster.set(os.path.normpath(f)); cfg["roster"] = roster.get(); save_cfg(); load_summary()
    ttk.Button(fr, text="เลือกไฟล์...", command=pick_roster).pack(side="left", padx=4)
    ttk.Checkbutton(t1, text="ไฟล์ไม่มีคอลัมน์ห้อง แต่เรียงทีละห้อง (ชายก่อนแล้วหญิง) ให้เดาห้องเอง",
                    variable=infer).pack(anchor="w", padx=8, pady=4)

    def do_compare():
        if not roster.get() or not need_pw():
            return
        cfg["infer"] = infer.get(); save_cfg()
        xlsx, res, link = paths()
        steps = []
        if roster.get().lower().endswith(".docx"):
            steps.append(("docx_roster_to_xlsx", [roster.get()]))
        steps.append(("jhcis_roster_compare", [xlsx] + (["--infer-room"] if infer.get() else [])))
        steps.append(("jhcis_link_existing", [res]))
        run_chain(steps, load_summary, "เทียบรายชื่อ")
    ttk.Button(t1, text="▶ เทียบกับ JHCIS (อ่านอย่างเดียว ไม่แก้อะไร)", command=do_compare).pack(fill="x", padx=8, pady=6)
    ttk.Label(t1, textvariable=summary, justify="left", font=("Tahoma", 11)).pack(anchor="w", padx=12, pady=6)

    def load_summary():
        xlsx, res, link = paths()
        if not res or not os.path.exists(res):
            summary.set("(ยังไม่ได้เทียบ)"); return
        try:
            from openpyxl import load_workbook
            ws = load_workbook(res, read_only=True)["สรุป"]
            lines = [f"{r[0]}: {r[1]}" for r in ws.iter_rows(values_only=True) if r and r[0] is not None]
            if os.path.exists(link):
                wl = load_workbook(link, read_only=True)["สรุป"]
                lines += ["", "เด็กที่มีอยู่แล้วแต่เลขบัตรไม่ตรง:"] + [f"   {r[0]}: {r[1]}" for r in wl.iter_rows(min_row=2, values_only=True) if r[0]]
            summary.set("\n".join(lines))
            guess_school()
        except Exception as e:
            summary.set(f"อ่านผลเทียบไม่ได้: {e}")

    # ---- แท็บ 2: ลงรายชื่อ ----
    t2 = ttk.Frame(nb); nb.add(t2, text=" 2 ลงรายชื่อเข้า JHCIS ")
    ttk.Label(t2, text="ลำดับของโรงเรียนนี้ในรายการ 'โรงเรียน' ของ JHCIS (นับแถวว่างบนสุดเป็นแถวที่ 1):").pack(anchor="w", padx=8, pady=(8, 2))
    f2 = ttk.Frame(t2); f2.pack(fill="x", padx=8)
    ttk.Spinbox(f2, from_=2, to=40, textvariable=spos, width=5).pack(side="left")
    sinfo = tk.StringVar(value="")
    ttk.Label(f2, textvariable=sinfo, foreground="#555").pack(side="left", padx=8)

    def school_name():
        _, res, _ = paths()
        try:
            from openpyxl import load_workbook
            return load_workbook(res, read_only=True)["สรุป"]["B1"].value
        except Exception:
            return None

    def guess_school():
        name = school_name()
        if not name:
            return
        saved = cfg.get("school_pos", {}).get(name)
        if saved:
            spos.set(str(saved)); sinfo.set(f"{name}: ใช้ลำดับที่เคยบันทึก"); return
        if pw.get():
            try:
                rc.HOST = "127.0.0.1" if local else srv
                names = [r["schoolname"] for r in rc.q("SELECT schoolname FROM villageschool", pw.get())]
                if name in names:
                    spos.set(str(names.index(name) + 2))
                sinfo.set(f"{name}: เดาจากฐานข้อมูล ตรวจกับรายการใน JHCIS ก่อน ({' / '.join(names)})")
                return
            except SystemExit:
                pass
        sinfo.set(f"{name}: ใส่ลำดับเอง")

    def spos_ok():
        if not spos.get().isdigit():
            messagebox.showwarning("โรงเรียน", "ใส่ลำดับโรงเรียนก่อน"); return False
        name = school_name()
        if name:
            cfg.setdefault("school_pos", {})[name] = int(spos.get()); save_cfg()
        return True

    def stu(limit, auto):
        if not (need_pw() and spos_ok()):
            return
        xlsx, res, link = paths()
        if not os.path.exists(res):
            messagebox.showwarning("ยังไม่เทียบ", "ทำแท็บ 1 (เทียบ) ก่อน"); return
        if not confirm_real("ใส่ชั้นเรียนให้นักเรียนที่มีใน JHCIS แล้ว"):
            return
        base = ["run", None, "--school-pos", spos.get(), "--save", "--wait", "1.5" if not local else "1.0"] + real_flags()
        base += (["--limit", "3"] if limit else []) + (["--auto"] if auto else [])
        steps = [("jhcis_student_fill", [res if i is None else i for i in base])]
        try:
            from openpyxl import load_workbook
            has_link = os.path.exists(link) and load_workbook(link, read_only=True)["3 เพิ่มเข้าโรงเรียน"].max_row > 1
        except Exception:
            has_link = False
        if has_link:                                           # เด็กที่มีอยู่แล้วแต่เลขบัตรไม่ตรง -> ใส่ชั้นให้คนเดิม
            steps.append(("jhcis_student_fill", [link if i is None else i for i in base]))
        run_chain(steps, None, "ใส่ชั้นเรียน")

    def newp(limit, auto):
        if not (need_pw() and spos_ok()):
            return
        xlsx, res, link = paths()
        if not os.path.exists(res):
            messagebox.showwarning("ยังไม่เทียบ", "ทำแท็บ 1 (เทียบ) ก่อน"); return
        if not confirm_real("เพิ่มเด็กที่ยังไม่มีใน JHCIS (สร้าง HN ใหม่)"):
            return
        if auto and not messagebox.askyesno("ยืนยัน", "โหมดอัตโนมัติ: ห้ามใครเพิ่มคนไข้ใหม่ใน JHCIS ระหว่างนี้\nทำต่อ?"):
            return
        args = ["run", res, "--school-pos", spos.get(), "--save", "--wait", "2.0" if not local else "1.5"] + real_flags()
        args += (["--limit", "3"] if limit else []) + (["--auto"] if auto else [])
        run_chain([("jhcis_newperson_fill", args)], None, "เพิ่มเด็กใหม่")

    for txt, fn in (("ใส่ชั้นเรียน: ลอง 3 คน (ถามทีละคน)", lambda: stu(True, False)),
                    ("ใส่ชั้นเรียน: ทั้งหมด (อัตโนมัติ)", lambda: stu(False, True)),
                    ("เพิ่มเด็กใหม่: ลอง 3 คน (ถามทีละคน)", lambda: newp(True, False)),
                    ("เพิ่มเด็กใหม่: ทั้งหมด (อัตโนมัติ / นอกเวลาทำการ)", lambda: newp(False, True)),
                    ("หาคนที่สร้างค้างไว้ (อ่านอย่างเดียว)", lambda: need_pw() and run_chain([("jhcis_newperson_fill", ["leftovers"])], None, "หาคนค้าง"))):
        ttk.Button(t2, text=txt, command=fn).pack(fill="x", padx=8, pady=3)
    ttk.Label(t2, text="ลำดับที่แนะนำ: ใส่ชั้น 3 คน -> ใส่ชั้นทั้งหมด -> เพิ่มเด็กใหม่ 3 คน -> เพิ่มเด็กใหม่ทั้งหมด -> กลับไปแท็บ 1 เทียบซ้ำ\n"
                       "ถ้าหยุดกลางทาง: อ่านข้อความในหน้าต่างดำ แก้ตามนั้น แล้วกดปุ่มเดิมซ้ำ (คนที่ทำแล้วจะข้ามเอง)",
              foreground="#555", justify="left").pack(anchor="w", padx=8, pady=8)

    # ---- แท็บ 3: ผลตรวจฟัน (ทั้งห้อง จับคู่ด้วยชื่อ) ----
    t3 = ttk.Frame(nb); nb.add(t3, text=" 3 ลงผลตรวจฟัน ")
    ttk.Label(t3, text="ไฟล์ ลงJHCIS_โรงเรียน.csv  (แอปกรอกผลตรวจฟัน > ⋯ เพิ่มเติม > ไฟล์สำหรับลง JHCIS)").pack(anchor="w", padx=8, pady=(8, 2))
    f3 = ttk.Frame(t3); f3.pack(fill="x", padx=8)
    ttk.Entry(f3, textvariable=dfile).pack(side="left", fill="x", expand=True)
    cb = ttk.Combobox(t3, textvariable=dsheet, state="readonly", width=12)
    dinfo = tk.StringVar()
    redo = tk.BooleanVar(value=False)

    def load_sheets():
        try:
            import csv as _csv
            rows = list(_csv.DictReader(open(dfile.get(), encoding="utf-8-sig")))
            import re as _re
            def k(r):
                m = _re.match(r"([อปม])\D*(\d+)", r)
                return ("อปม".index(m.group(1)), int(m.group(2))) if m else (9, 0)
            rooms = sorted({r["ROOM"] for r in rows}, key=k)
            cb["values"] = rooms
            if rooms and dsheet.get() not in rooms:
                dsheet.set(rooms[0])
            n = {x: sum(1 for r in rows if r["ROOM"] == x) for x in rooms}
            dinfo.set("ในไฟล์: " + "  ".join(f"{x} {n[x]} คน" for x in rooms))
        except Exception as e:
            cb["values"] = []; dinfo.set("อ่านไฟล์ไม่ได้: " + str(e))

    def pick_dental():
        f = filedialog.askopenfilename(filetypes=[("ไฟล์ลง JHCIS", "*.csv")])
        if f:
            dfile.set(os.path.normpath(f)); cfg["dental_csv"] = dfile.get(); save_cfg(); load_sheets()
    ttk.Button(f3, text="เลือกไฟล์...", command=pick_dental).pack(side="left", padx=4)
    ttk.Label(t3, textvariable=dinfo, foreground="#555").pack(anchor="w", padx=8, pady=(2, 0))
    f3b = ttk.Frame(t3); f3b.pack(fill="x", padx=8, pady=6)
    ttk.Label(f3b, text="ห้อง:").pack(side="left")
    cb.pack(in_=f3b, side="left", padx=4)
    ttk.Checkbutton(t3, text="บันทึกซ้ำคนที่ลงไปแล้วด้วย (ใช้ตอนแก้ผลหรือแก้ผู้ให้บริการ)", variable=redo).pack(anchor="w", padx=8)
    ttk.Label(t3, text="ก่อนกด: ใน JHCIS เปิด 'บันทึกข้อมูลงานอนามัยโรงเรียน' เลือกโรงเรียน/ชั้นให้ตรงกับห้อง แล้วกด 'ตรวจทันตกรรม' ค้างไว้\n"
                       "โปรแกรมจะเปิดฟอร์ม Dental ทีละคน อ่านชื่อจากฟอร์ม แล้วกรอกผลของเด็กคนนั้น (ลำดับไม่ต้องตรงกัน)\n"
                       "เลื่อนตารางลงเองจนครบ ข้ามคนไม่มา/ไม่มีในไฟล์/ที่บันทึกไปแล้ว  แป้นพิมพ์ต้องเป็น ENG",
              foreground="#555", justify="left").pack(anchor="w", padx=8)

    def dental(save):
        if not dfile.get() or not dsheet.get():
            messagebox.showwarning("ผลตรวจฟัน", "เลือกไฟล์และห้องก่อน"); return
        if save and not confirm_real("บันทึกผลตรวจฟัน ห้อง " + dsheet.get()):
            return
        run_chain([("jhcis_dental_room", [dsheet.get(), "--csv", dfile.get()] + (["--save", "--auto"] if save else []) + (["--redo"] if redo.get() else []))],
                  None, "ลงผลตรวจฟัน " + dsheet.get())
    ttk.Button(t3, text="ซ้อม (กรอกแต่ไม่บันทึก ถามทุกคน)", command=lambda: dental(False)).pack(fill="x", padx=8, pady=3)
    ttk.Button(t3, text="บันทึกจริงทั้งห้อง (ถาม 3 คนแรก แล้วทำต่อเอง)", command=lambda: dental(True)).pack(fill="x", padx=8, pady=3)
    ttk.Button(t3, text="เปิดโฟลเดอร์สรุปผล (ลงJHCIS_ผล_ห้อง_เวลา.txt)", command=lambda: os.startfile(HOME)).pack(fill="x", padx=8, pady=3)
    if dfile.get():
        load_sheets()

    # ---- แท็บ 4: ตรวจรายงาน ----
    t4 = ttk.Frame(nb); nb.add(t4, text=" 4 ตรวจความพร้อมรายงาน ")
    ttk.Label(t4, text="ตรวจว่านักเรียนปีนี้มีข้อมูลครบพอขึ้น 43 แฟ้ม / HDC (อ่านอย่างเดียว)", foreground="#555").pack(anchor="w", padx=8, pady=8)
    ttk.Button(t4, text="ตรวจภาพรวมทุกโรงเรียน",
               command=lambda: need_pw() and run_chain([("jhcis_check_report", [])], None, "ตรวจภาพรวม")).pack(fill="x", padx=8, pady=3)

    def detail():
        if not need_pw():
            return
        xs = [p for p in cfg.get("rosters_all", []) if os.path.exists(p)]
        x, _, _ = paths()
        if x and os.path.exists(x) and x not in xs:
            xs.append(x); cfg["rosters_all"] = xs; save_cfg()
        run_chain([("jhcis_check_detail", xs)], None, "ทำรายชื่อที่ต้องแก้")
    ttk.Button(t4, text="ทำรายชื่อเด็กที่ต้องแก้ (ไม่มีวันเกิด / สถานะจำหน่าย)", command=detail).pack(fill="x", padx=8, pady=3)

    def birth():
        fs = sorted(glob.glob(os.path.join(HOME, "ต้องแก้ในJHCIS_*.xlsx")))
        if not fs:
            messagebox.showwarning("ยังไม่มีรายชื่อ", "กดปุ่ม 'ทำรายชื่อเด็กที่ต้องแก้' ก่อน"); return
        if not (need_pw() and confirm_real("ใส่วันเกิดที่ขาด (จากไฟล์รายชื่อ)")):
            return
        run_chain([("jhcis_birth_fill", [fs[-1], "--save", "--auto"] + real_flags())], None, "ใส่วันเกิด")
    ttk.Button(t4, text="ใส่วันเกิดที่ขาด (อัตโนมัติ จากรายชื่อล่าสุด)", command=birth).pack(fill="x", padx=8, pady=3)

    ttk.Label(root, textvariable=busy, foreground="#1565c0").pack(anchor="w", padx=12, pady=4)
    load_summary()
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--tool":
        run_tool(sys.argv[2], sys.argv[3:])
    else:
        gui()
