# -*- coding: utf-8 -*-
"""
jhcis_roster_compare.py  (v0.5)  เทียบรายชื่อนักเรียนจากโรงเรียน กับ JHCIS  ** อ่านอย่างเดียว **

ใช้:   python jhcis_roster_compare.py ไฟล์รายชื่อ.xlsx [--infer-room]
       --infer-room  ไฟล์ไม่มีคอลัมน์ห้อง แต่เรียงทีละห้อง (ชายก่อนแล้วหญิง): ขึ้นห้องใหม่เมื่อกลับมาเป็นชายอีกรอบในชั้นเดียวกัน
ผล:    ไฟล์ "ผลเทียบ_<ชื่อไฟล์>.xlsx" แบ่งเด็กเป็น 5 กลุ่ม
         1 ถูกแล้ว            อยู่โรงเรียนนี้ ชั้นตรง
         2 แก้ชั้น             อยู่โรงเรียนนี้ แต่ชั้นไม่ตรง
         3 เพิ่มเข้าโรงเรียน     มีในทะเบียนบุคคล แต่ยังไม่ผูกกับโรงเรียนนี้
         4 ไม่มีใน JHCIS       ไม่พบเลขประจำตัวในทะเบียนบุคคล
         5 ค้างในโรงเรียน       JHCIS ยังผูกกับโรงเรียนนี้ แต่ไม่มีในรายชื่อปีนี้ (จบ/ย้าย)

- ต่อเฉพาะฐานข้อมูลในเครื่องนี้ (127.0.0.1) และใช้แต่คำสั่ง SELECT ไม่แก้ข้อมูลใด ๆ
- ไฟล์ผลลัพธ์มีข้อมูลส่วนบุคคลของเด็ก เก็บไว้ในเครื่อง อย่าส่งต่อ
- อ่านไฟล์ที่มีหัวคอลัมน์ เลขประจำตัวนักเรียน หรือ เลขประจำตัวประชาชน / ชั้น / ชื่อ / นามสกุล / วันเกิด (มีคอลัมน์ ห้อง หรือชั้นแบบ ป.1/2 ก็ได้)
"""
import os, re, sys, subprocess, getpass, datetime

MYSQL = os.environ.get("JHCIS_MYSQL", r"C:\Program Files\JHCIS\MySQL\bin\mysql.exe")
HOST, PORT, DB = "127.0.0.1", "3333", "jhcisdb"
USER = "root"


# ---------- อ่านฐานข้อมูล (SELECT เท่านั้น) ----------
def home():
    """โฟลเดอร์เก็บตำแหน่ง/บันทึก: ตอนเป็นโปรแกรมสำเร็จรูป = ข้างไฟล์ .exe แยกตามชื่อเครื่อง (จอแต่ละเครื่องไม่เท่ากัน)"""
    if getattr(sys, "frozen", False):
        d = os.path.join(os.path.dirname(sys.executable), "ข้อมูลเครื่อง_" + os.environ.get("COMPUTERNAME", "pc"))
        os.makedirs(d, exist_ok=True)
        return d
    return os.path.dirname(os.path.abspath(__file__))


def ask_pw(prompt):
    """รหัสผ่านจากหน้าต่างโปรแกรม (ส่งมาทาง env) ถ้าไม่มีค่อยถาม"""
    return os.environ.get("JHCIS_PW") or getpass.getpass(prompt)


def jhcis_server():
    try:
        for ln in open(r"C:\Program Files\JHCIS\database.properties", encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                return ln.split("=", 1)[1].strip()
    except OSError:
        return None


try:
    import pymysql                                             # ไม่ต้องพึ่ง mysql.exe (เครื่องลูกข่ายบางเครื่องไม่มี)
except ImportError:
    pymysql = None
_CONN = {}


def _sv(v):
    import datetime as _d
    if v is None:
        return None
    if isinstance(v, bytes):
        return v.decode("utf-8", "replace")
    if isinstance(v, _d.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _d.date):
        return v.isoformat()
    return str(v)


def q(sql, pw):
    if not re.match(r"\s*(SELECT|SHOW)\b", sql, re.I) or ";" in sql:
        raise RuntimeError("อนุญาตเฉพาะคำสั่งอ่านข้อมูล")
    if pymysql is not None:
        key = (HOST, PORT, USER, pw)
        c = _CONN.get(key)
        if c is None:
            try:
                c = pymysql.connect(host=HOST, port=int(PORT), user=USER, password=pw, database=DB,
                                    charset="utf8", connect_timeout=10, autocommit=True)
            except Exception as e:
                sys.exit("ต่อฐานข้อมูลไม่ได้: " + str(e))
            _CONN[key] = c
        with c.cursor() as cur:
            cur.execute(sql)
            cols = [d[0] for d in cur.description] if cur.description else []
            return [dict(zip(cols, [_sv(v) for v in r])) for r in cur.fetchall()]
    env = dict(os.environ, MYSQL_PWD=pw)
    r = subprocess.run([MYSQL, "-h", HOST, "-P", PORT, "-u", USER, "--default-character-set=utf8",
                        "-B", DB, "-e", sql], capture_output=True, env=env)
    if r.returncode != 0:
        sys.exit("ต่อฐานข้อมูลไม่ได้: " + r.stderr.decode("utf-8", "replace").strip())
    lines = r.stdout.decode("utf-8", "replace").splitlines()
    if not lines:
        return []
    head = lines[0].split("\t")
    return [dict(zip(head, [None if v == "NULL" else v for v in ln.split("\t")])) for ln in lines[1:]]


# ---------- ตัวช่วย ----------
def norm_id(v):
    return re.sub(r"[\s\-]", "", str(v or "")).upper()


def norm_name(v):
    return re.sub(r"[\s\.\-]", "", str(v or ""))


def norm_class(v):
    """'อ.1' 'อนุบาล 1' 'ป.1' 'ประถมศึกษาปีที่ 1' 'ประถม1' -> ('อ',1) / ('ป',1) / ('ม',1)"""
    s = str(v or "").replace(" ", "")
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    n = int(m.group(1))
    if "อนุบาล" in s or s.startswith("อ"):
        return ("อ", n)
    if "ประถม" in s or s.startswith("ป"):
        return ("ป", n)
    if "มัธยม" in s or s.startswith("ม"):
        return ("ม", n)
    return None


def birth_roster(v):
    """'14/05/2566' (พ.ศ.) หรือ datetime -> 'YYYY-MM-DD' ค.ศ."""
    if isinstance(v, (datetime.date, datetime.datetime)):
        y = v.year - 543 if v.year > 2400 else v.year
        return f"{y:04d}-{v.month:02d}-{v.day:02d}"
    m = re.match(r"\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", str(v or ""))
    if not m:
        return ""
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y > 2400:
        y -= 543
    return f"{y:04d}-{mo:02d}-{d:02d}"


# ---------- อ่านไฟล์รายชื่อ ----------
ID_HEADS = ("เลขประจำตัวนักเรียน", "เลขประจำตัวประชาชน", "เลขบัตรประชาชน", "เลขประจำตัว 13 หลัก")


# สัญชาติ/เชื้อชาติ -> รหัส JHCIS (cnation)  ตามที่ รพ.สต. เลือก (แบบ C)
NATION_MAP = {"ไทย": "099", "เมียนมา": "048", "เมียนมา (กลุ่มเชื้อสายไทย)": "048", "ไทยใหญ่": "048", "พม่า": "048",
              "บุคคลที่ไม่มีสถานะทางทะเบียน": "999", "ไม่ได้สัญชาติไทย": "999", "ไม่ปรากฏเชื้อชาติ": "999",
              "ไม่ปรากฎเชื้อชาติ": "999", "อื่นๆ": "999", "อื่น ๆ": "999"}


def nation_code(v):
    v = re.sub(r"\s+", " ", str(v or "")).strip()
    return NATION_MAP.get(v, "") if v else ""


def read_roster(path, infer_room=False):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    best = None
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for hi, r in enumerate(rows[:15]):
            h = [str(c or "").strip() for c in r]
            idh = next((n for n in ID_HEADS if n in h), None)
            if idh and "ชื่อ" in h and "ชั้น" in h:
                data = [x for x in rows[hi + 1:] if x and len(x) > h.index(idh) and x[h.index(idh)]]
                if best is None or len(data) > len(best[2]):
                    best = (ws.title, h, data, rows[0][0] if rows and rows[0] else "")
                break
    if not best:
        sys.exit("ไม่พบแผ่นงานที่มีหัวคอลัมน์ เลขประจำตัวนักเรียน(หรือประชาชน) / ชั้น / ชื่อ")
    title, h, data, top = best
    def col(*names):
        for n in names:
            if n in h:
                return h.index(n)
        return None
    ci = {"id": col(*ID_HEADS), "cls": col("ชั้น"), "pre": col("คำนำหน้าชื่อ", "คำนำหน้า"), "fn": col("ชื่อ"),
          "ln": col("นามสกุล"), "bd": col("วันเกิด"), "nat": col("สัญชาติ"), "race": col("เชื้อชาติ"), "rel": col("ศาสนา"), "school": col("ชื่อโรงเรียน"), "room": col("ห้อง"), "sex": col("เพศ"),
          "hno": col("บ้านเลขที่", "เลขที่บ้าน (ทะเบียนบ้าน)"), "moo": col("หมู่", "หมู่ (ทะเบียนบ้าน)"),
          "tambon": col("ตำบล", "ตำบล (ทะเบียนบ้าน)"), "amphur": col("อำเภอ", "อำเภอ (ทะเบียนบ้าน)")}
    g = lambda r, k: (r[ci[k]] if ci[k] is not None and ci[k] < len(r) else None)
    out = []
    last_cls, last_male, room_no = None, None, 1
    for r in data:
        cls = str(g(r, "cls") or "").strip()
        room = str(g(r, "room") or "").strip()
        m = re.match(r"(.+?)\s*/\s*(\d+)$", cls)
        if m and not room:
            cls, room = m.group(1).strip(), m.group(2)
        if isinstance(g(r, "room"), float):
            room = str(int(g(r, "room")))
        male = str(g(r, "sex") or g(r, "pre") or "").strip().startswith(("ช", "เด็กชาย", "นาย", "ด.ช"))
        if infer_room and not room:
            if cls != last_cls:
                room_no = 1
            elif male and last_male is False:
                room_no += 1
            last_cls, last_male, room = cls, male, str(room_no)
        out.append({"id": norm_id(g(r, "id")), "cls": cls, "room": room, "pre": g(r, "pre") or "",
                    "fn": str(g(r, "fn") or "").strip(),
                    "nat": str(g(r, "nat") or "").strip(), "race": str(g(r, "race") or "").strip(), "rel": str(g(r, "rel") or "").strip(), "ln": str(g(r, "ln") or "").strip(),
                    "bd": birth_roster(g(r, "bd")), "school": str(g(r, "school") or "").strip(),
                    "addr": " ".join(f"{t}{g(r, k)}" for t, k in [("", "hno"), ("ม.", "moo"), ("ต.", "tambon"), ("อ.", "amphur")] if g(r, k))})
    print(f"อ่านรายชื่อจากแผ่นงาน '{title}' ได้ {len(out)} คน")
    import collections
    cnt = collections.Counter((x["cls"], x["room"] or "-") for x in out)
    print("  ชั้น/ห้อง: " + "  ".join(f"{c}/{r}={n}" for (c, r), n in cnt.items()))
    return out


# ---------- หลัก ----------
def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    path = sys.argv[1]
    roster = read_roster(path, "--infer-room" in sys.argv)
    dup = {r["id"] for r in roster if sum(1 for x in roster if x["id"] == r["id"]) > 1}
    if dup:
        print("คำเตือน: เลขประจำตัวซ้ำในไฟล์:", ", ".join(sorted(dup)))

    global HOST
    srv = None                                                 # ต่อฐานเดียวกับที่ JHCIS เครื่องนี้ต่ออยู่ (อ่านจาก database.properties)
    try:
        for ln in open(r"C:\Program Files\JHCIS\database.properties", encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                srv = ln.split("=", 1)[1].strip()
    except OSError:
        pass
    if srv and srv.lower() not in ("localhost", "127.0.0.1"):
        HOST = srv
    print("=" * 60 + f"\nเทียบกับฐาน: {HOST} -> " + ("ฐานทดสอบในเครื่อง" if HOST == "127.0.0.1" else "*** ฐานข้อมูลจริง ***") + "\n" + "=" * 60)
    pw = ask_pw(f"รหัสผ่าน root ของ MySQL ที่ {HOST}: ")
    info = q("SELECT @@hostname AS h", pw)[0]["h"]
    print(f"ต่อกับฐานข้อมูลของเครื่อง: {info}  ({HOST}:{PORT})  -- อ่านอย่างเดียว")

    pcols = {r["Field"] for r in q("SHOW COLUMNS FROM person", pw)}
    need = {"pcucodeperson", "pid", "idcard", "fname", "lname", "birth"}
    if not need <= pcols:
        sys.exit("ตาราง person ไม่มีคอลัมน์ที่คาดไว้: " + ", ".join(sorted(need - pcols)) + "  ส่งข้อความนี้ให้ Claude")
    extra = [c for c in ("typelive", "dischargetype") if c in pcols]
    persons = q("SELECT pcucodeperson,pid,idcard,fname,lname,birth" + "".join("," + c for c in extra) + " FROM person", pw)
    studs = q("SELECT pcucodeperson,pid,villcode,schoolno,classeducate,classroom,yeareducate FROM personstudent", pw)
    schools = q("SELECT villcode,schoolno,schoolname FROM villageschool", pw)
    crow = q("SELECT * FROM cschoolclass", pw)
    ck = list(crow[0].keys()) if crow else []
    classes = {r[ck[0]]: r[ck[1]] for r in crow}          # รหัสชั้น -> ชื่อชั้น
    cls_by_norm = {norm_class(n): c for c, n in classes.items() if norm_class(n)}

    # เลือกโรงเรียน
    rname = next((r["school"] for r in roster if r["school"]), "")
    cand = [s for s in schools if rname and norm_name(rname) in norm_name(s["schoolname"])] or schools
    if len(cand) == 1:
        sc = cand[0]
    else:
        print("เลือกโรงเรียนใน JHCIS ที่ตรงกับไฟล์นี้" + (f" ({rname})" if rname else "") + ":")
        for i, s in enumerate(cand, 1):
            print(f"  {i}) {s['schoolname']}  [หมู่บ้าน {s['villcode']} เลขที่ {s['schoolno']}]")
        while True:                                            # พิมพ์ผิด/กด Enter เปล่า ให้ถามใหม่ ไม่ล้ม
            v = input("พิมพ์เลข แล้ว Enter: ").strip()
            if v.isdigit() and 1 <= int(v) <= len(cand):
                sc = cand[int(v) - 1]; break
            print(f"  ต้องพิมพ์ตัวเลข 1-{len(cand)} (แป้นพิมพ์ ENG)")
    skey = (sc["villcode"], sc["schoolno"])
    print("โรงเรียนใน JHCIS:", sc["schoolname"])

    by_id = {}
    for p in persons:
        by_id.setdefault(norm_id(p["idcard"]), []).append(p)
    by_nb = {}
    for p in persons:
        by_nb.setdefault((norm_name(p["fname"]), (p["birth"] or "")[:10]), []).append(p)
    by_nl = {}
    for p in persons:
        by_nl.setdefault((norm_name(p["fname"]), norm_name(p["lname"])), []).append(p)
    stud = {(s["pcucodeperson"], s["pid"]): s for s in studs}
    sname = {(s["villcode"], s["schoolno"]): s["schoolname"] for s in schools}

    G = {k: [] for k in "12345"}
    matched = set()
    for r in roster:
        base = [r["cls"], r["id"], r["pre"], r["fn"], r["ln"], r["bd"]]
        want = cls_by_norm.get(norm_class(r["cls"]))
        ps = by_id.get(r["id"], []) if r["id"] else []
        if not ps:
            c2 = by_nb.get((norm_name(r["fn"]), r["bd"]), []) if r["bd"] else []
            c3 = by_nl.get((norm_name(r["fn"]), norm_name(r["ln"])), []) if r["ln"] and r["ln"] != "-" else []
            seen, cand = set(), []
            for p in c2 + c3:                                   # ชื่อ+วันเกิดตรง หรือ ชื่อ+สกุลตรง (วันเกิดอาจบันทึกต่างกัน)
                if p["pid"] not in seen:
                    seen.add(p["pid"]); cand.append(p)
            hint = "; ".join(f"pid {p['pid']} {p['fname']} {p['lname']} เกิด {(p['birth'] or '')[:10]} เลข {p['idcard'] or '-'}" for p in cand[:3])
            G["4"].append(base + [r["room"], r["nat"], nation_code(r["nat"]) if r["nat"] else "099",
                                  r["race"], nation_code(r["race"]) if r["race"] else "099", r["rel"], r["addr"], hint])
            continue
        p = ps[0]
        note = "พบเลขนี้มากกว่า 1 คนใน JHCIS ตรวจด้วยคน" if len(ps) > 1 else ""
        if norm_name(p["fname"]) != norm_name(r["fn"]):
            note = (note + " | " if note else "") + f"ชื่อไม่ตรง JHCIS: {p['fname']} {p['lname']}"
        key = (p["pcucodeperson"], p["pid"])
        matched.add(key)
        s = stud.get(key)
        ext = [p["pid"]] + [p.get(c) or "" for c in extra]
        if want is None:
            note = (note + " | " if note else "") + f"ไม่พบรหัสชั้นของ '{r['cls']}' ใน JHCIS"
        if s and (s["villcode"], s["schoolno"]) == skey:
            cur = classes.get(s["classeducate"], s["classeducate"] or "")
            room_ok = (not r["room"]) or str(s["classroom"] or "") == r["room"]
            if want is not None and s["classeducate"] == want and room_ok:
                G["1"].append(base + ext + [cur, note])
            else:
                G["2"].append(base + ext + [cur + ("" if room_ok else f" ห้อง {s['classroom'] or '-'}"), classes.get(want, ""), note, r["room"]])
        else:
            where = "ยังไม่ผูกกับโรงเรียนใด" if not s else "ตอนนี้อยู่: " + sname.get((s["villcode"], s["schoolno"]), "โรงเรียนอื่น") + " " + str(classes.get(s["classeducate"], ""))
            G["3"].append(base + ext + [where, classes.get(want, ""), note, r["room"]])
    pmap = {(p["pcucodeperson"], p["pid"]): p for p in persons}
    for key, s in stud.items():
        if (s["villcode"], s["schoolno"]) == skey and key not in matched:
            p = pmap.get(key, {})
            G["5"].append([classes.get(s["classeducate"], s["classeducate"] or ""), s["yeareducate"] or "", key[1],
                           p.get("idcard", ""), p.get("fname", ""), p.get("lname", ""), (p.get("birth") or "")[:10]] + [p.get(c) or "" for c in extra])

    # เขียนผล
    from openpyxl import Workbook
    from openpyxl.styles import Font
    B = ["ชั้น(ไฟล์)", "เลขประจำตัว", "คำนำหน้า", "ชื่อ", "นามสกุล", "วันเกิด(ค.ศ.)"]
    E = ["pid"] + extra
    spec = [("1 ถูกแล้ว", B + E + ["ชั้นใน JHCIS", "หมายเหตุ"], "1"),
            ("2 แก้ชั้น", B + E + ["ชั้นใน JHCIS ตอนนี้", "ต้องแก้เป็น", "หมายเหตุ", "ห้อง"], "2"),
            ("3 เพิ่มเข้าโรงเรียน", B + E + ["สถานะใน JHCIS", "ชั้นที่ต้องใส่", "หมายเหตุ", "ห้อง"], "3"),
            ("4 ไม่มีใน JHCIS", B + ["ห้อง", "สัญชาติ(ไฟล์)", "รหัสสัญชาติ", "เชื้อชาติ(ไฟล์)", "รหัสเชื้อชาติ", "ศาสนา",
                                    "ที่อยู่จากไฟล์", "คนใน JHCIS ที่ชื่อตรง (ให้คนตัดสิน)"], "4"),
            ("5 ค้างในโรงเรียน", ["ชั้นใน JHCIS", "ปีการศึกษา", "pid", "เลขประจำตัว", "ชื่อ", "นามสกุล", "วันเกิด"] + extra, "5")]
    wb = Workbook()
    ws = wb.active
    ws.title = "สรุป"
    ws.append(["โรงเรียน", sc["schoolname"]]); ws.append(["ไฟล์รายชื่อ", os.path.basename(path)])
    ws.append(["ฐานข้อมูล", f"{info} ({HOST})"]); ws.append(["วันที่เทียบ", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws.append([]); ws.append(["กลุ่ม", "จำนวน (คน)"])
    for name, head, k in spec:
        ws.append([name, len(G[k])])
        w = wb.create_sheet(name)
        w.append(head)
        for c in w[1]:
            c.font = Font(bold=True)
        for row in sorted(G[k], key=lambda x: [str(v) for v in x[:1]]):
            w.append(row)
        for i, hname in enumerate(head, 1):
            w.column_dimensions[w.cell(1, i).column_letter].width = max(12, min(45, len(hname) + 6))
    ws.append(["รวมในไฟล์รายชื่อ", len(roster)])
    ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 40
    out = os.path.join(os.path.dirname(os.path.abspath(path)), "ผลเทียบ_" + os.path.splitext(os.path.basename(path))[0] + ".xlsx")
    wb.save(out)
    print()
    for name, head, k in spec:
        print(f"  {name:<22} {len(G[k]):>4} คน")
    print("\nบันทึกผลที่:", out)


def cli():
    main()


if __name__ == "__main__":
    cli()
