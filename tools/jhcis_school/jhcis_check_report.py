"""
jhcis_check_report.py  (อ่านอย่างเดียว ไม่แก้อะไรในฐาน)
ตรวจว่าข้อมูลนักเรียนใน JHCIS พร้อมขึ้นรายงาน (43 แฟ้ม / HDC) แค่ไหน
ต่อฐานเดียวกับที่ JHCIS เครื่องนี้ต่ออยู่ (อ่านจาก database.properties)

ใช้:  python jhcis_check_report.py
ได้:  สรุปบนจอ + ไฟล์ ตรวจรายงาน_<วันที่>.txt
"""
import os
import time
import getpass
import jhcis_roster_compare as rc


def cli():

    out = []


    def say(s=""):
        print(s); out.append(s)


    srv = None
    try:
        for ln in open(r"C:\Program Files\JHCIS\database.properties", encoding="utf-8", errors="replace"):
            if ln.strip().upper().startswith("SERVER="):
                srv = ln.split("=", 1)[1].strip()
    except OSError:
        pass
    if srv and srv.lower() not in ("localhost", "127.0.0.1"):
        rc.HOST = srv
    say(f"ตรวจฐาน: {rc.HOST}  เวลา {time.strftime('%Y-%m-%d %H:%M')}")
    pw = rc.ask_pw(f"รหัสผ่าน root ของ MySQL ที่ {rc.HOST}: ")
    q = lambda s: rc.q(s, pw)

    year = time.localtime().tm_year + 543 - (0 if time.localtime().tm_mon >= 5 else 1)
    schools = q("SELECT villcode,schoolno,schoolname FROM villageschool")
    say(f"\n1) จำนวนนักเรียนปีการศึกษา {year} แยกโรงเรียน/ชั้น")
    for sc in schools:
        rows = q(f"SELECT s.classeducate c, COUNT(*) n, SUM(s.classroom IS NULL OR TRIM(s.classroom)='') noroom, "
                 f"SUM(p.birth IS NULL) nobirth, SUM(p.idcard IS NULL OR TRIM(p.idcard)='') noid, "
                 f"SUM(p.typelive='4') out4, SUM(p.dischargetype IS NOT NULL AND p.dischargetype<>'9') gone "
                 f"FROM personstudent s JOIN person p ON p.pid=s.pid AND p.pcucodeperson=s.pcucodeperson "
                 f"WHERE s.villcode='{sc['villcode']}' AND s.schoolno='{sc['schoolno']}' AND s.yeareducate='{year}' "
                 f"GROUP BY s.classeducate ORDER BY s.classeducate+0")
        if not rows:
            continue
        tot = sum(int(r["n"]) for r in rows)
        say(f"  {sc['schoolname']}: รวม {tot} คน")
        for r in rows:
            warn = []
            for k, t in (("noroom", "ไม่มีห้อง"), ("nobirth", "ไม่มีวันเกิด"), ("noid", "ไม่มีเลขบัตร"), ("gone", "จำหน่ายแล้ว")):
                if int(r[k] or 0):
                    warn.append(f"{t} {r[k]}")
            say(f"     ชั้นรหัส {r['c']:>2}: {r['n']:>4} คน | นอกเขต(typelive 4) {r['out4']}" + ("  ⚠ " + ", ".join(warn) if warn else ""))
        old = q(f"SELECT COUNT(*) n FROM personstudent WHERE villcode='{sc['villcode']}' AND schoolno='{sc['schoolno']}' "
                f"AND (yeareducate IS NULL OR yeareducate<>'{year}')")
        say(f"     ยังค้างปีเก่า (กลุ่ม 5): {old[0]['n']} คน")

    say("\n2) คนซ้ำ: เลขบัตรเดียวกันมีหลาย HN (เฉพาะที่เป็นนักเรียนปีนี้)")
    dup = q(f"SELECT p.idcard, GROUP_CONCAT(p.pid) hns, COUNT(*) n FROM person p WHERE p.idcard IN "
            f"(SELECT p2.idcard FROM personstudent s JOIN person p2 ON p2.pid=s.pid AND p2.pcucodeperson=s.pcucodeperson "
            f"WHERE s.yeareducate='{year}' AND p2.idcard IS NOT NULL AND TRIM(p2.idcard)<>'') GROUP BY p.idcard HAVING COUNT(*)>1")
    say(f"  {len(dup)} เลข" + "".join(f"\n     {d['idcard']}: HN {d['hns']}" for d in dup[:30]))

    say("\n3) ตารางที่เก็บผลตรวจฟัน และช่องเกี่ยวกับโรงเรียน/ชั้น (ใช้ในแฟ้ม DENTAL)")
    for t in q("SHOW TABLES LIKE '%dent%'"):
        name = list(t.values())[0]
        cols = [c["Field"] for c in q(f"SHOW COLUMNS FROM {name}")]
        rel = [c for c in cols if any(k in c.lower() for k in ("school", "class", "room", "edu"))]
        say(f"  {name}: {len(cols)} ช่อง" + (f" | ช่องโรงเรียน/ชั้น: {', '.join(rel)}" if rel else ""))

    f = os.path.join(rc.home(), f"ตรวจรายงาน_{time.strftime('%Y%m%d_%H%M')}.txt")
    open(f, "w", encoding="utf-8").write("\n".join(out))
    say(f"\nบันทึกผลที่: {f}")



if __name__ == "__main__":
    cli()
