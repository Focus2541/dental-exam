"""
jhcis_compare_person.py  (อ่านอย่างเดียว)
เทียบข้อมูลคน 1 คน (ตาราง person) ระหว่าง ฐานทดสอบในเครื่อง (สำเนาวันที่ 15 พ.ย.) กับ ฐานจริง 192.168.1.2
ใช้ดูว่าช่องไหนถูกเปลี่ยนไป เพื่อแก้กลับให้ถูก

ใช้:  python jhcis_compare_person.py 17168
"""
import sys
import getpass
import jhcis_roster_compare as rc


def cli():

    if len(sys.argv) < 2 or not sys.argv[1].isdigit():
        sys.exit("ใช้: python jhcis_compare_person.py HN")
    pid = int(sys.argv[1])
    rows = {}
    for name, host in (("ฐานทดสอบ (สำเนา 15 พ.ย.)", "127.0.0.1"), ("ฐานจริง", "192.168.1.2")):
        rc.HOST = host
        pw = rc.ask_pw(f"รหัสผ่าน root ของ MySQL ที่ {host} ({name}): ")
        r = rc.q(f"SELECT * FROM person WHERE pid={pid}", pw)
        rows[name] = r[0] if r else {}
        print(f"  {name}: {'พบ' if r else 'ไม่พบ'} HN {pid}")
    a, b = rows.values()
    diff = [k for k in sorted(set(a) | set(b)) if str(a.get(k)) != str(b.get(k))]
    print(f"\nHN {pid}: ช่องที่ต่างกัน {len(diff)} ช่อง")
    print(f"{'ช่อง':<22}{'ค่าเดิม (15 พ.ย.)':<30}ค่าตอนนี้ (ฐานจริง)")
    for k in diff:
        print(f"{k:<22}{str(a.get(k)):<30}{b.get(k)}")



if __name__ == "__main__":
    cli()
