# -*- coding: utf-8 -*-
"""
ดูว่า JHCIS เขียนข้อมูลลงตารางไหน ตอนกรอกผลตรวจฟันนักเรียน 1 คน
วิธีใช้: ดับเบิลคลิก run_capture.bat

ขั้นตอนของโปรแกรม
 1. ต่อ MySQL ของ JHCIS (แนะนำ: ฉบับทดสอบในเครื่องนี้ localhost)
 2. เปิด general log ของ MySQL ชั่วคราว (จดคำสั่ง SQL ทุกคำสั่งที่ JHCIS ส่ง)
 3. ให้คุณกรอกผลตรวจฟันเด็ก 1 คนในโปรแกรม JHCIS ตามปกติ แล้วกด Enter
 4. เก็บเฉพาะคำสั่งในช่วงนั้น + โครงสร้างตารางที่ถูกเขียน ลงไฟล์ .txt
 5. ปิด general log คืนค่าเดิม

โปรแกรมนี้ไม่แก้ ไม่ลบ ข้อมูลผู้ป่วยใน JHCIS
"""
import datetime
import getpass
import os
import re
import sys
import time

try:
    import pymysql
except ImportError:
    print('ไม่พบ pymysql ให้รัน:  python -m pip install pymysql')
    input('กด Enter เพื่อปิด')
    sys.exit(1)

HERE = os.path.dirname(os.path.abspath(__file__))
WRITE_RE = re.compile(r'^\s*(INSERT|UPDATE|DELETE|REPLACE)\b', re.I)
TABLE_RE = re.compile(r'^\s*(?:INSERT|REPLACE)\s+(?:LOW_PRIORITY\s+|DELAYED\s+|IGNORE\s+)*(?:INTO\s+)?`?(\w+)`?'
                      r'|^\s*UPDATE\s+(?:LOW_PRIORITY\s+|IGNORE\s+)*`?(\w+)`?'
                      r'|^\s*DELETE\s+(?:LOW_PRIORITY\s+|QUICK\s+|IGNORE\s+)*FROM\s+`?(\w+)`?', re.I)


def ask(label, default):
    v = input(f'{label} [{default}]: ').strip()
    return v or default


def txt(x):
    if isinstance(x, (bytes, bytearray)):
        for enc in ('utf-8', 'tis-620', 'cp874', 'latin-1'):
            try:
                return x.decode(enc)
            except UnicodeDecodeError:
                pass
    return '' if x is None else str(x)


def main():
    print('=' * 60)
    print(' จับคำสั่งที่ JHCIS ใช้ตอนบันทึกผลตรวจฟัน')
    print(' แนะนำ: ใช้กับ JHCIS ฉบับทดสอบในเครื่องนี้ (localhost)')
    print('=' * 60)
    host = ask('เครื่อง MySQL (host)', 'localhost')
    port = int(ask('พอร์ต', '3333'))
    db = ask('ชื่อฐานข้อมูล', 'jhcisdb')
    user = ask('ผู้ใช้ (ต้องเป็น root หรือมีสิทธิ์ SUPER)', 'root')
    pw = getpass.getpass('รหัสผ่าน (พิมพ์แล้วจะไม่ขึ้นบนจอ): ')
    if host not in ('localhost', '127.0.0.1'):
        print('\n*** คุณกำลังต่อเครื่องอื่น ไม่ใช่ฉบับทดสอบในเครื่องนี้ ***')
        print('โปรแกรมไม่แก้ข้อมูล แต่การกรอกผลใน JHCIS จะลงข้อมูลจริง')
        if input('พิมพ์ yes เพื่อทำต่อ: ').strip().lower() != 'yes':
            return

    con = pymysql.connect(host=host, port=port, user=user, password=pw, database=db,
                          charset='utf8', autocommit=True, connect_timeout=10)
    cur = con.cursor()
    cur.execute('SELECT VERSION(), CONNECTION_ID(), NOW()')
    ver, my_id, _ = cur.fetchone()
    cur.execute("SHOW GLOBAL VARIABLES WHERE Variable_name IN ('general_log','log_output')")
    old = {k.lower(): v for k, v in cur.fetchall()}
    print(f'\nต่อได้แล้ว MySQL {ver}  (general_log เดิม={old.get("general_log")}, log_output เดิม={old.get("log_output")})')

    try:
        cur.execute("SET GLOBAL log_output = 'TABLE'")
        cur.execute("SET GLOBAL general_log = 'ON'")
    except pymysql.err.OperationalError as e:
        print('\nเปิด general log ไม่ได้: ผู้ใช้นี้ไม่มีสิทธิ์ SUPER ให้ใช้ root ของ MySQL ในเครื่องทดสอบ')
        print('รายละเอียด:', e)
        return

    time.sleep(1)
    cur.execute('SELECT NOW()')
    t0 = cur.fetchone()[0]
    print('\n' + '-' * 60)
    print(' ตอนนี้ให้เปิดโปรแกรม JHCIS (ที่ต่อฐานข้อมูลเดียวกันนี้)')
    print(' 1) เลือกเด็กนักเรียน 1 คนที่มี visit วันนี้ (เสียบบัตร/Authen แล้ว)')
    print(' 2) เข้าหน้าตรวจฟัน / อนามัยโรงเรียน กรอกผลตามปกติ')
    print('    ลองใส่ตัวเลขต่างกันทุกช่อง เช่น ฟันน้ำนม 20 ผุ 3 อุด 1 ถอน 2 ฟันแท้ 6 ผุ 1 ...')
    print('    และติ๊กเหงือกอักเสบ หินปูน ต้องอุด ต้องถอน Sealant ให้ครบ')
    print(' 3) กดบันทึกให้เรียบร้อย')
    print('-' * 60)
    note = input('จดสิ่งที่กรอกไว้สั้น ๆ (เช่น HN/ชื่อเด็ก และตัวเลขที่ใส่) แล้วกด Enter เมื่อบันทึกใน JHCIS เสร็จ:\n> ')
    cur.execute('SELECT NOW()')
    t1 = cur.fetchone()[0]

    cur.execute("SELECT event_time, thread_id, command_type, argument FROM mysql.general_log "
                "WHERE event_time BETWEEN %s AND %s AND thread_id <> %s "
                "AND command_type IN ('Query','Execute') ORDER BY event_time", (t0, t1, my_id))
    rows = cur.fetchall()

    # คืนค่าเดิม
    try:
        cur.execute(f"SET GLOBAL general_log = '{old.get('general_log', 'OFF')}'")
        cur.execute(f"SET GLOBAL log_output = '{old.get('log_output', 'FILE')}'")
        if str(old.get('general_log', 'OFF')).upper() == 'OFF':
            cur.execute('TRUNCATE TABLE mysql.general_log')
    except Exception as e:  # ไม่ร้ายแรง
        print('คืนค่า general log ไม่ครบ:', e)

    writes, tables = [], []
    for et, tid, ct, arg in rows:
        a = txt(arg)
        if WRITE_RE.match(a):
            writes.append((et, tid, a))
            m = TABLE_RE.match(a)
            t = m and (m.group(1) or m.group(2) or m.group(3))
            if t and t.lower() not in [x.lower() for x in tables]:
                tables.append(t)

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out = os.path.join(HERE, f'jhcis_capture_{stamp}.txt')
    with open(out, 'w', encoding='utf-8') as f:
        f.write(f'JHCIS capture {stamp}\nMySQL {ver} host={host} db={db}\nช่วงเวลา {t0} ถึง {t1}\n')
        f.write(f'สิ่งที่กรอก: {note}\n\n')
        f.write(f'== ตารางที่ถูกเขียน ({len(tables)}) ==\n' + '\n'.join(tables) + '\n\n')
        f.write(f'== คำสั่งที่เขียนข้อมูล ({len(writes)}) ==\n')
        for et, tid, a in writes:
            f.write(f'[{et}] thread {tid}\n{a}\n\n')
        for t in tables:
            try:
                cur.execute(f'SHOW CREATE TABLE `{t}`')
                f.write(f'== โครงสร้าง {t} ==\n{cur.fetchone()[1]}\n\n')
            except Exception as e:
                f.write(f'== โครงสร้าง {t} == อ่านไม่ได้: {e}\n\n')
        f.write(f'== คำสั่งทั้งหมดในช่วงนี้ รวม SELECT ({len(rows)}) ==\n')
        for et, tid, ct, arg in rows:
            f.write(f'[{et}] {tid} {ct}: {txt(arg)[:2000]}\n')
    con.close()

    print('\n' + '=' * 60)
    if not writes:
        print(' ไม่พบคำสั่งบันทึกข้อมูลในช่วงนี้')
        print(' ตรวจว่า JHCIS ต่อฐานข้อมูลเดียวกับที่ใส่ไว้ และกดบันทึกแล้ว แล้วลองใหม่')
    else:
        print(f' เจอคำสั่งบันทึก {len(writes)} คำสั่ง ใน {len(tables)} ตาราง: {", ".join(tables)}')
    print(f' ไฟล์ผล: {out}')
    print(' ส่งไฟล์นี้ให้ Claude ได้เลย')
    print('=' * 60)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('\nเกิดข้อผิดพลาด:', e)
    input('\nกด Enter เพื่อปิด')
