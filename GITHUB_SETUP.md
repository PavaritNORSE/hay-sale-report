# GitHub Setup — HAY Sale Report

> สร้างเมื่อ 11 Jun 2026 — ขั้นตอนนำโปรเจคขึ้น GitHub + รันอัตโนมัติ

## ⚠️ ข้อจำกัดสำคัญที่ต้องเข้าใจก่อน

Pipeline นี้ **export PDF ผ่าน Microsoft Excel COM (win32com)** ซึ่งหมายความว่า:

- **GitHub-hosted runner (cloud) รันไม่ได้** — เครื่อง cloud ของ GitHub ไม่มี Excel ติดตั้ง
- การรันอัตโนมัติผ่าน GitHub Actions ต้องใช้ **self-hosted runner** = ติดตั้งตัว runner
  บนเครื่อง Windows ที่มี Excel (เช่น laptop เครื่องนี้)
- เครื่องที่เป็น runner ยังต้องเปิดอยู่ + logged in ตอน 6 โมงเช้า
  (ข้อจำกัดเดียวกับ Task Scheduler ปัจจุบันทุกประการ)

**สิ่งที่ GitHub Actions ให้เพิ่มจาก Task Scheduler:**
- กดรันจากมือถือ/เครื่องอื่นได้ (workflow_dispatch + เลือกวันที่ได้)
- ดู log ทุกรอบบนเว็บ + ดาวน์โหลด PDF ย้อนหลังจาก Artifacts
- โค้ดมี version control — rollback ได้ทุกเวอร์ชัน

**ถ้าอยากรันบน cloud 100% (ไม่ง้อ laptop):** ต้องเขียนส่วน PDF ใหม่ให้ไม่ใช้ Excel
(เช่น LibreOffice headless) — เป็นงาน rewrite แยกต่างหาก ยังไม่ได้ทำ

---

## ขั้นตอนที่ 1 — สร้าง repo และ push (ทำครั้งเดียว)

Repo ในเครื่องถูก init + commit ไว้ให้แล้ว เหลือแค่ push:

1. ติดตั้ง GitHub CLI: `winget install GitHub.cli` (เปิด PowerShell ใหม่หลังติดตั้ง)
2. Login: `gh auth login` (เลือก GitHub.com → HTTPS → browser login)
3. สร้าง **private** repo แล้ว push:

```powershell
cd "C:\Users\WP Admin\Desktop\01-Sale Report"
gh repo create hay-sale-report --private --source . --push
```

> ⚠️ ต้องเป็น **private** เท่านั้น — โค้ดมีรายละเอียดธุรกิจ/email ภายใน

## ขั้นตอนที่ 2 — เพิ่ม Secret `.env`

ไฟล์ `.env` ไม่ถูก commit (มีรหัสผ่าน Odoo + Graph client secret)
ต้องเอาเนื้อหาไปใส่เป็น GitHub Secret เพื่อให้ workflow สร้าง `.env` เองตอนรัน:

1. เปิด repo บนเว็บ → **Settings → Secrets and variables → Actions**
2. **New repository secret** → Name: `ENV_FILE`
3. Value: copy เนื้อหา **ทั้งไฟล์** `.env` ในเครื่องนี้วางลงไป
4. เมื่อ rotate secret ในอนาคต → แก้ทั้ง `.env` ในเครื่อง และ Secret บน GitHub

## ขั้นตอนที่ 3 — ติดตั้ง Self-hosted Runner

1. เปิด repo บนเว็บ → **Settings → Actions → Runners → New self-hosted runner**
2. เลือก **Windows x64** แล้วทำตามคำสั่งที่หน้าจอแสดง (download + `config.cmd`)
3. ตอน config ใส่ label เพิ่ม: `Windows` (workflow ใช้ `runs-on: [self-hosted, Windows]`)
4. **สำคัญ:** start runner แบบ interactive ด้วย `run.cmd` — **อย่าติดตั้งเป็น Windows
   service** เพราะ Excel COM ใช้ไม่ได้ใน session แบบ non-interactive
5. ตั้งให้ `run.cmd` เปิดอัตโนมัติตอน login: ใส่ shortcut ใน
   `shell:startup` (Win+R → `shell:startup`)

## ขั้นตอนที่ 4 — ทดสอบ

1. เปิด repo → **Actions → Daily Sale Report → Run workflow**
2. ใส่วันที่ทดสอบ (เช่น `2026-06-10`) หรือเว้นว่าง = เมื่อวาน
3. ดู log สด ๆ บนหน้าเว็บ — ถ้าสำเร็จจะมีอีเมลส่ง + PDF อยู่ใน Artifacts

## ⚠️ อย่ารัน 2 ระบบซ้อนกัน

ถ้า GitHub Actions schedule ทำงานแล้ว ให้ **ปิด Task Scheduler ตัวเดิม** ไม่งั้น
อีเมลจะถูกส่ง 2 รอบทุกเช้า:

```powershell
# run as Administrator
.\unregister_scheduled_task.ps1
```

(หรือกลับกัน — ใช้ Task Scheduler ต่อ แล้วใช้ GitHub แค่เก็บโค้ด ก็ได้เช่นกัน
โดยลบ schedule block ออกจาก `.github/workflows/daily-report.yml`)

## การ push การแก้ไขในอนาคต

```powershell
cd "C:\Users\WP Admin\Desktop\01-Sale Report"
git add -A
git commit -m "describe change"
git push
```

> หมายเหตุ: self-hosted runner จะ checkout โค้ดจาก GitHub มารันใน folder ของ runner เอง
> (`actions-runner\_work\...`) ไม่ใช่ folder นี้ — แก้โค้ดแล้วต้อง push ก่อน รอบถัดไปถึงจะมีผล
