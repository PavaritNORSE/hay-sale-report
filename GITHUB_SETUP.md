# GitHub Setup — HAY Sale Report (Cloud 100%)

> อัปเดต 11 Jun 2026 (v3.13) — รันอัตโนมัติบน GitHub cloud **ไม่ต้องเปิดเครื่องตัวเองเลย**

## ระบบทำงานอย่างไร

- GitHub Actions รัน pipeline บนเครื่อง cloud ของ GitHub (`ubuntu-latest`) ทุกวัน
  **06:00 เวลาไทย** (23:00 UTC) — laptop ปิดเครื่องได้เลย
- ส่วน Export PDF เปลี่ยนจาก Microsoft Excel COM → **LibreOffice headless**
  (`pdf_libre.py`) เพราะเครื่อง cloud ไม่มี Excel
- บนเครื่อง Windows ของคุณ ทุกอย่างยังทำงานเหมือนเดิม 100% — `run_fetch_odoo.bat`
  ยังใช้ Excel ตามปกติ (โค้ดเลือกเองอัตโนมัติ: มี pywin32 → Excel, ไม่มี → LibreOffice)
- อีเมล + PDF attachment ส่งผ่าน Microsoft Graph API เหมือนเดิมทุกประการ
- PDF ทุกรอบเก็บเป็น Artifact บน GitHub 30 วัน (ดาวน์โหลดย้อนหลังได้)

## ⚠️ ข้อควรรู้ก่อนใช้

1. **PDF จาก LibreOffice อาจหน้าตาต่างจาก Excel เล็กน้อย** (font substitution,
   การตัดหน้า) — รอบแรก ๆ ให้เปิดเทียบกับ PDF เดิมที่ Excel สร้าง ถ้า layout เพี้ยน
   ให้แจ้ง Claude ปรับ
2. **Schedule ของ GitHub อาจคลาดเวลาได้ ~15-30 นาที** ช่วงที่ระบบหนาแน่น
   (อีเมลอาจมาถึง 6:00-6:45 น.)
3. **Repo ที่ไม่มี commit ใหม่เกิน 60 วัน → GitHub ปิด schedule อัตโนมัติ**
   จะมีอีเมลเตือนจาก GitHub ให้กด "Enable workflow" ใน tab Actions
4. ต้องเป็น **private repo** เท่านั้น

---

## ขั้นตอนที่ 1 — สร้าง repo และ push (ทำครั้งเดียว)

Repo ในเครื่องถูก init + commit ไว้ให้แล้ว เหลือแค่:

1. ติดตั้ง GitHub CLI: `winget install GitHub.cli` (เปิด PowerShell ใหม่หลังติดตั้ง)
2. Login: `gh auth login` (เลือก GitHub.com → HTTPS → browser login)
3. สร้าง **private** repo แล้ว push:

```powershell
cd "C:\Users\WP Admin\Desktop\01-Sale Report"
gh repo create hay-sale-report --private --source . --push
```

## ขั้นตอนที่ 2 — เพิ่ม Secret `.env`

ไฟล์ `.env` ไม่ถูก commit (มีรหัสผ่าน Odoo + Graph client secret)
Workflow จะสร้าง `.env` ขึ้นเองตอนรันจาก GitHub Secret:

1. เปิด repo บนเว็บ → **Settings → Secrets and variables → Actions**
2. **New repository secret** → Name: `ENV_FILE`
3. Value: copy เนื้อหา **ทั้งไฟล์** `.env` ในเครื่องนี้วางลงไป
4. เมื่อ rotate secret ในอนาคต → แก้ทั้ง `.env` ในเครื่อง **และ** Secret บน GitHub

## ขั้นตอนที่ 3 — ทดสอบ

1. เปิด repo → **Actions → Daily Sale Report → Run workflow**
2. ใส่วันที่ทดสอบ (เช่น `2026-06-10`) หรือเว้นว่าง = เมื่อวาน
3. ดู log สดบนหน้าเว็บ — สำเร็จ = มีอีเมลส่งถึงผู้รับ + PDF อยู่ใน Artifacts ท้ายหน้า run
4. **เปิด PDF เทียบกับของวันเดียวกันที่ Excel เคยสร้าง** (`Daily PDFs\2026\Jun\10\`)
   ว่า layout / ตัวเลข / ภาษาไทย ถูกต้อง

## ขั้นตอนที่ 4 — ปิด Task Scheduler ตัวเดิม

เมื่อ GitHub รันได้แล้ว ปิดตัวเดิมบน laptop ไม่งั้นอีเมลส่งซ้ำ 2 รอบทุกเช้า:

```powershell
# run as Administrator ใน folder นี้
.\unregister_scheduled_task.ps1
```

(`run_fetch_odoo.bat` ยังเก็บไว้กดรัน manual ในเครื่องได้ตามเดิม)

---

## การใช้งานประจำวัน

| ต้องการ | วิธี |
|---------|------|
| รายงานอัตโนมัติทุกเช้า | ไม่ต้องทำอะไร — GitHub รันเอง 6 โมงเช้า |
| สั่งรันทันที / เลือกวันย้อนหลัง | เว็บ/แอป GitHub → Actions → Run workflow → ใส่วันที่ |
| ดู log การรัน | Actions → เลือก run ที่ต้องการ |
| โหลด PDF ย้อนหลัง | Actions → run นั้น ๆ → Artifacts (เก็บ 30 วัน) |
| เปลี่ยนผู้รับอีเมล | แก้ Secret `ENV_FILE` บน GitHub (และ `.env` ในเครื่องให้ตรงกัน) |
| แก้โค้ด | แก้ในเครื่อง → `git add -A; git commit -m "..."; git push` |

## เปลี่ยนเวลารัน

แก้บรรทัด cron ใน [.github/workflows/daily-report.yml](.github/workflows/daily-report.yml)
— ค่าเป็น **UTC** (เวลาไทย −7 ชม.) เช่น รัน 07:00 ไทย = `0 0 * * *` แล้ว push
