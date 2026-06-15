# GitHub Setup — HAY Sale Report (Cloud 100%)

> อัปเดต 15 Jun 2026 (v3.14) — รันอัตโนมัติบน GitHub cloud **ไม่ต้องเปิดเครื่องตัวเองเลย**

## ระบบทำงานอย่างไร

- **cron-job.org** (ฟรี) เรียก GitHub API ตรงเวลา **06:00 เวลาไทย** (23:00 UTC) ทุกวัน
  → GitHub Actions รัน pipeline บนเครื่อง cloud (`ubuntu-latest`) — laptop ปิดเครื่องได้เลย
- Backup: ถ้า cron-job.org ล้มเหลว GitHub schedule จะ kick in ให้อัตโนมัติ (~08:47 น.)
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
2. **Primary trigger คือ cron-job.org** (ดูขั้นตอนที่ 3 ด้านล่าง) — ถ้ายังไม่ได้ตั้ง
   อีเมลจะมาถึง ~8:47 น. แทน (backup schedule ของ GitHub Actions)
3. **Repo ที่ไม่มี commit ใหม่เกิน 60 วัน → GitHub ปิด schedule อัตโนมัติ**
   จะมีอีเมลเตือนจาก GitHub ให้กด "Enable workflow" ใน tab Actions
4. ต้องเป็น **private repo** เท่านั้น

---

---

## ขั้นตอนที่ 3 — ตั้ง External Scheduler (cron-job.org) ⬅ สำคัญมาก

> **ทำไมต้องมี?** GitHub's built-in schedule ที่ 23:00 UTC (6 โมงเช้าไทย) ล่าช้าได้ถึง ~12 ชม.
> เพราะเป็นช่วง peak traffic ของ GitHub → อีเมลมาถึงตอน 5 โมงเย็น
> cron-job.org จะ "กด Run" แทนเราตรงเวลาผ่าน GitHub API — ฟรี ไม่ต้องเปิดเครื่อง

### 3.1 สร้าง GitHub PAT (Personal Access Token)

1. GitHub.com → ชื่อผู้ใช้ (มุมขวาบน) → **Settings**
2. เลื่อนลงสุด → **Developer settings** → **Personal access tokens** → **Fine-grained tokens**
3. **Generate new token**
   - Token name: `cron-job.org trigger`
   - Expiration: **1 year** (จะมีอีเมลเตือนก่อนหมดอายุ)
   - Repository access: **Only select repositories** → `hay-sale-report`
   - Permissions → **Actions** → `Read and write`
4. กด **Generate token** → copy token ทันที (ดูได้ครั้งเดียว)

### 3.2 ตั้ง cron-job.org

1. เปิด **https://cron-job.org** → สมัครฟรี → Login
2. **CREATE CRONJOB**
3. ตั้งค่าดังนี้:

| ช่อง | ค่า |
|------|-----|
| Title | HAY Sale Report trigger |
| URL | `https://api.github.com/repos/PavaritNORSE/hay-sale-report/actions/workflows/daily-report.yml/dispatches` |
| Execution schedule | เลือก **Custom** → เวลา `23:00` UTC ทุกวัน |
| Request method | **POST** |
| Request body | `{"ref":"main"}` |

4. กด **Advanced** → เพิ่ม Headers 2 ตัว:
   - `Authorization` = `Bearer ghp_XXXXXXXXXXXXXXXX` (ใส่ token จาก 3.1)
   - `Accept` = `application/vnd.github+json`
5. **CREATE** → ทดสอบกด **Run now** ดูว่า HTTP 204 = สำเร็จ

### 3.3 ทดสอบ end-to-end

1. กด **Run now** บน cron-job.org
2. เปิด GitHub → Actions → ควรเห็น run ใหม่ (`workflow_dispatch`)
3. รอ ~2 นาที → มีอีเมลส่งออกไป

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

เวลารันปรับที่ **cron-job.org** (Execution schedule) ตรง ๆ ไม่ต้องแก้โค้ด
— เวลาบนเว็บ cron-job.org คือ UTC (เวลาไทย −7 ชม.) เช่น อยาก 07:00 ไทย = 00:00 UTC
