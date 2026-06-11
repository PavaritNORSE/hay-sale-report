# HAY Sale Report — Project Summary
> สรุปครบทุกส่วน สำหรับ handoff / team plan session
> อัปเดตล่าสุด: 10 Jun 2026 | Version: **v3.12** (secret expiry warning)
> Project location: `C:\Users\USER\Desktop\Claude Cowork-Workspace\02-Projects\01-Sale Report\`

---

## 1. ภาพรวมโปรเจค

ระบบ automated pipeline ดึงข้อมูลจาก Odoo XML-RPC API แล้ว patch ข้อมูลลงไฟล์ Excel template `Sales Report V3.1.xlsm` โดยตรง (ไม่ผ่าน openpyxl) จากนั้น export PDF รายสาขา 7 ใบ แล้วส่งอีเมลพร้อม attachment ผ่าน Microsoft Graph API

**Flow ทั้งหมด (v3.7):**
```
run_fetch_odoo.bat
  → [1] auto-install: pywin32, requests, python-dotenv, tkcalendar
  → [2] kill stale Excel processes
  → [3] launch run_for_date.py (date picker dialog)
        - 🪟 calendar widget — default=today, max=today
        - countdown 30 วินาทีบนปุ่ม Generate & Send
        - user เลือกวัน (default = today) → คลิก Send หรือรอ auto
  → [4] daily_report.py --fetch-odoo --date YYYY-MM-DD
        - fetch 5 datasets จาก Odoo API (~6-12s, gzip + keep-alive)
        - cleanse data (mirrors VBA Data_Cleansing, + cross-doc leak fix)
        - build YTD rows + CR rows
        - patch 7 sheets ใน xlsm (zipfile bytes-mode XML patching)
        - fix formula cache + fullCalcOnLoad
        - generate 7 PDF files (win32com, ScreenUpdating=Off, Calc=manual)
        - PDFs → Daily PDFs/{YYYY}/{MMM}/{DD}/
  → [5] send email via Microsoft Graph API
        - subject: "Sale Report D MMM YYYY"
        - attach 7 PDFs (base64)
        - to: bharkbhum.best@norserepublics.com
  → [6] on any crash → send failure notification email (best-effort)
```

---

## 2. Prerequisites — สิ่งที่ต้องติดตั้ง

| รายการ | วิธีติดตั้ง | หมายเหตุ |
|--------|------------|----------|
| **Python 3.9+** | https://python.org | ติ๊ก "Add to PATH" ตอนติดตั้ง |
| **pywin32** | `pip install pywin32` | PDF export via Excel COM |
| **requests** | `pip install requests` | Graph API HTTP calls (email) |
| **python-dotenv** | `pip install python-dotenv` | โหลด `.env` credentials |
| **tkcalendar** | `pip install tkcalendar` | date picker dialog |
| Microsoft Excel | ติดตั้งอยู่แล้ว | ต้องใช้ Visible=True ถึงจะ COM call ได้ |

**หมายเหตุ:** bat file `run_fetch_odoo.bat` จะ auto-install ทุก package ถ้าไม่มี ติดตั้งครั้งเดียว ครั้งถัดไปข้ามทันที

---

## 3. ไฟล์ที่เกี่ยวข้องทั้งหมด (v3.8 — 5 Jun 2026 + xlsm renamed Sales Report V3.1)

### 📌 ACTIVE — ใช้จริงรันประจำ (root)

```
01-Sale Report/
├── daily_report.py                          ← v3.8 main script
├── run_fetch_odoo.bat                       ← ดับเบิลคลิก (manual run พร้อม date picker)
├── run_for_date.py                          ← date picker dialog (tkinter)
├── run_scheduled.bat                        ← ใช้โดย Task Scheduler (ไม่มี picker, log)
├── register_scheduled_task.ps1              ← run admin ครั้งเดียวเพื่อ setup scheduled
├── unregister_scheduled_task.ps1            ← ลบ scheduled task
├── .env                                     ← Graph API credentials (PRIVATE — never commit)
├── .env.example                             ← template สำหรับ handoff
├── Sales Report V3.1.xlsm                  ← output xlsm (overwrite ทุก run)
├── Sales Report V3.1.xlsm.bak              ← ⚠️ template base (ห้ามลบ ห้ามแก้)
├── PROJECT_SUMMARY.md                       ← ไฟล์นี้
├── INSTRUCTIONS.md                          ← instructions เก่า (April 2026)
├── Daily PDFs/                              ← PDF output organized by year/month/day
│   └── 2026/
│       ├── May/
│       │   ├── 14/  (7 PDFs)
│       │   ├── 18/  (7 PDFs)
│       │   └── 22/  (7 PDFs)
│       └── Jun/
│           └── 01/  (7 PDFs)
└── __pycache__/                             ← Python bytecode (auto-gen, ignore)
```

### 📦 _Archive/ — ของเก่าเก็บไว้ rollback / reference

```
_Archive/
├── 01-daily_report_backups/                 ← Python script versions (9 ไฟล์)
│   ├── daily_report_v1.0_backup.py
│   ├── daily_report_v2_stable.py
│   ├── daily_report_v3.0_backup.py          (pre-leak-fix)
│   ├── daily_report_v3.1_backup.py          (pre-M1 optimize)
│   ├── daily_report_v3.2_backup.py          (pre-M4)
│   ├── daily_report_v3.3_backup.py          (pre-M5)
│   ├── daily_report_v3.4_backup.py          (pre-M2)
│   ├── daily_report_v3.5b_backup.py         (pre-email)
│   ├── daily_report_v3.6_backup.py          (pre-date-picker)
│   ├── daily_report_v3.7_backup.py          (pre-cleanse-so-multi-line-fix)
│   └── candidates/                          ← 8 mirror files (รอ approve ก่อน apply)
│
├── 02-old_excel_versions/                   ← 8 ไฟล์ Excel เวอร์ชันเก่า
│   ├── Retail Sale Report V1.0 backup.xlsm
│   ├── Retail Sale Report V1.1 backup.xlsm
│   ├── Retail Sale Report V1.2 backup.xlsm + V1.2.xlsm + V1.2 - Copy.xlsm
│   ├── Retail Sale Report V2.0.xlsm
│   ├── Retail Sale Report V3.0.xlsm + V3.0 test.xlsm
│   └── Retail Sale Report V3.0 2026.xlsm + .bak  ← moved 5-Jun-2026 (renamed → Sales Report V3.1)
│
├── 03-reference/                            ← VBA reference สำหรับเปรียบเทียบ
│   └── Sales Report V3.0_reference.xlsm
│
├── 04-test_macros_2024/                     ← Excel debug macros ปี 2024
│   └── test macro + backups (5 ไฟล์)
│
├── 05-old_scripts/                          ← script เก่า/debug (7 ไฟล์)
│   ├── odoo_check.py, odoo_verify_fields.py
│   ├── run_daily_report.bat, run_macros.vbs, run_verify_fields.bat
│   ├── SOPay_before_after.xlsx
│   └── test_write.bin
│
└── 06-old_data/                             ← ไฟล์ data เก่า (4 ไฟล์)
    ├── Report Update 2024 (conflicted copy).xlsx
    ├── Sales Order (sale.order) (42).xlsx
    ├── Sales Report Prototype.xlsx
    └── SO update branch.xlsx
```

### Sheet mapping ใน xlsm (ไม่เปลี่ยน)

| Sheet name | sheet XML | ข้อมูล |
|-----------|-----------|--------|
| Cash Received | sheet1.xml | CR rows (first_row=5) |
| Daily Template | (ไม่ถูก patch) | PDF source sheet |
| YTD | sheet4.xml | YTD rows (first_row=5) |
| SO data | sheet5.xml | SO cleaned (first_row=1) |
| POS data | sheet6.xml | POS cleaned (first_row=1) |
| CN data | sheet7.xml | CN cleaned (first_row=1) |
| SO Payment data | sheet8.xml | SOPay cleaned (first_row=1) |
| POS Payment data | sheet9.xml | POSPay cleaned (first_row=1) |

---

## 4. Odoo Connection Settings

v3.13: connection settings ย้ายจาก `daily_report.py` ไปอยู่ใน `.env` (ห้าม commit):

```ini
ODOO_URL=https://norseodooproduction.com
ODOO_DB=norse-production
ODOO_UID=68
ODOO_PW=<ดูใน .env เท่านั้น — ห้ามเขียนลงไฟล์ที่ commit>
```

`ODOO_YEAR = 2026` ยังอยู่ใน `daily_report.py` (ไม่ใช่ secret)

ใน v3.2+ มี HTTP keep-alive + gzip + TLS warm-up เพื่อลด wall-clock ของ fetch (~25-35% ลดลง)

---

## 5. Email Configuration — Microsoft Graph API (v3.6)

ใช้ Graph API + OAuth 2.0 client credentials แทน SMTP เพื่อไม่ต้องซื้อ license ให้ shared mailbox

### Credentials ใน `.env` (private — ห้าม commit)

```ini
GRAPH_TENANT_ID=aa7e7791-8d80-4e21-9804-576175408655
GRAPH_CLIENT_ID=7e2fc405-0eb1-46d0-8fc0-4d097b552824
GRAPH_CLIENT_SECRET=<from Azure portal → Certificates & secrets>

EMAIL_FROM=noreply@norserepublics.com
EMAIL_FROM_NAME=Norse Republics — Sales

# v3.11: To/CC/BCC ทั้ง 3 ตัวรับ comma-separated list (เว้นบรรทัด/missing = empty)
EMAIL_TO=bharkbhum.best@norserepublics.com, person2@norserepublics.com
EMAIL_CC=manager@norserepublics.com
EMAIL_BCC=

# v3.12 (optional): warn ในอีเมลเมื่อ client secret ใกล้หมดอายุ
GRAPH_SECRET_EXPIRES=2028-06-04
```

**v3.12 — Secret expiry warning:**

| Days remaining | Severity | ในอีเมล body |
|----------------|----------|----------------|
| > 30 days | none | ไม่มี warning |
| 8-30 days | `warning` | "NOTICE: ... please renew before ..." |
| 0-7 days | `critical` | "!!! URGENT ..." |
| expired (<0) | `critical` | "!!! EXPIRED N day(s) ago. Pipeline will FAIL ..." |

ตอน admin สร้าง secret ใหม่ → user แก้ทั้ง `GRAPH_CLIENT_SECRET` + `GRAPH_SECRET_EXPIRES` ใน `.env` พร้อมกัน ไม่ต้อง restart

### Setup ที่ Admin ต้องทำใน Azure AD (ครั้งเดียว)

1. **App registration** ใน Entra (Azure AD) — Single tenant
2. **API permission** = Microsoft Graph > **Mail.Send** (Application type) + Admin consent
3. **Client secret** สร้างใน Certificates & secrets (อายุ 2 ปี — set calendar reminder)
4. **Application Access Policy** ⭐ จำกัด App ให้ส่งได้แค่ noreply mailbox (security):
   ```powershell
   New-ApplicationAccessPolicy `
       -AppId "<CLIENT_ID>" `
       -PolicyScopeGroupId "sale-report-mailers@norserepublics.com" `
       -AccessRight RestrictAccess
   ```

### Email behavior

| Trigger | Subject | Body |
|---------|---------|------|
| pipeline สำเร็จ | `Sale Report 4 Jun 2026` | static — `Dear All, Please see ...` |
| PDF ไม่ครบ 7 | เหมือนกัน | + `NOTE: incomplete — missing: ...` |
| pipeline ล้ม | `Sale Report FAILED 4 Jun 2026` | dynamic — `... while processing the report for 4 Jun 2026 at phase: ...` + traceback |

### Token cache
- Token จาก Microsoft Identity Platform อายุ ~1 ชม. — cache ใน module global ใช้ซ้ำใน run เดียวกัน
- Cost: $0 (Graph API ฟรีสำหรับ tenant ที่มี Exchange Online)

---

## 6. Date Picker — v3.7

### วิธีใช้
1. ดับเบิลคลิก `run_fetch_odoo.bat` (เหมือนเดิม)
2. หลัง auto-install + kill Excel → dialog เด้งขึ้น
3. ปฏิทินแสดงเดือนปัจจุบัน วันนี้ highlight ไว้
4. ปุ่ม `Generate & Send (30)` นับถอยหลัง

| User action | ผล |
|-------------|-----|
| ไม่ทำอะไร (30 วินาที) | auto-run ด้วยวันที่ที่ highlight (default = today) |
| คลิกวันที่อื่นใน calendar | วันที่เปลี่ยน — timer **ไม่ reset** |
| คลิก `Generate & Send` | รันด้วยวันที่เลือกตอนนั้น (timer หยุด) |
| คลิก `Cancel` / ปิด window | ออกจาก bat ไม่รัน |

### Date selector → daily_report.py

Picker เรียก: `subprocess.run([python, 'daily_report.py', '--fetch-odoo', '--date', 'YYYY-MM-DD'])`

`--date` ใน daily_report.py ทำให้:
- PDF folder: `Daily PDFs/{YYYY}/{MMM}/{DD}/` (overwrite ถ้ามี)
- PDF filenames: `*_DD-MMM-YYYY.pdf`
- Excel G1/G2/G3 cells = day/month/year
- Email subject: `Sale Report D MMM YYYY`
- Failure body: `... while processing the report for D MMM YYYY at phase: ...`

**Backward compat:** `daily_report.py --fetch-odoo` (no `--date`) ใช้ today เหมือนเดิม

### Fallback ใน bat
```bat
%PYTHON% -c "import tkcalendar" >nul 2>&1
if %ERRORLEVEL% == 0 (
    %PYTHON% run_for_date.py
) else (
    echo WARNING: tkcalendar not available — running for today's date
    %PYTHON% daily_report.py --fetch-odoo
)
```

ถ้า tkcalendar install ไม่สำเร็จ → fallback เป็น today อัตโนมัติ ไม่ break

---

## 6.1 Scheduled Task — v3.10 (6 AM Bangkok daily)

ใช้ Windows Task Scheduler รัน pipeline อัตโนมัติทุกวัน 6:00 น. เวลาประเทศไทย
เพื่อ generate รายงานของ "เมื่อวาน" (per v3.9 default)

### Setup (admin, ครั้งเดียว)

```powershell
# เปิด PowerShell AS ADMINISTRATOR แล้วรัน:
cd "C:\Users\USER\Desktop\Claude Cowork-Workspace\02-Projects\01-Sale Report"
.\register_scheduled_task.ps1
```

Script จะสร้าง task ชื่อ **"Sale Report - Daily 6 AM"** ใน Task Scheduler

### Settings ที่ตั้งไว้

| Setting | Value | เหตุผล |
|---------|-------|--------|
| Trigger | Daily at 06:00 local | ตามที่ user ต้องการ |
| WakeToRun | ✓ | Laptop ปิดฝา/sleep → ปลุกเครื่องตอน 6 AM |
| StartWhenAvailable | ✓ | ถ้าพลาด (เช่นเครื่องปิด) → รัน ASAP ตอนเปิดถัดมา |
| AllowStartIfOnBatteries | ✓ | Laptop ใช้แบตได้ |
| DontStopIfGoingOnBatteries | ✓ | กลางทางถอด AC ก็ไม่หยุด |
| ExecutionTimeLimit | 1 hour | safety timeout |
| MultipleInstances | IgnoreNew | ถ้า task เก่ายังรันอยู่ → ข้ามรอบใหม่ |
| LogonType | Interactive | user ต้อง login (จำเป็นสำหรับ Excel COM) |
| RunLevel | Limited | ไม่ใช้ admin privileges (Excel COM ทำงานปกติ) |

### Flow ของ scheduled run

```
6:00 AM → Task Scheduler ปลุกเครื่อง
       ↓ รัน run_scheduled.bat
       ↓
[1] เปิด log file: _Archive/scheduled_logs/2026-06-10_060000.log
[2] หา Python + ensure pywin32 / requests / dotenv installed
[3] taskkill Excel ที่ค้าง
[4] รัน daily_report.py --fetch-odoo (NO --date → ใช้เมื่อวาน per v3.9)
       ↓
       [pipeline]: Odoo fetch → cleanse → patch xlsm → PDFs → email
       ↓
[5] บันทึก exit code + finish timestamp ใน log
       ↓
ออกจาก bat (ไม่มี pause — สำคัญสำหรับ scheduled task)
```

### ⚠️ ข้อจำกัด / ข้อควรรู้

1. **User ต้อง logged in** — task ใช้ `LogonType=Interactive` เพื่อให้ Excel COM ทำงาน ถ้า logged out → task จะ pending
2. **Excel windows จะแสดงตอนรัน** — เพราะ `xl.Visible=True` ใน code (ข้อจำกัดของ COM) ถ้าเครื่อง lock screen — Excel เปิดอยู่ใต้ lock จะเห็นตอนปลดล็อค
3. **Wake from sleep** — ต้องเช็ค power settings ของ Windows ให้อนุญาต wake timers
4. **Closed-lid laptop:** Power Settings → Battery → "When I close the lid" — ถ้าตั้ง "Sleep" → wake ได้, ถ้าตั้ง "Hibernate" → wake ลำบาก, ถ้า "Shut down" → task ไม่รัน

### วิธีทดสอบ / ลบ

```powershell
# Test manual (admin not needed):
Task Scheduler UI → Right-click "Sale Report - Daily 6 AM" → Run
ดู log ที่: _Archive\scheduled_logs\<timestamp>.log

# Remove:
.\unregister_scheduled_task.ps1   # ต้อง run as admin
```

### Logs

ทุก scheduled run เขียน log file:
- Path: `_Archive\scheduled_logs\YYYY-MM-DD_HHMMSS.log`
- Contains: stdout + stderr ของทั้ง bat + Python pipeline
- ถ้า pipeline fail → log จะมี traceback + failure email ก็ส่งไป recipient ตามปกติ (v3.6)

### Manual run vs Scheduled run

| | Manual (run_fetch_odoo.bat) | Scheduled (run_scheduled.bat) |
|---|---|---|
| Date picker | ✓ มี (default=เมื่อวาน) | ✗ ไม่มี |
| Default date | เมื่อวาน (v3.9) | เมื่อวาน (v3.9) |
| Excel kill ก่อน | ✓ | ✓ |
| Auto-install deps | ✓ | ✓ |
| Log file | ✗ (output ที่ console) | ✓ (`_Archive\scheduled_logs\`) |
| Pause at end | ✓ (รอ user กด key) | ✗ (จบทันที — สำหรับ scheduled) |

ทั้งสองวิธีใช้ pipeline + email + folder structure เดียวกัน — ผลลัพธ์เหมือนกันทุกประการ

---

## 6.2 ⚠️ เงื่อนไขสำหรับ Scheduled Task — ทำ / ห้ามทำ

หลัง register task แล้ว เพื่อให้ task รันได้สม่ำเสมอทุก 6 AM **เครื่อง laptop ต้อง:**

### ✅ ต้องทำ / ต้องเป็นแบบนี้

| รายการ | เหตุผล |
|--------|--------|
| **ไม่ shut down เครื่องตอนกลางคืน** | Task ปลุกจาก Sleep ได้ แต่ปลุกจาก Power-off ไม่ได้ |
| **ปิดฝา laptop = Sleep** (ไม่ใช่ Hibernate / Shut down) | Wake timers ทำงานได้เฉพาะใน Sleep mode |
| **Stay logged in Windows** (อย่า Sign out) | Task settings: "Run only when user is logged on" — ถ้า sign out → task pending |
| **เปิด Allow wake timers** ใน Power Options | Windows ต้องอนุญาต task ปลุกเครื่อง |
| **เชื่อมต่อ Internet ตอน 6 AM** (Wi-Fi/Ethernet) | Pipeline ต้อง fetch Odoo + ส่ง email via Graph API |
| **Microsoft Excel ติดตั้ง + license ใช้งานได้** | PDF generation ใช้ Excel COM |
| **เก็บไฟล์ project ไว้ที่เดิม** (`C:\Users\USER\Desktop\Claude Cowork-Workspace\02-Projects\01-Sale Report\`) | Task ใช้ absolute path — ย้ายไฟล์ → ต้อง re-register |

### ❌ ห้ามทำ

| รายการ | ผลกระทบ |
|--------|---------|
| **ห้าม Shut Down เครื่อง** ตอนกลางคืน | Task พลาดรอบ → รัน ASAP ตอนเปิดถัดมา (อาจล่าช้าหลายชั่วโมง) |
| **ห้าม Sign out Windows** | Task ไม่รัน (รออยู่จนกว่าจะ sign in ใหม่) |
| **ห้ามลบ** `daily_report.py`, `.env`, `Sales Report V3.1.xlsm.bak`, `run_scheduled.bat` | Pipeline จะ fail (ไฟล์หาย) |
| **ห้ามแก้** code ใน `daily_report.py` / `run_scheduled.bat` ที่ไม่เข้าใจ | อาจ break pipeline |
| ✅ **แก้ `.env` ได้** เพื่อเปลี่ยน `EMAIL_TO` / `EMAIL_CC` / `EMAIL_BCC` | ⭐ ไม่ต้อง re-register task, ไม่ต้อง restart — มีผลตั้งแต่ run ครั้งถัดไป |
| **ห้ามเปลี่ยน timezone** ของเครื่อง | Task ใช้ local time — เปลี่ยน TZ → 6:00 ในเขตเวลาใหม่ |
| **ห้าม disable Task Scheduler service** ใน Windows Services | Task ทั้งหมดของระบบจะหยุด |
| **ห้ามให้ Antivirus / Firewall block** Python/Excel | Pipeline crash |

### 🟡 ทำได้ ไม่กระทบ

| รายการ |
|--------|
| Lock screen (Win + L) — task ยังรันได้ ตราบใดที่ login session ยังอยู่ |
| ปิดฝา laptop ค้างคืน (ถ้าตั้ง Lid action = Sleep) |
| ขาด Wi-Fi ระหว่างวัน — กลับมาเชื่อมต่อก่อน 6 AM |
| Reboot ตอนกลางวัน — แค่ login กลับเข้ามา task ยังอยู่ |
| ติดตั้ง Windows Update — เครื่อง reboot → login กลับมา → task ยังอยู่ |
| **แก้ `.env`** (เปลี่ยน EMAIL_TO/CC/BCC ใครก็ได้) — ผลใช้ตั้งแต่ run ครั้งถัดไป ไม่ต้อง re-register |

### 📋 ตรวจสอบ Power Settings ของ Windows (ครั้งเดียวพอ)

เปิด **Settings → System → Power & battery → Advanced power settings**

| Setting | Value ที่ต้องตั้ง |
|---------|--------------------|
| **Sleep → Allow wake timers** | Enable (ทั้ง "On battery" และ "Plugged in") |
| **Power buttons and lid → Lid close action** | Sleep (ไม่ใช่ Hibernate / Shut down) |
| **Sleep → Hybrid sleep** | Off (ปกติแล้ว Off — Hybrid sleep บางครั้งทำให้ wake ไม่ทำงาน) |

### 🔍 ถ้า task ไม่รันตอน 6 AM — ตรวจอะไรบ้าง

1. เปิด **Task Scheduler** → ดู task **"Sale Report - Daily 6 AM"** ใน Task Scheduler Library
2. ดูคอลัมน์ **"Last Run Result"**:
   - `0x0` (The operation completed successfully) — รันสำเร็จ ✓
   - `0x00041303` (Task has not yet run) — ยังไม่ถึงรอบ
   - `0x41306` (Task was terminated by user) — โดน user หยุด
   - อื่นๆ → ดู log file: `_Archive\scheduled_logs\latest.log`
3. ดูคอลัมน์ **"Last Run Time"** vs **"Next Run Time"** — มี gap ไหม?
4. เปิด **`_Archive\scheduled_logs\latest.log`** เพื่อดูรายละเอียดการรันครั้งล่าสุด

---

## 7. Field Mappings — Odoo API Fields (ไม่เปลี่ยน)

ดูใน `daily_report.py` ตัว `SO_FIELDS`, `POS_FIELDS`, `CN_FIELDS`, `SOPAY_FIELDS`, `POSPAY_FIELDS` — เนื้อหา 27/26/26/11/11 fields ตามลำดับ ยืนยันแล้วจาก Odoo export

**Highlights:**
- SO: `x_studio_actual_order_date`, `partner_id/title|name`, `order_line/*`
- POS: `date_order` (UTC → แปลงด้วย `pos_utc_range()`), `config_id/name`
- CN: `invoice_line_ids/price_total` ← ตัวที่เคย leak (v3.1 fix)
- SOPay: `is_pay_with_inv=True`, filter ด้วย `AP` set
- POSPay: filter ด้วย `AP2` set (รวม Void Credit Card)

---

## 8. Version History — รายละเอียดทุก iteration

### v1.0 — baseline (May 27, 2026)
- โค้ดต้นฉบับ มาจาก inject/cleanse/generate แยกกัน
- Sequential Odoo fetch (~12s)

### v2.0 — parallelized fetch + optimized cleanse (May 27)
- `ThreadPoolExecutor(max_workers=5)` fetch 5 datasets concurrent
- Hoisted attribute lookups, cached lengths, slice-extend padding
- Speedup: ~50% on fetch phase

### v3.0 — HTTP keep-alive + tighter cleanse hot paths (Jun 1)
- `KeepAliveTransport(SafeTransport)` reuses TCP/TLS per thread
- `_NUM_TYPES = (int, float)`, `type(x) is int` faster than `isinstance(x, (int, float))`
- `_COMBINE_POS_RE` pre-compiled regex
- `_parse_date_triple()` fused get_date + to_iso_date + ordinal
- Schwartzian sort with stable insertion order

### v3.1 — Cross-document leak fix (Jun 1)
**Symptom:** RINV2605.0004 (6-May) Total col = 351,180.42 (จริงต้อง 46,830)
**Cause:** zero-qty filter carries `price_total` to `d[i+1]` without same-doc check
**Fix:** add `if cur[1] == nxt[1]:` guard in `cleanse_so`, `cleanse_pos`, `cleanse_cn` (3 จุด)

### v3.2 — Odoo fetch transport optimizations (M1) (Jun 1)
- gzip Accept-Encoding (XML payload ลด ~80-90%)
- TLS warm-up via 1 throwaway `common.version()` call
- Expected: 25-35% wall-clock reduction (production network)

### v3.3 — xlsm patching optimizations (M4) (Jun 2)
- Single-pass `.bak` read (was open twice)
- `patch_xml` operates on bytes end-to-end (no decode/encode)
- Module-level pre-compiled regexes
- `col_letter` memoised, `make_cell` f-strings, ZIP_DEFLATED compresslevel=3
- Speedup: ~10% (771ms → 691ms in mock)

### v3.4 — PDF generation optimizations (M5) (Jun 2)
- `xl.ScreenUpdating = False` during loop
- `xl.Calculation = xlCalculationManual`
- PageSetup hoisted out of loop (only PrintArea per-iter)
- `Range('G1:G3').Value = (...)` bulk assign (1 COM call vs 3)
- **Bonus bug fix:** `_wait_excel_ready` now actually polls `xl.Ready` (was one-shot)
- Expected: 35-55% PDF phase reduction

### v3.5b — Cleanse hot-path optimizations (M2) (Jun 2)
- Fused inline ensure() + drop redundant guards
- `cleanse_so` 1.51× faster, `cleanse_pos` 1.21×, `cleanse_cn` 1.19×
- **Cherry-pick:** kept v3.4 `cleanse_pospay` (v3.5 regressed it ~4%)
- Aggregate: 1.27× (~21% reduction)

### v3.6 — Graph API email integration (Jun 4)
- OAuth 2.0 client credentials flow → access_token (cached)
- `send_daily_email(pdf_paths, missing_branches=None)` — POST `/sendMail`
- `send_failure_email(error, phase, traceback)` — best-effort notification
- `generate_daily_pdfs` returns list of (branch, pdf_path) tuples
- Top-level try/except in `if __name__` block

### v3.7 — Date selector + calendar UI (Jun 4)
- `--date YYYY-MM-DD` argparse (backward compat — no arg = today)
- 3 functions รับ `report_date=None` kwarg
- Failure email body mentions chosen date
- New `run_for_date.py` — tkinter + tkcalendar + 30s countdown
- `run_fetch_odoo.bat` auto-invokes date picker (with fallback to today)
- **Hotfix:** cross-platform date format (was using Linux-only `%-d`)

### v3.10 — Scheduled task setup (Jun 10)
- New `run_scheduled.bat` — auto-runs pipeline without date picker, logs to file
- New `register_scheduled_task.ps1` — one-time admin script to create Windows scheduled task
- New `unregister_scheduled_task.ps1` — cleanup
- Task settings: Daily 06:00, WakeToRun, StartWhenAvailable, Interactive logon (for Excel COM)
- Logs saved per-run to `_Archive\scheduled_logs\YYYY-MM-DD_HHMMSS.log`
- Manual run via `run_fetch_odoo.bat` unchanged (still has date picker)

### v3.12 — Graph secret expiry warning (Jun 10)
- Optional `GRAPH_SECRET_EXPIRES=YYYY-MM-DD` in `.env`
- Helper `_check_secret_expiry()` returns (days, severity, message)
- Warning prepended to email body (both success + failure)
- Warning logged at startup so it shows in scheduled task log too
- Severity: > 30d none | 8-30d warning | 0-7d critical | expired critical
- Renewal workflow: admin makes new secret → user updates both env values → done

### v3.11 — Multi-recipient + CC + BCC (Jun 10)
- `EMAIL_TO`, `EMAIL_CC`, `EMAIL_BCC` accept comma-separated email lists
- Both daily success email AND failure notification send to To+CC+BCC
- Backward compat: single email in `EMAIL_TO` + missing CC/BCC = legacy behavior
- New helper `_parse_recipients(s)` converts CSV → Graph API recipient list
- Log line shows To + CC + BCC clearly so you know who got the email
- Editing `.env` takes effect on the NEXT run — no code restart needed

### v3.10.1 — UTF-8 hotfix for scheduled task (Jun 10)
- Python 3.x on Windows defaults to cp1252 when stdout is redirected (Task Scheduler context)
- Unicode chars (→, ⚠, ✓, —) in pipeline output crashed with UnicodeEncodeError
- Fix: set `PYTHONIOENCODING=utf-8` + `PYTHONUTF8=1` in both run_scheduled.bat and run_fetch_odoo.bat

### v3.9 — Default to yesterday (Jun 10)
- `daily_report.py`: when `--date` is NOT provided → default = YESTERDAY (was today)
- `run_for_date.py`: calendar dialog highlights yesterday by default; maxdate still = today
- Rationale: sales reports are typically run in the MORNING for the previous day's
  (finalized) sales — yesterday is the right default
- Backward compat: explicit `--date YYYY-MM-DD` still wins

### v3.8 — Fix cleanse_so multi-line fill + xlsm rename (Jun 5)
**Bug:** v3.5b ตอน inline cleanse_so ไว้เก็บ keys_match comparison ก่อน fill col 1 (SO name) — ตั้งใจ preserve v1.0 semantics ที่จริงๆ เป็น bug. ผลคือ multi-line SOs (711 ตัว = 48.8% ของ data, 1,655 rows = 53% ของ SO sheet) มี campaign / source (branch) / foreigner / customer info / salesperson / payment_term / **rate** ว่างทั้งหมดบนบรรทัด 2+

**Fix:** ย้าย `keys_match = (cur[1] == prev[1])` ไป AFTER ที่ fill col 0,1 จาก prev row (5 บรรทัด edit)

**Impact:**
- PDF Daily Template filter by branch → ครบทุกบรรทัด (ก่อนนี้ตกหล่น lines 2+)
- YTD pivot ตาม salesperson / campaign / branch → ถูกต้อง
- Non-THB multi-line SOs (เช่น SO2604.0071 USD) → rate fill ครบ → THB conversion ถูก

**Xlsm rename (admin change):**
- `Retail Sale Report V3.0 2026.xlsm` → `Sales Report V3.1.xlsm`
- ย้ายไฟล์ V3.0 2026 เข้า `_Archive/02-old_excel_versions/`
- สร้าง `Sales Report V3.1.xlsm.bak` (template) จาก V3.1.xlsm
- Update path references ใน daily_report.py + INSTRUCTIONS.md + PROJECT_SUMMARY.md

---

## 9. Optimization Roadmap — Final Status

| Milestone | สถานะ | Wall-clock impact |
|-----------|--------|--------------------|
| **M1** Odoo fetch (gzip + keep-alive + warm-up) | ✅ v3.2 | ~25-35% (production) |
| **M2** Cleanse hot paths | ✅ v3.5b | ~21% (synthetic) |
| **M3** Build YTD/CR rows | — skipped (already optimal from v3.0) | — |
| **M4** Patch xlsm (bytes-mode + compresslevel) | ✅ v3.3 | ~10% |
| **M5** PDF generation (ScreenUpdate off + hoist PageSetup) | ✅ v3.4 | ~35-55% (estimated) |
| **M6** Email integration (Graph API) | ✅ v3.6 | feature add |
| **M7** Date selector (CLI + calendar UI) | ✅ v3.7 | feature add |
| **Multi-line SO fill fix** | ✅ **v3.8** | **correctness fix — affects 711 SOs / 1,655 rows** |
| xlsm rename to V3.1 | ✅ v3.8 (admin) | template renamed Sales Report V3.1 |
| Default date = yesterday | ✅ v3.9 | morning reporting workflow |
| Scheduled task (6 AM daily) | ✅ v3.10 | Task Scheduler + log files |
| Multi-recipient + CC + BCC | ✅ v3.11 | `.env`-driven — no code restart |
| Graph secret expiry warning | ✅ v3.12 | optional GRAPH_SECRET_EXPIRES in `.env` |

---

## 10. Excel XML Patching — Technical Details (ไม่เปลี่ยน)

**Approach:** ใช้ `zipfile` อ่าน/เขียน xlsm โดยตรง ไม่ใช้ openpyxl

**patch_xml(orig_bytes, new_rows, first_row):**
- เก็บ XML header rows 1 ถึง first_row-1 ไว้เดิม (formulas, formatting)
- แทนที่ rows ≥ first_row ด้วย data ใหม่
- Raw data sheets ใช้ first_row=1 (เขียนทับทั้งหมด)
- YTD/CR ใช้ first_row=5 (เก็บหัว rows 1-4)
- v3.3: operates on bytes end-to-end (no decode/encode round-trip)

**strip_formula_cache(xml_bytes):**
- ลบ `<v>cache</v>` ออกจาก formula cells
- ทำให้ Excel recalculate SUBTOTAL ใหม่จากข้อมูลจริงเมื่อเปิดไฟล์
- Apply เฉพาะ YTD และ CR sheets

**calcChain.xml:** ถูกลบออกจาก zip output ทุกครั้ง (Excel สร้างใหม่เอง)

**fullCalcOnLoad:** ใส่ `fullCalcOnLoad="1"` ใน `xl/workbook.xml` → `<calcPr>` บังคับ Excel recalculate ตาม dependency order

---

## 11. PDF Generation — Daily Template Sheet

**Sheet:** "Daily Template" ใน xlsm

**Cells ที่ต้องตั้งค่า:**
- G1 = วัน (day) — เลขไม่มี leading zero
- G2 = เดือน (month) — เลข
- G3 = ปี (year)
- J2 = ชื่อสาขา (dropdown)

**J2 Dropdown options:** Somkid, NORSE Store, Line Chat, Line My Shop, Lazada, HAY Store, Wholesale, All
- Loop ทุกสาขา **ยกเว้น All**

**v3.4 optimizations:**
- `xl.ScreenUpdating = False` + restore in finally
- `xl.Calculation = xlCalculationManual` + restore
- PageSetup hoisted: PrintTitleRows, Zoom, FitToPagesWide, FitToPagesTall, all margins ตั้งครั้งเดียวก่อน loop
- Per-iteration: เซตแค่ `PrintArea`
- บลก G1/G2/G3 ด้วย `Range('G1:G3').Value = ((d,),(m,),(y,))` — 1 COM call

**ชื่อไฟล์:** `{branch}_{DD-MMM-YYYY}.pdf` (เช่น `HAY_Store_18-May-2026.pdf`)

**Output folder (v3.1+):** `Daily PDFs\{YYYY}\{MMM}\{DD}\` (organized by date)

---

## 12. win32com — ข้อควรระวัง

| ปัญหา | สาเหตุ | วิธีแก้ |
|-------|--------|---------|
| `'bool' object is not callable` | `xl.Visible = False` block COM calls | ใช้ `xl.Visible = True` เสมอ |
| Keyword arguments ไม่ work | win32com late binding ไม่รองรับ kwargs | positional args เท่านั้น |
| `Call was rejected by callee` | Excel ยัง recalculate อยู่ | ใช้ `_wait_excel_ready(xl)` รอ `xl.Ready` |
| `Property can not be set` | Sheet มี protection | เรียก `ws.Unprotect()` ก่อน |
| Lock file error | Excel ค้างจาก run ก่อนหน้า | bat file มี `taskkill /F /IM EXCEL.EXE` อัตโนมัติ |
| `FitToPagesTall = 0` error | บาง Excel version ไม่รองรับ 0 | ใช้ 32767 แทน |
| `%-d` strftime crash | Linux-only directive | ใช้ `f"{d.day}"` แทน (v3.7 hotfix) |

---

## 13. Date Format

- **YTD/CR date column:** `DD-MMM-YYYY` string เช่น `15-Jan-2026`
- ฟังก์ชัน `_parse_date_triple(dv)` (v3.0) แปลง Excel serial หรือ ISO string → tuple
- Sort key ใช้ ordinal pre-computed (Schwartzian transform)

**POS UTC correction:**
- Odoo เก็บ date_order เป็น UTC
- ต้องแปลง TH timezone (UTC+7) → UTC ก่อน query
- `pos_utc_range()` คำนวณ range ที่ถูกต้อง

---

## 14. ผลการตรวจสอบความถูกต้อง

เปรียบเทียบ Python output vs VBA reference file (column AI = Total_THB):

| Metric | ผล |
|--------|-----|
| จำนวน rows | เท่ากัน |
| Sort order | ต่างกัน (Python sort ต่างจาก VBA) — **ค่าเหมือนกัน** |
| Sum difference | -2,505 บาท |
| สาเหตุ difference | SO2604.0070 และ SO2604.0071 ใน REF มีค่า 0 แต่ Python มีค่าจริง |
| ข้อสรุป | Python ให้ค่า accurate กว่า REF |

**v3.1 cross-doc leak fix verification:**
- 52/52 synthetic edge cases PASS
- Byte-equiv with baseline on normal data
- RINV2605.0004 case: 351,180.42 → 46,830 ✓ FIXED

**v3.5b cleanse correctness:**
- 5/5 functions byte-equiv with v3.4 baseline on ~47K synthetic rows
- All edge cases (Thai/Non-Thai, zero-qty boundary, fill-down chains, AP/AP2, Void, Combine, unicode, short rows)

**M1-M5 verifier results:** ทุก optimization milestone ผ่าน byte-equiv check ก่อน apply

**v3.8 cleanse_so multi-line fix verification:**
- Multi-line empty header → all cols fill correctly (campaign, source, foreigner, customer info, salesperson, payment_term, rate) ✓
- Full-data SO → byte-equiv with v3.7 baseline (no regression on already-populated rows) ✓
- cleanse_pos / cleanse_cn / cleanse_sopay / cleanse_pospay → untouched, byte-equiv ✓
- v3.1 leak fix preserved (SO-A → SO-B boundary, no carry across docs) ✓
- Non-THB multi-line rate now fills correctly (Sale_Closed_THB conversion accurate) ✓

---

## 15. PDF Verification Results (27-May-2026)

ตรวจ PDF 21 ไฟล์ใน `Daily PDFs\` (14/18/22-May × 7 สาขา):

| Aspect | Result |
|--------|--------|
| Branch dropdown (J2) | ✅ ทำงานถูก — แต่ละสาขาแสดงชื่อ + data ของตัวเอง |
| Date cells (G1/G2/G3) | ✅ ถูกต้อง (Day/Month/Year) |
| Print Area workaround | ✅ ทำงาน (max 41 / last col I) |
| Page size | ⚠️ US Letter Landscape 792x612pt (ไม่ใช่ A4) |
| Font / Layout | ✅ อ่านได้ ไม่หด, ภาษาไทยถูกต้อง |
| Multi-page | ✅ 14-May HAY Store 2 pages (112 รายการ) — FitToPagesTall=32767 ทำงาน |
| Branch Summary table (A-B) | ❌ Values ว่างเปล่าทั้งตาราง — อาจเป็น formula issue |

---

## 16. Pending / Future Work

- [x] ~~ส่งอีเมล~~ — ✅ done v3.6 (Microsoft Graph API)
- [x] ~~ตรวจสอบ PDF margin/font~~ — ✅ verified 27-May (margin/font/fit OK, Thai OK)
- [x] ~~Date picker UI~~ — ✅ done v3.7 (tkinter + tkcalendar + 30s countdown)
- [x] ~~Multi-line SO header fields empty on lines 2+~~ — ✅ done v3.8 (cleanse_so fix)
- [x] ~~Rate column missing on multi-line non-THB SOs~~ — ✅ done v3.8 (same fix)
- [x] ~~xlsm template rename to Sales Report V3.1~~ — ✅ done 5-Jun-2026
- [ ] **PrintTitleRows I5:V5 only** — Excel ไม่รองรับ column-restricted title rows, ตอนนี้ repeat ทั้ง row 5
- [ ] **Branch Summary ด้านซ้าย (col A-B) ว่างเปล่า** — formula ใน Daily Template อาจอ้างอิงผิด (ตาราง 7 สาขา × 3 รายการ Sales/Cash/AR ไม่มี value แม้ใน 14-May ที่มี data จริง) — รอ user confirm ว่าเป็น design intent หรือ bug
- [ ] **Page size = US Letter Landscape** — น่าตรวจสอบว่า user ต้องการ A4 Landscape (842x595pt) ไหม
- [ ] **Header row "C, G, H, I, J, ..."** (Excel column letters) — แสดงใน PDF ทุกหน้า — รอ user confirm ว่าเป็น design intent หรือควรซ่อน
- [ ] **Empty-day PDF: พื้นที่ว่างเยอะ** — `last_row = max(41, last_col_I)` ทำให้วันที่ไม่มี transaction ครอบพื้นที่ว่าง ~3/4 หน้า
- [x] ~~CC/BCC support~~ — ✅ done v3.11
- [x] ~~Multiple recipients~~ — ✅ done v3.11 (comma-separated in EMAIL_TO/CC/BCC)
- [x] ~~Schedule daily auto-run~~ — ✅ done v3.10 (Task Scheduler + register_scheduled_task.ps1)
- [ ] **Optional: Fetch only chosen date's incremental data** — ตอนนี้ fetch ทั้งปีทุก run (ช้าเกินจำเป็นถ้า regenerate ย้อนหลังเฉพาะ 1 วัน)

---

## 17. Rollback Path — ถ้าเจอปัญหา

ทุก version อยู่ใน `_Archive/01-daily_report_backups/`:

```cmd
cd "C:\Users\USER\Desktop\Claude Cowork-Workspace\02-Projects\01-Sale Report"

REM กลับ v3.6 (ก่อน date picker)
copy "_Archive\01-daily_report_backups\daily_report_v3.6_backup.py" daily_report.py

REM กลับ v3.5b (ก่อน email — กลับเป็น offline pipeline)
copy "_Archive\01-daily_report_backups\daily_report_v3.5b_backup.py" daily_report.py

REM กลับลึกกว่านี้: v3.4, v3.3, v3.2, v3.1, v3.0, v2_stable, v1.0
```

---

## 18. การรันซ้ำในวันเดียวกัน

- xlsm → overwrite ทับทันที
- PDF ชื่อเดิม → overwrite ทับทันที (ชื่อไฟล์ใช้วันที่ที่เลือก)
- วันถัดไป → PDF folder ใหม่ตามวันใหม่ ไฟล์เก่ายังอยู่ใน year/month/day structure
- `.bak` file → ไม่ถูกแตะต้องเลย (เป็น base template เสมอ)
- Email → ส่งใหม่ทุกครั้ง (recipient จะได้รับซ้ำในวันเดียวกัน)

---

## 19. วิธีเพิ่มปีใหม่

1. เปลี่ยนค่า `ODOO_YEAR` ในไฟล์ `daily_report.py`:
   ```python
   ODOO_YEAR = 2027
   ```
2. สร้าง/เปลี่ยน xlsm template ใหม่ให้ตรงกับปีนั้น
3. ทำ `.bak` ใหม่ (`copy "Sales Report V3.2.xlsm" "Sales Report V3.2.xlsm.bak"`)
4. ทดสอบรัน `run_fetch_odoo.bat`

---

## 20. ฟังก์ชันหลักและหน้าที่ (v3.8)

| ฟังก์ชัน | หน้าที่ |
|---------|---------|
| `fetch_odoo_all(year)` | ดึงข้อมูล 5 ชุดจาก Odoo XML-RPC (parallel + gzip + keep-alive) |
| `_KeepAliveSafeTransport` | HTTPS transport ที่ reuse TCP/TLS connection (v3.2) |
| `cleanse_so(raw)` | Cleanse SO data + cross-doc leak fix (v3.1) + multi-line header fill (v3.8) |
| `cleanse_pos/cn/sopay/pospay(raw)` | Cleanse data + cross-doc leak fix (v3.1) |
| `split_column_p(data)` | แยก Category → 4 cols |
| `build_ytd_rows(so_c, pos_c, cn_c)` | สร้าง rows สำหรับ YTD sheet (Schwartzian sort) |
| `build_cr_rows(sop_c, pop_c)` | สร้าง rows สำหรับ CR sheet |
| `patch_xml(orig, rows, first_row)` | Patch sheet XML ใน xlsm (bytes-mode v3.3) |
| `strip_formula_cache(xml_bytes)` | ลบ cached value จาก formula cells |
| `generate_daily_pdfs(xlsm_path, report_date=None)` | Export PDF 7 สาขา (v3.4 optimized, v3.7 dateable) |
| `_wait_excel_ready(xl)` | รอ Excel ก่อนส่ง COM command (v3.4 fixed bug) |
| `_get_access_token()` | OAuth 2.0 → Graph API access_token (cached) |
| `send_daily_email(pdf_paths, missing_branches, report_date)` | ส่งอีเมลพร้อม attachment ผ่าน Graph API |
| `send_failure_email(error, phase, traceback, report_date)` | แจ้งเตือนเมื่อ pipeline ล้ม |
| `_load_env()` | โหลด `.env` (python-dotenv fallback to stdlib) |
| `_parse_date_triple(dv)` | Fused date parser (v3.0) |
| `pos_utc_range(year)` | คำนวณ UTC range จาก TH timezone |

---

## 21. Quick Reference — สิ่งที่ต้องจำ

| สิ่งที่ทำ | คำสั่ง / ขั้นตอน |
|-----------|-------------------|
| รันรายงานวันนี้ | ดับเบิลคลิก `run_fetch_odoo.bat` → รอ 30 วินาที (หรือคลิก Send) |
| รันรายงานวันที่ผ่านมา | ดับเบิลคลิก `run_fetch_odoo.bat` → เลือกวัน → คลิก Generate & Send |
| Rollback version | `copy _Archive\01-daily_report_backups\daily_report_vX_backup.py daily_report.py` |
| ดู PDF ของวันใดวันหนึ่ง | เปิด `Daily PDFs\{YYYY}\{MMM}\{DD}\` |
| แก้ credentials | edit `.env` |
| แก้ recipient | edit `.env` → `EMAIL_TO` |
| ขึ้นปีใหม่ | แก้ `ODOO_YEAR` + xlsm template + .bak |

---

*สร้างโดย Claude (Anthropic) — HAY/Norse Sale Report Session*
*Last major update: 10 Jun 2026 (v3.12 — Graph client secret expiry warning)*
