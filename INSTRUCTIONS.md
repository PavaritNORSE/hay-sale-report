# คู่มือ HAY Daily Sale Report
## `daily_report.py` — เวอร์ชันปรับปรุง

---

## ภาพรวม

`daily_report.py` อ่านไฟล์ export จาก Odoo **5 ไฟล์** แล้วอัพเดท `Sales Report V3.1.xlsm`  
ในขั้นตอนเดียว ใช้เวลารวมประมาณ **4 วินาที**

แทนที่กระบวนการเดิม 46 ขั้นตอน (export + copy-paste + VBA macro) ให้เหลือแค่ 3 ขั้นตอน

---

## ขั้นตอนการใช้งาน

### 1. Export ข้อมูลจาก Odoo

Export ไฟล์ xlsx ทั้ง 5 ไฟล์ตามลำดับนี้ — **ลำดับสำคัญสำหรับ Payment**

#### SO — Sales Order
1. เข้า **Sales** module
2. กด **Favorites** → เลือก **Sale Report YYYY** (เช่น Sale Report 2026)
3. เลือก **All transactions**
4. กด **Action → Export** → เลือก template **Sale Report V.3** → **Export**
5. ไฟล์ที่ได้: `Sales Order (sale.order) *.xlsx`

#### POS — Point of Sale Orders
1. เข้า **Point of Sale** module → **Orders → Orders**
2. กด **Favorites** → เลือก **Sale Report YYYY**
3. เลือก **All transactions**
4. กด **Action → Export** → เลือก template **Sale Report V.3** → **Export**
5. ไฟล์ที่ได้: `Point of Sale Orders (pos.order) *.xlsx`

#### CN — Credit Notes
1. เข้า **Accounting** module → **Customers → Credit Notes**
2. กด **Favorites** → เลือก **Sale Report YYYY**
3. เลือก **All transactions**
4. กด **Action → Export** → เลือก template **Sale Report V.3** → **Export**
5. ไฟล์ที่ได้: `Journal Entry (account.move) *.xlsx`  
   *(ชื่อไฟล์จะขึ้นต้นด้วย Journal Entry ไม่ใช่ Credit Note)*

#### SO Payment — Payments with Invoices
1. ยังอยู่ใน **Accounting** → **Customers → Payments with Invoices**
2. กด **Favorites** → เลือก **Sale Report YYYY**
3. เลือก **All transactions**
4. กด **Action → Export** → เลือก template **SO Payment Sale Report** → **Export**
5. ไฟล์ที่ได้: `Payments (account.payment) *.xlsx`

#### POS Payment — Payments *(export หลังสุด)*
1. ยังอยู่ใน **Accounting** → **Customers → Payments**
2. กด **Favorites** → เลือก **Sale Report YYYY**
3. เลือก **All transactions**
4. กด **Action → Export** → เลือก template **POS Payment Sale Report** → **Export**
5. ไฟล์ที่ได้: `Payments (account.payment) *.xlsx`

> **หมายเหตุ Payment:** SO Payment และ POS Payment มีชื่อไฟล์เหมือนกันทุกอย่าง  
> Script แยกโดยดูเวลา export — **SO Payment ต้อง export ก่อน POS Payment เสมอ**  
> ไฟล์ที่เวลาเก่ากว่า = SO Payment / ไฟล์ที่เวลาใหม่กว่า = POS Payment

> **Sale Report YYYY Favorites:** Filter นี้ถูกตั้งค่าใน Odoo ให้กรองตาม Sale Order, วันที่ภายในปีนั้น และเงื่อนไขอื่นๆ — กดเลือกแค่ชื่อนี้ก็พอ ปีต่อไปจะมี filter ใหม่ชื่อ Sale Report YYYY+1

---

### 2. ปิด Excel ก่อนทุกครั้ง

ต้องปิดไฟล์ `Sales Report V3.1.xlsm` ใน Excel ก่อนรัน script  
Excel จะล็อคไฟล์ไว้ และ script จะหยุดโดยอัตโนมัติถ้าตรวจพบว่า Excel เปิดอยู่

---

### 3. รัน Script

**วิธีที่ 1 — Double-click (แนะนำ):**
```
ดับเบิ้ลคลิกที่ run_daily_report.bat
```

**วิธีที่ 2 — Auto-detect (ไฟล์ล่าสุดใน Downloads):**
```bash
python daily_report.py
```

**วิธีที่ 3 — จำกัดอายุไฟล์ (แนะนำถ้ามีไฟล์เก่าค้างใน Downloads):**
```bash
python daily_report.py --max-age 4
```
> `--max-age 4` = ดึงเฉพาะไฟล์ที่ export มาภายใน 4 ชั่วโมงล่าสุด  
> ป้องกันการดึงไฟล์เก่าจาก session ก่อนหน้าโดยผิดพลาด

**วิธีที่ 4 — ระบุไฟล์เอง:**
```bash
python daily_report.py \
  --so    "path/to/Sales Order.xlsx" \
  --pos   "path/to/Point of Sale Orders.xlsx" \
  --cn    "path/to/Journal Entry.xlsx" \
  --sopay "path/to/Payments_SO.xlsx" \
  --pospay "path/to/Payments_POS.xlsx"
```

---

### 4. ตรวจสอบ Output

Script จะแสดงผลดังนี้ — ตรวจสอบ timestamp ของแต่ละไฟล์ว่าตรงกับที่ export มาจริงๆ:

```
Auto-detecting xlsx files (ล่าสุดในโฟลเดอร์): C:/Users/USER/Downloads
Files detected (0.0s):
  SO    : Sales Order (sale.order) - 2026-04-22T....xlsx  [2026-04-22 09:15:32]
  POS   : Point of Sale Orders (pos.order) - ...xlsx      [2026-04-22 09:18:44]
  CN    : Journal Entry (account.move) - ...xlsx          [2026-04-22 09:21:05]
  SOPay : Payments (account.payment) - ...xlsx            [2026-04-22 09:25:11]
  POSPay: Payments (account.payment) - ...xlsx            [2026-04-22 09:27:33]

Reading xlsx files...
  SO=1234 POS=2406 CN=89 SOPay=456 POSPay=789 rows  (1.2s)
Cleansing...
  SO=1100 POS=2300 CN=85 SOPay=400 POSPay=700 clean rows  (1.8s)
Building report rows...
  YTD=3485 rows × 36 cols  CR=1100 rows × 13 cols  (2.0s)
Patching YTD and CR sheets...
Writing xlsm...
Done! Total time: 4.1s
```

เปิด Excel แล้วรัน Macro ต่อได้เลย

---

## โครงสร้าง Sheet ใน xlsm

| Sheet | ชื่อ | คำอธิบาย |
|-------|------|-----------|
| sheet1 | CR (Cash Receipt) | Payment summary 13 คอลัมน์ |
| sheet4 | YTD | ข้อมูลยอดขายรวม SO+POS+CN, 36 คอลัมน์ (A–AJ) |
| sheet5 | SO Data | ข้อมูล Sales Order ที่ cleanse แล้ว |
| sheet6 | POS Data | ข้อมูล POS Orders ที่ cleanse แล้ว |
| sheet7 | CN Data | ข้อมูล Credit Notes ที่ cleanse แล้ว |
| sheet8 | SO Payment | ข้อมูล SO Payments ที่ cleanse แล้ว |
| sheet9 | POS Payment | ข้อมูล POS Payments ที่ cleanse แล้ว |

### YTD Column Layout (A–AJ, 36 คอลัมน์)

```
A  = Month          B  = Weekday        C  = Date
D  = Order No.      E  = Salesperson     F  = Nationality
G  = Customer       H  = Source          I  = Channel
J  = Brand          K  = Collection      L  = Payment Terms
M  = SKU            N  = Product Name    O  = Fabric
P  = Full Category  Q  = Product Cat.    R  = Sub-Cat 1
S  = Sub-Cat 2      T  = Color           U  = Size
V  = UOM            W  = Qty Ordered     X  = Qty Delivered
Y  = Unit Price     Z  = Discount %      AA = Sale Closed
AB = Total          AC = Unpaid          AD = Currency
AE = Exchange Rate
AG = Unit Price THB   AH = Sale Closed THB
AI = Total THB        AJ = Unpaid THB
```

> **หมายเหตุ:** POS และ CN ไม่มีคอลัมน์ AC (Unpaid) — ใส่เป็นค่าว่างไว้แทน  
> เพราะ POS ชำระเงินที่หน้าร้านทันที / CN คือรายการคืนเงิน ไม่มียอดค้างชำระ

### CR Column Layout (A–M, 13 คอลัมน์)

```
A  = Month     B  = Date      C  = Order/Payment No.
D  = (ว่าง)    E  = (ว่าง)    F  = Journal
G  = (ว่าง)    H  = Customer   I  = Payment Method
J  = Amount    K  = Currency   L  = Exchange Rate
M  = Amount THB
```

---

## Troubleshooting

### ❌ "Please close the report in Excel first"
Excel เปิดไฟล์อยู่ → ปิด Excel แล้วรันใหม่

### ❌ "ไม่พบไฟล์ xlsx"
ไฟล์ไม่อยู่ใน Downloads หรือไม่ตรง pattern  
ลอง: `python daily_report.py --max-age 24` เพื่อขยายช่วงเวลาที่ค้นหา

### ❌ Script ดึงไฟล์เก่าผิด session
เพิ่ม `--max-age 4` เพื่อดึงเฉพาะไฟล์ที่ export มาภายใน 4 ชั่วโมง  
หรือระบุ path ไฟล์ตรงๆ ด้วย `--so / --pos / ...`

### ⚠ "SO Payment และ POS Payment เป็นไฟล์เดียวกัน"
พบ Payment file แค่ 1 ไฟล์ใน Downloads — ตรวจสอบว่า export ครบทั้ง SO Payment และ POS Payment แล้ว

### ❌ YTD มีข้อมูลน้อยกว่าที่คาดไว้
ตรวจ POS export — ต้อง export ทั้งปี (ตั้งแต่ 1 ม.ค.) ไม่ใช่แค่วันนี้  
ดูตัวเลขหลัง read: `POS=2406` ถ้าต่ำผิดปกติ ให้ export ใหม่

### ❌ Column ไม่ตรงใน Excel หลัง script รัน
Odoo อาจอัพเดทและเปลี่ยน column layout → แจ้งเพื่อปรับ `cleanse_*` functions

---

## Payment Methods ที่ระบบรองรับ (CR Sheet)

SO Payment:
- Credit Card (SCB), Bank Transfer (SCB)
- Credit Card (KBANK), Bank Transfer (KBANK)
- Beam (KBANK), Alipay (KBANK), WeChat Pay (KBANK)

POS Payment (รวม Void):
- ทุกรายการข้างต้น + Void Credit Card (SCB/KBANK)

---

## ไฟล์ในโฟลเดอร์ Sale Report

| ไฟล์ | หน้าที่ |
|------|---------|
| `daily_report.py` | Script หลัก — รันตัวนี้อย่างเดียว |
| `run_daily_report.bat` | ดับเบิ้ลคลิกเพื่อรัน (Windows) |
| `Sales Report V3.1.xlsm` | ไฟล์รายงานที่จะถูกอัพเดท |
| `Sales Report V3.1.xlsm.bak` | ไฟล์ template สำหรับ header rows (ห้ามลบ) |
| `INSTRUCTIONS.md` | คู่มือนี้ |

> **.bak file:** Script อ่าน header/template จาก `.bak` เสมอ ทำให้ไม่มีทางทับ header โดยไม่ตั้งใจ  
> ถ้าต้องการเปลี่ยน header → อัพเดท `.bak` ด้วย

---

## ปีหน้า (อัพเดท Report ประจำปี)

1. สร้าง Favorites filter ใหม่ใน Odoo ชื่อ **Sale Report YYYY+1**
2. สร้างไฟล์ xlsm ใหม่ เช่น `Sales Report V3.2.xlsm`
3. แก้ XLSM path ในบรรทัดนี้ของ `daily_report.py`:
   ```python
   XLSM = os.path.join(SCRIPT_DIR, "Sales Report V3.1.xlsm")
   ```
   เปลี่ยนเป็น `Sales Report V3.2.xlsm`
4. สร้าง `.bak` ไฟล์ใหม่จาก xlsm ใหม่

---

*อัพเดทล่าสุด: เมษายน 2026*
