# Stock Report HQ – Stock Daily

สร้างรายงาน **Stock Daily** จากไฟล์ export ประจำวัน แล้วเติมลงเทมเพลต
`(HAL)Stock form.xlsx` (ชีต `Daily`, `RF SEASONAL`, `NEW FORMAT`) พร้อมสร้างหน้า
dashboard `index.html` (หัวเรื่อง "Stock Report HQ") สำหรับ GitHub Pages

> คู่มือ prompt สำหรับสอนคนอื่นใช้งาน: ดูไฟล์ **`PROMPT-ใช้งาน.md`**

## ใช้งาน (ประจำวัน)

1. วางไฟล์ export ของวันนั้นลงในโฟลเดอร์นี้ (ดูตารางด้านล่าง)
2. ดับเบิลคลิก **`run_stock_hq.bat`** — หรือรันเอง:

```bash
python build_stock_hq.py                       # ไม่ใส่วันที่ = ใช้วันนี้
python build_stock_hq.py 2026-10-12            # ระบุวันที่เอง
python build_stock_hq.py --publish             # + push dashboard ขึ้น GitHub Pages
```

สคริปต์หาไฟล์ให้เองจากชื่อ/เนื้อหา และใช้เฉพาะไฟล์ชุดที่อัปโหลดวันล่าสุด (ไม่ต้องลบไฟล์เก่าออก)

| ไฟล์ | ใช้คำนวณ |
|---|---|
| `*ACTUAL*.xls` | FULL INBOUND (ไม่นับ Move Code OFD, ไม่นับตู้เกิน 30 days, IFV ไม่เกิน 30 days นับเป็น IFI) |
| `*STAYING*.xls` | CURRENT STOCK + RF SEASONAL (ไม่นับ Move Code OEP, OEV, OEW) |
| `*PENDING*.xls` + `*BKG*.xls` | BOOKING (รวมสองไฟล์ ถ้ามีไฟล์เดียวก็ใช้ไฟล์เดียว) |
| `*EMPTY*.xls` (ไม่บังคับ) | REPO (E/P) แถว 14 — THBKK → L14:T14, THLCH → AG14:AO14 |

ผลลัพธ์: `Stock Daily <YYYY-MM-DD>.xlsx` (รายงาน) และ `index.html` (dashboard — ยอดรวมรายกลุ่มสถานที่ ไม่มีชื่อลูกค้า/เลขบุ๊คกิ้ง)

หลังรันเสร็จ ให้สร้างโฟลเดอร์ของวัน (เช่น `10-13`) แล้วย้ายไฟล์ดิบทั้งหมด + รายงานเข้าไปเก็บ

## นิยาม

- **WK1ST** = วันล่าสุด → วันอาทิตย์สัปดาห์นี้ (รวมบุ๊คค้างเก่า ไม่มีขอบล่าง)
- **WK2ND** = วันจันทร์ → วันอาทิตย์สัปดาห์หน้า
- **วันศุกร์**: WK1ST รวมถึงอาทิตย์สัปดาห์หน้า และช่องที่สองใช้บุ๊ค WK3RD
- **STOCK BALANCE END** = CURRENT STOCK − BOOKING (คิดต่อเนื่องข้ามสัปดาห์)

## หมายเหตุ

- เครื่องที่ไม่มี LibreOffice: สคริปต์ตั้ง `fullCalcOnLoad` ไว้ สูตรจะคำนวณเมื่อเปิดใน Excel
- ไฟล์ export ดิบ (`*ACTUAL*`, `*STAYING*`, `*BKG*`, `*PENDING*`, `*EMPTY*`) มีข้อมูลลูกค้า ถูก `.gitignore` ไว้ ห้าม push ขึ้น repo

## ติดตั้งบนเครื่องใหม่ / ให้เพื่อนใช้

**แบบที่ 1 — ใช้สร้างรายงาน Excel ในเครื่องตัวเองอย่างเดียว (ไม่ต้องมีบัญชี GitHub)**

1. ติดตั้ง [Python 3](https://www.python.org/downloads/) (ตอนติดตั้งติ๊ก "Add python.exe to PATH")
2. ก็อปทั้งโฟลเดอร์นี้ไปเครื่องเพื่อน (zip/USB/LINE ก็ได้) — ต้องมีอย่างน้อย
   `build_stock_hq.py`, `requirements.txt`, `run_stock_hq.bat`, `(HAL)Stock form.xlsx`
3. เปิด Command Prompt ในโฟลเดอร์นั้น แล้วรันครั้งเดียว:
   ```bash
   pip install -r requirements.txt
   ```
4. ใช้งานประจำวัน: วางไฟล์ประจำวัน + ดับเบิลคลิก `run_stock_hq.bat` → ได้ไฟล์ Excel
   (ส่วน `--publish` จะข้ามอัตโนมัติถ้ายังไม่ได้ตั้ง dashboard-public — ดูแบบที่ 2)

**แบบที่ 2 — เชื่อม GitHub ด้วย เพื่ออัปเดต dashboard สาธารณะร่วมกัน**

ต้องมีบัญชี GitHub และถูกเชิญเป็น collaborator ของ repo นี้ (private) และ repo
[hal-stock-hq-dashboard](https://github.com/sirichai1265/hal-stock-hq-dashboard) (public)
ก่อน — ขอให้เจ้าของ repo (sirichai1265) เชิญผ่าน GitHub: Settings → Collaborators →
Add people

1. ทำแบบที่ 1 ให้ครบก่อน
2. ติดตั้ง [Git](https://git-scm.com/downloads) และ [GitHub CLI](https://cli.github.com/)
   แล้วล็อกอิน: `gh auth login`
3. Clone repo นี้แทนการก็อปโฟลเดอร์:
   ```bash
   git clone https://github.com/sirichai1265/hal-stock-hq-daily.git
   cd hal-stock-hq-daily
   pip install -r requirements.txt
   git clone https://github.com/sirichai1265/hal-stock-hq-dashboard.git dashboard-public
   ```
4. ใช้งานเหมือนแบบที่ 1 — แต่ตอนนี้ `--publish` (ปุ่ม `run_stock_hq.bat` ใช้ค่านี้อยู่แล้ว)
   จะ push dashboard ขึ้น https://sirichai1265.github.io/hal-stock-hq-dashboard/ ได้จริง
5. ถ้าอยาก push โค้ด/ไฟล์รายงานกลับขึ้น repo private ด้วย: `git add -A && git commit -m "..." && git push`
