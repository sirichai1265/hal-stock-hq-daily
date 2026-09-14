# HAL Stock HQ – Daily Container Stock

สร้างรายงาน **HAL Stock HQ – Daily** จากไฟล์ export 3 ไฟล์ประจำวัน แล้วเติมลงเทมเพลต
`(HAL)Stock form.xlsx` (ชีต `Daily`, `RF SEASONAL`, `NEW FORMAT`) พร้อมสร้างหน้า
dashboard `index.html` สำหรับ GitHub Pages

## ใช้งาน (ประจำวัน)

1. วางไฟล์ export ของวันนั้น 3 ไฟล์ (`*ACTUAL*.xls`, `*STAYING*.xls`, `*BKG*.xls`) ลงในโฟลเดอร์นี้
2. ดับเบิลคลิก **`run_stock_hq.bat`** — หรือรันเอง:

```bash
python build_stock_hq.py --publish             # ไม่ใส่วันที่ = ใช้วันนี้
python build_stock_hq.py 2026-09-10 --publish  # ระบุวันที่เอง
python build_stock_hq.py 2026-09-10            # สร้างรายงานอย่างเดียว ไม่ push dashboard
```

ไม่ต้องลบไฟล์เก่าออกก่อน — สคริปต์เลือกไฟล์ที่ **อัปโหลดล่าสุด** (เรียงตามเวลาแก้ไขไฟล์) ของแต่ละประเภทให้เอง

`--publish` จะ copy `index.html` ไปที่ `dashboard-public/` (local clone ของ repo
public `hal-stock-hq-dashboard`) แล้ว commit + push → อัปเดต
https://sirichai1265.github.io/hal-stock-hq-dashboard/

วางไฟล์ 3 ไฟล์ไว้ในโฟลเดอร์เดียวกัน สคริปต์ auto-detect จากชื่อ:

| ไฟล์ | ใช้คำนวณ |
|---|---|
| `*ACTUAL*.xls` | FULL INBOUND (ตู้ FULL คงเหลือ, Move Code ≠ OFD) |
| `*STAYING*.xls` | CURRENT STOCK + RF SEASONAL |
| `*BKG*.xls` | BOOKING WK1ST / WK2ND |
| `*EP2*.xls` (ไม่บังคับ) | REPO (E/P) แถว 14 — รวมตามชนิดตู้จากคอลัมน์ P.O.D (THBKK/THLCH) วันไหนไม่มีไฟล์นี้ แถว 14 จะเว้นว่างไว้ |

ผลลัพธ์:
- `HAL Stock HQ - Daily <YYYY-MM-DD>.xlsx` — ไฟล์รายงาน
- `index.html` — dashboard (ยอดรวมรายกลุ่มสถานที่เท่านั้น ไม่มีชื่อลูกค้า/เลขบุ๊คกิ้ง)

## นิยาม

- **WK 1ST** = วันล่าสุด → วันอาทิตย์สัปดาห์แรก (รวม backlog ที่ค้าง)
- **WK 2ND** = วันจันทร์ → วันอาทิตย์สัปดาห์ที่สอง
- **STOCK END WK** = CURRENT STOCK − BOOKING (คิดต่อเนื่องข้ามสัปดาห์)

## หมายเหตุ

- เครื่องที่ไม่มี LibreOffice: สคริปต์ตั้ง `fullCalcOnLoad` ไว้ สูตรจะคำนวณเมื่อเปิดใน Excel
- ไฟล์ export ดิบ (`*ACTUAL*`, `*STAYING*`, `*BKG*`, `*.docx`) ถูก `.gitignore` ไว้ ไม่ push ขึ้น repo

## ติดตั้งบนเครื่องใหม่ / ให้เพื่อนใช้

**แบบที่ 1 — ใช้สร้างรายงาน Excel ในเครื่องตัวเองอย่างเดียว (ไม่ต้องมีบัญชี GitHub)**

1. ติดตั้ง [Python 3](https://www.python.org/downloads/) (ตอนติดตั้งติ๊ก "Add python.exe to PATH")
2. ก็อปทั้งโฟลเดอร์นี้ไปเครื่องเพื่อน (zip/USB/LINE ก็ได้) — ต้องมีอย่างน้อย
   `build_stock_hq.py`, `requirements.txt`, `run_stock_hq.bat`, `(HAL)Stock form.xlsx`
3. เปิด Command Prompt ในโฟลเดอร์นั้น แล้วรันครั้งเดียว:
   ```bash
   pip install -r requirements.txt
   ```
4. ใช้งานประจำวัน: วางไฟล์ 3 ไฟล์ + ดับเบิลคลิก `run_stock_hq.bat` → ได้ไฟล์ Excel
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
