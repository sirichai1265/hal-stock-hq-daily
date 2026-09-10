# HAL Stock HQ – Daily Container Stock

สร้างรายงาน **HAL Stock HQ – Daily** จากไฟล์ export 3 ไฟล์ประจำวัน แล้วเติมลงเทมเพลต
`(HAL)Stock form.xlsx` (ชีต `Daily`, `RF SEASONAL`, `NEW FORMAT`) พร้อมสร้างหน้า
dashboard `index.html` สำหรับ GitHub Pages

## ใช้งาน

```bash
python build_stock_hq.py 2026-09-10            # สร้างรายงาน + index.html
python build_stock_hq.py 2026-09-10 --publish  # + push dashboard ขึ้น GitHub Pages
```

`--publish` จะ copy `index.html` ไปที่ `dashboard-public/` (local clone ของ repo
public `hal-stock-hq-dashboard`) แล้ว commit + push → อัปเดต
https://sirichai1265.github.io/hal-stock-hq-dashboard/

วางไฟล์ 3 ไฟล์ไว้ในโฟลเดอร์เดียวกัน สคริปต์ auto-detect จากชื่อ:

| ไฟล์ | ใช้คำนวณ |
|---|---|
| `*ACTUAL*.xls` | FULL INBOUND (ตู้ FULL คงเหลือ, Move Code ≠ OFD) |
| `*STAYING*.xls` | CURRENT STOCK + RF SEASONAL |
| `*BKG*.xls` | BOOKING WK1ST / WK2ND |

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
