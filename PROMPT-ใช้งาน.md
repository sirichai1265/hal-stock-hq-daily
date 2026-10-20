# คู่มือ Prompt – ทำรายงาน Stock Daily (สอนผู้ใช้ใหม่)

มี 2 วิธี เลือกวิธีที่ 1 ถ้ามีโฟลเดอร์ชุดนี้ครบ (สคริปต์ + เทมเพลต) เพราะกฎทั้งหมดอยู่ในสคริปต์แล้ว
วิธีที่ 2 ใช้เมื่อไม่มีสคริปต์ ให้ Claude สร้างรายงานเองจากเทมเพลต

---

## ตั้งค่าครั้งเดียว (วิธีที่ 1)

1. ติดตั้ง [Python 3](https://www.python.org/downloads/) — ตอนติดตั้งติ๊ก **"Add python.exe to PATH"**
2. แตก zip ไว้ที่โฟลเดอร์ เช่น `Desktop\STOCK HQ` (ต้องมี `build_stock_hq.py`, `(HAL)Stock form.xlsx`, `requirements.txt`, `run_stock_hq.bat`)
3. เปิด Command Prompt ในโฟลเดอร์นั้น แล้วรัน `pip install -r requirements.txt`
4. เปิด Claude Code ในโฟลเดอร์นั้น (หรือให้ Claude เข้าถึงโฟลเดอร์นี้)

ไม่ต้องใช้ Claude ก็ได้: วางไฟล์ประจำวันในโฟลเดอร์ แล้วดับเบิลคลิก `run_stock_hq.bat`

---

## วิธีที่ 1 – Prompt ประจำวัน (แนบไฟล์ประจำวัน 5 ไฟล์ แล้วพิมพ์)

```text
ทำรายงาน Stock Daily ของวันนี้จากไฟล์ที่แนบ (ACTUAL, STAYING, EMPTY, PENDING, BKG-3WK)
ใช้ build_stock_hq.py ในโฟลเดอร์ STOCK HQ ตามขั้นตอนนี้:
1. คัดลอกไฟล์เข้า STOCK HQ แล้วรัน build_stock_hq.py (ระบบหาไฟล์จากชื่อ/เนื้อหาเอง)
2. ตรวจตัวเลข CURRENT STOCK กับไฟล์ STAYING ดิบก่อนส่ง
3. สรุปช่องที่ STOCK BALANCE END ติดลบ แยกว่าตู้ FULL ขาเข้ารองรับได้ไหม
   (DEPOT ไม่มี FULL รองรับ = ขาดจริง)
4. สร้างโฟลเดอร์ของวัน (เช่น 10-13) ใน STOCK HQ ย้ายไฟล์ดิบทั้ง 5 + รายงานเข้าไป
   ลบสำเนาซ้ำที่ Desktop หลังตรวจว่าขนาดไฟล์ตรงกัน
5. ส่งไฟล์ Stock Daily YYYY-MM-DD.xlsx ให้ดู
(ถ้าต้องการอัปเดตเว็บ ให้รันด้วย --publish แล้วเช็คหน้าเว็บ)
```

---

## วิธีที่ 2 – Prompt สร้างรายงานเองจากเทมเพลต (ไม่มีสคริปต์)

วางครั้งเดียวต่อโปรเจกต์ (หรือใส่ไว้ในไฟล์ `CLAUDE.md`) แล้วแนบไฟล์ประจำวันทุกวัน

```text
สร้างรายงาน "Stock Daily YYYY-MM-DD.xlsx" โดยเติมชีต Daily ในเทมเพลต (HAL)Stock form.xlsx
(ชีต RF SEASONAL และ NEW FORMAT อ้างอิงจาก Daily) ชนิดตู้ 9 ชนิดตามลำดับ
22GP 42GP 45GP 22RE 45RE 22UT 42UT 22PC 42PC

ไฟล์ต้นทาง (จับคู่ตามเนื้อหา ไม่ยึดชื่อไฟล์ ใช้ชีทชื่อ "Sheet"):
- ACTUAL = FULL INBOUND: นับตาม Location+Size/Type ไม่นับ OFD, ไม่นับตู้เกิน 30 days (คอลัมน์ Days), IFV ที่ BKNO(Cur) ขึ้นต้น HASLS22 ไม่นับ แต่ IFV ไม่เกิน 30 days นับเป็น IFI
- STAYING = CURRENT STOCK: นับตาม Location+Size/Type (ไม่นับ Move Code OEP, OEV, OEW)
- PENDING + BKG-3WK = BOOKING (รวมสองไฟล์, BK No ซ้ำให้ใช้ตัวหลัง,
  ตัดแถวสรุปท้ายไฟล์, หัวตารางอาจอยู่แถว 0 หรือ 1)
- EMPTY = REPO (E/P) แถว 14: THBKK -> L14:T14, THLCH -> AG14:AO14
  (ไฟล์เป็นรายตู้ ให้นับตามคอลัมน์ Type Size)

กลุ่มสถานที่ BKK: PAT=BKK01+BKK04, UNITHAI=BKK02, ESCO #2=LCH55, TIFFA=LCH54,
SMART=BKK25 (DEPOT), BC2=BKK27 (DEPOT), BMTP=BKKY4, SCT=LCHY4
LCH: LCMT=LCH04+07, ESCO=LCH01, LCIT=LCH05+10, HPT=LCH02+06+08+09,
HAST=LCH27, CELLO=LCH28, PW=LCHY5 (HAST/CELLO/PW = DEPOT)

ช่วงบุ๊คกิ้ง: WK1ST = ทุกแถวถึงวันอาทิตย์สัปดาห์นี้ (รวมค้างเก่า ไม่มีขอบล่าง),
WK2ND = จันทร์-อาทิตย์สัปดาห์หน้า
วันศุกร์: WK1ST รวมถึงอาทิตย์สัปดาห์หน้า และช่องที่สองใช้ WK3RD
STOCK BALANCE END = CURRENT STOCK - BOOKING WK1ST และต่อเนื่องลบ WK2ND

รูปแบบ: ช่อง 0 ไม่แสดง (ยกเว้น TOTAL), DEPOT ตัวอักษรสีน้ำเงิน, TERMINAL/ICD สีดำ,
ช่องติดลบตัวหนาพื้นชมพู, ห้ามเขียนทับสูตรในเทมเพลต,
LCMT/ESCO/LCIT/HPT ไม่แสดง STOCK BALANCE END, NEW FORMAT แถว Balance
แดงเมื่อ <= 0 นอกนั้นน้ำเงิน, หัวตารางต่อท้าย WK1ST/WK2ND, RF SEASONAL ปี 2025/2023/2021/2020

หลังสร้างเสร็จ: ตรวจตัวเลขกับไฟล์ดิบ, สรุปช่องติดลบ, สร้างโฟลเดอร์ของวัน (เช่น 10-13)
ย้ายไฟล์ดิบทั้งหมด + รายงานเข้าไป, ห้ามอัปโหลดไฟล์ดิบขึ้น GitHub (มีข้อมูลลูกค้า)
```

---

## ข้อควรรู้

- ไฟล์ดิบ (ACTUAL / STAYING / PENDING / BKG / EMPTY) มีข้อมูลลูกค้าและเลขบุ๊คกิ้ง **ห้ามอัปโหลดขึ้น GitHub หรือส่งออกนอกบริษัท**
- ปุ่ม `--publish` อัปเดตหน้าเว็บ ต้องมีสิทธิ์เขียน repo `hal-stock-hq-dashboard` และล็อกอิน `gh` ไว้ ถ้าไม่ต้องการเว็บให้ข้าม จะได้ไฟล์ Excel ครบ
- เครื่องที่ไม่มี LibreOffice: สูตรจะคำนวณตอนเปิดไฟล์ใน Excel ครั้งแรก
- ชื่อไฟล์ต้นทางเปลี่ยนได้ (เช่น มีคำว่า Copy, ชื่อพิมพ์ผิด) สคริปต์ดูเนื้อหาไฟล์ช่วยด้วย แต่ถ้าหาไฟล์ไม่เจอให้ระบุเอง: `python build_stock_hq.py --actual ... --staying ... --bkg ... --bkg ... --ep2 ...`
- ตรวจผลทุกวัน: เทียบ CURRENT STOCK ของ SMART/BC2 กับจำนวนแถวในไฟล์ STAYING ดิบ ต้องตรงกัน
