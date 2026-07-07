# IC Weekly Report System — BPK Power Plant

## ไฟล์ในระบบ

| ไฟล์ | หน้าที่ |
|------|---------|
| `index.html` | Web Form — เปิดผ่าน Browser, Deploy บน GitHub Pages |
| `Code.gs` | Google Apps Script — วางใน Google Sheets, เป็น Backend API |
| `README.md` | วิธีติดตั้งและใช้งาน |

---

## วิธีติดตั้ง (ครั้งแรก ~20 นาที)

### ขั้นที่ 1 — ตั้งค่า Google Sheets + Apps Script

1. เปิด [Google Sheets](https://sheets.google.com) → สร้าง Spreadsheet ใหม่
2. ตั้งชื่อ: **IC Weekly Report 2568**
3. คัดลอก URL Spreadsheet ID (ส่วนที่อยู่ระหว่าง `/d/` และ `/edit`)
4. เมนู **Extensions → Apps Script**
5. ลบโค้ดเดิมทั้งหมด → วางโค้ดจากไฟล์ `Code.gs`
6. ที่บรรทัด `const SPREADSHEET_ID = "";` ใส่ ID ที่คัดลอกไว้ในขั้นที่ 3
7. กด **Save** → **Deploy → New Deployment**
8. ตั้งค่า:
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
9. กด **Deploy** → **คัดลอก Web App URL**

### ขั้นที่ 2 — ตั้งค่า index.html

เปิดไฟล์ `index.html` ค้นหาบรรทัด:
```javascript
const SCRIPT_URL = "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
```
แก้เป็น URL ที่ได้จากขั้นที่ 1 ข้อ 9

### ขั้นที่ 3 — Deploy บน GitHub Pages

1. สร้าง Repository ใหม่บน GitHub (เช่น `ic-weekly-report`)
2. Upload ไฟล์ `index.html` ขึ้นไป
3. Settings → Pages → Branch: **main** → Save
4. รอ 2-3 นาที URL จะเป็น: `https://YOUR-USERNAME.github.io/ic-weekly-report/`
5. แชร์ URL ให้ทีม

---

## การใช้งานรายสัปดาห์

```
ทุกวันจันทร์เช้า:
1. เปิด URL → กรอกข้อมูลสัปดาห์ที่ผ่านมา
2. ค่า Protection Readiness → เปิด PowerEx บน EGAT WiFi แล้วกรอกเอง
3. กด "Preview" ตรวจสอบ
4. กด "Submit" → บันทึกลง Google Sheets อัตโนมัติ
```

## โครงสร้าง Google Sheets หลัง Submit

```
Google Sheets
├── W_2025_W01   ← ข้อมูลสัปดาห์ที่ 1
├── W_2025_W02   ← ข้อมูลสัปดาห์ที่ 2
├── W_2025_W03   ← ข้อมูลสัปดาห์ที่ 3
│   ...
├── Monthly_ม.ค.2568   ← Monthly Summary (Auto-aggregate)
├── Monthly_ก.พ.2568
│   ...
```

## Monthly Summary — ดึงอัตโนมัติ

Sheet `Monthly_xxx` จะสรุปข้อมูลทุกสัปดาห์ในเดือนนั้นให้อัตโนมัติ ได้แก่:
- Total I&C Events ทั้งเดือน
- Average Protection Readiness
- Max WO Overdue
- จำนวนสัปดาห์ที่รายงาน

---

## หมายเหตุ PowerEx

URL `http://10.212.28.3/PowerEx/...` เป็น Private IP ใน EGAT Network
ไม่สามารถดึงข้อมูลอัตโนมัติจาก Web ภายนอกได้ เนื่องจาก:
- Browser Block Cross-origin request
- IP ไม่ได้ Expose ออก Internet
- Cybersecurity Policy ของ EGAT

**วิธีแก้ที่ใช้ได้จริง:** กรอกค่าเองในช่อง Protection Readiness
โดยเปิด PowerEx บน EGAT WiFi แล้วดูค่า → กรอกใน Form (ใช้เวลา < 1 นาที)
