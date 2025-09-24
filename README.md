# BNI TYFCB Auto Scraper

โปรแกรมสำหรับดึงข้อมูล TYFCB (Thank You For Closed Business) จาก BNI Connect Global อัตโนมัติ และบันทึกลง Google Sheets

## คุณสมบัติ

- 🔄 **การรันอัตโนมัติ**: ใช้ GitHub Actions รันทุกสัปดาห์
- 📊 **บันทึกข้อมูลอัตโนมัติ**: บันทึกลง Google Sheets โดยอัตโนมัติ
- 🖥️ **รองรับหลายแพลตฟอร์ม**: ใช้งานได้ทั้ง Windows และ Linux
- 🛡️ **ปลอดภัย**: เก็บรหัสผ่านใน GitHub Secrets

## การตั้งค่าเบื้องต้น

### 1. ติดตั้ง Dependencies

```bash
pip install -r requirements.txt
```

### 2. ตั้งค่า Google Sheets API

#### สร้าง Google Cloud Project และ Service Account:

1. ไปที่ [Google Cloud Console](https://console.cloud.google.com/)
2. สร้าง Project ใหม่หรือเลือก Project ที่มีอยู่
3. เปิดใช้งาน Google Sheets API และ Google Drive API
4. สร้าง Service Account:
   - ไปที่ **IAM & Admin** > **Service Accounts**
   - คลิก **Create Service Account**
   - กรอกชื่อและรายละเอียด
   - คลิก **Create and Continue**
5. สร้าง JSON Key:
   - คลิกที่ Service Account ที่สร้างไว้
   - ไปที่แท็บ **Keys**
   - คลิก **Add Key** > **Create new key**
   - เลือก **JSON** แล้วคลิก **Create**
   - ดาวน์โหลดไฟล์ JSON

#### สร้าง Google Sheet:

1. ไปที่ [Google Sheets](https://sheets.google.com/)
2. สร้าง Spreadsheet ใหม่ชื่อ **"BNI TYFCB Data"**
3. แชร์ Spreadsheet กับอีเมล Service Account (จากไฟล์ JSON)
   - คลิก **Share** บนขวาบน
   - ใส่อีเมล Service Account
   - ตั้งสิทธิ์เป็น **Editor**

### 3. การใช้งานแบบ Local

#### วางไฟล์ Credentials:
```
BNI_Lifetime/
├── google-sheets-credentials.json  # วางไฟล์ JSON ที่ดาวน์โหลดมา
├── BNI-Lifetime-Selenuim-V5.py
└── requirements.txt
```

#### รันโปรแกรม:
```bash
python BNI-Lifetime-Selenuim-V5.py
```

### 4. การตั้งค่า GitHub Actions สำหรับรันอัตโนมัติ

#### ตั้งค่า GitHub Secrets:

1. ไปที่ GitHub repository
2. **Settings** > **Secrets and variables** > **Actions**
3. เพิ่ม Secrets ต่อไปนี้:

| Secret Name | คำอธิบาย | ตัวอย่าง |
|------------|---------|---------|
| `BNI_USERNAME` | ชื่อผู้ใช้ BNI Connect Global | `your-email@example.com` |
| `BNI_PASSWORD` | รหัสผ่าน BNI Connect Global | `your-password` |
| `GOOGLE_SHEETS_CREDENTIALS` | เนื้อหาไฟล์ JSON ทั้งหมด | `{"type": "service_account"...}` |
| `GOOGLE_SHEET_NAME` | ชื่อ Google Sheet | `BNI TYFCB Data` |

#### เปิดใช้งาน GitHub Actions:

1. ไปที่แท็บ **Actions** ใน GitHub repository
2. เปิดใช้งาน GitHub Actions (ถ้ายังไม่ได้เปิด)
3. Workflow จะรันอัตโนมัติทุกวันจันทร์เวลา 16:00 น. (เวลาไทย)

#### รันด้วยตนเองได้:
1. ไปที่แท็บ **Actions**
2. เลือก **Weekly BNI TYFCB Scraping**
3. คลิก **Run workflow**

## โครงสร้างข้อมูลใน Google Sheets

| คอลัมน์ | คำอธิบาย | ตัวอย่าง |
|--------|---------|---------|
| Timestamp | วันที่และเวลาที่ดึงข้อมูล | `2024-01-15 16:30:45` |
| TYFCB Received | ยอดเงิน TYFCB ที่ได้รับ | `1,250,000` |
| Running User | ชื่อผู้ใช้ที่รันโปรแกรม | `john.doe@example.com` |
| Chapter | Chapter ที่สังกัด | `Bangkok Central` |
| Total Given Amount | ยอดรวม TYFCB ที่ให้ | `850,000` |
| Records Count | จำนวนรายการ TYFCB Given | `25` |

## การตรวจสอบและแก้ไขปัญหา

### 1. ตรวจสอบ GitHub Actions
```
Actions > Weekly BNI TYFCB Scraping > ดูล็อกการทำงาน
```

### 2. ดู Screenshot จากการรัน
GitHub Actions จะอัปโหลด screenshot ไว้ในส่วน **Artifacts**

### 3. ปัญหาที่พบบ่อย

#### ❌ Google Sheets permission denied
- ตรวจสอบว่า Service Account มีสิทธิ์ Editor ใน Google Sheet
- ตรวจสอบชื่อ Google Sheet ใน `GOOGLE_SHEET_NAME`

#### ❌ BNI login failed
- ตรวจสอบ `BNI_USERNAME` และ `BNI_PASSWORD`
- ลองล็อกอินด้วยตนเองที่ BNI Connect Global

#### ❌ Chrome/ChromeDriver issues
- GitHub Actions จะติดตั้ง Chrome อัตโนมัติ
- สำหรับ local ให้ดูที่ error message จาก webdriver-manager

### 4. การรันใน Development Mode

```bash
# ตั้งค่า environment variables
export BNI_USERNAME="your-username"
export BNI_PASSWORD="your-password"
export GOOGLE_SHEET_NAME="BNI TYFCB Data (Test)"

# รันโปรแกรม
python BNI-Lifetime-Selenuim-V5.py
```

## กำหนดการรัน

- **GitHub Actions**: ทุกวันจันทร์ เวลา 16:00 น. (เวลาไทย)
- **Manual**: สามารถรันได้ตลอดเวลาผ่าน GitHub Actions หรือ local

## ไฟล์ที่สำคัญ

```
BNI_Lifetime/
├── .github/workflows/weekly-bni-scraping.yml  # GitHub Actions workflow
├── BNI-Lifetime-Selenuim-V5.py               # โปรแกรมหลัก
├── requirements.txt                           # Python dependencies
├── google-sheets-credentials.json             # Google API credentials (local)
├── CLAUDE.md                                  # คู่มือสำหรับ Claude Code
└── README.md                                  # คู่มือนี้
```

## การสนับสนุน

หากพบปัญหาหรือต้องการความช่วยเหลือ:
1. ตรวจสอบ GitHub Actions logs
2. ดู screenshot ใน Artifacts
3. ตรวจสอบ Google Sheets permissions

---

⚠️ **หมายเหตุ**: โปรแกรมนี้เป็นเครื่องมือสำหรับใช้งานภายในองค์กร กรุณาใช้อย่างมีความรับผิดชอบและปฏิบัติตาม Terms of Service ของ BNI Connect Global