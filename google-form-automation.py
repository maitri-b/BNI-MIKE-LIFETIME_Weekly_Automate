# Google Form Automation for BNI Data
# -*- coding: utf-8 -*-

import requests
import json
import os
import time
from datetime import datetime
import re

# Google Sheets API imports
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    print("Google Sheets API ไม่พร้อมใช้งาน กรุณาติดตั้ง: pip install gspread google-auth")
    GOOGLE_SHEETS_AVAILABLE = False

class GoogleFormSubmitter:
    def __init__(self):
        self.form_url = "https://docs.google.com/forms/d/e/1FAIpQLSfBkXWsGZXP3IXJ8gR2vZbyAi7VP3R2FSF6YB9ohkr94rIb8g/formResponse"
        self.name_entry = "entry.683444359"  # ชื่อ (dropdown)
        self.business_entry = "entry.290745485"  # ยอดธุรกิจ Lifetime

        # ข้อมูลที่เคยส่งไปแล้ว (เพื่อป้องกันการส่งซ้ำ)
        self.sent_data_file = "sent_form_data.json"
        self.load_sent_data()

    def load_sent_data(self):
        """โหลดข้อมูลที่เคยส่งไปแล้ว"""
        try:
            if os.path.exists(self.sent_data_file):
                with open(self.sent_data_file, 'r', encoding='utf-8') as f:
                    self.sent_data = json.load(f)
            else:
                self.sent_data = {}
        except Exception as e:
            print(f"ไม่สามารถโหลดข้อมูลที่เคยส่งได้: {e}")
            self.sent_data = {}

    def save_sent_data(self):
        """บันทึกข้อมูลที่ส่งแล้ว"""
        try:
            with open(self.sent_data_file, 'w', encoding='utf-8') as f:
                json.dump(self.sent_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"ไม่สามารถบันทึกข้อมูลที่ส่งแล้ว: {e}")

    def clean_amount(self, amount_str):
        """ทำความสะอาดข้อมูลยอดเงิน เอาเฉพาะตัวเลข"""
        if not amount_str:
            return "0"
        # เอาเฉพาะตัวเลข จุดทศนิยม และคอมมา
        cleaned = re.sub(r'[^\d,.]', '', str(amount_str))
        # แทนที่คอมมาด้วยจุด (ถ้าใช้คอมมาเป็น decimal separator)
        cleaned = cleaned.replace(',', '')
        return cleaned if cleaned else "0"

    def submit_to_form(self, name, business_amount):
        """ส่งข้อมูลไปยัง Google Form"""
        try:
            # ทำความสะอาดข้อมูล
            clean_amount = self.clean_amount(business_amount)

            # ตรวจสอบว่าเคยส่งข้อมูลนี้ไปแล้วหรือไม่
            data_key = f"{name}_{clean_amount}"
            if data_key in self.sent_data:
                print(f"ข้ามการส่ง: ข้อมูลของ {name} ยอด {clean_amount} เคยส่งไปแล้วเมื่อ {self.sent_data[data_key]}")
                return False

            # เตรียมข้อมูลสำหรับส่ง
            form_data = {
                self.name_entry: name,
                self.business_entry: clean_amount
            }

            # ส่งข้อมูล
            print(f"กำลังส่งข้อมูล: {name} = {clean_amount}")

            response = requests.post(
                self.form_url,
                data=form_data,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                timeout=30
            )

            if response.status_code == 200:
                print(f"✅ ส่งข้อมูลสำเร็จ: {name} = {clean_amount}")
                # บันทึกว่าส่งแล้ว
                self.sent_data[data_key] = datetime.now().isoformat()
                self.save_sent_data()
                return True
            else:
                print(f"❌ ส่งข้อมูลไม่สำเร็จ: Status {response.status_code}")
                return False

        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการส่งข้อมูล: {e}")
            return False

class BNIDataMonitor:
    def __init__(self):
        self.form_submitter = GoogleFormSubmitter()
        self.last_data_file = "last_bni_data.json"
        self.load_last_data()

    def setup_google_sheets(self):
        """ตั้งค่าการเชื่อมต่อ Google Sheets API"""
        if not GOOGLE_SHEETS_AVAILABLE:
            print("Google Sheets API ไม่พร้อมใช้งาน")
            return None

        try:
            # ลองใช้ environment variable ก่อน
            credentials_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
            if credentials_json:
                credentials_info = json.loads(credentials_json)
                scope = [
                    "https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"
                ]
                credentials = Credentials.from_service_account_info(credentials_info, scopes=scope)
            else:
                # ใช้ไฟล์ local
                credentials_file = "google-sheets-credentials.json"
                if not os.path.exists(credentials_file):
                    print(f"ไม่พบไฟล์ {credentials_file}")
                    return None

                scope = [
                    "https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"
                ]
                credentials = Credentials.from_service_account_file(credentials_file, scopes=scope)

            client = gspread.authorize(credentials)
            print("เชื่อมต่อ Google Sheets สำเร็จ")
            return client

        except Exception as e:
            print(f"ไม่สามารถเชื่อมต่อ Google Sheets: {str(e)}")
            return None

    def load_last_data(self):
        """โหลดข้อมูลครั้งสุดท้ายที่ตรวจสอบ"""
        try:
            if os.path.exists(self.last_data_file):
                with open(self.last_data_file, 'r', encoding='utf-8') as f:
                    self.last_data = json.load(f)
            else:
                self.last_data = {}
        except Exception as e:
            print(f"ไม่สามารถโหลดข้อมูลครั้งสุดท้าย: {e}")
            self.last_data = {}

    def save_last_data(self, current_data):
        """บันทึกข้อมูลปัจจุบัน"""
        try:
            with open(self.last_data_file, 'w', encoding='utf-8') as f:
                json.dump(current_data, f, ensure_ascii=False, indent=2)
            self.last_data = current_data
        except Exception as e:
            print(f"ไม่สามารถบันทึกข้อมูลปัจจุบัน: {e}")

    def get_current_sheet_data(self):
        """ดึงข้อมูลปัจจุบันจาก Google Sheets"""
        try:
            client = self.setup_google_sheets()
            if not client:
                return {}

            sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'BNI TYFCB Data')
            spreadsheet = client.open(sheet_name)
            worksheet = spreadsheet.sheet1

            # ดึงข้อมูลทั้งหมด
            all_records = worksheet.get_all_records()

            # สร้าง dict จากข้อมูล โดยใช้ Running User เป็น key
            current_data = {}
            for record in all_records:
                running_user = record.get('Running User', '').strip()
                tyfcb_received = record.get('TYFCB Received', '').strip()

                if running_user and tyfcb_received:
                    current_data[running_user] = {
                        'tyfcb_received': tyfcb_received,
                        'timestamp': record.get('Timestamp', ''),
                        'chapter': record.get('Chapter', ''),
                        'total_amount': record.get('Total Given Amount', ''),
                        'records_count': record.get('Records Count', 0)
                    }

            print(f"ดึงข้อมูลจาก Google Sheets สำเร็จ: {len(current_data)} รายการ")
            return current_data

        except Exception as e:
            print(f"ไม่สามารถดึงข้อมูลจาก Google Sheets: {e}")
            return {}

    def detect_new_data(self, force_check=False):
        """ตรวจสอบข้อมูลใหม่และส่งไปยัง Google Form"""
        print("=" * 50)
        print("เริ่มตรวจสอบข้อมูลใหม่...")

        if force_check:
            print("🔄 โหมดตรวจสอบทั้งหมด (ไม่สนใจข้อมูลเก่า)")
            self.last_data = {}

        # ดึงข้อมูลปัจจุบัน
        current_data = self.get_current_sheet_data()
        if not current_data:
            print("ไม่พบข้อมูลใหม่")
            return

        # เปรียบเทียบกับข้อมูลครั้งสุดท้าย
        new_entries = []
        updated_entries = []

        for user, data in current_data.items():
            if user not in self.last_data:
                # ข้อมูลใหม่
                new_entries.append((user, data))
            elif self.last_data[user]['tyfcb_received'] != data['tyfcb_received']:
                # ข้อมูลที่อัปเดต
                updated_entries.append((user, data))

        total_to_process = len(new_entries) + len(updated_entries)

        if total_to_process == 0:
            print("ไม่พบข้อมูลใหม่หรือข้อมูลที่เปลี่ยนแปลง")
            return

        print(f"พบข้อมูลใหม่: {len(new_entries)} รายการ")
        print(f"พบข้อมูลที่อัปเดต: {len(updated_entries)} รายการ")

        # ส่งข้อมูลไปยัง Google Form
        success_count = 0

        # ส่งข้อมูลใหม่
        for user, data in new_entries:
            print(f"\n[ใหม่] {user}: {data['tyfcb_received']}")
            if self.form_submitter.submit_to_form(user, data['tyfcb_received']):
                success_count += 1
            time.sleep(2)  # หน่วงเวลาเพื่อไม่ให้ส่งเร็วเกินไป

        # ส่งข้อมูลที่อัปเดต
        for user, data in updated_entries:
            print(f"\n[อัปเดต] {user}: {data['tyfcb_received']} (เดิม: {self.last_data[user]['tyfcb_received']})")
            if self.form_submitter.submit_to_form(user, data['tyfcb_received']):
                success_count += 1
            time.sleep(2)  # หน่วงเวลาเพื่อไม่ให้ส่งเร็วเกินไป

        print(f"\n✅ ส่งข้อมูลสำเร็จ: {success_count}/{total_to_process} รายการ")

        # บันทึกข้อมูลปัจจุบัน
        self.save_last_data(current_data)

def main():
    """ฟังก์ชันหลักสำหรับตรวจสอบและส่งข้อมูลไปยัง Google Form"""
    print("โปรแกรม Google Form Automation สำหรับข้อมูล BNI TYFCB")
    print("=" * 60)

    # ตรวจสอบว่าเป็นโหมด force check หรือไม่
    force_check = os.getenv('FORCE_CHECK', 'false').lower() == 'true'

    monitor = BNIDataMonitor()
    monitor.detect_new_data(force_check=force_check)

    print("=" * 60)
    print("การตรวจสอบเสร็จสิ้น")

if __name__ == "__main__":
    main()