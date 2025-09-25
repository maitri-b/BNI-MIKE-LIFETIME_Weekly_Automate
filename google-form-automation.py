# Google Form Automation for BNI Data
# -*- coding: utf-8 -*-

import requests
import json
import os
import time
from datetime import datetime, timedelta
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

    def parse_timestamp(self, timestamp_str):
        """แปลง timestamp string เป็น datetime object"""
        if not timestamp_str:
            return None

        try:
            # ลองหลายรูปแบบ timestamp ที่อาจจะพบ
            formats = [
                "%Y-%m-%d %H:%M:%S",     # 2024-01-15 14:30:00
                "%m/%d/%Y %H:%M:%S",     # 01/15/2024 14:30:00
                "%d/%m/%Y %H:%M:%S",     # 15/01/2024 14:30:00
                "%Y-%m-%dT%H:%M:%S",     # ISO format
                "%Y-%m-%d",              # Date only
                "%m/%d/%Y",              # Date only US format
                "%d/%m/%Y",              # Date only EU format
            ]

            for fmt in formats:
                try:
                    return datetime.strptime(timestamp_str.strip(), fmt)
                except ValueError:
                    continue

            print(f"⚠️  ไม่สามารถแปลง timestamp: {timestamp_str}")
            return None

        except Exception as e:
            print(f"⚠️  ข้อผิดพลาดในการแปลง timestamp {timestamp_str}: {e}")
            return None

    def is_data_recent(self, timestamp_str, days_limit=7):
        """ตรวจสอบว่าข้อมูลอัปเดตมาไม่เกินจำนวนวันที่กำหนด"""
        timestamp = self.parse_timestamp(timestamp_str)
        if not timestamp:
            # ถ้าแปลง timestamp ไม่ได้ ให้ถือว่าเป็นข้อมูลเก่า
            return False

        current_time = datetime.now()
        time_diff = current_time - timestamp

        is_recent = time_diff <= timedelta(days=days_limit)

        if not is_recent:
            print(f"   📅 ข้ามข้อมูลเก่า: {timestamp_str} (เก่ากว่า {days_limit} วัน)")

        return is_recent

    def get_current_sheet_data(self):
        """ดึงข้อมูลปัจจุบันจาก Google Sheets (เฉพาะข้อมูลที่อัปเดตมาไม่เกิน 7 วัน)"""
        try:
            client = self.setup_google_sheets()
            if not client:
                return {}

            sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'BNI TYFCB Data')
            spreadsheet = client.open(sheet_name)
            worksheet = spreadsheet.sheet1

            # ดึงข้อมูลทั้งหมด
            all_records = worksheet.get_all_records()
            print(f"ดึงข้อมูลจาก Google Sheets: {len(all_records)} รายการทั้งหมด")

            # กรองเฉพาะข้อมูลใหม่ (ไม่เกิน 7 วัน)
            recent_data = {}
            old_data_count = 0

            for record in all_records:
                running_user = record.get('Running User', '').strip()
                tyfcb_received = record.get('TYFCB Received', '').strip()
                timestamp_str = record.get('Timestamp', '').strip()

                if running_user and tyfcb_received:
                    # ตรวจสอบว่าข้อมูลใหม่หรือไม่
                    if self.is_data_recent(timestamp_str, days_limit=7):
                        # ใช้ timestamp เป็นส่วนหนึ่งของ key เพื่อแยกข้อมูลเดียวกันที่มี timestamp ต่างกัน
                        data_key = f"{running_user}_{timestamp_str}"
                        recent_data[data_key] = {
                            'running_user': running_user,
                            'tyfcb_received': tyfcb_received,
                            'timestamp': timestamp_str,
                            'chapter': record.get('Chapter', ''),
                            'total_amount': record.get('Total Given Amount', ''),
                            'records_count': record.get('Records Count', 0)
                        }
                    else:
                        old_data_count += 1

            print(f"✅ ข้อมูลใหม่ (ไม่เกิน 7 วัน): {len(recent_data)} รายการ")
            print(f"⏰ ข้อมูลเก่า (เกิน 7 วัน): {old_data_count} รายการ")

            return recent_data

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

        # เปรียบเทียบกับข้อมูลครั้งสุดท้าย (เฉพาะข้อมูลใหม่ที่ไม่เกิน 7 วัน)
        new_entries = []

        for data_key, data in current_data.items():
            if data_key not in self.last_data:
                # ข้อมูลใหม่ที่ยังไม่เคยส่ง
                new_entries.append((data_key, data))

        if len(new_entries) == 0:
            print("ไม่พบข้อมูลใหม่ที่ต้องส่ง")
            return

        print(f"🔍 พบข้อมูลใหม่ที่ต้องส่ง: {len(new_entries)} รายการ")

        # ส่งข้อมูลไปยัง Google Form
        success_count = 0

        for data_key, data in new_entries:
            running_user = data['running_user']
            tyfcb_received = data['tyfcb_received']
            timestamp = data['timestamp']

            print(f"\n[ใหม่] {running_user}: {tyfcb_received} (Timestamp: {timestamp})")
            if self.form_submitter.submit_to_form(running_user, tyfcb_received):
                success_count += 1
            time.sleep(2)  # หน่วงเวลาเพื่อไม่ให้ส่งเร็วเกินไป

        print(f"\n✅ ส่งข้อมูลสำเร็จ: {success_count}/{len(new_entries)} รายการ")

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