# Google Form Automation using Selenium with Prefilled Data
# -*- coding: utf-8 -*-

import time
import os
import json
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# Google Sheets API imports
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    print("Google Sheets API ไม่พร้อมใช้งาน กรุณาติดตั้ง: pip install gspread google-auth")
    GOOGLE_SHEETS_AVAILABLE = False

class GoogleFormSeleniumAutomation:
    def __init__(self):
        self.form_url = "https://docs.google.com/forms/d/e/1FAIpQLSfBkXWsGZXP3IXJ8gR2vZbyAi7VP3R2FSF6YB9ohkr94rIb8g/viewform"
        self.prefill_name = "Maitri+Boonkijrungpaisan"  # URL encoded name
        self.sheet_id = "1MmuiQ2gRNbaA84YTXB2HvDR7MDIyW_buwELkkVm95Qs"  # Google Sheets ID
        self.sheet_name = "BNI TYFCB Data"
        self.driver = None

    def setup_google_sheets_client(self):
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
                service_email = credentials_info.get('client_email', 'ไม่พบ email')
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

                # อ่าน service email จากไฟล์
                with open(credentials_file, 'r') as f:
                    creds_data = json.load(f)
                    service_email = creds_data.get('client_email', 'ไม่พบ email')

            client = gspread.authorize(credentials)
            print(f"✅ เชื่อมต่อ Google Sheets สำเร็จ")
            print(f"📧 Service Account: {service_email}")
            return client

        except Exception as e:
            print(f"ไม่สามารถเชื่อมต่อ Google Sheets: {str(e)}")
            return None

    def get_latest_tyfcb_received(self):
        """ดึงข้อมูล TYFCB Received ล่าสุดจาก Google Sheets"""
        try:
            client = self.setup_google_sheets_client()
            if not client:
                return None

            # เปิด sheet โดยใช้ sheet ID
            print(f"📊 เปิด Google Sheet ID: {self.sheet_id}")
            spreadsheet = client.open_by_key(self.sheet_id)
            worksheet = spreadsheet.sheet1

            # ดึงข้อมูลทั้งหมด
            all_records = worksheet.get_all_records()
            print(f"📊 ดึงข้อมูลจาก Google Sheets: {len(all_records)} รายการ")

            if not all_records:
                print("ไม่พบข้อมูลใน Google Sheets")
                return None

            # แสดงข้อมูล headers เพื่อ debug
            headers = list(all_records[0].keys()) if all_records else []
            print(f"🔍 Headers ที่พบ: {headers}")

            # ค้นหา TYFCB Received ที่ไม่ว่าง โดยเริ่มจากแถวล่าสุด
            tyfcb_received = None
            valid_record = None

            # วนลูปจากแถวสุดท้ายไปแถวแรก
            for i in range(len(all_records) - 1, -1, -1):
                record = all_records[i]

                # ลองหา column ที่มีคำว่า TYFCB หรือ Received
                possible_columns = [
                    'TYFCB Received',
                    'TYFCB received',
                    'tyfcb received',
                    'TYFCB_Received',
                    'Received',
                    'ยอดธุรกิจ Lifetime'  # กรณีที่เป็นภาษาไทย
                ]

                for col in possible_columns:
                    if col in record:
                        value = record[col]
                        if value and str(value).strip() != '':
                            tyfcb_received = value
                            valid_record = record
                            print(f"🎯 พบข้อมูลใน column '{col}' ที่แถว {i+2}: {value}")
                            break

                if tyfcb_received:
                    break

            if not tyfcb_received:
                print("❌ ไม่พบข้อมูล TYFCB Received ที่ไม่ว่าง")
                # แสดงตัวอย่างข้อมูลจากแถวสุดท้าย 3 แถว
                print("📋 ตัวอย่างข้อมูลจากแถวสุดท้าย 3 แถว:")
                for i, record in enumerate(all_records[-3:]):
                    print(f"   แถว {len(all_records)-2+i}: {record}")
                return None

            # แปลงข้อมูลให้เป็นรูปแบบที่เหมาะสม
            if isinstance(tyfcb_received, (int, float)):
                tyfcb_received_str = f"{tyfcb_received:,.0f}"
            else:
                # ถ้าเป็น string ให้ทำความสะอาดและเช็คว่าเป็นตัวเลขไหม
                cleaned = re.sub(r'[^\d,.]', '', str(tyfcb_received))
                if cleaned and cleaned.replace(',', '').replace('.', '').isdigit():
                    try:
                        numeric_value = float(cleaned.replace(',', ''))
                        tyfcb_received_str = f"{numeric_value:,.0f}"
                    except:
                        tyfcb_received_str = cleaned
                else:
                    tyfcb_received_str = str(tyfcb_received).strip()

            print(f"💰 TYFCB Received ที่พบ: {tyfcb_received_str}")

            # แสดงข้อมูลเพิ่มเติมเพื่อ debug
            if valid_record:
                print(f"📋 ข้อมูลในแถวที่เลือก: {valid_record}")

            return tyfcb_received_str

        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูล TYFCB Received: {e}")
            return None

    def setup_driver(self):
        """ตั้งค่า Chrome WebDriver"""
        chrome_options = Options()

        # ตรวจสอบว่าทำงานใน GitHub Actions หรือไม่
        if os.getenv('GITHUB_ACTIONS'):
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            print("กำลังรันในโหมด headless สำหรับ GitHub Actions")
        else:
            chrome_options.add_argument("--start-maximized")

        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)

        try:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            return driver
        except Exception as e:
            print(f"ไม่สามารถตั้งค่า WebDriver ได้: {str(e)}")
            raise Exception("ไม่สามารถเริ่มต้น Chrome WebDriver ได้")

    def fill_and_submit_form(self, tyfcb_amount):
        """เปิด Google Form, กรอกข้อมูล และ submit"""
        try:
            print("🚀 เริ่มต้นการกรอก Google Form...")

            # ตั้งค่า WebDriver
            self.driver = self.setup_driver()

            # สร้าง URL พร้อม prefill data
            prefill_url = f"{self.form_url}?usp=pp_url&entry.683444359={self.prefill_name}&entry.290745485={tyfcb_amount}"
            print(f"🔗 URL: {prefill_url}")

            # เปิดหน้า form
            self.driver.get(prefill_url)
            print("📄 เปิดหน้า Google Form สำเร็จ")

            # รอให้หน้าเว็บโหลดเสร็จ
            time.sleep(3)

            # บันทึกภาพหน้าจอหน้าแรกเพื่อตรวจสอบ
            if not os.getenv('GITHUB_ACTIONS'):
                self.driver.save_screenshot("form_page1.png")
                print("📸 บันทึกภาพหน้าจอหน้าแรกไว้ที่ form_page1.png")

            # หาและคลิกปุ่ม Next เพื่อไปหน้าถัดไป
            print("🔍 กำลังหาปุ่ม Next...")
            try:
                # วิธีที่ 1: หาด้วย jsname="OCpkoe" (จาก HTML ที่ให้มา)
                next_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[jsname='OCpkoe']"))
                )
                print("✅ พบปุ่ม Next (วิธีที่ 1: jsname)")

            except TimeoutException:
                try:
                    # วิธีที่ 2: หาด้วย class และข้อความ "Next"
                    next_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@role='button']//span[contains(text(), 'Next')]"))
                    )
                    print("✅ พบปุ่ม Next (วิธีที่ 2: XPath)")

                except TimeoutException:
                    try:
                        # วิธีที่ 3: หาด้วย class combinations
                        next_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.uArJ5e.UQuaGc.YhQJj.zo8FOc.ctEux"))
                        )
                        print("✅ พบปุ่ม Next (วิธีที่ 3: CSS classes)")

                    except TimeoutException:
                        # วิธีที่ 4: หาทุกปุ่มและเลือกที่มีข้อความ Next
                        all_buttons = self.driver.find_elements(By.CSS_SELECTOR, "div[role='button']")
                        next_button = None
                        for button in all_buttons:
                            if 'Next' in button.text or 'ต่อไป' in button.text:
                                next_button = button
                                print("✅ พบปุ่ม Next (วิธีที่ 4: ค้นหาทั่วไป)")
                                break

                        if not next_button:
                            print("❌ ไม่พบปุ่ม Next")
                            return False

            # คลิกปุ่ม Next
            if next_button:
                print(f"🔤 ข้อความบนปุ่ม: '{next_button.text}'")
                try:
                    self.driver.execute_script("arguments[0].click();", next_button)
                    print("✅ คลิกปุ่ม Next สำเร็จ")

                    # รอให้หน้าถัดไปโหลด
                    time.sleep(3)

                    # บันทึกภาพหน้าจอหน้าที่ 2
                    if not os.getenv('GITHUB_ACTIONS'):
                        self.driver.save_screenshot("form_page2.png")
                        print("📸 บันทึกภาพหน้าจอหน้าที่ 2 ไว้ที่ form_page2.png")

                except Exception as e:
                    print(f"❌ ไม่สามารถคลิกปุ่ม Next: {e}")
                    return False


            # ตรวจสอบว่าข้อมูลถูก prefill หรือไม่
            try:
                # หาฟิลด์ชื่อ (entry.683444359)
                name_field = self.driver.find_element(By.XPATH, "//input[@data-params*='683444359']")
                name_value = name_field.get_attribute('value')
                print(f"👤 ชื่อที่ prefill: {name_value}")

                # หาฟิลด์ยอดธุรกิจ (entry.290745485)
                amount_field = self.driver.find_element(By.XPATH, "//input[@data-params*='290745485']")
                amount_value = amount_field.get_attribute('value')
                print(f"💰 ยอดธุรกิจที่ prefill: {amount_value}")

            except Exception as e:
                print(f"⚠️  ไม่สามารถตรวจสอบข้อมูล prefill: {e}")

                # ลองหาและกรอกข้อมูลด้วยวิธีอื่น
                try:
                    # หาฟิลด์โดยใช้ name attribute
                    name_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
                    if len(name_inputs) >= 2:
                        # ฟิลด์แรกน่าจะเป็นชื่อ
                        if not name_inputs[0].get_attribute('value'):
                            name_inputs[0].clear()
                            name_inputs[0].send_keys("Maitri Boonkijrungpaisan")
                            print("👤 กรอกชื่อแบบ manual")

                        # ฟิลด์ที่สองน่าจะเป็นยอดธุรกิจ
                        if not name_inputs[1].get_attribute('value'):
                            name_inputs[1].clear()
                            name_inputs[1].send_keys(str(tyfcb_amount))
                            print("💰 กรอกยอดธุรกิจแบบ manual")

                except Exception as e2:
                    print(f"⚠️  ไม่สามารถกรอกข้อมูลแบบ manual: {e2}")

            # หาปุ่ม Submit
            print("🔍 กำลังหาปุ่ม Submit...")
            try:
                # วิธีที่ 1: หาปุ่มที่มีข้อความ Submit หรือ ส่ง
                submit_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Submit') or contains(text(), 'ส่ง') or contains(text(), 'ยืนยัน')]/ancestor::div[@role='button']"))
                )

            except TimeoutException:
                # วิธีที่ 2: หาปุ่มด้วย role="button"
                try:
                    submit_buttons = self.driver.find_elements(By.CSS_SELECTOR, "div[role='button']")
                    submit_button = None

                    for button in submit_buttons:
                        button_text = button.text.lower()
                        if any(word in button_text for word in ['submit', 'ส่ง', 'ยืนยัน', 'send']):
                            submit_button = button
                            break

                    if not submit_button and submit_buttons:
                        # ใช้ปุ่มสุดท้าย (มักจะเป็นปุ่ม Submit)
                        submit_button = submit_buttons[-1]

                except Exception as e:
                    print(f"❌ ไม่สามารถหาปุ่ม Submit: {e}")
                    return False

            if submit_button:
                print("✅ พบปุ่ม Submit")
                print(f"🔤 ข้อความบนปุ่ม: '{submit_button.text}'")

                # คลิกปุ่ม Submit
                try:
                    self.driver.execute_script("arguments[0].click();", submit_button)
                    print("✅ คลิกปุ่ม Submit สำเร็จ")

                    # รอให้ form submit เสร็จ
                    time.sleep(5)

                    # ตรวจสอบว่า submit สำเร็จหรือไม่
                    current_url = self.driver.current_url
                    if "formResponse" in current_url or "thanks" in current_url.lower():
                        print("🎉 Submit form สำเร็จ!")
                        return True
                    else:
                        print("⚠️  ไม่แน่ใจว่า submit สำเร็จหรือไม่")
                        print(f"URL ปัจจุบัน: {current_url}")
                        return True

                except Exception as e:
                    print(f"❌ ไม่สามารถคลิกปุ่ม Submit: {e}")
                    return False
            else:
                print("❌ ไม่พบปุ่ม Submit")
                return False

        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการกรอก form: {e}")
            return False

        finally:
            # ปิด WebDriver
            if self.driver:
                if not os.getenv('GITHUB_ACTIONS'):
                    # บันทึกภาพหน้าจอสุดท้าย
                    try:
                        self.driver.save_screenshot("form_final.png")
                        print("📸 บันทึกภาพหน้าจอสุดท้ายไว้ที่ form_final.png")
                    except:
                        pass

                print("🔒 ปิด WebDriver")
                self.driver.quit()

    def run(self):
        """ฟังก์ชันหลักสำหรับรันโปรแกรม"""
        print("🤖 เริ่มต้นโปรแกรม Google Form Automation")
        print("=" * 60)

        # ดึงข้อมูล TYFCB Received จาก Google Sheets
        print("📊 ดึงข้อมูล TYFCB Received จาก Google Sheets...")
        tyfcb_amount = self.get_latest_tyfcb_received()

        if not tyfcb_amount:
            print("❌ ไม่สามารถดึงข้อมูล TYFCB Received ได้")
            return False

        print(f"💰 ยอดธุรกิจที่จะกรอก: {tyfcb_amount}")

        # กรอกและส่ง Google Form
        print("\n📝 เริ่มกรอก Google Form...")
        success = self.fill_and_submit_form(tyfcb_amount)

        print("\n" + "=" * 60)
        if success:
            print("✅ กรอกและส่ง Google Form สำเร็จ!")
            print(f"👤 ชื่อ: Maitri Boonkijrungpaisan")
            print(f"💰 ยอดธุรกิจ Lifetime: {tyfcb_amount}")
        else:
            print("❌ ไม่สามารถกรอกหรือส่ง Google Form ได้")

        print("🏁 โปรแกรมทำงานเสร็จสิ้น")
        return success


def main():
    """ฟังก์ชันหลัก"""
    automation = GoogleFormSeleniumAutomation()
    success = automation.run()

    # ตรวจสอบว่าทำงานใน automated environment หรือไม่
    if not os.getenv('GITHUB_ACTIONS') and not os.getenv('CI'):
        print("\nกด Enter เพื่อออกจากโปรแกรม...")
        input()

    return success


if __name__ == "__main__":
    main()