# BNI Integrated Automation - Scrape TYFCB Received and Submit to Google Form
# -*- coding: utf-8 -*-
"""
รวมการทำงานของ:
1. BNI-Lifetime-Selenuim-V5.py - ดึงข้อมูล TYFCB Received จาก BNI Connect
2. google-form-selenium-automation.py - กรอก Google Form อัตโนมัติ
"""

import time
import os
import getpass
import json
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

class BNIIntegratedAutomation:
    def __init__(self):
        # BNI Connect settings
        self.bni_login_url = "https://www.bniconnectglobal.com/login"
        self.bni_dashboard_url = "https://www.bniconnectglobal.com/web/dashboard"

        # Google Form settings
        self.form_url = "https://docs.google.com/forms/d/e/1FAIpQLSfBkXWsGZXP3IXJ8gR2vZbyAi7VP3R2FSF6YB9ohkr94rIb8g/viewform"
        self.prefill_name = "Maitri+Boonkijrungpaisan"  # URL encoded name

        self.driver = None
        self.tyfcb_received = None

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
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-background-timer-throttling")
            chrome_options.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options.add_argument("--disable-renderer-backgrounding")
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
            print("กำลังลองใช้วิธีสำรอง...")

            try:
                # วิธีสำรอง: ใช้ path ที่กำหนดเอง
                driver_path = input("กรุณาระบุ path ของ ChromeDriver (กด Enter เพื่อใช้ค่าเริ่มต้น './chromedriver'): ")
                if not driver_path:
                    driver_path = "./chromedriver"

                service = Service(driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                return driver
            except Exception as e2:
                print(f"ไม่สามารถตั้งค่า WebDriver ด้วยวิธีสำรองได้: {str(e2)}")
                raise Exception("ไม่สามารถเริ่มต้น Chrome WebDriver ได้")

    def get_password_with_stars(self):
        """รับรหัสผ่านจากผู้ใช้โดยแสดงดอกจัน (*) แทนตัวอักษรที่พิมพ์"""
        try:
            import msvcrt
            import sys

            password = ""
            print("กรุณาใส่รหัสผ่าน: ", end="", flush=True)

            while True:
                try:
                    char = msvcrt.getch().decode('utf-8')
                except:
                    try:
                        char = msvcrt.getch().decode('cp874')  # สำหรับภาษาไทยบน Windows
                    except:
                        char = msvcrt.getch().decode('latin-1')  # fallback

                # กด Enter เพื่อสิ้นสุด
                if char == '\r' or char == '\n':
                    print()
                    break
                # กด Backspace เพื่อลบตัวอักษรสุดท้าย
                elif char == '\b':
                    if len(password) > 0:
                        password = password[:-1]
                        print('\b \b', end="", flush=True)
                # พิมพ์ตัวอักษรปกติ
                else:
                    password += char
                    print('*', end="", flush=True)

            return password
        except:
            # ถ้าเกิดข้อผิดพลาด ให้ใช้ getpass ธรรมดา
            return getpass.getpass("กรุณาใส่รหัสผ่าน: ")

    def login_to_bni(self, username, password):
        """ล็อกอินเข้า BNI Connect"""
        try:
            print("\n🔐 กำลังล็อกอินเข้า BNI Connect...")

            # เข้าสู่หน้าล็อกอิน
            self.driver.get(self.bni_login_url)
            time.sleep(3)

            # กรอกข้อมูลล็อกอิน
            try:
                # ค้นหาฟิลด์ username
                username_field = self.driver.find_element(By.NAME, "username")
                username_field.clear()
                username_field.send_keys(username)

                # ค้นหาฟิลด์ password
                password_field = self.driver.find_element(By.NAME, "password")
                password_field.clear()
                password_field.send_keys(password)

                # คลิกปุ่มล็อกอิน
                try:
                    login_buttons = self.driver.find_elements(By.XPATH, "//button[contains(text(), 'Login') or contains(text(), 'Sign In') or contains(text(), 'Sign-in') or contains(text(), 'เข้าสู่ระบบ')]")
                    if login_buttons:
                        self.driver.execute_script("arguments[0].click();", login_buttons[0])
                    else:
                        submit_buttons = self.driver.find_elements(By.CSS_SELECTOR, "form button[type='submit']")
                        if submit_buttons:
                            self.driver.execute_script("arguments[0].click();", submit_buttons[0])
                        else:
                            password_field.send_keys(Keys.RETURN)
                except Exception as e:
                    print(f"ไม่สามารถคลิกปุ่มล็อกอิน: {str(e)}")
                    password_field.send_keys(Keys.RETURN)

                # รอให้ล็อกอินเสร็จ
                print("⏳ กำลังรอการล็อกอิน...")
                time.sleep(10)

                # ตรวจสอบว่าล็อกอินสำเร็จ
                current_url = self.driver.current_url
                if "login" in current_url:
                    return False, "ล็อกอินไม่สำเร็จ กรุณาตรวจสอบชื่อผู้ใช้และรหัสผ่าน"

                print("✅ ล็อกอินสำเร็จ!")
                return True, "ล็อกอินสำเร็จ"

            except Exception as e:
                return False, f"เกิดข้อผิดพลาดในการล็อกอิน: {str(e)}"

        except Exception as e:
            return False, f"เกิดข้อผิดพลาดในการเข้าสู่หน้าล็อกอิน: {str(e)}"

    def get_tyfcb_received_from_bni(self):
        """ดึงข้อมูล TYFCB Received จาก BNI Connect Dashboard"""
        try:
            print("\n📊 กำลังดึงข้อมูล TYFCB Received...")

            # เข้าสู่หน้า Dashboard
            self.driver.get(self.bni_dashboard_url)
            time.sleep(7)

            # ค้นหาและคลิกที่ Lifetime
            print("🔍 กำลังค้นหา Lifetime section...")
            try:
                # วิธีที่ 1: ค้นหาด้วยข้อความ Lifetime
                lifetime_elements = self.driver.find_elements(By.XPATH, "//p[contains(text(), 'Lifetime')]")
                if lifetime_elements:
                    for element in lifetime_elements:
                        print(f"📍 พบ Lifetime element: {element.text}")
                        try:
                            self.driver.execute_script("arguments[0].click();", element)
                            print("✅ คลิก Lifetime สำเร็จ")
                            break
                        except:
                            print("⚠️ ไม่สามารถคลิกที่ element นี้ ลองต่อไป...")
                else:
                    # วิธีที่ 2: ค้นหาด้วย attribute isbackground
                    lifetime_elements = self.driver.find_elements(By.XPATH, "//*[@isbackground='true']")
                    if lifetime_elements:
                        for element in lifetime_elements:
                            if "Lifetime" in element.text:
                                print(f"📍 พบ Lifetime จาก isbackground: {element.text}")
                                try:
                                    self.driver.execute_script("arguments[0].click();", element)
                                    print("✅ คลิก Lifetime จาก isbackground สำเร็จ")
                                    break
                                except:
                                    print("⚠️ ไม่สามารถคลิกที่ element นี้ ลองต่อไป...")

                # รอให้หน้าเว็บอัปเดต
                time.sleep(5)

            except Exception as e:
                print(f"⚠️ เกิดข้อผิดพลาดในการคลิก Lifetime: {str(e)}")

            # ค้นหา TYFCB Received
            print("🔍 กำลังค้นหาข้อมูล TYFCB Received...")
            tyfcb_received = "ไม่พบข้อมูล TYFCB Received"

            try:
                # วิธีที่ 1: หาจากข้อความ TYFCB และตัวเลขใกล้ๆ
                tyfcb_indicators = self.driver.find_elements(By.XPATH, "//*[contains(text(), 'TYFCB') and contains(text(), 'Received')]")
                if tyfcb_indicators:
                    for indicator in tyfcb_indicators:
                        print(f"📍 พบข้อความเกี่ยวกับ TYFCB Received: {indicator.text}")
                        try:
                            # หาตัวเลขใกล้ๆ
                            parent_div = indicator.find_element(By.XPATH, "./ancestor::div[contains(@class, 'MuiBox-root')][1]")
                            money_spans = parent_div.find_elements(By.XPATH, ".//span[contains(text(), '฿')]")
                            if money_spans:
                                for span in money_spans:
                                    print(f"💰 พบข้อมูลเงิน TYFCB Received: {span.text}")
                                    tyfcb_received = span.text
                                    break
                        except Exception as e:
                            print(f"⚠️ ไม่สามารถหาข้อมูลเงินใกล้ TYFCB Received: {str(e)}")
                            continue

                # วิธีที่ 2: หาจากสัญลักษณ์สกุลเงินบาทโดยตรง
                if tyfcb_received == "ไม่พบข้อมูล TYFCB Received":
                    money_elements = self.driver.find_elements(By.XPATH, "//span[contains(text(), '฿')]")
                    if money_elements:
                        values = [element.text for element in money_elements]
                        print(f"💰 พบข้อมูลเงินทั้งหมด: {values}")
                        if values:
                            tyfcb_received = values[0]  # ใช้ค่าแรกที่พบ

                # ทำความสะอาดข้อมูล
                if tyfcb_received and tyfcb_received != "ไม่พบข้อมูล TYFCB Received":
                    # ลบสัญลักษณ์และเก็บเฉพาะตัวเลขกับจุลภาค
                    cleaned_amount = re.sub(r'[^\d,.]', '', tyfcb_received)
                    self.tyfcb_received = cleaned_amount
                    print(f"✅ TYFCB Received: {self.tyfcb_received}")
                    return True, self.tyfcb_received
                else:
                    return False, "ไม่พบข้อมูล TYFCB Received"

            except Exception as e:
                return False, f"เกิดข้อผิดพลาดในการดึงข้อมูล TYFCB Received: {str(e)}"

        except Exception as e:
            return False, f"เกิดข้อผิดพลาดในการเข้าสู่ Dashboard: {str(e)}"

    def submit_to_google_form(self, tyfcb_amount):
        """กรอกและส่ง Google Form ด้วยข้อมูล TYFCB"""
        try:
            print(f"\n📝 กำลังกรอก Google Form ด้วยยอดธุรกิจ: {tyfcb_amount}")

            # สร้าง URL พร้อม prefill data
            prefill_url = f"{self.form_url}?usp=pp_url&entry.683444359={self.prefill_name}&entry.290745485={tyfcb_amount}"
            print(f"🔗 URL: {prefill_url}")

            # เปิดหน้า form
            self.driver.get(prefill_url)
            print("📄 เปิดหน้า Google Form สำเร็จ")

            # รอให้หน้าเว็บโหลดเสร็จ
            time.sleep(3)

            # บันทึกภาพหน้าจอหน้าแรก
            if not os.getenv('GITHUB_ACTIONS'):
                self.driver.save_screenshot("form_page1.png")
                print("📸 บันทึกภาพหน้าจอหน้าแรกไว้ที่ form_page1.png")

            # หาและคลิกปุ่ม Next เพื่อไปหน้าถัดไป
            print("🔍 กำลังหาปุ่ม Next...")
            try:
                # วิธีที่ 1: หาด้วย jsname="OCpkoe"
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
                print(f"🔤 ข้อความบนปุ่ม Next: '{next_button.text}'")
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

            # บันทึกภาพหน้าจอหลังกรอกข้อมูล
            if not os.getenv('GITHUB_ACTIONS'):
                self.driver.save_screenshot("form_filled_page2.png")
                print("📸 บันทึกภาพหน้าจอหลังกรอกข้อมูลไว้ที่ form_filled_page2.png")

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
                print(f"🔤 ข้อความบนปุ่ม Submit: '{submit_button.text}'")

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
            print(f"❌ เกิดข้อผิดพลาดในการกรอก Google Form: {e}")
            return False

    def run_integrated_automation(self, username, password):
        """ฟังก์ชันหลักสำหรับรันการทำงานแบบรวม"""
        success_steps = []

        try:
            print("🤖 เริ่มต้น BNI Integrated Automation")
            print("=" * 70)

            # ขั้นตอนที่ 1: ตั้งค่า WebDriver
            print("\n🚀 ขั้นตอนที่ 1: ตั้งค่า WebDriver...")
            self.driver = self.setup_driver()
            success_steps.append("Setup WebDriver")

            # ขั้นตอนที่ 2: ล็อกอินเข้า BNI Connect
            print("\n🔐 ขั้นตอนที่ 2: ล็อกอินเข้า BNI Connect...")
            login_success, login_message = self.login_to_bni(username, password)
            if not login_success:
                return False, f"ล็อกอินไม่สำเร็จ: {login_message}", success_steps
            success_steps.append("Login to BNI Connect")

            # ขั้นตอนที่ 3: ดึงข้อมูล TYFCB Received
            print("\n📊 ขั้นตอนที่ 3: ดึงข้อมูล TYFCB Received...")
            tyfcb_success, tyfcb_result = self.get_tyfcb_received_from_bni()
            if not tyfcb_success:
                return False, f"ไม่สามารถดึงข้อมูล TYFCB Received: {tyfcb_result}", success_steps
            success_steps.append(f"Get TYFCB Received: {tyfcb_result}")

            # ขั้นตอนที่ 4: กรอกและส่ง Google Form
            print(f"\n📝 ขั้นตอนที่ 4: กรอกและส่ง Google Form...")
            form_success = self.submit_to_google_form(tyfcb_result)
            if not form_success:
                return False, "ไม่สามารถกรอกหรือส่ง Google Form ได้", success_steps
            success_steps.append("Submit Google Form")

            return True, "การทำงานทั้งหมดเสร็จสิ้นสำเร็จ", success_steps

        except Exception as e:
            return False, f"เกิดข้อผิดพลาดในการทำงาน: {str(e)}", success_steps

        finally:
            # ปิด WebDriver
            if self.driver:
                if not os.getenv('GITHUB_ACTIONS'):
                    try:
                        self.driver.save_screenshot("final_screen.png")
                        print("📸 บันทึกภาพหน้าจอสุดท้ายไว้ที่ final_screen.png")
                    except:
                        pass

                print("🔒 ปิด WebDriver")
                self.driver.quit()


def main():
    """ฟังก์ชันหลัก"""
    print("🤖 BNI Integrated Automation - รวมการดึงข้อมูล TYFCB และกรอก Google Form")
    print("=" * 80)

    # รองรับ environment variables สำหรับ GitHub Actions
    username = os.getenv('BNI_USERNAME')
    password = os.getenv('BNI_PASSWORD')

    if not username:
        username = input("กรุณาใส่ชื่อผู้ใช้หรืออีเมล BNI Connect: ")

    if not password:
        # รับรหัสผ่านโดยแสดงดอกจันหรือใช้ getpass
        automation = BNIIntegratedAutomation()
        password = automation.get_password_with_stars()
    else:
        print("ใช้รหัสผ่านจาก environment variable")

    print(f"\n👤 ผู้ใช้: {username}")
    print("🔐 รหัสผ่าน: ********")

    print("\n🚀 กำลังดำเนินการ... โปรดรอสักครู่")

    # รันการทำงานแบบรวม
    automation = BNIIntegratedAutomation()
    success, message, completed_steps = automation.run_integrated_automation(username, password)

    # แสดงผลลัพธ์
    print("\n" + "=" * 70)
    print("📋 สรุปผลการทำงาน:")
    print("=" * 70)

    print("✅ ขั้นตอนที่ทำสำเร็จ:")
    for i, step in enumerate(completed_steps, 1):
        print(f"   {i}. {step}")

    if success:
        print(f"\n🎉 {message}")
        print(f"💰 TYFCB Received ที่ดึงได้: {automation.tyfcb_received if automation.tyfcb_received else 'N/A'}")
        print("📝 Google Form ถูกส่งเรียบร้อยแล้ว")
    else:
        print(f"\n❌ {message}")
        print(f"📊 ขั้นตอนที่ทำสำเร็จ: {len(completed_steps)}/4")

    print("=" * 70)
    print("🏁 โปรแกรมทำงานเสร็จสิ้น")

    # ตรวจสอบว่าทำงานใน automated environment หรือไม่
    if not os.getenv('GITHUB_ACTIONS') and not os.getenv('CI'):
        print("\nกด Enter เพื่อออกจากโปรแกรม...")
        input()

    return success


if __name__ == "__main__":
    main()