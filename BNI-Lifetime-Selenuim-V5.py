# V.5 เพิ่ม Google Sheets API integration สำหรับบันทึกข้อมูล TYFCB อัตโนมัติ
# -*- coding: utf-8 -*-
import time
import os
import getpass
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
import csv
import json
import re

# Google Sheets API imports
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    print("Google Sheets API ไม่พร้อมใช้งาน กรุณาติดตั้ง: pip install gspread google-auth")
    GOOGLE_SHEETS_AVAILABLE = False

def setup_google_sheets():
    """
    ตั้งค่าการเชื่อมต่อ Google Sheets API
    """
    if not GOOGLE_SHEETS_AVAILABLE:
        print("Google Sheets API ไม่พร้อมใช้งาน")
        return None

    try:
        # ลองใช้ environment variable ก่อน (สำหรับ GitHub Actions)
        credentials_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
        if credentials_json:
            credentials_info = json.loads(credentials_json)
            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"
            ]
            credentials = Credentials.from_service_account_info(credentials_info, scopes=scope)
        else:
            # ถ้าไม่มี env variable ให้ใช้ไฟล์ local
            credentials_file = "google-sheets-credentials.json"
            if not os.path.exists(credentials_file):
                print(f"ไม่พบไฟล์ {credentials_file} กรุณาวางไฟล์ credentials ไว้ในโฟลเดอร์เดียวกัน")
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

def save_to_google_sheet(tyfcb_received, tyfcb_given_data=None):
    """
    บันทึกข้อมูล TYFCB ลง Google Sheets

    Parameters:
    -----------
    tyfcb_received : str
        ยอดเงิน TYFCB Received
    tyfcb_given_data : dict
        ข้อมูล TYFCB Given (optional)
    """
    try:
        client = setup_google_sheets()
        if not client:
            return False

        # ชื่อ Google Sheet (สามารถเปลี่ยนได้)
        sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'BNI TYFCB Data')

        try:
            spreadsheet = client.open(sheet_name)
        except gspread.SpreadsheetNotFound:
            print(f"ไม่พบ Google Sheet ชื่อ '{sheet_name}' กรุณาสร้างก่อน")
            return False

        # ใช้ worksheet แรก หรือสร้างใหม่
        try:
            worksheet = spreadsheet.sheet1
        except:
            worksheet = spreadsheet.add_worksheet(title="TYFCB Data", rows="1000", cols="10")

        # ตรวจสอบว่ามี header หรือไม่
        try:
            headers = worksheet.row_values(1)
            if not headers:
                # สร้าง header
                header_row = [
                    'Timestamp',
                    'TYFCB Received',
                    'Running User',
                    'Chapter',
                    'Total Given Amount',
                    'Records Count'
                ]
                worksheet.insert_row(header_row, 1)
        except:
            pass

        # เตรียมข้อมูลที่จะบันทึก - ใช้ serial number สำหรับ Google Sheets
        from datetime import datetime
        import time

        # ใช้ Unix timestamp แล้วแปลงเป็น Google Sheets serial number
        # Google Sheets ใช้ serial date จาก 1899-12-30 เป็น day 1
        now = datetime.now()

        # คำนวณ serial number สำหรับ Google Sheets
        # Google Sheets epoch: December 30, 1899
        sheets_epoch = datetime(1899, 12, 30)
        days_since_epoch = (now - sheets_epoch).total_seconds() / (24 * 60 * 60)

        # ใช้ serial number แทน string
        timestamp = days_since_epoch

        # ดึงค่าตัวเลขจาก TYFCB Received และแปลงเป็นตัวเลข
        tyfcb_amount_clean = re.sub(r'[^\d,.]', '', tyfcb_received) if tyfcb_received else '0'
        try:
            # แปลงเป็นตัวเลข เพื่อป้องกัน apostrophe prefix
            tyfcb_amount = float(tyfcb_amount_clean.replace(',', ''))
            print(f"💰 แปลง TYFCB Received: '{tyfcb_received}' → {tyfcb_amount}")
        except ValueError:
            print(f"⚠️  ไม่สามารถแปลง TYFCB Received เป็นตัวเลข: '{tyfcb_received}' - ใช้เป็น string")
            tyfcb_amount = str(tyfcb_amount_clean)

        # แปลง Total Given Amount เป็นตัวเลขเช่นกัน
        total_amount = ''
        if tyfcb_given_data and tyfcb_given_data.get('total_amount'):
            total_amount_clean = re.sub(r'[^\d,.]', '', str(tyfcb_given_data['total_amount']))
            try:
                total_amount = float(total_amount_clean.replace(',', '')) if total_amount_clean else 0
                print(f"💰 แปลง Total Given Amount: '{tyfcb_given_data['total_amount']}' → {total_amount}")
            except ValueError:
                print(f"⚠️  ไม่สามารถแปลง Total Given Amount เป็นตัวเลข - ใช้เป็น string")
                total_amount = str(total_amount_clean)

        row_data = [
            timestamp,                  # Google Sheets serial number
            tyfcb_amount,              # number
            tyfcb_given_data.get('running_user', '') if tyfcb_given_data else '',  # string
            tyfcb_given_data.get('chapter', '') if tyfcb_given_data else '',       # string
            total_amount,              # number
            len(tyfcb_given_data.get('report_data', [])) if tyfcb_given_data else 0  # number
        ]

        # เพิ่มข้อมูลใหม่
        worksheet.append_row(row_data)

        # หา row ที่เพิ่มไป
        last_row_num = len(worksheet.get_all_values())

        # Format timestamp cell ให้เป็น datetime format
        try:
            cell_range = f'A{last_row_num}'
            worksheet.format(cell_range, {
                'numberFormat': {
                    'type': 'DATE_TIME',
                    'pattern': 'mm/dd/yyyy hh:mm:ss'
                }
            })
            print(f"✅ Format timestamp cell {cell_range} เป็น datetime สำเร็จ")
        except Exception as format_error:
            print(f"⚠️  ไม่สามารถ format timestamp cell: {format_error}")

        # ตรวจสอบว่า cell เป็น datetime หรือไม่
        try:
            # อ่านค่ากลับมาเพื่อยืนยันว่าเป็น datetime
            cell_value = worksheet.acell(f'A{last_row_num}').value
            print(f"🔍 ค่าที่บันทึกใน cell A{last_row_num}: '{cell_value}' (type: {type(cell_value).__name__})")
        except Exception as check_error:
            print(f"⚠️  ไม่สามารถตรวจสอบค่า cell: {check_error}")

        print(f"✅ บันทึกข้อมูลลง Google Sheets สำเร็จ")
        print(f"   Timestamp: {timestamp:.6f} (Google Sheets serial number)")
        print(f"   TYFCB Received: {tyfcb_amount}")

        return True

    except Exception as e:
        print(f"❌ ไม่สามารถบันทึกลง Google Sheets: {str(e)}")
        return False

def setup_driver():
    """
    ตั้งค่า WebDriver สำหรับ Chrome
    """
    # ตั้งค่า Chrome options
    chrome_options = Options()

    # ตรวจสอบว่าทำงานใน GitHub Actions หรือไม่
    if os.getenv('GITHUB_ACTIONS'):
        # ตั้งค่าสำหรับ GitHub Actions (headless mode)
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
        # ตั้งค่าสำหรับรันในเครื่องส่วนตัว
        chrome_options.add_argument("--start-maximized")  # เปิดแบบเต็มหน้าจอ

    chrome_options.add_argument("--disable-notifications")  # ปิดการแจ้งเตือน
    # อัปเดต User-Agent เพื่อให้เหมือนเบราว์เซอร์ทั่วไป
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    # เพิ่มการตั้งค่าอื่นๆ ที่อาจจำเป็น
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # เพิ่ม experimental options เพื่อซ่อนการใช้ automation
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # ติดตั้ง WebDriver อัตโนมัติ
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"ไม่สามารถตั้งค่า WebDriver ได้: {str(e)}")
        print("กำลังลองใช้วิธีสำรอง...")
        
        # วิธีสำรอง: ใช้ path ที่กำหนดเอง
        try:
            # ผู้ใช้อาจต้องดาวน์โหลด ChromeDriver เองและระบุ path
            driver_path = input("กรุณาระบุ path ของ ChromeDriver (กด Enter เพื่อใช้ค่าเริ่มต้น './chromedriver'): ")
            if not driver_path:
                driver_path = "./chromedriver"
            
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            return driver
        except Exception as e2:
            print(f"ไม่สามารถตั้งค่า WebDriver ด้วยวิธีสำรองได้: {str(e2)}")
            raise Exception("ไม่สามารถเริ่มต้น Chrome WebDriver ได้ โปรดตรวจสอบการติดตั้ง Chrome และ ChromeDriver")

def get_tyfcb_given_report_data(driver):
    """
    ดึงข้อมูลจากรายงาน TYFCB Given Report ที่อยู่ใน iframe ซ้อน
    """
    report_data = {
        "running_user": "",
        "run_at": "",
        "chapter": "",
        "report_data": [],
        "total_amount": ""
    }
    
    try:
        # บันทึกภาพหน้าจอก่อนเข้า iframe
        # driver.save_screenshot("before_iframe.png") # Disabled file save
        
        # หา iframe ทั้งหมดในหน้า
        all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"พบ iframe ทั้งหมด {len(all_iframes)} อัน")
        
        if all_iframes:
            # สลับไปยัง iframe แรก
            driver.switch_to.frame(all_iframes[0])
            print("สลับไปยัง iframe แรกสำเร็จ")
            
            # บันทึกโค้ด HTML ของ iframe หลัก
            # with open("iframe_main_content.html", "w", encoding="utf-8") as f:
            #     f.write(driver.page_source)
            # print("บันทึกโค้ด HTML ของ iframe หลักไว้ที่ iframe_main_content.html สำหรับตรวจสอบ") # Disabled file save
            
            # ตรวจสอบว่ามี iframe ซ้อนหรือไม่
            inner_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            print(f"พบ iframe ซ้อนทั้งหมด {len(inner_iframes)} อัน")
            
            if inner_iframes:
                # สลับไปยัง iframe ซ้อน
                driver.switch_to.frame(inner_iframes[0])
                print("สลับไปยัง iframe ซ้อนสำเร็จ")
                
                # บันทึกโค้ด HTML ของ iframe ซ้อน
                # with open("iframe_nested_content.html", "w", encoding="utf-8") as f:
                #     f.write(driver.page_source)
                # print("บันทึกโค้ด HTML ของ iframe ซ้อนไว้ที่ iframe_nested_content.html สำหรับตรวจสอบ") # Disabled file save
                
                # ตรวจสอบโครงสร้าง HTML
                page_source = driver.page_source
                print(f"ได้โค้ด HTML ของ iframe ซ้อน ความยาว: {len(page_source)} ตัวอักษร")
                
                # ดึงข้อมูล Running User จาก iframe ซ้อน
                try:
                    # วิธีที่ 1: หาจาก reporttoolbar
                    running_user_elements = driver.find_elements(By.XPATH, "//div[text()='Running User']")
                    if running_user_elements:
                        # หาค่าที่อยู่ใน div params_1
                        params_elements = driver.find_elements(By.XPATH, "//div[@id='params_1']")
                        if params_elements:
                            report_data["running_user"] = params_elements[0].text
                            print(f"Running User: {report_data['running_user']}")
                except Exception as e:
                    print(f"ไม่สามารถดึงข้อมูล Running User: {str(e)}")
                
                # ดึง Run At จาก iframe ซ้อน
                try:
                    # วิธีที่ 1: หาจาก reporttoolbar
                    run_at_elements = driver.find_elements(By.XPATH, "//div[text()='Run At']")
                    if run_at_elements:
                        # หาค่าที่อยู่ใน div params_2
                        params_elements = driver.find_elements(By.XPATH, "//div[@id='params_2']")
                        if params_elements:
                            report_data["run_at"] = params_elements[0].text
                            print(f"Run At: {report_data['run_at']}")
                except Exception as e:
                    print(f"ไม่สามารถดึงข้อมูล Run At: {str(e)}")
                
                # ดึง Chapter จาก iframe ซ้อน
                try:
                    # วิธีที่ 1: หาจาก reporttoolbar
                    chapter_elements = driver.find_elements(By.XPATH, "//div[text()='Chapter']")
                    if chapter_elements:
                        # หาค่าที่อยู่ใน div params_5
                        params_elements = driver.find_elements(By.XPATH, "//div[@id='params_5']")
                        if params_elements:
                            report_data["chapter"] = params_elements[0].text
                            print(f"Chapter: {report_data['chapter']}")
                except Exception as e:
                    print(f"ไม่สามารถดึงข้อมูล Chapter: {str(e)}")
                
                # ดึงข้อมูลตาราง
                try:
                    # วิธีที่ 1: หาตารางที่มี ID __bookmark_3
                    tables = driver.find_elements(By.ID, "__bookmark_3")
                    
                    if not tables:
                        # วิธีที่ 2: หาตารางทั้งหมดและเลือกตารางที่ดูเหมือนจะมีข้อมูล
                        tables = driver.find_elements(By.TAG_NAME, "table")
                        print(f"พบตารางทั้งหมด {len(tables)} ตาราง")
                        
                        # กรองเฉพาะตารางที่มีแถวมากกว่า 1 แถว
                        data_tables = []
                        for table in tables:
                            rows = table.find_elements(By.TAG_NAME, "tr")
                            if len(rows) > 1:
                                data_tables.append(table)
                        
                        if data_tables:
                            # เลือกตารางที่มีแถวมากที่สุด (น่าจะเป็นตารางข้อมูลหลัก)
                            table = max(data_tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")))
                        else:
                            print("ไม่พบตารางที่มีข้อมูล")
                            table = None
                    else:
                        table = tables[0]
                    
                    if table:
                        print(f"พบตารางข้อมูล")
                        
                        # ดึงแถวทั้งหมดจากตาราง
                        rows = table.find_elements(By.TAG_NAME, "tr")
                        print(f"พบแถวทั้งหมด {len(rows)} แถว")
                        
                        # ข้ามแถวแรก (เป็นหัวตาราง)
                        header_row = None
                        data_rows = []
                        total_row = None
                        
                        for i, row in enumerate(rows):
                            # ตรวจสอบว่าเป็นแถวหัวตาราง
                            if i == 0 or row.find_elements(By.TAG_NAME, "th"):
                                header_row = row
                                continue
                            
                            # ตรวจสอบว่าเป็นแถวรวม
                            if "total_row" in row.get_attribute("id") or "Total" in row.text:
                                total_row = row
                            else:
                                data_rows.append(row)
                        
                        # ดึงข้อมูลจากแถวข้อมูล
                        for row in data_rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 6:  # ตรวจสอบว่ามีคอลัมน์ครบตามที่คาดหวัง
                                row_data = {
                                    "date": cells[0].text.strip(),
                                    "thank_you_to": cells[1].text.strip(),
                                    "amount": cells[2].text.strip(),
                                    "new_repeat": cells[3].text.strip() if len(cells) > 3 else "",
                                    "inside_outside": cells[4].text.strip() if len(cells) > 4 else "",
                                    "comments": cells[5].text.strip() if len(cells) > 5 else "",
                                    "status": cells[6].text.strip() if len(cells) > 6 else ""
                                }
                                report_data["report_data"].append(row_data)
                        
                        print(f"ดึงข้อมูลตารางสำเร็จ: พบ {len(report_data['report_data'])} รายการ")
                        
                        # ดึงข้อมูลแถวรวม
                        if total_row:
                            total_cells = total_row.find_elements(By.TAG_NAME, "td")
                            if len(total_cells) > 2:
                                report_data["total_amount"] = total_cells[2].text.strip()
                                print(f"Total Amount: {report_data['total_amount']}")
                        
                        # บันทึกภาพหน้าจอของตาราง
                        # table.screenshot("table_screenshot.png") # Disabled file save
                        print("ข้ามการบันทึกภาพตารางเพื่อลดการสร้างไฟล์ที่ไม่จำเป็น")
                except Exception as e:
                    print(f"ไม่สามารถดึงข้อมูลตาราง: {str(e)}")
                    # บันทึกหน้าจอเพื่อตรวจสอบ
                    # driver.save_screenshot("table_error.png") # Disabled file save
                
                # สลับกลับไปยัง iframe หลัก
                driver.switch_to.default_content()
                driver.switch_to.frame(all_iframes[0])
                print("สลับกลับไปยัง iframe หลักสำเร็จ")
            else:
                print("ไม่พบ iframe ซ้อน ลองดึงข้อมูลจาก iframe หลัก...")
                
                # ดึงข้อมูลจาก iframe หลักโดยตรง (ถ้าไม่มี iframe ซ้อน)
                # ใส่โค้ดสำหรับดึงข้อมูลจาก iframe หลักที่นี่...
            
            # สลับกลับมาที่ main content
            driver.switch_to.default_content()
            print("สลับกลับไปยัง main content สำเร็จ")
            
        else:
            print("ไม่พบ iframe ใดๆ ในหน้า")
            
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการดึงข้อมูลรายงาน TYFCB Given: {str(e)}")
        # พยายามกลับมาที่ main content
        try:
            driver.switch_to.default_content()
        except:
            pass
    
    return report_data

def export_tyfcb_given_report(driver):
    """
    คลิกปุ่ม Export เพื่อดาวน์โหลดรายงาน TYFCB Given เป็นไฟล์ Excel
    """
    try:
        # บันทึกภาพหน้าจอก่อนเข้า iframe
        # driver.save_screenshot("before_export.png") # Disabled file save
        
        # สลับไปยัง iframe หลัก
        iframe_main = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )
        driver.switch_to.frame(iframe_main)
        print("สลับไปยัง iframe หลักสำเร็จ")
        
        # ตรวจสอบ iframe ซ้อน
        inner_iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if inner_iframes:
            driver.switch_to.frame(inner_iframes[0])
            print("สลับไปยัง iframe ซ้อนสำเร็จ")
        
        # ค้นหาปุ่ม Export
        try:
            # วิธีที่ 1: หาปุ่มที่มีข้อความแน่นอนว่า "Export" (ไม่ใช่ "Export without Headers")
            export_buttons = driver.find_elements(By.XPATH, "//a[text()='Export' and not(contains(text(), 'without'))]")
            if export_buttons:
                driver.execute_script("arguments[0].click();", export_buttons[0])
                print("คลิกปุ่ม Export สำเร็จ (วิธีที่ 1)")
            else:
                # วิธีที่ 2: หาลิงก์ทั้งหมดและตรวจสอบข้อความ
                all_links = driver.find_elements(By.TAG_NAME, "a")
                for link in all_links:
                    if link.text == "Export" and "without" not in link.text:
                        driver.execute_script("arguments[0].click();", link)
                        print("คลิกปุ่ม Export สำเร็จ (วิธีที่ 2)")
                        break
        except Exception as e:
            print(f"ไม่สามารถคลิกปุ่ม Export: {str(e)}")
        
        # รอให้ดาวน์โหลดเริ่มต้น
        time.sleep(5)
        
        # สลับกลับไปยัง content หลัก
        driver.switch_to.default_content()
        print("สลับกลับไปยัง content หลักสำเร็จ")
        
        return True
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการ Export รายงาน: {str(e)}")
    
    # กรณีมีข้อผิดพลาด ให้พยายามกลับไปยัง content หลัก
    try:
        driver.switch_to.default_content()
    except:
        pass
    
    return False

def login_and_get_tyfcb(username, password):
    """
    ล็อกอินและดึงข้อมูล TYFCB Received และ TYFCB Given
    """
    driver = None
    try:
        print("\nกำลังเริ่มต้น WebDriver...")
        driver = setup_driver()
        
        # 1. เข้าสู่หน้าล็อกอิน
        print("\nกำลังเข้าสู่หน้าล็อกอิน...")
        driver.get("https://www.bniconnectglobal.com/login")
        time.sleep(3)
        
        # 2. กรอกข้อมูลล็อกอิน
        print("\nกำลังกรอกข้อมูลล็อกอิน...")
        try:
            # ค้นหาฟิลด์ username
            username_field = driver.find_element(By.NAME, "username")
            username_field.clear()
            username_field.send_keys(username)
            
            # ค้นหาฟิลด์ password
            password_field = driver.find_element(By.NAME, "password")
            password_field.clear()
            password_field.send_keys(password)
            
            # คลิกปุ่มล็อกอิน - ลองหลายวิธี
            try:
                # หาปุ่มที่มีข้อความเกี่ยวกับล็อกอิน
                login_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login') or contains(text(), 'Sign In') or contains(text(), 'Sign-in') or contains(text(), 'เข้าสู่ระบบ')]")
                if login_buttons:
                    # ใช้ JavaScript คลิก
                    driver.execute_script("arguments[0].click();", login_buttons[0])
                else:
                    # ลองค้นหาปุ่ม submit ในฟอร์ม
                    submit_buttons = driver.find_elements(By.CSS_SELECTOR, "form button[type='submit']")
                    if submit_buttons:
                        driver.execute_script("arguments[0].click();", submit_buttons[0])
                    else:
                        # ใช้วิธีส่งคีย์ Enter ที่ฟิลด์ password
                        password_field.send_keys(Keys.RETURN)
            except Exception as e:
                print(f"ไม่สามารถคลิกปุ่มล็อกอิน: {str(e)}")
                # ลองกด Enter ที่ฟิลด์ password
                password_field.send_keys(Keys.RETURN)
            
            # รอให้ล็อกอินเสร็จและเปลี่ยนเส้นทาง
            print("\nกำลังรอการล็อกอินและเปลี่ยนเส้นทาง...")
            time.sleep(10)
            
            # ตรวจสอบว่าล็อกอินสำเร็จโดยดูที่ URL
            current_url = driver.current_url
            if "login" in current_url:
                print("\nล็อกอินไม่สำเร็จ ยังอยู่ที่หน้าล็อกอิน")
                if driver:
                    driver.quit()
                return False, "ล็อกอินไม่สำเร็จ กรุณาตรวจสอบชื่อผู้ใช้และรหัสผ่าน", None
            
        except Exception as e:
            print(f"\nเกิดข้อผิดพลาดในการล็อกอิน: {str(e)}")
            if driver:
                driver.quit()
            return False, f"เกิดข้อผิดพลาดในการล็อกอิน: {str(e)}", None
        
        # 3. เข้าสู่หน้า Dashboard
        print("\nล็อกอินสำเร็จ! กำลังเข้าสู่หน้า Dashboard...")
        driver.get("https://www.bniconnectglobal.com/web/dashboard")
        time.sleep(7)
        
        # 4. ค้นหาและคลิกที่ Lifetime
        print("\nกำลังค้นหาและคลิกที่ Lifetime...")
        try:
            # วิธีที่ 1: ค้นหาด้วยข้อความ Lifetime
            lifetime_elements = driver.find_elements(By.XPATH, "//p[contains(text(), 'Lifetime')]")
            if lifetime_elements:
                for element in lifetime_elements:
                    print(f"พบ Lifetime element: {element.text}")
                    try:
                        driver.execute_script("arguments[0].click();", element)
                        print("คลิก Lifetime สำเร็จ")
                        break
                    except:
                        print("ไม่สามารถคลิกที่ element นี้ ลองต่อไป...")
            else:
                print("ไม่พบข้อความ Lifetime")
                
                # วิธีที่ 2: ค้นหาด้วย attribute isbackground
                lifetime_elements = driver.find_elements(By.XPATH, "//*[@isbackground='true']")
                if lifetime_elements:
                    for element in lifetime_elements:
                        if "Lifetime" in element.text:
                            print(f"พบ Lifetime จาก isbackground: {element.text}")
                            try:
                                driver.execute_script("arguments[0].click();", element)
                                print("คลิก Lifetime จาก isbackground สำเร็จ")
                                break
                            except:
                                print("ไม่สามารถคลิกที่ element นี้ ลองต่อไป...")
                else:
                    print("ไม่พบ element ที่มี isbackground='true'")
            
            # รอให้หน้าเว็บอัปเดตหลังคลิก Lifetime
            time.sleep(5)
            
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการคลิก Lifetime: {str(e)}")
        
        # 5. ค้นหา TYFCB Received
        print("\nกำลังค้นหาข้อมูล TYFCB Received...")
        tyfcb_received = "ไม่พบข้อมูล TYFCB Received"
        try:
            # ค้นหาด้วยหลายวิธี
            
            # วิธีที่ 1: หาจากข้อความ TYFCB และจากนั้นหาตัวเลขใกล้ๆ
            tyfcb_indicators = driver.find_elements(By.XPATH, "//*[contains(text(), 'TYFCB') and contains(text(), 'Received')]")
            if tyfcb_indicators:
                for indicator in tyfcb_indicators:
                    print(f"พบข้อความเกี่ยวกับ TYFCB Received: {indicator.text}")
                    try:
                        # หาตัวเลขใกล้ๆ
                        parent_div = indicator.find_element(By.XPATH, "./ancestor::div[contains(@class, 'MuiBox-root')][1]")
                        money_spans = parent_div.find_elements(By.XPATH, ".//span[contains(text(), '฿')]")
                        if money_spans:
                            for span in money_spans:
                                print(f"พบข้อมูลเงิน TYFCB Received: {span.text}")
                                tyfcb_received = span.text
                                break
                    except Exception as e:
                        print(f"ไม่สามารถหาข้อมูลเงินใกล้ TYFCB Received: {str(e)}")
                        continue
            
            # วิธีที่ 2: หาจากสัญลักษณ์สกุลเงินบาทโดยตรง
            if tyfcb_received == "ไม่พบข้อมูล TYFCB Received":
                money_elements = driver.find_elements(By.XPATH, "//span[contains(text(), '฿')]")
                if money_elements:
                    values = [element.text for element in money_elements]
                    print(f"พบข้อมูลเงินทั้งหมด: {values}")
                    if values:
                        tyfcb_received = values[0]  # ใช้ค่าแรกที่พบ
            
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการดึงข้อมูล TYFCB Received: {str(e)}")
        
        # 6. คลิกที่ปุ่ม Review ของ TYFCB Given
        print("\nกำลังค้นหาและคลิกที่ปุ่ม Review ของ TYFCB Given...")
        tyfcb_given_report = "ไม่พบข้อมูล TYFCB Given Report"
        try:
            # ค้นหาข้อความ TYFCB Given
            given_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'TYFCB') and contains(text(), 'Given')]")
            if given_elements:
                given_element = given_elements[0]
                print(f"พบข้อความ TYFCB Given: {given_element.text}")
                
                # หาปุ่ม Review ที่อยู่ใกล้ๆ
                try:
                    # วิธีที่ 1: หาจากข้อความ Review โดยตรง
                    review_elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Review')]")
                    if review_elements:
                        for review_element in review_elements:
                            # ตรวจสอบว่าอยู่ใกล้กับ TYFCB Given หรือไม่
                            print(f"พบปุ่ม Review: {review_element.text}")
                            try:
                                driver.execute_script("arguments[0].click();", review_element)
                                print("คลิกปุ่ม Review สำเร็จ")
                                break
                            except Exception as e:
                                print(f"ไม่สามารถคลิกปุ่ม Review นี้: {str(e)}")
                                continue
                    else:
                        # วิธีที่ 2: หาจาก class ตามที่เห็นในรูป
                        review_elements = driver.find_elements(By.CSS_SELECTOR, "div.MuiBox-root.css-13xuqvq")
                        if review_elements:
                            for review_element in review_elements:
                                if "Review" in review_element.text:
                                    print(f"พบปุ่ม Review จาก class: {review_element.text}")
                                    try:
                                        driver.execute_script("arguments[0].click();", review_element)
                                        print("คลิกปุ่ม Review จาก class สำเร็จ")
                                        break
                                    except Exception as e:
                                        print(f"ไม่สามารถคลิกปุ่ม Review จาก class นี้: {str(e)}")
                                        continue
                
                except Exception as e:
                    print(f"เกิดข้อผิดพลาดในการหาปุ่ม Review: {str(e)}")
                
                # รอให้ popup ปรากฏ
                print("\nรอให้ popup ปรากฏ...")
                time.sleep(3)
                
                # คลิกปุ่ม Go เท่านั้น (ไม่มีการกรอกวันที่)
                print("\nกำลังค้นหาและคลิกปุ่ม Go โดยตรง...")
                try:
                    # บันทึกภาพหน้าจอก่อนการคลิกปุ่ม Go
                    # driver.save_screenshot("popup_before_go.png") # Disabled file save
                    
                    # วิธีที่ 1: หาปุ่มที่มีข้อความแน่นอนว่า "Go"
                    go_button = driver.find_element(By.XPATH, "//button[text()='Go']")
                    driver.execute_script("arguments[0].click();", go_button)
                    print("คลิกปุ่ม 'Go' สำเร็จ (วิธีที่ 1)")
                except Exception as e:
                    print(f"ไม่สามารถคลิกปุ่ม Go ด้วยวิธีที่ 1: {str(e)}")
                    try:
                        # วิธีที่ 2: หาปุ่มสีแดงในป็อปอัพ
                        go_buttons = driver.find_elements(By.CSS_SELECTOR, "button.MuiButton-contained")
                        if go_buttons:
                            for button in go_buttons:
                                if button.is_displayed():
                                    driver.execute_script("arguments[0].click();", button)
                                    print(f"คลิกปุ่ม MuiButton-contained ที่พบ: {button.text}")
                                    break
                        else:
                            # วิธีที่ 3: หาปุ่มทั้งหมดและพิมพ์ออกมาเพื่อตรวจสอบ
                            all_buttons = driver.find_elements(By.TAG_NAME, "button")
                            print(f"พบปุ่มทั้งหมด {len(all_buttons)} ปุ่ม")
                            for idx, button in enumerate(all_buttons):
                                if button.is_displayed():
                                    button_text = button.text.strip()
                                    print(f"ปุ่มที่ {idx+1}: '{button_text}'")
                                    if button_text.lower() == "go":
                                        driver.execute_script("arguments[0].click();", button)
                                        print(f"คลิกปุ่ม '{button_text}' สำเร็จ")
                                        break
                            else:
                                # วิธีที่ 4: ใช้ JavaScript โดยตรง
                                try:
                                    driver.execute_script("document.querySelector('button.MuiButton-contained').click();")
                                    print("คลิกปุ่ม MuiButton-contained ด้วย JavaScript โดยตรง")
                                except Exception as js_error:
                                    print(f"ไม่สามารถคลิกปุ่มด้วย JavaScript โดยตรง: {str(js_error)}")
                    except Exception as e2:
                        print(f"เกิดข้อผิดพลาดในการคลิกปุ่ม Go: {str(e2)}")
                
                # รอให้รายงานแสดง
                print("\nรอให้รายงานแสดง...")
                time.sleep(10)

                # ถ่ายภาพหน้าจอของรายงาน
                # driver.save_screenshot("tyfcb_given_report.png") # Disabled file save
                print("ข้ามการบันทึกภาพรายงาน TYFCB Given เพื่อลดการสร้างไฟล์ที่ไม่จำเป็น")

                # ดึงข้อมูลจากรายงาน - ทำก่อนคลิก Export
                tyfcb_given_data = get_tyfcb_given_report_data(driver)

                # ทำการสร้างรายงานสรุปจากข้อมูลที่ดึงได้
                report_summary = f"Running User: {tyfcb_given_data['running_user']}\n"
                report_summary += f"Run At: {tyfcb_given_data['run_at']}\n"
                report_summary += f"Chapter: {tyfcb_given_data['chapter']}\n\n"
                report_summary += "รายการ TYFCB Given:\n"
                report_summary += "-" * 80 + "\n"
                report_summary += "{:<12} {:<25} {:>10} {:<10} {:<15} {:<30}\n".format(
                    "วันที่", "Thank you to", "จำนวนเงิน", "ประเภท", "แหล่งที่มา", "หมายเหตุ")
                report_summary += "-" * 80 + "\n"

                for item in tyfcb_given_data["report_data"]:
                    report_summary += "{:<12} {:<25} {:>10} {:<10} {:<15} {:<30}\n".format(
                        item["date"],
                        item["thank_you_to"][:25],
                        item["amount"],
                        item["new_repeat"],
                        item["inside_outside"],
                        item["comments"][:30]
                    )

                if "total_amount" in tyfcb_given_data and tyfcb_given_data["total_amount"]:
                    report_summary += "-" * 80 + "\n"
                    report_summary += "{:<12} {:<25} {:>10}\n".format(
                        "", "Total", tyfcb_given_data["total_amount"])

                tyfcb_given_report = report_summary

                # บันทึกข้อมูลเป็นไฟล์ CSV - ยกเลิกการบันทึกไฟล์เพื่อลดการสร้างไฟล์ที่ไม่จำเป็น
                # try:
                #     # ใช้ utf-8-sig แทน utf-8 เพื่อให้แสดงภาษาไทยได้ถูกต้อง
                #     with open("tyfcb_given_report.csv", "w", encoding="utf-8-sig", newline='') as f:
                #         writer = csv.writer(f)
                #         # เพิ่มคอลัมน์ Thank you From หลังจาก Date
                #         writer.writerow(["Date", "Thank you From", "Thank you to", "Amount", "New/Repeat", "Inside/Outside", "Comments", "Status"])
                #         for item in tyfcb_given_data["report_data"]:
                #             writer.writerow([
                #                 item["date"],
                #                 tyfcb_given_data['running_user'], # เพิ่มชื่อ Running User
                #                 item["thank_you_to"],
                #                 item["amount"],
                #                 item["new_repeat"],
                #                 item["inside_outside"],
                #                 item["comments"],
                #                 item["status"]
                #             ])
                #     print("บันทึกข้อมูลรายงานเป็นไฟล์ CSV เรียบร้อย")
                # except Exception as e:
                #     print(f"ไม่สามารถบันทึกไฟล์ CSV: {str(e)}")
                print("ข้อมูล CSV พร้อมประมวลผลแล้ว (ไม่มีการบันทึกไฟล์)")

                # คลิกปุ่ม Export เพื่อดาวน์โหลดรายงานเป็นไฟล์ Excel - ทำหลังจากดึงข้อมูลแล้ว
                export_success = export_tyfcb_given_report(driver)
                if export_success:
                    print("ดาวน์โหลดรายงานเป็นไฟล์ Excel สำเร็จ")
                
            else:
                print("ไม่พบข้อความ TYFCB Given")
                
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการคลิกปุ่ม Review: {str(e)}")
        
        # บันทึกข้อมูลลง Google Sheets (ถ้าพร้อมใช้งาน)
        if GOOGLE_SHEETS_AVAILABLE:
            print("\n=== บันทึกข้อมูลลง Google Sheets ===")
            save_to_google_sheet(tyfcb_received, tyfcb_given_data)

        return True, tyfcb_received, tyfcb_given_report
        
    except Exception as e:
        print(f"\nเกิดข้อผิดพลาดในโปรแกรม: {str(e)}")
        if driver:
            # driver.save_screenshot("error.png") # Disabled file save
            pass
        return False, f"เกิดข้อผิดพลาด: {str(e)}", None
        
    finally:
        if driver:
            print("\nกำลังปิดเบราว์เซอร์...")
            try:
                # driver.save_screenshot("final_screen.png") # Disabled file save
                print("ข้ามการบันทึกภาพหน้าจอเพื่อลดการสร้างไฟล์ที่ไม่จำเป็น")
            except:
                pass
            
            print("\nกำลังปิด WebDriver...")
            driver.quit()

def get_password_with_stars():
    """
    รับรหัสผ่านจากผู้ใช้โดยแสดงดอกจัน (*) แทนตัวอักษรที่พิมพ์
    สำหรับ Windows
    """
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

def save_report_to_excel(tyfcb_given_data, filename="tyfcb_given_report_export.xlsx"):
    """
    บันทึกข้อมูลรายงาน TYFCB Given เป็นไฟล์ Excel โดยใช้ pandas และ openpyxl
    """
    try:
        import pandas as pd
        
        if not tyfcb_given_data or not tyfcb_given_data.get("report_data"):
            print("ไม่มีข้อมูลรายงานสำหรับบันทึกลงไฟล์ Excel")
            return False
        
        # สร้าง DataFrame จากข้อมูลรายงาน
        data = []
        for item in tyfcb_given_data["report_data"]:
            data.append({
                "Date": item["date"],
                "Thank you to": item["thank_you_to"],
                "Amount": item["amount"],
                "New/Repeat": item["new_repeat"],
                "Inside/Outside": item["inside_outside"],
                "Comments": item["comments"],
                "Status": item["status"]
            })
        
        df = pd.DataFrame(data)
        
        # เพิ่มข้อมูลสรุป
        info_data = [
            ["Running User", tyfcb_given_data.get("running_user", "")],
            ["Run At", tyfcb_given_data.get("run_at", "")],
            ["Chapter", tyfcb_given_data.get("chapter", "")],
            ["Total Amount", tyfcb_given_data.get("total_amount", "")]
        ]
        info_df = pd.DataFrame(info_data, columns=["Information", "Value"])
        
        # บันทึกลงไฟล์ Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            info_df.to_excel(writer, sheet_name='Report Info', index=False)
            df.to_excel(writer, sheet_name='TYFCB Given Data', index=False)
        
        print(f"บันทึกรายงานเป็นไฟล์ Excel ที่ {filename} เรียบร้อย")
        return True
    
    except ImportError:
        print("ไม่พบโมดูล pandas หรือ openpyxl สำหรับการบันทึกไฟล์ Excel")
        print("คุณสามารถติดตั้งได้โดยใช้คำสั่ง: pip install pandas openpyxl")
        return False
    
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการบันทึกไฟล์ Excel: {str(e)}")
        return False

def main():
    print("โปรแกรมดึงข้อมูล TYFCB Received และ TYFCB Given Report จาก BNI Connect Global")
    print("=" * 70)

    # รองรับ environment variables สำหรับ GitHub Actions
    username = os.getenv('BNI_USERNAME')
    password = os.getenv('BNI_PASSWORD')

    if not username:
        username = input("กรุณาใส่ชื่อผู้ใช้หรืออีเมล: ")

    if not password:
        # รับรหัสผ่านโดยแสดงดอกจันหรือใช้ getpass
        password = get_password_with_stars()
    else:
        print("ใช้รหัสผ่านจาก environment variable")
    
    print("\nกำลังดำเนินการ... โปรดรอสักครู่")
    success, tyfcb_received, tyfcb_given_report = login_and_get_tyfcb(username, password)
    
    print("\n" + "=" * 40)
    if success:
        print("ดึงข้อมูลสำเร็จ!")
        print(f"TYFCB Received (Lifetime): {tyfcb_received}")
        
        print("\nTYFCB Given Report (1 ปีย้อนหลัง):")
        print("-" * 30)
        if tyfcb_given_report and tyfcb_given_report != "ไม่พบข้อมูล TYFCB Given Report":
            # จำกัดความยาวของรายงานที่แสดงในคอนโซล
            if len(tyfcb_given_report) > 2000:
                print(tyfcb_given_report[:2000] + "...\n(ข้อมูลเพิ่มเติมดูได้จากการแสดงผลบนหน้าจอ)")
            else:
                print(tyfcb_given_report)
            print("\nข้อมูลรายงานพร้อมใช้งานแล้ว (ไม่มีการบันทึกไฟล์)")
        else:
            print(tyfcb_given_report)
    else:
        print("ไม่สามารถดึงข้อมูลได้")
        print(f"สาเหตุ: {tyfcb_received}")
    print("=" * 40)
    
    print("\nโปรแกรมทำงานเสร็จสิ้นแล้ว ไม่มีการบันทึกไฟล์เพิ่มเติม")

    # ตรวจสอบว่าทำงานใน automated environment หรือไม่
    if not os.getenv('GITHUB_ACTIONS') and not os.getenv('CI'):
        print("\nกด Enter เพื่อออกจากโปรแกรม...")
        input()

if __name__ == "__main__":
     main()    
