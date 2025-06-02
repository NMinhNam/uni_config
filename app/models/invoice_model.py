import os
import time
from uuid import uuid4
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import random
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
import base64
import re
from requests.exceptions import JSONDecodeError
from config import Config
import tempfile
import shutil
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.chrome import ChromeDriverManager

class InvoiceModel:
    def __init__(self):
        self.base_url = Config.BASE_URL
        self.api_url = Config.API_URL
        self.invoice_url_template = Config.INVOICE_URL_TEMPLATE
        self.download_root = Config.DOWNLOAD_ROOT
        self.excel_filename = Config.EXCEL_FILENAME
        self.driver = None
        self.wait = None
        self.session_cookies = None
        self.last_captcha_url = None

    def initialize_driver(self):
        if self.driver is None:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")
            
            # Create a unique temporary directory for user data
            temp_dir = tempfile.mkdtemp()
            options.add_argument(f"--user-data-dir={temp_dir}")
            
            try:
                # Check environment
                if os.environ.get('RENDER'):  # If running on Render
                    # Try different possible Chrome paths
                    chrome_paths = [
                        '/usr/bin/google-chrome',
                        '/usr/bin/google-chrome-stable',
                        '/usr/bin/chrome',
                        '/usr/bin/chrome-browser'
                    ]
                    
                    chrome_path = None
                    for path in chrome_paths:
                        if os.path.exists(path):
                            chrome_path = path
                            break
                    
                    if not chrome_path:
                        raise Exception("Chrome binary not found in any of the expected locations")
                    
                    options.binary_location = chrome_path
                    service = webdriver.chrome.service.Service(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(service=service, options=options)
                else:  # If running locally
                    self.driver = webdriver.Chrome(options=options)
                
                self.wait = WebDriverWait(self.driver, 15)
                self.driver.get(f"{self.base_url}/dang-nhap")
                time.sleep(random.uniform(1, 2))
            except Exception as e:
                # Clean up temp directory if driver initialization fails
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                print(f"Error initializing driver: {str(e)}")
                raise e

    def get_captcha_image(self):
        try:
            self.initialize_driver()
            captcha_path = self.download_captcha_from_element()
            if not captcha_path:
                return False, "Không thể tải ảnh CAPTCHA"
            
            self.last_captcha_url = f"/static/{os.path.basename(captcha_path)}"
            self.session_cookies = self.driver.get_cookies()
            print(f"New CAPTCHA URL: {self.last_captcha_url}")
            return True, self.last_captcha_url
        except Exception as e:
            self.close_driver()
            return False, f"Lỗi khi tải CAPTCHA: {str(e)}"

    def download_captcha_from_element(self):
        try:
            captcha_image = self.wait.until(EC.presence_of_element_located((By.ID, "CaptchaImage")))
            
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(random.uniform(0.5, 1.5))
            
            self.driver.execute_script("document.getElementById('CaptchaImage').style.width = '150px';")
            self.driver.execute_script("document.getElementById('CaptchaImage').style.height = '50px';")
            time.sleep(1)
            
            size = captcha_image.size
            if size['width'] < 50 or size['height'] < 50:
                print("Kích thước ảnh CAPTCHA không hợp lệ")
                return None
            
            static_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'static')
            os.makedirs(static_dir, exist_ok=True)
            captcha_filename = f"captcha_{uuid4()}.png"
            captcha_path = os.path.join(static_dir, captcha_filename)
            
            if not captcha_image.screenshot(captcha_path):
                print("Không thể lưu screenshot CAPTCHA")
                return None
            
            if not os.path.exists(captcha_path):
                print("File CAPTCHA không được tạo")
                return None
                
            with Image.open(captcha_path) as img:
                if img.size[0] < 50 or img.size[1] < 50:
                    print("Kích thước ảnh CAPTCHA sau khi lưu không hợp lệ")
                    try:
                        os.remove(captcha_path)
                    except:
                        pass
                    return None
            
            return captcha_path
        except Exception as e:
            print(f"Lỗi khi chụp ảnh CAPTCHA: {e}")
            return None

    def validate_input(self, username, password, from_date, to_date, company):
        try:
            if not all([username, password, from_date, to_date, company]):
                return False, "Tất cả các trường (username, password, from_date, to_date, company) đều bắt buộc"
            
            date_pattern = r"^\d{2}/\d{2}/\d{4}$"
            if not re.match(date_pattern, from_date) or not re.match(date_pattern, to_date):
                return False, "Ngày phải có định dạng DD/MM/YYYY"
            
            from_date_obj = pd.to_datetime(from_date, format="%d/%m/%Y")
            to_date_obj = pd.to_datetime(to_date, format="%d/%m/%Y")
            
            if from_date_obj > to_date_obj:
                return False, "Từ ngày không được lớn hơn đến ngày"
            
            return True, (from_date_obj, to_date_obj)
        except ValueError as e:
            return False, f"Lỗi định dạng ngày: {str(e)}"

    def login(self, username, password, captcha_code, max_attempts=3):
        for attempt in range(max_attempts):
            try:
                username_field = self.wait.until(EC.presence_of_element_located((By.ID, "form-username")))
                username_field.clear()
                username_field.send_keys(username)
                time.sleep(random.uniform(0.5, 1))
                
                password_field = self.wait.until(EC.presence_of_element_located((By.ID, "form-password")))
                password_field.clear()
                password_field.send_keys(password)
                time.sleep(random.uniform(0.5, 1))
                
                captcha_field = self.wait.until(EC.presence_of_element_located((By.ID, "CaptchaInputText")))
                captcha_field.clear()
                captcha_field.send_keys(captcha_code)
                time.sleep(random.uniform(0.5, 1))
                
                login_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @class='btn']")))
                login_button.click()
                
                time.sleep(3)
                
                error_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.red")
                captcha_error = False
                for element in error_elements:
                    if "Vui lòng nhập đúng mã captcha" in element.text:
                        captcha_error = True
                        break
                
                if captcha_error:
                    print("CAPTCHA sai, đóng trình duyệt và mở lại...")
                    self.close_driver()
                    self.initialize_driver()
                    success, new_captcha_url = self.get_captcha_image()
                    if success:
                        return False, "Mã CAPTCHA sai", None, new_captcha_url
                    return False, "Mã CAPTCHA sai, không thể làm mới CAPTCHA", None, None
                
                try:
                    user_info = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".pull-left.info.width190px"))
                    )
                    if user_info:
                        cookies = self.driver.get_cookies()
                        cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}
                        
                        cookie_order = [
                            "ASP.NET_SessionId",
                            "__RequestVerificationToken",
                            "_frontend",
                            "TPHCM_LOGIN_TOKEN",
                            "TS013c59b1"
                        ]
                        cookie_parts = []
                        for key in cookie_order:
                            if key in cookie_dict:
                                cookie_parts.append(f"{key}={cookie_dict[key]}")
                        cookie_string = "; ".join(cookie_parts)
                        
                        token = None
                        html = self.driver.page_source
                        soup = BeautifulSoup(html, "html.parser")
                        token_input = soup.find("input", {"name": "__RequestVerificationToken", "type": "hidden"})
                        if token_input and token_input.get("value"):
                            token = token_input["value"]
                        
                        if not token:
                            for cookie in cookies:
                                if cookie['name'] == "__RequestVerificationToken":
                                    token = cookie['value']
                                    break
                        
                        if not token:
                            token = self.driver.execute_script("return localStorage.getItem('__RequestVerificationToken');") or \
                                    self.driver.execute_script("return localStorage.getItem('TPHCM_LOGIN_TOKEN');")
                        
                        if not token:
                            for cookie in cookies:
                                if cookie['name'] == "TPHCM_LOGIN_TOKEN":
                                    token = cookie['value']
                                    break
                        
                        if not token:
                            return False, "Không tìm thấy token", None, None
                        
                        print("Đăng nhập thành công!")
                        return True, cookie_string, token, self.driver
                except:
                    continue

            except Exception as e:
                return False, f"Lỗi đăng nhập: {str(e)}", None, None
        
        return False, "Không thể đăng nhập sau tất cả các lần thử", None, None

    def cleanup_static_files(self):
        try:
            static_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'static')
            if os.path.exists(static_dir):
                for old_file in os.listdir(static_dir):
                    if old_file.startswith("captcha_") and old_file.endswith(".png"):
                        file_path = os.path.join(static_dir, old_file)
                        try:
                            os.remove(file_path)
                            print(f"Đã xóa file: {file_path}")
                        except Exception as e:
                            print(f"Không thể xóa file {file_path}: {str(e)}")
        except Exception as e:
            print(f"Lỗi khi dọn dẹp thư mục static: {str(e)}")

    def process_invoices(self, username, password, captcha_code, from_date, to_date, company):
        driver = None
        temp_dir = None
        try:
            is_valid, result = self.validate_input(username, password, from_date, to_date, company)
            if not is_valid:
                return {"status": "error", "message": result, "new_captcha_url": None}
            
            current_year = pd.Timestamp.now().strftime("%Y")
            current_month = pd.Timestamp.now().strftime("%B")
            company = re.sub(r'[<>:"/\\|?*]', '_', company)
            download_folder = os.path.join(self.download_root, current_year, current_month, company)
            os.makedirs(download_folder, exist_ok=True)
            
            success, message, token, driver = self.login(username, password, captcha_code)
            
            if not success:
                return {"status": "captcha_error" if "CAPTCHA sai" in message else "error", 
                        "message": message, 
                        "new_captcha_url": message if "CAPTCHA sai" in message else None}
            
            headers = {
                "__requestverificationtoken": token,
                "accept": "*/*",
                "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                "x-requested-with": "XMLHttpRequest",
                "origin": self.base_url,
                "referer": f"{self.base_url}/",
                "Cookie": f"SERVERID=web1; {message}",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
            }
            
            data = {
                "TU_NGAY": from_date,
                "DEN_NGAY": to_date,
                "SO_TK": "",
                "SO_THONG_BAO": "",
                "MA_DV": "",
                "LOAI_DULIEU": "",
                "LOAI_THANH_TOAN": "",
                "TRANG_THAI_TOKHAI": "-3",
                "LOAI_TOKHAI": "0",
                "NHOM_LOAIHINH": "",
                "MA_MIEN_GIAM": "",
                "pageNum": "1"
            }
            
            res = requests.post(self.api_url, headers=headers, data=data, timeout=10)
            if res.status_code != 200:
                return {"status": "error", "message": f"API request failed with status {res.status_code}", "new_captcha_url": None}
            
            try:
                response_json = res.json()
                invoices = response_json.get("DANHSACH", [])
            except JSONDecodeError:
                return {"status": "error", "message": "JSONDecodeError", "new_captcha_url": None}
            
            self.close_driver()  # Close driver before opening new one
            
            options = webdriver.ChromeOptions()
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")
            
            # Create a unique temporary directory for user data
            temp_dir = tempfile.mkdtemp()
            options.add_argument(f"--user-data-dir={temp_dir}")
            
            try:
                # Check environment
                if os.environ.get('RENDER'):  # If running on Render
                    options.binary_location = '/usr/bin/google-chrome'
                    service = webdriver.chrome.service.Service(ChromeDriverManager().install())
                    service = webdriver.chrome.service.Service(ChromeDriverManager().install())
                    driver = webdriver.Chrome(service=service, options=options)
                else:  # If running locally
                    driver = webdriver.Chrome(options=options)
            except Exception as e:
                # Clean up temp directory if driver initialization fails
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                print(f"Error initializing driver: {str(e)}")
                raise e
            
            records = []
            for i, invoice in enumerate(invoices, 1):
                so_tk = invoice.get("SO_TKHQ", "")
                mhd = invoice.get("INVOICELINK", "")
                url = self.invoice_url_template.format(mhd)
                
                try:
                    driver.get(url)
                    time.sleep(3)
                    html = driver.page_source
                    soup = BeautifulSoup(html, "html.parser")
                    
                    invoice_no = soup.find("span", style=re.compile("font-size:20px"))
                    invoice_no = invoice_no.text.strip() if invoice_no else ""
                    
                    td_candidates = soup.find_all(lambda tag: tag.name == "td" and "text-right" in tag.get("class", []) and "boderTop" in tag.get("class", []))
                    amount = ""
                    for td in td_candidates:
                        div = td.find("div", class_="dott")
                        if div:
                            amount = div.get_text(strip=True).replace("\n", "")
                            break
                    
                    date_text = soup.find(string=re.compile(r"TP\.HCM, ngày \d{2} tháng \d{2} năm \d{4}"))
                    invoice_date, month, day = "", "XX", "XX"
                    if date_text:
                        match = re.search(r"ngày (\d{2}) tháng (\d{2}) năm (\d{4})", date_text)
                        if match:
                            day, month, year = match.groups()
                            invoice_date = f"{day}/{month}/{year}"
                    
                    short_date = f"{month}.{day}"
                    file_name = f"{short_date}_{so_tk}_{invoice_no}.pdf"
                    file_path = os.path.join(download_folder, file_name)
                    
                    result = driver.execute_cdp_cmd("Page.printToPDF", {
                        "printBackground": True,
                        "paperWidth": 8.27,
                        "paperHeight": 11.69,
                        "scale": 0.98,
                        "marginTop": 0.5,
                        "marginBottom": 0,
                        "marginLeft": 0,
                        "marginRight": 0
                    })
                    with open(file_path, "wb") as f:
                        f.write(base64.b64decode(result['data']))
                    
                    records.append({
                        "Số TK": so_tk,
                        "Số Hoá đơn": invoice_no,
                        "Số Tiền (VND)": amount,
                        "Ngày hoá đơn": invoice_date,
                        "Tên file hoá đơn": file_name
                    })
                
                except Exception as e:
                    print(f"Lỗi khi xử lý hóa đơn {i}: {str(e)}")
                    continue
            
            if records:
                df = pd.DataFrame(records)
                excel_path = os.path.join(download_folder, self.excel_filename)
                df.to_excel(excel_path, index=False)
            
            return {"status": "success", "message": f"Đã tải {len(records)} hóa đơn và lưu tại {download_folder}", "new_captcha_url": None}
        
        except Exception as e:
            return {"status": "error", "message": f"Lỗi khi xử lý hóa đơn: {str(e)}", "new_captcha_url": None}
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            self.cleanup_static_files()

    def close_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
            self.wait = None
            self.session_cookies = None