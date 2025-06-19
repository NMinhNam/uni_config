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
import base64
import re
from requests.exceptions import JSONDecodeError
from config import Config
import tempfile
import shutil

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

    def _is_production(self):
        """Detect if running in production environment"""
        return (
            os.path.exists('/.dockerenv') or
            os.environ.get('RENDER') or
            os.environ.get('RAILWAY_ENVIRONMENT') or
            os.environ.get('VERCEL')
        )

    def _get_chrome_options(self):
        """Get Chrome options for current environment"""
        options = webdriver.ChromeOptions()
        
        # Essential options for all environments
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
        
        if self._is_production():
            # Production-specific options
            options.add_argument("--disable-software-rasterizer")
            options.add_argument("--disable-background-timer-throttling")
            options.add_argument("--disable-backgrounding-occluded-windows")
            options.add_argument("--disable-renderer-backgrounding")
            options.add_argument("--remote-debugging-port=9222")
            options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
            
            # Set Chromium binary for production
            chrome_bin = os.environ.get('CHROME_BIN', '/usr/bin/chromium')
            if os.path.exists(chrome_bin):
                options.binary_location = chrome_bin
                print(f"✅ Using Chromium: {chrome_bin}")
        else:
            # Local development
            options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
        
        return options

    def _create_driver(self):
        """Create Chrome driver - Production ready"""
        options = self._get_chrome_options()
        temp_dir = tempfile.mkdtemp()
        options.add_argument(f"--user-data-dir={temp_dir}")
        
        try:
            # ✅ STEP 1: Try system ChromeDriver first
            system_paths = [
                '/usr/bin/chromedriver',
                '/usr/local/bin/chromedriver',
                os.environ.get('CHROMEDRIVER_PATH', '')
            ]
            
            for path in system_paths:
                if path and os.path.exists(path) and os.access(path, os.X_OK):
                    service = webdriver.chrome.service.Service(path)
                    driver = webdriver.Chrome(service=service, options=options)
                    print(f"✅ Using system ChromeDriver: {path}")
                    return driver, temp_dir
            
            # ✅ STEP 2: Try webdriver-manager
            from webdriver_manager.chrome import ChromeDriverManager
            
            manager = ChromeDriverManager()
            driver_path = manager.install()
            
            # ✅ STEP 3: Validate and find correct executable
            if driver_path and 'chromedriver' in driver_path:
                # Find actual chromedriver binary
                search_dirs = [
                    os.path.dirname(driver_path),
                    os.path.join(os.path.dirname(driver_path), 'chromedriver-linux64'),
                    os.path.join(os.path.dirname(driver_path), '..', 'chromedriver-linux64')
                ]
                
                for search_dir in search_dirs:
                    if os.path.exists(search_dir):
                        chromedriver_path = os.path.join(search_dir, 'chromedriver')
                        if os.path.exists(chromedriver_path) and os.access(chromedriver_path, os.X_OK):
                            service = webdriver.chrome.service.Service(chromedriver_path)
                            driver = webdriver.Chrome(service=service, options=options)
                            print(f"✅ Using webdriver-manager ChromeDriver: {chromedriver_path}")
                            return driver, temp_dir
            
            # ✅ STEP 4: Last resort - try downloaded path directly
            if os.path.exists(driver_path):
                service = webdriver.chrome.service.Service(driver_path)
                driver = webdriver.Chrome(service=service, options=options)
                print(f"✅ Using downloaded ChromeDriver: {driver_path}")
                return driver, temp_dir
            
            raise Exception("Cannot find working ChromeDriver")
            
        except Exception as e:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
            print(f"❌ ChromeDriver error: {str(e)}")
            raise Exception(f"ChromeDriver setup failed: {str(e)}")

    def initialize_driver(self):
        """Initialize Chrome driver"""
        if self.driver is None:
            try:
                self.driver, self.temp_dir = self._create_driver()
                self.wait = WebDriverWait(self.driver, 15)
                
                print(f"🌐 Navigating to: {self.base_url}/dang-nhap")
                self.driver.get(f"{self.base_url}/dang-nhap")
                time.sleep(random.uniform(1, 2))
                print("✅ Driver initialized successfully")
                
            except Exception as e:
                print(f"❌ Error initializing driver: {str(e)}")
                raise e

    def download_captcha(self, captcha_url, driver):
        """Download CAPTCHA using requests - Always return base64"""
        try:
            # Create session with cookies from driver
            session = requests.Session()
            
            # Get cookies from Selenium WebDriver
            selenium_cookies = driver.get_cookies()
            
            # Add cookies to session
            for cookie in selenium_cookies:
                session.cookies.set(cookie['name'], cookie['value'])
            
            # Headers
            headers = {
                "user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
            }
            
            # Download CAPTCHA
            print(f"🔄 Downloading CAPTCHA from: {captcha_url}")
            response = session.get(captcha_url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                # ✅ ALWAYS convert to base64 - No file saving
                img_base64 = base64.b64encode(response.content).decode('utf-8')
                data_url = f"data:image/png;base64,{img_base64}"
                
                print(f"✅ CAPTCHA downloaded, size: {len(response.content)} bytes")
                print(f"✅ Base64 data URL created, length: {len(data_url)} chars")
                return data_url
            else:
                print(f"❌ HTTP Error: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"❌ Error in download_captcha: {str(e)}")
            return None

    def get_captcha_image(self):
        """Get CAPTCHA image for user input - Base64 only"""
        try:
            self.initialize_driver()
            
            # Wait for CAPTCHA element
            captcha_element = self.wait.until(EC.presence_of_element_located((By.ID, "CaptchaImage")))
            print("✅ CAPTCHA element found")
            
            # ✅ ONLY USE BASE64 METHOD - No file system
            try:
                # Method 1: Screenshot as base64 (most reliable)
                img_base64 = captcha_element.screenshot_as_base64
                data_url = f"data:image/png;base64,{img_base64}"
                
                self.last_captcha_url = data_url
                self.session_cookies = self.driver.get_cookies()
                
                print("✅ CAPTCHA obtained via base64 screenshot")
                print(f"🔍 Data URL length: {len(data_url)} chars")
                return True, self.last_captcha_url
                
            except Exception as e:
                print(f"❌ Base64 screenshot failed: {str(e)}")
                
                # Method 2: Download via requests + convert to base64
                try:
                    captcha_url = captcha_element.get_attribute('src')
                    if captcha_url:
                        print(f"🔍 Trying download method: {captcha_url}")
                        data_url = self.download_captcha(captcha_url, self.driver)
                        if data_url:
                            self.last_captcha_url = data_url
                            self.session_cookies = self.driver.get_cookies()
                            print("✅ CAPTCHA obtained via download + base64")
                            return True, self.last_captcha_url
                except Exception as e2:
                    print(f"❌ Download method failed: {str(e2)}")
                
                return False, f"Không thể tải CAPTCHA: {str(e)}"
            
        except Exception as e:
            self.close_driver()
            print(f"❌ Error getting CAPTCHA: {str(e)}")
            return False, f"Lỗi khi tải CAPTCHA: {str(e)}"

    def _get_download_directory(self, company):
        """Get download directory for invoices"""
        current_year = pd.Timestamp.now().strftime("%Y")
        current_month = pd.Timestamp.now().strftime("%B")
        safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company))
        
        # Try multiple base directories
        base_candidates = [
            self.download_root,  # From config
            '/app/downloads',    # Docker standard
            './downloads',       # Local fallback
            '/tmp/downloads'     # Last resort
        ]
        
        for base_dir in base_candidates:
            try:
                download_folder = os.path.join(base_dir, current_year, current_month, safe_company)
                os.makedirs(download_folder, exist_ok=True)
                
                # Test write permission
                test_file = os.path.join(download_folder, 'test.tmp')
                with open(test_file, 'w') as f:
                    f.write('test')
                os.remove(test_file)
                
                print(f"✅ Using download directory: {download_folder}")
                return download_folder
                
            except Exception as e:
                print(f"⚠️ Cannot use {base_dir}: {str(e)}")
                continue
        
        raise Exception("Cannot create writable download directory")

    def validate_input(self, username, password, from_date, to_date, company):
        """Validate input parameters"""
        try:
            if not all([username, password, from_date, to_date, company]):
                return False, "Tất cả các trường đều bắt buộc"
            
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
        """Login to the website"""
        for attempt in range(max_attempts):
            try:
                # Fill username
                username_field = self.wait.until(EC.presence_of_element_located((By.ID, "form-username")))
                username_field.clear()
                username_field.send_keys(username)
                time.sleep(random.uniform(0.5, 1))
                
                # Fill password
                password_field = self.wait.until(EC.presence_of_element_located((By.ID, "form-password")))
                password_field.clear()
                password_field.send_keys(password)
                time.sleep(random.uniform(0.5, 1))
                
                # Fill CAPTCHA
                captcha_field = self.wait.until(EC.presence_of_element_located((By.ID, "CaptchaInputText")))
                captcha_field.clear()
                captcha_field.send_keys(captcha_code)
                time.sleep(random.uniform(0.5, 1))
                
                # Submit form
                login_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @class='btn']")))
                login_button.click()
                time.sleep(3)
                
                # Check for CAPTCHA error
                error_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.red")
                captcha_error = False
                for element in error_elements:
                    if "Vui lòng nhập đúng mã captcha" in element.text:
                        captcha_error = True
                        break
                
                if captcha_error:
                    print("❌ CAPTCHA incorrect, refreshing...")
                    self.close_driver()
                    self.initialize_driver()
                    success, new_captcha_url = self.get_captcha_image()
                    if success:
                        return False, "Mã CAPTCHA sai", None, new_captcha_url
                    return False, "Mã CAPTCHA sai, không thể làm mới CAPTCHA", None, None
                
                # Check login success
                try:
                    user_info = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".pull-left.info.width190px"))
                    )
                    if user_info:
                        # Get cookies and token
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
                        
                        # Get token
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
                            return False, "Không tìm thấy token", None, None
                        
                        print("✅ Login successful!")
                        return True, cookie_string, token, self.driver
                except:
                    continue

            except Exception as e:
                print(f"❌ Login error: {str(e)}")
                return False, f"Lỗi đăng nhập: {str(e)}", None, None
        
        return False, "Không thể đăng nhập sau tất cả các lần thử", None, None

    def _get_sharepoint_service(self):
        """Get SharePoint service for invoice uploads"""
        try:
            from app.services.sharepoint_service import SharePointService
            sp_service = SharePointService('clearance_flow')  # Dùng site ClearanceFlowAutomation
            return sp_service
        except Exception as e:
            print(f"⚠️ SharePoint service not available: {str(e)}")
            return None

    def _get_fallback_directory(self, company):
        """Create fallback local directory if SharePoint fails"""
        try:
            current_year = pd.Timestamp.now().strftime("%Y")
            current_month = pd.Timestamp.now().strftime("%B")
            safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company))
            
            # Try temp directory first
            base_candidates = [
                tempfile.gettempdir(),
                '/tmp',
                './temp_invoices'
            ]
            
            for base_dir in base_candidates:
                try:
                    fallback_folder = os.path.join(base_dir, 'invoice_fallback', current_year, current_month, safe_company)
                    os.makedirs(fallback_folder, exist_ok=True)
                    
                    # Test write permission
                    test_file = os.path.join(fallback_folder, 'test.tmp')
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                    
                    print(f"✅ Fallback directory: {fallback_folder}")
                    return fallback_folder
                    
                except Exception as e:
                    print(f"⚠️ Cannot use fallback {base_dir}: {str(e)}")
                    continue
            
            return None
        except Exception as e:
            print(f"❌ Cannot create fallback directory: {str(e)}")
            return None

    def process_invoices(self, username, password, captcha_code, from_date, to_date, company):
        """Process and upload invoices to SharePoint"""
        pdf_driver = None
        try:
            # Validate input
            is_valid, result = self.validate_input(username, password, from_date, to_date, company)
            if not is_valid:
                return {"status": "error", "message": result, "new_captcha_url": None}
            
            # ✅ SHAREPOINT SETUP
            sp_service = self._get_sharepoint_service()
            use_sharepoint = sp_service is not None
            
            current_year = pd.Timestamp.now().strftime("%Y")
            current_month = pd.Timestamp.now().strftime("%B")
            safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company))
            
            # Fallback directory if SharePoint fails
            fallback_dir = None
            if not use_sharepoint:
                fallback_dir = self._get_fallback_directory(company)
                if not fallback_dir:
                    return {"status": "error", "message": "Không thể tạo thư mục lưu trữ và SharePoint không khả dụng", "new_captcha_url": None}
            
            print(f"📡 SharePoint: {'✅ Available' if use_sharepoint else '❌ Not available'}")
            print(f"📁 Fallback: {'✅ Ready' if fallback_dir else '❌ None'}")
            
            # Login
            success, message, token, driver = self.login(username, password, captcha_code)
            if not success:
                return {
                    "status": "captcha_error" if "CAPTCHA sai" in message else "error", 
                    "message": message, 
                    "new_captcha_url": message if "CAPTCHA sai" in message else None
                }
            
            # Prepare API request
            headers = {
                "__requestverificationtoken": token,
                "accept": "*/*",
                "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                "x-requested-with": "XMLHttpRequest",
                "origin": self.base_url,
                "referer": f"{self.base_url}/",
                "Cookie": f"SERVERID=web1; {message}",
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
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
            
            # Get invoice list
            print("🔍 Fetching invoice list...")
            res = requests.post(self.api_url, headers=headers, data=data, timeout=10)
            if res.status_code != 200:
                return {"status": "error", "message": f"API request failed with status {res.status_code}", "new_captcha_url": None}
            
            try:
                response_json = res.json()
                invoices = response_json.get("DANHSACH", [])
                print(f"✅ Found {len(invoices)} invoices")
            except JSONDecodeError:
                return {"status": "error", "message": "JSONDecodeError", "new_captcha_url": None}
            
            if not invoices:
                return {"status": "success", "message": "Không tìm thấy hóa đơn nào trong khoảng thời gian này", "new_captcha_url": None}
            
            # Close login driver and create new one for PDF downloads
            self.close_driver()
            
            # Create new driver for PDF processing
            print("🔄 Creating new driver for PDF downloads...")
            pdf_driver, temp_dir = self._create_driver()
            
            # ✅ PROCESS EACH INVOICE - SHAREPOINT UPLOAD
            records = []
            upload_results = []
            successful_uploads = 0
            
            for i, invoice in enumerate(invoices, 1):
                print(f"📄 Processing invoice {i}/{len(invoices)}")
                
                so_tk = invoice.get("SO_TKHQ", "")
                mhd = invoice.get("INVOICELINK", "")
                url = self.invoice_url_template.format(mhd)
                
                try:
                    pdf_driver.get(url)
                    time.sleep(3)
                    
                    # Parse invoice details
                    html = pdf_driver.page_source
                    soup = BeautifulSoup(html, "html.parser")
                    
                    # Extract invoice number
                    invoice_no = soup.find("span", style=re.compile("font-size:20px"))
                    invoice_no = invoice_no.text.strip() if invoice_no else ""
                    
                    # Extract amount
                    td_candidates = soup.find_all(lambda tag: tag.name == "td" and "text-right" in tag.get("class", []) and "boderTop" in tag.get("class", []))
                    amount = ""
                    for td in td_candidates:
                        div = td.find("div", class_="dott")
                        if div:
                            amount = div.get_text(strip=True).replace("\n", "")
                            break
                    
                    # Extract date
                    date_text = soup.find(string=re.compile(r"TP\.HCM, ngày \d{2} tháng \d{2} năm \d{4}"))
                    invoice_date, month, day = "", "XX", "XX"
                    if date_text:
                        match = re.search(r"ngày (\d{2}) tháng (\d{2}) năm (\d{4})", date_text)
                        if match:
                            day, month, year = match.groups()
                            invoice_date = f"{day}/{month}/{year}"
                    
                    # Generate PDF in memory
                    short_date = f"{month}.{day}"
                    file_name = f"{short_date}_{so_tk}_{invoice_no}.pdf"
                    
                    # Create PDF using Chrome DevTools Protocol
                    result = pdf_driver.execute_cdp_cmd("Page.printToPDF", {
                        "printBackground": True,
                        "paperWidth": 8.27,
                        "paperHeight": 11.69,
                        "scale": 0.98,
                        "marginTop": 0.5,
                        "marginBottom": 0,
                        "marginLeft": 0,
                        "marginRight": 0
                    })
                    
                    pdf_content = base64.b64decode(result['data'])
                    
                    # ✅ UPLOAD TO SHAREPOINT
                    upload_status = "❌ Failed"
                    
                    if use_sharepoint:
                        try:
                            upload_success, upload_result = sp_service.upload_invoice_pdf(
                                pdf_content, file_name, safe_company, current_year, current_month
                            )
                            
                            if upload_success:
                                print(f"📤 Uploaded to SharePoint: {file_name}")
                                upload_status = "✅ SharePoint"
                                upload_results.append(f"✅ {file_name}")
                                successful_uploads += 1
                            else:
                                print(f"❌ SharePoint upload failed: {upload_result}")
                                upload_status = f"❌ SP Error: {upload_result}"
                                upload_results.append(f"❌ {file_name}: {upload_result}")
                                
                                # Fallback to local
                                if fallback_dir:
                                    file_path = os.path.join(fallback_dir, file_name)
                                    with open(file_path, "wb") as f:
                                        f.write(pdf_content)
                                    upload_status = "⚠️ Local fallback"
                                    print(f"💾 Saved locally as fallback: {file_name}")
                                    
                        except Exception as sp_error:
                            print(f"❌ SharePoint error: {str(sp_error)}")
                            upload_status = f"❌ SP Exception: {str(sp_error)}"
                            
                            # Fallback to local
                            if fallback_dir:
                                file_path = os.path.join(fallback_dir, file_name)
                                with open(file_path, "wb") as f:
                                    f.write(pdf_content)
                                upload_status = "⚠️ Local fallback"
                                print(f"💾 Saved locally as fallback: {file_name}")
                    else:
                        # Local only mode
                        if fallback_dir:
                            file_path = os.path.join(fallback_dir, file_name)
                            with open(file_path, "wb") as f:
                                f.write(pdf_content)
                            upload_status = "⚠️ Local only"
                            print(f"💾 Saved locally: {file_name}")
                    
                    records.append({
                        "Số TK": so_tk,
                        "Số Hoá đơn": invoice_no,
                        "Số Tiền (VND)": amount,
                        "Ngày hoá đơn": invoice_date,
                        "Tên file hoá đơn": file_name,
                        "Upload Status": upload_status
                    })
                    
                except Exception as e:
                    print(f"❌ Error processing invoice {i}: {str(e)}")
                    records.append({
                        "Số TK": so_tk or "Unknown",
                        "Số Hoá đơn": "Error",
                        "Số Tiền (VND)": "Error",
                        "Ngày hoá đơn": "Error",
                        "Tên file hoá đơn": f"Error processing {mhd}",
                        "Upload Status": f"❌ Processing Error: {str(e)}"
                    })
                    continue
            
            # ✅ UPLOAD EXCEL SUMMARY
            excel_status = "❌ Failed"
            sharepoint_url = None
            
            if records:
                df = pd.DataFrame(records)
                excel_filename = f"Invoice_Summary_{current_year}_{current_month}_{safe_company}.xlsx"
                
                if use_sharepoint:
                    try:
                        upload_success, upload_result = sp_service.upload_invoice_excel(
                            df, excel_filename, safe_company, current_year, current_month
                        )
                        if upload_success:
                            print(f"📊 Excel summary uploaded to SharePoint: {excel_filename}")
                            excel_status = "✅ SharePoint"
                            sharepoint_url = sp_service.get_invoice_folder_url(safe_company, current_year, current_month)
                        else:
                            print(f"❌ Excel upload failed: {upload_result}")
                            excel_status = f"❌ Excel Error: {upload_result}"
                    except Exception as excel_error:
                        print(f"❌ Excel SharePoint error: {str(excel_error)}")
                        excel_status = f"❌ Excel Exception: {str(excel_error)}"
                
                # Local fallback for Excel
                if excel_status.startswith("❌") and fallback_dir:
                    try:
                        excel_path = os.path.join(fallback_dir, excel_filename)
                        df.to_excel(excel_path, index=False)
                        excel_status = "⚠️ Local fallback"
                        print(f"📊 Excel saved locally: {excel_path}")
                    except Exception as local_error:
                        excel_status = f"❌ Local save failed: {str(local_error)}"
            
            # ✅ RETURN RESULTS
            total_invoices = len(records)
            
            if use_sharepoint and sharepoint_url and successful_uploads > 0:
                return {
                    "status": "success", 
                    "message": f"✅ Đã upload {successful_uploads}/{total_invoices} hóa đơn lên SharePoint. Excel: {excel_status}", 
                    "sharepoint_url": sharepoint_url,
                    "upload_results": upload_results,
                    "total_invoices": total_invoices,
                    "successful_uploads": successful_uploads,
                    "excel_status": excel_status,
                    "storage_type": "SharePoint",
                    "new_captcha_url": None
                }
            elif fallback_dir:
                return {
                    "status": "success", 
                    "message": f"⚠️ SharePoint không khả dụng. Đã lưu {total_invoices} hóa đơn tại: {fallback_dir}", 
                    "local_path": fallback_dir,
                    "total_invoices": total_invoices,
                    "excel_status": excel_status,
                    "storage_type": "Local",
                    "new_captcha_url": None
                }
            else:
                return {
                    "status": "warning", 
                    "message": f"⚠️ Xử lý {total_invoices} hóa đơn nhưng có vấn đề lưu trữ. Vui lòng kiểm tra logs.", 
                    "total_invoices": total_invoices,
                    "storage_type": "Error",
                    "new_captcha_url": None
                }
        
        except Exception as e:
            print(f"❌ Error in process_invoices: {str(e)}")
            return {"status": "error", "message": f"Lỗi khi xử lý hóa đơn: {str(e)}", "new_captcha_url": None}
        
        finally:
            if pdf_driver:
                try:
                    pdf_driver.quit()
                except:
                    pass

    def close_driver(self):
        """Close driver and cleanup"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
            self.wait = None
            self.session_cookies = None
            
        # Cleanup temp directory
        if hasattr(self, 'temp_dir') and self.temp_dir:
            try:
                shutil.rmtree(self.temp_dir)
            except:
                pass