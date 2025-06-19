from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import os
import io
import pandas as pd
from config import Config
from flask import session

class SharePointService:
    def __init__(self, site_key='clearance_flow'):
        self.site_key = site_key
        self.site_url = Config.get_sharepoint_site_url(site_key)
        if not self.site_url:
            raise ValueError(f"Không tìm thấy cấu hình cho site key: {site_key}")
        self.ctx = None

    def authenticate_from_session(self):
        """Xác thực sử dụng thông tin đăng nhập từ session"""
        if 'user' not in session or not session['user'].get('authenticated'):
            raise Exception("Chưa đăng nhập")
        
        # Lấy thông tin đăng nhập từ session
        username = session['user'].get('username')
        password = session['user'].get('password')
        
        if not username or not password:
            raise Exception("Không tìm thấy thông tin đăng nhập trong session")
            
        return self.authenticate_user(username, password)

    def authenticate_user(self, username, password):
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_user(username, password)
            self.ctx = ClientContext(self.site_url, auth_context)
            
            # Verify authentication by making a simple request
            self.ctx.load(self.ctx.web)
            self.ctx.execute_query()
            return True
        except Exception as e:
            # Log error without sensitive information
            print("SharePoint authentication error occurred")
            return False

    def get_context(self):
        if not self.ctx:
            # Tự động xác thực từ session nếu chưa có context
            self.authenticate_from_session()
        return self.ctx

    def get_list_by_key(self, list_key):
        """Lấy list SharePoint dựa trên list key"""
        list_name = Config.get_sharepoint_list_name(self.site_key, list_key)
        if not list_name:
            raise ValueError(f"Không tìm thấy cấu hình cho list key: {list_key}")
        return self.get_context().web.lists.get_by_title(list_name)

    # ✅ === INVOICE UPLOAD METHODS - FIXED VERSION ===
    
    def create_invoice_folders(self, company, year, month):
        """Tạo folder structure: Invoice_Documents/Year/Month/Company"""
        try:
            ctx = self.get_context()
            web = ctx.web
            
            # Get Documents library
            doc_library = web.lists.get_by_title("Documents")
            root_folder = doc_library.root_folder
            
            # Sanitize company name
            import re
            safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company)).strip()
            if not safe_company:
                safe_company = "Unknown_Company"
            
            # Folder structure
            folder_names = ["Invoice_Documents", str(year), month, safe_company]
            current_folder = root_folder
            folder_path = ""
            
            for folder_name in folder_names:
                folder_path = f"{folder_path}/{folder_name}" if folder_path else folder_name
                
                try:
                    # Try to get existing folder
                    next_folder = current_folder.folders.get_by_url(folder_name)
                    ctx.load(next_folder)
                    ctx.execute_query()
                    current_folder = next_folder
                    print(f"📁 Folder exists: {folder_name}")
                except:
                    # Create new folder
                    current_folder = current_folder.folders.add(folder_name)
                    ctx.execute_query()
                    print(f"📁 Created folder: {folder_name}")
            
            print(f"✅ Folder structure ready: {folder_path}")
            return current_folder, folder_path
            
        except Exception as e:
            print(f"❌ Error creating folders: {str(e)}")
            return None, None

    def upload_file_to_folder(self, file_content, file_name, target_folder):
        """Upload file vào SharePoint folder - SIMPLIFIED METHOD"""
        try:
            ctx = self.get_context()
            
            # ✅ SỬ DỤNG BUILT-IN UPLOAD METHOD - No import needed
            uploaded_file = target_folder.upload_file(file_name, file_content)
            ctx.execute_query()
            
            # Load file properties to get URL
            ctx.load(uploaded_file)
            ctx.execute_query()
            
            print(f"✅ Uploaded: {file_name}")
            return True, uploaded_file.serverRelativeUrl
            
        except Exception as e:
            print(f"❌ Upload failed {file_name}: {str(e)}")
            return False, str(e)

    def upload_invoice_pdf(self, pdf_content, file_name, company, year, month):
        """Upload single PDF invoice"""
        try:
            target_folder, folder_path = self.create_invoice_folders(company, year, month)
            if not target_folder:
                return False, "Cannot create folder structure"
            
            success, result = self.upload_file_to_folder(pdf_content, file_name, target_folder)
            return success, result
            
        except Exception as e:
            print(f"❌ PDF upload error: {str(e)}")
            return False, str(e)

    def upload_invoice_excel(self, excel_df, file_name, company, year, month):
        """Upload Excel summary file"""
        try:
            # Convert DataFrame to Excel bytes
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                excel_df.to_excel(writer, sheet_name='Invoice Summary', index=False)
            excel_content = excel_buffer.getvalue()
            
            # Get target folder
            target_folder, folder_path = self.create_invoice_folders(company, year, month)
            if not target_folder:
                return False, "Cannot create folder structure"
            
            # Upload Excel
            success, result = self.upload_file_to_folder(excel_content, file_name, target_folder)
            return success, result
            
        except Exception as e:
            print(f"❌ Excel upload error: {str(e)}")
            return False, str(e)

    def get_invoice_folder_url(self, company, year, month):
        """Get SharePoint folder URL for user access"""
        import re
        safe_company = re.sub(r'[<>:"/\\|?*]', '_', str(company)).strip()
        if not safe_company:
            safe_company = "Unknown_Company"
        
        return f"{self.site_url}/Shared Documents/Invoice_Documents/{year}/{month}/{safe_company}"

    def upload_invoice_batch(self, invoice_data_list, company, year, month):
        """Upload batch of invoices (PDFs + Excel summary)"""
        try:
            # Create folder structure once
            target_folder, folder_path = self.create_invoice_folders(company, year, month)
            if not target_folder:
                return False, "Cannot create folder structure", []
            
            upload_results = []
            
            # Upload each PDF
            for invoice_data in invoice_data_list:
                pdf_content = invoice_data.get('pdf_content')
                file_name = invoice_data.get('file_name')
                
                if pdf_content and file_name:
                    success, result = self.upload_file_to_folder(pdf_content, file_name, target_folder)
                    status = "✅ Success" if success else f"❌ {result}"
                    upload_results.append({
                        'file_name': file_name,
                        'status': status,
                        'success': success
                    })
            
            return True, "Batch upload completed", upload_results
            
        except Exception as e:
            print(f"❌ Batch upload error: {str(e)}")
            return False, str(e), []
