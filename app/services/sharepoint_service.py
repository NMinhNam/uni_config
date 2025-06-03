from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import os
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
