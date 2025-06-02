from flask import session
from app.services.sharepoint_service import SharePointService
from datetime import datetime

class AuthService:
    def __init__(self):
        self.sharepoint_service = SharePointService()

    def login(self, username, password):
        if not username or not password:
            return False, 'Vui lòng nhập đầy đủ thông tin đăng nhập'

        try:
            # Authenticate with SharePoint
            if self.sharepoint_service.authenticate_user(username, password):
                # Store user session with timestamp
                session.permanent = True
                session['user'] = {
                    'username': username,
                    'authenticated': True,
                    'login_time': datetime.now().timestamp()
                }
                return True, 'Đăng nhập thành công'
            return False, 'Đăng nhập thất bại! Vui lòng kiểm tra lại thông tin đăng nhập'
        except Exception as e:
            # Log error without sensitive information
            print("Authentication error occurred")
            return False, 'Đăng nhập thất bại! Vui lòng kiểm tra lại thông tin đăng nhập'

    def logout(self):
        session.clear()
        return True

    def is_authenticated(self):
        if 'user' not in session or not session['user'].get('authenticated'):
            return False
        return True
