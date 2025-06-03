import os
from datetime import timedelta

class Config:
    # Cấu hình cơ bản
    # SECRET_KEY = os.environ.get('SECRET_KEY') or 'you-will-never-guess'
    UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    # Cấu hình SharePoint
    SHAREPOINT_CONFIG = {
        'clearance_flow': {
            'site_url': 'https://uniconsulting079.sharepoint.com/sites/ClearanceFlowAutomation',
            'lists': {
                'debit': 'Debit list',
                'declaration': 'Debit list'
            }
        },
        'test_data': {
            'site_url': 'https://uniconsulting079.sharepoint.com/sites/Test_data',
            'lists': {
                'debit': 'Debit list',
                'declaration': 'Debit list'
            }
        }
    }

    @staticmethod
    def get_sharepoint_site_url(site_key):
        if site_key in Config.SHAREPOINT_CONFIG:
            return Config.SHAREPOINT_CONFIG[site_key]['site_url']
        raise ValueError(f"Không tìm thấy cấu hình cho site key: {site_key}")

    @staticmethod
    def get_sharepoint_list_name(site_key, list_key):
        if site_key in Config.SHAREPOINT_CONFIG and list_key in Config.SHAREPOINT_CONFIG[site_key]['lists']:
            return Config.SHAREPOINT_CONFIG[site_key]['lists'][list_key]
        raise ValueError(f"Không tìm thấy cấu hình cho site key: {site_key} hoặc list key: {list_key}")

    @classmethod
    def get_all_sharepoint_sites(cls):
        """Lấy danh sách tất cả các site SharePoint"""
        return list(cls.SHAREPOINT_CONFIG.keys())

    @classmethod
    def get_all_sharepoint_lists(cls, site_key):
        """Lấy danh sách tất cả các list trong một site"""
        if site_key in cls.SHAREPOINT_CONFIG:
            return list(cls.SHAREPOINT_CONFIG[site_key]['lists'].keys())
        return []

    # Cấu hình URL
    BASE_URL = "https://thuphihatang.tphcm.gov.vn"
    API_URL = "https://thuphihatang.tphcm.gov.vn/DToKhaiNP/GetList/"
    INVOICE_URL_TEMPLATE = "https://thuphihatang.tphcm.gov.vn:8081/Viewer/HoaDonViewer.aspx?mhd={}"

    # Cấu hình thư mục
    DOWNLOAD_ROOT = "DOWNLOAD"
    EXCEL_FILENAME = "Invoice_List.xlsx"

    # Cấu hình database (nếu cần)
    # SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///app.db'
    # SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Cấu hình bảo mật
    SESSION_COOKIE_SECURE = False  # Chỉ bật True khi chạy trên HTTPS
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=30)
    SESSION_REFRESH_EACH_REQUEST = True  # Refresh session mỗi request

    # Cấu hình thêm (nếu cần)
    # ADDITIONAL_CONFIG_KEY = 'additional_value' 