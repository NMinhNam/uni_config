import os
from datetime import timedelta

class Config:
    # Cấu hình cơ bản
    # SECRET_KEY = os.environ.get('SECRET_KEY') or 'you-will-never-guess'
    UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
    SESSION_COOKIE_SECURE = True
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=1)  # 1 phút
    SESSION_REFRESH_EACH_REQUEST = True  # Refresh session mỗi request

    # Cấu hình thêm (nếu cần)
    # ADDITIONAL_CONFIG_KEY = 'additional_value' 