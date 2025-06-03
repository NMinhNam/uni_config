from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import pandas as pd
import re
import os
from config import Config

class DeclarationService:
    def __init__(self):
        self.site_key = 'test_data'
        self.list_key = 'declaration'
        self.site_url = Config.get_sharepoint_site_url(self.site_key)
        self.list_name = Config.get_sharepoint_list_name(self.site_key, self.list_key)
        self.client_options = [
            "INZI VINA", "ALS ĐỒNG NAI", "AJ SOLUTIONS", "CJU", "YN VIETNAM", "SEOGWANG", "N unquestionably", "VINA DUKE", "HWA SUNG", "GST", "EMAX", "YKK", "KOOKIL DYEING", "UNITEX", "THUÝ MỸ TƯ",
            "CHUNSHIN PRECISION VINA", "AINA VIỆT NAM", "COUNT VINA", "DAE MYUNG", "Daeduck Mesh Vina",
            "Damy Vina", "DEA SIN VINA", "ĐIỆN TỬ ND", "DIOPS VINA", "DS VINA", "ENS FOAM", "GLORY V.T",
            "GLORYTEX VINA", "GST", "HANIL POLYTEC VINA", "HWA SUNG", "I HOA ECO", "IN ẤN SONG KHÁNH",
            "JEONG", "K.A.S.S.A", "KOOKIL DYEING", "KORVIET VINA", "LAON ĐỨC HÒA VIỆT NAM",
            "NHỰA JINGGUANG ĐỒNG NAI", "NIFCO", "OK SUNG VINA", "PHÚ LỘC L.A", "Quang Nghi",
            "PRINT DESIGN VNB", "Phản Quang Việt - Hàn", "SAMSUNG HCMC CE COMPLEX", "SAM JIN TEXTILE",
            "SHINWON Việt Nam", "SJ GLOMAX", "SMART FIBER VĨNH LONG VINA", "SPEED VINA", "SUNGHO VINA",
            "U.K Vina", "Taedoo Vina", "UNITEX", "VIET SMART MOBILITY"
        ]

    def validate_input(self, from_date, to_date):
        try:
            # Chuyển đổi từ dd/mm/yyyy sang datetime
            from_date_dt = datetime.strptime(from_date, '%d/%m/%Y')
            to_date_dt = datetime.strptime(to_date, '%d/%m/%Y')
            
            # Đặt thời gian về 00:00:00 cho ngày bắt đầu và 23:59:59 cho ngày kết thúc
            from_date_dt = from_date_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            to_date_dt = to_date_dt.replace(hour=23, minute=59, second=59, microsecond=999999)
            
            if from_date_dt > to_date_dt:
                return {'error': 'Ngày bắt đầu không thể lớn hơn ngày kết thúc.'}
            return {'success': True, 'from_date': from_date_dt, 'to_date': to_date_dt}
        except ValueError as e:
            print(f"Debug - Date parsing error: {str(e)}")
            return {'error': 'Định dạng ngày không hợp lệ. Vui lòng sử dụng định dạng dd/mm/yyyy.'}
        
    def connect_sharepoint(self, username, password):
        try:
            ctx_auth = AuthenticationContext(self.site_url)
            if ctx_auth.acquire_token_for_user(username, password):
                return {'success': True, 'context': ClientContext(self.site_url, ctx_auth)}
            return {'error': 'Không thể xác thực với SharePoint'}
        except Exception as e:
            return {'error': f'Lỗi kết nối SharePoint: {str(e)}'}