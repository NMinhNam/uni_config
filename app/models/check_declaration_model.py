from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import pandas as pd
import re
import os
from app.services.check_declaration_service import DeclarationService
from flask import session

class DeclarationModel:
    def __init__(self):
        self.service = DeclarationService()
        self.client_options = self.service.client_options

    def get_clients(self):
        """Trả về danh sách khách hàng"""
        try:
            return {'clients': self.client_options}
        except Exception as e:
            print(f"Error getting clients: {str(e)}")
            return {'error': 'Không thể tải danh sách khách hàng'}

    def clean_cd_no(self, series):
        cleaned = set(
            series.dropna()
            .apply(lambda x: str(int(float(str(x).strip().replace(' ', '')))) if str(x).replace(' ', '').replace('.', '').isdigit() else str(x).strip().replace(' ', ''))
        )
        print(f"Debug - Cleaned cd_no_set: {cleaned}")
        return cleaned
    
    def find_cd_no_column(self, df):
        print(f"Debug - DataFrame columns: {df.columns.tolist()}")
        print(f"Debug - DataFrame head:\n{df.head().to_string()}")
        for idx, row in df.iterrows():
            for col in df.columns:
                cell_value = str(row[col]).strip().upper()
                print(f"Debug - Checking cell [{idx}, {col}]: {cell_value}")
                if cell_value in ["CD NO", "CD NO.", "CD_NO", "CD_NO."]:
                    cd_no_col = df[col].iloc[idx + 1:].reset_index(drop=True)
                    print(f"Debug - Found CD No column at [{idx}, {col}]: {cd_no_col.tolist()}")
                    return cd_no_col
        return None
    
    def validate_input(self, from_date, to_date):
        result = self.service.validate_input(from_date, to_date)
        if 'error' in result:
            return result
        return {
            'success': True,
            'from_date': result['from_date'],
            'to_date': result['to_date']
        }
    
    def process_excel_file(self, file_path):
        try:
            print(f"Debug - Processing Excel file: {file_path}")
            # Đọc file Excel
            df = pd.read_excel(file_path, header=None)
            
            # Tìm cột CD NO
            cd_no_series = self.find_cd_no_column(df)
            if cd_no_series is None:
                return {
                    'error': "Không tìm thấy cột 'CD NO' (hoặc tương tự) trong file Excel."
                }
            
            # Chuyển đổi dữ liệu thành tập hợp số tờ khai
            cd_no_set = self.clean_cd_no(cd_no_series)
            if not cd_no_set:
                return {
                    'error': 'Không có dữ liệu hợp lệ trong cột CD NO.'
                }
            
            return {
                'success': True,
                'data': cd_no_set
            }
            
        except Exception as e:
            print(f"Debug - Error reading Excel file: {str(e)}")
            return {
                'error': f'Lỗi khi đọc file Excel: {str(e)}'
            }
    
    def connect_sharepoint(self):
        try:
            if 'user' not in session or not session['user'].get('authenticated'):
                print("Debug - User not authenticated in model")
                return {'error': 'Vui lòng đăng nhập lại'}
                
            username = session['user'].get('username')
            password = session['user'].get('password')
            
            print(f"Debug - Attempting to connect with username: {username}")
            print(f"Debug - Password exists: {bool(password)}")
            
            if not username or not password:
                print("Debug - Missing credentials in model")
                return {'error': 'Vui lòng đăng nhập lại'}
                
            result = self.service.connect_sharepoint(username, password)
            if 'error' in result:
                print(f"Debug - SharePoint connection error: {result['error']}")
                return result
                
            print("Debug - SharePoint connection successful")
            return {
                'success': True,
                'context': result['context']
            }
        except Exception as e:
            print(f"Debug - Error in connect_sharepoint: {str(e)}")
            return {'error': f'Lỗi kết nối SharePoint: {str(e)}'}
    
    def check_declarations(self, ctx, cd_no_set, from_date, to_date, selected_client):
        try:
            print(f"Debug - Starting check_declarations with from_date={from_date}, to_date={to_date}, selected_client={selected_client}")
            # Lấy danh sách từ SharePoint
            list_obj = ctx.web.lists.get_by_title("Debit list")
            ctx.load(list_obj)
            ctx.execute_query()

            all_items = []
            seen_ids = set()
            while True:
                items = list_obj.items.top(1000).get().execute_query()
                if not items or all(item.properties['Id'] in seen_ids for item in items):
                    break
                for item in items:
                    seen_ids.add(item.properties['Id'])
                    all_items.append(item)

            print(f"Debug - all_items count: {len(all_items)}")
            for item in all_items:
                print(f"Debug - Item: CDNo={item.properties.get('CDNo')}, InvoiceDate={item.properties.get('InvoiceDate')}, Client={item.properties.get('Client')}")

            # Chuẩn hóa số tờ khai từ Excel
            def normalize_cd_no(cd_no):
                try:
                    cleaned = str(cd_no).strip().replace(' ', '')
                    return str(int(float(cleaned))) if cleaned.replace('.', '').isdigit() else cleaned
                except:
                    return str(cd_no).strip()

            excel_cd_no_set = set(normalize_cd_no(cd_no) for cd_no in cd_no_set)
            print(f"Debug - excel_cd_no_set: {excel_cd_no_set}")

            filtered_items = []
            updated_items = []

            # Lọc và kiểm tra các tờ khai
            for item in all_items:
                invoice_date_str = item.properties.get("InvoiceDate", "").strip()
                if not invoice_date_str:
                    print(f"Debug - Missing InvoiceDate for CDNo={item.properties.get('CDNo')}")
                    continue

                try:
                    invoice_date = datetime.strptime(invoice_date_str, "%d/%m/%Y")
                    print(f"Debug - Parsed InvoiceDate for CDNo={item.properties.get('CDNo')}: {invoice_date}")
                except ValueError:
                    print(f"Debug - Invalid date format for CDNo={item.properties.get('CDNo')}: {invoice_date_str}")
                    continue

                client = item.properties.get("Client", "").strip().upper()
                so_tk = normalize_cd_no(item.properties.get("CDNo", ""))

                # Kiểm tra điều kiện thời gian và khách hàng
                client_match = selected_client == "Tất cả khách hàng" or selected_client.upper() == client
                print(f"Debug - Client match for Client={client}, selected_client={selected_client}: {client_match}")
                print(f"Debug - Date check: {from_date} <= {invoice_date} <= {to_date}")
                if from_date <= invoice_date <= to_date and client_match:
                    print(f"Debug - Matching item: CDNo={so_tk}, Client={client}, InvoiceDate={invoice_date_str}")
                    if so_tk not in excel_cd_no_set:
                        # Tờ khai thiếu (không có trong Excel)
                        filtered_items.append({
                            'CDNo': so_tk,
                            'Client': item.properties.get("Client", "N/A"),
                            'InvoiceDate': invoice_date_str
                        })
                    else:
                        # Tờ khai trùng, cập nhật DebitStatus
                        try:
                            item.set_property("DebitStatus", "Debit")
                            item.update()
                            ctx.execute_query()
                            updated_items.append({
                                'CDNo': so_tk,
                                'Client': item.properties.get("Client", "N/A"),
                                'InvoiceDate': invoice_date_str
                            })
                            print(f"Debug - Updated CDNo: {so_tk}")
                        except Exception as e:
                            print(f"Debug - Error updating CDNo {so_tk}: {str(e)}")
                            return {'error': f'Lỗi khi cập nhật Debit Status cho CDNo {so_tk}: {str(e)}'}

            # Kết quả trả về
            result = {
                'missing_count': len(filtered_items),
                'updated_count': len(updated_items),
                'missing_details': [
                    f"CDNo: {item['CDNo']}, Client: {item['Client']}, InvoiceDate: {item['InvoiceDate']}"
                    for item in filtered_items
                ],
                'updated_details': [
                    f"CDNo: {item['CDNo']}, Client: {item['Client']}, InvoiceDate: {item['InvoiceDate']}"
                    for item in updated_items
                ]
            }
            print(f"Debug - Final result: {result}")
            return result

        except Exception as e:
            print(f"Debug - Error in check_declarations: {str(e)}")
            return {'error': f'Lỗi khi kiểm tra tờ khai: {str(e)}'}