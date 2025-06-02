# File: models/debit_model.py

import requests
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import List
from office365.runtime.queries.client_query import ClientQuery
from datetime import datetime

class DebitModel:
    def __init__(self):
        self.site_url = "https://uniconsulting079.sharepoint.com/sites/ClearanceFlowAutomation"
        self.list_name = "Debit list"
        self.username = "giangvo@eximuni.com"
        self.password = "Thiien2004"

    def get_all_items(self):
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_user(self.username, self.password)
            ctx = ClientContext(self.site_url, auth_context)
            target_list = ctx.web.lists.get_by_title(self.list_name)
            
            items = target_list.items.top(5000)  # Get up to 5000 items
            ctx.load(items)
            ctx.execute_query()
            
            all_items = []
            for item in items:
                all_items.append({
                    "ID": item.properties.get("ID"),
                    "Invoice": item.properties.get("Invoice"),
                    "InvoiceDate": item.properties.get("InvoiceDate"),
                    "ParkingList": item.properties.get("ParkingList"),
                    "DebitStatus": item.properties.get("DebitStatus"),
                    "Client": item.properties.get("Client"),
                    "Import_x002f_Export": item.properties.get("Import_x002f_Export"),
                    "CD No": item.properties.get("CDNo"),
                    "Term": item.properties.get("OData__x0110_i_x1ec1_u_x0020_ki_x1ec7_"),
                    "Ma dia diem do hang": item.properties.get("M_x00e3__x0111__x1ecb_a_x0111_i_"),
                    "Ma dia diem xep hang": item.properties.get("M_x00e3__x0111__x1ecb_a_x0111_i_0"),
                    "Tong tri gia": item.properties.get("T_x1ed5_ng_x0020_tr_x1ecb__x00200"),
                    "Ma loai hinh": item.properties.get("M_x00e3__x0020_lo_x1ea1_i_x0020_"),
                    "Fist CD No": item.properties.get("S_x1ed1__x0020_TK_x0020__x0111__"),
                    "Loai CD": item.properties.get("Lo_x1ea1_i_x0020_TK"),
                    "Phan luong": item.properties.get("Ph_x00e2_n_x0020_lu_x1ed3_ng"),
                    "HBL": item.properties.get("V_x1ead_n_x0020__x0111__x01a1_n"),
                    "CD Date": item.properties.get("Date_x0020_of_x0020_CD"),
                    "So luong kien": item.properties.get("S_x1ed1__x0020_l_x01b0__x1ee3_ng"),
                    "tong trong luong": item.properties.get("T_x1ed5_ng_x0020_l_x01b0__x1ee3_"),
                    "so luong cont": item.properties.get("S_x1ed1__x0020_l_x01b0__x1ee3_ng0"),
                    "cont number": item.properties.get("Cont_x0020_Number"),
                    "ETA": item.properties.get("ArrivalDate"),
                    "ETD": item.properties.get("Ng_x00e0_ykh_x1edf_ih_x00e0_nh_x"),
                })
            
            return all_items
        except Exception as e:
            return {"error": str(e)}

    def get_debit_list(self, page=1, per_page=10, client_filter=None, from_date=None, to_date=None, import_export_filter=None):
        try:
            auth_context = AuthenticationContext(self.site_url)
            auth_context.acquire_token_for_user(self.username, self.password)
            ctx = ClientContext(self.site_url, auth_context)
            target_list = ctx.web.lists.get_by_title(self.list_name)
            
            items = target_list.items.top(5000)  # Get up to 5000 items
            ctx.load(items)
            ctx.execute_query()
            
            # Chuyển đổi items thành list các dictionary
            all_items = []
            for item in items:
                all_items.append({
                    "ID": item.properties.get("ID"),
                    "Invoice": item.properties.get("Invoice"),
                    "InvoiceDate": item.properties.get("InvoiceDate"),
                    "ParkingList": item.properties.get("ParkingList"),
                    "DebitStatus": item.properties.get("DebitStatus"),
                    "Client": item.properties.get("Client"),
                    "Import_x002f_Export": item.properties.get("Import_x002f_Export"),
                    "CD No": item.properties.get("CDNo"),
                    "Term": item.properties.get("OData__x0110_i_x1ec1_u_x0020_ki_x1ec7_"),
                    "Ma dia diem do hang": item.properties.get("M_x00e3__x0111__x1ecb_a_x0111_i_"),
                    "Ma dia diem xep hang": item.properties.get("M_x00e3__x0111__x1ecb_a_x0111_i_0"),
                    "Tong tri gia": item.properties.get("T_x1ed5_ng_x0020_tr_x1ecb__x00200"),
                    "Ma loai hinh": item.properties.get("M_x00e3__x0020_lo_x1ea1_i_x0020_"),
                    "Fist CD No": item.properties.get("S_x1ed1__x0020_TK_x0020__x0111__"),
                    "Loai CD": item.properties.get("Lo_x1ea1_i_x0020_TK"),
                    "Phan luong": item.properties.get("Ph_x00e2_n_x0020_lu_x1ed3_ng"),
                    "HBL": item.properties.get("V_x1ead_n_x0020__x0111__x01a1_n"),
                    "CD Date": item.properties.get("Date_x0020_of_x0020_CD"),
                    "So luong kien": item.properties.get("S_x1ed1__x0020_l_x01b0__x1ee3_ng"),
                    "tong trong luong": item.properties.get("T_x1ed5_ng_x0020_l_x01b0__x1ee3_"),
                    "so luong cont": item.properties.get("S_x1ed1__x0020_l_x01b0__x1ee3_ng0"),
                    "cont number": item.properties.get("Cont_x0020_Number"),
                    "ETA": item.properties.get("ArrivalDate"),
                    "ETD": item.properties.get("Ng_x00e0_ykh_x1edf_ih_x00e0_nh_x"),
                })

            # Áp dụng filter
            filtered_items = all_items
            if client_filter:
                filtered_items = [item for item in filtered_items if item.get("Client", "").lower().find(client_filter.lower()) != -1]

            if import_export_filter:
                 filtered_items = [item for item in filtered_items if item.get("Import_x002f_Export") == import_export_filter]

            if from_date or to_date:
                filtered_items = [item for item in filtered_items if item.get("InvoiceDate")]
                
                # Chuyển đổi ngày từ chuỗi DD/MM/YYYY thành đối tượng Date
                def parse_date(date_str):
                    if not date_str:
                        return None
                    day, month, year = map(int, date_str.split('/'))
                    return datetime(year, month, day)

                # Lọc theo khoảng thời gian
                if from_date:
                    from_date_obj = datetime.strptime(from_date, '%Y-%m-%d')
                    filtered_items = [item for item in filtered_items if parse_date(item["InvoiceDate"]) and parse_date(item["InvoiceDate"]) >= from_date_obj]
                
                if to_date:
                    to_date_obj = datetime.strptime(to_date, '%Y-%m-%d')
                    filtered_items = [item for item in filtered_items if parse_date(item["InvoiceDate"]) and parse_date(item["InvoiceDate"]) <= to_date_obj]

            # Tính toán phân trang cho dữ liệu đã filter
            total_count = len(filtered_items)
            start_index = (page - 1) * per_page
            end_index = start_index + per_page
            paginated_items = filtered_items[start_index:end_index]
            
            return {
                "items": paginated_items,
                "total_items": total_count,
                "total_pages": (total_count + per_page - 1) // per_page,
                "current_page": page,
                "per_page": per_page
            }
        except Exception as e:
            return {"error": str(e)}

    def get_debit_statistics(self):
        try:
            all_items = self.get_all_items()
            if isinstance(all_items, dict) and "error" in all_items:
                return all_items

            # Đếm tất cả invoice
            all_invoices = [item.get("Invoice") for item in all_items if item.get("Invoice")]
            invoice_count = len(all_invoices)
            
            # Đếm số lượng theo trạng thái
            debit_count = sum(1 for item in all_items if item.get("DebitStatus") == "Debit")
            not_debit_count = sum(1 for item in all_items if item.get("DebitStatus") == "Not Debit")
            client_count = len(set(item.get("Client") for item in all_items if item.get("Client")))

            return {
                "debit_count": debit_count,
                "not_debit_count": not_debit_count,
                "client_count": client_count,
                "invoice_count": invoice_count
            }
        except Exception as e:
            return {"error": str(e)}