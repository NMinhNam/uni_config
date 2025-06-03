from flask import Blueprint, render_template, jsonify, request, Response
from app.models.dashboard_cd_model import DebitModel
from app.services.sharepoint_service import SharePointService
from app.controllers.auth_controller import login_required
import logging
import json
import time
import datetime

# Cấu hình logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

dashboard_bp = Blueprint('dashboard', __name__, url_prefix='/dashboard')
model = DebitModel()

# Thêm biến toàn cục để lưu trữ thời gian cập nhật cuối cùng
last_update_time = time.time()

@dashboard_bp.route('/')
@login_required
def index():
    return render_template('dashboard_cd.html')

@dashboard_bp.route('/api/debit_list', methods=['GET'])
@login_required
def debit_list():
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 10))
        client_filter = request.args.get('client_filter', None)
        from_date = request.args.get('from_date', None)
        to_date = request.args.get('to_date', None)
        import_export_filter = request.args.get('import_export_filter', None)

        debit_model = DebitModel()
        result = debit_model.get_debit_list(
            page=page,
            per_page=per_page,
            client_filter=client_filter,
            from_date=from_date,
            to_date=to_date,
            import_export_filter=import_export_filter
        )

        if "error" in result:
            return jsonify({"error": result["error"]}), 500

        print(f"Retrieved {len(result['items'])} items")
        return jsonify(result)

    except Exception as e:
        print(f"Error in debit_list: {str(e)}")
        return jsonify({"error": str(e)}), 500

@dashboard_bp.route('/api/debit_statistics', methods=['GET'])
@login_required
def debit_statistics():
    try:
        logger.debug("Đang gọi API debit_statistics")
        data = model.get_debit_statistics()
        if "error" in data:
            logger.error(f"Lỗi khi lấy thống kê: {data['error']}")
            return jsonify({"error": data["error"]}), 500
        logger.debug(f"Lấy thống kê thành công: {data}")
        return jsonify(data)
    except Exception as e:
        logger.error(f"Lỗi không xác định: {str(e)}")
        return jsonify({"error": str(e)}), 500

@dashboard_bp.route('/api/update_item/<int:item_id>', methods=['POST'])
@login_required
def update_item(item_id):
    try:
        data = request.get_json()
        model = DebitModel()
        sharepoint = model.sharepoint
        ctx = sharepoint.get_context()
        target_list = sharepoint.get_list_by_key('debit')

        # Lấy thông tin về kiểu dữ liệu của list
        ctx.load(target_list)
        ctx.execute_query()
        list_item_type = target_list.properties['ListItemEntityTypeFullName']
        logger.debug(f"List item type: {list_item_type}")

        item_to_update = target_list.get_item_by_id(item_id)
        ctx.load(item_to_update)
        ctx.execute_query()

        # Chuẩn bị dữ liệu để cập nhật
        update_properties = {}
        # Thêm thông tin kiểu dữ liệu cho item
        update_properties['__metadata'] = {'type': list_item_type}

        # Ánh xạ tên trường từ frontend sang SharePoint
        field_mapping = {
            "Invoice": "Invoice",
            "InvoiceDate": "InvoiceDate",
            "ParkingList": "ParkingList",
            "DebitStatus": "DebitStatus",
            "Client": "Client",
            "Import_x002f_Export": "Import_x002f_Export",
            "CD No": "CDNo",
            "Term": "OData__x0110_i_x1ec1_u_x0020_ki_x1ec7_",
            "Ma dia diem do hang": "M_x00e3__x0111__x1ecb_a_x0111_i_",
            "Ma dia diem xep hang": "M_x00e3__x0111__x1ecb_a_x0111_i_0",
            "Tong tri gia": "T_x1ed5_ng_x0020_tr_x1ecb__x00200",
            "Ma loai hinh": "M_x00e3__x0020_lo_x1ea1_i_x0020_",
            "Fist CD No": "S_x1ed1__x0020_TK_x0020__x0111__",
            "Loai CD": "Lo_x1ea1_i_x0020_TK",
            "Phan luong": "Ph_x00e2_n_x0020_lu_x1ed3_ng",
            "HBL": "V_x1ead_n_x0020__x0111__x01a1_n",
            "CD Date": "Date_x0020_of_x0020_CD",
            "So luong kien": "S_x1ed1__x0020_l_x01b0__x1ee3_ng",
            "tong trong luong": "T_x1ed5_ng_x0020_l_x01b0__x1ee3_",
            "so luong cont": "S_x1ed1__x0020_l_x01b0__x1ee3_ng0",
            "cont number": "Cont_x0020_Number",
            "ETA": "ArrivalDate",
            "ETD": "Ng_x00e0_ykh_x1edf_ih_x00e0_nh_x"
        }

        for frontend_field, value in data.items():
            if frontend_field in field_mapping:
                sharepoint_field_name = field_mapping[frontend_field]
                # Chuyển đổi giá trị sang chuỗi, xử lý None
                converted_value = str(value) if value is not None else None
                update_properties[sharepoint_field_name] = converted_value
                logger.debug(f"Chuẩn bị cập nhật trường {frontend_field} ({sharepoint_field_name}) với giá trị: {converted_value}")
            else:
                logger.warning(f"Bỏ qua trường không xác định từ frontend: {frontend_field}")

        # Gán các thuộc tính cần cập nhật vào item
        for field_name, value in update_properties.items():
            if field_name != '__metadata':  # Không cập nhật metadata trực tiếp
                item_to_update.set_property(field_name, value)

        # Cập nhật và thực thi
        item_to_update.update()
        ctx.execute_query()

        # Load lại item để kiểm tra giá trị sau khi cập nhật
        ctx.load(item_to_update)
        ctx.execute_query()
        logger.debug(f"Item sau khi cập nhật: {item_to_update.properties}")

        logger.info(f"Cập nhật item {item_id} thành công")
        return jsonify({"message": f"Cập nhật item {item_id} thành công"}), 200

    except Exception as e:
        logger.error(f"Lỗi tổng quát khi cập nhật mục {item_id}: {str(e)}")
        if hasattr(e, 'response') and e.response is not None:
            logger.error(f"SharePoint response error: {e.response.text}")
        return jsonify({"error": f"Lỗi khi cập nhật mục: {str(e)}"}), 500

@dashboard_bp.route('/api/check_updates')
@login_required
def check_updates():
    try:
        global last_update_time
        current_time = time.time()
        
        # Kiểm tra thay đổi từ SharePoint
        sharepoint = model.sharepoint
        ctx = sharepoint.get_context()
        target_list = sharepoint.get_list_by_key('debit')
        
        # Lấy thời gian cập nhật cuối cùng của list
        ctx.load(target_list)
        ctx.execute_query()
        
        has_updates = False
        if target_list.properties.get('LastItemModifiedDate'):
            last_modified = target_list.properties['LastItemModifiedDate']
            if last_modified and last_modified.timestamp() > last_update_time:
                has_updates = True
                last_update_time = last_modified.timestamp()
        
        return jsonify({
            'has_updates': has_updates,
            'last_update': last_update_time
        })
        
    except Exception as e:
        logger.error(f"Error checking updates: {str(e)}")
        return jsonify({
            'has_updates': False,
            'error': str(e)
        }), 500

@dashboard_bp.route('/api/events')
@login_required
def events():
    def event_stream():
        global last_update_time
        try:
            # Tạo context và auth một lần duy nhất
            sharepoint = model.sharepoint
            ctx = sharepoint.get_context()
            target_list = sharepoint.get_list_by_key('debit')
            
            # Load target list một lần duy nhất
            ctx.load(target_list)
            ctx.execute_query()
            logger.info(f"SSE: Initial list load successful. Last Modified Date: {target_list.properties.get('LastItemModifiedDate')}")
            
            while True:
                try:
                    # Cập nhật thời gian kiểm tra cuối cùng của server
                    server_check_time = time.time()
                    
                    if server_check_time - last_update_time >= 60:  # Kiểm tra mỗi 60 giây
                        logger.debug('SSE: Performing periodic check for updates...')
                        # Chỉ load lại properties của list, không tạo kết nối mới
                        ctx.load(target_list)
                        ctx.execute_query()
                        
                        current_last_modified = target_list.properties.get('LastItemModifiedDate')
                        
                        sharepoint_dt = None
                        # Xử lý giá trị LastItemModifiedDate nhận được
                        if current_last_modified is not None:
                            if isinstance(current_last_modified, str):
                                # Nếu là chuỗi, phân tích từ ISO 8601
                                try:
                                     sharepoint_dt = datetime.datetime.fromisoformat(current_last_modified.replace('Z', '+00:00'))
                                except ValueError as ve:
                                     logger.error(f"SSE: Error parsing SharePoint date string '{current_last_modified}': {ve}")
                            elif isinstance(current_last_modified, datetime.datetime):
                                # Nếu đã là datetime object
                                if current_last_modified.tzinfo is None: # Nếu là naive datetime
                                    # Giả định naive datetime từ SharePoint là UTC
                                    sharepoint_dt = current_last_modified.replace(tzinfo=datetime.timezone.utc)
                                else: # Nếu đã là timezone-aware
                                    # Chuyển đổi sang UTC
                                    sharepoint_dt = current_last_modified.astimezone(datetime.timezone.utc)
                            else:
                                 logger.warning(f"SSE: Unexpected type for LastItemModifiedDate: {type(current_last_modified)}")

                        # Chuyển last_update_time (timestamp) sang datetime object có nhận biết múi giờ (UTC)
                        server_last_update_dt = datetime.datetime.fromtimestamp(last_update_time, datetime.timezone.utc)
                        
                        logger.debug(f"SSE: SharePoint Last Modified DateTime (UTC): {sharepoint_dt}. Server last_update_dt (UTC): {server_last_update_dt}")

                        if sharepoint_dt and sharepoint_dt > server_last_update_dt:
                            logger.info(f"SSE: Updates detected! SharePoint: {sharepoint_dt}, Server: {server_last_update_dt}")
                            last_update_time = sharepoint_dt.timestamp()
                            yield f"data: {json.dumps({'has_updates': True, 'last_update': last_update_time})}\n\n"
                        else:
                            logger.debug('SSE: No updates detected')
                            yield f"data: {json.dumps({'has_updates': False, 'last_update': last_update_time})}\n\n"
                    
                    time.sleep(1)  # Đợi 1 giây trước khi kiểm tra lại
                    
                except Exception as e:
                    logger.error(f"SSE: Error in update check loop: {str(e)}")
                    yield f"data: {json.dumps({'error': str(e)})}\n\n"
                    time.sleep(5)  # Đợi lâu hơn nếu có lỗi
                    
        except Exception as e:
            logger.error(f"SSE: Error in event stream initialization: {str(e)}")
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
            
    return Response(event_stream(), mimetype='text/event-stream')


