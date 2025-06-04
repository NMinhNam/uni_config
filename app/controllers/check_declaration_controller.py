from flask import Flask, request, jsonify, render_template, session, Blueprint
from flask_cors import CORS
from app.models.check_declaration_model import DeclarationModel
import sys
import os
import re
from datetime import datetime
from app.controllers.auth_controller import login_required
from werkzeug.utils import secure_filename
import pandas as pd
import json

sys.stdout.reconfigure(encoding='utf-8')

# Tạo thư mục uploads nếu chưa tồn tại
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

# Khởi tạo Flask với thư mục templates chính xác
app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config.from_object('config.Config')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['SESSION_TYPE'] = 'filesystem'
CORS(app, supports_credentials=True)

declaration_model = DeclarationModel()

check_declaration_bp = Blueprint('check_declaration', __name__, url_prefix='/check-declaration')

@app.route('/', methods=['GET'])
def index():
    return render_template('home.html')

@check_declaration_bp.route('/')
@check_declaration_bp.route('/check-declarations')
@login_required
def check_declarations_page():
    return render_template('check_declaration.html')

@check_declaration_bp.route('/api/clients')
@login_required
def get_clients():
    try:
        model = DeclarationModel()
        result = model.get_clients()
        if 'error' in result:
            return jsonify({'error': result['error']}), 500
        return jsonify({'clients': result['clients']})
    except Exception as e:
        print(f"Error in get_clients: {str(e)}")
        return jsonify({'error': 'Không thể tải danh sách khách hàng'}), 500

@check_declaration_bp.route('/api/check', methods=['POST'])
@login_required
def check_declarations():
    try:
        print("Debug - Starting check_declarations")
        print(f"Debug - Session user: {session.get('user')}")
        
        if 'user' not in session or not session['user'].get('authenticated'):
            print("Debug - User not authenticated")
            return jsonify({'error': 'Vui lòng đăng nhập lại'}), 401
            
        # Lấy thông tin đăng nhập từ session
        username = session['user'].get('username')
        password = session['user'].get('password')
        
        if not username or not password:
            print("Debug - Missing credentials in session")
            return jsonify({'error': 'Vui lòng đăng nhập lại'}), 401
            
        # Lấy dữ liệu từ form
        from_date_str = request.form.get('from_date')
        to_date_str = request.form.get('to_date')
        selected_client = request.form.get('client')
        excel_file = request.files.get('excelFile')
        
        print(f"Debug - Form data: from_date={from_date_str}, to_date={to_date_str}, client={selected_client}")
        print(f"Debug - File received: {excel_file.filename if excel_file else None}")
        
        # Kiểm tra dữ liệu đầu vào
        if not from_date_str or not to_date_str:
            return jsonify({'error': 'Vui lòng chọn khoảng thời gian'}), 400
            
        if not selected_client:
            return jsonify({'error': 'Vui lòng chọn khách hàng'}), 400
            
        if not excel_file or not excel_file.filename:
            return jsonify({'error': 'Vui lòng chọn file Excel'}), 400
            
        # Kiểm tra định dạng file
        if not excel_file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'File không đúng định dạng. Vui lòng chọn file Excel (.xlsx hoặc .xls)'}), 400
        
        # Tạo thư mục uploads nếu chưa tồn tại
        uploads_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'uploads')
        os.makedirs(uploads_dir, exist_ok=True)
        
        # Lưu file tạm thời
        file_path = os.path.join(uploads_dir, excel_file.filename)
        excel_file.save(file_path)
        print(f"Debug - File saved to: {file_path}")
        
        try:
            model = DeclarationModel()  # Tạo instance mới của model
            
            # Xử lý file Excel
            excel_result = model.process_excel_file(file_path)
            if 'error' in excel_result:
                return jsonify(excel_result), 400
                
            cd_no_set = excel_result['data']
            print(f"Debug - Processed {len(cd_no_set)} declarations from Excel")
            
            # Kết nối SharePoint
            sharepoint_result = model.connect_sharepoint()
            if 'error' in sharepoint_result:
                return jsonify(sharepoint_result), 401
                
            ctx = sharepoint_result['context']
            
            # Kiểm tra tờ khai
            try:
                from_date = datetime.strptime(from_date_str, '%d/%m/%Y')
                to_date = datetime.strptime(to_date_str, '%d/%m/%Y')
            except Exception as e:
                return jsonify({'error': 'Định dạng ngày không hợp lệ. Vui lòng nhập đúng định dạng dd/mm/yyyy'}), 400
            
            result = model.check_declarations(ctx, cd_no_set, from_date, to_date, selected_client)
            if 'error' in result:
                return jsonify(result), 400
                
            return jsonify(result)
            
        finally:
            # Xóa file tạm
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Debug - Removed file: {file_path}")
                
    except Exception as e:
        print(f"Debug - Error in check_declarations: {str(e)}")
        return jsonify({'error': f'Lỗi khi kiểm tra tờ khai: {str(e)}'}), 500