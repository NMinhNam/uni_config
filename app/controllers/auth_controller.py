from flask import Blueprint, request, jsonify, session, redirect, url_for, render_template
import os

# Import Service Layer
from app.services.sharepoint_service import SharepointService

auth_bp = Blueprint('auth', __name__)

# Tạo instance của Service Layer
sharepoint_service = SharepointService()

# Route hiển thị trang đăng nhập
@auth_bp.route('/login-page', methods=['GET'])
def login_page():
    # Kiểm tra xem người dùng đã đăng nhập chưa để chuyển hướng
    if 'user' in session and session['user'].get('authenticated'):
        # Chuyển hướng đến trang tải hóa đơn
        return redirect(url_for('invoice.render_download_invoices'))  # Cập nhật endpoint
    return render_template('login.html')  # Render file login.html

# Route xử lý yêu cầu POST đăng nhập từ form/AJAX
@auth_bp.route('/login', methods=['POST'])
def login():
    data = request.get_json()  # Lấy dữ liệu JSON từ request body
    username = data.get('username')
    password = data.get('password')

    if not username or not password:
        return jsonify({'success': False, 'message': 'Vui lòng nhập đầy đủ thông tin đăng nhập'}), 400

    # GỌI ĐẾN SERVICE LAYER ĐỂ XÁC THỰC
    try:
        ctx = sharepoint_service.authenticate_user(username, password)
        if ctx is None:  # Nếu xác thực thất bại
            return jsonify({'success': False, 'message': 'Đăng nhập thất bại! Vui lòng kiểm tra lại thông tin đăng nhập'}), 401
        
        # Thử thực hiện một truy vấn đơn giản để xác minh kết nối
        ctx.load(ctx.web, 'Title')
        ctx.execute_query()
        
        # Nếu đến được đây, nghĩa là xác thực thành công
        session['user'] = {
            'username': username,
            'authenticated': True
        }
        return jsonify({
            'success': True, 
            'message': 'Đăng nhập thành công',
            'redirect': url_for('invoice.render_download_invoices')  # Cập nhật endpoint
        }), 200
    except Exception as e:
        print(f"Lỗi đăng nhập: {str(e)}")
        return jsonify({'success': False, 'message': 'Đăng nhập thất bại! Vui lòng kiểm tra lại thông tin đăng nhập'}), 401

@auth_bp.route('/logout', methods=['GET'])
def logout():
    session.clear()  # Xóa session
    # Chuyển hướng về trang đăng nhập
    return redirect(url_for('auth.login_page'))  # Sử dụng endpoint name của Blueprint

@auth_bp.route('/check-auth', methods=['GET'])
def check_auth():
    if 'user' in session and session['user'].get('authenticated'):
        return jsonify({'authenticated': True}), 200
    return jsonify({'authenticated': False}), 401