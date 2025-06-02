from flask import Blueprint, request, jsonify, session, redirect, url_for, render_template
from app.services.auth_service import AuthService
from functools import wraps

auth_bp = Blueprint('auth', __name__)
auth_service = AuthService()

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not auth_service.is_authenticated():
            return redirect(url_for('auth.login_page'))
        return f(*args, **kwargs)
    return decorated_function

# Route hiển thị trang đăng nhập
@auth_bp.route('/login-page', methods=['GET'])
def login_page():
    if auth_service.is_authenticated():
        return redirect(url_for('invoice.render_download_invoices'))
    return render_template('login.html')

# Route xử lý yêu cầu POST đăng nhập từ form/AJAX
@auth_bp.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    
    success, message = auth_service.login(username, password)
    if success:
        return jsonify({
            'success': True, 
            'message': message,
            'redirect': url_for('invoice.render_download_invoices')
        })
    return jsonify({'success': False, 'message': message}), 401

@auth_bp.route('/logout')
def logout():
    auth_service.logout()
    return redirect(url_for('auth.login_page'))

@auth_bp.route('/check-auth')
@login_required
def check_auth():
    return jsonify({'authenticated': True})