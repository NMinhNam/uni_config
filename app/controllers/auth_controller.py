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
@auth_bp.route('/login-page')
def login_page():
    if 'user' in session and session['user'].get('authenticated'):
        return redirect(url_for('dashboard.index'))  # Chuyển hướng đến dashboard
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
            'redirect': url_for('dashboard.index')  # Chuyển hướng đến dashboard
        })
    return jsonify({
        'success': False,
        'message': message
    })

@auth_bp.route('/logout')
def logout():
    auth_service.logout()
    return redirect(url_for('auth.login_page'))

@auth_bp.route('/check-auth')
@login_required
def check_auth():
    return jsonify({'authenticated': True})