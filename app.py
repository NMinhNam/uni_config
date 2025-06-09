from flask import Flask, session, redirect, url_for
from flask_cors import CORS
from app.controllers.auth_controller import auth_bp
from app.controllers.invoice_controller import invoice_bp
from app.controllers.dashboard_cd_controller import dashboard_bp
from app.controllers.check_declaration_controller import check_declaration_bp
from app.controllers.split_file_controller import split_bp
from app.controllers.report_soa_controller import report_soa_bp
import os
from functools import wraps
from config import Config
from datetime import timedelta, datetime
# from flask_session import Session  # Bỏ Flask-Session

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
STATIC_DIR = os.path.join(BASE_DIR, 'app', 'static')

app = Flask(__name__, 
            template_folder=TEMPLATE_DIR,
            static_folder=STATIC_DIR)

# Cấu hình session
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['SESSION_TYPE'] = 'filesystem'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)  # 30 phút
app.config['SESSION_REFRESH_EACH_REQUEST'] = True

# Cấu hình CORS
CORS(app, supports_credentials=True)

# app.config.from_object(Config)

# Session(app)  # Bỏ Flask-Session

@app.before_request
def check_session_timeout():
    if 'user' in session and session['user'].get('authenticated'):
        # Lấy thời gian login từ session
        login_time = session['user'].get('login_time', 0)
        current_time = datetime.now().timestamp()
        
        # Nếu đã quá 30 phút
        if current_time - login_time > 30 * 60:
            session.clear()
            return redirect(url_for('auth.login_page'))
        else:
            # Cập nhật lại thời gian truy cập
            session['user']['login_time'] = current_time

@app.route('/')
def index():
    print('Session:', dict(session))
    if 'user' in session and session['user'].get('authenticated'):
        return redirect(url_for('dashboard.index'))  # Chuyển hướng đến dashboard
    return redirect(url_for('auth.login_page'))

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or not session['user'].get('authenticated'):
            return redirect(url_for('auth.login_page'))
        return f(*args, **kwargs)
    return decorated_function

app.register_blueprint(auth_bp)
app.register_blueprint(invoice_bp)
app.register_blueprint(dashboard_bp)  # Đăng ký dashboard blueprint
app.register_blueprint(check_declaration_bp)  # Đăng ký check declaration blueprint
app.register_blueprint(split_bp, url_prefix='/split')  # Đăng ký split file blueprint
app.register_blueprint(report_soa_bp, url_prefix='/report-soa')  # Đăng ký report soa blueprint

if __name__ == '__main__':
    app.run(debug=True)