from flask import Flask, session, redirect, url_for
from flask_cors import CORS
from app.controllers.auth_controller import auth_bp
from app.controllers.invoice_controller import invoice_bp
import os
from functools import wraps
from config import Config
from datetime import timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
STATIC_DIR = os.path.join(BASE_DIR, 'app', 'static')

app = Flask(__name__, 
            template_folder=TEMPLATE_DIR,
            static_folder=STATIC_DIR)

# Cấu hình session
app.secret_key = os.urandom(24)
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=1)
app.config['SESSION_COOKIE_SECURE'] = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_REFRESH_EACH_REQUEST'] = True
app.config['SESSION_TYPE'] = 'filesystem'

CORS(app)
app.config.from_object(Config)

@app.route('/')
def index():
    if 'user' in session and session['user'].get('authenticated'):
        return redirect(url_for('invoice.render_download_invoices'))  # Cập nhật endpoint
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

if __name__ == '__main__':
    app.run(debug=True)