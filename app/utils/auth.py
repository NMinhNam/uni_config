from functools import wraps
from flask import session, redirect, url_for

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session or not session['user'].get('authenticated'):
            return redirect(url_for('auth.login_page'))
        return f(*args, **kwargs)
    return decorated_function 