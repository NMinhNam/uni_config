#!/usr/bin/env python3
"""
WSGI Entry Point for Production Deployment
Sử dụng cho Gunicorn, uWSGI, và các production servers khác
"""

import os
import sys

# Thêm project directory vào Python path
project_home = os.path.dirname(os.path.abspath(__file__))
if project_home not in sys.path:
    sys.path.insert(0, project_home)

# Import Flask app
from app import app

# WSGI application object
application = app

# Production configuration
if __name__ == "__main__":
    # Chỉ chạy khi gọi trực tiếp python wsgi.py
    # Trong production, WSGI server sẽ import 'application'
    port = int(os.environ.get('PORT', 8000))
    print(f"🚀 WSGI Server starting on port {port}")
    application.run(host='0.0.0.0', port=port, debug=False) 