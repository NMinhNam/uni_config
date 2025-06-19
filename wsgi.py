#!/usr/bin/env python3
"""
WSGI Entry Point for Production Deployment
Sử dụng cho Gunicorn, uWSGI, và các production servers khác
"""

import os
import sys
import importlib.util

# Thêm project directory vào Python path
project_home = os.path.dirname(os.path.abspath(__file__))
if project_home not in sys.path:
    sys.path.insert(0, project_home)

# Import Flask app từ app.py bằng cách chỉ định đường dẫn tuyệt đối
# để tránh xung đột với package app/
app_path = os.path.join(project_home, 'app.py')
spec = importlib.util.spec_from_file_location('main_app', app_path)
main_app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(main_app)

# WSGI application object
application = main_app.app

# Production configuration
if __name__ == "__main__":
    # Chỉ chạy khi gọi trực tiếp python wsgi.py
    # Trong production, WSGI server sẽ import 'application'
    port = int(os.environ.get('PORT', 8000))
    print(f"🚀 WSGI Server starting on port {port}")
    application.run(host='0.0.0.0', port=port, debug=False) 