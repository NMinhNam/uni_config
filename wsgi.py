#!/usr/bin/env python3
"""
WSGI Entry Point for Production Deployment
Sử dụng cho Gunicorn, uWSGI, và các production servers khác
"""

import os
import sys
import importlib.util
import subprocess
import glob
import platform
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("wsgi")

# Fix ChromeDriver permissions at startup
def fix_chromedriver_permissions():
    """Find and fix ChromeDriver permissions in common locations"""
    try:
        logger.info("🔍 Checking system info before startup")
        system = platform.system()
        machine = platform.machine()
        python_version = platform.python_version()
        logger.info(f"🖥️ System: {system}, Architecture: {machine}, Python: {python_version}")
        
        # Common locations for ChromeDriver
        home_dir = os.path.expanduser("~")
        locations = [
            os.path.join(home_dir, ".wdm", "drivers", "chromedriver"),
            os.path.join("/tmp", ".wdm", "drivers", "chromedriver"),
            "/usr/local/bin/chromedriver",
            "/usr/bin/chromedriver",
            os.path.join(os.getcwd(), "chromedriver")
        ]
        
        fixed_count = 0
        for location in locations:
            # Use glob to find all chromedriver binaries
            for driver_path in glob.glob(f"{location}/**/*chromedriver*", recursive=True):
                if os.path.isfile(driver_path) and not driver_path.endswith('.zip'):
                    try:
                        current_mode = os.stat(driver_path).st_mode
                        os.chmod(driver_path, current_mode | 0o111)  # Add executable bit
                        logger.info(f"✅ Fixed permissions for: {driver_path}")
                        fixed_count += 1
                    except Exception as e:
                        logger.warning(f"⚠️ Could not fix permissions for {driver_path}: {str(e)}")
        
        # Fix THIRD_PARTY_NOTICES.chromedriver if it exists and causes problems
        for notice_path in glob.glob(f"{home_dir}/.wdm/**/THIRD_PARTY_NOTICES.chromedriver", recursive=True):
            try:
                if os.path.exists(notice_path):
                    # Create a fixed version without executable bit
                    with open(notice_path, 'rb') as notice_file:
                        content = notice_file.read()
                    
                    os.remove(notice_path)
                    with open(notice_path, 'wb') as notice_file:
                        notice_file.write(content)
                    
                    logger.info(f"✅ Fixed problematic THIRD_PARTY_NOTICES file: {notice_path}")
            except Exception as e:
                logger.warning(f"⚠️ Could not fix notices file {notice_path}: {str(e)}")
                
        logger.info(f"🔧 ChromeDriver permission check complete. Fixed {fixed_count} binaries.")
    except Exception as e:
        logger.error(f"❌ Error fixing permissions: {str(e)}")

# Run the permission fix at startup
fix_chromedriver_permissions()

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
    logger.info(f"🚀 WSGI Server starting on port {port}")
    application.run(host='0.0.0.0', port=port, debug=False)