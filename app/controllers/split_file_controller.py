from flask import Blueprint, render_template, request, jsonify, send_from_directory
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename
from app.models.split_file_model import SplitModel
from app.services.split_file_service import SplitFileService
from app.controllers.auth_controller import login_required

# Tạo blueprint với prefix là /split
split_bp = Blueprint('split', __name__, url_prefix='/split')
CORS(split_bp)

split_model = SplitModel()
split_service = SplitFileService(split_model)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@split_bp.route('/')
@login_required
def index():
    """Render the main page for splitting Excel files"""
    return render_template('split_file.html')

@split_bp.route('/api/split-file', methods=['POST'])
@login_required
def split_file():
    input_file_path = None
    template_file_path = None
    
    try:
        # Add debug logging
        print("Debug - Request files:", request.files)
        
        if 'input_file' not in request.files or 'template_file' not in request.files:
            return jsonify({
                'status': 'error',
                'message': 'Vui lòng upload cả file tổng và file mẫu'
            }), 400

        input_file = request.files['input_file']
        template_file = request.files['template_file']

        # Add more validation
        if not input_file.filename or not template_file.filename:
            return jsonify({
                'status': 'error',
                'message': 'File không hợp lệ'
            }), 400

        # Log file details
        print(f"Debug - Input file: {input_file.filename}")
        print(f"Debug - Template file: {template_file.filename}")

        # Secure filenames
        input_filename = secure_filename(input_file.filename)
        template_filename = secure_filename(template_file.filename)

        # Create necessary directories with absolute paths
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        uploads_dir = os.path.join(base_dir, 'uploads')
        output_dir = os.path.join(base_dir, 'Data_Export')
        
        os.makedirs(uploads_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # Save uploaded files with absolute paths
        input_file_path = os.path.join(uploads_dir, input_filename)
        template_file_path = os.path.join(uploads_dir, template_filename)
        
        print(f"Debug - Saving files to: {uploads_dir}")
        input_file.save(input_file_path)
        template_file.save(template_file_path)

        # Process files
        result = split_service.split_file(input_file_path, template_file_path, output_dir)
        return jsonify(result), 200 if result['status'] == 'success' else 400

    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Lỗi xử lý: {str(e)}'
        }), 500
    
    finally:
        # Clean up temporary files
        for file_path in [input_file_path, template_file_path]:
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print(f"Debug - Removed temp file: {file_path}")
                except Exception as e:
                    print(f"Debug - Error removing temp file: {str(e)}")

@split_bp.route('/download/<filename>')
@login_required
def download_file(filename):
    """Handle file download requests"""
    try:
        filename = secure_filename(filename)
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Data_Export')
        
        if not os.path.exists(os.path.join(output_dir, filename)):
            return jsonify({
                'status': 'error',
                'message': 'File không tồn tại'
            }), 404
            
        return send_from_directory(
            output_dir, 
            filename, 
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        print(f"Debug - Error downloading file: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': f'Lỗi tải file: {str(e)}'
        }), 500