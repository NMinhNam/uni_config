from flask import Blueprint, request, jsonify, render_template, session, send_from_directory
from app.services.report_soa_service import ReportSoaService
from app.controllers.auth_controller import login_required
import os

report_soa_bp = Blueprint('report_soa', __name__)
report_soa_service = ReportSoaService()

@report_soa_bp.route('/')
@login_required
def index():
    return render_template('report_soa.html')

@report_soa_bp.route('/generate', methods=['POST'])
@login_required
def generate_report():
    try:
        data = request.get_json()
        selected_company = data.get('selected_company')
        from_date = data.get('from_date')
        to_date = data.get('to_date')
        file_format = data.get('file_format', 'docx')  # Default to docx if not specified

        if not all([from_date, to_date, selected_company]):
            return jsonify({'error': 'Vui lòng điền đầy đủ thông tin'}), 400

        # Tạo thư mục SOA nếu chưa tồn tại
        soa_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'soa')
        os.makedirs(soa_dir, exist_ok=True)

        # Lấy dữ liệu từ SharePoint
        result = report_soa_service.fetch_sharepoint_data(from_date, to_date, selected_company)
        
        # Kiểm tra lỗi
        if isinstance(result, dict) and 'error' in result:
            if 'đăng nhập lại' in result['error']:
                # Clear session và yêu cầu đăng nhập lại
                session.clear()
                return jsonify({'error': result['error'], 'redirect': '/login-page'}), 401
            return jsonify(result), 400

        filtered_data, company_info = result

        if not filtered_data:
            return jsonify({'error': 'Không tìm thấy dữ liệu phù hợp'}), 404

        # Tạo file theo định dạng được chọn
        file_name = report_soa_service.create_word_file(
            filtered_data, 
            selected_company, 
            company_info, 
            from_date, 
            to_date, 
            soa_dir,
            file_format
        )
        
        return jsonify({
            'success': True,
            'message': f'Tạo Statement of Account thành công ({file_format.upper()})',
            'file_name': file_name,
            'download_url': f'/report-soa/download/{file_name}'
        })

    except Exception as e:
        print(f"Error in generate_report: {str(e)}")
        return jsonify({'error': 'Có lỗi xảy ra khi tạo Statement of Account'}), 500

@report_soa_bp.route('/download/<filename>')
@login_required
def download_file(filename):
    try:
        soa_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'soa')
        return send_from_directory(soa_dir, filename, as_attachment=True)
    except Exception as e:
        print(f"Error in download_file: {str(e)}")
        return jsonify({'error': 'Không thể tải file'}), 404 