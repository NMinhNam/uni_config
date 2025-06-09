from flask import Blueprint, render_template, request, jsonify, send_from_directory
from app.services.expense_service import ExpenseService
from app.utils.auth import login_required
import os

# Thêm url_prefix cho blueprint
expense_bp = Blueprint('expense', __name__, 
                      template_folder='templates', 
                      static_folder='static',
                      url_prefix='/expense')  # Thêm prefix

expense_service = ExpenseService()

@expense_bp.route('/')  # Route chính đổi thành '/'
@login_required
def index():
    return render_template('expense.html')

@expense_bp.route('/api/process', methods=['POST'])  # Đơn giản hóa API endpoint
@login_required
def process_pdf_excel():
    try:
        print("Request files:", request.files)
        if 'excel_file' not in request.files:
            return jsonify({'status': 'error', 'message': 'Vui lòng upload file Excel'}), 400

        excel_file = request.files['excel_file']
        pdf_files = request.files.getlist('pdf_files[]')

        if not excel_file or not pdf_files:
            return jsonify({'status': 'error', 'message': 'Vui lòng chọn file Excel và ít nhất một file PDF'}), 400

        if not excel_file.filename.endswith('.xlsx'):
            return jsonify({'status': 'error', 'message': 'Vui lòng upload file Excel (.xlsx)'}), 400

        uploads_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'uploads')
        os.makedirs(uploads_dir, exist_ok=True)

        # Lưu file Excel
        excel_path = os.path.join(uploads_dir, excel_file.filename)
        excel_file.save(excel_path)
        print("Saved Excel to:", excel_path)

        # Lưu các file PDF
        pdf_paths = []
        for pdf_file in pdf_files:
            if pdf_file.filename.endswith('.pdf'):
                pdf_filename = os.path.basename(pdf_file.filename)
                pdf_path = os.path.join(uploads_dir, pdf_filename)
                pdf_file.save(pdf_path)
                pdf_paths.append(pdf_path)
                print("Saved PDF to:", pdf_path)

        if not pdf_paths:
            return jsonify({'status': 'error', 'message': 'Không tìm thấy file PDF hợp lệ'}), 400

        # Xử lý và cập nhật Excel sử dụng service
        output_path = expense_service.update_excel_with_pdf_data(excel_path, pdf_paths)
        print("Output path:", output_path)

        # Xóa các file PDF tạm thời sau khi xử lý
        for pdf_path in pdf_paths:
            try:
                os.remove(pdf_path)
                print("Deleted PDF:", pdf_path)
            except Exception as e:
                print(f"Error deleting PDF {pdf_path}: {e}")

        return jsonify({
            'status': 'success', 
            'message': 'Xử lý thành công [File được lưu tại Folder TONG_HOP_EXPENSE]', 
            'file': os.path.basename(output_path)
        }), 200

    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@expense_bp.route('/download/<filename>')
@login_required
def download_file(filename):
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'uploads')
    return send_from_directory(output_dir, filename, as_attachment=True)