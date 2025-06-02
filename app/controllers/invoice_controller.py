from flask import Blueprint, request, jsonify, render_template
from app.models.invoice_model import InvoiceModel
from app.controllers.auth_controller import login_required
import os
import re
import traceback

invoice_bp = Blueprint('invoice', __name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

invoice_model_instance = None

def get_invoice_model():
    global invoice_model_instance
    if invoice_model_instance is None:
        invoice_model_instance = InvoiceModel()
    return invoice_model_instance

@invoice_bp.route('/download-invoices')
@login_required
def render_download_invoices():
    invoice_model = get_invoice_model()
    invoice_model.close_driver()
    return render_template('download_invoices.html')

@invoice_bp.route('/api/get-captcha', methods=['GET'])
@login_required
def get_captcha():
    try:
        invoice_model = get_invoice_model()
        success, result = invoice_model.get_captcha_image()
        if not success:
            return jsonify({'status': 'error', 'message': result, 'new_captcha_url': None}), 400
        return jsonify({'status': 'success', 'captcha_url': result}), 200
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Lỗi khi tải CAPTCHA: {str(e)}', 'new_captcha_url': None}), 400

@invoice_bp.route('/api/download-invoices', methods=['POST'])
@login_required
def download_invoices():
    invoice_model = get_invoice_model()
    try:
        company = request.form.get('company')
        username = request.form.get('username')
        password = request.form.get('password')
        from_date = request.form.get('from_date')
        to_date = request.form.get('to_date')
        captcha_code = request.form.get('captcha_code')

        if not all([company, username, password, from_date, to_date, captcha_code]):
            return jsonify({'status': 'error', 'message': 'Vui lòng điền đầy đủ các trường!'}), 400

        is_valid, message_or_dates = invoice_model.validate_input(username, password, from_date, to_date, company)
        if not is_valid:
            return jsonify({'status': 'error', 'message': message_or_dates}), 400

        result = invoice_model.process_invoices(username, password, captcha_code, from_date, to_date, company)
        if result['status'] == 'captcha_error':
            return jsonify({
                'status': 'captcha_error',
                'message': result['message'],
                'new_captcha_url': result['new_captcha_url']
            }), 400
        elif result['status'] == 'error':
            return jsonify({'status': 'error', 'message': result['message']}), 400
        return jsonify({'status': 'success', 'message': result['message']}), 200
    except Exception as e:
        error_message = f"Đã xảy ra lỗi khi tải hóa đơn: {str(e)}\n{traceback.format_exc()}"
        print(error_message)  
        return jsonify({'status': 'error', 'message': 'Đã xảy ra lỗi phía server, vui lòng thử lại sau.'}), 500