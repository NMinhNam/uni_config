import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment

class SplitFileService:
    def __init__(self, split_model):
        self.split_model = split_model
        self.sheets = ['SEA', 'AIR']

    def split_file(self, input_file_path, template_file_path, output_dir):
        """Process Excel file sheets"""
        try:
            processed_files = []
            for sheet_name in self.sheets:
                try:
                    result = self.split_model.process_sheet(
                        sheet_name,
                        input_file_path,
                        template_file_path,
                        output_dir
                    )
                    if result['status'] == 'success':
                        processed_files.extend(result['files'])
                except Exception as e:
                    print(f"Warning - Could not process sheet {sheet_name}: {str(e)}")
                    continue

            if not processed_files:
                return {
                    'status': 'error',
                    'message': 'Không có sheet nào được xử lý'
                }

            return {
                'status': 'success',
                'message': 'Xử lý file thành công',
                'files': processed_files
            }

        except Exception as e:
            return {
                'status': 'error',
                'message': f'Lỗi xử lý file: {str(e)}'
            }