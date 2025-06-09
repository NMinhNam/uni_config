const form = document.getElementById('pdfExcelForm');
const resultDiv = document.getElementById('result');
const pdfFolderInput = document.getElementById('pdfFolder');
const pdfFolderDisplay = document.getElementById('pdf_folder_display');

function selectFolder() {
    pdfFolderInput.click();
    pdfFolderInput.addEventListener('change', (e) => {
        const files = e.target.files;
        if (files.length > 0) {
            const folderName = files[0].webkitRelativePath.split('/')[0];
            pdfFolderDisplay.value = `${folderName} (${files.length} file PDF đã chọn)`;
        } else {
            pdfFolderDisplay.value = 'Chưa chọn thư mục...';
        }
    });
}

form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const excelFile = document.getElementById('excel_file').files[0];
    const pdfFiles = document.getElementById('pdfFolder').files;

    if (!excelFile || pdfFiles.length === 0) {
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Vui lòng chọn file Excel và ít nhất một file PDF!'
        });
        return;
    }

    const formData = new FormData();
    formData.append('excel_file', excelFile);
    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append('pdf_files[]', pdfFiles[i]);
    }

    resultDiv.innerHTML = '<p class="text-info">Đang xử lý...</p>';

    try {
        const response = await axios.post('/expense/api/process', formData, {
            headers: {
                'Content-Type': 'multipart/form-data'
            }
        });

        const data = response.data;
        if (data.status === 'success') {
            resultDiv.innerHTML = `<p class="text-success">${data.message}</p>`;
            Swal.fire({
                icon: 'success',
                title: 'Thành công',
                text: data.message
            });
        } else {
            resultDiv.innerHTML = `<p class="text-danger">${data.message}</p>`;
            Swal.fire({
                icon: 'error',
                title: 'Lỗi',
                text: data.message
            });
        }
    } catch (error) {
        resultDiv.innerHTML = `<p class="text-danger">Lỗi: ${error.message}</p>`;
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Đã có lỗi xảy ra, vui lòng thử lại.'
        });
    }
});
