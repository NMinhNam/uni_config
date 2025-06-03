document.addEventListener('DOMContentLoaded', function() {
    // Initialize datepickers
    $('.datepicker').datepicker({
        format: 'dd/mm/yyyy',
        autoclose: true,
        todayHighlight: true,
        language: 'vi'
    });

    // Load client list
    loadClients();

    // Form submission
    const form = document.getElementById('checkForm');
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        // Show loading state
        const submitButton = form.querySelector('button[type="submit"]');
        const originalText = submitButton.innerHTML;
        submitButton.disabled = true;
        submitButton.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Đang xử lý...';

        try {
            // Validate form
            if (!form.checkValidity()) {
                form.classList.add('was-validated');
                throw new Error('Vui lòng điền đầy đủ thông tin');
            }

            // Get form data
            const formData = new FormData(form);
            console.log('Debug - Form data being sent:');
            for (let pair of formData.entries()) {
                console.log(pair[0] + ': ' + pair[1]);
            }

            // Validate dates
            const fromDate = formData.get('from_date');
            const toDate = formData.get('to_date');
            if (!fromDate || !toDate) {
                throw new Error('Vui lòng chọn khoảng thời gian');
            }

            // Validate client
            const client = formData.get('client');
            if (!client) {
                throw new Error('Vui lòng chọn khách hàng');
            }

            // Validate file
            const file = formData.get('excelFile');
            if (!file || !file.name) {
                throw new Error('Vui lòng chọn file Excel');
            }
            if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                throw new Error('File không hợp lệ. Vui lòng chọn file Excel (.xlsx hoặc .xls)');
            }

            // Send request
            const response = await fetch('/check-declaration/api/check', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();
            console.log('Debug - Server response:', result);

            if (!response.ok) {
                throw new Error(result.error || 'Có lỗi xảy ra');
            }

            // Show success message
            Swal.fire({
                icon: 'success',
                title: 'Thành công!',
                text: 'Đã kiểm tra tờ khai thành công',
                showConfirmButton: false,
                timer: 1500
            });

            // Display results
            displayResults(result); // Sửa ở đây: bỏ .result

        } catch (error) {
            console.error('Error:', error);
            Swal.fire({
                icon: 'error',
                title: 'Lỗi!',
                text: error.message || 'Có lỗi xảy ra khi kiểm tra tờ khai'
            });
        } finally {
            // Reset button state
            submitButton.disabled = false;
            submitButton.innerHTML = originalText;
        }
    });
});

async function loadClients() {
    try {
        const response = await fetch('/check-declaration/api/clients');
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.error || 'Không thể tải danh sách khách hàng');
        }

        const clientSelect = document.getElementById('client');
        clientSelect.innerHTML = '<option value="">Chọn khách hàng...</option>';
        data.clients.forEach(client => {
            const option = document.createElement('option');
            option.value = client;
            option.textContent = client;
            clientSelect.appendChild(option);
        });
    } catch (error) {
        console.error('Error loading clients:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi!',
            text: error.message || 'Không thể tải danh sách khách hàng'
        });
    }
}

function displayResults(result) {
    const resultDiv = document.getElementById('result');
    const missingCount = document.getElementById('missing_count');
    const updatedCount = document.getElementById('updated_count');
    const missingList = document.getElementById('missing_list');
    const updatedList = document.getElementById('updated_list');

    // Update counts
    missingCount.textContent = result.missing_count || 0;
    updatedCount.textContent = result.updated_count || 0;

    // Clear previous results
    missingList.innerHTML = '';
    updatedList.innerHTML = '';

    // Display missing declarations
    if (result.missing_details && result.missing_details.length > 0) {
        result.missing_details.forEach(detail => {
            const li = document.createElement('li');
            li.className = 'list-group-item';
            // Parse the string to extract CDNo, Client, and InvoiceDate
            const match = detail.match(/CDNo: (\S+), Client: (.+?), InvoiceDate: (\S+)/); // Cập nhật regex để xử lý trường hợp Client có nhiều từ
            if (match) {
                const [, cdNo, client, invoiceDate] = match;
                li.textContent = `${cdNo} - ${client} - ${invoiceDate}`;
            } else {
                li.textContent = detail; // Fallback to full string if parsing fails
            }
            missingList.appendChild(li);
        });
    } else {
        const li = document.createElement('li');
        li.className = 'list-group-item';
        li.textContent = 'Không có tờ khai thiếu';
        missingList.appendChild(li);
    }

    // Display updated declarations
    if (result.updated_details && result.updated_details.length > 0) {
        result.updated_details.forEach(detail => {
            const li = document.createElement('li');
            li.className = 'list-group-item';
            // Parse the string to extract CDNo, Client, and InvoiceDate
            const match = detail.match(/CDNo: (\S+), Client: (.+?), InvoiceDate: (\S+)/); // Cập nhật regex
            if (match) {
                const [, cdNo, client, invoiceDate] = match;
                li.textContent = `${cdNo} - ${client} - ${invoiceDate}`;
            } else {
                li.textContent = detail; // Fallback to full string if parsing fails
            }
            updatedList.appendChild(li);
        });
    } else {
        const li = document.createElement('li');
        li.className = 'list-group-item';
        li.textContent = 'Không có tờ khai được cập nhật';
        updatedList.appendChild(li);
    }

    // Show result section
    resultDiv.style.display = 'block';
}