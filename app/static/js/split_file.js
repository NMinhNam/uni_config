const form = document.getElementById('splitForm');
        const resultDiv = document.getElementById('result');

        // Form Submit Handler
        form.addEventListener('submit', async (e) => {
            e.preventDefault();

            // Get file inputs
            const inputFile = document.getElementById('input_file').files[0];
            const templateFile = document.getElementById('template_file').files[0];

            // Validate file selection
            if (!inputFile || !templateFile) {
                Swal.fire({
                    icon: 'error',
                    title: 'Lỗi',
                    text: 'Vui lòng chọn cả file tổng và file mẫu!'
                });
                return;
            }

            // Prepare form data
            const formData = new FormData();
            formData.append('input_file', inputFile);
            formData.append('template_file', templateFile);

            // Show processing status
            resultDiv.innerHTML = '<p class="text-info">Đang xử lý...</p>';

            try {
                // Send request to server
                const response = await axios.post('/split/api/split-file', formData, {
                    headers: {
                        'Content-Type': 'multipart/form-data'
                    }
                });

                const data = response.data;
                if (data.status === 'success') {
                    // Create download links for processed files
                    let downloadLinks = data.files.map(file => 
                        `<a href="/split/download/${file}" class="text-primary" target="_blank">${file}</a>`
                    ).join('<br>');
                    resultDiv.innerHTML = `<p class="text-success">${data.message}</p>${downloadLinks}`;
                    
                    // Show success message
                    Swal.fire({
                        icon: 'success',
                        title: 'Thành công',
                        text: data.message
                    });
                } else {
                    // Show error message
                    resultDiv.innerHTML = `<p class="text-danger">${data.message}</p>`;
                    Swal.fire({
                        icon: 'error',
                        title: 'Lỗi',
                        text: data.message
                    });
                }
            } catch (error) {
                // Handle request errors
                resultDiv.innerHTML = `<p class="text-danger">Lỗi: ${error.message}</p>`;
                Swal.fire({
                    icon: 'error',
                    title: 'Lỗi',
                    text: 'Đã có lỗi xảy ra, vui lòng thử lại.'
                });
            }
        });