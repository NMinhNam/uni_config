document.addEventListener('DOMContentLoaded', function () {
    // Khởi tạo datepicker
    $('.datepicker').datepicker({
      format: 'dd/mm/yyyy',
      autoclose: true,
      todayHighlight: true
    });

    async function showCaptchaPopup(company, username, password, from_date, to_date) {
      let captchaResult;
      try {
        // Hiển thị thông báo đang lấy CAPTCHA
        Swal.fire({
          title: 'Đang lấy ảnh CAPTCHA...',
          allowOutsideClick: false,
          didOpen: () => {
            Swal.showLoading();
          }
        });

        captchaResult = await fetch('/api/get-captcha', {
          method: 'GET'
        }).then(response => response.json());

        // Đóng thông báo loading
        Swal.close();
      } catch (e) {
        Swal.fire({
          icon: 'error',
          title: 'Lỗi',
          text: 'Không thể tải ảnh CAPTCHA: ' + e.message
        });
        return;
      }

      if (captchaResult.status !== 'success') {
        Swal.fire({
          icon: 'error',
          title: 'Lỗi',
          text: captchaResult.message
        });
        return;
      }

      let currentCaptchaUrl = captchaResult.captcha_url;

      while (true) {
        const result = await Swal.fire({
          title: 'Nhập mã CAPTCHA',
          imageUrl: currentCaptchaUrl,
          imageAlt: 'CAPTCHA Image',
          html: '<input type="text" id="captcha-input" class="swal2-input" placeholder="Nhập mã CAPTCHA">',
          confirmButtonText: 'Xác nhận',
          showCancelButton: true,
          cancelButtonText: 'Hủy',
          preConfirm: () => {
            const captchaCode = document.getElementById('captcha-input').value;
            if (!captchaCode) {
              Swal.showValidationMessage('Vui lòng nhập mã CAPTCHA!');
            }
            return captchaCode;
          }
        });

        if (!result.isConfirmed) {
          return;
        }

        const captchaCode = result.value;

        const resultDiv = document.getElementById('result');
        if (!resultDiv) {
          console.error('Phần tử #result không tồn tại trong DOM');
          Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Không tìm thấy phần tử kết quả trên trang'
          });
          return;
        }

        let downloadResult;
        try {
          // Hiển thị thông báo đang tải file
          Swal.fire({
            title: 'Đang tải hóa đơn...',
            html: 'Vui lòng đợi trong giây lát...',
            allowOutsideClick: false,
            didOpen: () => {
              Swal.showLoading();
            }
          });

          downloadResult = await fetch('/api/download-invoices', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: new URLSearchParams({
              'company': company,
              'username': username,
              'password': password,
              'from_date': from_date,
              'to_date': to_date,
              'captcha_code': captchaCode
            })
          }).then(response => response.json());

          // Đóng thông báo loading
          Swal.close();
        } catch (e) {
          Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Lỗi khi tải hóa đơn: ' + e.message
          });
          return;
        }

        if (downloadResult.status === 'success') {
          resultDiv.style.display = 'block';
          resultDiv.className = 'success';
          resultDiv.querySelector('#invoice_count').textContent = downloadResult.message.match(/\d+/)?.[0] || '0';
          Swal.fire({
            icon: 'success',
            title: 'Thành công',
            text: downloadResult.message
          });
          break;
        } else if (downloadResult.status === 'captcha_error') {
          await Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Mã CAPTCHA sai, đang làm mới CAPTCHA...',
            timer: 1500,
            showConfirmButton: false
          });

          // Hiển thị thông báo đang lấy CAPTCHA mới
          Swal.fire({
            title: 'Đang lấy ảnh CAPTCHA mới...',
            allowOutsideClick: false,
            didOpen: () => {
              Swal.showLoading();
            }
          });

          // Lấy CAPTCHA mới
          try {
            captchaResult = await fetch('/api/get-captcha', {
              method: 'GET'
            }).then(response => response.json());
            
            // Đóng thông báo loading
            Swal.close();
            
            if (captchaResult.status === 'success') {
              currentCaptchaUrl = captchaResult.captcha_url;
              console.log("New CAPTCHA URL:", currentCaptchaUrl);
            } else {
              Swal.fire({
                icon: 'error',
                title: 'Lỗi',
                text: 'Không thể làm mới CAPTCHA: ' + captchaResult.message
              });
              return;
            }
          } catch (e) {
            Swal.fire({
              icon: 'error',
              title: 'Lỗi',
              text: 'Không thể làm mới CAPTCHA: ' + e.message
            });
            return;
          }
        } else {
          resultDiv.style.display = 'block';
          resultDiv.className = 'error';
          resultDiv.querySelector('#invoice_count').textContent = '0';
          Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: downloadResult.message
          });
          break;
        }
      }
    }

    window.submitForm = function () {
      const resultDiv = document.getElementById('result');
      if (!resultDiv) {
        console.error('Phần tử #result không tồn tại trong DOM');
        Swal.fire({
          icon: 'error',
          title: 'Lỗi',
          text: 'Không tìm thấy phần tử kết quả trên trang'
        });
        return;
      }

      const company = document.getElementById('company').value;
      const username = document.getElementById('username').value;
      const password = document.getElementById('password').value;
      const from_date = document.getElementById('from_date').value;
      const to_date = document.getElementById('to_date').value;

      if (!company || !username || !password || !from_date || !to_date) {
        resultDiv.style.display = 'block';
        resultDiv.className = 'error';
        resultDiv.querySelector('#invoice_count').textContent = '0';
        Swal.fire({
          icon: 'error',
          title: 'Lỗi',
          text: 'Vui lòng điền đầy đủ các trường!'
        });
        return;
      }

      const datePattern = /^\d{2}\/\d{2}\/\d{4}$/;
      if (!datePattern.test(from_date) || !datePattern.test(to_date)) {
        resultDiv.style.display = 'block';
        resultDiv.className = 'error';
        resultDiv.querySelector('#invoice_count').textContent = '0';
        Swal.fire({
          icon: 'error',
          title: 'Lỗi',
          text: 'Ngày phải có định dạng DD/MM/YYYY!'
        });
        return;
      }

      showCaptchaPopup(company, username, password, from_date, to_date);
    };
  });