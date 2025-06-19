document.addEventListener('DOMContentLoaded', function () {
  // Khởi tạo datepicker
  $('.datepicker').datepicker({
    format: 'dd/mm/yyyy',
    autoclose: true,
    todayHighlight: true
  });

  // ✅ FUNCTION TẠO SUCCESS MESSAGE VỚI LINK
  function createSuccessMessage(result) {
      let message = result.message;
      let htmlContent = `<div style="text-align: left;">`;
      
      // Basic info
      htmlContent += `<p><strong>📊 Tổng số hóa đơn:</strong> ${result.total_invoices || 0}</p>`;
      
      if (result.storage_type === 'SharePoint' && result.sharepoint_url) {
          // SharePoint success
          htmlContent += `<p><strong>📤 Đã upload:</strong> ${result.successful_uploads || 0} files lên SharePoint</p>`;
          htmlContent += `<p><strong>📋 Excel:</strong> ${result.excel_status || 'Unknown'}</p>`;
          htmlContent += `<hr style="margin: 15px 0;">`;
          htmlContent += `<p><strong>🔗 Truy cập files tại:</strong></p>`;
          htmlContent += `<p><a href="${result.sharepoint_url}" target="_blank" style="color: #007bff; text-decoration: none; background: #f8f9fa; padding: 8px 12px; border-radius: 4px; display: inline-block; margin: 5px 0;">
              <i class="fas fa-external-link-alt"></i> Mở SharePoint Folder
          </a></p>`;
          htmlContent += `<p style="font-size: 12px; color: #6c757d; margin-top: 10px;">
              <i class="fas fa-info-circle"></i> Link sẽ mở trong tab mới. Đăng nhập SharePoint nếu được yêu cầu.
          </p>`;
      } else if (result.storage_type === 'Local' && result.local_path) {
          // Local fallback
          htmlContent += `<p><strong>⚠️ Lưu tạm thời:</strong> ${result.local_path}</p>`;
          htmlContent += `<p style="color: #856404; background: #fff3cd; padding: 8px; border-radius: 4px; font-size: 14px;">
              <i class="fas fa-exclamation-triangle"></i> SharePoint không khả dụng. Files được lưu tạm thời trên server.
          </p>`;
      }
      
      htmlContent += `</div>`;
      
      return {
          title: '🎉 Hoàn thành!',
          html: htmlContent,
          icon: 'success',
          confirmButtonText: 'Đóng',
          width: '600px'
      };
  }

  // ✅ FUNCTION TẠO ERROR MESSAGE CHI TIẾT
  function createErrorMessage(result) {
      let htmlContent = `<div style="text-align: left;">`;
      htmlContent += `<p><strong>❌ Lỗi:</strong> ${result.message}</p>`;
      
      if (result.upload_details && result.upload_details.length > 0) {
          htmlContent += `<hr style="margin: 15px 0;">`;
          htmlContent += `<p><strong>📋 Chi tiết:</strong></p>`;
          htmlContent += `<div style="max-height: 200px; overflow-y: auto; background: #f8f9fa; padding: 10px; border-radius: 4px;">`;
          result.upload_details.forEach(detail => {
              htmlContent += `<p style="margin: 2px 0; font-size: 12px;">${detail}</p>`;
          });
          htmlContent += `</div>`;
      }
      
      htmlContent += `</div>`;
      
      return {
          title: '❌ Có lỗi xảy ra',
          html: htmlContent,
          icon: 'error',
          confirmButtonText: 'Đóng',
          width: '500px'
      };
  }

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
        // Hiển thị thông báo đang tải file với progress
        Swal.fire({
          title: 'Đang xử lý hóa đơn...',
          html: `
              <div style="text-align: left;">
                  <p>🔄 Đang đăng nhập...</p>
                  <p>📋 Lấy danh sách hóa đơn...</p>
                  <p>📄 Tạo PDF files...</p>
                  <p>📤 Upload lên SharePoint...</p>
                  <div style="margin-top: 15px; padding: 10px; background: #f8f9fa; border-radius: 4px;">
                      <small style="color: #6c757d;">⏱️ Quá trình này có thể mất vài phút...</small>
                  </div>
              </div>
          `,
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
        resultDiv.querySelector('#invoice_count').textContent = downloadResult.total_invoices || '0';
        
        // ✅ HIỂN THỊ SUCCESS MESSAGE VỚI LINK
        const successConfig = createSuccessMessage(downloadResult);
        Swal.fire(successConfig);
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
        
        // ✅ HIỂN THỊ ERROR MESSAGE CHI TIẾT
        const errorConfig = createErrorMessage(downloadResult);
        Swal.fire(errorConfig);
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