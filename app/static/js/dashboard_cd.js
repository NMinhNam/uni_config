// Thêm biến toàn cục cho phân trang và filter
let currentPage = 1;
let totalPages = 1;
let totalItems = 0;
let isLoading = false;
let allItems = []; // Lưu trữ tất cả items
let isInitialized = false; // Biến kiểm tra đã khởi tạo chưa

// Thêm biến toàn cục cho SSE
let eventSource = null;

// Hàm kiểm tra sự tồn tại của phần tử DOM
function getElement(id) {
    const element = document.getElementById(id);
    if (!element) {
        return null;
    }
    return element;
}

// Hàm khởi tạo các phần tử cần thiết
function initializeElements() {
    if (isInitialized) return true;

    const requiredElements = [
        'debitTableBody',
        'clientFilter',
        'fromDateFilter',
        'toDateFilter',
        'importExportFilter',
        'clearFilter',
        'pagination'
    ];

    const missingElements = requiredElements.filter(id => !getElement(id));
    if (missingElements.length > 0) {
        return false;
    }

    isInitialized = true;
    return true;
}

// Hàm khởi tạo event listeners
function initializeEventListeners() {
    const clientFilter = getElement('clientFilter');
    const fromDateFilter = getElement('fromDateFilter');
    const toDateFilter = getElement('toDateFilter');
    const importExportFilter = getElement('importExportFilter');
    const clearFilter = getElement('clearFilter');

    if (clientFilter) {
        clientFilter.addEventListener('input', handleFilterChange);
    }

    if (fromDateFilter) {
        fromDateFilter.addEventListener('change', handleFilterChange);
    }

    if (toDateFilter) {
        toDateFilter.addEventListener('change', handleFilterChange);
    }

    if (importExportFilter) {
        importExportFilter.addEventListener('change', handleFilterChange);
    }

    if (clearFilter) {
        clearFilter.addEventListener('click', function() {
            if (clientFilter) clientFilter.value = '';
            if (fromDateFilter) fromDateFilter.value = '';
            if (toDateFilter) toDateFilter.value = '';
            if (importExportFilter) importExportFilter.value = '';
            handleFilterChange();
        });
    }
}

// Hàm lấy ngày 30 ngày trước
function getDate30DaysAgo() {
    const today = new Date();
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(today.getDate() - 30);
    return thirtyDaysAgo;
}

// Hàm format ngày thành DD/MM/YYYY
function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

// Hàm kiểm tra xem một ngày có nằm trong khoảng thời gian không
function isDateInRange(dateStr, startDate, endDate) {
    if (!dateStr) return false;
    const [day, month, year] = dateStr.split('/').map(Number);
    const date = new Date(year, month - 1, day);
    return date >= startDate && date <= endDate;
}

// Thêm hàm cuộn bảng
function scrollTable(direction) {
    const container = document.querySelector('.table-container');
    const scrollAmount = 200; // Số pixel cuộn mỗi lần
    if (direction === 'left') {
        container.scrollLeft -= scrollAmount;
    } else {
        container.scrollLeft += scrollAmount;
    }
}

// Hàm lọc dữ liệu theo các điều kiện
function filterItems(items) {
    const clientFilterValue = getElement('clientFilter')?.value.toLowerCase().trim() || '';
    const fromDateValue = getElement('fromDateFilter')?.value || '';
    const toDateValue = getElement('toDateFilter')?.value || '';
    const importExportFilterElement = getElement('importExportFilter');
    const importExportValue = importExportFilterElement?.value.toLowerCase().trim() || '';

    const startDate30DaysAgo = getDate30DaysAgo();
    const endDateToday = new Date();
    endDateToday.setHours(23, 59, 59, 999); // Set to end of today

    // Kiểm tra xem người dùng có đặt filter ngày tùy chỉnh không
    const userHasCustomDateFilter = fromDateValue !== '' || toDateValue !== '';

    let fromDateFilterObj = null;
    let toDateFilterObjInclusive = null;

    // Xử lý From Date và To Date từ input filter
    if (userHasCustomDateFilter) {
        if (fromDateValue) {
            // Parse YYYY-MM-DD và tạo Date object tại 00:00:00 local time
            const [y, m, d] = fromDateValue.split('-').map(Number);
            fromDateFilterObj = new Date(y, m - 1, d);
            fromDateFilterObj.setHours(0, 0, 0, 0);
        }
        if (toDateValue) {
            // Parse YYYY-MM-DD và tạo Date object tại 00:00:00 local time
            const [y, m, d] = toDateValue.split('-').map(Number);
            toDateFilterObjInclusive = new Date(y, m - 1, d);
            // Set to the very end of the 'To Date'
            toDateFilterObjInclusive.setHours(23, 59, 59, 999);
        }
        console.log('filterItems: Using custom date range', fromDateFilterObj, 'to', toDateFilterObjInclusive);
    } else {
        // Default filter for last 30 days
        fromDateFilterObj = new Date(startDate30DaysAgo);
        fromDateFilterObj.setHours(0, 0, 0, 0); // Ensure start of day

        toDateFilterObjInclusive = new Date(endDateToday);
        toDateFilterObjInclusive.setHours(23, 59, 59, 999); // Ensure end of day
        console.log('filterItems: Using default 30-day range', fromDateFilterObj, 'to', toDateFilterObjInclusive);
    }

    return items.filter(item => {
        // Lấy ngày của item (định dạng DD/MM/YYYY)
        const itemDateStr = item.InvoiceDate;

        // Nếu item không có ngày InvoiceDate hoặc sai định dạng, bỏ qua
        if (!itemDateStr) {
            return false;
        }

        // Phân tích ngày của item
        const dateParts = itemDateStr.split('/');
        if (dateParts.length !== 3) {
            return false;
        }
        const [dayStr, monthStr, yearStr] = dateParts;

        // Chuyển các phần của ngày thành số để so sánh chính xác
        const itemDayNum = parseInt(dayStr);
        const itemMonthNum = parseInt(monthStr);
        const itemYearNum = parseInt(yearStr);

        // Kiểm tra tính hợp lệ của số ngày, tháng, năm
        if (isNaN(itemDayNum) || isNaN(itemMonthNum) || isNaN(itemYearNum)) {
            return false;
        }

        // Chuyển đổi ngày của item thành đối tượng Date tại 00:00:00 local time
        const itemDate = new Date(itemYearNum, itemMonthNum - 1, itemDayNum);
        itemDate.setHours(0, 0, 0, 0); 

        // 1. Kiểm tra Filter Client (áp dụng nếu có giá trị)
        const clientMatch = !clientFilterValue || item.Client?.toLowerCase()?.includes(clientFilterValue);
        
        // 2. Kiểm tra Filter Import/Export (áp dụng nếu có giá trị)
        const rawItemImportExportValue = item['Import_x002f_Export'] || '';
        const processedItemImportExportValue = rawItemImportExportValue.toLowerCase().trim();
        
        let normalizedItemImportExportValue = '';
        if (processedItemImportExportValue === 'im') {
            normalizedItemImportExportValue = 'import';
        } else if (processedItemImportExportValue === 'ex') {
            normalizedItemImportExportValue = 'export';
        }

        const importExportMatch = importExportValue === '' || normalizedItemImportExportValue === importExportValue;

        // 3. Kiểm tra Filter Ngày
        let dateMatch = true; 

        if (fromDateFilterObj && itemDate < fromDateFilterObj) {
            dateMatch = false;
        }
        if (toDateFilterObjInclusive && itemDate > toDateFilterObjInclusive) {
            dateMatch = false;
        }

        // Item được giữ lại chỉ khi khớp tất cả các điều kiện
        return clientMatch && dateMatch && importExportMatch;
    });
}

// Hàm hiển thị loading
function showLoading() {
    const tableBody = getElement('debitTableBody');
    if (tableBody) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="24" class="text-center">
                    <div class="loading-spinner" style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 20px;">
                        <div class="spinner-border text-primary" role="status" style="width: 3rem; height: 3rem;">
                            <span class="visually-hidden">Loading...</span>
                        </div>
                        <p class="mt-2" style="margin-top: 10px;">Đang tải dữ liệu...</p>
                    </div>
                </td>
            </tr>
        `;
    }
}

// Hàm ẩn loading
function hideLoading() {
    const loadingSpinner = document.querySelector('.loading-spinner');
    if (loadingSpinner) {
        loadingSpinner.style.display = 'none';
    }
}

// Fetch data from Flask backend
async function fetchDebitList() {
    try {
        if (!initializeElements()) {
            throw new Error('Không thể khởi tạo các phần tử cần thiết');
        }

        isLoading = true;
        showLoading();

        console.log('fetchDebitList: Starting fetch, allItems.length =', allItems.length);

        // Lấy tất cả dữ liệu nếu chưa có. Chỉ fetch 1 lần duy nhất.
        if (allItems.length === 0) {
            console.log('fetchDebitList: allItems is empty, fetching data from API');
            const response = await axios.get('/dashboard/api/debit_list?per_page=5000'); // Lấy tối đa 5000 items
            if (response.data.error) {
                throw new Error(response.data.error);
            }
            allItems = response.data.items || [];
            console.log('fetchDebitList: Fetched', allItems.length, 'items from API');
        }

        // Lọc dữ liệu từ toàn bộ danh sách dựa trên filter hiện tại (mặc định hoặc người dùng)
        const filteredItems = filterItems(allItems);

        console.log('fetchDebitList: Filtered down to', filteredItems.length, 'items');

        // Cập nhật biến toàn cục dựa trên dữ liệu đã filter
        totalItems = filteredItems.length;
        totalPages = Math.ceil(totalItems / 10);
        currentPage = Math.min(currentPage, totalPages);

        console.log('fetchDebitList: Rendering page', currentPage, 'of', totalPages, 'with', totalItems, 'total items');

        // Lấy items cho trang hiện tại từ danh sách đã filter
        const startIndex = (currentPage - 1) * 10;
        const endIndex = startIndex + 10;
        const pageItems = filteredItems.slice(startIndex, endIndex);

        // Render bảng và phân trang
        renderTable(pageItems);
        renderPaginationControls();

        // Cập nhật dropdowns với dữ liệu đã lọc
        if (filteredItems.length > 0) {
            populateDropdowns(filteredItems);
        } else {
            // Nếu không có dữ liệu sau filter, reset dropdowns
            populateSelect('itemDebitStatus', []);
            populateSelect('itemPhanLuong', []);
            populateSelect('itemClient', []);
        }

        console.log('fetchDebitList: Rendering complete');

        return {
            items: pageItems,
            total_items: totalItems,
            total_pages: totalPages,
            current_page: currentPage
        };

    } catch (error) {
        console.error('fetchDebitList: Lỗi khi gọi API debit_list:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi!',
            text: error.message || 'Đã xảy ra lỗi khi tải danh sách Debit List.',
        });

        const tableBody = getElement('debitTableBody');
        if (tableBody) {
            tableBody.innerHTML = '';
        }
        
        totalItems = 0;
        totalPages = 0;
        currentPage = 1;
        renderPaginationControls();
    } finally {
        isLoading = false;
        hideLoading();
        console.log('fetchDebitList: Finished');
    }
}

async function fetchDebitStatistics() {
    try {
        const response = await axios.get('/dashboard/api/debit_statistics');
        return response.data;
    } catch (error) {
        console.error('Lỗi khi gọi API debit_statistics:', error);
        throw new Error(error.response?.data?.error || error.message || 'Network Error');
    }
}

// Hàm gọi khi filter thay đổi (client, date, month, clear)
function handleFilterChange() {
    currentPage = 1; // Luôn reset về trang 1 khi filter thay đổi
    fetchDebitList(); // Fetch lại dữ liệu với filter mới
}

// Cập nhật hàm filterTable (có thể xóa hoặc đổi tên thành handleFilterChange)
// Giữ lại hàm filterTable để tương thích với các event listener đã đặt
const filterTable = handleFilterChange;

// Hàm khởi tạo SSE
function initializeSSE() {
    if (eventSource) {
        console.log('Closing existing SSE connection');
        eventSource.close();
    }

    console.log('Initializing new SSE connection to /dashboard/api/events');
    eventSource = new EventSource('/dashboard/api/events');
    
    eventSource.onmessage = async function(event) {
        try {
            const data = JSON.parse(event.data);
            if (data.type === 'update') {
                console.log('SSE: Received update event', data);
                // Reset allItems để fetch lại dữ liệu mới
                allItems = [];
                console.log('SSE: Reset allItems, fetching new data and statistics');
                // Cập nhật bảng và thống kê
                await Promise.all([
                    fetchDebitList(),
                    fetchDebitStatistics().then(stats => {
                        updateDebitDashboard(stats);
                    })
                ]);
                console.log('SSE: Data and statistics updated');
            }
        } catch (error) {
            console.error('SSE: Lỗi khi xử lý message:', error);
        }
    };

    eventSource.onerror = function(error) {
        console.error('SSE Error:', error);
        // Đóng kết nối hiện tại
        eventSource.close();
        console.log('SSE Error: Attempting to reconnect in 60 seconds');
        // Thử kết nối lại sau 5 giây
        setTimeout(initializeSSE, 60000);
    };

    eventSource.onopen = function() {
        console.log('SSE connection opened');
    };

    eventSource.onclose = function() {
        console.log('SSE connection closed');
    };
}

// Cập nhật hàm updateItem để cập nhật last_update_time
async function updateItem() {
    const itemId = document.getElementById('itemId').value;
    const updatedData = {
        "Invoice": document.getElementById('itemInvoice').value,
        "Client": document.getElementById('itemClient').value,
        "Term": document.getElementById('itemTerm').value,
        "CD Date": document.getElementById('itemCDDate').value,
        "InvoiceDate": document.getElementById('itemInvoiceDate').value,
        "ParkingList": document.getElementById('itemParkingList').value,
        "DebitStatus": document.getElementById('itemDebitStatus').value,
        "Import_x002f_Export": document.getElementById('itemImportExport').value,
        "CD No": document.getElementById('itemCDNo').value,
        "Ma dia diem do hang": document.getElementById('itemMaDiaDiemDoHang').value,
        "Ma dia diem xep hang": document.getElementById('itemMaDiaDiemXepHang').value,
        "Tong tri gia": document.getElementById('itemTongTriGia').value,
        "Ma loai hinh": document.getElementById('itemMaLoaiHinh').value,
        "Fist CD No": document.getElementById('itemFirstCDNo').value,
        "Loai CD": document.getElementById('itemLoaiCD').value,
        "Phan luong": document.getElementById('itemPhanLuong').value,
        "HBL": document.getElementById('itemHBL').value,
        "So luong kien": document.getElementById('itemSoLuongKien').value,
        "tong trong luong": document.getElementById('itemTongTrongLuong').value,
        "so luong cont": document.getElementById('itemSoLuongCont').value,
        "cont number": document.getElementById('itemContNumber').value,
        "ETA": document.getElementById('itemETA').value,
        "ETD": document.getElementById('itemETD').value
    };

    try {
        const response = await axios.post(`/dashboard/api/update_item/${itemId}`, updatedData);

        Swal.fire({
            icon: 'success',
            title: 'Cập nhật thành công!',
            text: response.data.message,
            timer: 1500,
            showConfirmButton: false
        });

        // Cập nhật bảng và thống kê ngay lập tức
        await Promise.all([
            fetchDebitList(),
            fetchDebitStatistics().then(stats => {
                updateDebitDashboard(stats);
            })
        ]);

    } catch (error) {
        console.error('Lỗi khi cập nhật:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi cập nhật',
            text: error.response?.data?.error || error.message || 'Network Error',
        });
    }
}

// Function to load dashboard specific data
async function loadDashboardData(dashboardId) {
    try {
        let response;
        switch(dashboardId) {
            case 'debit':
                response = await fetchDebitStatistics();
                updateDebitDashboard(response);
                break;
            case 'hr':
                response = await fetchHRStatistics();
                updateHRDashboard(response);
                break;
        }
    } catch (error) {
        console.error('Error loading dashboard data:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Không thể tải dữ liệu dashboard: ' + error.message
        });
    }
}

// Update functions for each dashboard
function updateDebitDashboard(data) {
    document.querySelector('#debit-dashboard .dashboard-card:nth-child(1) p').textContent = data.debit_count;
    document.querySelector('#debit-dashboard .dashboard-card:nth-child(2) p').textContent = data.not_debit_count;
    document.querySelector('#debit-dashboard .dashboard-card:nth-child(3) p').textContent = data.client_count;
    document.querySelector('#debit-dashboard .dashboard-card:nth-child(4) p').textContent = data.invoice_count;
}

function updateHRDashboard(data) {
    // Update HR dashboard cards
    document.querySelector('#hr-dashboard .dashboard-card:nth-child(1) p').textContent = data.total_employees;
    document.querySelector('#hr-dashboard .dashboard-card:nth-child(2) p').textContent = data.new_employees;
    document.querySelector('#hr-dashboard .dashboard-card:nth-child(3) p').textContent = data.on_leave;
    document.querySelector('#hr-dashboard .dashboard-card:nth-child(4) p').textContent = data.open_positions;
}

// Add new API fetch functions
async function fetchHRStatistics() {
    try {
        const response = await axios.get('/dashboard/api/hr_statistics');
        return response.data;
    } catch (error) {
        throw new Error(error.response?.data?.error || error.message || 'Network Error');
    }
}

// Function to populate days based on selected month
function populateDays() {
    const monthSelect = getElement('monthFilter');
    const daySelect = getElement('dateFilter');
    
    if (!monthSelect || !daySelect) {
        return;
    }

    const selectedMonth = parseInt(monthSelect.value);
    const currentYear = new Date().getFullYear();
    
    // Clear existing options except the first one ("Select day")
    while (daySelect.options.length > 1) {
        daySelect.remove(1);
    }
    
    // Add options for each day
    // Add days 1-31 if no month is selected, otherwise add days for the selected month
    const daysToAdd = selectedMonth ? new Date(new Date().getFullYear(), selectedMonth, 0).getDate() : 31;

    for (let i = 1; i <= daysToAdd; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        daySelect.appendChild(option);
    }
}

// Add event listener for month selection
const monthFilter = getElement('monthFilter');
if (monthFilter) {
    monthFilter.addEventListener('change', populateDays);
}

// Initialize days when page loads
document.addEventListener('DOMContentLoaded', function() {
    populateDays();
});

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', async () => {
    try {
        await new Promise(resolve => setTimeout(resolve, 100));

        // Khởi tạo các phần tử
        if (!initializeElements()) {
            throw new Error('Không thể khởi tạo các phần tử cần thiết');
        }

        // Khởi tạo event listeners
        initializeEventListeners();

        // Reset các filter về giá trị mặc định
        const clientFilter = getElement('clientFilter');
        const fromDateFilter = getElement('fromDateFilter');
        const toDateFilter = getElement('toDateFilter');
        const importExportFilter = getElement('importExportFilter');
        
        if (clientFilter) clientFilter.value = '';
        if (fromDateFilter) fromDateFilter.value = '';
        if (toDateFilter) toDateFilter.value = '';
        if (importExportFilter) importExportFilter.value = '';

        // Fetch debit list với filter mặc định (30 ngày)
        const data = await fetchDebitList();

        // Fetch and display statistics
        const stats = await fetchDebitStatistics();
        updateDebitDashboard(stats);

        // Khởi tạo SSE
        initializeSSE();

        Swal.fire({
            icon: 'success',
            title: 'Tải dữ liệu thành công',
            text: 'Danh sách Debit List đã được tải thành công!',
            timer: 1500,
            showConfirmButton: false
        });
    } catch (error) {
        console.error('Lỗi tổng thể trong DOMContentLoaded:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Không thể tải dữ liệu: ' + error.message,
        });
        
        const tableBody = getElement('debitTableBody');
        if (tableBody) {
            tableBody.innerHTML = '';
        }
        
        totalItems = 0;
        totalPages = 0;
        currentPage = 1;
        renderPaginationControls();
    }

    // Add event listeners for dashboard switching
    const dashboardSelectors = document.querySelectorAll('.dashboard-selector');
    const dashboardContents = document.querySelectorAll('.dashboard-content');
    const navLinks = document.querySelectorAll('.nav-link');

    dashboardSelectors.forEach(selector => {
        selector.addEventListener('click', function(e) {
            e.preventDefault();
            const dashboardId = this.getAttribute('data-dashboard');

            // Hide all dashboards
            dashboardContents.forEach(content => {
                content.classList.remove('active');
            });

            // Show selected dashboard
            document.getElementById(`${dashboardId}-dashboard`).classList.add('active');

            // Update active state in menu
            navLinks.forEach(link => {
                link.classList.remove('active');
            });
            this.closest('.nav-link').classList.add('active');

            // Load dashboard specific data
            loadDashboardData(dashboardId);
        });
    });

    // Add event listener to the form submission
    document.getElementById('itemUpdateForm').addEventListener('submit', async (event) => {
        event.preventDefault(); // Prevent default form submission
        updateItem(); // Call the update function
    });
});

// Function to render the table body
function renderTable(items) {
    const tableBody = getElement('debitTableBody');
    if (!tableBody) {
        return;
    }

    tableBody.innerHTML = ''; // Clear existing rows

    if (!items || items.length === 0) {
        return;
    }

    items.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.ID || ''}</td>
            <td>${item.Invoice || ''}</td>
            <td>${item.InvoiceDate || ''}</td>
            <td>${item.ParkingList || ''}</td>
            <td>${item.DebitStatus || ''}</td>
            <td>${item.Client || ''}</td>
            <td>${item['Import_x002f_Export'] || ''}</td>
            <td>${item['CD No'] || ''}</td>
            <td>${item.Term || ''}</td>
            <td>${item['Ma dia diem do hang'] || ''}</td>
            <td>${item['Ma dia diem xep hang'] || ''}</td>
            <td>${item['Tong tri gia'] || ''}</td>
            <td>${item['Ma loai hinh'] || ''}</td>
            <td>${item['Fist CD No'] || ''}</td>
            <td>${item['Loai CD'] || ''}</td>
            <td>${item['Phan luong'] || ''}</td>
            <td>${item.HBL || ''}</td>
            <td>${item['CD Date'] || ''}</td>
            <td>${item['So luong kien'] || ''}</td>
            <td>${item['tong trong luong'] || ''}</td>
            <td>${item['so luong cont'] || ''}</td>
            <td>${item['cont number'] || ''}</td>
            <td>${item.ETA || ''}</td>
            <td>${item.ETD || ''}</td>
        `;
        row.dataset.item = JSON.stringify(item);
        tableBody.appendChild(row);
    });

    // Add click event listeners to new rows
    document.querySelectorAll('#debitTableBody tr').forEach(row => {
        row.addEventListener('click', () => {
            const item = JSON.parse(row.dataset.item);
            // Populate all input and select fields
            const fields = {
                'itemId': item.ID,
                'itemInvoice': item.Invoice,
                'itemInvoiceDate': item.InvoiceDate,
                'itemParkingList': item.ParkingList,
                'itemDebitStatus': item.DebitStatus,
                'itemClient': item.Client,
                'itemImportExport': item['Import_x002f_Export'],
                'itemCDNo': item['CD No'],
                'itemTerm': item.Term,
                'itemMaDiaDiemDoHang': item['Ma dia diem do hang'],
                'itemMaDiaDiemXepHang': item['Ma dia diem xep hang'],
                'itemTongTriGia': item['Tong tri gia'],
                'itemMaLoaiHinh': item['Ma loai hinh'],
                'itemFirstCDNo': item['Fist CD No'],
                'itemLoaiCD': item['Loai CD'],
                'itemPhanLuong': item['Phan luong'],
                'itemHBL': item.HBL,
                'itemCDDate': item['CD Date'],
                'itemSoLuongKien': item['So luong kien'],
                'itemTongTrongLuong': item['tong trong luong'],
                'itemSoLuongCont': item['so luong cont'],
                'itemContNumber': item['cont number'],
                'itemETA': item.ETA,
                'itemETD': item.ETD
            };

            for (const [id, value] of Object.entries(fields)) {
                const element = getElement(id);
                if (element) {
                    element.value = value || '';
                }
            }
        });
    });
}

// Render pagination controls
function renderPaginationControls() {
    const paginationContainer = getElement('pagination');
    if (!paginationContainer) {
        return;
    }
    paginationContainer.innerHTML = '';

    // Previous button
    const prevButton = document.createElement('button');
    prevButton.innerHTML = '&laquo; Previous';
    prevButton.className = 'pagination-item';
    // Disable Previous nếu ở trang 1 hoặc đang loading hoặc không có item nào
    prevButton.disabled = currentPage <= 1 || isLoading || totalItems === 0;
    prevButton.onclick = () => {
        if (currentPage > 1 && !isLoading && totalItems > 0) {
            currentPage--;
            fetchDebitList(); // Fetch lại dữ liệu (đã filter) cho trang mới
        }
    };
    paginationContainer.appendChild(prevButton);

    // Page input
    const pageInput = document.createElement('input');
    pageInput.type = 'number';
    // Đảm bảo min/max/value hợp lệ ngay cả khi totalPages/currentPage là 0
    pageInput.min = 1;
    pageInput.max = totalPages > 0 ? totalPages : 1; 
    pageInput.value = currentPage > 0 && currentPage <= totalPages ? currentPage : 1;
    pageInput.className = 'pagination-input';
    // Disable input nếu chỉ có 1 trang hoặc không có item nào
    pageInput.disabled = totalPages <= 1 || totalItems === 0;
    pageInput.onchange = (e) => {
        const newPage = parseInt(e.target.value);
        // Kiểm tra tính hợp lệ của newPage, nằm trong phạm vi trang và không đang loading
        if (!isNaN(newPage) && newPage >= 1 && newPage <= totalPages && !isLoading) {
            currentPage = newPage;
            fetchDebitList(); // Fetch lại dữ liệu (đã filter) cho trang mới
        } else {
            // Đặt lại giá trị input nếu không hợp lệ
            e.target.value = currentPage > 0 && currentPage <= totalPages ? currentPage : 1;
        }
    };
    paginationContainer.appendChild(pageInput);

    // Page info
    const pageInfo = document.createElement('span');
    pageInfo.className = 'pagination-info';
    pageInfo.textContent = `of ${totalPages > 0 ? totalPages : 0}`; // Hiển thị 0 trang nếu không có item
    paginationContainer.appendChild(pageInfo);

    // Next button
    const nextButton = document.createElement('button');
    nextButton.innerHTML = 'Next &raquo;';
    nextButton.className = 'pagination-item';
    // Disable Next nếu ở trang cuối hoặc đang loading hoặc không có item nào
    nextButton.disabled = currentPage >= totalPages || isLoading || totalItems === 0;
     nextButton.onclick = () => {
        if (currentPage < totalPages && !isLoading && totalItems > 0) {
            currentPage++;
            fetchDebitList(); // Fetch lại dữ liệu (đã filter) cho trang mới
        }
    };
    paginationContainer.appendChild(nextButton);

    // Items info
    const itemsInfo = document.createElement('div');
    itemsInfo.className = 'pagination-items-info';
    
    // Tính toán start và end dựa trên totalItems và currentPage
    const start = totalItems === 0 ? 0 : (currentPage - 1) * 10 + 1;
    const end = totalItems === 0 ? 0 : Math.min(currentPage * 10, totalItems);
    
    // Hiển thị thông tin số mục
    itemsInfo.textContent = `Showing ${start} to ${end} of ${totalItems} items`;
    
    paginationContainer.appendChild(itemsInfo);

    // Kiểm tra và điều chỉnh currentPage nếu nó lớn hơn totalPages sau khi filter
    // Điều này có thể xảy ra nếu filter làm giảm đáng kể số lượng mục
     if (totalItems > 0 && currentPage > totalPages) {
         console.warn(`Current page (${currentPage}) exceeds total pages (${totalPages}) after filter. Resetting to page 1.`);
         currentPage = 1;
         // Gọi lại fetchDebitList để hiển thị trang 1 của dữ liệu đã filter
         fetchDebitList(); 
     } else if (totalItems > 0 && currentPage === 0) { // Đảm bảo trang không bao giờ là 0 khi có items
          console.warn(`Current page is 0 with total items > 0. Resetting to page 1.`);
          currentPage = 1;
          fetchDebitList();
     }
}

// Function to populate dropdowns with unique values from data
function populateDropdowns(data) {
    const uniqueDebitStatuses = new Set();
    const uniquePhanLuongs = new Set();
    const uniqueClients = new Set();

    if (data && Array.isArray(data)) { // Thêm kiểm tra data là mảng
        data.forEach(item => {
            if (item.DebitStatus) uniqueDebitStatuses.add(item.DebitStatus);
            if (item['Phan luong']) uniquePhanLuongs.add(item['Phan luong']);
            if (item.Client) uniqueClients.add(item.Client);
        });
    }

    // Convert Sets to sorted Arrays and populate dropdowns
    populateSelect('itemDebitStatus', Array.from(uniqueDebitStatuses).sort());
    populateSelect('itemPhanLuong', Array.from(uniquePhanLuongs).sort());
    populateSelect('itemClient', Array.from(uniqueClients).sort());
}

// Helper function to populate a select element
function populateSelect(selectId, options) {
    const selectElement = document.getElementById(selectId);
    selectElement.innerHTML = ''; // Clear existing options

    // Add a default empty option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = `-- Select ${selectId.replace('item', '')} --`;
    selectElement.appendChild(defaultOption);

    options.forEach(optionText => {
        const optionElement = document.createElement('option');
        optionElement.value = optionText;
        optionElement.textContent = optionText;
        selectElement.appendChild(optionElement);
    });
}

// Thêm cleanup khi trang bị đóng
window.addEventListener('beforeunload', () => {
    if (eventSource) {
        eventSource.close();
    }
}); 