/**
 * Module xử lý việc chọn và tổng hợp báo cáo theo tháng
 */

/**
 * Lấy danh sách tất cả các sheet báo cáo, bao gồm cả sheet bị ẩn
 * @returns {Array} Danh sách các sheet báo cáo
 */
function getAllReportSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var reportSheets = [];
  
  // Lấy tất cả sheet, bao gồm cả sheet bị ẩn
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Kiểm tra các định dạng báo cáo
    var isPXReport = sheetName.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    
    if (isPXReport) {
      // Trích xuất thông tin tháng/năm
      var monthYear = isPXReport[2] + '/' + isPXReport[3];
      
      reportSheets.push({
        name: sheetName,
        visible: !sheet.isSheetHidden(),
        monthYear: monthYear,
        unitCode: isPXReport[1]
      });
    }
  }
  
  return reportSheets;
}

/**
 * Lọc danh sách sheet theo tháng/năm
 * @param {string} monthYear - Tháng/năm được chọn (MM/YYYY)
 * @returns {Array} Danh sách các sheet phù hợp với tháng/năm đã chọn
 */
function filterSheetsByMonth(monthYear) {
  var allSheets = getAllReportSheets();
  
  if (!monthYear) {
    return allSheets; // Trả về tất cả nếu không có lọc
  }
  
  // Lọc theo tháng/năm
  return allSheets.filter(function(sheet) {
    return sheet.monthYear === monthYear;
  });
}

/**
 * Lấy danh sách sheet báo cáo theo tháng/năm
 * @param {string} monthYear - Tháng/năm được chọn (MM/YYYY)
 * @returns {Array} Danh sách sheet báo cáo phù hợp
 */
function getReportSheetsByMonth(monthYear) {
  return filterSheetsByMonth(monthYear);
}

/**
 * Tổng hợp báo cáo từ các sheet được chọn
 * @param {Object} data - Dữ liệu từ form
 * @returns {boolean} Kết quả xử lý
 */
function consolidateMonthlyReports(data) {
  try {
    // Kiểm tra xem có dữ liệu hợp lệ không
    if (!data || !data.monthYear) {
      throw new Error('Không có thông tin tháng/năm');
    }
    
    // Lấy danh sách sheet đã chọn
    var selectedSheets = data.selectedSheets;
    if (!selectedSheets || selectedSheets.length === 0) {
      throw new Error('Vui lòng chọn ít nhất một đơn vị');
    }
    
    // Gọi hàm xử lý báo cáo hiện có
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var monthYear = data.monthYear;
    
    // Tạo hoặc lấy sheet OUTPUT
    var outputSheet = getOrCreateSummarySheet(monthYear);
    
    // Tìm các sheet INPUT theo tên đã chọn
    var inputSheets = [];
    for (var i = 0; i < selectedSheets.length; i++) {
      var sheet = ss.getSheetByName(selectedSheets[i]);
      if (sheet) {
        inputSheets.push(sheet);
      }
    }
    
    if (inputSheets.length === 0) {
      throw new Error('Không tìm thấy sheet báo cáo nào');
    }
    
    // Tổng hợp dữ liệu từ các sheet đã chọn
    var aggregatedData = aggregateMonthlyData(inputSheets, 12); // Bắt đầu từ dòng 13 (index 12)
    
    // Cập nhật sheet tổng hợp
    updateSummarySheet(outputSheet, aggregatedData, 10); // Bắt đầu từ dòng 11 (index 10)
    
    // Nếu cần, gọi hàm cập nhật dữ liệu sản phẩm
    try {
      copyProductDataToOutput(monthYear);
    } catch (err) {
      console.error('Lỗi khi xử lý dữ liệu sản phẩm:', err);
      // Không làm gián đoạn quy trình chính nếu có lỗi ở đây
    }
    
    return true;
  } catch (error) {
    SpreadsheetApp.getUi().alert('Lỗi: ' + error.message);
    return false;
  }
}