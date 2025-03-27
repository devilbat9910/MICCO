/**
 * Module hỗ trợ xử lý dữ liệu theo tháng
 */

/**
 * Chuyển đổi định dạng tháng từ "YYYY-MM" sang "MM/YYYY"
 * @param {string} monthYearInput - Tháng/năm định dạng "YYYY-MM"
 * @returns {string} - Tháng/năm định dạng "MM/YYYY"
 */
function convertMonthYearFormat(monthYearInput) {
  if (!monthYearInput) {
    throw new Error('Giá trị Tháng/Năm không hợp lệ.');
  }
  
  // Chuyển đổi "YYYY-MM" thành "MM/YYYY"
  const parts = monthYearInput.split("-");
  if (parts.length !== 2) {
    throw new Error('Định dạng Tháng/Năm không hợp lệ.');
  }
  return parts[1] + '/' + parts[0];
}

/**
 * Trích xuất thông tin tháng/năm từ tên sheet báo cáo
 * @param {string} sheetName - Tên sheet (định dạng: PX{unit}_Báo cáo MM/YYYY)
 * @returns {Object} - Thông tin về tháng, năm, mã đơn vị
 */
function extractMonthYearFromSheetName(sheetName) {
  const match = sheetName.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
  if (!match) {
    throw new Error(`Tên sheet không đúng định dạng: ${sheetName}`);
  }
  
  return {
    unitCode: match[1],
    month: parseInt(match[2], 10),
    year: parseInt(match[3], 10),
    monthStr: match[2],
    yearStr: match[3],
    monthYear: `${match[2]}/${match[3]}`
  };
}

/**
 * Tạo tên sheet báo cáo từ thông tin tháng, năm và mã đơn vị
 * @param {string} unitCode - Mã đơn vị (ĐT, CP, QN,...)
 * @param {string|number} month - Tháng (1-12)
 * @param {string|number} year - Năm
 * @returns {string} - Tên sheet báo cáo
 */
function createReportSheetName(unitCode, month, year) {
  const monthStr = String(month).padStart(2, '0');
  return `PX${unitCode}_Báo cáo ${monthStr}/${year}`;
}

/**
 * Tạo tên sheet tổng hợp theo tháng
 * @param {string|number} month - Tháng (1-12)
 * @param {string|number} year - Năm
 * @returns {string} - Tên sheet tổng hợp
 */
function createMonthlySummarySheetName(month, year) {
  const monthStr = String(month).padStart(2, '0');
  return `BC_TCT_${monthStr}/${year}`;
}

/**
 * Tìm các sheet báo cáo theo tháng/năm cụ thể
 * @param {string} monthYear - Tháng/năm định dạng "MM/YYYY"
 * @returns {Array} - Danh sách các sheet báo cáo
 */
function findReportSheetsByMonthYear(monthYear) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  // Xây dựng regex để nhận dạng sheet báo cáo cho tháng đang xét:
  // Ví dụ: "PXNB_Báo cáo 02/2025" khi monthYear là "02/2025"
  const inputSheetRegex = new RegExp(`^PX(ĐT|CP|QN|TB|NB|ĐN|VT)_Báo cáo ${monthYear}$`);
  
  return sheets.filter(function(sheet) {
    return inputSheetRegex.test(sheet.getName());
  });
}

/**
 * Tạo hoặc lấy sheet tổng hợp theo tháng/năm
 * @param {string} monthYear - Tháng/năm định dạng "MM/YYYY"
 * @returns {Sheet} - Sheet tổng hợp
 */
function getOrCreateSummarySheet(monthYear) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheetName = 'BC_TCT_' + monthYear;
  
  // Tìm sheet mẫu
  const templateSheet = spreadsheet.getSheetByName('BC_TCT');
  if (!templateSheet) {
    throw new Error('Không tìm thấy sheet mẫu BC_TCT.');
  }
  
  // Kiểm tra và tạo sheet nếu chưa tồn tại
  let outputSheet = spreadsheet.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = templateSheet.copyTo(spreadsheet).setName(outputSheetName);
    outputSheet.showSheet(); // Hiển thị sheet ngay cả khi sheet gốc bị ẩn
    
    // Chuẩn bị định dạng ngày
    const partsDate = monthYear.split('/');  // monthYear có dạng "MM/YYYY"
    const formattedDate = "THÁNG " + partsDate[0] + " NĂM " + partsDate[1];
    
    // Cập nhật ô tiêu đề với định dạng tháng năm
    outputSheet.getRange('A6:I6').setValue(formattedDate);
  }
  
  return outputSheet;
}

/**
 * Kiểm tra xem một giá trị có là số hợp lệ không
 * @param {*} value - Giá trị cần kiểm tra
 * @returns {boolean} - true nếu là số hợp lệ, false nếu không
 */
function isValidNumber(value) {
  return typeof value === 'number' && !isNaN(value) && isFinite(value);
}

/**
 * Gom dữ liệu từ các sheet tháng vào mảng
 * @param {Array} sheets - Danh sách các sheet cần xử lý
 * @param {number} startRow - Dòng bắt đầu của dữ liệu
 * @returns {Object} - Dữ liệu đã được tổng hợp
 */
function aggregateMonthlyData(sheets, startRow) {
  let aggregatedData = {};
  
  sheets.forEach(function(sheet) {
    var inputData = sheet.getDataRange().getValues();
    for (var i = startRow; i < inputData.length; i++) {
      var key = inputData[i][0] + '|' + inputData[i][1]; // index + target
      
      if (!aggregatedData[key]) {
        aggregatedData[key] = {
          index: inputData[i][0],
          target: inputData[i][1],
          allocation: 0,
          execution: 0
        };
      }
      
      // Cộng dồn giá trị ở cột E (index 4) - Allocation
      if (isValidNumber(inputData[i][4])) {
        aggregatedData[key].allocation += inputData[i][4];
      }
      
      // Cộng dồn giá trị ở cột G (index 6) - Execution
      if (isValidNumber(inputData[i][6])) {
        aggregatedData[key].execution += inputData[i][6];
      }
    }
  });
  
  return aggregatedData;
}

/**
 * Cập nhật sheet tổng hợp với dữ liệu đã gom
 * @param {Sheet} outputSheet - Sheet tổng hợp
 * @param {Object} aggregatedData - Dữ liệu đã được tổng hợp
 * @param {number} startRow - Dòng bắt đầu ghi dữ liệu
 */
function updateSummarySheet(outputSheet, aggregatedData, startRow) {
  // Lấy dữ liệu hiện có của sheet OUTPUT để tìm kiếm các hàng
  var outputData = outputSheet.getDataRange().getValues();
  
  for (var key in aggregatedData) {
    var data = aggregatedData[key];
    
    // Tìm dòng tương ứng trong OUTPUT
    var targetRow = -1;
    for (var i = startRow; i < outputData.length; i++) {
      if (outputData[i][0] === data.index && outputData[i][1] === data.target) {
        targetRow = i + 1; // +1 vì output sheet là 1-indexed
        break;
      }
    }
    
    if (targetRow === -1) {
      Logger.log("Không tìm thấy dòng cho key: " + key);
      continue;
    }
    
    // Cập nhật dữ liệu
    outputSheet.getRange(targetRow, 5).setValue(data.allocation); // Cột E
    outputSheet.getRange(targetRow, 7).setValue(data.execution); // Cột G
  }
}