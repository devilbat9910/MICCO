/**
 * Hiển thị giao diện HTML cho chức năng "Tổng hợp báo cáo theo tháng".
 */
function summarizeMonthlyReports() {
  var html = HtmlService.createHtmlOutputFromFile('html_monthly_report_selector')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Tổng hợp báo cáo theo tháng');
  SpreadsheetApp.getUi().showModalDialog(html, 'Tổng hợp báo cáo theo tháng');
}

/**
 * Xử lý dữ liệu nhận từ giao diện HTML.
 * @param {Object} data - Dữ liệu nhận từ form HTML
 * @returns {boolean} - Kết quả xử lý
 */
function summarizeMonthlyReportsHtml(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // Lấy giá trị monthYear từ HTML (ví dụ: "2025-02")
  var monthYearInput = data.monthYear;
  if (!monthYearInput) {
    ui.alert('Giá trị Tháng/Năm không hợp lệ.');
    return false;
  }
  
  // Chuyển đổi "YYYY-MM" thành "MM/YYYY"
  var parts = monthYearInput.split("-");
  if (parts.length !== 2) {
    ui.alert('Định dạng Tháng/Năm không hợp lệ.');
    return false;
  }
  var monthYear = parts[1] + '/' + parts[0];
  
  // Tạo hoặc lấy sheet OUTPUT từ mẫu "BC_TCT"
  var outputSheetName = 'BC_TCT_' + monthYear;
  var templateSheet = ss.getSheetByName('BC_TCT');
  if (!templateSheet) {
    ui.alert('Không tìm thấy sheet mẫu BC_TCT.');
    return false;
  }
  
  // Chuẩn bị định dạng ngày
  var partsDate = monthYear.split('/');  // monthYear có dạng "MM/YYYY"
  var formattedDate = "THÁNG " + partsDate[0] + " NĂM " + partsDate[1];
  
  // Kiểm tra và tạo sheet nếu chưa tồn tại
  var outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = templateSheet.copyTo(ss).setName(outputSheetName);
    outputSheet.showSheet(); // Hiển thị sheet ngay cả khi sheet gốc bị ẩn
  }
  
  // Cập nhật ô tiêu đề với định dạng tháng năm
  outputSheet.getRange('A6:I6').setValue(formattedDate);
  
  // Tối ưu phần 1: Chỉ đọc dữ liệu cần thiết từ sheet OUTPUT
  var lastRow = outputSheet.getLastRow();
  var numRows = lastRow - 10; // Số dòng cần xử lý, bắt đầu từ dòng 11
  
  if (numRows <= 0) {
    ui.alert('Không tìm thấy dữ liệu trong sheet mẫu.');
    return false;
  }
  
  // Chỉ lấy cột A, B, E, G để tạo map và xử lý
  var outputDataRange = outputSheet.getRange(11, 1, numRows, 9);
  var outputData = outputDataRange.getValues();
  
  // Tạo map từ dữ liệu OUTPUT (chỉ sử dụng cột A và B làm key)
  var outputMap = {};
  for (var i = 0; i < outputData.length; i++) {
    var key = outputData[i][0] + '|' + outputData[i][1];
    outputMap[key] = i;
  }
  
  // Mảng tích lũy dữ liệu cho cột E và G để cập nhật một lần duy nhất
  var colEValues = Array(numRows).fill(0);
  var colGValues = Array(numRows).fill(0);
  
  // Khởi tạo giá trị ban đầu từ dữ liệu OUTPUT
  for (var i = 0; i < outputData.length; i++) {
    var eValue = outputData[i][4];
    var gValue = outputData[i][6];
    
    colEValues[i] = (eValue === "" || eValue === null || isNaN(eValue)) ? 0 : Number(eValue);
    colGValues[i] = (gValue === "" || gValue === null || isNaN(gValue)) ? 0 : Number(gValue);
  }
  
  // Tối ưu phần 2: Xử lý các worksheet được chọn
  var missingSheets = [];
  var selectedSheets = data.selectedSheets;
  if (!selectedSheets || selectedSheets.length === 0) {
    ui.alert('Vui lòng chọn ít nhất một worksheet.');
    return false;
  }
  
  // Xử lý từng worksheet
  for (var sheetIndex = 0; sheetIndex < selectedSheets.length; sheetIndex++) {
    var sheetName = selectedSheets[sheetIndex];
    var inputSheet = ss.getSheetByName(sheetName);
    
    if (!inputSheet) {
      missingSheets.push(sheetName);
      continue; // Bỏ qua worksheet không tìm thấy
    }
    
    // Tối ưu phần 3: Chỉ đọc phạm vi cần thiết
    var inputLastRow = inputSheet.getLastRow();
    if (inputLastRow <= 12) continue; // Bỏ qua nếu không có dữ liệu
    
    // Chỉ lấy các cột cần thiết: A, B, E, G (index 0, 1, 4, 6)
    var inputRange = inputSheet.getRange(13, 1, inputLastRow - 12, 9);
    var inputData = inputRange.getValues();
    
    // Xử lý và tích lũy dữ liệu
    for (var i = 0; i < inputData.length; i++) {
      var key = inputData[i][0] + '|' + inputData[i][1];
      if (outputMap.hasOwnProperty(key)) {
        var rowIndex = outputMap[key];
        
        // Xử lý cột E (index 4)
        var eValue = inputData[i][4];
        if (eValue !== "" && eValue !== null && !isNaN(eValue)) {
          colEValues[rowIndex] += Number(eValue);
        }
        
        // Xử lý cột G (index 6)
        var gValue = inputData[i][6];
        if (gValue !== "" && gValue !== null && !isNaN(gValue)) {
          colGValues[rowIndex] += Number(gValue);
        }
      }
    }
  }
  
  // Tối ưu phần 4: Cập nhật dữ liệu hàng loạt một lần duy nhất
  // Chuyển mảng thành định dạng cần thiết cho setValues
  var formattedColE = colEValues.map(function(value) { return [value]; });
  var formattedColG = colGValues.map(function(value) { return [value]; });
  
  // Cập nhật cột E và G trong sheet OUTPUT
  outputSheet.getRange(11, 5, numRows, 1).setValues(formattedColE);
  outputSheet.getRange(11, 7, numRows, 1).setValues(formattedColG);
  
  // Hiển thị cảnh báo về các sheet không tìm thấy
  if (missingSheets.length > 0) {
    ui.alert("Không tìm thấy worksheet có tên: " + missingSheets.join(", "));
  }
  
  // Cập nhật dữ liệu sản phẩm
  try {
    copyProductDataToOutput(monthYear);
  } catch (e) {
    console.error("Lỗi khi xử lý dữ liệu sản phẩm:", e);
  }
  
  return true;
}