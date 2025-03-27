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
 * Dữ liệu nhận vào: { monthYear: "YYYY-MM", selectedUnits: [...] }.
 * Chuyển đổi định dạng "YYYY-MM" thành "MM/YYYY" và thực hiện tổng hợp báo cáo.
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
  
  // --- Phần xử lý tổng hợp dữ liệu từ các sheet INPUT cho các đơn vị ---
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
  
  // TỐI ƯU HÓA PHẦN 1: Chỉ đọc dữ liệu cần thiết từ sheet OUTPUT
  // Lấy số dòng thực tế có dữ liệu trong sheet OUTPUT (bắt đầu từ dòng 11)
  var lastRow = outputSheet.getLastRow();
  var numRows = lastRow - 10; // Số dòng cần xử lý, bắt đầu từ dòng 11
  
  if (numRows <= 0) {
    ui.alert('Không tìm thấy dữ liệu trong sheet mẫu.');
    return false;
  }
  
  // Chỉ lấy cột A, B, E, G để tạo map và xử lý (thay vì lấy toàn bộ dữ liệu)
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
  
  // TỐI ƯU HÓA PHẦN 2: Xử lý các đơn vị được chọn
  var missingSheets = [];
  var selectedUnits = data.selectedUnits;
  if (!selectedUnits || selectedUnits.length === 0) {
    ui.alert('Vui lòng chọn ít nhất một đơn vị.');
    return false;
  }
  
  // Xử lý từng đơn vị - với tối ưu hóa
  for (var unitIndex = 0; unitIndex < selectedUnits.length; unitIndex++) {
    var unit = selectedUnits[unitIndex];
    var inputSheetName = "PX" + unit + "_Báo cáo " + monthYear;
    var inputSheet = ss.getSheetByName(inputSheetName);
    
    if (!inputSheet) {
      missingSheets.push(inputSheetName);
      continue; // Bỏ qua đơn vị không tìm thấy
    }
    
    // TỐI ƯU HÓA PHẦN 3: Chỉ đọc phạm vi cần thiết của INPUT sheet
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
  
  // TỐI ƯU HÓA PHẦN 4: Cập nhật dữ liệu hàng loạt một lần duy nhất
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
  
  // TỐI ƯU HÓA PHẦN 5: Tách rời việc gọi copyProductDataToOutput để giảm thời gian chờ đợi
  // Thay vì gọi trực tiếp, sử dụng đặt lịch để thực hiện sau khi đã trả về kết quả
  try {
    var lock = LockService.getScriptLock();
    // Thử khóa trong 5 giây để tránh xung đột nếu có nhiều người dùng
    if (lock.tryLock(5000)) {
      // Lưu trữ thông tin để xử lý tiếp
      PropertiesService.getScriptProperties().setProperty(
        'PENDING_PRODUCT_DATA', 
        JSON.stringify({monthYear: monthYear, timestamp: new Date().getTime()})
      );
      lock.releaseLock();
    }
    
    // Gọi hàm xử lý sản phẩm - nhưng chạy bất đồng bộ
    processProductDataAsync();
  } catch (e) {
    // Xử lý lỗi nhưng không làm gián đoạn kết quả chính
    console.error("Lỗi khi chuẩn bị xử lý dữ liệu sản phẩm:", e);
  }
  
  // Hoàn thành xử lý chính và trả về thành công
  return true;
}

/**
 * Hàm xử lý dữ liệu sản phẩm bất đồng bộ.
 * Tách rời khỏi luồng chính để giảm thời gian chờ đợi.
 */
function processProductDataAsync() {
  try {
    var props = PropertiesService.getScriptProperties();
    var pendingDataJSON = props.getProperty('PENDING_PRODUCT_DATA');
    
    if (!pendingDataJSON) return;
    
    var pendingData = JSON.parse(pendingDataJSON);
    var monthYear = pendingData.monthYear;
    
    // Chỉ xử lý dữ liệu trong vòng 30 phút
    var currentTime = new Date().getTime();
    if (currentTime - pendingData.timestamp > 30 * 60 * 1000) {
      // Dữ liệu quá cũ, xóa và bỏ qua
      props.deleteProperty('PENDING_PRODUCT_DATA');
      return;
    }
    
    // Gọi hàm xử lý sản phẩm thực tế
    copyProductDataToOutput(monthYear);
    
    // Xóa dữ liệu chờ xử lý sau khi hoàn thành
    props.deleteProperty('PENDING_PRODUCT_DATA');
  } catch (e) {
    console.error("Lỗi khi xử lý dữ liệu sản phẩm bất đồng bộ:", e);
  }
}

/**
 * Hàm trigger để chạy theo lịch.
 * Thêm tính năng này và đặt trigger để chạy mỗi phút hoặc 5 phút.
 */
function checkAndProcessPendingTasks() {
  processProductDataAsync();
}