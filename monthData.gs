/**
 * Hàm được gọi từ giao diện HTML khi người dùng nhấn nút "Tạo báo cáo".
 * @param {Object} data - Đối tượng chứa:
 *    - monthYear: chuỗi tháng/năm theo định dạng "MM/YYYY"
 *    - selectedProducts: mảng danh mục sản phẩm đã chọn (có thể dùng cho xử lý mở rộng)
 */
function generateAndCopySheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthYear = data.monthYear; // Đã có định dạng "MM/YYYY"
  var outputSheetName = 'BC_TCT_' + monthYear;
  
  var templateSheet = ss.getSheetByName('BC_TCT');
  if (!templateSheet) {
    Logger.log('Không tìm thấy sheet mẫu BC_TCT');
    SpreadsheetApp.getUi().alert('Không tìm thấy sheet mẫu BC_TCT.');
    return;
  }
  

  // Nếu chưa có sheet OUTPUT, tạo bằng cách nhân bản sheet mẫu
  var outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = templateSheet.copyTo(ss).setName(outputSheetName);
    outputSheet.showSheet(); // Hiển thị sheet ngay cả khi sheet gốc bị ẩn
  }

  
  // Tìm các sheet INPUT có tên theo định dạng: PX<PHAN_XUONG>_Báo cáo <MM/YYYY>
  var regex = new RegExp('^PX(ĐT|QN|CP|TB|NB|ĐN|VT)_Báo cáo ' + monthYear + '$');
  var sheets = ss.getSheets().filter(function(sheet) {
    return regex.test(sheet.getName());
  });
  
  if (sheets.length === 0) {
    Logger.log('Không tìm thấy worksheet nào phù hợp với tháng/năm ' + monthYear);
    SpreadsheetApp.getUi().alert('Không tìm thấy worksheet nào phù hợp với tháng/năm ' + monthYear);
    return;
  }
  
  // Lấy dữ liệu hiện có của sheet OUTPUT và tạo map (key = "cột A|cột B") từ dòng 11 (index 10)
  var outputData = outputSheet.getDataRange().getValues();
  var outputMap = {};
  for (var i = 10; i < outputData.length; i++) {
    var key = outputData[i][0] + '|' + outputData[i][1];
    outputMap[key] = i;
  }
  
  // Duyệt qua từng sheet INPUT (với dữ liệu bắt đầu từ hàng 13, index 12)
  sheets.forEach(function(sheet) {
    var inputData = sheet.getDataRange().getValues();
    for (var i = 12; i < inputData.length; i++) {
      var key = inputData[i][0] + '|' + inputData[i][1];
      if (outputMap.hasOwnProperty(key)) {
        var rowIndex = outputMap[key];
        // Cộng dồn giá trị ở cột E (index 4) và cột G (index 6)
        [4, 6].forEach(function(col) {
          var inputValue = inputData[i][col];
          if (typeof inputValue === 'number') {
            if (typeof outputData[rowIndex][col] !== 'number') {
              Logger.log('Giá trị trong OUTPUT không phải số tại sheet ' + outputSheetName + ' dòng ' + (rowIndex + 1) + ', cột ' + (col + 1));
            } else {
              outputData[rowIndex][col] += inputValue;
            }
          } else {
            Logger.log('Dữ liệu không hợp lệ tại ' + sheet.getName() + ' - Hàng ' + (i + 1) + ', Cột ' + (col + 1) + ': ' + inputValue);
          }
        });
      } else {
        Logger.log('Bỏ qua dòng không khớp: ' + sheet.getName() + ' - Hàng ' + (i + 1));
      }
    }
  });
  
  // Ghi lại dữ liệu đã cập nhật vào sheet OUTPUT
  outputSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  SpreadsheetApp.getUi().alert('Báo cáo đã được tạo thành công!');
}