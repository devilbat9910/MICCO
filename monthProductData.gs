/**
 * Tổng hợp dữ liệu sản lượng của từng loại sản phẩm từ các sheet INPUT (PX…_Báo cáo <MM/YYYY>)
 * và copy giá trị số (không bao gồm công thức) sang sheet OUTPUT (BC_TCT_<MM/YYYY>) theo "index" của sản phẩm.
 *
 * @param {string} monthYear - Chuỗi tháng/năm theo định dạng "MM/YYYY".
 */
function copyProductDataToOutput(monthYear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheetName = "BC_TCT_" + monthYear;
  var outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    SpreadsheetApp.getUi().alert("Không tìm thấy sheet OUTPUT: " + outputSheetName);
    return;
  }
  
  // Lấy header của sheet OUTPUT ở hàng 10 (giả sử hàng 10 chứa mã đơn vị)
  var headerRow = outputSheet.getRange(10, 1, 1, outputSheet.getLastColumn()).getValues()[0];
  
  // Xây dựng regex để nhận dạng sheet INPUT cho tháng đang xét
  var inputSheetRegex = new RegExp("^PX(ĐT|CP|QN|TB|NB|ĐN|VT)_Báo cáo " + monthYear + "$");
  
  // Lấy tất cả các sheet phù hợp với định dạng tên INPUT
  var allSheets = ss.getSheets();
  var inputSheets = allSheets.filter(function(sheet) {
    return inputSheetRegex.test(sheet.getName());
  });
  
  if (inputSheets.length === 0) {
    SpreadsheetApp.getUi().alert("Không tìm thấy sheet INPUT nào cho tháng: " + monthYear);
    return;
  }
  
  // Tạo regex để tìm các index phù hợp:
  // - {Chữ cái}.1
  // - {Chữ cái}.1.{số}
  // - {Chữ cái}.1.{số}.{số}...
  var allValidIndexRegex = /^([A-Za-z]\.1)(\.\d+)*$/;
  
  // Lấy toàn bộ dữ liệu của sheet OUTPUT để tìm kiếm các hàng
  var outputRangeData = outputSheet.getRange(1, 1, outputSheet.getLastRow(), 1).getValues();
  
  // Duyệt qua từng sheet INPUT
  inputSheets.forEach(function(inputSheet) {
    var sheetName = inputSheet.getName();
    // Trích xuất mã đơn vị từ tên sheet: từ "PX{unit}_Báo cáo ..." => unit = {unit}
    var match = sheetName.match(/^PX(ĐT|CP|QN|TB|NB|ĐN|VT)_Báo cáo/);
    if (!match) {
      Logger.log("Sheet không khớp định dạng: " + sheetName);
      return;
    }
    var unitCode = match[1]; // ví dụ "NB"
    
    // Xác định cột OUTPUT tương ứng với đơn vị bằng cách quét hàng 10 của OUTPUT
    var outputCol = -1;
    for (var col = 0; col < headerRow.length; col++) {
      if (headerRow[col] && headerRow[col].toString().trim() === unitCode) {
        outputCol = col + 1; // chuyển từ 0-index sang 1-index
        break;
      }
    }
    if (outputCol === -1) {
      Logger.log("Không tìm thấy cột cho đơn vị " + unitCode + " ở sheet OUTPUT.");
      return;
    }
    
    // Lấy toàn bộ dữ liệu của sheet INPUT để xác định các hàng cần copy
    var inputData = inputSheet.getDataRange().getValues();
    var numRowsInput = inputData.length;
    
    // Duyệt qua từng hàng của sheet INPUT
    for (var i = 0; i < numRowsInput; i++) {
      var indexVal = inputData[i][0]; // giả sử cột A chứa chỉ số sản phẩm
      if (typeof indexVal !== "string") continue;
      
      // Chuẩn hóa index để tránh lỗi do khoảng trắng
      var normalizedIndex = indexVal.trim();
      
      // Kiểm tra xem index có phù hợp với định dạng cần thiết không
      if (!allValidIndexRegex.test(normalizedIndex)) continue;
      
      // Kiểm tra ô cột G (index 6) có giá trị không
      var cellValue = inputData[i][6];
      if (cellValue === "" || cellValue === null || cellValue === undefined) continue;
      
      // Xác định ô trong sheet OUTPUT có cùng index
      var targetRow = -1;
      for (var j = 0; j < outputRangeData.length; j++) {
        var outIndex = outputRangeData[j][0];
        if (typeof outIndex === "string" && outIndex.trim() === normalizedIndex) {
          targetRow = j + 1; // chuyển từ 0-index sang 1-index
          break;
        }
      }
      if (targetRow === -1) {
        Logger.log("Không tìm thấy dòng OUTPUT có chỉ số " + normalizedIndex + " để copy dữ liệu từ sheet " + sheetName);
        continue;
      }
      
      // Lấy giá trị số (không phải công thức) từ ô INPUT
      var valueToSet = inputSheet.getRange(i + 1, 7).getValue(); // Cột G là 7
      
      // Đặt giá trị (không phải công thức) vào ô OUTPUT
      outputSheet.getRange(targetRow, outputCol).setValue(valueToSet);
      
      // Lấy định dạng từ ô INPUT và áp dụng cho ô OUTPUT
      var sourceFormat = inputSheet.getRange(i + 1, 7).getNumberFormat();
      outputSheet.getRange(targetRow, outputCol).setNumberFormat(sourceFormat);
    }
  });
}