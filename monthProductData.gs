/**
 * Tổng hợp dữ liệu sản lượng của từng loại sản phẩm từ các sheet INPUT (PX…_Báo cáo <MM/YYYY>)
 * và copy giá trị số (không bao gồm công thức) sang sheet OUTPUT (BC_TCT_<MM/YYYY>) theo "index" của sản phẩm.
 *
 * Quy trình:
 * 1. Nhận tham số monthYear với định dạng "MM/YYYY". Từ đó xác định sheet OUTPUT có tên "BC_TCT_<MM/YYYY>".
 * 2. Lấy danh sách các sheet trong spreadsheet có tên theo định dạng: 
 *      "PX{unit}_Báo cáo <MM/YYYY>" (unit là một trong các mã: ĐT, CP, QN, TB, NB, ĐN, VT).
 * 3. Với mỗi sheet INPUT:
 *    a. Trích xuất mã đơn vị (unit) từ tên sheet.
 *    b. Ở sheet OUTPUT, quét hàng 10 để xác định cột chứa mã đơn vị đó.
 *    c. Quét dữ liệu của sheet INPUT: 
 *       - Tìm các hàng có chỉ số sản phẩm theo các định dạng sau:
 *          + {Chữ cái}.1 (ví dụ: A.1, B.1, J.1, ...)
 *          + {Chữ cái}.1.{số} (ví dụ: A.1.2, B.1.1, ...)
 *          + tất cả các mục con, cháu của những index trên (ví dụ: A.1.2.1, J.1.4.2.2, v.v.)
 *       - Nếu ở cột G (ở hàng đó) có giá trị (sản lượng) thì:
 *            i. Lấy giá trị số từ ô tại cột G của sheet INPUT.
 *           ii. Tìm ở sheet OUTPUT hàng có "index" (giá trị trong cột A) khớp với chỉ số sản phẩm đó.
 *          iii. Copy giá trị số từ ô INPUT sang ô OUTPUT tại vị trí (hàng tìm được, cột đã xác định cho đơn vị).
 *
 * Lưu ý: 
 * - Việc "trùng index" giữa INPUT và OUTPUT có nghĩa là sản phẩm được xác định bởi giá trị trong cột A của cả 2 sheet.
 * - Các cột OUTPUT được xác định bằng cách quét hàng 10 của sheet OUTPUT, nơi đã được điền sẵn mã đơn vị.
 * - Chỉ copy giá trị số, không copy công thức.
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
  
  // Xây dựng regex để nhận dạng sheet INPUT cho tháng đang xét:
  // Ví dụ: "PXNB_Báo cáo 02/2025" khi monthYear là "02/2025"
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
    
    // Lấy toàn bộ dữ liệu của sheet OUTPUT để tìm kiếm các hàng
    var outputRangeData = outputSheet.getRange(1, 1, outputSheet.getLastRow(), 1).getValues();
    
    // Tạo regex để tìm các index cấp 1: {Chữ cái}.1
    var level1IndexRegex = /^([A-Za-z]\.1)$/;
    
    // Tạo regex để tìm các index cấp 2: {Chữ cái}.1.{số}
    var level2IndexRegex = /^([A-Za-z]\.1\.\d+)$/;
    
    // Tạo regex mở rộng để tìm tất cả các index phù hợp:
    // - {Chữ cái}.1
    // - {Chữ cái}.1.{số}
    // - {Chữ cái}.1.{số}.{số}...
    var allValidIndexRegex = /^([A-Za-z]\.1)(\.\d+)*$/;
    
    // Duyệt qua từng hàng của sheet INPUT
    for (var i = 0; i < numRowsInput; i++) {
      var indexVal = inputData[i][0]; // giả sử cột A chứa chỉ số sản phẩm
      if (typeof indexVal !== "string") continue;
      
      // Chuẩn hóa index để tránh lỗi do khoảng trắng
      var normalizedIndex = indexVal.trim();
      
      // Kiểm tra xem index có phù hợp với các định dạng cần thiết không
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
    } // end for input rows
  }); // end for each input sheet
}