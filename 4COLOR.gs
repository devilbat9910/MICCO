/** 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Number color')
      .addItem('Thực hiện', 'highlightCells')
      .addToUi();
}
*/

function highlightCells() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var sheets = spreadsheet.getSheets(); // Lấy tất cả các sheet (dòng chính)
  var sheets = [spreadsheet.getActiveSheet()]; // Dự phòng: Chỉ xử lý sheet hiện tại

  for (var s = 0; s < sheets.length; s++) { // Duyệt qua từng sheet
    var sheet = sheets[s];
    var range = sheet.getDataRange(); // Toàn bộ dữ liệu trong sheet
    var formulas = range.getFormulas();
    var values = range.getValues();
    var fontColors = range.getFontColors(); // Lấy màu chữ của toàn bộ phạm vi

    // Xác định phạm vi hàng từ cột A
    var colAValues = sheet.getRange("A:A").getValues(); // Lấy toàn bộ cột A
    var startRow = -1;
    var endRow = -1;

    // Tìm hàng bắt đầu (ô có giá trị 'TT') và hàng kết thúc (ô cuối cùng có giá trị)
    for (var i = 0; i < colAValues.length; i++) {
      if (colAValues[i][0] === 'TT' && startRow === -1) {
        startRow = i + 1; // +1 vì chỉ số bắt đầu từ 1 trong sheet
      }
      if (colAValues[i][0] !== '' && startRow !== -1) {
        endRow = i + 1; // Cập nhật hàng cuối cùng có giá trị
      }
      if (colAValues[i][0] === '' && endRow !== -1) {
        break; // Thoát vòng lặp khi gặp ô trống sau ô cuối có giá trị
      }
    }

    // Giới hạn cột từ E (5) đến G (7)
    var startCol = 5; // Cột E
    var endCol = 7;   // Cột G

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        var cell = sheet.getRange(i + 1, j + 1);
        var value = values[i][j];
        var formula = formulas[i][j];
        var currentFontColor = fontColors[i][j]; // Lấy màu chữ hiện tại của ô

        if (formula) { // Nếu ô có công thức
          // Chỉ thay đổi màu chữ nếu màu hiện tại không phải là trắng
          if (currentFontColor !== "white" && currentFontColor !== "#ffffff") {
            if (formula.includes('!')) {
              cell.setFontColor("magenta");
            } else if (/^=\s*[A-Za-z]+\d+\s*$/.test(formula)) {
              cell.setFontColor("#34a853");
            } else {
              cell.setFontColor("black");
            }
          }
          cell.setBackground("#ffffff"); // Đặt nền trắng cho ô có công thức
        } 
        else if (typeof value === 'number') { // Nếu ô chứa số
          // Chỉ thay đổi màu chữ nếu màu hiện tại không phải là trắng
          if (currentFontColor !== "white" && currentFontColor !== "#ffffff") {
            cell.setFontColor("CornflowerBlue");
          }
          cell.setBackground("#ffffff"); // Nền trắng khi có số
        } 
        // Kiểm tra ô trống chỉ trong phạm vi từ hàng startRow đến endRow và cột E đến G
        else if (value === "" && 
                 i + 1 >= startRow && 
                 i + 1 <= endRow && 
                 j + 1 >= startCol && 
                 j + 1 <= endCol) {
          cell.setBackground("CornflowerBlue"); // Nền xanh dương nhạt cho ô trống trong phạm vi
          cell.setFontColor(null); // Xóa màu chữ (nếu có)
        }
      }
    }
  }
}