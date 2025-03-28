/**
 * Cập nhật worksheet cơ sở dựa trên phân xưởng được chọn với hiệu suất cao hơn.
 * @param {string} selectedUnit Mã phân xưởng hoặc 'all'.
 */
function updateBaseSheets(selectedUnit) {
  const startTime = new Date().getTime();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const latestReports = getLatestReportSheets();
  const unitsToUpdate = selectedUnit === 'all' ? Object.keys(latestReports) : [selectedUnit];
  
  // Xử lý hàng loạt thay vì từng đơn vị
  const batchUpdates = [];
  
  unitsToUpdate.forEach(unitCode => {
    const baseSheet = spreadsheet.getSheetByName(`PX${unitCode}_BCTH`);
    const reportSheet = latestReports[unitCode];

    if (!baseSheet || !reportSheet) {
      ui.alert(`Không tìm thấy sheet cho phân xưởng ${unitCode}.`);
      return;
    }

    batchUpdates.push({unitCode, baseSheet, reportSheet});
  });
  
  if (batchUpdates.length === 0) {
    ui.alert('Không có sheet nào cần cập nhật.');
    return;
  }
  
  // Hiển thị hộp thoại xác nhận một lần cho tất cả các sheet
  const unitsList = batchUpdates.map(update => update.unitCode).join(", ");
  const response = ui.alert(
    'Xác nhận',
    `Cập nhật các sheet cơ sở cho các phân xưởng: ${unitsList}?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Tạo hộp thoại tự đóng bằng HTML
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 20px;
          }
          .message {
            font-size: 16px;
            margin-bottom: 20px;
          }
          .status {
            color: #444;
            font-style: italic;
          }
        </style>
        <script>
          // Thiết lập thời gian tự đóng hộp thoại sau 1 phút
          setTimeout(function() {
            google.script.host.close();
          }, 60000);
        </script>
      </head>
      <body>
        <div class="message">Đang cập nhật worksheet cơ sở...</div>
        <div class="status">Vui lòng đợi trong khi hệ thống xử lý. Cửa sổ này sẽ tự đóng khi hoàn thành.</div>
      </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(150);
  
  const dialog = ui.showModelessDialog(htmlOutput, 'Đang cập nhật');
  
  try {
    // Thực hiện các cập nhật
    for (let i = 0; i < batchUpdates.length; i++) {
      const {unitCode, baseSheet, reportSheet} = batchUpdates[i];
      updateSingleBaseSheet(unitCode, baseSheet, reportSheet);
    }
    
    // Đóng dialog và hiện thông báo hoàn thành
    const endTime = new Date().getTime();
    const executionTime = ((endTime - startTime) / 1000).toFixed(2);
    
    // Tạo thông báo hoàn thành sẽ tự đóng hộp thoại trước đó
    const completionHtml = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <script>
            // Đóng hộp thoại đang xử lý
            try {
              google.script.host.close();
              window.top.close();
            } catch(e) {
              console.log(e);
            }
            
            // Thông báo hoàn thành cho người dùng
            window.onload = function() {
              google.script.run
                .withSuccessHandler(function() {
                  google.script.host.close();
                })
                .showCompletionAlert("Đã cập nhật thành công ${batchUpdates.length} phân xưởng sau ${executionTime} giây.");
            };
          </script>
        </head>
        <body>
          <div style="text-align:center; padding:10px;">
            Hoàn thành cập nhật!
          </div>
        </body>
      </html>
    `)
    .setWidth(10)
    .setHeight(10);
    
    ui.showModelessDialog(completionHtml, 'Hoàn thành');
    
  } catch (error) {
    ui.alert(`Lỗi khi cập nhật: ${error.message}`);
  }
}

/**
 * Hiển thị thông báo hoàn thành.
 * @param {string} message Thông báo cần hiển thị
 */
function showCompletionAlert(message) {
  SpreadsheetApp.getUi().alert(message);
  return true;
}

/**
 * Cập nhật một sheet cơ sở riêng biệt với tối ưu hóa hiệu suất.
 * @param {string} unitCode Mã phân xưởng 
 * @param {Sheet} baseSheet Sheet cơ sở
 * @param {Sheet} reportSheet Sheet báo cáo
 */
function updateSingleBaseSheet(unitCode, baseSheet, reportSheet) {
  try {
    reportSheet.showRows(1, reportSheet.getMaxRows());
    const reportRows = Math.max(reportSheet.getLastRow() - 12, 0);
    const baseRows = Math.max(baseSheet.getLastRow() - 12, 0);
    
    // Đảm bảo đủ số cột trong baseSheet
    const requiredColumns = 45; // Số cột tối thiểu cần có (A đến AS)
    ensureSufficientColumns(baseSheet, requiredColumns);

    // Điều chỉnh số hàng 
    if (reportRows > baseRows) {
      baseSheet.insertRowsAfter(12 + baseRows, reportRows - baseRows);
    } else if (reportRows < baseRows) {
      baseSheet.deleteRows(13 + reportRows, baseRows - reportRows);
    }

    // Sao chép dữ liệu - sử dụng batch operation
    if (reportRows > 0) {
      // Tối ưu: Sao chép theo nhóm thay vì từng cột
      const sourceRange = reportSheet.getRange(13, 1, reportRows, 4);
      const targetRange = baseSheet.getRange(13, 1, reportRows, 4);
      
      // Sao chép nội dung và định dạng
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
      
      // Sao chép cột E (định mức) nếu có
      if (reportSheet.getLastColumn() >= 5) {
        reportSheet.getRange(13, 5, reportRows, 1).copyTo(
          baseSheet.getRange(13, 5, reportRows, 1), 
          SpreadsheetApp.CopyPasteType.PASTE_VALUES
        );
      }
      
      // Sao chép định dạng nền và chữ theo batch
      const sourceRangeFull = reportSheet.getRange(13, 1, reportRows, 5);
      baseSheet.getRange(13, 1, reportRows, 5).setBackgrounds(sourceRangeFull.getBackgrounds());
      baseSheet.getRange(13, 1, reportRows, 5).setFontColors(sourceRangeFull.getFontColors());
      baseSheet.getRange(13, 1, reportRows, 5).setFontWeights(sourceRangeFull.getFontWeights());
    }

    // Vẽ đường viền - một lần cho cả vùng
    const lastRow = baseSheet.getLastRow();
    if (lastRow >= 13) {
      baseSheet.getRange(13, 1, lastRow - 12, 40).setBorder(true, true, true, true, true, true);
    }
    
    // Xử lý công thức và tiêu đề 
    const reportSheetName = reportSheet.getName();
    const match = reportSheetName.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    if (match) {
      const month = parseInt(match[2], 10);
      const year = parseInt(match[3], 10);
      
      // Batch áp dụng công thức cho cột F và G
      applyFormulaBatch(baseSheet, match[2], match[3]);
      
      // Cập nhật tiêu đề
      updateHeader(baseSheet, year, month);
    }

    // Ẩn các hàng dư thừa
    hideExtraRows(reportSheet);
    return true;
  } catch (error) {
    Logger.log(`Lỗi khi cập nhật ${unitCode}: ${error.message}`);
    return false;
  }
}

/**
 * Áp dụng công thức hàng loạt cho cột F và G.
 * @param {Sheet} sheet Sheet cần áp dụng công thức
 * @param {string} month Tháng
 * @param {string} year Năm
 */
function applyFormulaBatch(sheet, month, year) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 13) return;
  
  const dataRange = sheet.getRange(13, 1, lastRow - 12, 2);
  const values = dataRange.getValues();
  const fontWeights = dataRange.getFontWeights();
  
  // Mảng công thức cho cột F và G
  const formulasF = [];
  const formulasG = [];
  
  // Danh sách các nội dung cần loại trừ khỏi quy tắc in đậm
  const excludedBoldContents = ["Sản lượng", "Điện - động lực"];
  
  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 13;
    const cellBValue = values[i][1];
    const isBold = fontWeights[i][1] === "bold";
    
    // Quyết định có áp dụng công thức hay không
    const needsFormula = !isBold || excludedBoldContents.includes(cellBValue);
    
    if (needsFormula) {
      formulasF.push([`=G${rowIndex}*1000/E${rowIndex}`]);
      
      // Công thức tính tổng cho cột G
      const columnLetters = ['I', 'L', 'O', 'R', 'U', 'X', 'AA', 'AD', 'AG', 'AJ', 'AM', 'AP'];
      formulasG.push([`=SUM(${columnLetters.map(letter => `${letter}${rowIndex}`).join(';')})`]);
    } else {
      formulasF.push([""]);
      formulasG.push([""]);
    }
  }
  
  // Áp dụng công thức theo batch
  if (formulasF.length > 0) {
    sheet.getRange(13, 6, formulasF.length, 1).setFormulas(formulasF);
    sheet.getRange(13, 7, formulasG.length, 1).setFormulas(formulasG);
  }
}

/**
 * Đảm bảo sheet có đủ số cột cần thiết.
 * @param {Sheet} sheet Sheet cần kiểm tra
 * @param {number} requiredColumns Số cột cần thiết
 */
function ensureSufficientColumns(sheet, requiredColumns) {
  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < requiredColumns) {
    const columnsToAdd = requiredColumns - currentColumns;
    sheet.insertColumnsAfter(currentColumns, columnsToAdd);
  }
}

/**
 * Lấy worksheet báo cáo mới nhất cho từng phân xưởng.
 * @returns {Object} Danh sách sheet mới nhất theo mã phân xưởng.
 */
function getLatestReportSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sheet => /^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/.test(sheet.getName()));
  
  const latest = {};
  sheets.forEach(sheet => {
    const match = sheet.getName().match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    if (match) {
      const unitCode = match[1];
      const month = parseInt(match[2], 10);
      const year = parseInt(match[3], 10);
      const dateValue = year * 100 + month;
      
      if (!latest[unitCode] || dateValue > latest[unitCode].dateValue) {
        latest[unitCode] = { sheet, dateValue };
      }
    }
  });

  const result = {};
  for (const unitCode in latest) result[unitCode] = latest[unitCode].sheet;
  return result;
}

/**
 * Ẩn các hàng dư thừa trong worksheet.
 * @param {Sheet} sheet Worksheet cần xử lý.
 */
function hideExtraRows(sheet) {
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (maxRows > lastRow) sheet.hideRows(lastRow + 1, maxRows - lastRow);
}