/**
 * Hiển thị hộp thoại để tạo báo cáo cho nhiều phân xưởng
 */
function showBatchReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BatchReportDialog')
    .setWidth(600)
    .setHeight(550)
    .setTitle('Tạo báo cáo cho Phân xưởng');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo báo cáo cho Phân xưởng');
}

/**
 * Lấy danh sách phân xưởng và các sản phẩm liên quan
 * @return {Array} Danh sách phân xưởng với sản phẩm
 */
function getWorkshopsWithProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Danh mục sản phẩm');
  
  try {
    // Lấy dữ liệu từ sheet danh mục sản phẩm
    const data = sheet.getRange('F1:H20').getValues();
    const workshops = [];
    
    // Bỏ qua dòng tiêu đề
    let startRow = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && typeof data[i][0] === 'string' && data[i][0].trim() === 'Mã PX') {
        startRow = i + 1;
        break;
      }
    }
    
    // Xử lý từng dòng phân xưởng
    for (let i = startRow; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const workshopCode = data[i][0].toString().trim();
        const workshopName = data[i][1].toString().trim();
        
        // Kiểm tra xem có phải là dòng tiêu đề hay không
        if (workshopCode === 'Mã PX' || workshopName === 'Tên PX') {
          continue;
        }
        
        // Lấy danh sách sản phẩm từ cột H
        let productIndexes = [];
        if (data[i][2]) {
          // Chuyển đổi thành chữ hoa và loại bỏ khoảng trắng
          productIndexes = data[i][2].toString()
            .split(',')
            .map(index => index.trim().toUpperCase())
            .filter(index => index); // Lọc ra các phần tử rỗng
        }
        
        workshops.push({
          code: workshopCode,
          name: workshopName,
          productIndexes: productIndexes
        });
      }
    }
    
    return workshops;
  } catch (error) {
    Logger.log('Lỗi khi lấy danh sách phân xưởng và sản phẩm: ' + error.message);
    return [];
  }
}

/**
 * Xác định phạm vi hàng cho mỗi loại sản phẩm trong sheet "Sản lượng ngày"
 * @return {Object} Phạm vi hàng cho mỗi index sản phẩm
 */
function getProductRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sản lượng ngày');
  
  try {
    const lastRow = sheet.getLastRow();
    const dataA = sheet.getRange('A1:A' + lastRow).getValues();
    
    // Khởi tạo đối tượng lưu trữ phạm vi
    const ranges = {};
    const indexPositions = [];
    
    // Tìm vị trí của tất cả các index chữ cái
    for (let i = 0; i < dataA.length; i++) {
      const value = dataA[i][0];
      if (value && typeof value === 'string' && /^[A-HJ-Z]$/.test(value.trim())) {
        indexPositions.push({
          index: value.trim().toUpperCase(),
          row: i + 1 // 1-based index
        });
      }
    }
    
    // Tính toán phạm vi cho mỗi index
    for (let i = 0; i < indexPositions.length; i++) {
      const current = indexPositions[i];
      let endRow;
      
      if (i < indexPositions.length - 1) {
        // Nếu không phải index cuối cùng, endRow là hàng trước index tiếp theo
        endRow = indexPositions[i + 1].row - 1;
      } else {
        // Nếu là index cuối cùng, tìm hàng rỗng cuối cùng hoặc hàng cuối của sheet
        let emptyRow = lastRow;
        for (let j = current.row; j <= lastRow; j++) {
          if (!dataA[j - 1][0]) {
            emptyRow = j - 1;
            break;
          }
        }
        endRow = emptyRow;
      }
      
      ranges[current.index] = {
        startRow: current.row,
        endRow: endRow
      };
    }
    
    return ranges;
  } catch (error) {
    Logger.log('Lỗi khi xác định phạm vi sản phẩm: ' + error.message);
    return {};
  }
}

/**
 * Tạo báo cáo cho các phân xưởng đã chọn
 * @param {Object} data - Dữ liệu từ form
 * @return {Object} Kết quả tạo báo cáo
 */
function createBatchReports(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Parse dữ liệu đầu vào
    const monthYear = data.monthYear;
    const selectedWorkshops = data.selectedWorkshops || [];
    const deleteAfterSend = data.deleteAfterSend || false;
    
    if (!monthYear || selectedWorkshops.length === 0) {
      throw new Error('Vui lòng chọn tháng/năm và ít nhất một phân xưởng');
    }
    
    const parts = monthYear.split('/');
    if (parts.length !== 2) {
      throw new Error('Định dạng tháng/năm không hợp lệ');
    }
    
    // Lấy danh sách phân xưởng và sản phẩm
    const workshopsWithProducts = getWorkshopsWithProducts();
    
    // Lấy phạm vi sản phẩm
    const productRanges = getProductRanges();
    
    // Danh sách kết quả
    const results = [];
    
    // Tạo báo cáo cho từng phân xưởng đã chọn
    for (const selectedCode of selectedWorkshops) {
      const workshop = workshopsWithProducts.find(w => w.code === selectedCode);
      if (!workshop) {
        results.push({
          code: selectedCode,
          success: false,
          message: 'Không tìm thấy thông tin phân xưởng'
        });
        continue;
      }
      
      // Tạo báo cáo sản lượng từ template
      const result = createProductionReport(workshop, monthYear, productRanges, deleteAfterSend);
      results.push({
        code: workshop.code,
        name: workshop.name,
        success: result.success,
        message: result.message
      });
    }
    
    return {
      success: true,
      results: results
    };
    
  } catch (error) {
    Logger.log('Lỗi khi tạo báo cáo hàng loạt: ' + error.message);
    return {
      success: false,
      message: 'Lỗi: ' + error.message
    };
  }
}

/**
 * Tạo báo cáo sản lượng cho một phân xưởng
 * @param {Object} workshop - Thông tin phân xưởng
 * @param {string} monthYear - Tháng/năm (MM/YYYY)
 * @param {Object} productRanges - Phạm vi hàng cho mỗi sản phẩm
 * @param {boolean} deleteAfterSend - Có xóa sheet sau khi gửi không
 * @return {Object} Kết quả tạo báo cáo
 */
function createProductionReport(workshop, monthYear, productRanges, deleteAfterSend) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Lấy template "Sản lượng ngày"
    const templateSheet = ss.getSheetByName('Sản lượng ngày');
    if (!templateSheet) {
      throw new Error('Không tìm thấy sheet mẫu "Sản lượng ngày"');
    }
    
    // Tạo tên sheet duy nhất cho sheet tạm
    const tempSheetName = `temp_${workshop.code.toLowerCase()}_${new Date().getTime()}`;
    
    // Sao chép template sang sheet tạm
    const reportSheet = templateSheet.copyTo(ss);
    reportSheet.setName(tempSheetName);
    
    // Lọc dữ liệu dựa trên index sản phẩm
    filterProductsByIndexes(reportSheet, workshop.productIndexes, productRanges);
    
    // Gửi báo cáo đến phân xưởng
    const sendResult = sendReportToWorkshop(workshop.code, monthYear, reportSheet);
    
    // Xóa sheet tạm nếu được yêu cầu
    if (deleteAfterSend) {
      ss.deleteSheet(reportSheet);
    }
    
    if (!sendResult.success) {
      return {
        success: false,
        message: sendResult.message
      };
    }
    
    return {
      success: true,
      message: `Đã tạo và gửi báo cáo cho phân xưởng ${workshop.name} thành công`
    };
    
  } catch (error) {
    Logger.log(`Lỗi khi tạo báo cáo cho ${workshop.name}: ${error.message}`);
    return {
      success: false,
      message: `Lỗi: ${error.message}`
    };
  }
}

/**
 * Lọc dữ liệu trong sheet theo danh sách index sản phẩm
 * @param {Sheet} sheet - Sheet cần lọc
 * @param {Array} indexes - Danh sách index sản phẩm cần giữ lại
 * @param {Object} productRanges - Phạm vi hàng cho mỗi index
 */
function filterProductsByIndexes(sheet, indexes, productRanges) {
  const lastRow = sheet.getLastRow();
  
  // Danh sách hàng cần ẩn
  const rowsToHide = [];
  
  // Xác định hàng cần ẩn dựa trên index
  for (const index in productRanges) {
    // Nếu index không nằm trong danh sách cần giữ, ẩn toàn bộ phạm vi
    if (!indexes.includes(index)) {
      const range = productRanges[index];
      for (let row = range.startRow; row <= range.endRow; row++) {
        rowsToHide.push(row);
      }
    }
  }
  
  // Ẩn các hàng
  if (rowsToHide.length > 0) {
    // Sắp xếp hàng
    rowsToHide.sort((a, b) => a - b);
    
    // Nhóm các hàng liên tiếp để tối ưu hiệu suất
    let startRow = rowsToHide[0];
    let count = 1;
    
    for (let i = 1; i < rowsToHide.length; i++) {
      if (rowsToHide[i] === rowsToHide[i-1] + 1) {
        count++;
      } else {
        // Ẩn nhóm hàng hiện tại
        sheet.hideRows(startRow, count);
        
        // Bắt đầu nhóm mới
        startRow = rowsToHide[i];
        count = 1;
      }
    }
    
    // Ẩn nhóm cuối cùng
    if (count > 0) {
      sheet.hideRows(startRow, count);
    }
  }
}

/**
 * Gửi báo cáo đến phân xưởng
 * @param {string} workshopCode - Mã phân xưởng
 * @param {string} monthYear - Tháng/năm (MM/YYYY)
 * @param {Sheet} sourceSheet - Sheet báo cáo nguồn
 * @return {Object} Kết quả gửi báo cáo
 */
function sendReportToWorkshop(workshopCode, monthYear, sourceSheet) {
  try {
    // Lấy URL của phân xưởng
    const workshopUrl = getWorkshopUrl(workshopCode);
    
    if (!workshopUrl) {
      throw new Error(`Không tìm thấy URL cho phân xưởng ${workshopCode}`);
    }
    
    // Mở bảng tính của phân xưởng
    const workshopSS = SpreadsheetApp.openByUrl(workshopUrl);
    
    // Luôn tạo sheet tạm trong spreadsheet đích trước
    let tempSheet;
    try {
      tempSheet = workshopSS.insertSheet("TempSheet_" + new Date().getTime());
    } catch (e) {
      Logger.log(`Không thể tạo sheet mới: ${e.message}. Thử phương pháp khác...`);
      throw new Error(`Không thể tạo sheet mới trong bảng tính đích: ${e.message}`);
    }
    
    // Sao chép dữ liệu và định dạng từ sheet nguồn sang sheet tạm
    copySheetData(sourceSheet, tempSheet);
    
    // Sau khi đã sao chép dữ liệu, xác định tên phù hợp cho sheet
    let targetSheetName = monthYear;
    let suffix = 1;
    
    // Kiểm tra xem sheet đã tồn tại chưa và tìm tên thích hợp
    const existingSheets = workshopSS.getSheets().map(s => s.getName());
    
    while (existingSheets.includes(targetSheetName)) {
      suffix++;
      targetSheetName = `${monthYear} (${suffix})`;
    }
    
    // Đổi tên sheet tạm thành tên đích
    tempSheet.setName(targetSheetName);
    
    return {
      success: true,
      message: `Đã gửi báo cáo thành công đến phân xưởng ${workshopCode} với tên "${targetSheetName}"`
    };
  } catch (error) {
    Logger.log(`Lỗi khi gửi báo cáo đến phân xưởng: ${error.message}`);
    return {
      success: false,
      message: `Lỗi: ${error.message}`
    };
  }
}

/**
 * Lấy URL của phân xưởng dựa vào mã
 * @param {string} workshopCode - Mã phân xưởng
 * @return {string|null} URL của phân xưởng hoặc null nếu không tìm thấy
 */
function getWorkshopUrl(workshopCode) {
  const workshopUrls = {
    'CP': 'https://docs.google.com/spreadsheets/d/1fS7bRnPy2xJChqoLVr1AgEmMyoeOJlaNC0Plt_JS7N8/edit?usp=sharing',
    'ĐN': 'https://docs.google.com/spreadsheets/d/1OxLqZDL6sWXa3vg0inM_0d8CbGvAQCrEWSKpQTx-U84/edit?usp=sharing',
    'TB': 'https://docs.google.com/spreadsheets/d/1Nnn3_ElEiYGs2eanwH5O8fv7YJIMiwpzUreU_pVCNP8/edit?usp=sharing',
    'QN': 'https://docs.google.com/spreadsheets/d/1R9lMIQjzL_eDkMCCEdUenImUBOwxdE_LA3d78h-QriQ/edit?usp=sharing',
    'NB': 'https://docs.google.com/spreadsheets/d/1QT7fJvY7573VB-UJNCq3uJxVd-UMWTg57tDBn4U7FqU/edit?usp=sharing',
    'VT': 'https://docs.google.com/spreadsheets/d/1ojKesIV8nDd495U28GBEUoDTsUDfUqSqv-MW-Xza8vU/edit?usp=sharing',
    'ĐT': 'https://docs.google.com/spreadsheets/d/1RRO_RK2dZJcEsGtYxM4OUYPcv5BXP_vr0Od_idap8PA/edit?usp=sharing'
  };
  
  return workshopUrls[workshopCode] || null;
}

/**
 * Sao chép dữ liệu và định dạng từ sheet nguồn sang sheet đích
 * @param {Sheet} sourceSheet - Sheet nguồn
 * @param {Sheet} targetSheet - Sheet đích
 */
function copySheetData(sourceSheet, targetSheet) {
  try {
    // Lấy số hàng và cột của sheet nguồn
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    
    Logger.log(`Copying data: lastRow=${lastRow}, lastCol=${lastCol}`);
    
    if (lastRow <= 0 || lastCol <= 0) {
      Logger.log(`Warning: Source sheet dimensions are invalid: rows=${lastRow}, cols=${lastCol}`);
      return;
    }
    
    // Lấy dữ liệu từ sheet nguồn
    const formulas = sourceSheet.getRange(1, 1, lastRow, lastCol).getFormulas();
    const values = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const formats = sourceSheet.getRange(1, 1, lastRow, lastCol).getNumberFormats();
    const backgrounds = sourceSheet.getRange(1, 1, lastRow, lastCol).getBackgrounds();
    const fontColors = sourceSheet.getRange(1, 1, lastRow, lastCol).getFontColors();
    const fontWeights = sourceSheet.getRange(1, 1, lastRow, lastCol).getFontWeights();
    const horizontalAlignments = sourceSheet.getRange(1, 1, lastRow, lastCol).getHorizontalAlignments();
    const verticalAlignments = sourceSheet.getRange(1, 1, lastRow, lastCol).getVerticalAlignments();
    
    // Thiết lập dữ liệu cho sheet đích
    // Đầu tiên đặt giá trị (cần đặt giá trị trước khi đặt công thức)
    targetSheet.getRange(1, 1, lastRow, lastCol).setValues(values);
    
    // Sau đó đặt công thức (ghi đè lên các ô có công thức)
    // Chỉ đặt công thức cho các ô thực sự có công thức (khác rỗng)
    for (let i = 0; i < formulas.length; i++) {
      for (let j = 0; j < formulas[i].length; j++) {
        if (formulas[i][j] !== '') {
          targetSheet.getRange(i + 1, j + 1).setFormula(formulas[i][j]);
        }
      }
    }
    
    // Đặt các thuộc tính định dạng
    targetSheet.getRange(1, 1, lastRow, lastCol).setNumberFormats(formats);
    targetSheet.getRange(1, 1, lastRow, lastCol).setBackgrounds(backgrounds);
    targetSheet.getRange(1, 1, lastRow, lastCol).setFontColors(fontColors);
    targetSheet.getRange(1, 1, lastRow, lastCol).setFontWeights(fontWeights);
    targetSheet.getRange(1, 1, lastRow, lastCol).setHorizontalAlignments(horizontalAlignments);
    targetSheet.getRange(1, 1, lastRow, lastCol).setVerticalAlignments(verticalAlignments);
    
    // Điều chỉnh độ rộng của cột
    for (let i = 1; i <= lastCol; i++) {
      targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
    }
    
    // Điều chỉnh độ cao của hàng
    for (let i = 1; i <= lastRow; i++) {
      targetSheet.setRowHeight(i, sourceSheet.getRowHeight(i));
    }
    
    // Sao chép các hàng ẩn
    for (let i = 1; i <= lastRow; i++) {
      try {
        if (sourceSheet.isRowHiddenByUser(i)) {
          targetSheet.hideRows(i);
        }
      } catch (e) {
        Logger.log(`Could not copy hidden state for row ${i}: ${e.message}`);
      }
    }
    
    // Sao chép các cột ẩn
    for (let i = 1; i <= lastCol; i++) {
      try {
        if (sourceSheet.isColumnHiddenByUser(i)) {
          targetSheet.hideColumns(i);
        }
      } catch (e) {
        Logger.log(`Could not copy hidden state for column ${i}: ${e.message}`);
      }
    }
    
    Logger.log(`Successfully copied data from ${sourceSheet.getName()} to ${targetSheet.getName()}`);
  } catch (error) {
    Logger.log(`Error copying sheet data: ${error.message}`);
    throw error;
  }
}