
/**
 * Hiển thị hộp thoại để người dùng chọn các thông số báo cáo
 */
function showProductReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ProductReportDialog')
    .setWidth(600)
    .setHeight(550)
    .setTitle('Tạo báo cáo theo sản phẩm');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo báo cáo theo sản phẩm');
}

/**
 * Lấy danh sách các phân xưởng có sẵn
 * @return {Array} Mảng các đối tượng phân xưởng {code, name}
 */
function getWorkshops() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Danh mục sản phẩm');
  
  try {
    // Giả sử danh sách phân xưởng bắt đầu từ ô F1
    const data = sheet.getRange('F1:G20').getValues();
    const workshops = [];
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        workshops.push({
          code: data[i][0].toString().trim(),
          name: data[i][1].toString().trim()
        });
      }
    }
    
    return workshops;
  } catch (error) {
    Logger.log('Lỗi khi lấy danh sách phân xưởng: ' + error.message);
    return [];
  }
}

/**
 * Lấy danh sách tất cả sản phẩm từ sheet "Sản lượng ngày"
 * @return {Array} Mảng các đối tượng sản phẩm
 */
function getAllProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sản lượng ngày');
  
  try {
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(1, 1, lastRow, 3).getValues(); // Lấy 3 cột: INDEX, Chủng loại, Tổng
    
    const products = [];
    let currentCategory = null;
    
    for (let i = 1; i < data.length; i++) { // Bỏ qua hàng tiêu đề
      const indexVal = data[i][0];
      const productName = data[i][1];
      const totalVal = data[i][2];
      
      if (!indexVal || !productName) continue;
      
      const index = indexVal.toString().trim();
      
      // Kiểm tra nếu đây là một loại sản phẩm (level 1)
      if (/^[A-HJ-Z]$/.test(index)) {
        currentCategory = {
          index: index,
          name: productName,
          isCategory: true,
          hasProduction: totalVal > 0,
          children: []
        };
        products.push(currentCategory);
      } 
      // Kiểm tra nếu đây là sản phẩm cụ thể hoặc chỉ tiêu
      else if (index.includes('.') && currentCategory) {
        const levels = index.split('.');
        
        if (levels.length >= 2) {
          // Kiểm tra nếu đây là sản phẩm cụ thể (level 2, thành phần cố định '.1')
          const isProduct = levels[1] === '1';
          
          // Tạo đối tượng sản phẩm hoặc chỉ tiêu
          const item = {
            index: index,
            name: productName,
            isCategory: false, 
            isProduct: isProduct,
            hasProduction: totalVal > 0,
            parentIndex: currentCategory.index
          };
          
          currentCategory.children.push(item);
        }
      }
    }
    
    return products;
  } catch (error) {
    Logger.log('Lỗi khi lấy danh sách sản phẩm: ' + error.message);
    return [];
  }
}

/**
 * Lấy dữ liệu sản lượng từ sheet của phân xưởng
 * @param {string} workshopCode - Mã phân xưởng
 * @param {string} monthYear - Tháng/năm (định dạng "MM/YYYY")
 * @return {Object} Đối tượng chứa dữ liệu sản lượng theo index sản phẩm
 */
function getProductionData(workshopCode, monthYear) {
  try {
    // Tên sheet theo định dạng "MM/YYYY"
    // Giả sử đã truy cập file của phân xưởng liên quan
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(monthYear);
    
    if (!sheet) {
      throw new Error(`Không tìm thấy sheet "${monthYear}" cho phân xưởng ${workshopCode}`);
    }
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(1, 1, lastRow, 3).getValues(); // INDEX, Chủng loại, Tổng
    
    const productionData = {};
    
    for (let i = 1; i < data.length; i++) {
      const index = data[i][0];
      const total = data[i][2];
      
      if (index && total > 0) {
        productionData[index.toString().trim()] = total;
      }
    }
    
    return productionData;
  } catch (error) {
    Logger.log(`Lỗi khi lấy dữ liệu sản lượng: ${error.message}`);
    return {};
  }
}

/**
 * Tạo báo cáo mới dựa trên sản phẩm đã chọn
 * @param {Object} data - Dữ liệu từ form
 * @return {Object} Kết quả tạo báo cáo
 */
function createProductReport(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Parse dữ liệu đầu vào
    const workshopCode = data.workshopCode;
    const workshopName = data.workshopName;
    const monthYear = data.monthYear;
    const selectedProducts = data.selectedProducts || [];
    const sendToWorkshop = data.sendToWorkshop || false;
    
    if (!workshopCode || !monthYear || selectedProducts.length === 0) {
      throw new Error('Vui lòng chọn đầy đủ thông tin phân xưởng, tháng/năm và sản phẩm');
    }
    
    // Tạo tên cho báo cáo mới
    const reportName = `PX${workshopCode}_Báo cáo ${monthYear}`;
    
    // Kiểm tra xem báo cáo đã tồn tại chưa
    let reportSheet = ss.getSheetByName(reportName);
    
    if (reportSheet) {
      const response = ui.alert(
        'Báo cáo đã tồn tại',
        `Báo cáo "${reportName}" đã tồn tại. Bạn có muốn tạo lại không?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        return { success: false, message: 'Hủy tạo báo cáo' };
      }
      
      // Xóa báo cáo cũ nếu đồng ý tạo lại
      ss.deleteSheet(reportSheet);
    }
    
    // Tạo báo cáo mới từ BC_TCT
    const templateSheet = ss.getSheetByName('BC_TCT');
    if (!templateSheet) {
      throw new Error('Không tìm thấy sheet mẫu "BC_TCT"');
    }
    
    // Sao chép sheet mẫu
    reportSheet = templateSheet.copyTo(ss);
    reportSheet.setName(reportName);
    
    // Lấy dữ liệu sản lượng từ sheet của phân xưởng
    const productionData = getProductionData(workshopCode, monthYear);
    
    // Cập nhật tiêu đề báo cáo (giả sử tiêu đề ở ô A5)
    const parts = monthYear.split('/');
    if (parts.length === 2) {
      reportSheet.getRange('A5').setValue(`THÁNG ${parts[0]} NĂM ${parts[1]}`);
      reportSheet.getRange('A7').setValue(`Phân xưởng: ${workshopName}`);
    }
    
    // Tìm những hàng cần giữ lại dựa trên sản phẩm đã chọn
    const rowsToKeep = findRowsToKeep(reportSheet, selectedProducts, productionData);
    
    // Ẩn những hàng không cần thiết
    hideUnusedRows(reportSheet, rowsToKeep);
    
    // Truyền dữ liệu sản lượng vào báo cáo
    updateProductionData(reportSheet, rowsToKeep, productionData);
    
    // Gửi báo cáo đến phân xưởng nếu được yêu cầu
    if (sendToWorkshop) {
      const sendResult = sendReportToWorkshop(workshopCode, monthYear, reportName);
      
      if (sendResult.success) {
        return { 
          success: true, 
          message: `Đã tạo báo cáo "${reportName}" và gửi đến phân xưởng ${workshopName} thành công`,
          sheetName: reportName 
        };
      } else {
        return { 
          success: true, 
          message: `Đã tạo báo cáo "${reportName}" thành công nhưng không gửi được đến phân xưởng: ${sendResult.message}`,
          sheetName: reportName 
        };
      }
    }
    
    return { 
      success: true, 
      message: `Đã tạo báo cáo "${reportName}" thành công`,
      sheetName: reportName 
    };
    
  } catch (error) {
    Logger.log('Lỗi khi tạo báo cáo: ' + error.message);
    return { success: false, message: 'Lỗi: ' + error.message };
  }
}

/**
 * Tìm các hàng cần giữ lại trong báo cáo dựa trên sản phẩm đã chọn
 * @param {Sheet} sheet - Sheet báo cáo
 * @param {Array} selectedProducts - Mảng index của sản phẩm đã chọn
 * @param {Object} productionData - Dữ liệu sản lượng
 * @return {Object} Danh sách các hàng cần giữ lại
 */
function findRowsToKeep(sheet, selectedProducts, productionData) {
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 2).getValues(); // INDEX, Chủng loại
  
  const rowsToKeep = {};
  const parentMap = {}; // Map để theo dõi quan hệ cha-con
  
  // Tạo map của tất cả các hàng theo index
  for (let i = 10; i < data.length; i++) { // Giả sử dữ liệu bắt đầu từ hàng 11 (index 10)
    const rowIndex = i + 1; // 1-based index
    const cellValue = data[i][0];
    
    if (cellValue) {
      const index = cellValue.toString().trim();
      
      // Lưu trữ index và hàng tương ứng
      rowsToKeep[index] = {
        row: rowIndex,
        keep: false,
        hasChildren: false
      };
      
      // Xây dựng quan hệ cha-con
      if (index.includes('.')) {
        const parts = index.split('.');
        parts.pop(); // Loại bỏ phần tử cuối
        
        if (parts.length > 0) {
          const parentIndex = parts.join('.');
          
          if (!parentMap[parentIndex]) {
            parentMap[parentIndex] = [];
          }
          
          parentMap[parentIndex].push(index);
        }
      }
    }
  }
  
  // Đánh dấu các sản phẩm đã chọn và cha của chúng
  selectedProducts.forEach(productIndex => {
    // Đánh dấu sản phẩm
    if (rowsToKeep[productIndex]) {
      rowsToKeep[productIndex].keep = true;
      
      // Đánh dấu tất cả các mục cha
      let currentIndex = productIndex;
      
      while (currentIndex.includes('.')) {
        const parts = currentIndex.split('.');
        parts.pop();
        const parentIndex = parts.join('.');
        
        if (rowsToKeep[parentIndex]) {
          rowsToKeep[parentIndex].keep = true;
          rowsToKeep[parentIndex].hasChildren = true;
        }
        
        currentIndex = parentIndex;
      }
      
      // Đánh dấu mục cha cấp cao nhất (chữ cái)
      if (currentIndex.length === 1 && rowsToKeep[currentIndex]) {
        rowsToKeep[currentIndex].keep = true;
      }
    }
  });
  
  // Thêm các hàng tiêu đề
  for (let i = 1; i <= 10; i++) {
    rowsToKeep['header_' + i] = {
      row: i,
      keep: true
    };
  }
  
  return rowsToKeep;
}

/**
 * Ẩn các hàng không cần thiết trong báo cáo
 * @param {Sheet} sheet - Sheet báo cáo
 * @param {Object} rowsToKeep - Danh sách các hàng cần giữ lại
 */
function hideUnusedRows(sheet, rowsToKeep) {
  const lastRow = sheet.getLastRow();
  const rowsToHide = [];
  
  // Tìm các hàng cần ẩn
  for (let i = 11; i <= lastRow; i++) {
    let shouldKeep = false;
    
    // Kiểm tra xem hàng có trong danh sách cần giữ lại không
    for (const index in rowsToKeep) {
      if (rowsToKeep[index].row === i && rowsToKeep[index].keep) {
        shouldKeep = true;
        break;
      }
    }
    
    if (!shouldKeep) {
      rowsToHide.push(i);
    }
  }
  
  // Ẩn các hàng theo nhóm để tối ưu hiệu suất
  if (rowsToHide.length > 0) {
    // Sắp xếp các hàng cần ẩn
    rowsToHide.sort((a, b) => a - b);
    
    // Nhóm các hàng liên tiếp
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
    sheet.hideRows(startRow, count);
  }
}

/**
 * Cập nhật dữ liệu sản lượng vào báo cáo
 * @param {Sheet} sheet - Sheet báo cáo
 * @param {Object} rowsToKeep - Danh sách các hàng cần giữ lại
 * @param {Object} productionData - Dữ liệu sản lượng
 */
function updateProductionData(sheet, rowsToKeep, productionData) {
  // Cập nhật dữ liệu sản lượng vào cột G
  for (const index in rowsToKeep) {
    if (productionData[index] !== undefined && productionData[index] > 0) {
      const row = rowsToKeep[index].row;
      sheet.getRange(row, 7).setValue(productionData[index]);
    }
  }
}

/**
 * Gửi báo cáo đến phân xưởng
 * @param {string} workshopCode - Mã phân xưởng
 * @param {string} monthYear - Tháng/năm (định dạng "MM/YYYY")
 * @param {string} reportSheetName - Tên sheet báo cáo nguồn
 * @return {Object} Kết quả gửi báo cáo
 */
function sendReportToWorkshop(workshopCode, monthYear, reportSheetName) {
  try {
    // Lấy URL của phân xưởng
    const workshopUrl = getWorkshopUrl(workshopCode);
    
    if (!workshopUrl) {
      throw new Error(`Không tìm thấy URL cho phân xưởng ${workshopCode}`);
    }
    
    // Mở bảng tính của phân xưởng
    const workshopSS = SpreadsheetApp.openByUrl(workshopUrl);
    
    // Lấy sheet báo cáo từ bảng tính hiện tại
    const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSS.getSheetByName(reportSheetName);
    
    if (!sourceSheet) {
      throw new Error(`Không tìm thấy sheet báo cáo "${reportSheetName}"`);
    }
    
    // Tạo hoặc cập nhật sheet báo cáo trong bảng tính của phân xưởng
    let targetSheet = workshopSS.getSheetByName(monthYear);
    
    if (!targetSheet) {
      // Tạo sheet mới
      targetSheet = workshopSS.insertSheet(monthYear);
    }
    
    // Sao chép dữ liệu và định dạng từ sheet nguồn
    copySheetData(sourceSheet, targetSheet);
    
    return {
      success: true,
      message: `Đã gửi báo cáo thành công đến phân xưởng ${workshopCode}`
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
  // Xóa dữ liệu hiện có trong sheet đích
  targetSheet.clear();
  
  // Lấy số hàng và cột của sheet nguồn
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    // Lấy dữ liệu từ sheet nguồn
    const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastCol);
    const targetRange = targetSheet.getRange(1, 1, lastRow, lastCol);
    
    // Sao chép giá trị và định dạng
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    
    // Điều chỉnh độ rộng của cột
    for (let i = 1; i <= lastCol; i++) {
      targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
    }
    
    // Điều chỉnh độ cao của hàng
    for (let i = 1; i <= lastRow; i++) {
      targetSheet.setRowHeight(i, sourceSheet.getRowHeight(i));
    }
    
    // Sao chép hàng ẩn
    for (let i = 1; i <= lastRow; i++) {
      if (sourceSheet.isRowHiddenByUser(i)) {
        targetSheet.hideRows(i);
      }
    }
    
    // Sao chép cột ẩn
    for (let i = 1; i <= lastCol; i++) {
      if (sourceSheet.isColumnHiddenByUser(i)) {
        targetSheet.hideColumns(i);
      }
    }
  }
}