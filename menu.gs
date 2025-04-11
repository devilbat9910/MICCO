/**
 * Thêm menu tùy chỉnh trong Google Sheets để gọi các hàm xử lý.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tổng hợp báo cáo')
    .addItem('Cập nhật báo cáo theo PX', 'menuGenerateReport')
    .addItem('Tổng hợp báo cáo theo tháng', 'menuMonthlySummary')
    .addItem('Thu gọn báo cáo', 'shrinkReportUI')
    //.addItem('Tô màu theo quy tắc 4color', 'highlightCells')
    .addItem('Cập nhật worksheet cơ sở', 'menuUpdateBaseSheets')
    .addItem('Tạo báo cáo từ gốc', 'showProductReportDialog')
    .addToUi();
}

/**
 * Hàm menu để gọi chức năng tổng hợp báo cáo theo tháng.
 */
function menuMonthlySummary() {
  try {
    summarizeMonthlyReports();
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Lỗi: ${error.message}`);
  }
}

/**
 * Hàm menu để mở giao diện HTML cho chức năng cập nhật báo cáo theo PX.
 * Phiên bản tối ưu hiệu suất.
 */
function menuGenerateReport() {
  const html = HtmlService.createHtmlOutputFromFile('html_update_report')
    .setWidth(400)
    .setHeight(400)
    .setTitle('Cập nhật báo cáo theo PX');
  SpreadsheetApp.getUi().showModalDialog(html, 'Cập nhật báo cáo theo PX');
}

/**
 * Hàm menu để mở giao diện HTML cho chức năng cập nhật worksheet cơ sở.
 */
function menuUpdateBaseSheets() {
  const html = HtmlService.createHtmlOutputFromFile('html_update_base')
    .setWidth(300)
    .setHeight(200)
    .setTitle('Cập nhật worksheet cơ sở');
  SpreadsheetApp.getUi().showModalDialog(html, 'Cập nhật worksheet cơ sở');
}

/**
 * Lấy danh sách worksheet hợp lệ (không ẩn và có tên đúng định dạng).
 * @returns {Array} - Danh sách các worksheet hợp lệ.
 */
function getValidWorksheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const validSheets = sheets.filter(sheet => {
    if (sheet.isSheetHidden()) return false;
    const name = sheet.getName();
    // Regex hỗ trợ ký tự 'Đ' và các chữ cái A-Z
    const match = name.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    return match !== null;
  });

  return validSheets.map(sheet => {
    const name = sheet.getName();
    const match = name.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    const unitCode = match[1];
    const month = match[2];
    const year = match[3];
    const unitName = getUnitName(unitCode);
    return {
      name: name,
      unitCode: unitCode,
      unitName: unitName,
      month: parseInt(month, 10), // Chuyển thành số để sắp xếp
      year: parseInt(year, 10),   // Chuyển thành số để sắp xếp
      displayName: `${unitName} ${month}/${year}`
    };
  });
}

/**
 * Ánh xạ mã đơn vị thành tên đầy đủ.
 * @param {string} code - Mã đơn vị.
 * @returns {string} - Tên đầy đủ của đơn vị.
 */
function getUnitName(code) {
  const units = {
    'ĐT': 'Đông Triều',
    'CP': 'Cẩm Phả',
    'QN': 'Quảng Ninh',
    'TB': 'Tây Bắc',
    'NB': 'Ninh Bình',
    'ĐN': 'Đà Nẵng',
    'VT': 'Vũng Tàu'
  };
  return units[code] || code;
}

/**
 * Xử lý danh sách worksheet đã chọn từ giao diện HTML với cải tiến hiệu suất và hiển thị tiến trình.
 * @param {Array} selectedNames - Danh sách tên worksheet đã chọn.
 */
function processSelectedWorksheets(selectedNames) {
  // Hiển thị hộp thoại tiến trình
  const html = HtmlService.createHtmlOutputFromFile('html_progress_dialog')
    .setWidth(400)
    .setHeight(300)
    .setTitle('Tiến trình cập nhật báo cáo');
  
  const ui = SpreadsheetApp.getUi();
  const dialog = ui.showModelessDialog(html, 'Tiến trình cập nhật báo cáo');
  
  // Khởi tạo tiến trình và thông tin
  PropertiesService.getScriptProperties().setProperty('progressConfig', JSON.stringify({
    total: selectedNames.length
  }));
  
  // Thực hiện xử lý trong background
  try {
    // Nhóm các worksheet theo đơn vị và năm để tối ưu xử lý
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const groups = {};
    
    updateProgressDialog(0, selectedNames.length, 'Đang phân tích worksheet...', '');
    
    // Tối ưu 1: Nhóm các sheet một lần duy nhất
    selectedNames.forEach((name, index) => {
      const sheet = spreadsheet.getSheetByName(name);
      if (!sheet) return;
      
      const match = name.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
      if (!match) return;
      
      const unitCode = match[1];
      const month = parseInt(match[2], 10);
      const year = match[3];
      const key = `${unitCode}_${year}`;
      
      if (!groups[key]) {
        groups[key] = {
          unitCode: unitCode,
          year: year,
          items: []
        };
      }
      
      groups[key].items.push({ sheet: sheet, month: month });
      
      // Cập nhật tiến trình phân tích
      updateProgressDialog(index + 1, selectedNames.length, 
                          'Đang phân tích worksheet...', 
                          `Đã phân tích ${index + 1}/${selectedNames.length}: ${name}`);
    });
    
    // Tối ưu 2: Xử lý từng nhóm
    const groupKeys = Object.keys(groups);
    
    for (let groupIndex = 0; groupIndex < groupKeys.length; groupIndex++) {
      const key = groupKeys[groupIndex];
      const group = groups[key];
      
      // Cập nhật tiến trình
      updateProgressDialog(groupIndex, groupKeys.length,
                          `Đang xử lý nhóm ${groupIndex + 1}/${groupKeys.length}`,
                          `Phân xưởng: ${group.unitCode}, Năm: ${group.year}`);
      
      // Tìm hoặc tạo sheet đầu ra
      const baseSheetName = `PX${group.unitCode}_BCTH`;
      const baseSheet = spreadsheet.getSheetByName(baseSheetName);
      
      if (!baseSheet) {
        updateProgressDialog(groupIndex, groupKeys.length,
                            `Bỏ qua phân xưởng ${group.unitCode}`,
                            `Không tìm thấy sheet cơ sở ${baseSheetName}`);
        continue;
      }
      
      const outputSheetName = `${baseSheetName}_${group.year}`;
      let outputSheet = spreadsheet.getSheetByName(outputSheetName);
      
      if (!outputSheet) {
        updateProgressDialog(groupIndex, groupKeys.length,
                            `Đang tạo sheet đầu ra mới`,
                            `Sheet: ${outputSheetName}`);
                            
        outputSheet = baseSheet.copyTo(spreadsheet);
        outputSheet.setName(outputSheetName);
      }
      
      // Cập nhật tiêu đề
      updateHeader(outputSheet, parseInt(group.year, 10));
      
      // Tối ưu 3: Xử lý từng sheet trong nhóm
      const items = group.items;
      
      for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
        const item = items[itemIndex];
        
        // Tính toán tiến trình tổng thể
        const overallProgress = (groupIndex / groupKeys.length) + 
                               (itemIndex / items.length / groupKeys.length);
        
        updateProgressDialog(Math.floor(overallProgress * selectedNames.length), 
                            selectedNames.length,
                            `Đang xử lý worksheet ${itemIndex + 1}/${items.length}`,
                            `Sheet: ${item.sheet.getName()}, Tháng: ${item.month}`);
        
        // Gọi hàm xử lý dữ liệu đã tối ưu
        try {
          processData(item.sheet, outputSheet, item.month);
        } catch (error) {
          // Ghi nhật ký lỗi nhưng vẫn tiếp tục xử lý các sheet khác
          Logger.log(`Lỗi khi xử lý ${item.sheet.getName()}: ${error.message}`);
          
          updateProgressDialog(Math.floor(overallProgress * selectedNames.length),
                              selectedNames.length,
                              `Lỗi khi xử lý ${item.sheet.getName()}`,
                              error.message);
        }
      }
    }
    
    // Hoàn thành xử lý
    updateProgressDialog(selectedNames.length, selectedNames.length,
                        'Đã hoàn thành cập nhật báo cáo',
                        `Tổng số worksheet đã xử lý: ${selectedNames.length}`);
    
    return true;
  } catch (error) {
    // Xử lý lỗi
    updateProgressDialog(0, 1, 'Lỗi xử lý', error.message, true);
    Logger.log(`Lỗi: ${error.message}`);
    return false;
  }
}

/**
 * Lấy cấu hình tiến trình.
 * @returns {Object} - Cấu hình tiến trình.
 */
function getProgressConfig() {
  const configStr = PropertiesService.getScriptProperties().getProperty('progressConfig');
  return configStr ? JSON.parse(configStr) : { total: 0 };
}

/**
 * Cập nhật hộp thoại tiến trình.
 * @param {number} current - Số lượng đã xử lý.
 * @param {number} total - Tổng số lượng cần xử lý.
 * @param {string} status - Trạng thái hiện tại.
 * @param {string} details - Chi tiết bổ sung.
 * @param {boolean} isError - Có phải là lỗi không.
 */
function updateProgressDialog(current, total, status, details, isError = false) {
  try {
    // Cập nhật thuộc tính
    PropertiesService.getScriptProperties().setProperty('progressStatus', JSON.stringify({
      current: current,
      total: total,
      status: status,
      details: details,
      isError: isError,
      timestamp: new Date().getTime()
    }));
  } catch (e) {
    Logger.log(`Lỗi khi cập nhật tiến trình: ${e.message}`);
  }
}

/**
 * Lấy trạng thái tiến trình hiện tại.
 * @returns {Object} - Trạng thái tiến trình.
 */
function getProgressStatus() {
  const statusStr = PropertiesService.getScriptProperties().getProperty('progressStatus');
  return statusStr ? JSON.parse(statusStr) : null;
}

/**
 * Lấy danh sách các phân xưởng có worksheet báo cáo.
 * @returns {Array} - Danh sách các phân xưởng với mã và tên.
 */
function getAvailableUnits() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const unitCodes = new Set();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    // Regex hỗ trợ ký tự 'Đ' và các chữ cái A-Z
    const match = name.match(/^PX([A-ZĐ]{2})_Báo cáo \d{2}\/\d{4}$/);
    if (match) unitCodes.add(match[1]);
  });

  return Array.from(unitCodes).map(code => ({
    code: code,
    name: getUnitName(code)
  }));
}