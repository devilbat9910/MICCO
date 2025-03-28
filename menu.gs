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
      month: month,
      year: year,
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
 * Xử lý danh sách worksheet đã chọn từ giao diện HTML.
 * @param {Array} selectedNames - Danh sách tên worksheet đã chọn.
 */
function processSelectedWorksheets(selectedNames) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const groups = {};

  selectedNames.forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    const match = name.match(/^PX([A-ZĐ]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    if (!match) return;
    const unitCode = match[1];
    const month = parseInt(match[2], 10);
    const year = match[3];
    const key = `${unitCode}_${year}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push({ sheet: sheet, month: month });
  });

  for (const key in groups) {
    const [unitCode, year] = key.split('_');
    const baseSheetName = `PX${unitCode}_BCTH`;
    const baseSheet = spreadsheet.getSheetByName(baseSheetName);
    if (!baseSheet) {
      Logger.log(`Base sheet ${baseSheetName} not found.`);
      continue;
    }
    const outputSheetName = `${baseSheetName}_${year}`;
    let outputSheet = spreadsheet.getSheetByName(outputSheetName);
    if (!outputSheet) {
      outputSheet = baseSheet.copyTo(spreadsheet);
      outputSheet.setName(outputSheetName);
    }

    updateHeader(outputSheet, parseInt(year, 10));
    groups[key].forEach(item => {
      processData(item.sheet, outputSheet, item.month);
    });
  }
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