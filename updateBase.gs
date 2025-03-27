/**
 * Cập nhật worksheet cơ sở dựa trên phân xưởng được chọn.
 * @param {string} selectedUnit Mã phân xưởng hoặc 'all'.
 */
function updateBaseSheets(selectedUnit) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const latestReports = getLatestReportSheets();
  const unitsToUpdate = selectedUnit === 'all' ? Object.keys(latestReports) : [selectedUnit];

  unitsToUpdate.forEach(unitCode => {
    const baseSheet = spreadsheet.getSheetByName(`PX${unitCode}_BCTH`);
    const reportSheet = latestReports[unitCode];

    if (!baseSheet || !reportSheet) {
      ui.alert(`Không tìm thấy sheet cho phân xưởng ${unitCode}.`);
      return;
    }

    const response = ui.alert(
      'Xác nhận',
      `Cập nhật ${baseSheet.getName()} từ ${reportSheet.getName()}?`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;

    try {
      reportSheet.showRows(1, reportSheet.getMaxRows());
      const reportRows = Math.max(reportSheet.getLastRow() - 12, 0);
      const baseRows = Math.max(baseSheet.getLastRow() - 12, 0);

      // Sao chép cột A-D
      if (reportRows > 0) {
        const source = reportSheet.getRange(13, 1, reportRows, 4);
        const target = baseSheet.getRange(13, 1, reportRows, 4);
        source.copyTo(target, { contentsOnly: true });
      }

      // Điều chỉnh số hàng
      if (reportRows > baseRows) {
        baseSheet.insertRowsAfter(12 + baseRows, reportRows - baseRows);
      } else if (reportRows < baseRows) {
        baseSheet.deleteRows(13 + reportRows, baseRows - reportRows);
      }

      // Vẽ đường viền từ D đến AQ
      const lastRow = baseSheet.getLastRow();
      if (lastRow >= 13) {
        baseSheet.getRange(13, 4, lastRow - 12, 40).setBorder(true, true, true, true, true, true);
      }

      shrinkReport(reportSheet);
      ui.alert(`Đã cập nhật thành công cho ${unitCode}.`);
    } catch (error) {
      ui.alert(`Lỗi khi cập nhật ${unitCode}: ${error.message}`);
    }
  });
}

/**
 * Lấy worksheet báo cáo mới nhất cho từng phân xưởng.
 * @returns {Object} Danh sách sheet mới nhất theo mã phân xưởng.
 */
function getLatestReportSheets() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(sheet => /^PX([A-Z]{2})_Báo cáo (\d{2})\/(\d{4})$/.test(sheet.getName()));
  
  const latest = {};
  sheets.forEach(sheet => {
    const [, unitCode, month, year] = sheet.getName().match(/^PX([A-Z]{2})_Báo cáo (\d{2})\/(\d{4})$/);
    const dateValue = parseInt(year) * 100 + parseInt(month);
    if (!latest[unitCode] || dateValue > latest[unitCode].dateValue) {
      latest[unitCode] = { sheet, dateValue };
    }
  });

  const result = {};
  for (const unitCode in latest) result[unitCode] = latest[unitCode].sheet;
  return result;
}

/**
 * Thu gọn worksheet báo cáo bằng cách ẩn các dòng dư.
 * @param {Sheet} sheet Worksheet cần thu gọn.
 */
function shrinkReport(sheet) {
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (maxRows > lastRow) sheet.hideRows(lastRow + 1, maxRows - lastRow);
}