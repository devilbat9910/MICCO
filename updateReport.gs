/**
 * Hàm chính để tạo báo cáo từ worksheet nguồn và cập nhật dữ liệu.
 * @param {Sheet} inputSheet - Sheet nguồn (INPUT) chứa dữ liệu, tên dạng 'PXXX_Báo cáo mm/yyyy'.
 * @param {string} baseSheetPrefix - Tiền tố worksheet cơ sở ('PX').
 */
function updateReport(inputSheet, baseSheetPrefix) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheetName = inputSheet.getName();

  // Lấy mã phân xưởng và tháng/năm từ tên worksheet INPUT
  const match = inputSheetName.match(/(PX[A-Z]{2})_Báo cáo (\d{2})\/(\d{4})/);
  if (!match) {
    throw new Error('Tên worksheet không hợp lệ. Định dạng phải là "PX[A-Z]{2}_Báo cáo mm/yyyy".');
  }
  const unitCode = match[1]; // Mã phân xưởng (VD: PXĐT)
  const month = parseInt(match[2], 10);
  const year = parseInt(match[3], 10);

  // Tạo tên worksheet cơ sở từ tiền tố và mã phân xưởng
  const baseSheetName = `${unitCode}_BCTH`;

  // Nhân bản worksheet cơ sở nếu chưa tồn tại báo cáo năm
  const baseSheet = spreadsheet.getSheetByName(baseSheetName);
  if (!baseSheet) {
    throw new Error(`Worksheet cơ sở "${baseSheetName}" không tồn tại.`);
  }
  const outputSheetName = `${baseSheetName}_${year}`;
  let outputSheet = spreadsheet.getSheetByName(outputSheetName);

  if (!outputSheet) {
    outputSheet = baseSheet.copyTo(spreadsheet);
    outputSheet.setName(outputSheetName);
    outputSheet.showSheet();
  }

  // Cập nhật tiêu đề trong worksheet OUTPUT
  updateHeader(outputSheet, year, month);

  // Cập nhật dữ liệu từ worksheet INPUT sang OUTPUT
  processData(inputSheet, outputSheet, month);

  SpreadsheetApp.getUi().alert(`Báo cáo tháng ${month}/${year} cho phân xưởng ${unitCode} đã được cập nhật thành công.`);
}
