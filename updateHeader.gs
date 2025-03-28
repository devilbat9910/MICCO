/**
 * Cập nhật tiêu đề trong worksheet OUTPUT.
 * @param {Sheet} sheet - Worksheet OUTPUT cần cập nhật.
 * @param {number} year - Năm lấy từ tên worksheet INPUT.
 * @param {number} [month] - Tháng lấy từ tên worksheet INPUT (tùy chọn).
 */
function updateHeader(sheet, year, month = null) {
  const headerRow = 11;
  const startColumn = 8;
  const totalColumns = 36;
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < startColumn + totalColumns - 1) {
    throw new Error(`Sheet không đủ số cột để thiết lập tiêu đề. Yêu cầu tối thiểu: ${startColumn + totalColumns - 1}, hiện tại: ${lastColumn}`);
  }

  sheet.getRange(headerRow, startColumn, 1, totalColumns).clearContent();

  for (let i = 0; i < 12; i++) {
    const currentMonth = i + 1;
    const columnOffset = i * 3;
    const range = sheet.getRange(headerRow, startColumn + columnOffset, 1, 3);
    range.merge();
    range.setValue(`${String(currentMonth).padStart(2, '0')}/${year}`);
    range.setHorizontalAlignment("center");
    range.setFontWeight("bold");
  }

  if (month) {
    const monthOffset = (month - 1) * 3;
    const monthRange = sheet.getRange(headerRow, startColumn + monthOffset, 1, 3);
    monthRange.setBackground("#FFFF00");
  }

  updateOutputHeaders(sheet, year);
}

/**
 * Cập nhật các ô tiêu đề khác trong worksheet OUTPUT.
 * @param {Sheet} sheet - Worksheet OUTPUT.
 * @param {number} year - Năm.
 */
function updateOutputHeaders(sheet, year) {
  // Cập nhật tiêu đề trong hàng 5 và 7
  [5, 7].forEach(row => {
    const range = sheet.getRange(row, 1, 1, 9);
    const values = range.getValues()[0];
    const updatedValues = values.map(cell =>
      typeof cell === 'string' ? cell.replace(/yyyy/g, year.toString()) : cell
    );
    range.setValues([updatedValues]);
  });

  // Cập nhật tiêu đề trong hàng 11, cột D và E
  const rangeD11E11 = sheet.getRange(11, 4, 1, 2);
  const valuesD11E11 = rangeD11E11.getValues()[0];
  const updatedValuesD11E11 = valuesD11E11.map(cell =>
    typeof cell === 'string' ? cell.replace(/yyyy/g, year.toString()) : cell
  );
  rangeD11E11.setValues([updatedValuesD11E11]);
}