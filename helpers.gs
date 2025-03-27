function isRowMergedOrHidden(rowIndex, mergedRanges, hiddenRows) {
  if (mergedRanges && mergedRanges.length > 0) {
    for (let range of mergedRanges) {
      if (range.getRow() <= rowIndex && range.getLastRow() >= rowIndex) {
        return true; // Dòng bị gộp
      }
    }
  }
  return hiddenRows.includes(rowIndex); // Kiểm tra dòng ẩn
}
/**
 * Lấy danh sách các dòng bị ẩn trong sheet
 * @param {Sheet} sheet - Sheet cần kiểm tra
 * @returns {number[]} - Danh sách các chỉ số dòng bị ẩn
 */
function getHiddenRows(sheet) {
  const hiddenRows = [];
  const lastRow = sheet.getLastRow();

  for (let i = 1; i <= lastRow; i++) {
    try {
      const range = sheet.getRange(i, 1, 1, 1); // Lấy một ô đại diện cho dòng i
      if (range.isRowHiddenByUser()) {
        hiddenRows.push(i); // Thêm dòng vào danh sách nếu bị ẩn
      }
    } catch (e) {
      Logger.log(`Không thể kiểm tra dòng ${i}: ${e.message}`);
    }
  }
  return hiddenRows;
}
function copyBatchData(batch, batchRowIndexes, summarySheet, allocationColumn, executionColumn) {
  batch.forEach((row, i) => {
    const [index, target, , , allocation, , execution] = row;
    const summaryRow = findOrCreateSummaryRow(summarySheet, index, target);

    if (allocation) summarySheet.getRange(summaryRow, allocationColumn).setValue(allocation);
    if (execution) summarySheet.getRange(summaryRow, executionColumn).setValue(execution);
  });
}
function processSingleRow(sheet, rowIndex, row, allocationColumn, executionColumn) {
  const [index, target, , , allocation, , execution] = row;
  const summaryRow = findOrCreateSummaryRow(sheet, index, target);

  if (allocation) sheet.getRange(summaryRow, allocationColumn).setValue(allocation);
  if (execution) sheet.getRange(summaryRow, executionColumn).setValue(execution);
}
function findOrCreateSummaryRow(sheet, index, target) {
  const dataRange = sheet.getRange(11, 1, sheet.getLastRow() - 10, 2);
  const values = dataRange.getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === index && values[i][1] === target) {
      return 11 + i; // Trả về dòng hiện có
    }
  }

  // Nếu không tìm thấy, tạo dòng mới
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(index);
  sheet.getRange(newRow, 2).setValue(target);
  return newRow;
}
