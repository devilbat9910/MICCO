
/**
 * Cập nhật dữ liệu từ worksheet INPUT sang OUTPUT với cấu trúc mới.
 * Chỉ truyền dữ liệu từ các cột tương ứng với tháng được chỉ định.
 * Bỏ qua các giá trị bằng 0 hoặc lỗi (Div 0, Ref, ...).
 * @param {Sheet} inputSheet - Worksheet nguồn (INPUT).
 * @param {Sheet} outputSheet - Worksheet đích (OUTPUT).
 * @param {number} month - Tháng lấy từ tên worksheet INPUT (1-based index).
 */
function processData(inputSheet, outputSheet, month) {
  if (!inputSheet) {
    throw new Error('Input Sheet không tồn tại.');
  }
  if (!outputSheet) {
    throw new Error('Output Sheet không tồn tại.');
  }

  Logger.log(`Processing data for Input Sheet: ${inputSheet.getName()}, Output Sheet: ${outputSheet.getName()}`);

  const inputData = inputSheet.getDataRange().getValues();
  const outputData = outputSheet.getDataRange().getValues();

  // Hàm phụ để kiểm tra Index Lv1
  function isLevel1Index(value) {
    return /^[A-Z]$/.test(value?.trim());
  }

  // Tìm các vùng dựa trên Index Lv1 trong INPUT
  let inputRanges = [];
  let startRow = null;

  for (let i = 12; i < inputData.length; i++) {
    const index = inputData[i][0]?.trim();
    if (isLevel1Index(index)) {
      if (startRow !== null) {
        inputRanges.push({ start: startRow, end: i - 1 });
      }
      startRow = i;
    }
  }
  if (startRow !== null) {
    inputRanges.push({ start: startRow, end: inputData.length - 1 });
  }

  Logger.log(`Detected INPUT ranges: ${JSON.stringify(inputRanges)}`);

  // Tìm các vùng dựa trên Index Lv1 trong OUTPUT
  let outputRanges = [];
  startRow = null;

  for (let i = 12; i < outputData.length; i++) {
    const index = outputData[i][0]?.trim();
    if (isLevel1Index(index)) {
      if (startRow !== null) {
        outputRanges.push({ start: startRow, end: i - 1 });
      }
      startRow = i;
    }
  }
  if (startRow !== null) {
    outputRanges.push({ start: startRow, end: outputData.length - 1 });
  }

  Logger.log(`Detected OUTPUT ranges: ${JSON.stringify(outputRanges)}`);

  // Xác định các cột INPUT và OUTPUT tương ứng với tháng
  const inputColumns = { allocation: 4, execution: 6, percentage: 7 }; // Cột E, G, H
  const startMonthColumn = 8; // Tháng 1 bắt đầu từ cột H
  const outputColumns = {
    allocation: startMonthColumn + (month - 1) * 3,
    execution: startMonthColumn + (month - 1) * 3 + 1,
    percentage: startMonthColumn + (month - 1) * 3 + 2,
  };

  // Truyền dữ liệu giữa các vùng
  inputRanges.forEach((inputRange) => {
    const inputStart = inputRange.start;
    const inputEnd = inputRange.end;
    const inputIndex = inputData[inputStart][0]?.trim();

    const matchingOutputRange = outputRanges.find((outputRange) => {
      const outputStart = outputRange.start;
      const outputIndex = outputData[outputStart][0]?.trim();
      return inputIndex === outputIndex;
    });

    if (matchingOutputRange) {
      const outputStart = matchingOutputRange.start;
      const outputEnd = matchingOutputRange.end;

      Logger.log(`Mapping data from INPUT range (${inputStart}, ${inputEnd}) to OUTPUT range (${outputStart}, ${outputEnd})`);

      for (let i = 0; i <= inputEnd - inputStart; i++) {
        const inputRow = inputData[inputStart + i];
        const outputRowIndex = outputStart + i;

        if (outputRowIndex > outputEnd) break;

        // Truyền dữ liệu từ các cột INPUT sang OUTPUT, bỏ qua giá trị bằng 0 hoặc lỗi
        const allocationValue = inputRow[inputColumns.allocation];
        const executionValue = inputRow[inputColumns.execution];
        const percentageValue = inputRow[inputColumns.percentage];

        if (isValidValue(allocationValue)) {
          outputSheet.getRange(outputRowIndex + 1, outputColumns.allocation).setValue(allocationValue);
        }
        if (isValidValue(executionValue)) {
          outputSheet.getRange(outputRowIndex + 1, outputColumns.execution).setValue(executionValue);
        }
        if (isValidValue(percentageValue)) {
          outputSheet.getRange(outputRowIndex + 1, outputColumns.percentage).setValue(percentageValue);
        }
      }
    } else {
      Logger.log(`No matching OUTPUT range for INPUT Index: ${inputIndex}`);
    }
  });
}

/**
 * Kiểm tra giá trị hợp lệ (không bằng 0 hoặc lỗi).
 * @param {*} value - Giá trị cần kiểm tra.
 * @returns {boolean} - True nếu giá trị hợp lệ, ngược lại false.
 */
function isValidValue(value) {
  return value !== 0 && value !== null && value !== undefined && typeof value !== 'string';
}

/**
 * Tìm hàng tương ứng trong OUTPUT dựa trên "Chỉ tiêu" và "Index".
 * @param {Array[]} outputData - Dữ liệu trong OUTPUT.
 * @param {string} index - Index cần tìm (cột A).
 * @param {string} target - Chỉ tiêu cần tìm (cột B).
 * @param {number} headerRows - Số hàng tiêu đề trong OUTPUT.
 * @returns {number} - Chỉ số hàng tương ứng (0-based index). Nếu không tìm thấy, trả về -1.
 */
function findOutputRow(outputData, index, target, headerRows) {
  for (let i = headerRows; i < outputData.length; i++) {
    const [outputIndex, outputTarget] = outputData[i];
    if (outputIndex === index && outputTarget === target) {
      return i; // Trả về chỉ số hàng tương ứng
    }
  }
  return -1; // Không tìm thấy
}
