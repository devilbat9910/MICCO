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
  const fontWeights = outputSheet.getDataRange().getFontWeights(); // Lấy thông tin định dạng in đậm

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
        const absoluteRowIndex = outputRowIndex + 1; // 1-based row index for Google Sheets

        if (outputRowIndex > outputEnd) break;

        // Truyền dữ liệu từ các cột INPUT sang OUTPUT, bỏ qua giá trị bằng 0 hoặc lỗi
        const allocationValue = inputRow[inputColumns.allocation];
        const executionValue = inputRow[inputColumns.execution];
        const percentageValue = inputRow[inputColumns.percentage];

        if (isValidValue(allocationValue)) {
          outputSheet.getRange(absoluteRowIndex, outputColumns.allocation).setValue(allocationValue);
        }
        if (isValidValue(executionValue)) {
          outputSheet.getRange(absoluteRowIndex, outputColumns.execution).setValue(executionValue);
        }
        if (isValidValue(percentageValue)) {
          outputSheet.getRange(absoluteRowIndex, outputColumns.percentage).setValue(percentageValue);
        }

        // Áp dụng công thức theo quy tắc mới
        applyFormulas(outputSheet, outputRowIndex, absoluteRowIndex, fontWeights);
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
 * Áp dụng công thức cho cột F và G dựa trên quy tắc.
 * @param {Sheet} outputSheet - Worksheet đích.
 * @param {number} rowIndex - Chỉ số hàng (0-based).
 * @param {number} absoluteRowIndex - Chỉ số hàng (1-based).
 * @param {Array} fontWeights - Mảng chứa thông tin định dạng in đậm.
 */
function applyFormulas(outputSheet, rowIndex, absoluteRowIndex, fontWeights) {
  try {
    // Kiểm tra nội dung và định dạng của cột B
    const cellB = outputSheet.getRange(absoluteRowIndex, 2);
    const cellBValue = cellB.getValue();
    const isBold = fontWeights[rowIndex][1] === "bold";
    
    // Danh sách các nội dung cần loại trừ khỏi quy tắc in đậm
    const excludedBoldContents = ["Sản lượng", "Điện - động lực"];
    
    // Quyết định có áp dụng công thức hay không
    const needsFormula = !isBold || excludedBoldContents.includes(cellBValue);
    
    if (needsFormula) {
      // Áp dụng công thức cho cột F (tính theo tỷ lệ G*1000/E)
      const formulaF = `=G${absoluteRowIndex}*1000/E${absoluteRowIndex}`;
      outputSheet.getRange(absoluteRowIndex, 6).setFormula(formulaF);
      
      // Áp dụng công thức cho cột G (tổng các giá trị trên các tháng)
      const columnLetters = ['I', 'L', 'O', 'R', 'U', 'X', 'AA', 'AD', 'AG', 'AJ', 'AM', 'AP'];
      const sumFormula = `=SUM(${columnLetters.map(letter => `${letter}${absoluteRowIndex}`).join(';')})`;
      outputSheet.getRange(absoluteRowIndex, 7).setFormula(sumFormula);
    } else {
      // Xóa công thức nếu là hàng tiêu đề (in đậm và không nằm trong danh sách loại trừ)
      outputSheet.getRange(absoluteRowIndex, 6).clearContent();
      outputSheet.getRange(absoluteRowIndex, 7).clearContent();
    }
  } catch (error) {
    Logger.log(`Lỗi khi áp dụng công thức cho hàng ${absoluteRowIndex}: ${error.message}`);
  }
}