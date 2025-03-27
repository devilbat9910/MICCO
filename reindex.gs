/**
 * Khi mở bảng tính, hàm onOpen sẽ tạo menu tùy chỉnh "Tạo báo cáo"
 * với mục "Đánh lại STT CP" gọi hàm reindexSTT.
 */
/** 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tạo báo cáo')
    .addItem('Đánh lại STT CP', 'reindexSTT')
    .addToUi();
}
*/
/**
 * Hàm đánh lại STT cho các phần trong sheet "Bản sao của BC_TCT"
 * dựa theo cấu trúc ở cột A và nội dung nhận dạng (và định dạng) của cột B.
 *
 * – Level 1: Nếu ô cột A chứa một mã hợp lệ (lấy từ sheet "Danh mục sản phẩm")
 *            thì xem đó là tiêu đề Level 1, cập nhật biến currentLevel1 và reset các cấp dưới.
 *
 * – Level 2: Nếu nội dung ô cột B bắt đầu với một trong các từ khóa level2,
 *            thì tăng currentLevel2, reset Level 3 và Level 4,
 *            và ghi STT theo mẫu: currentLevel1.currentLevel2
 *
 * – Level 3: Nếu nội dung ô cột B được định dạng in đậm (bold) (và không phải Level 1 hay Level 2)
 *            thì tăng currentLevel3, reset Level 4, và ghi theo mẫu: currentLevel1.currentLevel2.currentLevel3
 *
 * – Level 4: Nếu ô cột B có nội dung (không rỗng) nhưng không thỏa mãn các điều kiện trên,
 *            thì tăng currentLevel4 và ghi theo mẫu: currentLevel1.currentLevel2.currentLevel3.currentLevel4.
 *
 * Các danh mục sản phẩm (mã sản phẩm hợp lệ) được lấy từ sheet "Danh mục sản phẩm"
 * trong phạm vi dữ liệu (thường là A1:B11 hoặc ít hơn).
 */
function reindexSTT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Đổi tên sheet làm việc thành "Bản sao của BC_TCT"
  const sheet = ss.getSheetByName("BC_TCT");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Lỗi: Không tìm thấy sheet 'Bản sao của BC_TCT'.");
    return;
  }
  
  // Lấy toàn bộ dữ liệu (và định dạng font) của sheet làm việc
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const fontWeights = dataRange.getFontWeights(); // Dùng để kiểm tra định dạng in đậm của cột B
  
  // Lấy danh mục sản phẩm từ sheet "Danh mục sản phẩm"
  const prodSheet = ss.getSheetByName("Danh mục sản phẩm");
  if (!prodSheet) {
    SpreadsheetApp.getUi().alert("Lỗi: Không tìm thấy sheet 'Danh mục sản phẩm'.");
    return;
  }
  const prodData = prodSheet.getDataRange().getValues();
  // Tạo mảng validHeaders từ các mã sản phẩm trong cột A của sheet "Danh mục sản phẩm"
  let validHeaders = [];
  for (let i = 0; i < prodData.length; i++) {
    let code = prodData[i][0];
    if (code && typeof code === 'string') {
      code = code.trim();
      if (code !== "") {
        validHeaders.push(code);
      }
    }
  }
  
  // Định nghĩa các từ khóa cho Level 2 (các header phụ)
  const level2Keywords = ["Sản lượng", "Các chỉ tiêu tiêu hao", "Điện - động lực", "Vật tư chính"];
  
  // Các biến lưu trạng thái của từng cấp
  let currentLevel1 = ""; // Level 1: mã sản phẩm
  let currentLevel2 = 0;    // Level 2: số nguyên dương
  let currentLevel3 = 0;    // Level 3: số nguyên dương
  let currentLevel4 = 0;    // Level 4: số nguyên dương
  
  // Duyệt qua từng dòng trong sheet làm việc
  for (let i = 0; i < data.length; i++) {
    // Lấy giá trị ô ở cột A và cột B
    const colA = data[i][0];
    const colB = data[i][1] ? data[i][1].toString().trim() : "";
    const isBold = (fontWeights[i][1] === "bold");
    
    // Debug (có thể uncomment nếu cần):
    // console.log(`Row ${i+1}: colA = ${colA}, colB = ${colB}, isBold = ${isBold}`);
    
    // --- Xét Level 1: Nếu ô cột A có giá trị nằm trong validHeaders (mã sản phẩm)
    if (typeof colA === 'string' && validHeaders.includes(colA.trim())) {
      currentLevel1 = colA.trim();
      currentLevel2 = 0;  // Reset các cấp dưới
      currentLevel3 = 0;
      currentLevel4 = 0;
      sheet.getRange(i + 1, 1).setValue(currentLevel1);
    }
    // --- Xét Level 2: Nếu ô cột B bắt đầu với một trong các từ khóa level2
    else if (level2Keywords.some(keyword => colB.startsWith(keyword))) {
      currentLevel2++;
      currentLevel3 = 0;
      currentLevel4 = 0;
      sheet.getRange(i + 1, 1).setValue(`${currentLevel1}.${currentLevel2}`);
    }
    // --- Xét Level 3: Nếu ô cột B được định dạng in đậm (bold)
    else if (isBold) {
      currentLevel3++;
      currentLevel4 = 0;
      sheet.getRange(i + 1, 1).setValue(`${currentLevel1}.${currentLevel2}.${currentLevel3}`);
    }
    // --- Xét Level 4: Nếu ô cột B có nội dung (không rỗng) nhưng không thỏa mãn các điều kiện trên
    else if (colB) {
      // Nếu đã có Level 3 thì tăng Level 4
      if (currentLevel3 > 0) {
        currentLevel4++;
        sheet.getRange(i + 1, 1).setValue(`${currentLevel1}.${currentLevel2}.${currentLevel3}.${currentLevel4}`);
      } else {
        // Nếu chưa có Level 3 (trường hợp “mồ côi”), đặt mặc định Level 3 = 1 và Level 4 = 1
        currentLevel3 = 1;
        currentLevel4 = 1;
        sheet.getRange(i + 1, 1).setValue(`${currentLevel1}.${currentLevel2}.${currentLevel3}.${currentLevel4}`);
      }
    }
    // Nếu không có nội dung ở cột B và không rơi vào các trường hợp trên,
    // ta không cập nhật số thứ tự (cột A giữ nguyên).
  }
  
  SpreadsheetApp.getUi().alert("Đã đánh lại STT thành công theo logic mới.");
}