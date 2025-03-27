/**
 * Kiểm tra xem một ô có được xem là "hợp lệ" hay không.
 * Một ô được xem là hợp lệ nếu:
 *   - Không phải null hoặc chuỗi rỗng,
 *   - Không phải số 0 (bao gồm 0.0),
 *   - Và không chứa chuỗi "div/0" hay "ref!" (không phân biệt chữ hoa thường).
 *
 * @param {*} value - Giá trị cần kiểm tra.
 * @return {boolean} - true nếu ô hợp lệ, false nếu không.
 */
function isNonEmpty(value) {
  if (value === null || value === "") return false;
  if (typeof value === "number" && value === 0) return false;
  if (typeof value === "string") {
    let lower = value.toLowerCase();
    if (lower.indexOf("div/0") !== -1 || lower.indexOf("ref!") !== -1) return false;
  }
  return true;
}

/**
 * Thu gọn báo cáo theo logic:
 *
 * - Lượt 1: Chỉ xét vùng của mỗi nhóm báo cáo (được xác định bởi dòng cấp 1)
 *   + Trong nhóm, tìm dòng con có cột B = "Sản lượng"
 *   + Lấy tập con của "Sản lượng" (các dòng có index bắt đầu bằng [sanLuong_index] + ".")
 *   + Nếu không có dòng con nào có (cột E hoặc G hợp lệ) thì toàn bộ nhóm bị đánh dấu ẩn.
 *   + Ngược lại, nhóm ban đầu được gán trạng thái "hiện".
 *
 * - Lượt 2: Duyệt các dòng từ tầng sâu nhất lên tầng cao:
 *   + Nếu một dòng không có con trực tiếp "hiện" và chính nó không có giá trị hợp lệ ở cột G
 *     thì đánh dấu dòng đó là ẩn.
 *
 * - Bổ sung: Nếu một dòng (cháu) đang hiện thì tất cả các dòng tổ tiên (cha, ông, …)
 *   đều phải được hiện, bất kể giá trị của chính chúng.
 *
 * - Cuối cùng, gom nhóm các hàng cần ẩn và gọi hideRows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} worksheet - Worksheet cần thu gọn.
 * @return {number} - Số hàng đã ẩn.
 */
function shrinkReport(worksheet) {
  const dataRange = worksheet.getDataRange();
  const data = dataRange.getValues();
  const headers = 10; // Hàng tiêu đề: từ hàng 1 đến hàng 10
  const lastRow = worksheet.getLastRow();

  // --- Bước 1: Thu thập thông tin các dòng (từ hàng 13 trở đi) ---
  // Mỗi đối tượng rowInfo có dạng:
  // { rowNum, indexStr, level, colB, colE, colG }
  let rowsInfo = [];
  for (let i = headers; i < lastRow; i++) {
    const rowNum = i + 1;
    const indexVal = data[i][0]; // Cột A: chỉ số dạng "1", "1.1", "1.1.1",…
    let indexStr = (indexVal ? indexVal.toString() : "");
    let level = (indexStr ? indexStr.split('.').length : 0);
    rowsInfo.push({
      rowNum: rowNum,
      indexStr: indexStr,
      level: level,
      colB: data[i][1], // Cột B: mô tả
      colE: data[i][4], // Cột E
      colG: data[i][6]  // Cột G
    });
  }

  // --- Bước 2: Phân nhóm theo nhóm cấp 1 ---
  // Mỗi nhóm có:
  // { level1Row: dòng cấp 1, groupRows: [các dòng thuộc nhóm] }
  let groups = [];
  let currentGroup = null;
  for (let row of rowsInfo) {
    if (row.indexStr && row.level === 1) {
      // Nếu gặp dòng cấp 1: bắt đầu nhóm mới
      if (currentGroup) groups.push(currentGroup);
      currentGroup = { level1Row: row, groupRows: [row] };
    } else {
      if (currentGroup) {
        currentGroup.groupRows.push(row);
      } else {
        // Nếu không có dòng cấp 1 nào trước, nhóm riêng (sẽ bị ẩn sau)
        groups.push({ level1Row: null, groupRows: [row] });
      }
    }
  }
  if (currentGroup) groups.push(currentGroup);

  // --- Bước 3: Lượt 1 – Xét nhóm "Sản lượng" ---
  // Với mỗi nhóm, tìm dòng con có cột B = "Sản lượng" (so sánh theo chữ thường, trim)
  // Sau đó, lấy các dòng liền sau dòng "Sản lượng" mà có chỉ số bắt đầu bằng
  // [sanLuong_index] + "." làm tập con của "Sản lượng" và kiểm tra dữ liệu (E hoặc G).
  groups.forEach(group => {
    // Nếu nhóm không có dòng cấp 1, đánh dấu ẩn luôn.
    if (!group.level1Row) {
      group.visibilityFirst = false;
      return;
    }
    let sanLuongRow = null;
    for (let row of group.groupRows) {
      if (row.level > 1 && typeof row.colB === "string" &&
          row.colB.trim().toLowerCase() === "sản lượng") {
        sanLuongRow = row;
        break;
      }
    }
    if (!sanLuongRow) {
      // Nếu không có dòng "Sản lượng", đánh dấu ẩn cả nhóm.
      group.visibilityFirst = false;
      return;
    }
    // Thu thập các dòng con của "Sản lượng"
    let sanLuongChildren = [];
    let startCollect = false;
    for (let row of group.groupRows) {
      if (row.rowNum === sanLuongRow.rowNum) {
        startCollect = true;
        continue;
      }
      if (startCollect) {
        if (row.indexStr &&
            row.indexStr.indexOf(sanLuongRow.indexStr + ".") === 0 &&
            row.level > sanLuongRow.level) {
          sanLuongChildren.push(row);
        } else {
          // Giả sử các con của "Sản lượng" liền nhau.
          break;
        }
      }
    }
    // Nếu không có dòng con nào có giá trị hợp lệ (cột E hoặc G) thì đánh dấu nhóm ẩn.
    let hasValidChild = sanLuongChildren.some(child => (isNonEmpty(child.colE) || isNonEmpty(child.colG)));
    group.visibilityFirst = hasValidChild;
  });

  // --- Bước 4: Gán trạng thái "hiện" ban đầu theo Lượt 1 ---
  // Nếu nhóm bị đánh dấu ẩn (visibilityFirst = false) thì toàn bộ các dòng của nhóm được gán false;
  // Ngược lại, gán true.
  let visibility = {}; // key: rowNum, value: true (hiện) hay false (ẩn)
  groups.forEach(group => {
    if (group.visibilityFirst === false) {
      group.groupRows.forEach(row => { visibility[row.rowNum] = false; });
    } else {
      group.groupRows.forEach(row => { visibility[row.rowNum] = true; });
    }
  });

  // --- Bước 5: Lượt 2 – Duyệt từ tầng sâu nhất lên (bottom-up) ---
  // Xây dựng quan hệ cha-con: với mỗi dòng có level > 1, cha của nó có chỉ số được xác định
  // bằng cách loại bỏ phần sau dấu chấm cuối cùng.
  let childrenMap = {}; // key: chỉ số cha, value: mảng các dòng con
  rowsInfo.forEach(row => {
    if (row.level > 1) {
      let parts = row.indexStr.split('.');
      parts.pop();
      let parentIndex = parts.join('.');
      if (!childrenMap[parentIndex]) childrenMap[parentIndex] = [];
      childrenMap[parentIndex].push(row);
    }
  });

  // Sắp xếp các dòng theo thứ tự giảm dần của level (tầng sâu nhất xử lý trước)
  let sortedRows = rowsInfo.slice().sort((a, b) => {
    if (b.level !== a.level) return b.level - a.level;
    return b.rowNum - a.rowNum;
  });

  // Với mỗi dòng, nếu dòng đang hiện mà không có con trực tiếp nào hiện,
  // thì kiểm tra giá trị của chính nó (chỉ cột G). Nếu không hợp lệ, đánh dấu ẩn.
  sortedRows.forEach(row => {
    if (!visibility[row.rowNum]) return; // nếu đã ẩn từ lượt 1 hoặc trước đó
    let children = childrenMap[row.indexStr] || [];
    let anyChildVisible = children.some(child => visibility[child.rowNum]);
    if (!anyChildVisible) {
      // Nếu dòng không có con hiện và chính nó không có giá trị hợp lệ ở cột G
      if (!isNonEmpty(row.colG)) {
        visibility[row.rowNum] = false;
      }
    }
  });

  // --- Bước 6: Truy vết ngược từ con lên cha (nâng dần tầng) ---
  // Nếu một dòng (cháu) đang hiện thì tất cả các dòng tổ tiên (cha, ông, …) cũng phải được hiện.
  let rowByIndex = new Map();
  rowsInfo.forEach(row => {
    if (row.indexStr) {
      rowByIndex.set(row.indexStr, row);
    }
  });

  // Hàm đệ quy để truyền trạng thái "hiện" lên tổ tiên.
  function propagateParentVisibility(row) {
    let parts = row.indexStr.split('.');
    if (parts.length > 1) {
      parts.pop();
      let parentIndex = parts.join('.');
      if (rowByIndex.has(parentIndex)) {
        let parentRow = rowByIndex.get(parentIndex);
        if (!visibility[parentRow.rowNum]) {
          visibility[parentRow.rowNum] = true;
        }
        propagateParentVisibility(parentRow);
      }
    }
  }

  // Với mỗi dòng đang hiện (với level > 1) thì truyền trạng thái lên.
  sortedRows.forEach(row => {
    if (visibility[row.rowNum] && row.level > 1) {
      propagateParentVisibility(row);
    }
  });

  // --- Bước 7: Gom nhóm các hàng cần ẩn và thực hiện ẩn ---
  let rowsToHide = [];
  rowsInfo.forEach(row => {
    if (!visibility[row.rowNum]) {
      rowsToHide.push(row.rowNum);
    }
  });
  // Loại bỏ trùng lặp và sắp xếp theo thứ tự tăng dần
  rowsToHide = Array.from(new Set(rowsToHide)).sort((a, b) => a - b);

  // Gom nhóm các hàng liền kề để gọi hideRows hiệu quả
  if (rowsToHide.length > 0) {
    let startRow = rowsToHide[0];
    let count = 1;
    for (let i = 1; i < rowsToHide.length; i++) {
      if (rowsToHide[i] === rowsToHide[i - 1] + 1) {
        count++;
      } else {
        worksheet.hideRows(startRow, count);
        startRow = rowsToHide[i];
        count = 1;
      }
    }
    worksheet.hideRows(startRow, count);
  }

  return rowsToHide.length;
}

/**
 * Giao diện thu gọn báo cáo qua UI Google Sheets.
 */
function shrinkReportUI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const newestWorksheet = sheets[sheets.length - 1];
  const newestWorksheetName = newestWorksheet.getName();

  const response = ui.prompt(
    'Thu gọn báo cáo',
    `Báo cáo mới nhất là "${newestWorksheetName}". Nhập tên worksheet nếu muốn chọn khác hoặc để trống để sử dụng báo cáo mới nhất:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputName = response.getResponseText().trim();
    let targetWorksheet;
    if (inputName) {
      targetWorksheet = ss.getSheetByName(inputName);
      if (!targetWorksheet) {
        ui.alert(`Worksheet "${inputName}" không tồn tại. Vui lòng kiểm tra lại!`);
        return;
      }
    } else {
      targetWorksheet = newestWorksheet;
    }

    const confirm = ui.alert(
      'Xác nhận',
      `Bạn có chắc chắn muốn thu gọn báo cáo trong worksheet "${targetWorksheet.getName()}"?`,
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      const hiddenCount = shrinkReport(targetWorksheet);
      ui.alert(`Đã hoàn thành. Đã ẩn ${hiddenCount} hàng.`);
    } else {
      ui.alert('Thu gọn báo cáo đã bị huỷ.');
    }
  } else {
    ui.alert('Thu gọn báo cáo đã bị huỷ.');
  }
}