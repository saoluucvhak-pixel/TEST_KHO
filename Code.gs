/**
 * HỆ THỐNG QUẢN LÝ HIỆP VĨNH AN - BẢN FULL 100% (IMPORT + UPDATE 4 SHEETS)
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Hệ Thống Quản Lý Hiệp Vĩnh An')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- HÀM 1: IMPORT DỮ LIỆU TỪ WEB APP ---
function processImportGeneral(rawData, target) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (target === 'THUE') ? "TK_THUE" : "TK_NB";
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const cleanedData = [];
  
  // Duyệt từ dòng 6 Excel (index 5)
  for (let i = 5; i < rawData.length; i++) {
    let row = rawData[i];
    if (!row || row.length < 2) continue;
    
    let colA_Ex = String(row[0] || "").trim();
    let colB_Ex = String(row[1] || "").trim();
    
    if (colA_Ex.includes("Số dòng") || colB_Ex.includes("Số dòng") || colA_Ex.includes("Cộng") || colB_Ex.includes("Cộng")) break;
    if (colB_Ex === "" || colB_Ex === "Mã hàng") continue;

    if (target === 'NB') {
      // NỘI BỘ: Lấy B-N Excel (idx 1-13) sang A-M Sheet
      cleanedData.push(row.slice(1, 14));
    } else {
      // THUẾ: Lấy A-L Excel (idx 0-11) sang A-L Sheet
      cleanedData.push(row.slice(0, 12));
    }
  }

  if (cleanedData.length > 0) {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    sheet.getRange(2, 1, cleanedData.length, cleanedData[0].length).setValues(cleanedData);
    return "✅ Đã Import " + cleanedData.length + " dòng vào " + sheetName;
  }
  throw new Error("Không tìm thấy dữ liệu hợp lệ.");
}

// --- HÀM 2: CẬP NHẬT TỔNG LỰC 4 SHEET ---
function updateThueAndCheckRemain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileDM_ID = "112xJKUdkgwsjPlTRqW5oSuvH5Mv560o2EhqiYMzMBAo";
  
  let ssDM;
  try { ssDM = SpreadsheetApp.openById(fileDM_ID); } catch (e) {
    throw new Error("Lỗi kết nối file DM_CHUNG!");
  }
  
  const shTkNB = ss.getSheetByName("TK_NB");
  const shTkThue = ss.getSheetByName("TK_THUE");
  const shDM = ssDM.getSheetByName("DM_CHUNG");
  if (!shTkNB || !shTkThue) throw new Error("Chưa có dữ liệu TK_NB / TK_THUE. Vui lòng Import Excel trước khi Cập Nhật Tổng Lực.");
  if (!shDM) throw new Error("Không tìm thấy sheet 'DM_CHUNG' trong file DM_CHUNG!");
  const shThueNB = ss.getSheetByName("THUE_NB") || ss.insertSheet("THUE_NB");
  const shTonNB = ss.getSheetByName("TONKHO_NB") || ss.insertSheet("TONKHO_NB");
  const shTonThue = ss.getSheetByName("TONKHO_THUE") || ss.insertSheet("TONKHO_THUE");
  const shRemain = ss.getSheetByName("KHO_NB_CONLAI") || ss.insertSheet("KHO_NB_CONLAI");

  const dataTkNB = shTkNB.getDataRange().getValues();
  const dataTkThue = shTkThue.getDataRange().getValues();
  const dataDM = shDM.getDataRange().getValues();
  // ID tuần tự theo run (prefix + thời điểm chạy + số thứ tự dòng) để đảm bảo không trùng,
  // thay cho số ngẫu nhiên 6 chữ số cũ (dễ đụng nhau khi có hàng trăm/ngàn dòng).
  const runStamp = Date.now().toString(36).toUpperCase();
  const genID = (i) => "HVA-" + runStamp + "-" + i;

  const mapMaBSangChungA = new Map(), mapTenCSangChungA = new Map();
  const mapTenCSangMaB = new Map(), mapMaBSangMaL = new Map();

  for (let i = 1; i < dataDM.length; i++) {
    let maA = String(dataDM[i][0] || "").trim();
    let maB = String(dataDM[i][1] || "").trim();
    let tenC = String(dataDM[i][2] || "").trim();
    let maL = String(dataDM[i][11] || "").trim();
    if (maA) {
      if (maB) mapMaBSangChungA.set(maB, maA);
      if (tenC) mapTenCSangChungA.set(tenC, maA);
    }
    if (tenC) mapTenCSangMaB.set(tenC, maB);
    if (maB) mapMaBSangMaL.set(maB, maL);
  }

  // Cập nhật TONKHO_THUE & TONKHO_NB
  const resTonThue = dataTkThue.map((r, i) => i == 0 ? [...r, "Mã Chung", "ID_KEY"] : [...r, mapMaBSangChungA.get(String(r[1]).trim()) || "", genID("T" + i)]);
  shTonThue.clear().getRange(1, 1, resTonThue.length, resTonThue[0].length).setValues(resTonThue);

  const resTonNB = dataTkNB.map((r, i) => i == 0 ? [...r, "Mã Chung", "ID_KEY"] : [...r, mapTenCSangChungA.get(String(r[2]).trim()) || "", genID("N" + i)]);
  shTonNB.clear().getRange(1, 1, resTonNB.length, resTonNB[0].length).setValues(resTonNB);

  // Xử lý THUE_NB (Có D, K diễn giải)
  const sumThue = new Map();
  for (let k = 1; k < dataTkThue.length; k++) {
    let mB = String(dataTkThue[k][1]).trim();
    if (!mB) continue;
    if (!sumThue.has(mB)) sumThue.set(mB, { ten: dataTkThue[k][2], dvt: dataTkThue[k][3], sl: 0, tt: 0, ct: [] });
    let o = sumThue.get(mB);
    o.sl += Number(dataTkThue[k][10]) || 0;
    o.tt += Number(dataTkThue[k][11]) || 0;
    if (dataTkThue[k][0]) o.ct.push(dataTkThue[k][0] + " (" + dataTkThue[k][10] + ")");
  }

  const sumNB = new Map(), remData = [];
  for (let j = 1; j < dataTkNB.length; j++) {
    let tenC = String(dataTkNB[j][2]).trim(), slL = Number(dataTkNB[j][11]) || 0, tKho = String(dataTkNB[j][1]).trim();
    let maB = mapTenCSangMaB.get(tenC);
    if (maB && sumThue.has(maB)) {
      if (!sumNB.has(maB)) sumNB.set(maB, { sl: 0, kho: [] });
      let o = sumNB.get(maB); o.sl += slL;
      if (slL != 0) o.kho.push(tKho + " (" + slL + ")");
    } else {
      remData.push([dataTkNB[j][1], dataTkNB[j][2], dataTkNB[j][3], dataTkNB[j][4], slL]);
    }
  }

  const report = [];
  sumThue.forEach((v, maB) => {
    let nb = sumNB.get(maB) || { sl: 0, kho: [] };
    let clI = v.sl - nb.sl;
    report.push([maB, v.ten, v.dvt, v.ct.join(", "), v.sl, v.tt, (v.sl?Math.round(v.tt/v.sl):0), nb.sl, clI, (clI>0?"Thừa":clI<0?"Thiếu":"Khớp"), nb.kho.join(", "), mapMaBSangMaL.get(maB)||""]);
  });

  shThueNB.clear();
  shThueNB.getRange(1,1,1,12).setValues([["Mã Thuế","Tên Hàng","ĐVT","Diễn giải Thuế (Số CT)","SL Thuế","Thành tiền","Giá vốn","SL Nội Bộ","Chênh lệch","Ghi chú","Diễn giải Nội Bộ (Kho)","Mã Chung (L)"]]).setFontWeight("bold");
  if (report.length > 0) shThueNB.getRange(2,1,report.length,12).setValues(report);

  shRemain.clear().getRange(1,1,1,5).setValues([["Tên Kho","Mã Hàng","Tên Hàng","Quy cách","Số lượng"]]).setFontWeight("bold");
  if (remData.length > 0) shRemain.getRange(2,1,remData.length,5).setValues(remData);

  return "🚀 Đã Import & Cập nhật thành công 4 Sheet!";
}

/**
 * Tự động tạo Menu khi mở file Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 HỆ THỐNG HVA')
      .addItem('📱 Mở Bảng Điều Khiển (Sidebar)', 'showImportSidebar')
      .addSeparator()
      .addItem('🔄 Cập Nhật & Tổng Lực 4 Sheet', 'runUpdateWithNotification')
      .addSeparator()
      .addItem('📖 Hướng Dẫn Sử Dụng', 'showHelp')
      .addToUi();
}

/**
 * Hàm mở Sidebar ở phía bên phải Google Sheets
 */
function showImportSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('Index') // Đảm bảo bạn có file Index.html
        .evaluate()
        .setTitle('BẢNG ĐIỀU KHIỂN HVA')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    // Hiển thị Sidebar bên phải
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Lỗi: Không tìm thấy file 'Index.html' hoặc file có lỗi cú pháp.");
  }
}

/**
 * Hàm trung gian để chạy cập nhật và thông báo kết quả qua Alert
 */
function runUpdateWithNotification() {
  const ui = SpreadsheetApp.getUi();
  
  // Hiển thị thông báo nhỏ ở góc dưới bên phải để người dùng biết hệ thống đang chạy
  SpreadsheetApp.getActiveSpreadsheet().toast("Đang xử lý dữ liệu 4 sheet... Vui lòng đợi trong giây lát.", "🕒 ĐANG CHẠY", 10);
  
  try {
    // Gọi hàm xử lý chính (updateThueAndCheckRemain phải nằm trong cùng dự án này)
    const result = updateThueAndCheckRemain(); 
    ui.alert('THÀNH CÔNG', result, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('LỖI HỆ THỐNG', "Chi tiết lỗi: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Hiển thị hộp thoại hướng dẫn nhanh
 */
function showHelp() {
  const msg = "HƯỚNG DẪN NHANH:\n\n" +
              "1. Chọn 'Mở Bảng Điều Khiển': Giao diện Import File Excel sẽ hiện ra ở cột bên phải.\n" +
              "2. Thực hiện Import dữ liệu MISA từ Web App vào Sidebar ghi vào 2 sheet (TK_THUE, TK_NB).\n" +
              "3. Chọn 'Cập Nhật & Tổng Lực': Hệ thống tự động đồng bộ 4 Sheet (TONKHO_NB, TONKHO_THUE, THUE_NB, KHO_NB_CONLAI).";
  
  SpreadsheetApp.getUi().alert('📖 HƯỚNG DẪN SỬ DỤNG', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
