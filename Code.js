/**
 * HỆ THỐNG NTTPRO 2026 - TÁC GIẢ: NGUYỄN THANH TÙNG
 */

const MON_COLS = { "TOAN": 5, "LY": 6, "HOA": 7, "SINH": 8, "DIA": 9, "SU": 10, "GDKT-PL": 11, "TIN": 12, "CNCN": 13, "CNNN": 14, "NN": 15 };
const MON_NAMES = { "TOAN": "Toán", "LY": "Vật lí", "HOA": "Hóa học", "SINH": "Sinh học", "DIA": "Địa lí", "SU": "Lịch sử", "GDKT-PL": "GD KT-PL", "TIN": "Tin học", "CNCN": "CN Công nghiệp", "CNNN": "CN Nông nghiệp", "NN": "Ngoại ngữ" };
const TENSHEET_MAP = { "TOAN": "Điểm Toán", "LY": "Điểm Vật lí", "HOA": "Điểm Hóa học", "SINH": "Điểm Sinh học", "DIA": "Điểm Địa lí", "SU": "Điểm Lịch sử", "GDKT-PL": "Điểm Giáo dục KT-PL", "TIN": "Điểm Tin", "CNCN": "Điểm CNCN", "CNNN": "Điểm CNNN", "NN": "Điểm Ngoại ngữ" };

function doGet(e) {
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : '';
  var made = (e && e.parameter && e.parameter.made) ? e.parameter.made : '';
  
  if (mode === 'login') {
    var tmp = HtmlService.createTemplateFromFile('Login');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Đăng nhập - Hệ thống DTC').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (mode === 'exam') {
    var tmp = HtmlService.createTemplateFromFile('Index');
    tmp.made = made; 
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Phòng thi Online').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  } else if (mode === 'packager') {
    var tmp = HtmlService.createTemplateFromFile('Packager');
    return tmp.evaluate().setTitle('Trung Tâm Xử Lý Đề Thi').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    var tmp = HtmlService.createTemplateFromFile('Admin');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Hệ Thống Quản Trị').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// KHÔI PHỤC THUẬT TOÁN ĐĂNG NHẬP ADMIN CHUẨN XÁC
function checkAdminLogin(username, password) {
  try {
    var user = username.toString().trim().toLowerCase(); 
    var pass = password.toString().trim();
    
    // TK Dự phòng
    if (user === "vinhxuan" && pass === "123456") return true;

    // ID file admin chuẩn có số 0
    var adminSheetId = "1T8VgoeqUD4tfXAafjC-t-0VN76lLNccOuNt276fLdfE";
    var ss = SpreadsheetApp.openById(adminSheetId);
    var data = ss.getSheets()[0].getDataRange().getDisplayValues(); 
    
    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
      var dbPass = data[i][1] ? data[i][1].toString().trim() : "";
      if (dbUser !== "" && user === dbUser && pass === dbPass) return true;
    }
    return false;
  } catch (e) {
    return "LỖI BẢO MẬT: " + e.message; 
  }
}

function checkLogin(maHS, matKhau) {
  let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
  if (!dbId) return { status: 'error', success: false, message: "Hệ thống chưa liên kết với File Danh sách học sinh!" };
  
  try {
    let ss = SpreadsheetApp.openById(dbId);
    let khois = ["KHOI_10", "KHOI_11", "KHOI_12"];
    for (let k = 0; k < khois.length; k++) {
      let sheet = ss.getSheetByName(khois[k]);
      if (!sheet) continue;
      let sheetData = sheet.getDataRange().getDisplayValues();
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][0] && sheetData[i][0].toString().trim() === maHS.trim()) {
          if (sheetData[i][3] && sheetData[i][3].toString().trim() === matKhau.trim()) {
            let dsDe = [];
            for (let mon in MON_COLS) {
              let code = (sheetData[i][MON_COLS[mon] - 1] || "").toString().replace(/['"]/g, '').trim();
              if (code) dsDe.push({ mon: mon, tenMon: MON_NAMES[mon], maDe: code });
            }
            return { status: 'success', success: true, studentInfo: { tenHS: sheetData[i][1].toString().trim(), lop: sheetData[i][2].toString().trim(), danhSachDe: dsDe } };
          } else return { status: 'error', success: false, message: "Mật khẩu không chính xác!" };
        }
      }
    }
    return { status: 'error', success: false, message: "Mã học sinh không tồn tại!" };
  } catch (e) { return { status: 'error', success: false, message: "Lỗi CSDL: " + e.message }; }
}

function luuDiemTuDong(duLieu) {
  try {
    let folderId = "1UwHrdypLAfsfiRC5qnE_Esl3DRUFydDS";
    let dataFolder = DriveApp.getFolderById(folderId);
    let ss;
    let files = dataFolder.getFilesByName("Diem_" + duLieu.lop);
    
    if (files.hasNext()) ss = SpreadsheetApp.open(files.next());
    else {
        ss = SpreadsheetApp.create("Diem_" + duLieu.lop);
        DriveApp.getFileById(ss.getId()).moveTo(dataFolder);
    }
    
    let sheetName = TENSHEET_MAP[duLieu.mon] || "Điểm Khác";
    let sheetCanLuu = ss.getSheetByName(sheetName);
    if (!sheetCanLuu) {
        sheetCanLuu = ss.insertSheet(sheetName);
        sheetCanLuu.appendRow(["THỜI GIAN NỘP", "MÃ HS", "HỌ VÀ TÊN", "LỚP", "MÃ ĐỀ", "ĐIỂM TRẮC NGHIỆM", "CHI TIẾT"]);
        sheetCanLuu.getRange("A1:G1").setFontWeight("bold").setBackground("#e65100").setFontColor("white").setHorizontalAlignment("center");
        sheetCanLuu.setFrozenRows(1);
    }
    sheetCanLuu.appendRow([ Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss"), duLieu.maHS, duLieu.tenHS, duLieu.lop, duLieu.maDe, duLieu.diemTN, duLieu.chiTiet ]);
    SpreadsheetApp.flush(); 
    return {success: true, msg: "Lưu điểm thành công!"};
  } catch (e) { return {success: false, msg: e.toString()}; }
}

function doPost(e) {
  try { return ContentService.createTextOutput(JSON.stringify(luuDiemTuDong(JSON.parse(e.postData.contents)))).setMimeType(ContentService.MimeType.JSON);
  } catch (error) { return ContentService.createTextOutput(JSON.stringify({success: false, msg: error.toString()})).setMimeType(ContentService.MimeType.JSON); }
}

// KHÔI PHỤC HÀM PHÁT ĐỀ VÀ THỐNG KÊ
function thucHienGiaoDe(khoi, lop, maDe, mon) {
  try {
    let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
    if (!dbId) return "Lỗi: Hệ thống chưa kết nối với File Danh sách Học sinh.";
    let ss = SpreadsheetApp.openById(dbId);
    let sheet = ss.getSheetByName(khoi);
    if (!sheet) return "Lỗi: Không tìm thấy tab " + khoi + " trong file Học sinh.";
    
    let data = sheet.getDataRange().getValues();
    let colIdx = MON_COLS[mon] - 1; 
    let lopArr = lop.split(',').map(l => l.trim().toUpperCase());
    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      let hsLop = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
      if (lopArr.includes(hsLop)) {
        sheet.getRange(i + 1, colIdx + 1).setValue("'" + maDe); 
        count++;
      }
    }
    return "Thành công! Đã phát mã đề " + maDe + " cho " + count + " học sinh.";
  } catch(e) { return "Lỗi hệ thống: " + e.message; }
}

function layDanhSachHocSinhThi(khoi, lop, mon, made) {
  try {
    let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
    let ssHS = SpreadsheetApp.openById(dbId);
    let sheetHS = ssHS.getSheetByName(khoi);
    if (!sheetHS) return { success: false, message: "Không tìm thấy dữ liệu khối " + khoi };
    
    let dataHS = sheetHS.getDataRange().getValues();
    let dsLop = [];
    for (let i = 1; i < dataHS.length; i++) {
      if (dataHS[i][2] && dataHS[i][2].toString().trim().toUpperCase() === lop.toUpperCase()) {
        dsLop.push({ maHS: dataHS[i][0].toString().trim(), ten: dataHS[i][1].toString().trim(), lop: dataHS[i][2].toString().trim() });
      }
    }
    if(dsLop.length === 0) return { success: false, message: "Không có học sinh nào trong lớp " + lop };
    
    let diemMap = {};
    let files = DriveApp.getFolderById("1UwHrdypLAfsfiRC5qnE_Esl3DRUFydDS").getFilesByName("Diem_" + lop.toUpperCase());
    if (files.hasNext()) {
       let sheetDiem = SpreadsheetApp.open(files.next()).getSheetByName(TENSHEET_MAP[mon] || "Điểm Khác");
       if (sheetDiem) {
         let dataDiem = sheetDiem.getDataRange().getValues();
         for (let i = 1; i < dataDiem.length; i++) { 
            if (dataDiem[i][1]) diemMap[dataDiem[i][1].toString().trim()] = dataDiem[i][5] || ""; 
         }
       }
    }
    
    let result = dsLop.map((hs, i) => {
       let daNop = diemMap.hasOwnProperty(hs.maHS);
       return { stt: i + 1, ten: hs.ten, lop: hs.lop, trangThai: daNop ? "Đã nộp bài" : "Chưa nộp bài", diem: daNop ? diemMap[hs.maHS] : "-" };
    });
    return { success: true, data: result };
  } catch(e) { return { success: false, message: "Lỗi kết nối CSDL: " + e.message }; }
}
