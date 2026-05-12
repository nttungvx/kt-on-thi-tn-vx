/**
 * HỆ THỐNG NTTPRO 2026 - TÁC GIẢ: NGUYỄN THANH TÙNG
 * PHIÊN BẢN HOÀN CHỈNH 2025
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
    return tmp.evaluate().setTitle('Học sinh đăng nhập - Hệ thống DTC').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (mode === 'exam') {
    var tmp = HtmlService.createTemplateFromFile('Index');
    tmp.made = made; 
    tmp.appUrl = ScriptApp.getService().getUrl();
    tmp.hinhNen = PropertiesService.getScriptProperties().getProperty('HINH_NEN_URL') || ''; 
    return tmp.evaluate().setTitle('Phòng thi Online - Hệ thống DTC').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (mode === 'packager') {
    var tmp = HtmlService.createTemplateFromFile('Packager');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Đóng Gói Đề - ÔN THI TỐT NGHIỆP THPT').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    var tmp = HtmlService.createTemplateFromFile('Admin');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('HỆ THỐNG KIỂM TRA & THI TRỰC TUYẾN').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function checkAdminLogin(username, password) {
  var user = username.toString().trim();
  var pass = password.toString().trim();
  if (user === "vinhxuan" && pass === "123456") return true;
  try {
    var rootFolders = DriveApp.getFoldersByName("DU_AN_NTTPRO_2026"); 
    if (!rootFolders.hasNext()) return false; 
    var rootFolder = rootFolders.next();
    var files = rootFolder.searchFiles("title = 'admin' and mimeType = '" + MimeType.GOOGLE_SHEETS + "'"); 
    var adminSheet;
    
    if (!files.hasNext()) {
      adminSheet = SpreadsheetApp.create("admin");
      DriveApp.getFileById(adminSheet.getId()).moveTo(rootFolder);
      var sheet = adminSheet.getSheets()[0];
      sheet.setName("Danh sách Quản trị viên");
      sheet.appendRow(["TÀI KHOẢN", "MẬT KHẨU", "GHI CHÚ / HỌ TÊN GV"]);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#1976d2").setFontColor("white");
      sheet.appendRow(["vinhxuan", "123456", "Tài khoản mặc định"]);
      sheet.setFrozenRows(1);
    } else { adminSheet = SpreadsheetApp.open(files.next()); }
    
    var data = adminSheet.getSheets()[0].getDataRange().getValues();
    for (var i = 0; i < data.length; i++) { 
      if (data[i][0] != null && data[i][1] != null && data[i][0].toString().trim() === user && data[i][1].toString().trim() === pass) return true;
    }
    return false; 
  } catch(e) { return false; }
}

function generateExamFromAI(base64Data, fileName, struct) {
  try {
    var GEMINI_API_KEY = layApiKey();
    if (!GEMINI_API_KEY || GEMINI_API_KEY === "") {
      return "<div style='color:#ef4444; padding:20px; background:rgba(239, 68, 68, 0.1); border-radius:8px;'>❌ <b>Lỗi API Key:</b> Thầy chưa cấu hình mã API Key. Vui lòng dán mã vào ô cấu hình góc trái dưới cùng và nhấn 'LƯU MÃ BẢO MẬT'.</div>";
    }

    var base64String = base64Data.split(',')[1] || base64Data;
    var listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + GEMINI_API_KEY;
    var listRes = UrlFetchApp.fetch(listUrl, {muteHttpExceptions: true});
    var listData = JSON.parse(listRes.getContentText());
    var modelName = "";
    if (listData.error) return "<div style='color:#ef4444; padding:20px; background:rgba(239, 68, 68, 0.1); border-radius:8px;'>❌ <b>Lỗi xác thực API:</b> " + listData.error.message + "</div>";
    
    if (listData.models) {
      for (var i = 0; i < listData.models.length; i++) {
        var mName = listData.models[i].name;
        if (listData.models[i].supportedGenerationMethods && listData.models[i].supportedGenerationMethods.includes("generateContent") && mName.includes("gemini")) {
            modelName = mName.replace("models/", "");
            if (mName.includes("flash")) break; 
        }
      }
    }
    if (modelName === "") return "<div style='color:#ef4444; padding:20px; background:rgba(239, 68, 68, 0.1); border-radius:8px;'>❌ <b>Lỗi API Key:</b> Không tìm thấy AI model.</div>";
    
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + GEMINI_API_KEY;
    var prompt = "Bạn là một chuyên gia số hóa đề thi môn " + struct.name + ". Hãy đọc file PDF đính kèm (bao gồm Đề thi và Hướng dẫn giải) và trích xuất ra HTML thuần túy. Tuân thủ:\n";
    prompt += "1. KHÔNG dùng thẻ <html>, <head>, <body> hay bọc markdown.\n";
    prompt += "2. Mọi công thức BẮT BUỘC chuyển sang LaTeX ($...$ hoặc $$...$$).\n";
    prompt += "3. ĐỊNH DẠNG ĐÁP ÁN (Chuẩn Bộ GD&ĐT):\n";
    prompt += "   - PHẦN I: Trình bày 4 đáp án trên cùng 1 hoặc 2 hàng ngang (tùy độ dài). Các chữ cái đáp án BẮT BUỘC in đậm và có dấu chấm theo sau: <b>A.</b> , <b>B.</b> , <b>C.</b> , <b>D.</b>\n";
    prompt += "   - PHẦN II: 4 ý của câu Đúng/Sai BẮT BUỘC in đậm và có dấu ngoặc đơn: <b>a)</b> , <b>b)</b> , <b>c)</b> , <b>d)</b>\n";
    prompt += "   - KHÔNG bọc đáp án trong class 'circle-opt' nữa.\n";
    prompt += "4. Tổ chức lại cấu trúc đề thành:\n";
    if (struct.p1 > 0) prompt += "<h5 style='color:#ff6d00; font-size:18px;'>PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn</h5>\n";
    if (struct.p2 > 0) prompt += "<h5 style='color:#ff6d00; font-size:18px; margin-top:30px;'>PHẦN II. Câu trắc nghiệm đúng sai</h5>\n";
    if (struct.p3 > 0) prompt += "<h5 style='color:#ff6d00; font-size:18px; margin-top:30px;'>PHẦN III. Câu trắc nghiệm trả lời ngắn</h5>\n";
    prompt += "5. Mỗi câu hỏi bọc trong <div style='margin-bottom: 20px;'>. In đậm 'Câu X:'.\n";
    prompt += "6. LỜI GIẢI CHI TIẾT: Tìm phần Hướng dẫn giải ở cuối file PDF. Khớp Lời giải của câu nào vào đúng bên trong khối <div> của câu hỏi đó. BẮT BUỘC bọc toàn bộ nội dung lời giải trong thẻ <div class='loi-giai-chi-tiet' style='display:none;'>...</div>.";
    
    var payload = { "contents": [{ "parts": [ {"text": prompt}, { "inlineData": { "mimeType": "application/pdf", "data": base64String } } ] }], "generationConfig": { "temperature": 0.1 } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    var maxRetries = 4;
    var attempt = 0; var json; var success = false;
    while (attempt < maxRetries && !success) {
      var response = UrlFetchApp.fetch(url, options);
      json = JSON.parse(response.getContentText());
      if (json.error) {
        let errMsg = json.error.message;
        if (errMsg.includes("high demand") || errMsg.includes("quota") || errMsg.includes("429") || errMsg.includes("503")) {
          attempt++;
          if (attempt < maxRetries) Utilities.sleep(5000 * attempt);
        } else return "<div style='color:#ef4444; padding:20px; background:rgba(239, 68, 68, 0.1); border-radius:8px;'>❌ <b>Lỗi AI:</b> " + errMsg + "</div>";
      } else success = true;
    }

    if (!success) return "<div style='color:#ef4444; padding:20px; background:rgba(239, 68, 68, 0.1); border-radius:8px;'>❌ <b>Máy chủ AI của Google đang bị giới hạn hoặc quá tải.</b> Vui lòng thử lại!</div>";
    return json.candidates[0].content.parts[0].text.replace(/```html/g, "").replace(/```/g, "");

  } catch (e) { return "<div style='color:#ef4444; padding:20px;'>❌ <b>Lỗi hệ thống kết nối:</b> " + e.toString() + "</div>"; }
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function luuDeLenDrive(made, htmlContent, monKey) {
  try {
    var cleanMade = made.toString().replace(/['"]/g, '').trim().toUpperCase();
    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), "DU_AN_NTTPRO_2026");
    var deCacMonFolder = getOrCreateFolder(rootFolder, "DE_CAC_MON");
    var onThiFolder = getOrCreateFolder(deCacMonFolder, "ON_THI_TN");
    var finalFolder = monKey ? getOrCreateFolder(onThiFolder, monKey) : onThiFolder;
    var fileName = "NTTPRO_DE_" + cleanMade + ".txt";
    
    var files = finalFolder.getFilesByName(fileName);
    while (files.hasNext()) { files.next().setTrashed(true); }
    
    finalFolder.createFile(fileName, htmlContent, MimeType.PLAIN_TEXT);
    return "✅ Đã LƯU ĐỀ mã [" + cleanMade + "] thành công vào hệ thống!";
  } catch (e) { return "❌ Lỗi lưu Drive: " + e.toString(); }
}

function layDeTuDrive(made) {
  try {
    var cleanMade = made.toString().replace(/['"]/g, '').trim().toUpperCase();
    var fileName = "NTTPRO_DE_" + cleanMade + ".txt";
    var activeFiles = DriveApp.searchFiles("title = '" + fileName + "' and trashed = false");
    if (activeFiles.hasNext()) return { success: true, data: activeFiles.next().getBlob().getDataAsString() };
    var allFiles = DriveApp.getFilesByName(fileName);
    if (allFiles.hasNext()) return { success: true, data: allFiles.next().getBlob().getDataAsString() };
    return { success: false, message: "Chưa tìm thấy đề thi mã [" + cleanMade + "]. Vui lòng kiểm tra lại hệ thống lưu trữ!" };
  } catch (e) { return { success: false, message: "Lỗi tải đề: " + e.toString() }; }
}

function thucHienGiaoDe(khoi, chuoiLop, maDe, monKey) {
  let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
  if(!dbId) return "Lỗi: Chưa chạy hàm khởi tạo Database hoặc Database chưa được liên kết!";
  let sheet = SpreadsheetApp.openById(dbId).getSheetByName(khoi);
  let data = sheet.getDataRange().getValues(), count = 0;
  let colIndex = MON_COLS[monKey];
  if(!colIndex) return "Lỗi: Môn học không hợp lệ!";
  
  let danhSachLop = chuoiLop.split(',').map(l => l.trim().toUpperCase()).filter(l => l !== "");
  if (danhSachLop.length === 0) return "Lỗi: Vui lòng nhập tên lớp hợp lệ!";
  let cleanMaDe = maDe.toString().replace(/['"]/g, '').trim().toUpperCase();
  
  for (let i = 1; i < data.length; i++) {
    if (danhSachLop.includes(data[i][2].toString().trim().toUpperCase())) { 
        sheet.getRange(i + 1, colIndex).setValue(cleanMaDe);
        count++; 
    }
  }
  if(count === 0) return "Không tìm thấy học sinh thuộc các lớp [" + danhSachLop.join(", ") + "]";
  return "Thành công! Đã phát đề môn " + MON_NAMES[monKey] + " cho " + count + " học sinh lớp: " + danhSachLop.join(", ");
}

function kiemTraDangNhapHS(maHS, matKhau) {
  let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
  if(!dbId) return {success: false, message: "Hệ thống chưa thiết lập dữ liệu!"};
  
  let ss = SpreadsheetApp.openById(dbId);
  let khois = ["KHOI_10", "KHOI_11", "KHOI_12"];
  
  for(let k = 0; k < khois.length; k++) {
    let sheet = ss.getSheetByName(khois[k]);
    if(!sheet) continue;
    
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] != null && data[i][0].toString().trim() === maHS.trim()) {
        if (data[i][3] != null && data[i][3].toString().trim() === matKhau.trim()) {
          
          let dsDe = [];
          let monKeys = ["TOAN", "LY", "HOA", "SINH", "DIA", "SU", "GDKT-PL", "TIN", "CNCN", "CNNN", "NN"];
          for(let j = 0; j < monKeys.length; j++) {
              let rawCode = data[i][4 + j] ? data[i][4 + j].toString() : "";
              let code = rawCode.replace(/['"]/g, '').trim();
              if (code !== "") dsDe.push({ mon: monKeys[j], tenMon: MON_NAMES[monKeys[j]], maDe: code });
          }
          if(dsDe.length === 0) return {success: false, message: "Hiện tại bạn chưa được giao bài thi nào!"};
          return { success: true, studentInfo: { tenHS: data[i][1].toString().trim(), lop: data[i][2].toString().trim(), khoi: khois[k], danhSachDe: dsDe } };
        } else {
          return {success: false, message: "Mật khẩu không chính xác!"};
        }
      }
    }
  }
  return {success: false, message: "Tài khoản không tồn tại!"};
}

function luuDiemTuDong(duLieu) {
  try {
    let tenThuMucGoc = "DU_AN_NTTPRO_2026", tenThuMucDiem = "Data_Diem", tenFileDiem = "Diem_" + duLieu.lop;
    let rootFolders = DriveApp.getFoldersByName(tenThuMucGoc);
    if (!rootFolders.hasNext()) return {success: false, msg: "Không tìm thấy thư mục gốc!"};
    let rootFolder = rootFolders.next();
    let dataFolders = rootFolder.getFoldersByName(tenThuMucDiem);
    let dataFolder = dataFolders.hasNext() ? dataFolders.next() : rootFolder.createFolder(tenThuMucDiem);
    const THUTU_MON = ["TOAN", "LY", "HOA", "SINH", "DIA", "SU", "GDKT-PL", "TIN", "CNCN", "CNNN", "NN"];
    const TIEUDE_COT = ["THỜI GIAN NỘP", "MÃ HS", "HỌ VÀ TÊN", "LỚP", "MÃ ĐỀ", "ĐIỂM TRẮC NGHIỆM", "CHI TIẾT"];
    let files = dataFolder.getFilesByName(tenFileDiem), ss;
    if (files.hasNext()) { 
        ss = SpreadsheetApp.open(files.next());
    } else {
        ss = SpreadsheetApp.create(tenFileDiem);
        DriveApp.getFileById(ss.getId()).moveTo(dataFolder); 
        let firstSheet = ss.getSheets()[0]; firstSheet.setName(TENSHEET_MAP[THUTU_MON[0]]);
        firstSheet.appendRow(TIEUDE_COT);
        firstSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e65100").setFontColor("white"); firstSheet.setFrozenRows(1);
        for (let i = 1; i < THUTU_MON.length; i++) {
            let newSheet = ss.insertSheet(TENSHEET_MAP[THUTU_MON[i]]);
            newSheet.appendRow(TIEUDE_COT); newSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e65100").setFontColor("white"); newSheet.setFrozenRows(1);
        }
    }
    
    let monThi = duLieu.mon || "TOAN"; 
    let sheetCanLuu = ss.getSheetByName(TENSHEET_MAP[monThi]);
    if (!sheetCanLuu) {
        sheetCanLuu = ss.insertSheet(TENSHEET_MAP[monThi]);
        sheetCanLuu.appendRow(TIEUDE_COT); sheetCanLuu.getRange("A1:G1").setFontWeight("bold").setBackground("#e65100").setFontColor("white");
        sheetCanLuu.setFrozenRows(1);
    }

    let thoiGian = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    sheetCanLuu.appendRow([ thoiGian, duLieu.maHS, duLieu.tenHS, duLieu.lop, duLieu.maDe, duLieu.diemTN, duLieu.chiTiet ]);
    return {success: true};
  } catch (e) { return {success: false, msg: e.toString()}; }
}

function setupDatabase3Khoi() {
  let ss = SpreadsheetApp.create("NTTPRO_2026_HS_3KHOI");
  let khois = ["KHOI_10", "KHOI_11", "KHOI_12"];
  let headers = ["Mã HS", "Họ và tên HS", "Lớp", "Mật khẩu", "Made_TOAN", "Made_LY", "Made_HOA", "Made_SINH", "Made_DIA", "Made_SU", "Made_GDKT-PL", "Made_TIN", "Made_CNCN", "Made_CNNN", "Made_NN"];
  let sheet1 = ss.getSheets()[0]; 
  sheet1.setName(khois[0]); 
  sheet1.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
  
  for(let i = 1; i < khois.length; i++) { 
    let sheet = ss.insertSheet(khois[i]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white"); 
  }
  
  let folders = DriveApp.getFoldersByName("DU_AN_NTTPRO_2026");
  if (folders.hasNext()) {
    DriveApp.getFileById(ss.getId()).moveTo(folders.next());
  }
  
  PropertiesService.getScriptProperties().setProperty('DB_HS_3KHOI', ss.getId());
}

function luuApiKey(newKey) {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', newKey.trim());
  return "✅ Đã cập nhật mã API Key mới thành công!";
}

function layApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function luuHinhNenHeThong(base64Data, fileName) {
  try {
    var base64String = base64Data.split(',')[1] || base64Data;
    var blob = Utilities.newBlob(Utilities.base64Decode(base64String), "image/jpeg", fileName);
    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), "DU_AN_NTTPRO_2026");
    var bgFolder = getOrCreateFolder(rootFolder, "HINH_NEN");
    
    var oldFiles = bgFolder.getFiles();
    while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
    
    var file = bgFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var url = "https://drive.google.com/uc?export=view&id=" + file.getId();
    PropertiesService.getScriptProperties().setProperty('HINH_NEN_URL', url);
    return "✅ Đã lưu hình nền thành công!";
  } catch(e) { return "❌ Lỗi: " + e.toString(); }
}

// BỔ SUNG: HÀM QUẢN LÝ TIẾN ĐỘ HỌC SINH
function layDanhSachHocSinhThi(khoi, lop, monKey, maDe) {
  try {
    let dbId = PropertiesService.getScriptProperties().getProperty('DB_HS_3KHOI');
    if(!dbId) return { success: false, message: "Chưa cấu hình Database học sinh!" };

    let ssHS = SpreadsheetApp.openById(dbId);
    let sheetHS = ssHS.getSheetByName(khoi);
    if(!sheetHS) return { success: false, message: "Không tìm thấy dữ liệu khối " + khoi };

    let dataHS = sheetHS.getDataRange().getValues();
    let danhSachHocSinhLop = [];
    let colMon = MON_COLS[monKey];

    let cleanLop = lop.trim().toUpperCase();
    let cleanMaDe = maDe.trim().toUpperCase();

    for(let i = 1; i < dataHS.length; i++) {
      let row = dataHS[i];
      if(row[2] && row[2].toString().trim().toUpperCase() === cleanLop) {
         danhSachHocSinhLop.push({
           maHS: row[0].toString().trim(),
           tenHS: row[1].toString().trim(),
           lop: row[2].toString().trim(),
           maDeDaGiao: row[colMon] ? row[colMon].toString().trim().toUpperCase() : ""
         });
      }
    }

    if(danhSachHocSinhLop.length === 0) return { success: false, message: "Không có học sinh nào ở lớp " + cleanLop };

    // Tìm điểm trong file Diem
    let tenThuMucGoc = "DU_AN_NTTPRO_2026";
    let tenThuMucDiem = "Data_Diem";
    let tenFileDiem = "Diem_" + cleanLop;
    let rootFolders = DriveApp.getFoldersByName(tenThuMucGoc);

    let diemMap = {};
    if (rootFolders.hasNext()) {
       let dataFolders = rootFolders.next().getFoldersByName(tenThuMucDiem);
       if (dataFolders.hasNext()) {
          let files = dataFolders.next().getFilesByName(tenFileDiem);
          if (files.hasNext()) {
             let ssDiem = SpreadsheetApp.open(files.next());
             let sheetDiem = ssDiem.getSheetByName(TENSHEET_MAP[monKey]);
             if(sheetDiem) {
                let dataDiem = sheetDiem.getDataRange().getValues();
                for(let j = 1; j < dataDiem.length; j++) {
                   let mHS = dataDiem[j][1] ? dataDiem[j][1].toString().trim() : "";
                   let mDe = dataDiem[j][4] ? dataDiem[j][4].toString().trim().toUpperCase() : "";
                   let d = dataDiem[j][5] !== "" ? dataDiem[j][5] : "";
                   if (mDe === cleanMaDe) {
                      diemMap[mHS] = d;
                   }
                }
             }
          }
       }
    }

    // Kết hợp dữ liệu
    let ketQua = [];
    for(let i=0; i<danhSachHocSinhLop.length; i++) {
       let hs = danhSachHocSinhLop[i];
       let trangThai = "Chưa nộp bài";
       let diem = "";

       if(diemMap[hs.maHS] !== undefined) {
          trangThai = "Đã nộp bài";
          diem = diemMap[hs.maHS];
       } else {
          if(hs.maDeDaGiao === cleanMaDe) {
             trangThai = "Chưa nộp bài (Đang làm)";
          } else {
             trangThai = "Chưa được giao đề này";
          }
       }

       ketQua.push({
         stt: i + 1,
         ten: hs.tenHS,
         lop: hs.lop,
         trangThai: trangThai,
         diem: diem
       });
    }
    return { success: true, data: ketQua };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}