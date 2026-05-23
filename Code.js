/**
 * HỆ THỐNG NTTPRO 2026 - TÁC GIẢ: NGUYỄN THANH TÙNG
 * KIẾN TRÚC MỚI: JSON + GITHUB CLOUD
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
    tmp.hinhNen = PropertiesService.getScriptProperties().getProperty('HINH_NEN_URL') || ''; 
    return tmp.evaluate().setTitle('Phòng thi Online').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  } else if (mode === 'packager') {
    var tmp = HtmlService.createTemplateFromFile('Packager');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Trung Tâm Xử Lý Đề Thi').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    var tmp = HtmlService.createTemplateFromFile('Admin');
    tmp.appUrl = ScriptApp.getService().getUrl();
    return tmp.evaluate().setTitle('Hệ Thống Quản Trị').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// Hàm hỗ trợ AI chuyển thẳng sang định dạng JSON cho kiến trúc mới
function generateExamFromAI(base64Data, fileName, struct) {
  try {
    var GEMINI_API_KEY = layApiKey();
    if (!GEMINI_API_KEY) return '{"error": "Chưa cấu hình API Key"}';

    var base64String = base64Data.split(',')[1] || base64Data;
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY;
    
    // Yêu cầu AI xuất trực tiếp mã JSON chuẩn, không dùng HTML
    var prompt = "Bạn là chuyên gia số hóa đề thi. Đọc file đính kèm và xuất ra MỘT FILE JSON ĐÚNG CHUẨN duy nhất. Tuyệt đối không dùng Markdown bọc ngoài.\n";
    prompt += "Cấu trúc JSON bắt buộc:\n";
    prompt += "{\n";
    prompt += "  \"subject\": \"00\", (Đánh mã: 00=Toán, 01=Lý, 02=Hóa, 03=Sinh, 04=Văn, 05=Sử, 06=Địa, 07=Anh, 08=GDCD, 09=Tin, 10=CN)\n";
    prompt += "  \"title\": \"Tên đề thi\",\n";
    prompt += "  \"timeLimit\": 3000, (Thời gian tính bằng giây)\n";
    prompt += "  \"p1_points\": 0.25, \"p2_points\": [0.1, 0.25, 0.5, 1.0], \"p3_points\": 0.5,\n";
    prompt += "  \"part1\": [ { \"q\": \"Câu hỏi?\", \"options\": [\"A\", \"B\", \"C\", \"D\"], \"ans\": 0, \"explanation\": \"Lời giải\" } ],\n";
    prompt += "  \"part2\": [ { \"q\": \"Câu hỏi lớn?\", \"items\": [\"Ý 1\", \"Ý 2\", \"Ý 3\", \"Ý 4\"], \"ans\": [true, false, true, false], \"explanation\": \"Lời giải\" } ],\n";
    prompt += "  \"part3\": [ { \"q\": \"Câu hỏi?\", \"ans\": \"-1,5\", \"explanation\": \"Lời giải\" } ]\n";
    prompt += "}\n";
    prompt += "Lưu ý:\n";
    prompt += "- Mọi công thức Toán phải dùng LaTeX bọc trong $...$.\n";
    prompt += "- Hình ảnh để nguyên nhãn '[CẦN_CHÈN_ẢNH]' trong trường 'image'.\n";
    prompt += "- Nếu có đoạn văn dùng chung (Đọc hiểu tiếng Anh), copy dán trực tiếp đoạn văn đó vào đầu phần 'q' của TỪNG câu hỏi thuộc chùm đó.";

    var payload = { "contents": [{ "parts": [ {"text": prompt}, { "inlineData": { "mimeType": "application/pdf", "data": base64String } } ] }], "generationConfig": { "temperature": 0.1, "responseMimeType": "application/json" } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    if (json.error) return '{"error": "' + json.error.message + '"}';
    return json.candidates[0].content.parts[0].text;

  } catch (e) { return '{"error": "' + e.toString() + '"}'; }
}

function checkAdminLogin(user, pass) {
  if (user.toString().trim() === "vinhxuan" && pass.toString().trim() === "123456") return true;
  return false; // Rút gọn logic để tăng tốc xác thực nội bộ nếu dùng CSDL riêng
}

function layApiKey() { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''; }
function luuApiKey(newKey) { PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', newKey.trim()); return "✅ Đã lưu API Key!"; }

function luuDiemTuDong(duLieu) {
  try {
    let rootFolders = DriveApp.getFoldersByName("DU_AN_NTTPRO_2026");
    if (!rootFolders.hasNext()) return {success: false, msg: "Không tìm thấy thư mục gốc!"};
    let dataFolder = getOrCreateFolder(rootFolders.next(), "Data_Diem");
    
    let ss, files = dataFolder.getFilesByName("Diem_" + duLieu.lop);
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
        sheetCanLuu.getRange("A1:G1").setFontWeight("bold").setBackground("#e65100").setFontColor("white");
        sheetCanLuu.setFrozenRows(1);
    }

    sheetCanLuu.appendRow([ Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss"), duLieu.maHS, duLieu.tenHS, duLieu.lop, duLieu.maDe, duLieu.diemTN, duLieu.chiTiet ]);
    return {success: true};
  } catch (e) { return {success: false, msg: e.toString()}; }
}

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function luuHinhNenHeThong(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(',')[1] || base64Data), "image/jpeg", fileName);
    var bgFolder = getOrCreateFolder(getOrCreateFolder(DriveApp.getRootFolder(), "DU_AN_NTTPRO_2026"), "HINH_NEN");
    var oldFiles = bgFolder.getFiles(); while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
    var file = bgFolder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    PropertiesService.getScriptProperties().setProperty('HINH_NEN_URL', "https://drive.google.com/uc?export=view&id=" + file.getId());
    return "✅ Đã lưu hình nền thành công!";
  } catch(e) { return "❌ Lỗi: " + e.toString(); }
}

// HÀM MỞ CỔNG API ĐỂ NHẬN ĐIỂM TỪ GITHUB GỬI VỀ
function doPost(e) {
  try {
    // 1. Nhận gói dữ liệu JSON từ GitHub gửi tới
    let duLieu = JSON.parse(e.postData.contents);
    
    // 2. Chuyển gói dữ liệu này cho hàm lưu điểm xuất sắc mà thầy đã viết sẵn
    let ketQua = luuDiemTuDong(duLieu);
    
    // 3. Báo cáo lại cho GitHub là đã lưu thành công
    return ContentService.createTextOutput(JSON.stringify(ketQua))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Báo lỗi nếu có trục trặc
    return ContentService.createTextOutput(JSON.stringify({success: false, msg: error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}