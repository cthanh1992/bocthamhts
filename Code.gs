// Code.gs - BẢN FINAL ĐÃ SỬA LỖI
function doGet(e) {
  var page = e.parameter.page || "main"; 
  var template;
  if (page == "control") template = HtmlService.createTemplateFromFile('Controller');
  else if (page == "result") template = HtmlService.createTemplateFromFile('Result');
  else template = HtmlService.createTemplateFromFile('Main');
  
  return template.evaluate()
      .setTitle("Lucky Draw - Tất Niên")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- LẤY DANH SÁCH GIẢI THƯỞNG & THỐNG KÊ ---
function getPrizesWithStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPrizes = ss.getSheetByName("Prizes");
  var sheetWinners = ss.getSheetByName("Winners");
  
  if (!sheetPrizes || sheetPrizes.getLastRow() < 2) return [];
  
  // Lấy 6 cột: Tên(A), Nội dung(B), SL(C), Ảnh(D), Thứ tự(E), Thời gian(F)
  var prizes = sheetPrizes.getRange(2, 1, sheetPrizes.getLastRow()-1, 6).getValues();
  
  var winners = [];
  if (sheetWinners && sheetWinners.getLastRow() >= 2) {
    var winnerData = sheetWinners.getRange(2, 4, sheetWinners.getLastRow()-1, 1).getValues();
    winners = winnerData.flat().map(String);
  }
  
  return prizes.map(function(p) {
    var prizeName = String(p[0]);
    var totalQty = Number(p[2]);
    var usedQty = winners.filter(function(w){ return w === prizeName }).length;
    var remainQty = totalQty - usedQty;
    if(remainQty < 0) remainQty = 0;
    
    // Xử lý thời gian: Mặc định 10s nếu để trống
    var durationSec = p[5];
    if (!durationSec || isNaN(durationSec)) durationSec = 10; 

    return {
      name: prizeName,
      content: p[1],
      total: totalQty,
      remain: remainQty,
      img: p[3],
      duration: Number(durationSec)
    };
  });
}

// --- LẤY DANH SÁCH NHÂN VIÊN ---
function getEmployees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetData = ss.getSheetByName("Data");
    var sheetWinners = ss.getSheetByName("Winners");

    if (!sheetData) return [["ERROR", "Không tìm thấy sheet 'Data'"]];
    if (sheetData.getLastRow() < 2) return [["ERROR", "Sheet Data chưa có dữ liệu"]];

    var rawData = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, 4).getValues();

    var winners = [];
    if (sheetWinners && sheetWinners.getLastRow() >= 2) {
      winners = sheetWinners.getRange(2, 2, sheetWinners.getLastRow() - 1, 1).getValues().flat().map(function(s){ return String(s).trim() });
    }

    var cleanList = [];
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      var id = String(row[0]).trim();
      if (id === "" || winners.indexOf(id) !== -1) continue;
      cleanList.push([id, row[1], row[2], row[3]]);
    }
    
    if (cleanList.length === 0) return [["WARNING", "Đã hết danh sách nhân viên!"]];

    return cleanList;
  } catch (e) { return [["ERROR", e.toString()]]; }
}

// --- CÁC HÀM HỖ TRỢ KHÁC ---
function triggerSpin() { CacheService.getScriptCache().put("SPIN_STATUS", "RUN", 60); }
function checkStatus() { return CacheService.getScriptCache().get("SPIN_STATUS"); }
function ackSpin() { CacheService.getScriptCache().remove("SPIN_STATUS"); }

function saveWinner(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Winners");
  sheet.appendRow([new Date(), data.id, data.name, data.prizeName, data.prizeContent]);
}

// Code.gs
function getWinnersList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWinners = ss.getSheetByName("Winners"); // Cột: Time, ID, Name, PrizeName, PrizeContent
  var sheetPrizes = ss.getSheetByName("Prizes");
  var sheetData = ss.getSheetByName("Data"); // Cột: ID, Name, Dept, ImgID
  
  if (!sheetWinners || sheetWinners.getLastRow() < 2) return [];

  // 1. Tạo Map giải thưởng để lấy Order (sắp xếp)
  var prizeMap = {};
  if (sheetPrizes && sheetPrizes.getLastRow() >= 2) {
    var prizeData = sheetPrizes.getRange(2, 1, sheetPrizes.getLastRow() - 1, 5).getValues();
    prizeData.forEach(function(row) {
      // Key: Tên giải -> Value: Order
      var pOrder = row[4]; 
      if (pOrder === "" || pOrder === null) pOrder = 999; 
      prizeMap[String(row[0])] = Number(pOrder);
    });
  }

  // 2. Tạo Map nhân viên để lấy Phòng ban & Ảnh
  var empMap = {};
  if (sheetData && sheetData.getLastRow() >= 2) {
     var empData = sheetData.getRange(2, 1, sheetData.getLastRow()-1, 4).getValues(); // ID, Name, Dept, Img
     empData.forEach(function(row){
        var id = String(row[0]).trim();
        empMap[id] = { dept: row[2], img: row[3] };
     });
  }

  // 3. Ghép dữ liệu
  var winnersData = sheetWinners.getRange(2, 1, sheetWinners.getLastRow() - 1, 5).getValues();
  
  return winnersData.map(function(row){
    // row: [Time, ID, Name, PrizeName, PrizeContent]
    var wId = String(row[1]).trim();
    var pName = String(row[3]);
    
    // Format Time
    var timeStr = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "HH:mm:ss");
    
    // Lấy thông tin bổ sung
    var pOrder = prizeMap[pName] ? prizeMap[pName] : 999;
    var empInfo = empMap[wId] || { dept: "N/A", img: "" };
    
    // Output: [Time, ID, Name, PrizeName, PrizeContent, PrizeOrder, EmpDept, EmpImg]
    return [timeStr, wId, row[2], pName, row[4], pOrder, empInfo.dept, empInfo.img];
  });
}

function getImageData(id) {
  try {
    if(!id) return {status: false};
    if (id.toString().indexOf("http") !== -1) {
       var match = id.match(/[-\w]{25,}/);
       if (match) id = match[0];
    }
    var blob = DriveApp.getFileById(id).getBlob();
    return { status: true, mime: blob.getContentType(), bytes: Utilities.base64Encode(blob.getBytes()) };
  } catch (e) { return { status: false, error: e.toString() }; }
}
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
