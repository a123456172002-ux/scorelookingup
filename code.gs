// ==================== 全域設定 ====================
const ADMIN_PASSWORD = "quentin0503"; // 管理員開關密碼
const MGMT_SPREADSHEET_ID = '15VEkkkt9G4FlS-xkOCYSZTvv3sJWeJoTUVSGvURWs-k'; // 管理系統試算表
const SCORE_SPREADSHEET_ID = '13nE4cNv-Fk6DUWoWsX6d76IZ5GsbV86kjaGbQfiaRA0'; // 成績查詢試算表
const SHEET_NAME = '總表';

// 1. 網頁進入點 (統一入口)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('成績管理與查詢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== 系統開關與權限邏輯 ====================

// 檢查系統目前是否開放 (預設為 true 開放)
function getSystemStatus() {
  var props = PropertiesService.getScriptProperties();
  var status = props.getProperty("SYSTEM_ENABLED");
  if (status === null) {
    props.setProperty("SYSTEM_ENABLED", "true");
    return { enabled: true };
  }
  return { enabled: status === "true" };
}

// 切換系統開關 (需密碼)
function toggleSystemStatus(password, targetStatus) {
  if (password !== ADMIN_PASSWORD) {
    return { success: false, message: "❌ 密碼錯誤，無法變更狀態！" };
  }
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SYSTEM_ENABLED", targetStatus ? "true" : "false");
  
  var stateText = targetStatus ? "【開放使用】" : "【關閉維護中】";
  return { success: true, message: "✅ 系統狀態已成功切換為：" + stateText, enabled: targetStatus };
}

// 內部檢查：未開放時阻擋後端查詢
function isSystemOpen() {
  var props = PropertiesService.getScriptProperties();
  var status = props.getProperty("SYSTEM_ENABLED");
  return status !== "false";
}

// ==================== 成績查詢邏輯 (學生成績) ====================

// 2. 獲取試算表內所有的班級清單
function getClassList() {
  if (!isSystemOpen()) return [];
  try {
    var ss = SpreadsheetApp.openById(SCORE_SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var header = data[0].map(h => h.toString().trim());
    var clsIdx = header.indexOf("班級");
    if (clsIdx === -1) clsIdx = header.indexOf("班");
    
    if (clsIdx === -1) return [];
    
    var classes = [];
    for (var i = 1; i < data.length; i++) {
      var cName = data[i][clsIdx].toString().trim();
      if (cName && !classes.includes(cName)) {
        classes.push(cName);
      }
    }
    return classes.sort();
  } catch(e) {
    return [];
  }
}

// 3. 抓取班級排名資料
function queryClassScores(className) {
  if (!isSystemOpen()) return "⚠️ 系統目前暫停開放，請稍後再試。";
  try {
    var ss = SpreadsheetApp.openById(SCORE_SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var header = data[0].map(h => h.toString().trim());

    var colMap = {};
    header.forEach((name, i) => colMap[name] = i);
    var getCol = list => { for (var n of list) { if (colMap[n] !== undefined) return colMap[n]; } return undefined; };

    var sidIdx = getCol(["學號", "SID"]);
    var clsIdx = getCol(["班級", "班"]);
    var nameIdx = getCol(["姓名"]);
    var seatIdx = getCol(["座號"]);
    var totalIdx = getCol(["總分"]);
    var avgIdx = getCol(["平均"]);
    var rankIdx = getCol(["班級排名", "排名", "名次", "班排名"]);
    
    var subjects = { "國文": ["國", "國文"], "英語": ["英", "英語", "英文"], "數學": ["數", "數學"], "自然": ["自", "自然"], "社會": ["社", "社會"] };

    var classStudents = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][clsIdx].toString().trim() === className) {
        var row = data[i];
        var rawSeat = row[seatIdx];
        var formattedSeat = (rawSeat !== "" && !isNaN(rawSeat)) ? ("0" + rawSeat).slice(-2) : rawSeat;
        
        var studentObj = {
          "座號": formattedSeat,
          "學號": row[sidIdx],
          "姓名": row[nameIdx],
          "scores": {},
          "總分": row[totalIdx] || "-",
          "平均": row[avgIdx] ? parseFloat(row[avgIdx]).toFixed(1) : "-",
          "排名": (rankIdx !== undefined && row[rankIdx] !== "") ? row[rankIdx].toString().trim() : "-" 
        };
        
        for (var key in subjects) {
          var idx = getCol(subjects[key]);
          var score = (idx !== undefined) ? row[idx] : "";
          studentObj.scores[key] = (score === "" || score === null) ? "-" : score;
        }
        classStudents.push(studentObj);
      }
    }
    
    classStudents.sort((a, b) => parseInt(a.座號) - parseInt(b.座號));
    return classStudents;
  } catch (e) {
    return "班級查詢異常：" + e.toString();
  }
}

// 4. 個人學號查詢邏輯
function queryScore(studentId) {
  if (!isSystemOpen()) return "⚠️ 系統目前暫停開放，請稍後再試。";
  try {
    var ss = SpreadsheetApp.openById(SCORE_SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var header = data[0].map(h => h.toString().trim());

    var colMap = {};
    header.forEach((name, i) => colMap[name] = i);

    var getCol = function(list) {
      for (var n of list) { if (colMap[n] !== undefined) return colMap[n]; }
      return undefined;
    };

    var sidIdx = getCol(["學號", "SID"]);
    var clsIdx = getCol(["班級", "班"]);
    
    var sRow = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][sidIdx].toString().trim() === studentId.toString().trim()) {
        sRow = data[i]; break;
      }
    }
    if (!sRow) return "查無此學號：" + studentId;

    var fullDisplayList = [
      "1140002", "1140006", "1140007", "1140008", "1140025", 
      "1140013", "1140016", "1140059", "1140042", "1140049", 
      "1140070", "1140072", "1140084", "1140096", "1140101", 
      "1140074", "1140099", "1140115"
    ];
    var isFullDisplay = fullDisplayList.includes(studentId.toString().trim());

    var rawSeat = sRow[getCol(["座號"])];
    var formattedSeat = (rawSeat !== "" && !isNaN(rawSeat)) ? ("0" + rawSeat).slice(-2) : rawSeat;

    var personal = {
      "info": { "班級": sRow[clsIdx], "座號": formattedSeat, "姓名": sRow[getCol(["姓名"])], "學號": sRow[sidIdx] },
      "scores": {}, 
      "stats": {}
    };

    var subjects = {
      "國文": ["國", "國文"], 
      "英語": ["英", "英語", "英文"], 
      "數學": ["數", "數學"],
      "自然": ["自", "自然"], 
      "社會": ["社", "社會"]
    };

    var targetClass = sRow[clsIdx];
    var classAvg = {};

    for (var key in subjects) {
      if (!isFullDisplay && key !== "英語") continue;

      var idx = getCol(subjects[key]);
      var score = (idx !== undefined) ? sRow[idx] : "";
      
      personal.scores[key] = (score === "" || score === null) ? "尚未上傳" : score;

      var avgValue = "N/A";
      if (key === "英語") {
        if (targetClass === "七年四班") {
          avgValue = sheet.getRange("O3").getValue();
        } else if (targetClass === "八年二班") {
          avgValue = sheet.getRange("P3").getValue();
        } else if (targetClass === "八年三班") {
          avgValue = sheet.getRange("Q3").getValue();
        } else {
          var classScores = data.filter((r, i) => i > 0 && r[clsIdx] === targetClass && r[idx] !== "" && !isNaN(r[idx]))
                               .map(r => parseFloat(r[idx]));
          avgValue = classScores.length ? (classScores.reduce((a, b) => a + b, 0) / classScores.length).toFixed(1) : "N/A";
        }
      } else if (idx !== undefined) {
        var cScores = data.filter((r, i) => i > 0 && r[clsIdx] === targetClass && r[idx] !== "" && !isNaN(r[idx]))
                         .map(r => parseFloat(r[idx]));
        avgValue = cScores.length ? (cScores.reduce((a, b) => a + b, 0) / cScores.length).toFixed(1) : "N/A";
      }
      
      classAvg[key] = (avgValue === "" || avgValue === null) ? "尚未上傳" : avgValue;
    }

    if (isFullDisplay) {
      personal.stats["總分"] = sRow[getCol(["總分"])] || "-";
      personal.stats["平均"] = parseFloat(sRow[getCol(["平均"])] || 0).toFixed(1);
    }

    return { personal: personal, classAvg: classAvg, isFullDisplay: isFullDisplay };
  } catch (e) {
    return "系統異常：" + e.toString();
  }
}

// ==================== 成績輸入/儲存邏輯 (老師管理) ====================

function getExistingData(query) {
  try {
    const ss = SpreadsheetApp.openById(MGMT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] && data[i][1].toString() === query.className.toString() &&
          data[i][2] === query.type &&
          data[i][3] === query.range &&
          data[i][4] === query.name) {
        let scores = {};
        for (let j = 1; j <= 26; j++) {
          let val = data[i][j + 4];
          scores[j] = (val === undefined || val === null) ? "" : val;
        }
        return { scores: scores };
      }
    }
  } catch(e) { return { error: e.toString() }; }
  return null;
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(MGMT_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    saveToSheet(sheet, formData);

    if (formData.name.includes("(原始)")) {
      let calcData = JSON.parse(JSON.stringify(formData));
      calcData.name = calcData.name.replace("(原始)", "(轉換)");
      
      const total = parseFloat(formData.totalWords);
      if (!isNaN(total) && total > 0) {
        for (let num in calcData.scores) {
          let raw = parseFloat(calcData.scores[num]);
          let weight = parseFloat(formData.levels[num] || 1.0);
          if (!isNaN(raw)) {
            calcData.scores[num] = Math.round((raw / total) * weight * 100);
          }
        }
        saveToSheet(sheet, calcData);
      }
    }
    return "✅ 原始與轉換成績已同步處理完成！";
  } catch (e) { return "❌ 錯誤：" + e.toString(); }
}

function saveToSheet(sheet, dataObj) {
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] && data[i][1].toString() === dataObj.className.toString() &&
        data[i][2] === dataObj.type &&
        data[i][3] === dataObj.range &&
        data[i][4] === dataObj.name) {
      targetRowIndex = i + 1;
      break;
    }
  }
  let rowOutput = [dataObj.date, dataObj.className, dataObj.type, dataObj.range, dataObj.name];
  for (let i = 1; i <= 26; i++) {
    let val = dataObj.scores[i];
    rowOutput.push((val === undefined || val === null) ? "" : val);
  }
  if (targetRowIndex > 0) {
    sheet.getRange(targetRowIndex, 1, 1, rowOutput.length).setValues([rowOutput]);
  } else {
    sheet.appendRow(rowOutput);
  }
}
