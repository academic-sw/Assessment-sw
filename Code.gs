/**
 * ==========================================================================
 * ระบบสารสนเทศงานวัดและประเมินผล โรงเรียนเสาไห้ "วิมลวิทยานุกูล"
 * Google Apps Script Backend (Google Sheets DB + Google Drive Upload)
 * Google Drive Folder ID: 1ehgB_CLugTr96TwYLTbSISYyO4K5rGob
 * ==========================================================================
 */

// โฟลเดอร์สำหรับบันทึกไฟล์รูปภาพและเอกสารแนบจากเว็บแอป
const FOLDER_ID = "1ehgB_CLugTr96TwYLTbSISYyO4K5rGob";

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAll";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let result = {};
  
  try {
    if (action === "getAll" || action === "getNews") {
      result.news = getSheetData(ss, "News");
    }
    if (action === "getAll" || action === "getSystems") {
      result.systems = getSheetData(ss, "Systems");
    }
    if (action === "getAll" || action === "getTheme") {
      result.theme = getSheetTheme(ss);
    }
    result.status = "success";
  } catch (err) {
    result.status = "error";
    result.message = err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = { status: "success" };
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "saveNews") {
      const newsList = data.news || [];
      // อัปโหลดไฟล์และรูปภาพเข้า Google Drive
      const processedNews = newsList.map(item => processNewsDriveUploads(item));
      saveSheetData(ss, "News", processedNews);
      result.news = processedNews;
    } else if (action === "saveSystems") {
      saveSheetData(ss, "Systems", data.systems || []);
      result.systems = data.systems;
    } else if (action === "saveTheme") {
      let theme = data.theme || {};
      if (theme.logoUrl && theme.logoUrl.indexOf("data:") === 0) {
        theme.logoUrl = uploadBase64ToDrive(theme.logoUrl, "school_logo_" + Date.now());
      }
      saveSheetTheme(ss, theme);
      result.theme = theme;
    }
  } catch (err) {
    result.status = "error";
    result.message = err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ฟังก์ชันบันทึกไฟล์ Base64 ลงใน Google Drive Folder (ID: 1ehgB_CLugTr96TwYLTbSISYyO4K5rGob)
 */
function uploadBase64ToDrive(base64Data, filenameStr) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const parts = base64Data.split(",");
    const meta = parts[0];
    const rawData = parts[1];
    
    let contentType = "image/png";
    const mimeMatch = meta.match(/data:(.*?);/);
    if (mimeMatch && mimeMatch[1]) {
      contentType = mimeMatch[1];
    }
    
    let ext = "png";
    if (contentType.indexOf("jpeg") !== -1 || contentType.indexOf("jpg") !== -1) ext = "jpg";
    if (contentType.indexOf("pdf") !== -1) ext = "pdf";
    if (contentType.indexOf("svg") !== -1) ext = "svg";
    
    const decoded = Utilities.base64Decode(rawData);
    const blob = Utilities.newBlob(decoded, contentType, filenameStr + "." + ext);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // ลิงก์ตรงแสดงผล/ดาวน์โหลดไฟล์บน Google Drive
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    Logger.log("Drive Upload Error: " + e.toString());
    return base64Data;
  }
}

function processNewsDriveUploads(newsItem) {
  // รูปภาพประกอบข่าว
  if (newsItem.images && newsItem.images.length > 0) {
    newsItem.images = newsItem.images.map((img, i) => {
      if (img && img.indexOf("data:") === 0) {
        return uploadBase64ToDrive(img, "news_img_" + newsItem.id + "_" + (i + 1));
      }
      return img;
    });
  }
  // ไฟล์เอกสารแนบ
  if (newsItem.files && newsItem.files.length > 0) {
    newsItem.files = newsItem.files.map((fileObj, i) => {
      if (fileObj && fileObj.url && fileObj.url.indexOf("data:") === 0) {
        const fileUrl = uploadBase64ToDrive(fileObj.url, "news_file_" + newsItem.id + "_" + (i + 1));
        return { name: fileObj.name || ("ไฟล์แนบ " + (i + 1)), url: fileUrl };
      }
      return fileObj;
    });
  }
  return newsItem;
}

function getSheetData(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    let row = rows[i];
    let item = {};
    headers.forEach((h, j) => {
      let val = row[j];
      if (typeof val === "string" && (val.indexOf("[") === 0 || val.indexOf("{") === 0)) {
        try { val = JSON.parse(val); } catch (e) {}
      }
      item[h] = val;
    });
    if (item.id) data.push(item);
  }
  return data;
}

function saveSheetData(ss, sheetName, items) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();
  if (!items || items.length === 0) return;
  
  const headers = Object.keys(items[0]);
  sheet.appendRow(headers);
  
  items.forEach(item => {
    let row = headers.map(h => {
      let val = item[h];
      if (typeof val === "object") val = JSON.stringify(val);
      return val;
    });
    sheet.appendRow(row);
  });
}

function getSheetTheme(ss) {
  let sheet = ss.getSheetByName("Theme");
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return null;
  let theme = {};
  for (let i = 1; i < rows.length; i++) {
    theme[rows[i][0]] = rows[i][1];
  }
  return theme;
}

function saveSheetTheme(ss, theme) {
  let sheet = ss.getSheetByName("Theme");
  if (!sheet) sheet = ss.insertSheet("Theme");
  sheet.clear();
  sheet.appendRow(["Key", "Value"]);
  Object.keys(theme).forEach(key => {
    sheet.appendRow([key, theme[key]]);
  });
}
