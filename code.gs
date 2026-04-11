/**
 * ============================================================
 * GOOGLE APPS SCRIPT — Smart Attendance System (Auto-Setup)
 * ระบบจะสร้าง Sheet และ Headers ให้โดยอัตโนมัติเมื่อรันครั้งแรก
 * ============================================================
 */

function doGet(e) {
  const action = e.parameter.action;
  checkAndInitSheets(); // ตรวจสอบและสร้าง Sheet อัตโนมัติทุกครั้งที่มีการเรียกใช้
  
  let result;
  try {
    switch (action) {
      case 'getConfig': result = getConfig(); break;
      case 'getKnownFaces': result = getKnownFaces(); break;
      case 'getAttendanceReport': result = getAttendanceReport(); break;
      case 'getSubjects': result = getSubjects(); break;
      case 'getStudents': result = getStudents(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  checkAndInitSheets(); // ตรวจสอบและสร้าง Sheet อัตโนมัติ
  
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ error: 'Invalid JSON body' });
  }

  const action = data.action;
  let result;

  try {
    switch (action) {
      case 'registerUser':
        result = registerUser(data.name, data.faceDescriptor, data.studentId, data.year);
        break;
      case 'logAttendance':
        result = logAttendance(data.name, data.subject, data.lat, data.lng);
        break;
      case 'saveConfig':
        result = saveConfig(data.lat, data.lng, data.radius);
        break;
      case 'addSubject':
        result = addSubject(data.code, data.name);
        break;
      case 'deleteItem':
        result = deleteItem(data.type, data.id);
        break;
      default:
        result = { error: 'Unknown POST action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return createJsonResponse(result);
}

/**
 * ฟังก์ชันตรวจสอบและสร้างแผ่นงาน (Sheets) อัตโนมัติ
 */
function checkAndInitSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet: Users (เก็บข้อมูลนักศึกษาและใบหน้า)
  if (!ss.getSheetByName('Users')) {
    const sheet = ss.insertSheet('Users');
    sheet.appendRow(['ID', 'Name', 'Year', 'FaceDescriptor', 'Timestamp']);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }

  // 2. Sheet: Attendance (เก็บประวัติการเช็คชื่อ)
  if (!ss.getSheetByName('Attendance')) {
    const sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Subject', 'Time', 'Date', 'Lat', 'Lng', 'MapLink']);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }

  // 3. Sheet: Subjects (เก็บรายวิชา)
  if (!ss.getSheetByName('Subjects')) {
    const sheet = ss.insertSheet('Subjects');
    sheet.appendRow(['Code', 'Name']);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }

  // 4. Sheet: Config (เก็บค่า GPS และระยะทาง)
  if (!ss.getSheetByName('Config')) {
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Parameter', 'Value']);
    sheet.appendRow(['Target Latitude', 13.7563]); // Default BKK
    sheet.appendRow(['Target Longitude', 100.5018]);
    sheet.appendRow(['Allowed Radius (KM)', 0.1]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ฟังก์ชันจัดการข้อมูล ---

function registerUser(name, descriptor, studentId, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  sheet.appendRow([studentId || '-', name, year || '-', JSON.stringify(descriptor), new Date()]);
  return { success: true, message: 'ลงทะเบียนสำเร็จ' };
}

function getStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ id: r[0], name: r[1], year: r[2] }));
}

function getKnownFaces() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ name: r[1], descriptor: JSON.parse(r[3]) }));
}

function addSubject(code, name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  sheet.appendRow([code, name]);
  return { success: true, message: 'เพิ่มรายวิชาสำเร็จ' };
}

function getSubjects() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ code: r[0], name: r[1] }));
}

function logAttendance(name, subject, lat, lng) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
  const now = new Date();
  const timeZone = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, timeZone, 'HH:mm:ss');
  const mapLink = lat ? `https://www.google.com/maps?q=${lat},${lng}` : '-';
  
  sheet.appendRow([name, subject || 'ทั่วไป', timeStr, dateStr, lat || '-', lng || '-', mapLink]);
  return { success: true, message: 'เช็คชื่อสำเร็จ' };
}

function getAttendanceReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).reverse().map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function saveConfig(lat, lng, radius) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  sheet.getRange(2, 2).setValue(lat);
  sheet.getRange(3, 2).setValue(lng);
  sheet.getRange(4, 2).setValue(radius);
  return { success: true, message: 'บันทึกการตั้งค่าแล้ว' };
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  return {
    lat: data[1][1],
    lng: data[2][1],
    radius: data[3][1]
  };
}

function deleteItem(type, id) {
  const sheetName = (type === 'student') ? 'Users' : 'Subjects';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบข้อมูลสำเร็จ' };
    }
  }
  return { error: 'ไม่พบข้อมูลที่ต้องการลบ' };
}
