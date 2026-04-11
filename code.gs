/**
 * GOOGLE APPS SCRIPT - Backend v3.2 (รองรับข้อมูลนักศึกษาครบถ้วน)
 * ระบบจะจัดเก็บ ID, Name, Year แยกคอลัมน์ให้อัตโนมัติ
 */

function doGet(e) {
  const action = e.parameter.action;
  checkAndInitSheets();
  let result;

  try {
    switch (action) {
      case 'getConfig': result = getConfig(); break;
      case 'getKnownFaces': result = getKnownFaces(); break;
      case 'getAttendanceReport': result = getAttendanceReport(); break;
      case 'getSubjects': result = getSubjects(); break;
      case 'getStudents': result = getStudents(); break;
      default: result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  checkAndInitSheets();
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ error: 'Invalid JSON' });
  }

  const action = data.action;
  let result;

  try {
    switch (action) {
      case 'registerUser':
        // รับค่า studentId และ year เพิ่มจากเดิม
        result = registerUser(data.studentId, data.name, data.year, data.faceDescriptor);
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
        result = { error: 'Unknown POST action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return createJsonResponse(result);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function checkAndInitSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Users': ['StudentID', 'Name', 'Year', 'FaceDescriptor', 'Timestamp'],
    'Attendance': ['Name', 'Subject', 'Time', 'Date', 'Lat', 'Lng', 'MapLink'],
    'Subjects': ['Code', 'Name'],
    'Config': ['Parameter', 'Value']
  };

  for (let name in sheets) {
    if (!ss.getSheetByName(name)) {
      let sh = ss.insertSheet(name);
      sh.appendRow(sheets[name]);
      sh.getRange("1:1").setFontWeight("bold").setBackground("#f1f5f9");
      if (name === 'Config') {
        sh.appendRow(['Target Latitude', 13.7563]);
        sh.appendRow(['Target Longitude', 100.5018]);
        sh.appendRow(['Allowed Radius (KM)', 0.1]);
      }
    }
  }
}

function registerUser(id, name, year, descriptor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  // บันทึกแยกคอลัมน์ ID, Name, Year, Descriptor
  sheet.appendRow([id, name, year, JSON.stringify(descriptor), new Date()]);
  return { success: true, message: 'ลงทะเบียนรหัส ' + id + ' สำเร็จ' };
}

function logAttendance(name, subject, lat, lng) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance');
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  const mapLink = lat ? "https://www.google.com/maps?q=" + lat + "," + lng : "-";
  sheet.appendRow([name, subject || 'ทั่วไป', timeStr, dateStr, lat || '-', lng || '-', mapLink]);
  return { success: true, message: 'บันทึกเวลาสำเร็จ' };
}

function getKnownFaces() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  // Descriptor อยู่คอลัมน์ที่ 4 (Index 3)
  return data.slice(1).map(r => ({ name: r[1], descriptor: JSON.parse(r[3]) }));
}

function getStudents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => ({ id: r[0], name: r[1], year: r[2] }));
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  return { lat: data[1][1], lng: data[2][1], radius: data[3][1] };
}

function saveConfig(lat, lng, radius) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  sheet.getRange(2, 2).setValue(lat);
  sheet.getRange(3, 2).setValue(lng);
  sheet.getRange(4, 2).setValue(radius);
  return { success: true, message: 'บันทึกตั้งค่าสำเร็จ' };
}

function addSubject(code, name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  sheet.appendRow([code, name]);
  return { success: true };
}

function getSubjects() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subjects');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => ({ code: r[0], name: r[1] }));
}

function deleteItem(type, id) {
  const sheetName = (type === 'student') ? 'Users' : 'Subjects';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'ไม่พบข้อมูล' };
}
