/**
 * @license
 * Copyright 2024 Google LLC
 * This script is the backend for a Student Attendance System.
 * Version: 3.8 (Added summary report and scoring functions)
 */

// --- CONFIGURATION ---
const SPREADSHEET_ID = '1CX7ZiydEDvLtu9N9GvFdoR78gwbinDnCh1ZoHpMEeg0'; // <-- ใส่ ID ของ Google Sheet ของคุณที่นี่
const STUDENT_SHEET_NAME = 'Students';
const ATTENDANCE_SHEET_NAME = 'Attendance';
const SUBJECT_SHEET_NAME = 'Subjects';
const SCORE_SHEET_NAME = 'Scores'; // [NEW] For storing scores
const STUDENT_HEADERS = ['StudentID', 'Name', 'Class', 'ImageDataBase64'];
const ATTENDANCE_HEADERS = ['StudentID', 'Name', 'Class', 'Subject', 'Timestamp', 'Location', 'Status'];
const SUBJECT_HEADERS = ['SubjectName'];
const SCORE_HEADERS = ['StudentID', 'Class', 'Subject', 'Score']; // [NEW] Headers for the Scores sheet

// --- GLOBAL VARIABLES ---
let spreadsheet, studentSheet, attendanceSheet, subjectSheet, scoreSheet;

// --- Main API Endpoint ---
function doPost(e) {
  try {
    initializeApp();
    const payload = JSON.parse(e.postData.contents);
    const { action, data } = payload;
    
    let result;

    switch (action) {
      // Student Actions
      case 'getStudents': result = handleGetStudents(); break;
      case 'addStudent': result = handleAddStudent(data); break;
      case 'updateStudent': result = handleUpdateStudent(data); break;
      case 'deleteStudent': result = handleDeleteStudent(data); break;
      // Attendance Actions
      case 'recordAttendanceBulk': result = handleRecordAttendanceBulk(data); break;
      case 'getAttendance': result = handleGetAttendance(); break;
      // Subject Actions
      case 'getSubjects': result = handleGetSubjects(); break;
      // Report & Dashboard Actions
      case 'getDashboardSummary': result = getDashboardSummaryBySubject(data); break;
      case 'getAttendanceReport': result = handleGetAttendanceReport(data); break;
      case 'getTodaysAttendanceBySubject': result = getTodaysAttendanceBySubject(data); break;
      case 'getAttendanceSummaryReport': result = getAttendanceSummaryReport(data); break;
      case 'updateAttendanceScores': result = updateAttendanceScores(data); break;
      case 'sendLineNotification': result = handleSendLineNotification(data); break;

      default:
        result = { success: false, error: `Unknown action: ${action}` };
    }
    return createJsonResponse(result);

  } catch (error) {
    Logger.log(`CRITICAL ERROR in doPost: ${error.stack}`);
    return createJsonResponse({ success: false, error: `Server error: ${error.message}` });
  }
}

// --- ACTION HANDLERS ---

function handleGetStudents() {
  const data = studentSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, data: [] };
  const headers = data.shift();
  const students = data.map((row, index) => {
    let studentObj = {};
    headers.forEach((header, i) => {
        studentObj[header.toLowerCase()] = row[i];
    });
    studentObj.rowIndex = index + 2;
    return studentObj;
  });
  return { success: true, data: students };
}

function handleAddStudent(data) {
  const { id, name, 'class': studentClass, imageData } = data;
  studentSheet.appendRow([id, name, studentClass, imageData || '']);
  return { success: true, message: 'Student added successfully' };
}

function handleUpdateStudent(data) {
  const { rowIndex, id, name, 'class': studentClass, imageData } = data;
  if (imageData && imageData.startsWith('data:image')) {
    studentSheet.getRange(rowIndex, 1, 1, 4).setValues([[id, name, studentClass, imageData]]);
  } else {
    studentSheet.getRange(rowIndex, 1, 1, 3).setValues([[id, name, studentClass]]);
  }
  return { success: true, message: 'Student updated successfully' };
}

function handleDeleteStudent(data) {
  studentSheet.deleteRow(parseInt(data.rowIndex));
  return { success: true, message: 'Student deleted successfully.' };
}

function handleRecordAttendanceBulk(data) {
  const { records, subject, date } = data;
  if (!records || !subject || !date) {
    return { success: false, error: "Missing data for replacement." };
  }

  try {
    const allData = attendanceSheet.getDataRange().getValues();
    const headers = allData.shift(); 
    const targetDateStr = new Date(date).toDateString();

    const remainingData = allData.filter(row => {
      const studentSubject = row[3]; 
      const timestamp = new Date(row[4]); 
      const recordDateStr = timestamp.toDateString();
      
      return !(recordDateStr === targetDateStr && studentSubject === subject);
    });

    const newRows = records.map(rec => [rec.id, rec.name, rec.class, rec.subject, rec.timestamp, rec.location, rec.status]);
    const finalData = remainingData.concat(newRows);
    
    attendanceSheet.clearContents();
    attendanceSheet.appendRow(headers);
    if (finalData.length > 0) {
      attendanceSheet.getRange(2, 1, finalData.length, headers.length).setValues(finalData);
    }
    
    return { success: true, message: 'Attendance updated successfully.' };
  } catch (error) {
    Logger.log(`ERROR in handleRecordAttendanceBulk (replacement logic): ${error.stack}`);
    return { success: false, error: error.message };
  }
}

function handleGetAttendance() {
  const data = attendanceSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, data: [] };
  const headers = data.shift();
  const records = data.map(row => {
    const record = {};
    headers.forEach((header, i) => record[header.toLowerCase()] = row[i]);
    return record;
  });
  return { success: true, data: records };
}

function handleGetSubjects() {
  if (!subjectSheet || subjectSheet.getLastRow() < 2) {
    return { success: true, data: [] };
  }
  const data = subjectSheet.getRange(2, 1, subjectSheet.getLastRow() - 1, 1).getValues();
  const subjects = data.map(row => row[0]).filter(subject => subject.toString().trim() !== '');
  return { success: true, data: subjects };
}

// --- DASHBOARD & REPORTS ---

function getDashboardSummaryBySubject(filters) {
    try {
        const { date, period } = filters;
        const allAttendanceRecords = handleGetAttendance().data.map(rec => ({...rec, timestamp: new Date(rec.timestamp) }));
        const selectedDate = new Date(date);

        let periodFilteredAttendance;
        if (period === 'day') {
            periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp.toDateString() === selectedDate.toDateString());
        } else if (period === 'week') {
            const startOfWeek = new Date(selectedDate);
            startOfWeek.setDate(selectedDate.getDate() - selectedDate.getDay());
            startOfWeek.setHours(0,0,0,0);
            const endOfWeek = new Date(startOfWeek);
            endOfWeek.setDate(startOfWeek.getDate() + 6);
            endOfWeek.setHours(23,59,59,999);
            periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp >= startOfWeek && rec.timestamp <= endOfWeek);
        } else { // month
            periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp.getFullYear() === selectedDate.getFullYear() && rec.timestamp.getMonth() === selectedDate.getMonth());
        }

        const summaryBySubject = {};
        periodFilteredAttendance.forEach(rec => {
            if (!rec.subject) return;
            if (!summaryBySubject[rec.subject]) {
                summaryBySubject[rec.subject] = { present: 0, absentOrLeave: 0 };
            }
            if (rec.status === 'มาเรียน') {
                summaryBySubject[rec.subject].present++;
            } else if (rec.status === 'ขาด' || rec.status === 'ลา') {
                summaryBySubject[rec.subject].absentOrLeave++;
            }
        });

        const summaryArray = Object.keys(summaryBySubject).map(subjectName => ({
            subject: subjectName,
            ...summaryBySubject[subjectName]
        })).sort((a,b) => a.subject.localeCompare(b.subject));

        return { success: true, data: summaryArray };
    } catch(error) {
        Logger.log(`ERROR in getDashboardSummaryBySubject: ${error.stack}`);
        return { success: false, error: error.message };
    }
}

function handleGetAttendanceReport(filters) {
  try {
    const studentData = studentSheet.getDataRange().getValues();
    const studentHeaders = studentData.shift();
    const allStudents = studentData.map(row => {
        let obj = {};
        studentHeaders.forEach((h, i) => obj[h.toLowerCase()] = row[i]);
        return obj;
    });

    const attendanceData = attendanceSheet.getDataRange().getValues();
    const attendanceHeaders = attendanceData.shift();
    let allAttendanceRecords = attendanceData.map(row => {
      let obj = {};
      attendanceHeaders.forEach((h, i) => obj[h.toLowerCase()] = row[i]);
      obj.timestamp = new Date(obj.timestamp);
      return obj;
    });

    const { classFilter, periodFilter, dateFilter } = filters;
    const selectedDate = new Date(dateFilter);

    let periodFilteredAttendance;
    if (periodFilter === 'day') { periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp.toDateString() === selectedDate.toDateString()); }
    else if (periodFilter === 'week') { const startOfWeek = new Date(selectedDate); startOfWeek.setDate(selectedDate.getDate() - selectedDate.getDay()); startOfWeek.setHours(0,0,0,0); const endOfWeek = new Date(startOfWeek); endOfWeek.setDate(startOfWeek.getDate() + 6); endOfWeek.setHours(23,59,59,999); periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp >= startOfWeek && rec.timestamp <= endOfWeek); }
    else { periodFilteredAttendance = allAttendanceRecords.filter(rec => rec.timestamp.getFullYear() === selectedDate.getFullYear() && rec.timestamp.getMonth() === selectedDate.getMonth()); }

    const classesToProcess = classFilter === 'all' ? [...new Set(allStudents.map(s => s.class))].filter(Boolean) : [classFilter];
    const summaryByClassAndSubject = {};

    classesToProcess.forEach(className => {
        const studentsInClass = allStudents.filter(s => s.class === className);
        if (studentsInClass.length === 0) return;
        const attendanceForClass = periodFilteredAttendance.filter(rec => rec.class === className);
        const subjectsTaught = [...new Set(attendanceForClass.map(rec => rec.subject))].filter(Boolean);
        if (subjectsTaught.length === 0) return;

        subjectsTaught.forEach(subjectName => {
            const key = `${className} - ${subjectName}`;
            const summary = { class: className, subject: subjectName, totalStudents: studentsInClass.length, present: 0, absent: 0, leave: 0 };
            const attendanceForSubject = attendanceForClass.filter(rec => rec.subject === subjectName);
            const studentIdsWithRecords = new Set();

            attendanceForSubject.forEach(rec => {
                studentIdsWithRecords.add(rec.studentid);
                if (rec.status === 'มาเรียน') summary.present++;
                else if (rec.status === 'ลา') summary.leave++;
                else if (rec.status === 'ขาด') summary.absent++;
            });

            const absentWithNoRecord = studentsInClass.filter(s => !studentIdsWithRecords.has(s.studentid)).length;
            summary.absent += absentWithNoRecord;
            summaryByClassAndSubject[key] = summary;
        });
    });

    const summaryArray = Object.values(summaryByClassAndSubject).map(item => {
      item.totalPresent = item.present;
      item.attendancePercent = item.totalStudents > 0 ? ((item.totalPresent / item.totalStudents) * 100).toFixed(2) : "0.00";
      return item;
    }).sort((a,b) => a.class.localeCompare(b.class) || a.subject.localeCompare(b.subject));

    const absenteeLineData = periodFilteredAttendance
      .filter(r => r.status === 'ขาด' || r.status === 'ลา')
      .map(r => ({ date: r.timestamp.toLocaleDateString('th-TH'), class: r.class, subject: r.subject || 'N/A', id: r.studentid, name: r.name, status: r.status }))
      .sort((a,b) => {
          const classCompare = a.class.localeCompare(b.class);
          if (classCompare !== 0) return classCompare;
          return String(a.id).localeCompare(String(b.id), undefined, { numeric: true });
      });

    return { success: true, data: { summary: summaryArray, lineData: absenteeLineData }};
  } catch (error) {
    Logger.log(`ERROR in handleGetAttendanceReport: ${error.stack}`);
    return { success: false, error: error.message };
  }
}

function getTodaysAttendanceBySubject(data) {
  try {
    const { subject, date } = data;
    if (!subject || !date) {
      return { success: false, error: 'Subject and date are required.' };
    }
    
    const targetDate = new Date(date).toDateString();
    const attendanceData = handleGetAttendance().data;
    
    const records = attendanceData.filter(rec => {
      const recordDate = new Date(rec.timestamp).toDateString();
      return recordDate === targetDate && rec.subject === subject;
    });

    return { success: true, data: records };
  } catch (error) {
    Logger.log(`ERROR in getTodaysAttendanceBySubject: ${error.stack}`);
    return { success: false, error: error.message };
  }
}

function getAttendanceSummaryReport(filters) {
  try {
    const { classFilter, period, date } = filters;
    
    const allStudents = handleGetStudents().data;
    const studentCountByClass = {};
    allStudents.forEach(s => {
      studentCountByClass[s.class] = (studentCountByClass[s.class] || 0) + 1;
    });

    const allAttendance = handleGetAttendance().data;
    const selectedDate = new Date(date);
    let filteredAttendance;
    if (period === 'month') {
      filteredAttendance = allAttendance.filter(rec => {
        const recDate = new Date(rec.timestamp);
        return recDate.getFullYear() === selectedDate.getFullYear() && recDate.getMonth() === selectedDate.getMonth();
      });
    } else {
      filteredAttendance = allAttendance;
    }

    if (classFilter !== 'all') {
      filteredAttendance = filteredAttendance.filter(rec => rec.class === classFilter);
    }
    
    const summary = {};
    filteredAttendance.forEach(rec => {
      const key = `${rec.class}|${rec.subject}`;
      if (!summary[key]) {
        summary[key] = {
          class: rec.class,
          subject: rec.subject,
          totalStudents: studentCountByClass[rec.class] || 0,
          present: 0,
          absent: 0,
          leave: 0,
        };
      }
      if (rec.status === 'มาเรียน') summary[key].present++;
      else if (rec.status === 'ขาด') summary[key].absent++;
      else if (rec.status === 'ลา') summary[key].leave++;
    });

    const scoreData = scoreSheet.getDataRange().getValues();
    scoreData.shift();
    const scores = {};
    scoreData.forEach(row => {
        const key = `${row[0]}|${row[2]}`;
        scores[key] = row[3];
    });

    const summaryArray = Object.values(summary).sort((a, b) => a.class.localeCompare(b.class) || a.subject.localeCompare(b.subject));
    
    return { success: true, data: { summary: summaryArray, details: filteredAttendance, scores: scores, allStudents: allStudents } };

  } catch (e) {
    Logger.log(`ERROR in getAttendanceSummaryReport: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function updateAttendanceScores(data) {
  try {
    const scoresToUpdate = data;
    if (!scoresToUpdate || scoresToUpdate.length === 0) {
      return { success: false, error: 'No scores to update.' };
    }
    const existingData = scoreSheet.getDataRange().getValues();
    const existingScoresMap = new Map();
    existingData.forEach((row, index) => {
      if (index > 0) {
        const key = `${row[0]}|${row[2]}`;
        existingScoresMap.set(key, index + 1);
      }
    });

    scoresToUpdate.forEach(scoreInfo => {
      const studentId = scoreInfo[0];
      const subject = scoreInfo[2];
      const score = scoreInfo[3];
      const key = `${studentId}|${subject}`;
      
      if (existingScoresMap.has(key)) {
        const rowIndex = existingScoresMap.get(key);
        scoreSheet.getRange(rowIndex, 4).setValue(score);
      } else {
        scoreSheet.appendRow(scoreInfo);
      }
    });
    
    return { success: true, message: 'Scores updated successfully.' };
  } catch(e) {
    Logger.log(`ERROR in updateAttendanceScores: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

function handleSendLineNotification(lineData) {
  const CHANNEL_ACCESS_TOKEN = "YOUR_CHANNEL_ACCESS_TOKEN";
  const GROUP_ID = "YOUR_GROUP_ID";
  if (CHANNEL_ACCESS_TOKEN === "YOUR_CHANNEL_ACCESS_TOKEN" || GROUP_ID === "YOUR_GROUP_ID") { return { success: false, error: 'กรุณาตั้งค่า CHANNEL_ACCESS_TOKEN และ GROUP_ID ในไฟล์ Code.gs' }; }
  try {
    if (!lineData || lineData.length === 0) return { success: true, message: 'ไม่มีนักเรียนที่ขาดหรือลาในวันนี้' };
    let messageText = 'แจ้งเตือนนักเรียนขาด/ลา 📋'; let currentClass = '';
    let currentSubject = '';
    lineData.forEach(item => { if (item.class !== currentClass) { currentClass = item.class; messageText += `\n\n🏫 *ชั้นเรียน: ${currentClass}*`; currentSubject = ''; } if(item.subject !== currentSubject) { currentSubject = item.subject; messageText += `\n*วิชา: ${currentSubject}* (${item.date})`; } messageText += `\n  - ${item.id} ${item.name} (${item.status})`; });
    if (messageText.length > 4950) messageText = messageText.substring(0, 4950) + "\n...(ข้อมูลยาวเกินไป)...";
    const url = "https://api.line.me/v2/bot/message/push";
    const payload = { to: GROUP_ID, messages: [{ type: "text", text: messageText }] };
    const options = { method: "post", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN }, payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) return { success: true, message: 'Notification sent.' };
    else throw new Error(`LINE API error: ${response.getContentText()}`);
  } catch (error) { Logger.log(`CRITICAL ERROR in handleSendLineNotification: ${error.stack}`);
  return { success: false, error: error.message }; }
}

// --- HELPER FUNCTIONS ---
function initializeApp() {
  if (spreadsheet) return;
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    Logger.log(`Could not open spreadsheet with ID: ${SPREADSHEET_ID}.`);
    throw new Error("Failed to open spreadsheet. Please check SPREADSHEET_ID.");
  }
  studentSheet = setupSheet(STUDENT_SHEET_NAME, STUDENT_HEADERS);
  attendanceSheet = setupSheet(ATTENDANCE_SHEET_NAME, ATTENDANCE_HEADERS);
  subjectSheet = setupSheet(SUBJECT_SHEET_NAME, SUBJECT_HEADERS);
  scoreSheet = setupSheet(SCORE_SHEET_NAME, SCORE_HEADERS);
}

function setupSheet(sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function createJsonResponse(responseObject) {
  return ContentService.createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}