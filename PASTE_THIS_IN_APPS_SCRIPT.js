/**
 * Google Apps Script for AIM Students Webapp Integration
 * 
 * Instructions:
 * 1. Open your Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Clear any default code and paste this script in.
 * 4. Make sure your Sheet has these header row titles in Row 1:
 *    A: Timestamp
 *    B: Grade
 *    C: Activity Name
 *    D: Student Name
 *    E: School Name
 *    F: Work Link
 * 
 * 5. Click Save.
 * 6. Click Deploy > New deployment. Select "Web app".
 * 7. Set "Execute as" to "Me" and "Who has access" to "Anyone". Click Deploy.
 * 8. Copy the Web App Executable URL and paste it inside the GOOGLE_SCRIPT_URL variable in app.js.
 */

function doGet(e) {
  var lock = LockService.getScriptLock();
  try {
    // Acquire a public lock that blocks up to 30 seconds for concurrent writes
    lock.waitLock(30000);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var action = e.parameter.action;
    
    if (action === "checkDuplicateLink") {
      var checkLink = (e.parameter.workLink || "").trim().toLowerCase();
      var data = sheet.getDataRange().getValues();
      var isDuplicate = false;
      for (var i = 1; i < data.length; i++) {
        var existingLink = (data[i][5] || "").toString().trim().toLowerCase();
        if (checkLink && existingLink === checkLink) {
          isDuplicate = true;
          break;
        }
      }
      return makeJSONResponse({ success: true, isDuplicate: isDuplicate }, e);
    }
    
    if (action === "submitStudentWork") {
      var grade = e.parameter.grade || "";
      var activityName = e.parameter.activityName || "";
      var studentName = e.parameter.studentName || "";
      var schoolName = e.parameter.schoolName || "";
      var workLink = (e.parameter.workLink || "").trim();
      
      // Duplicate Link Check in Column F (Work Link)
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var existingLink = (data[i][5] || "").toString().trim().toLowerCase();
        if (workLink && existingLink === workLink.toLowerCase()) {
          return makeJSONResponse({ success: false, isDuplicate: true, error: "Duplicate submission link" }, e);
        }
      }
      
      var newRow = [
        new Date(), // Timestamp
        grade.trim(),
        activityName.trim(),
        studentName.trim(),
        schoolName.trim(),
        workLink
      ];
      
      sheet.appendRow(newRow);
      return makeJSONResponse({ success: true }, e);
    }
    
    return makeJSONResponse({ success: false, error: "Invalid action" }, e);
  } catch (err) {
    return makeJSONResponse({ success: false, error: err.toString() }, e);
  } finally {
    // Always release lock safely
    lock.releaseLock();
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Acquire a public lock that blocks up to 30 seconds for concurrent writes
    lock.waitLock(30000);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var grade = e.parameter.grade;
    var activityName = e.parameter.activityName;
    var studentName = e.parameter.studentName;
    var schoolName = e.parameter.schoolName;
    var workLink = e.parameter.workLink;
    
    // Fallback if data is sent in JSON request body
    if (!grade && e.postData && e.postData.contents) {
      try {
        var data = JSON.parse(e.postData.contents);
        grade = data.grade;
        activityName = data.activityName;
        studentName = data.studentName;
        schoolName = data.schoolName;
        workLink = data.workLink;
      } catch (parseErr) {
        // Ignore parse error and keep url query param values
      }
    }
    
    var newRow = [
      new Date(), // Timestamp
      (grade || "").toString().trim(),
      (activityName || "").toString().trim(),
      (studentName || "").toString().trim(),
      (schoolName || "").toString().trim(),
      (workLink || "").toString().trim()
    ];
    
    sheet.appendRow(newRow);
    return makeJSONResponse({ success: true }, e);
  } catch (err) {
    return makeJSONResponse({ success: false, error: err.toString() }, e);
  } finally {
    // Always release lock safely
    lock.releaseLock();
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("");
}

function makeJSONResponse(object, e) {
  var prefix = e && e.parameter && e.parameter.prefix;
  var jsonString = JSON.stringify(object);
  
  var response;
  if (prefix) {
    response = ContentService.createTextOutput(prefix + '(' + jsonString + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    response = ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
  }
  return response;
}
