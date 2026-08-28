/**
 * google_script.js - Google Apps Script Backend for Annual Health Check Registration System
 * 
 * Instructions for User:
 * 1. Open your Google Sheet (https://docs.google.com/spreadsheets/d/1iQGQdCQRdTgU3F6LwJnWloeRR5mNzvbz4TMaPWaOpMo/edit)
 *    Note: If you are using a Copy of the sheet, open your Copy.
 * 2. Go to Extensions (ส่วนขยาย) > Apps Script
 * 3. Delete any existing code and paste this code.
 * 4. Save and click "Deploy" (การทำให้ใช้งานได้) > "New deployment" (การทำให้ใช้งานได้ใหม่)
 * 5. Choose Type: "Web app" (เว็บแอป)
 * 6. Set settings:
 *    - Execute as (เรียกใช้งานในฐานะ): "Me" (ฉัน - บัญชีอีเมลของคุณ)
 *    - Who has access (ผู้มีสิทธิ์เข้าถึง): "Anyone" (ทุกคน)
 * 7. Click Deploy, authorize the permissions, and copy the Web App URL (ending in /exec).
 * 8. Paste this Web App URL into the UI app.js settings.
 */

// Keep blank to default to the active spreadsheet copy that the script is running in.
var SPREADSHEET_ID = "";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (err) {
      console.warn("Could not open spreadsheet by ID: " + err.toString());
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Handle incoming API requests
 */
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch(err) {
    return createJsonResponse({ success: false, error: "Invalid JSON format" });
  }
  
  var action = payload.action;
  var args = payload.args || [];
  var result;
  
  try {
    if (action === "getEmployeeData") {
      result = getEmployeeData(args[0]);
    } else if (action === "getEmployeeAndRegistration") {
      result = getEmployeeAndRegistration(args[0]);
    } else if (action === "getConfigAndSlots") {
      result = getConfigAndSlots();
    } else if (action === "saveRegistration") {
      result = saveRegistration(args[0]);
    } else if (action === "getRegistrationByEmpId") {
      result = getRegistrationByEmpId(args[0]);
    } else if (action === "deleteRegistration") {
      result = deleteRegistration(args[0]);
    } else if (action === "initializeSheets") {
      result = initializeSheets();
    } else if (action === "getAdminDashboardData") {
      result = getAdminDashboardData();
    } else {
      throw new Error("Action not found: " + action);
    }
    return createJsonResponse({ success: true, data: result });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Annual Health Check Registration API is running. Send POST requests to interact.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fetch employee detail by Employee ID
 * Uses getDisplayValues() to preserve exact formatting (including leading zeros).
 */
function getEmployeeData(employeeId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Name");
  if (!sheet) {
    throw new Error("Sheet 'Name' not found. Please initialize sheets first.");
  }
  
  var idToFind = String(employeeId).trim();
  if (/^\d+$/.test(idToFind)) {
    idToFind = idToFind.padStart(6, '0'); // Pad to 6 digits as per prompt
  }

  // Performance Optimization: Check script cache first to support 2,000 concurrent users
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("emp_" + idToFind);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.warn("Cache read error in getEmployeeData: " + e.toString());
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  
  // Read headers only (1 row) to dynamically map column letter
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var colId = headers.indexOf("รหัสพนักงาน");
  if (colId === -1) colId = headers.indexOf("fempno");
  if (colId === -1) throw new Error("Header 'รหัสพนักงาน' or 'fempno' not found in Name sheet.");
  
  var colLetter = getColumnLetter(colId + 1);
  
  // Fast search using Google Sheets native TextFinder on the ID column range (milliseconds lookup)
  var searchRange = sheet.getRange(colLetter + "2:" + colLetter + lastRow);
  var cell = searchRange.createTextFinder(idToFind).matchEntireCell(true).findNext();
  if (!cell) return null;
  
  // Read ONLY the matching employee's row
  var rowIdx = cell.getRow();
  var row = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  
  var colName = headers.indexOf("ชื่อ");
  if (colName === -1) colName = headers.indexOf("fempnamet");
  
  var colLastName = headers.indexOf("นามสกุล");
  if (colLastName === -1) colLastName = headers.indexOf("fsurnamet");
  
  var colDept = headers.indexOf("แผนก");
  if (colDept === -1) colDept = headers.indexOf("fdeptcode");
  
  var colLoc = headers.indexOf("สถานที่");
  
  var colProg = -1;
  // 1. Try exact matches first
  colProg = headers.indexOf("โปรแกรมตรวจ");
  if (colProg === -1) colProg = headers.indexOf("โปรแกรมตรวจสุขภาพ");
  if (colProg === -1) colProg = headers.indexOf("โปรแกรม");
  if (colProg === -1) colProg = headers.indexOf("Program");
  
  // 2. Fallback to substring containing 'โปรแกรม'/'program'/'prog' but excluding 'ปัจจัยเสี่ยง'/'risk'
  if (colProg === -1) {
    for (var h = 0; h < headers.length; h++) {
      var headerLower = headers[h].toLowerCase();
      if (headerLower.indexOf("ปัจจัยเสี่ยง") === -1 && headerLower.indexOf("risk") === -1) {
        if (headerLower.indexOf("โปรแกรม") !== -1 || headerLower.indexOf("program") !== -1 || headerLower.indexOf("prog") !== -1) {
          colProg = h;
          break;
        }
      }
    }
  }
  
  var colAge = headers.indexOf("อายุ");
  if (colAge === -1) colAge = headers.indexOf("fbirth");
  
  var colRisk = headers.indexOf("โปรแกรมปัจจัยเสี่ยง");
  
  var colGender = headers.indexOf("เพศ");
  if (colGender === -1) colGender = headers.indexOf("gender");
  if (colGender === -1) colGender = headers.indexOf("sex");
  
  // Find pregnancy column dynamically by checking if header contains "ตั้งครรภ์"
  var colPreg = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h].indexOf("ตั้งครรภ์") !== -1) {
      colPreg = h;
      break;
    }
  }
  
  // Find checkup right column dynamically by checking if header contains "สิทธิ์"
  var colRight = -1;
  for (var h = 0; h < headers.length; h++) {
    if (headers[h].indexOf("สิทธิ์") !== -1) {
      colRight = h;
      break;
    }
  }
  
  var ageVal = 0;
  if (colAge !== -1) {
    var ageRaw = row[colAge];
    if (ageRaw instanceof Date) {
      ageVal = 2026 - ageRaw.getFullYear();
    } else if (ageRaw) {
      var ageStr = String(ageRaw).trim();
      var parsedAge = parseInt(ageStr, 10);
      if (!isNaN(parsedAge) && parsedAge > 0 && parsedAge < 120) {
        ageVal = parsedAge;
      } else {
        // Try parsing date string like "25-Mar-67"
        var dateParts = ageStr.split("-");
        if (dateParts.length === 3) {
          var yearPart = parseInt(dateParts[2], 10);
          if (!isNaN(yearPart)) {
            // Assume 1900s for two-digit years > current year
            var birthYear = yearPart < 100 ? (yearPart > 26 ? 1900 + yearPart : 2000 + yearPart) : yearPart;
            ageVal = 2026 - birthYear;
          }
        }
      }
    }
  }
  var nameVal = colName !== -1 ? String(row[colName]).trim() : "";
  var lastNameVal = colLastName !== -1 ? String(row[colLastName]).trim() : "";
  var deptVal = colDept !== -1 ? String(row[colDept]).trim() : "";
  var locVal = colLoc !== -1 ? String(row[colLoc]).trim() : "";
  var progVal = colProg !== -1 ? String(row[colProg]).trim() : "";
  var riskVal = colRisk !== -1 ? String(row[colRisk]).trim() : "";
  var pregVal = colPreg !== -1 ? String(row[colPreg]).trim().toLowerCase() : "";
  var rightVal = colRight !== -1 ? String(row[colRight]).trim() : "";
  var genderVal = colGender !== -1 ? String(row[colGender]).trim().toUpperCase() : "";
  
  var isPregnant = (pregVal === "yes" || pregVal === "y" || pregVal.indexOf("ตั้งครรภ์") !== -1 || pregVal === "จริง" || pregVal === "มี");
  var checkupRightVal = rightVal !== "" ? rightVal : "มีสิทธิ์";
  
  // Separate program group based on age and program name
  var programGroup = "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี";
  if (progVal.indexOf("MGR") !== -1 || progVal.toLowerCase().indexOf("mgr") !== -1) {
    programGroup = "โปรแกรม MGR";
  } else if (progVal.indexOf("35 ปีขึ้นไป") !== -1 || ageVal >= 35) {
    programGroup = "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป";
  }
  
  var result = {
    employeeId: idToFind,
    firstName: nameVal,
    lastName: lastNameVal,
    department: deptVal,
    defaultLocation: locVal,
    programName: progVal,
    age: ageVal,
    gender: genderVal,
    programGroup: programGroup,
    riskProgram: riskVal,
    isPregnant: isPregnant,
    checkupRight: checkupRightVal
  };

  // Cache employee detail for 5 minutes (300 seconds)
  try {
    var cache = CacheService.getScriptCache();
    cache.put("emp_" + idToFind, JSON.stringify(result), 300);
  } catch (e) {
    console.warn("Cache write error in getEmployeeData: " + e.toString());
  }
  
  return result;
}

/**
 * Combined API to get both employee details and registration status in a single round-trip.
 * Drastically reduces search time from 10+ seconds to under 2 seconds.
 */
function getEmployeeAndRegistration(employeeId) {
  var idToFind = String(employeeId).trim();
  if (/^\d+$/.test(idToFind)) {
    idToFind = idToFind.padStart(6, '0');
  }
  
  var cache = CacheService.getScriptCache();
  var empCached = null;
  var regCached = null;
  
  try {
    empCached = cache.get("emp_" + idToFind);
    regCached = cache.get("reg_" + idToFind);
  } catch (e) {
    console.warn("Cache read error in getEmployeeAndRegistration: " + e.toString());
  }
  
  var employee = null;
  if (empCached) {
    employee = JSON.parse(empCached);
  } else {
    employee = getEmployeeData(idToFind);
  }
  
  var registration = null;
  if (regCached) {
    registration = JSON.parse(regCached);
  } else {
    registration = getRegistrationByEmpId(idToFind);
    if (registration) {
      try {
        cache.put("reg_" + idToFind, JSON.stringify(registration), 120); // cache registration for 2 minutes
      } catch (e) {}
    }
  }
  
  return {
    employee: employee,
    registration: registration
  };
}

/**
 * Returns configuration settings (dates and slot caps) and calculated current booking counts.
 * Uses getDisplayValues() to prevent auto date object conversion of dates columns.
 */
function getConfigAndSlots() {
  // Performance Optimization: Check script cache first to support 2,000 concurrent users
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("config_and_slots");
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.warn("Cache read error in getConfigAndSlots: " + e.toString());
  }

  var ss = getSpreadsheet();
  
  var datesSheet = ss.getSheetByName("Config_Dates");
  var timesSheet = ss.getSheetByName("Config_TimeSlots");
  var regSheet = ss.getSheetByName("Registration");
  
  var dates = [];
  if (datesSheet) {
    var dData = datesSheet.getDataRange().getDisplayValues();
    var dHeaders = dData[0].map(function(h) { return String(h).trim(); });
    var colLoc = dHeaders.indexOf("สถานที่");
    var colDate = dHeaders.indexOf("วันที่ตรวจ");
    var colTeam = dHeaders.indexOf("ทีม");
    
    for (var i = 1; i < dData.length; i++) {
      dates.push({
        location: colLoc !== -1 ? String(dData[i][colLoc]).trim() : "",
        dateString: colDate !== -1 ? String(dData[i][colDate]).trim() : "",
        team: colTeam !== -1 ? String(dData[i][colTeam]).trim() : ""
      });
    }
  }
  
  var timeSlots = [];
  if (timesSheet) {
    var tData = timesSheet.getDataRange().getDisplayValues();
    var tHeaders = tData[0].map(function(h) { return String(h).trim(); });
    var colSlot = tHeaders.indexOf("รอบเวลา");
    var colLimit = tHeaders.indexOf("จำนวนจำกัด");
    
    for (var i = 1; i < tData.length; i++) {
      timeSlots.push({
        slotTime: colSlot !== -1 ? String(tData[i][colSlot]).trim() : "",
        limit: colLimit !== -1 ? parseInt(tData[i][colLimit], 10) : 50
      });
    }
  }
  
  // Count current registrations
  var registrationCounts = {};
  if (regSheet) {
    var rData = regSheet.getDataRange().getDisplayValues();
    if (rData.length > 1) {
      var rHeaders = rData[0].map(function(h) { return String(h).trim(); });
      var colLoc = rHeaders.indexOf("สถานที่");
      var colDate = rHeaders.indexOf("วันที่ตรวจ");
      var colTime = rHeaders.indexOf("เวลาที่ตรวจ");
      
      for (var i = 1; i < rData.length; i++) {
        var loc = colLoc !== -1 ? String(rData[i][colLoc]).trim() : "";
        var dateStr = colDate !== -1 ? String(rData[i][colDate]).trim() : "";
        var timeStr = colTime !== -1 ? String(rData[i][colTime]).trim() : "";
        
        var key = loc + "|" + dateStr + "|" + timeStr;
        registrationCounts[key] = (registrationCounts[key] || 0) + 1;
      }
    }
  }
  
  var result = {
    dates: dates,
    timeSlots: timeSlots,
    registrationCounts: registrationCounts
  };
  
  // Store in cache for 15 seconds (distributes load of 2000 users)
  try {
    var cache = CacheService.getScriptCache();
    cache.put("config_and_slots", JSON.stringify(result), 15);
  } catch (e) {
    console.warn("Cache write error in getConfigAndSlots: " + e.toString());
  }
  
  return result;
}

/**
 * Save user registration. Lock applied for race conditions.
 */
function saveRegistration(regData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    throw new Error("ระบบหนาแน่นชั่วคราว กรุณาลองใหม่อีกครั้ง (Lock timeout)");
  }
  
  try {
    var ss = getSpreadsheet();
    var regSheet = ss.getSheetByName("Registration");
    if (!regSheet) {
      regSheet = ss.insertSheet("Registration");
      regSheet.appendRow([
        "รหัสพนักงาน", "ชื่อ", "นามสกุล", "แผนก", "เบอร์โทรภายใน", 
        "กะทำงาน", "สถานที่", "วันที่ตรวจ", "เวลาที่ตรวจ", "รายการตรวจมะเร็งที่เลือก", "โปรแกรมปัจจัยเสี่ยง", "ตั้งครรภ์", "Timestamp"
      ]);
      regSheet.getRange("A1:M1").setFontWeight("bold").setBackground("#d9ead3");
    } else {
      // Dynamic migration: check if 'โปรแกรมปัจจัยเสี่ยง', 'ตั้งครรภ์', and 'การยินยอมใช้สิทธิ์ประกันสังคม' exist, if not, insert before Timestamp
      var rHeaders = regSheet.getDataRange().getDisplayValues()[0].map(function(h) { return String(h).trim(); });
      
      var colRiskIdx = rHeaders.indexOf("โปรแกรมปัจจัยเสี่ยง");
      if (colRiskIdx === -1) {
        var timeIdx = rHeaders.indexOf("Timestamp");
        if (timeIdx !== -1) {
          regSheet.insertColumnBefore(timeIdx + 1); // 1-indexed
          regSheet.getRange(1, timeIdx + 1).setValue("โปรแกรมปัจจัยเสี่ยง");
          regSheet.getRange(1, timeIdx + 1).setFontWeight("bold").setBackground("#d9ead3");
          rHeaders = regSheet.getDataRange().getDisplayValues()[0].map(function(h) { return String(h).trim(); });
        }
      }
      
      var colPregIdx = rHeaders.indexOf("ตั้งครรภ์");
      if (colPregIdx === -1) {
        var timeIdx = rHeaders.indexOf("Timestamp");
        if (timeIdx !== -1) {
          regSheet.insertColumnBefore(timeIdx + 1); // 1-indexed
          regSheet.getRange(1, timeIdx + 1).setValue("ตั้งครรภ์");
          regSheet.getRange(1, timeIdx + 1).setFontWeight("bold").setBackground("#d9ead3");
          rHeaders = regSheet.getDataRange().getDisplayValues()[0].map(function(h) { return String(h).trim(); });
        }
      }
      
      var colSsoIdx = rHeaders.indexOf("การยินยอมใช้สิทธิ์ประกันสังคม");
      if (colSsoIdx === -1) {
        var timeIdx = rHeaders.indexOf("Timestamp");
        if (timeIdx !== -1) {
          regSheet.insertColumnBefore(timeIdx + 1); // 1-indexed
          regSheet.getRange(1, timeIdx + 1).setValue("การยินยอมใช้สิทธิ์ประกันสังคม");
          regSheet.getRange(1, timeIdx + 1).setFontWeight("bold").setBackground("#d9ead3");
        }
      }
    }
    
    var employeeId = String(regData.employeeId).trim();
    if (/^\d+$/.test(employeeId)) {
      employeeId = employeeId.padStart(6, '0');
    }
    
    // Check if employee is eligible to register (is not blocked by HR)
    var empDetail = getEmployeeData(employeeId);
    if (empDetail && empDetail.checkupRight && empDetail.checkupRight.indexOf("ไม่มีสิทธิ์") !== -1) {
      throw new Error("ขออภัย ท่านยังไม่สามารถตรวจสุขภาพประจำปีนี้ได้ เนื่องจากเข้างานยังไม่ครบ 6 เดือน");
    }
    
    // Check if employee is already registered. Overwrite existing record if found.
    var data = regSheet.getDataRange().getDisplayValues();
    var existingRowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      var rawId = String(data[i][0]).trim().replace(/^'/, '');
      if (/^\d+$/.test(rawId)) {
        rawId = rawId.padStart(6, '0');
      }
      if (rawId === employeeId) {
        existingRowIndex = i + 1; // 1-indexed
        break;
      }
    }
    
    // Check slot limit
    var configAndSlots = getConfigAndSlots();
    var key = regData.location + "|" + regData.dateString + "|" + regData.timeString;
    var currentCount = configAndSlots.registrationCounts[key] || 0;
    
    // If it's a new registration or changing slots, check if slot is full
    var isChangingSlot = true;
    if (existingRowIndex !== -1) {
      var oldLoc = String(data[existingRowIndex - 1][6]).trim();
      var oldDate = String(data[existingRowIndex - 1][7]).trim();
      var oldTime = String(data[existingRowIndex - 1][8]).trim();
      if (oldLoc === regData.location && oldDate === regData.dateString && oldTime === regData.timeString) {
        isChangingSlot = false;
      }
    }
    
    if (isChangingSlot) {
      var limit = 50;
      for (var s = 0; s < configAndSlots.timeSlots.length; s++) {
        if (configAndSlots.timeSlots[s].slotTime === regData.timeString) {
          limit = configAndSlots.timeSlots[s].limit;
          break;
        }
      }
      if (currentCount >= limit) {
        throw new Error("รอบเวลา " + regData.timeString + " ในวันที่ " + regData.dateString + " สำหรับสถานที่ " + regData.location + " เต็มแล้ว (จำกัด " + limit + " คน)");
      }
    }
    
    var timestampStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    
    var finalHeaders = regSheet.getDataRange().getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    var rowDataMap = {
      "รหัสพนักงาน": "'" + employeeId,
      "ชื่อ": regData.firstName,
      "นามสกุล": regData.lastName,
      "แผนก": regData.department,
      "เบอร์โทรภายใน": "'" + regData.phone,
      "กะทำงาน": regData.shift,
      "สถานที่": regData.location,
      "วันที่ตรวจ": regData.dateString,
      "เวลาที่ตรวจ": regData.timeString,
      "รายการตรวจมะเร็งที่เลือก": regData.cancerTest || "",
      "โปรแกรมปัจจัยเสี่ยง": regData.riskProgram || "",
      "ตั้งครรภ์": regData.isPregnant ? "Yes" : "",
      "การยินยอมใช้สิทธิ์ประกันสังคม": regData.ssoConsent || "",
      "Timestamp": timestampStr
    };
    
    var rowValues = [];
    for (var h = 0; h < finalHeaders.length; h++) {
      var headerName = finalHeaders[h];
      rowValues.push(rowDataMap[headerName] !== undefined ? rowDataMap[headerName] : "");
    }
    
    if (existingRowIndex !== -1) {
      var range = regSheet.getRange(existingRowIndex, 1, 1, rowValues.length);
      range.setValues([rowValues]);
    } else {
      regSheet.appendRow(rowValues);
    }
    
    // Write "ลงทะเบียนแล้ว" to Column G of the Name sheet
    try {
      updateNameRegistrationStatus(employeeId, "ลงทะเบียนแล้ว");
    } catch (e) {
      console.warn("Could not update name registration status:", e);
    }
    
    // Clear cache to keep counts and employee details up to date under high concurrency (2000 users)
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("config_and_slots");
      cache.remove("emp_" + employeeId);
      cache.remove("reg_" + employeeId);
    } catch (e) {
      console.warn("Cache eviction error in saveRegistration: " + e.toString());
    }
    
    return { success: true, isNew: existingRowIndex === -1 };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Find user registration record
 */
function getRegistrationByEmpId(employeeId) {
  var idToFind = String(employeeId).trim();
  if (/^\d+$/.test(idToFind)) {
    idToFind = idToFind.padStart(6, '0');
  }

  // Check cache first
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("reg_" + idToFind);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.warn("Cache read error in getRegistrationByEmpId: " + e.toString());
  }

  var ss = getSpreadsheet();
  var regSheet = ss.getSheetByName("Registration");
  if (!regSheet) return null;
  
  var lastRow = regSheet.getLastRow();
  if (lastRow <= 1) return null;
  
  // Read headers only (1 row) to dynamically map column letter
  var headers = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var colIdIdx = headers.indexOf("รหัสพนักงาน");
  if (colIdIdx === -1) return null;
  
  var colLetter = getColumnLetter(colIdIdx + 1);
  
  // Fast search using Google Sheets native TextFinder on the ID column range (milliseconds lookup)
  var searchRange = regSheet.getRange(colLetter + "2:" + colLetter + lastRow);
  var cell = searchRange.createTextFinder(idToFind).matchEntireCell(true).findNext();
  if (!cell) return null;
  
  // Read ONLY the matching registration row
  var rowIdx = cell.getRow();
  var row = regSheet.getRange(rowIdx, 1, 1, headers.length).getDisplayValues()[0];
  
  var colPhone = headers.indexOf("เบอร์โทรภายใน");
  var colShift = headers.indexOf("กะทำงาน");
  var colLoc = headers.indexOf("สถานที่");
  var colDate = headers.indexOf("วันที่ตรวจ");
  var colTime = headers.indexOf("เวลาที่ตรวจ");
  var colCancer = headers.indexOf("รายการตรวจมะเร็งที่เลือก");
  var colRisk = headers.indexOf("โปรแกรมปัจจัยเสี่ยง");
  var colPreg = headers.indexOf("ตั้งครรภ์");
  var colSso = headers.indexOf("การยินยอมใช้สิทธิ์ประกันสังคม");
  var colTimeCreated = headers.indexOf("Timestamp");
  
  var empDetail = getEmployeeData(idToFind) || {};
  
  var result = {
    employeeId: idToFind,
    firstName: empDetail.firstName || String(row[1]).trim(),
    lastName: empDetail.lastName || String(row[2]).trim(),
    department: empDetail.department || String(row[3]).trim(),
    programGroup: empDetail.programGroup || "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป",
    age: empDetail.age || 0,
    gender: empDetail.gender || "",
    phone: colPhone !== -1 ? String(row[colPhone]).trim().replace(/^'/, '') : "",
    shift: colShift !== -1 ? String(row[colShift]).trim() : "",
    location: colLoc !== -1 ? String(row[colLoc]).trim() : "",
    dateString: colDate !== -1 ? String(row[colDate]).trim() : "",
    timeString: colTime !== -1 ? String(row[colTime]).trim() : "",
    cancerTest: colCancer !== -1 ? String(row[colCancer]).trim() : "",
    riskProgram: empDetail.riskProgram || (colRisk !== -1 ? String(row[colRisk]).trim() : ""),
    isPregnant: colPreg !== -1 ? String(row[colPreg]).trim() === "Yes" : false,
    ssoConsent: colSso !== -1 ? String(row[colSso]).trim() : "",
    timestamp: colTimeCreated !== -1 ? (row[colTimeCreated] instanceof Date ? Utilities.formatDate(row[colTimeCreated], "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss") : String(row[colTimeCreated]).trim()) : ""
  };

  // Cache result for 2 minutes (120 seconds)
  try {
    var cache = CacheService.getScriptCache();
    cache.put("reg_" + idToFind, JSON.stringify(result), 120);
  } catch (e) {}

  return result;
}

/**
 * Delete a user registration. Lock applied.
 */
function deleteRegistration(employeeId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    throw new Error("ระบบหนาแน่นชั่วคราว กรุณาลองใหม่อีกครั้ง (Lock timeout)");
  }
  
  try {
    var ss = getSpreadsheet();
    var regSheet = ss.getSheetByName("Registration");
    if (!regSheet) return { success: false, error: "Registration sheet not found" };
    
    var idToFind = String(employeeId).trim();
    if (/^\d+$/.test(idToFind)) {
      idToFind = idToFind.padStart(6, '0');
    }
    
    var data = regSheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: false, error: "No registrations found" };
    
    for (var i = 1; i < data.length; i++) {
      var rawId = String(data[i][0]).trim().replace(/^'/, '');
      if (/^\d+$/.test(rawId)) {
        rawId = rawId.padStart(6, '0');
      }
      if (rawId === idToFind) {
        regSheet.deleteRow(i + 1); // 1-indexed, header is row 1
        
        // Clear "ลงทะเบียนแล้ว" in Column G of the Name sheet
        try {
          updateNameRegistrationStatus(idToFind, "");
        } catch (e) {
          console.warn("Could not clear name registration status:", e);
        }
        
        // Clear caches under high concurrency (2000 users)
        try {
          var cache = CacheService.getScriptCache();
          cache.remove("config_and_slots");
          cache.remove("emp_" + idToFind);
          cache.remove("reg_" + idToFind);
        } catch (e) {
          console.warn("Cache eviction error in deleteRegistration: " + e.toString());
        }
        
        return { success: true };
      }
    }
    return { success: false, error: "Registration record not found" };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Initialize all spreadsheet sheets, column structure and default configs
 */
function initializeSheets() {
  var ss = getSpreadsheet();
  
  // 1. Name Sheet
  var nameSheet = ss.getSheetByName("Name");
  if (!nameSheet) {
    nameSheet = ss.insertSheet("Name");
    nameSheet.appendRow(["รหัสพนักงาน", "ชื่อ", "นามสกุล", "แผนก", "สถานที่", "โปรแกรมตรวจ", "อายุ", "สิทธิ์การตรวจ"]);
    nameSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#c9daf8");
    
    // Add default values matching user's image exactly (with padded zeros)
    var demoEmployees = [
      ["'003049", "วิชัย", "สุขประเสริฐกุล", "OPT", "LPN1", "โปรแกรม MGR", 59, ""],
      ["'004148", "ประภาพร", "ศรีประดู่", "HRDS", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 57, ""],
      ["'004379", "ประคอง", "อ้อยงาม", "QM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 54, ""],
      ["'004766", "พวงเพชร", "มณีฉาย", "CADT", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 55, ""],
      ["'005933", "ระเบียบ", "ปาละรัตน์", "QM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 60, ""],
      ["'006078", "อดิเรก", "อ่อนพรม", "OP2S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 57, ""],
      ["'006125", "ยุทธนา", "สุยะนันทน์", "OP2S", "LPN2", "โปรแกรม MGR", 56, ""],
      ["'006585", "มะลิ", "ยอดสิงห์", "OP4", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 53, ""],
      ["'006665", "กุหลาบ", "ศรไชย", "OP4", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 51, ""],
      ["'006704", "เยี่ยมรัก", "โดยอาษา", "QM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 60, ""],
      ["'006764", "อัมพร", "มาลิสา", "OP2S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 54, ""],
      ["'007046", "ดาราวรรณ", "คำสกุล", "FIN", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 55, ""],
      ["'007113", "พรณภัทร", "สุธรรมแจ่ม", "CADT", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 53, ""],
      ["'007114", "สายสุนีย์", "เขียวงาม", "OP2S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 52, ""],
      ["'007131", "มัทณานันต์", "พันธุ์สมบัติ", "QM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 50, ""],
      ["'007236", "ไพพรรณณ์", "ป้อมรักษา", "QM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 51, ""],
      ["'007321", "ทวี", "นพพรพิทักษ์", "TECH", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 58, ""],
      ["'007705", "มณฑา", "จันทร์เสน", "TRF", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 53, ""],
      ["'007860", "วิไล", "แสวงศรี", "OP4", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 49, ""],
      ["'008790", "วาสนา", "แซ่มวงศ์", "OP2S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 59, ""],
      ["'009115", "ธวัชชัย", "ช่างทอง", "PMS", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 54, ""],
      ["'009597", "สังเวียน", "มีดี", "PM", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 54, ""],
      ["'009847", "ศศิกานต์", "มะโนวงศ์", "OP1", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 55, ""],
      ["'009892", "สมชาย", "ม่วงไหม", "OP2S", "LPN2", "โปรแกรม MGR", 56, ""],
      ["'010268", "ยอดธง", "กรวิรัตน์", "OP2S", "LPN2", "โปรแกรม MGR", 55, ""],
      ["'010285", "กฤษณะ", "เรือนเดื่อ", "OP2S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 54, ""],
      ["'010858", "กิตติ", "ปลิวมา", "IC", "LPN1", "โปรแกรมอายุ 35 ปีขึ้นไป", 56, ""],
      ["'011016", "ซ่อนกลิ่น", "ศรีอ่อน", "OP1S", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 51, ""],
      ["'011333", "กัญจน์ชญา", "ปัญญา", "QMSS", "LPN2", "โปรแกรมอายุ 35 ปีขึ้นไป", 58, ""],
      ["'011382", "เพียงอัมพร", "องค์วิศิษฐ์", "TRF", "LPN1", "โปรแกรม MGR", 57, ""],
      ["'011999", "ณัฐพงษ์", "รักเรียน", "IT", "LPN1", "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี", 28, "ไม่มีสิทธิ์ (อายุงานไม่ถึง 6 เดือน)"]
    ];
    for (var j = 0; j < demoEmployees.length; j++) {
      nameSheet.appendRow(demoEmployees[j]);
    }
  }
  
  // 2. Config_Dates Sheet
  var datesSheet = ss.getSheetByName("Config_Dates");
  if (!datesSheet) {
    datesSheet = ss.insertSheet("Config_Dates");
    datesSheet.appendRow(["สถานที่", "วันที่ตรวจ", "ทีม"]);
    datesSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9ead3");
    
    var defaultDates = [
      ["LPN1", "30 กันยายน 2569", "ทีม A"],
      ["LPN1", "1 ตุลาคม 2569", "ทีม A"],
      ["LPN1", "6 ตุลาคม 2569", "ทีม B"],
      ["LPN1", "7 ตุลาคม 2569", "ทีม B"],
      ["LPN2", "2 ตุลาคม 2569", "ทีม A"],
      ["LPN2", "5 ตุลาคม 2569", "ทีม B"]
    ];
    for (var k = 0; k < defaultDates.length; k++) {
      datesSheet.appendRow(defaultDates[k]);
    }
  }
  
  // 3. Config_TimeSlots Sheet
  var timesSheet = ss.getSheetByName("Config_TimeSlots");
  if (!timesSheet) {
    timesSheet = ss.insertSheet("Config_TimeSlots");
    timesSheet.appendRow(["รอบเวลา", "จำนวนจำกัด"]);
    timesSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#fce5cd");
    
    var defaultTimes = [
      ["08:00 - 08:30", 50],
      ["08:30 - 09:00", 50],
      ["09:00 - 09:30", 50],
      ["09:30 - 10:00", 50],
      ["10:00 - 10:30", 50],
      ["10:30 - 11:00", 50],
      ["11:00 - 11:30", 50],
      ["11:30 - 12:00", 50],
      ["12:00 - 12:30", 50],
      ["12:30 - 13:00", 50],
      ["13:00 - 13:30", 50],
      ["13:30 - 14:00", 50]
    ];
    for (var m = 0; m < defaultTimes.length; m++) {
      timesSheet.appendRow(defaultTimes[m]);
    }
  }
  
  // 4. Registration Sheet
  var regSheet = ss.getSheetByName("Registration");
  if (!regSheet) {
    regSheet = ss.insertSheet("Registration");
    regSheet.appendRow([
      "รหัสพนักงาน", "ชื่อ", "นามสกุล", "แผนก", "เบอร์โทรภายใน", 
      "กะทำงาน", "สถานที่", "วันที่ตรวจ", "เวลาที่ตรวจ", "รายการตรวจมะเร็งที่เลือก", "โปรแกรมปัจจัยเสี่ยง", "ตั้งครรภ์", "การยินยอมใช้สิทธิ์ประกันสังคม", "Timestamp"
    ]);
    regSheet.getRange("A1:N1").setFontWeight("bold").setBackground("#d9ead3");
  }
  
  return "Initialization successful. 'Name', 'Config_Dates', 'Config_TimeSlots', and 'Registration' sheets created/verified.";
}

/**
 * Helper to convert 1-based column index to A-Z / AA-ZZ Google Sheet column letter
 */
function getColumnLetter(colIndex) {
  var letter = "";
  while (colIndex > 0) {
    var temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

/**
 * Fetch all eligible employees and all registrations for the Admin Dashboard
 */
function getAdminDashboardData() {
  var ss = getSpreadsheet();
  
  var nameSheet = ss.getSheetByName("Name");
  var regSheet = ss.getSheetByName("Registration");
  
  var employees = [];
  if (nameSheet) {
    var headers = nameSheet.getRange(1, 1, 1, nameSheet.getLastColumn()).getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    var colId = headers.indexOf("รหัสพนักงาน");
    if (colId === -1) colId = headers.indexOf("fempno");
    
    var colName = headers.indexOf("ชื่อ");
    if (colName === -1) colName = headers.indexOf("fempnamet");
    
    var colLastName = headers.indexOf("นามสกุล");
    if (colLastName === -1) colLastName = headers.indexOf("fsurnamet");
    
    var colDept = headers.indexOf("แผนก");
    if (colDept === -1) colDept = headers.indexOf("fdeptcode");
    
    var colRight = -1;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h].indexOf("สิทธิ์") !== -1) {
        colRight = h;
        break;
      }
    }
    
    var data = nameSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      var rightVal = colRight !== -1 ? String(data[i][colRight]).trim() : "";
      // Exclude those marked as no checkup eligibility
      if (rightVal.indexOf("ไม่มีสิทธิ์") !== -1) {
        continue;
      }
      
      var idVal = colId !== -1 ? String(data[i][colId]).trim().replace(/^'/, '') : "";
      if (idVal === "") continue;
      if (/^\d+$/.test(idVal)) {
        idVal = idVal.padStart(6, '0');
      }
      
      employees.push({
        employeeId: idVal,
        firstName: colName !== -1 ? String(data[i][colName]).trim() : "",
        lastName: colLastName !== -1 ? String(data[i][colLastName]).trim() : "",
        department: colDept !== -1 ? String(data[i][colDept]).trim() : ""
      });
    }
  }
  
  var registrations = [];
  if (regSheet) {
    var rHeaders = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getDisplayValues()[0].map(function(h) { return String(h).trim(); });
    var colIdIdx = rHeaders.indexOf("รหัสพนักงาน");
    var colLoc = rHeaders.indexOf("สถานที่");
    var colDate = rHeaders.indexOf("วันที่ตรวจ");
    var colTime = rHeaders.indexOf("เวลาที่ตรวจ");
    
    var rData = regSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < rData.length; i++) {
      var idVal = colIdIdx !== -1 ? String(rData[i][colIdIdx]).trim().replace(/^'/, '') : "";
      if (idVal === "") continue;
      if (/^\d+$/.test(idVal)) {
        idVal = idVal.padStart(6, '0');
      }
      
      registrations.push({
        employeeId: idVal,
        location: colLoc !== -1 ? String(rData[i][colLoc]).trim() : "",
        dateString: colDate !== -1 ? String(rData[i][colDate]).trim() : "",
        timeString: colTime !== -1 ? String(rData[i][colTime]).trim() : ""
      });
    }
  }
  
  return {
    employees: employees,
    registrations: registrations
  };
}

/**
 * Update the 'ลงทะเบียน' column in the Name sheet for a specific employee
 */
function updateNameRegistrationStatus(employeeId, status) {
  var ss = getSpreadsheet();
  var nameSheet = ss.getSheetByName("Name");
  if (!nameSheet) return;
  
  var idToFind = String(employeeId).trim();
  if (/^\d+$/.test(idToFind)) {
    idToFind = idToFind.padStart(6, '0');
  }
  
  var data = nameSheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return;
  
  // Find headers and locate "ลงทะเบียน" (Column K, or index 10)
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colRegIdx = headers.indexOf("ลงทะเบียน");
  
  // If not found in headers, let's use Column G (column 7, index 6)
  if (colRegIdx === -1) {
    colRegIdx = 6; // 0-indexed index 6 is Column G
    // If sheet has fewer than 7 columns, write header first
    if (headers.length < 7) {
      nameSheet.getRange(1, 7).setValue("ลงทะเบียน").setFontWeight("bold");
    }
  }
  
  // Find employee row
  var colId = headers.indexOf("รหัสพนักงาน");
  if (colId === -1) colId = 0; // fallback to column A
  
  for (var i = 1; i < data.length; i++) {
    var rawId = String(data[i][colId]).trim().replace(/^'/, '');
    if (/^\d+$/.test(rawId)) {
      rawId = rawId.padStart(6, '0');
    }
    if (rawId === idToFind) {
      nameSheet.getRange(i + 1, colRegIdx + 1).setValue(status);
      try {
        CacheService.getScriptCache().remove("emp_" + idToFind);
      } catch (e) {
        console.warn("Cache eviction error in updateNameRegistrationStatus: " + e.toString());
      }
      break;
    }
  }
}

/**
 * Run this function once from the Apps Script editor to update the 'ลงทะเบียน' column 
 * in the Name sheet for all employees who have already registered.
 */
function backfillRegistrationStatus() {
  var ss = getSpreadsheet();
  var nameSheet = ss.getSheetByName("Name");
  var regSheet = ss.getSheetByName("Registration");
  if (!nameSheet || !regSheet) {
    return "Error: Name or Registration sheet not found.";
  }
  
  var nameData = nameSheet.getDataRange().getDisplayValues();
  var regData = regSheet.getDataRange().getDisplayValues();
  
  if (nameData.length <= 1) return "No employees in Name sheet.";
  if (regData.length <= 1) return "No registrations in Registration sheet.";
  
  // Find "รหัสพนักงาน" and "ลงทะเบียน" columns in Name sheet
  var nameHeaders = nameData[0].map(function(h) { return String(h).trim(); });
  var colNameId = nameHeaders.indexOf("รหัสพนักงาน");
  if (colNameId === -1) colNameId = nameHeaders.indexOf("fempno");
  if (colNameId === -1) colNameId = 0;
  
  var colRegIdx = nameHeaders.indexOf("ลงทะเบียน");
  if (colRegIdx === -1) {
    colRegIdx = 6; // Column G
    if (nameHeaders.length < 7) {
      nameSheet.getRange(1, 7).setValue("ลงทะเบียน").setFontWeight("bold");
    }
  }
  
  // Get all registered IDs from Registration sheet
  var regHeaders = regData[0].map(function(h) { return String(h).trim(); });
  var colRegId = regHeaders.indexOf("รหัสพนักงาน");
  if (colRegId === -1) colRegId = 0;
  
  var registeredIds = {};
  for (var j = 1; j < regData.length; j++) {
    var regId = String(regData[j][colRegId]).trim().replace(/^'/, '');
    if (/^\d+$/.test(regId)) {
      regId = regId.padStart(6, '0');
    }
    if (regId !== "") {
      registeredIds[regId] = true;
    }
  }
  
  // Update Name sheet
  var updatedCount = 0;
  for (var i = 1; i < nameData.length; i++) {
    var empId = String(nameData[i][colNameId]).trim().replace(/^'/, '');
    if (/^\d+$/.test(empId)) {
      empId = empId.padStart(6, '0');
    }
    
    var isRegistered = registeredIds[empId] === true;
    var currentStatus = colRegIdx < nameData[i].length ? String(nameData[i][colRegIdx]).trim() : "";
    
    if (isRegistered && currentStatus !== "ลงทะเบียนแล้ว") {
      nameSheet.getRange(i + 1, colRegIdx + 1).setValue("ลงทะเบียนแล้ว");
      updatedCount++;
    } else if (!isRegistered && currentStatus === "ลงทะเบียนแล้ว") {
      nameSheet.getRange(i + 1, colRegIdx + 1).setValue("");
      updatedCount++;
    }
  }
  
  return "Success! Updated " + updatedCount + " employees.";
}

