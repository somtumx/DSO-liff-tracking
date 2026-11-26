// ==========================================
// 1. CONFIGURATION (ตั้งค่าระบบ)
// ==========================================
// Setup in Google Apps Script -> Project Settings -> Script Properties:
//   - CHANNEL_ACCESS_TOKEN : LINE Messaging API Channel Access Token (required for LINE Bot)
//   - SPREADSHEET_ID       : (Optional) Explicit Spreadsheet ID if deployed standalone
//                            Leave blank when container-bound to WorkingStatusReport
//   - LINE_CHANNEL_ID      : (Optional) LINE Login Channel ID if enabling getMe / REQUIRE_AUTH
//   - REQUIRE_AUTH         : (Optional) "false" (default) or "true" for read endpoints
//
// Deploy -> New deployment -> Web app:
//   - Execute as: Me
//   - Who has access: Anyone

const LIFF_URL = 'https://liff.line.me/2008983827-ZbZI2YWu';
const LOGO_URL = 'https://upload.wikimedia.org/wikipedia/commons/thumb/c/ce/Coca-Cola_logo.svg/1200px-Coca-Cola_logo.svg.png';

const SHEET_LOG_NAME = 'Log';
const SHEET_BRANCH_NAME = 'Branch';
const SHEET_CONFIG_NAME = 'Config';
const SHEET_SYSTEM_LOG = 'SystemLog';
const SHEET_USERS_NAME = 'Users';
const SHEET_REPORT_NAME = 'Work status';

// SystemLog auto-trim settings (keeps sheet fast and within cell limits)
const SYSTEM_LOG_MAX_ROWS = 2000;
const SYSTEM_LOG_TRIM_TO = 1500;

// Script lock wait timeout (ms)
const LOCK_TIMEOUT_MS = 15000;

// ==========================================
// 2. MAIN SYSTEM (LINE Webhook & Web App API)
// ==========================================
function doPost(e) {
  try {
    const rawData = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
    if (!rawData) {
      return jsonResponse({ status: 'error', code: 'EMPTY_BODY', message: 'Empty request body' });
    }

    const data = JSON.parse(rawData);

    // 2A. Web App API POST (e.g. DSO-liff Cancel Order button / future POST actions)
    if (data && typeof data.action === 'string' && !Array.isArray(data.events)) {
      return handleWebPostAction(data);
    }

    // 2B. LINE Messaging API Webhook POST
    if (data && Array.isArray(data.events)) {
      data.events.forEach(event => {
        if (event && event.type === 'message' && event.message && event.message.type === 'text') {
          handleMessage(event);
        }
      });
    }

    return jsonResponse({ status: 'success' });
  } catch (error) {
    logSystem('ERROR', 'DO_POST_CRASH', error.message || String(error), '');
    return jsonResponse({ status: 'error', code: 'SERVER_ERROR', message: error.message || 'Internal server error' });
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = sanitizeString(params.action);

    // Never expose bulk employee Users sheet via GET
    if (action === 'getUsers') {
      return jsonResponse({ status: 'forbidden', code: 'ACTION_DISABLED', message: 'Bulk user listing is disabled for privacy' });
    }

    if (action === 'getTrackingData') {
      return getDataFromSheet();
    }
    if (action === 'getBranches') {
      return getBranchesList();
    }
    return jsonResponse({ status: 'ready' });
  } catch (error) {
    logSystem('ERROR', 'DO_GET_CRASH', error.message || String(error), '');
    return jsonResponse({ status: 'error', code: 'SERVER_ERROR', message: 'Internal server error' });
  }
}

function handleWebPostAction(body) {
  const action = sanitizeString(body.action);

  if (action === 'getUsers') {
    return jsonResponse({ status: 'forbidden', code: 'ACTION_DISABLED', message: 'Bulk user listing is disabled for privacy' });
  }

  if (action === 'cancelOrder') {
    return jsonResponse(cancelOrderFromWeb(body));
  }

  if (action === 'getTrackingData') {
    return getDataFromSheet();
  }

  if (action === 'getBranches') {
    return getBranchesList();
  }

  if (action === 'getMe') {
    const idToken = sanitizeString(body.idToken);
    if (!idToken) {
      return jsonResponse({ status: 'unauthorized', code: 'MISSING_TOKEN', message: 'LINE ID token is required' });
    }
    const claims = verifyLineIdToken(idToken);
    if (!claims || !claims.sub) {
      return jsonResponse({ status: 'unauthorized', code: 'INVALID_TOKEN', message: 'Invalid or expired LINE ID token' });
    }
    const registeredName = getUserName(getSpreadsheet(), claims.sub);
    return jsonResponse({
      status: 'success',
      data: {
        displayName: registeredName || claims.name || 'LINE User',
        role: registeredName ? 'staff' : 'user'
      }
    });
  }

  return jsonResponse({ status: 'error', code: 'INVALID_ACTION', message: 'Unsupported action' });
}

// ==========================================
// 3. SYSTEM LOGGING UTILITY (WITH AUTO-TRIM)
// ==========================================
function logSystem(type, action, message, detail, existingSs) {
  try {
    const ss = existingSs || getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SYSTEM_LOG);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SYSTEM_LOG);
      sheet.appendRow(['Timestamp', 'Type', 'Action', 'Message', 'Detail']);
      sheet.getRange('A1:E1').setFontWeight('bold');
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    let detailStr = typeof detail === 'object' && detail !== null ? JSON.stringify(detail) : String(detail || '');
    if (detailStr.length > 1000) {
      detailStr = detailStr.slice(0, 1000) + '...';
    }

    sheet.appendRow([timestamp, type, action, message, detailStr]);
    trimSystemLogSheet(sheet);
  } catch (e) {
    console.error('Logging Failed: ' + e.message);
  }
}

/**
 * Automatically trims old rows from SystemLog when total data rows exceed SYSTEM_LOG_MAX_ROWS.
 * Detects whether Row 1 is a header or raw data so no stale row is left behind.
 */
function trimSystemLogSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= SYSTEM_LOG_MAX_ROWS + 1) return;

  const firstCell = String(sheet.getRange(1, 1).getValue() || '').trim().toLowerCase();
  const hasHeader = firstCell === 'timestamp';
  const startRow = hasHeader ? 2 : 1;
  const dataRows = lastRow - (hasHeader ? 1 : 0);

  if (dataRows > SYSTEM_LOG_MAX_ROWS) {
    const rowsToDelete = dataRows - SYSTEM_LOG_TRIM_TO;
    if (rowsToDelete > 0) {
      sheet.deleteRows(startRow, rowsToDelete);
    }
  }
}

/**
 * Optional helper: run manually or via Time-driven trigger to trim SystemLog immediately.
 */
function trimSystemLogNow() {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SYSTEM_LOG);
    if (sheet) {
      trimSystemLogSheet(sheet);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. MESSAGE HANDLING (LINE BOT)
// ==========================================
function handleMessage(event) {
  const text = (event.message.text || '').trim();
  // Ignore non-command chat messages immediately (saves quota & prevents SystemLog bloat)
  if (!text.startsWith('/')) return;

  const ss = getSpreadsheet();
  const userId = (event.source && event.source.userId) ? event.source.userId : 'UNKNOWN_USER';
  const groupId = (event.source && (event.source.groupId || event.source.roomId)) ? (event.source.groupId || event.source.roomId) : 'NO_GROUP';

  // ---------------------------------------------------
  // ระบบลงทะเบียนผู้ใช้ (บันทึกลงชีท Users)
  // ---------------------------------------------------
  if (/^\/ลงทะเบียน(?:\s|$)/i.test(text)) {
    let name = text.replace(/^\/ลงทะเบียน\s*/i, '').trim();
    name = name.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s*(ครับ|คับ|คะ|ค่ะ|จ้า|จ๊ะ)\s*$/g, '').trim();

    if (name) {
      const regResult = saveUserToSheet(ss, userId, name);
      if (regResult.isUpdate) {
        replyText(event.replyToken, `✅ อัปเดตชื่อลงทะเบียนเป็น "${name}" เรียบร้อยแล้วครับ`, ss);
      } else {
        replyText(event.replyToken, `✅ ลงทะเบียนชื่อ "${name}" เรียบร้อยแล้วครับ ระบบจะใช้ชื่อนี้ในการแจ้งรายงานประจำวัน`, ss);
      }
      logSystem('INFO', 'REGISTER_USER', `ลงทะเบียนผู้ใช้ ${name}`, userId, ss);
    } else {
      replyText(event.replyToken, '⚠️ กรุณาระบุชื่อหลังคำสั่งด้วยครับ\nตัวอย่าง: /ลงทะเบียน สมชาย', ss);
    }
    return;
  }

  const cleanCommandText = text.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s*(ครับ|คับ|คะ|ค่ะ|จ้า|จ๊ะ|นะ)\s*$/g, '').trim();

  // ตั้งค่า / เพิ่มกลุ่ม Admin (รองรับหลายกลุ่มไม่จำกัด เช่น /AdminGroup หรือ /AdminGroup ชื่อกลุ่ม)
  const matchAddAdmin = cleanCommandText.match(/^\/AdminGroup(?:\s+(.+))?$/i);
  if (matchAddAdmin) {
    if (groupId === 'NO_GROUP') {
      replyText(event.replyToken, 'ใช้คำสั่งนี้ได้เฉพาะในกลุ่มไลน์เท่านั้น', ss);
      return;
    }
    const label = matchAddAdmin[1] ? matchAddAdmin[1].trim() : '';
    const res = addAdminGroupId(ss, groupId, label);
    if (res.added) {
      const labelInfo = label ? ` (${label})` : '';
      replyText(
        event.replyToken,
        `✅ เพิ่มกลุ่มนี้${labelInfo}เป็น Admin Group เรียบร้อยแล้วครับ\nกลุ่ม Admin ทั้งหมดตอนนี้: ${res.total} กลุ่ม`,
        ss
      );
      logSystem('INFO', 'SETUP_ADMIN', 'เพิ่มกลุ่มแอดมินใหม่', `${groupId}${labelInfo} (รวม ${res.total} กลุ่ม)`, ss);
    } else if (res.labelUpdated) {
      replyText(
        event.replyToken,
        `✅ อัปเดตชื่อกลุ่ม Admin เป็น "${label}" เรียบร้อยแล้วครับ\nกลุ่ม Admin ทั้งหมดตอนนี้: ${res.total} กลุ่ม`,
        ss
      );
      logSystem('INFO', 'UPDATE_ADMIN_LABEL', 'อัปเดตชื่อกลุ่มแอดมิน', `${groupId} (${label})`, ss);
    } else {
      replyText(
        event.replyToken,
        `ℹ️ กลุ่มนี้มีสิทธิ์ Admin Group อยู่แล้วครับ\nกลุ่ม Admin ทั้งหมดตอนนี้: ${res.total} กลุ่ม`,
        ss
      );
    }
    return;
  }

  // ถอนสิทธิ์กลุ่ม Admin ปัจจุบัน (เฉพาะกลุ่มที่เป็น Admin อยู่แล้ว)
  if (/^\/(?:RemoveAdminGroup|ลบAdminGroup)$/i.test(cleanCommandText)) {
    if (!isAdminGroup(ss, groupId)) {
      logSystem('WARNING', 'UNAUTHORIZED_ADMIN_CMD', 'พยายามลบสิทธิ์ Admin จากกลุ่มอื่น', `Group ID: ${groupId}`, ss);
      return;
    }
    const res = removeAdminGroupId(ss, groupId);
    replyText(
      event.replyToken,
      `✅ ลบกลุ่มนี้ออกจากสิทธิ์ Admin Group เรียบร้อยแล้วครับ\nเหลือกลุ่ม Admin ทั้งหมด: ${res.total} กลุ่ม`,
      ss
    );
    logSystem('INFO', 'REMOVE_ADMIN', 'ลบกลุ่มแอดมินออก', `${groupId} (เหลือ ${res.total} กลุ่ม)`, ss);
    return;
  }

  // ดูรายชื่อกลุ่ม Admin ทั้งหมด (เฉพาะกลุ่มที่เป็น Admin)
  if (/^\/AdminList$/i.test(cleanCommandText)) {
    if (!isAdminGroup(ss, groupId)) {
      logSystem('WARNING', 'UNAUTHORIZED_ADMIN_CMD', 'พยายามดูรายชื่อ Admin จากกลุ่มอื่น', `Group ID: ${groupId}`, ss);
      return;
    }
    const entries = getAdminGroupEntries(ss);
    const lines = entries.map((item, idx) => {
      const currentTag = item.groupId === groupId ? ' (กลุ่มนี้)' : '';
      const noteTag = item.note ? ` — ${item.note}` : '';
      return `${idx + 1}. ${item.groupId}${noteTag}${currentTag}`;
    });
    replyText(
      event.replyToken,
      `🛡️ รายชื่อกลุ่ม Admin ทั้งหมด (${entries.length} กลุ่ม)\n${lines.join('\n')}`,
      ss
    );
    return;
  }

  // เรียกดูรายงานสรุป (เฉพาะกลุ่ม Admin) รองรับระบุวันที่ เช่น /Dispatch done หรือ /Dispatch done 26.11.26
  const matchReport = cleanCommandText.match(/^\/Dispatch\s+done(?:\s+(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{2,4}))?$/i);
  if (matchReport) {
    if (isAdminGroup(ss, groupId)) {
      let targetDateStr = null;
      if (matchReport[1]) {
        const parsed = parseCommandDate(matchReport[1]);
        if (!parsed) {
          replyText(event.replyToken, `⚠️ รูปแบบวันที่ไม่ถูกต้อง: ${matchReport[1]}\nตัวอย่าง: /Dispatch done 26.11.26`, ss);
          return;
        }
        targetDateStr = parsed.shortDate;
      }
      sendDailySummary(ss, event.replyToken, targetDateStr);
    } else {
      logSystem('WARNING', 'UNAUTHORIZED_REPORT', 'พยายามเรียก Report จากกลุ่มอื่น', `Group ID: ${groupId}`, ss);
    }
    return;
  }

  // คำนวณเวลาใน Log ใหม่ + แก้ไขทับ Work status + ส่งรายงานสรุป (เฉพาะกลุ่ม Admin)
  // รองรับ: /Recalc, /Recalc 26.11.26, /คำนวณใหม่, /คำนวณใหม่ 26.11.69
  const matchRecalc = cleanCommandText.match(/^\/(?:Recalc|คำนวณใหม่)(?:\s+(.+))?$/i);
  if (matchRecalc) {
    if (!isAdminGroup(ss, groupId)) {
      logSystem('WARNING', 'UNAUTHORIZED_RECALC', 'พยายามสั่งคำนวณใหม่จากกลุ่มอื่น', `Group ID: ${groupId}`, ss);
      return;
    }

    const nowParts = toBangkokParts(new Date());
    let targetShortDate = `${nowParts.dd}/${nowParts.mm}/${nowParts.yy}`;
    let targetFullDate = `${nowParts.dd}/${nowParts.mm}/${nowParts.yyyy}`;

    if (matchRecalc[1] && matchRecalc[1].trim()) {
      const parsed = parseCommandDate(matchRecalc[1].trim());
      if (!parsed) {
        replyText(event.replyToken, `⚠️ รูปแบบวันที่ไม่ถูกต้อง: ${matchRecalc[1].trim()}\nตัวอย่าง: /Recalc 26.11.26 หรือ /คำนวณใหม่ 26.11.69`, ss);
        return;
      }
      targetShortDate = parsed.shortDate;
      targetFullDate = parsed.fullDate;
    }

    handleRecalcCommand(ss, event.replyToken, targetShortDate, targetFullDate);
    return;
  }

  // ประมวลผลคำสั่งงาน (BILL / CONFIRM / SHIPMENT)
  // หมายเหตุ: ยกเลิกคำสั่ง /เปลี่ยนรถ, /ช่วยส่ง และ /ยกเลิก ใน LINE แล้ว
  //           (การยกเลิกรายการย้ายไปกดปุ่ม "ยกเลิก" บนหน้าเว็บ DSO-liff แทน)
  logSystem('INFO', 'COMMAND_START', 'เริ่มประมวลผลคำสั่ง', text, ss);
  processCommand(ss, event.replyToken, text, userId);
}

// ==========================================
// 5. CORE TRANSACTION LOGIC (SILENT MODE)
// ==========================================
function processCommand(ss, replyToken, rawText, userId) {
  // 1. ตัดคำสุภาพ และตัวอักษรพิเศษที่ไม่ต้องการออกก่อน
  let text = rawText.replace(/[\u200B-\u200D\uFEFF]/g, '');
  text = text.replace(/\s*(ครับ|คับ|คะ|ค่ะ|จ้า|จ๊ะ|นะ)\s*$/g, '').trim();

  // 2. Regex ตรวจสอบคำสั่ง (แยกกลุ่มวันที่ 3 ส่วนชัด + บังคับเว้นวรรคก่อนจำนวน SH กัน backtracking)
  const regexBill     = /^\/\s*([A-Z]+)\s*(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{2,4})\s*DO\s*([\d,]+)/i;
  const regexConfirm  = /^\/ยืนยัน\s*([A-Z]+)\s*(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{2,4})\s*DO\s*([\d,]+)/i;
  const regexShipment = /^\/SH\s*([A-Z]+)\s*(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{2,4})\s+([\d,]+)\s*SH\b/i;

  let match, actionType = '', branchCode = '', dateInput = '', rawAmount = '';

  if ((match = text.match(regexConfirm))) {
    actionType = 'CONFIRM';
    branchCode = match[1];
    dateInput = match[2];
    rawAmount = match[3];
  } else if ((match = text.match(regexShipment))) {
    actionType = 'SHIPMENT';
    branchCode = match[1];
    dateInput = match[2];
    rawAmount = match[3];
  } else if ((match = text.match(regexBill))) {
    actionType = 'BILL';
    branchCode = match[1];
    dateInput = match[2];
    rawAmount = match[3];
  } else {
    logSystem('WARNING', 'INVALID_FORMAT', 'รูปแบบคำสั่งไม่ถูกต้องหรือไม่มีในระบบ', text, ss);
    return;
  }

  branchCode = branchCode.toUpperCase();
  const branchInfo = getBranchInfo(ss, branchCode);
  if (!branchInfo.found) {
    logSystem('WARNING', 'INVALID_BRANCH', `ไม่พบรหัสสาขา "${branchCode}" ในชีท Branch`, text, ss);
    return;
  }

  // 3. ตรวจสอบและแปลงวันที่ (รองรับทั้ง พ.ศ. 2/4 หลัก และ ค.ศ. 2/4 หลัก)
  const parsedDate = parseCommandDate(dateInput);
  if (!parsedDate) {
    logSystem('WARNING', 'INVALID_DATE', `รูปแบบวันที่ไม่ถูกต้อง: ${dateInput}`, text, ss);
    return;
  }

  const registeredName = getUserName(ss, userId);
  const uidStr = userId !== 'UNKNOWN_USER' ? userId.substring(0, 4) : 'ไม่ระบุ';
  const displayName = registeredName ? registeredName : ('เจ้าหน้าที่ (' + uidStr + ')');

  const amountVal = parseCleanInt(rawAmount);
  if (amountVal <= 0) {
    logSystem('WARNING', 'INVALID_AMOUNT', `จำนวนไม่ถูกต้อง: ${rawAmount}`, text, ss);
    return;
  }

  try {
    const result = updateLogTransaction(ss, actionType, branchCode, parsedDate.shortDate, displayName, amountVal);

    if (result.status !== 'success') {
      logSystem('WARNING', 'TRANSACTION_FAIL', 'ทำรายการไม่สำเร็จ', JSON.stringify(result), ss);
      return;
    }

    logSystem('SUCCESS', 'TRANSACTION_OK', 'บันทึกลงชีท Log สำเร็จ', `Action: ${actionType}, Branch: ${branchCode}, Amount: ${amountVal}`, ss);

    if (actionType === 'BILL' && result.queueData) {
      const replyMessage = `ลำดับการแจ้ง: ${result.queueData.sequence}\n` +
                           `สาขาที่กำลังจัด: ${result.queueData.processingList}\n` +
                           `สาขาที่รอคิว: ${result.queueData.waitingList}`;
      replyText(replyToken, replyMessage, ss);
    } else if (actionType === 'SHIPMENT') {
      writeToReportSheet(ss, {
        workingDate: result.history.billDateFull,
        deliveryDate: parsedDate.fullDate,
        salesOffice: branchInfo.salesOffice,
        branchCode: branchCode,
        billTime: result.history.billTime + ':00',
        confTime: result.history.confTime + ':00',
        shipTime: result.history.shipTime + ':00',
        amount: amountVal,
        operationTime: formatDuration(result.history.confTime, result.history.shipTime),
        respondence: registeredName ? registeredName : displayName
      });
      SpreadsheetApp.flush();
    }
  } catch (error) {
    logSystem('ERROR', 'PROCESS_COMMAND_CRASH', 'โค้ดพังขณะประมวลผลคำสั่ง', error.stack || error.message, ss);
  }
}

/**
 * Parses DD.MM.YY / DD/MM/YYYY (BE or CE) and validates calendar bounds.
 */
function parseCommandDate(dateInput) {
  const parts = String(dateInput || '').split(/[\.\/]/);
  if (parts.length !== 3) return null;

  const dayNum = parseInt(parts[0], 10);
  const monthNum = parseInt(parts[1], 10);
  const yearNum = parseInt(parts[2], 10);

  if (isNaN(dayNum) || isNaN(monthNum) || isNaN(yearNum)) return null;
  if (dayNum < 1 || dayNum > 31 || monthNum < 1 || monthNum > 12) return null;

  let fullYear;
  if (yearNum >= 2500) {
    fullYear = yearNum - 543;
  } else if (yearNum >= 2000) {
    fullYear = yearNum;
  } else if (yearNum >= 50 && yearNum <= 99) {
    fullYear = 2000 + (yearNum - 43);
  } else if (yearNum >= 0 && yearNum <= 49) {
    fullYear = 2000 + yearNum;
  } else {
    return null;
  }

  if (fullYear < 2020 || fullYear > 2100) return null;

  // Verify real calendar date (e.g. reject 31/02)
  const check = new Date(fullYear, monthNum - 1, dayNum);
  if (check.getFullYear() !== fullYear || check.getMonth() !== monthNum - 1 || check.getDate() !== dayNum) {
    return null;
  }

  const dayStr = String(dayNum).padStart(2, '0');
  const monthStr = String(monthNum).padStart(2, '0');
  const shortYearStr = String(fullYear % 100).padStart(2, '0');

  return {
    shortDate: `${dayStr}/${monthStr}/${shortYearStr}`,
    fullDate: `${dayStr}/${monthStr}/${fullYear}`
  };
}

function updateLogTransaction(ss, action, branch, dateStr, user, amount) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);

  try {
    const sheet = ss.getSheetByName(SHEET_LOG_NAME);
    const lastRow = sheet.getLastRow();
    const timestamp = new Date();
    const timeStr = Utilities.formatDate(timestamp, 'Asia/Bangkok', 'HH:mm');
    const todayStr = Utilities.formatDate(timestamp, 'Asia/Bangkok', 'dd/MM/yy');

    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 18).getValues() : [];

    let targetRow = -1;
    let historyData = {};
    let currentTotal = parseCleanInt(amount);
    let confirmedAmount = 0;
    let currentEditCount = 0;
    let isExistingRecord = false;
    let candidateIndex = -1;

    for (let i = 0; i < data.length; i++) {
      const rowBranch = String(data[i][1] || '').trim().toUpperCase();
      if (rowBranch === branch && formatDateForComparison(data[i][2]) === dateStr) {
        // Skip completed rows (already shipped)
        if (data[i][12] && data[i][12].toString().trim() !== '') continue;

        if (action === 'BILL') {
          if (formatDateForComparison(data[i][0]) === todayStr) {
            targetRow = i + 2;
            isExistingRecord = true;
            break;
          }
          continue;
        }

        if (action === 'CONFIRM') {
          const rowCurrentTotal = parseCleanInt(data[i][6] || data[i][3]);
          const inputAmount = parseCleanInt(amount);
          if (rowCurrentTotal === inputAmount) {
            targetRow = i + 2;
            break;
          } else {
            candidateIndex = i;
          }
        } else if (action === 'SHIPMENT') {
          if (data[i][8] && data[i][8].toString().trim() !== '') {
            targetRow = i + 2;
            break;
          } else {
            candidateIndex = i;
          }
        }
      }
    }

    if (targetRow === -1 && candidateIndex !== -1) {
      targetRow = candidateIndex + 2;
    }

    if (targetRow !== -1) {
      const i = targetRow - 2;
      historyData = {
        billDateFull: formatDateFullYear(data[i][0]),
        billUser: data[i][4],
        billTime: formatTimeOnly(data[i][5]),
        confUser: data[i][7],
        confTime: formatTimeOnly(data[i][8])
      };
      currentTotal = parseCleanInt(data[i][6] || data[i][3]);
      confirmedAmount = parseCleanInt(data[i][9]);
      currentEditCount = parseCleanInt(data[i][16]);
    }

    if (action === 'BILL') {
      const cleanAmt = parseCleanInt(amount);

      // Compute daily queue & processing/waiting lists from already-loaded `data` (single read)
      const dailyRecords = [];
      for (let k = 0; k < data.length; k++) {
        if (formatDateForComparison(data[k][0]) === todayStr) {
          const rowTime = data[k][0] instanceof Date ? data[k][0].getTime() : k;
          const bCode = String(data[k][1] || '').trim();
          const isUpdatingThisRow = isExistingRecord && (k + 2 === targetRow);
          const isConfirmed = isUpdatingThisRow
            ? false
            : Boolean(data[k][8] && data[k][8].toString().trim() !== '');
          const isShipped = Boolean(data[k][12] && data[k][12].toString().trim() !== '');
          dailyRecords.push({ time: rowTime, branch: bCode, isConfirmed: isConfirmed, isShipped: isShipped });
        }
      }

      dailyRecords.sort((a, b) => a.time - b.time);
      const processingBranches = [];
      const waitingBranches = [];

      dailyRecords.forEach(item => {
        if (!item.isShipped) {
          if (item.isConfirmed) {
            processingBranches.push(item.branch);
          } else {
            waitingBranches.push(item.branch);
          }
        }
      });

      if (isExistingRecord) {
        // Keep original DO_Num (col 4), update Amt_Bill (col 7), increment Edit Count (col 17)
        sheet.getRange(targetRow, 7).setValue(cleanAmt);
        sheet.getRange(targetRow, 17).setValue(currentEditCount + 1);
        if (historyData.confTime && historyData.confTime !== '-') {
          sheet.getRange(targetRow, 8, 1, 4).clearContent();
        }
        SpreadsheetApp.flush();

        const existingSeq = parseCleanInt(data[targetRow - 2][15]) || dailyRecords.length;
        return {
          status: 'success',
          isUpdate: true,
          currentTimeStr: getThaiDateTime(timestamp),
          queueData: {
            sequence: existingSeq,
            processingList: processingBranches.length > 0 ? processingBranches.join(',') : '-',
            waitingList: waitingBranches.length > 0 ? waitingBranches.join(',') : '-'
          },
          currentTotal: cleanAmt
        };
      }

      waitingBranches.push(branch);
      const newSeq = dailyRecords.length + 1;
      sheet.appendRow([timestamp, branch, dateStr, cleanAmt, user, timeStr, cleanAmt, '', '', '', '', '', '', '', '', newSeq, 0, '']);
      SpreadsheetApp.flush();

      return {
        status: 'success',
        isUpdate: false,
        currentTimeStr: getThaiDateTime(timestamp),
        queueData: {
          sequence: newSeq,
          processingList: processingBranches.length > 0 ? processingBranches.join(',') : '-',
          waitingList: waitingBranches.join(',')
        },
        currentTotal: cleanAmt
      };
    }

    if (targetRow !== -1) {
      if (action === 'CONFIRM') {
        const inputConfirmAmt = parseCleanInt(amount);
        if (inputConfirmAmt !== currentTotal) {
          return { status: 'error_mismatch', input: inputConfirmAmt, current: currentTotal };
        }
        const diff1 = calculateTimeDiff(historyData.billTime, timeStr);
        sheet.getRange(targetRow, 8, 1, 4).setValues([[user, timeStr, inputConfirmAmt, diff1]]);
        SpreadsheetApp.flush();
        return { status: 'success' };
      }

      if (action === 'SHIPMENT') {
        if (!historyData.confTime || historyData.confTime === '-') {
          return { status: 'error_not_confirmed' };
        }
        if (confirmedAmount !== currentTotal) {
          return { status: 'error_amount_changed' };
        }

        const diff2 = calculateTimeDiff(historyData.confTime, timeStr);
        const cleanShAmt = parseCleanInt(amount);

        sheet.getRange(targetRow, 12, 1, 4).setValues([[user, timeStr, cleanShAmt, diff2]]);
        SpreadsheetApp.flush();

        return {
          status: 'success',
          currentTimeStr: getThaiDateTime(timestamp),
          history: { ...historyData, shipUser: user, shipTime: timeStr },
          timeStats: {
            diff1: calculateTimeDiff(historyData.billTime, historyData.confTime),
            diff2: diff2,
            total: calculateTimeDiff(historyData.billTime, timeStr)
          }
        };
      }
    }

    return { status: 'error_not_found', action: action, branch: branch, date: dateStr };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 6. WEB CANCEL ORDER (ปุ่มยกเลิกจากหน้า DSO-liff)
// ==========================================
/**
 * Deletes an unshipped order row immediately and recalculates Daily_Seq for that day.
 * Protected by ScriptLock + row fingerprint matching (rowIndex + branch + targetDate + timestamp).
 */
function cancelOrderFromWeb(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG_NAME);
    if (!sheet) {
      return { status: 'error', code: 'SHEET_NOT_FOUND', message: 'ไม่พบชีท Log' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'error', code: 'NOT_FOUND', message: 'ไม่พบรายการในชีท Log' };
    }

    const reqRowIndex = parseInt(body.rowIndex, 10);
    const reqBranch = sanitizeString(body.branch).toUpperCase();
    const reqTargetDate = sanitizeString(body.targetDate);
    const reqTimestamp = sanitizeString(body.timestamp);
    const actorName = sanitizeString(body.actorName) || 'Web User';

    if (!reqBranch || !reqTargetDate) {
      return { status: 'error', code: 'INVALID_PARAMS', message: 'ข้อมูลรายการสำหรับยกเลิกไม่ครบถ้วน' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();

    const rowMatchesFingerprint = (row) => {
      const rBranch = String(row[1] || '').trim().toUpperCase();
      const rTarget = formatDateForComparison(row[2]);
      const rTs = formatDateFull(row[0]);
      if (rBranch !== reqBranch || rTarget !== reqTargetDate) return false;
      if (reqTimestamp && rTs !== reqTimestamp) return false;
      return true;
    };

    // 1. Check requested rowIndex directly first
    let matchedIndex = -1;
    if (!isNaN(reqRowIndex) && reqRowIndex >= 2 && reqRowIndex <= lastRow) {
      const idx = reqRowIndex - 2;
      if (rowMatchesFingerprint(data[idx])) {
        matchedIndex = idx;
      }
    }

    // 2. Fallback scan (in case another row above was inserted/deleted right before this click)
    if (matchedIndex === -1) {
      for (let i = data.length - 1; i >= 0; i--) {
        if (rowMatchesFingerprint(data[i])) {
          matchedIndex = i;
          break;
        }
      }
    }

    if (matchedIndex === -1) {
      logSystem('WARNING', 'WEB_CANCEL_FAIL', `ไม่พบรายการสำหรับยกเลิกสาขา ${reqBranch}`, `Row: ${reqRowIndex}, Date: ${reqTargetDate}, TS: ${reqTimestamp}`, ss);
      return { status: 'error', code: 'NOT_FOUND', message: 'ไม่พบรายการนี้ หรือรายการอาจถูกยกเลิกไปแล้ว' };
    }

    const matchedRow = data[matchedIndex];
    const isShipped = Boolean(matchedRow[12] && matchedRow[12].toString().trim() !== '');
    if (isShipped) {
      logSystem('WARNING', 'WEB_CANCEL_FAIL', `ไม่สามารถยกเลิกสาขา ${reqBranch} เพราะปิด Shipment แล้ว`, `Row: ${matchedIndex + 2}, Date: ${reqTargetDate}`, ss);
      return { status: 'error', code: 'ALREADY_SHIPPED', message: 'งานนี้ปิด Shipment ไปแล้ว ไม่สามารถยกเลิกได้' };
    }

    const orderDayStr = formatDateForComparison(matchedRow[0]);
    const deletedRowNumber = matchedIndex + 2;
    const deletedAmount = parseCleanInt(matchedRow[6] || matchedRow[3]);

    sheet.deleteRow(deletedRowNumber);
    recalculateDailySequence(sheet, orderDayStr);

    logSystem(
      'SUCCESS',
      'WEB_CANCEL_OK',
      `ยกเลิกรายการสาขา ${reqBranch} จากหน้าเว็บสำเร็จ`,
      `Date: ${reqTargetDate}, DO: ${deletedAmount}, Row: ${deletedRowNumber}, By: ${actorName}`,
      ss
    );

    return {
      status: 'success',
      message: `ยกเลิกรายการสาขา ${reqBranch} เรียบร้อยแล้ว`,
      deleted: {
        rowIndex: deletedRowNumber,
        branch: reqBranch,
        targetDate: reqTargetDate
      }
    };
  } catch (error) {
    logSystem('ERROR', 'WEB_CANCEL_CRASH', 'ยกเลิกรายการจากหน้าเว็บไม่สำเร็จ', error.stack || error.message);
    return { status: 'error', code: 'SERVER_ERROR', message: 'เกิดข้อผิดพลาดระหว่างยกเลิกรายการ' };
  } finally {
    lock.releaseLock();
  }
}

function recalculateDailySequence(sheet, targetDayStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const numRows = lastRow - 1;
  const timestamps = sheet.getRange(2, 1, numRows, 1).getValues();
  const seqRange = sheet.getRange(2, 16, numRows, 1);
  const seqValues = seqRange.getValues();

  let seq = 1;
  let changed = false;

  for (let i = 0; i < numRows; i++) {
    if (formatDateForComparison(timestamps[i][0]) === targetDayStr) {
      if (seqValues[i][0] !== seq) {
        seqValues[i][0] = seq;
        changed = true;
      }
      seq++;
    }
  }

  if (changed) {
    seqRange.setValues(seqValues);
  }
}

// ==========================================
// 7. SECONDARY ROUTING REPORT & RECALCULATION (สำหรับกลุ่ม Admin / ชีท)
// ==========================================
function buildDailySummaryFlex(ss, targetDateStr) {
  const sheet = ss.getSheetByName(SHEET_LOG_NAME);
  const lastRow = sheet ? sheet.getLastRow() : 0;

  if (!sheet || lastRow <= 1) {
    return { empty: true, text: `ไม่มีข้อมูลงานประจำวันที่ ${targetDateStr}` };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const branchMap = {};
  let totalDO = 0;
  let totalSH = 0;

  data.forEach(row => {
    if (formatDateForComparison(row[0]) === targetDateStr) {
      const b = row[1];
      if (!b) return;
      if (!branchMap[b]) {
        branchMap[b] = { bill: 0, sh: 0, bTime: '-', sTime: '-' };
      }
      const billVal = parseCleanInt(row[6] || row[3]);
      const shVal = parseCleanInt(row[13]);

      branchMap[b].bill += billVal;
      branchMap[b].sh += shVal;
      totalDO += billVal;
      totalSH += shVal;

      if (row[5]) branchMap[b].bTime = formatTimeOnly(row[5]);
      if (row[12]) branchMap[b].sTime = formatTimeOnly(row[12]);
    }
  });

  const branches = Object.keys(branchMap);
  if (branches.length === 0) {
    return { empty: true, text: `ไม่มีรายการประจำวันที่ ${targetDateStr}` };
  }

  const rows = branches.map(b => {
    const s = branchMap[b];
    return {
      type: 'box',
      layout: 'horizontal',
      margin: 'md',
      contents: [
        { type: 'text', text: String(b), size: 'sm', weight: 'bold', flex: 2 },
        { type: 'text', text: s.bill.toLocaleString(), size: 'sm', color: '#F40009', align: 'center', flex: 2 },
        { type: 'text', text: s.sh.toLocaleString(), size: 'sm', color: '#006747', align: 'center', flex: 2 },
        { type: 'text', text: `${s.bTime}/${s.sTime}`, size: 'xxs', color: '#666666', align: 'end', flex: 3 }
      ]
    };
  });

  const flexDashboard = {
    type: 'bubble',
    size: 'mega',
    header: {
      type: 'box',
      layout: 'horizontal',
      paddingBottom: '0px',
      contents: [
        {
          type: 'box',
          layout: 'vertical',
          flex: 3,
          contents: [
            { type: 'text', text: 'SECONDARY ROUTING', color: '#F40009', size: 'xs', weight: 'bold' },
            { type: 'text', text: 'Daily Summary', size: 'xl', weight: 'bold', margin: 'xs' },
            { type: 'text', text: `ประจำวันที่ ${targetDateStr}`, size: 'sm', color: '#888888' }
          ]
        },
        { type: 'image', url: LOGO_URL, aspectMode: 'fit', flex: 2, gravity: 'center' }
      ]
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'box',
          layout: 'horizontal',
          spacing: 'md',
          contents: [
            {
              type: 'box', layout: 'vertical', flex: 1, backgroundColor: '#F40009', paddingAll: 'md', cornerRadius: 'md',
              contents: [
                { type: 'text', text: 'TOTAL DO', size: 'xxs', color: '#FFFFFF' },
                { type: 'text', text: totalDO.toLocaleString(), size: 'xl', weight: 'bold', color: '#FFFFFF' }
              ]
            },
            {
              type: 'box', layout: 'vertical', flex: 1, backgroundColor: '#006747', paddingAll: 'md', cornerRadius: 'md',
              contents: [
                { type: 'text', text: 'TOTAL SH', size: 'xxs', color: '#FFFFFF' },
                { type: 'text', text: totalSH.toLocaleString(), size: 'xl', weight: 'bold', color: '#FFFFFF' }
              ]
            }
          ]
        },
        {
          type: 'box',
          layout: 'horizontal',
          margin: 'xl',
          contents: [
            { type: 'text', text: 'BRANCH', size: 'xxs', color: '#AAAAAA', flex: 2 },
            { type: 'text', text: 'DO', size: 'xxs', color: '#AAAAAA', align: 'center', flex: 2 },
            { type: 'text', text: 'SH', size: 'xxs', color: '#AAAAAA', align: 'center', flex: 2 },
            { type: 'text', text: 'TIME (B/S)', size: 'xxs', color: '#AAAAAA', align: 'end', flex: 3 }
          ]
        },
        { type: 'separator', margin: 'sm' },
        ...rows
      ]
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      contents: [
        { type: 'text', text: 'HAADTHIP PUBLIC COMPANY LIMITED', size: 'xxs', color: '#AAAAAA', align: 'center', weight: 'bold' }
      ]
    }
  };

  return { empty: false, flex: flexDashboard, totalDO: totalDO, totalSH: totalSH, branchCount: branches.length };
}

function sendDailySummary(ss, replyToken, customDateStr) {
  const nowParts = toBangkokParts(new Date());
  const targetDateStr = customDateStr || `${nowParts.dd}/${nowParts.mm}/${nowParts.yy}`;
  const built = buildDailySummaryFlex(ss, targetDateStr);

  if (built.empty) {
    replyText(replyToken, built.text, ss);
    return;
  }

  sendToLine('reply', {
    replyToken: replyToken,
    messages: [{ type: 'flex', altText: `รายงานสรุปงาน ${targetDateStr}`, contents: built.flex }]
  }, ss);
  logSystem('INFO', 'REPORT_SENT', `ส่งสรุปรายวันสำเร็จ (${targetDateStr})`, '', ss);
}

/**
 * LINE Admin command handler for /Recalc [DD.MM.YY] or /คำนวณใหม่ [DD.MM.YY]
 * 1) Recalculates Diff_1 & Diff_2 in Log for the selected Working Date
 * 2) Overwrites existing matching rows in Work status in-place (preserving col K:M notes)
 * 3) Replies with both summary status and the updated Daily Summary Flex report
 */
function handleRecalcCommand(ss, replyToken, targetShortDate, targetFullDate) {
  const res = recalculateAndSyncByDate(ss, targetShortDate, targetFullDate);
  if (res.status !== 'success') {
    replyText(replyToken, `❌ คำนวณเวลาใหม่ไม่สำเร็จ: ${res.message || 'เกิดข้อผิดพลาด'}`, ss);
    return;
  }

  if (res.totalLogRows === 0) {
    replyText(replyToken, `⚠️ ไม่พบรายการในชีท Log สำหรับวันที่ ${targetShortDate}`, ss);
    return;
  }

  const summaryText =
    `✅ คำนวณเวลาและอัปเดต Work status เรียบร้อยแล้ว\n` +
    `วันที่: ${targetShortDate}\n` +
    `- ตรวจสอบ Log: ${res.totalLogRows} รายการ (แก้เวลา ${res.logUpdated} แถว)\n` +
    `- แก้ไขทับ Work status: ${res.wsUpdated} บรรทัด` +
    (res.wsInserted > 0 ? `\n- เพิ่มบรรทัดที่ตกหล่น: ${res.wsInserted} บรรทัด` : '');

  const messages = [{ type: 'text', text: summaryText }];
  const built = buildDailySummaryFlex(ss, targetShortDate);
  if (!built.empty) {
    messages.push({ type: 'flex', altText: `รายงานสรุปงาน ${targetShortDate}`, contents: built.flex });
  }

  sendToLine('reply', { replyToken: replyToken, messages: messages }, ss);
}

/**
 * Core recalculation engine:
 * - Recalculates Diff_1 (col K) and Diff_2 (col O) in Log for targetShortDate (dd/MM/yy)
 * - Overwrites matching Work status rows (cols A:J) in-place without duplicating rows or wiping K:M notes
 * - Appends any completed shipments missing from Work status
 */
function recalculateAndSyncByDate(ss, targetShortDate, targetFullDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);

  try {
    const logSheet = ss.getSheetByName(SHEET_LOG_NAME);
    if (!logSheet || logSheet.getLastRow() <= 1) {
      return { status: 'success', date: targetShortDate, totalLogRows: 0, logUpdated: 0, wsUpdated: 0, wsInserted: 0 };
    }

    // Build Branch lookup map once (BranchCode -> SalesOffice)
    const branchOfficeMap = {};
    const branchSheet = ss.getSheetByName(SHEET_BRANCH_NAME);
    if (branchSheet && branchSheet.getLastRow() > 1) {
      const bData = branchSheet.getRange(2, 1, branchSheet.getLastRow() - 1, 3).getValues();
      for (let i = 0; i < bData.length; i++) {
        const code = String(bData[i][0] || '').trim().toUpperCase();
        if (code) branchOfficeMap[code] = bData[i][2] || '-';
      }
    }

    const logLastRow = logSheet.getLastRow();
    const logData = logSheet.getRange(2, 1, logLastRow - 1, 16).getValues();

    let totalLogRows = 0;
    let logUpdated = 0;
    const completedOrders = [];

    for (let i = 0; i < logData.length; i++) {
      const row = logData[i];
      if (!row[1]) continue;
      if (formatDateForComparison(row[0]) !== targetShortDate) continue;

      totalLogRows++;
      const rowNum = i + 2;
      const branchCode = String(row[1]).trim().toUpperCase();
      const billTime = formatTimeOnly(row[5]);
      const confTime = formatTimeOnly(row[8]);
      const shipTime = formatTimeOnly(row[12]);

      const newDiff1 = (billTime !== '-' && confTime !== '-') ? calculateTimeDiff(billTime, confTime) : '';
      const newDiff2 = (confTime !== '-' && shipTime !== '-') ? calculateTimeDiff(confTime, shipTime) : '';

      const oldDiff1 = row[10] === null || row[10] === undefined ? '' : row[10];
      const oldDiff2 = row[14] === null || row[14] === undefined ? '' : row[14];

      let rowChanged = false;
      if (String(oldDiff1) !== String(newDiff1)) {
        logSheet.getRange(rowNum, 11).setValue(newDiff1);
        rowChanged = true;
      }
      if (String(oldDiff2) !== String(newDiff2)) {
        logSheet.getRange(rowNum, 15).setValue(newDiff2);
        rowChanged = true;
      }
      if (rowChanged) logUpdated++;

      if (shipTime !== '-') {
        const deliveryFull = formatDateFullYear(row[2]);
        const deliveryShort = formatDateForComparison(row[2]);
        completedOrders.push({
          workingDate: formatDateFullYear(row[0]) !== '-' ? formatDateFullYear(row[0]) : targetFullDate,
          deliveryDate: deliveryFull,
          deliveryShortDate: deliveryShort,
          salesOffice: branchOfficeMap[branchCode] || '-',
          branchCode: branchCode,
          billTime: billTime !== '-' ? billTime + ':00' : '-',
          confTime: confTime !== '-' ? confTime + ':00' : '-',
          shipTime: shipTime + ':00',
          amount: parseCleanInt(row[13]),
          operationTime: formatDuration(confTime, shipTime),
          respondence: String(row[11] || row[7] || row[4] || '').trim()
        });
      }
    }

    // Sync completed orders into Work status (overwrite matching rows in-place)
    let wsSheet = ss.getSheetByName(SHEET_REPORT_NAME);
    if (!wsSheet) {
      wsSheet = ss.insertSheet(SHEET_REPORT_NAME);
      wsSheet.appendRow([
        'Working Date', 'Delivery Date', 'Sales office', 'เขตการขาย',
        'Time to send Order', 'Confirm order time', 'Shipment created time',
        'Shipments', 'Operation time', 'Respondence',
        'สถานะ SH', 'ข้อติดขัด', 'หมายเหตุ', 'Month', 'Weekday', 'Weeknum'
      ]);
      wsSheet.getRange('A1:P1').setFontWeight('bold');
    }

    const wsLastRow = wsSheet.getLastRow();
    const wsExactMap = {};   // key: BRANCH|deliveryShortDate -> [rowNum, ...]
    const wsBranchMap = {};  // fallback key: BRANCH -> [rowNum, ...]
    const usedWsRows = {};
    const wsExistingKP = {};

    if (wsLastRow > 1) {
      const wsData = wsSheet.getRange(2, 1, wsLastRow - 1, 16).getValues();
      for (let i = 0; i < wsData.length; i++) {
        if (formatDateForComparison(wsData[i][0]) !== targetShortDate) continue;
        const rNum = i + 2;
        const bCode = String(wsData[i][3] || '').trim().toUpperCase();
        const dShort = formatDateForComparison(wsData[i][1]);
        if (!bCode) continue;

        const exactKey = `${bCode}|${dShort}`;
        if (!wsExactMap[exactKey]) wsExactMap[exactKey] = [];
        wsExactMap[exactKey].push(rNum);

        if (!wsBranchMap[bCode]) wsBranchMap[bCode] = [];
        wsBranchMap[bCode].push(rNum);

        wsExistingKP[rNum] = wsData[i].slice(10, 16);
      }
    }

    const takeUnusedRow = (list) => {
      if (!list) return -1;
      while (list.length > 0) {
        const candidate = list.shift();
        if (!usedWsRows[candidate]) {
          usedWsRows[candidate] = true;
          return candidate;
        }
      }
      return -1;
    };

    let wsUpdated = 0;
    let wsInserted = 0;

    completedOrders.forEach(item => {
      const exactKey = `${item.branchCode}|${item.deliveryShortDate}`;
      let matchRow = takeUnusedRow(wsExactMap[exactKey]);
      if (matchRow === -1) {
        matchRow = takeUnusedRow(wsBranchMap[item.branchCode]);
      }

      const coreCols = [
        item.workingDate,
        item.deliveryDate,
        item.salesOffice,
        item.branchCode,
        item.billTime,
        item.confTime,
        item.shipTime,
        item.amount,
        item.operationTime,
        item.respondence
      ];

      if (matchRow !== -1) {
        // Overwrite A:J in-place and preserve existing K:M notes (fill N:P formulas only if empty)
        wsSheet.getRange(matchRow, 1, 1, 10).setValues([coreCols]);
        const kp = wsExistingKP[matchRow] || [];
        if (!kp[3] || !kp[4] || !kp[5]) {
          wsSheet.getRange(matchRow, 14, 1, 3).setValues([[
            `=MONTH(A${matchRow})`,
            `=WEEKDAY(A${matchRow})`,
            `=WEEKNUM(A${matchRow})`
          ]]);
        }
        wsUpdated++;
      } else {
        const nextRow = wsSheet.getLastRow() + 1;
        wsSheet.getRange(nextRow, 1, 1, 16).setValues([[
          ...coreCols,
          'ปกติ',
          'ไม่มี',
          'ไม่มี',
          `=MONTH(A${nextRow})`,
          `=WEEKDAY(A${nextRow})`,
          `=WEEKNUM(A${nextRow})`
        ]]);
        wsInserted++;
      }
    });

    SpreadsheetApp.flush();

    logSystem(
      'SUCCESS',
      'RECALC_OK',
      `คำนวณเวลาและซิงก์ Work status วันที่ ${targetShortDate} สำเร็จ`,
      `Log: ${totalLogRows} (Updated ${logUpdated}), WorkStatus Updated: ${wsUpdated}, Inserted: ${wsInserted}`,
      ss
    );

    return {
      status: 'success',
      date: targetShortDate,
      totalLogRows: totalLogRows,
      logUpdated: logUpdated,
      wsUpdated: wsUpdated,
      wsInserted: wsInserted
    };
  } catch (error) {
    logSystem('ERROR', 'RECALC_CRASH', `คำนวณเวลาใหม่วันที่ ${targetShortDate} ล้มเหลว`, error.stack || error.message, ss);
    return { status: 'error', message: error.message || String(error) };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------
// Google Sheets Menu & Manual Editor Helpers
// ------------------------------------------
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('⚙️ DSO Tools')
      .addItem('คำนวณเวลาและอัปเดต Work status ใหม่ (เลือกวันที่)...', 'promptRecalculateDate')
      .addItem('ตัดแถวเก่าใน SystemLog ทันที', 'trimSystemLogNow')
      .addToUi();
  } catch (_e) {}
}

function promptRecalculateDate() {
  const ui = SpreadsheetApp.getUi();
  const nowParts = toBangkokParts(new Date());
  const defaultInput = `${nowParts.dd}.${nowParts.mm}.${nowParts.yy}`;

  const response = ui.prompt(
    'คำนวณเวลาและอัปเดต Work status ใหม่',
    `ระบุวันที่ (Working Date) ที่ต้องการคำนวณใหม่ เช่น ${defaultInput} หรือ ${nowParts.dd}.${nowParts.mm}.${parseInt(nowParts.yyyy, 10) + 543}\n(เว้นว่างเพื่อใช้วันนี้: ${defaultInput})`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const inputStr = response.getResponseText().trim() || defaultInput;
  const parsed = parseCommandDate(inputStr);
  if (!parsed) {
    ui.alert('รูปแบบวันที่ไม่ถูกต้อง', `กรุณาระบุวันที่ในรูปแบบ DD.MM.YY เช่น ${defaultInput}`, ui.ButtonSet.OK);
    return;
  }

  const res = recalculateAndSyncByDate(getSpreadsheet(), parsed.shortDate, parsed.fullDate);
  if (res.status === 'success') {
    ui.alert(
      'คำนวณเวลาใหม่สำเร็จ',
      `วันที่: ${res.date}\n- ตรวจสอบ Log: ${res.totalLogRows} รายการ (แก้เวลา ${res.logUpdated} แถว)\n- แก้ไขทับ Work status: ${res.wsUpdated} บรรทัด\n- เพิ่มบรรทัดที่ตกหล่น: ${res.wsInserted} บรรทัด`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('เกิดข้อผิดพลาด', res.message || 'ไม่สามารถคำนวณเวลาใหม่ได้', ui.ButtonSet.OK);
  }
}

function recalculateToday() {
  const nowParts = toBangkokParts(new Date());
  return recalculateAndSyncByDate(
    getSpreadsheet(),
    `${nowParts.dd}/${nowParts.mm}/${nowParts.yy}`,
    `${nowParts.dd}/${nowParts.mm}/${nowParts.yyyy}`
  );
}

function recalculateByDate(dateInputStr) {
  const parsed = parseCommandDate(dateInputStr);
  if (!parsed) throw new Error('Invalid date format: ' + dateInputStr + ' (expected DD.MM.YY)');
  return recalculateAndSyncByDate(getSpreadsheet(), parsed.shortDate, parsed.fullDate);
}

// ==========================================
// 8. UTILITIES & HELPERS
// ==========================================
function getSpreadsheet() {
  const explicitId = getProperty('SPREADSHEET_ID', '');
  return explicitId
    ? SpreadsheetApp.openById(explicitId)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getProperty(key, defaultValue) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  return (val !== null && val !== undefined && String(val).trim() !== '') ? String(val).trim() : defaultValue;
}

function sanitizeString(val) {
  if (val === null || val === undefined) return '';
  return String(val).replace(/[\u0000-\u001F\u007F]/g, '').trim();
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUsersDataStartIndex(data) {
  if (!data || data.length === 0) return 0;
  const firstCell = String(data[0][0] || '').trim();
  // If Row 1 already starts with a LINE userId (no header row), include Row 1
  if (/^U[0-9a-fA-F]{16,}$/.test(firstCell)) {
    return 0;
  }
  return 1;
}

function saveUserToSheet(ss, userId, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);

  try {
    let sheet = ss.getSheetByName(SHEET_USERS_NAME);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USERS_NAME);
      sheet.appendRow(['User ID', 'Name', 'Updated At']);
      sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#d9ead3');
    }

    const data = sheet.getDataRange().getValues();
    const startIdx = getUsersDataStartIndex(data);
    let rowIndex = -1;

    for (let i = startIdx; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === userId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2, 1, 2).setValues([[name, timestamp]]);
      SpreadsheetApp.flush();
      return { status: 'success', isUpdate: true };
    }

    sheet.appendRow([userId, name, timestamp]);
    SpreadsheetApp.flush();
    return { status: 'success', isUpdate: false };
  } finally {
    lock.releaseLock();
  }
}

function getUserName(ss, userId) {
  if (!userId || userId === 'UNKNOWN_USER') return null;
  const sheet = ss.getSheetByName(SHEET_USERS_NAME);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const startIdx = getUsersDataStartIndex(data);
  for (let i = startIdx; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === userId) {
      return data[i][1];
    }
  }
  return null;
}

function writeToReportSheet(ss, data) {
  try {
    let sheet = ss.getSheetByName(SHEET_REPORT_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_REPORT_NAME);
      sheet.appendRow([
        'Working Date', 'Delivery Date', 'Sales office', 'เขตการขาย',
        'Time to send Order', 'Confirm order time', 'Shipment created time',
        'Shipments', 'Operation time', 'Respondence',
        'สถานะ SH', 'ข้อติดขัด', 'หมายเหตุ', 'Month', 'Weekday', 'Weeknum'
      ]);
      sheet.getRange('A1:P1').setFontWeight('bold');
    }
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 16).setValues([[
      data.workingDate,
      data.deliveryDate,
      data.salesOffice,
      data.branchCode,
      data.billTime,
      data.confTime,
      data.shipTime,
      data.amount,
      data.operationTime,
      data.respondence,
      'ปกติ',
      'ไม่มี',
      'ไม่มี',
      `=MONTH(A${nextRow})`,
      `=WEEKDAY(A${nextRow})`,
      `=WEEKNUM(A${nextRow})`
    ]]);
  } catch (e) {
    logSystem('ERROR', 'REPORT_WRITE_FAIL', 'บันทึกรายงาน Work status ไม่สำเร็จ', e.message, ss);
  }
}

function isAdminConfigKey(key) {
  const k = String(key || '').trim().toUpperCase();
  return /^ADMIN_GROUP_ID(?:S|_\d+)?$/.test(k);
}

function computeAdminRowFlags(data) {
  const flags = [];
  let inAdminBlock = false;
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (key) {
      inAdminBlock = isAdminConfigKey(key);
    }
    flags.push(inAdminBlock);
  }
  return flags;
}

/**
 * Reads all Admin Group entries from the Config sheet (unlimited groups).
 * Supports:
 *  - Multiple rows with ADMIN_GROUP_ID (or ADMIN_GROUP_ID_1, ADMIN_GROUP_ID_2, ..., ADMIN_GROUP_IDS)
 *  - Continuation rows with blank Column A under an ADMIN_GROUP_ID row
 *  - Multiple IDs separated by comma, semicolon, newline, or whitespace inside Column B
 *  - Optional human-readable label/note in Column C
 */
function getAdminGroupEntries(ss) {
  const sheet = ss.getSheetByName(SHEET_CONFIG_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  const data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 3)).getValues();
  const adminFlags = computeAdminRowFlags(data);

  const entries = [];
  const seen = {};
  for (let i = 0; i < data.length; i++) {
    if (!adminFlags[i]) continue;
    const rawVal = String(data[i][1] || '').trim();
    if (!rawVal) continue;
    const note = data[i].length >= 3 ? String(data[i][2] || '').trim() : '';
    const tokens = rawVal.split(/[,;\n\r\s]+/).map(s => s.trim()).filter(Boolean);
    for (let t = 0; t < tokens.length; t++) {
      const id = tokens[t];
      if (!seen[id]) {
        seen[id] = true;
        entries.push({ groupId: id, note: note, rowIndex: i + 1, isMultiCell: tokens.length > 1 });
      }
    }
  }
  return entries;
}

function getAdminGroupIds(ss) {
  return getAdminGroupEntries(ss).map(e => e.groupId);
}

function isAdminGroup(ss, groupId) {
  if (!groupId || groupId === 'NO_GROUP') return false;
  return getAdminGroupIds(ss).indexOf(String(groupId).trim()) !== -1;
}

function addAdminGroupId(ss, groupId, label) {
  const targetId = String(groupId || '').trim();
  const cleanLabel = String(label || '').trim();
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(10000);
  try {
    let sheet = ss.getSheetByName(SHEET_CONFIG_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG_NAME);

    const entries = getAdminGroupEntries(ss);
    const existing = entries.find(e => e.groupId === targetId);
    if (existing) {
      let labelUpdated = false;
      if (cleanLabel && !existing.isMultiCell && existing.note !== cleanLabel) {
        sheet.getRange(existing.rowIndex, 3).setValue(cleanLabel);
        labelUpdated = true;
      }
      return { added: false, labelUpdated: labelUpdated, total: entries.length, ids: entries.map(e => e.groupId) };
    }

    const p = toBangkokParts(new Date());
    const noteToWrite = cleanLabel || `เพิ่มเมื่อ ${p.dd}/${p.mm}/${p.yyyy} ${p.HH}:${p.mm}`;
    sheet.appendRow(['ADMIN_GROUP_ID', targetId, noteToWrite]);

    const updatedIds = entries.map(e => e.groupId).concat([targetId]);
    return { added: true, labelUpdated: false, total: updatedIds.length, ids: updatedIds };
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function removeAdminGroupId(ss, groupId) {
  const targetId = String(groupId || '').trim();
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(10000);
  try {
    const sheet = ss.getSheetByName(SHEET_CONFIG_NAME);
    if (!sheet || sheet.getLastRow() < 1) {
      return { removed: false, total: 0, ids: [] };
    }

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(1, 1, lastRow, 2).getValues();
    const adminFlags = computeAdminRowFlags(data);
    let removed = false;

    for (let i = data.length - 1; i >= 0; i--) {
      if (!adminFlags[i]) continue;
      const rawVal = String(data[i][1] || '').trim();
      if (!rawVal) continue;
      const tokens = rawVal.split(/[,;\n\r\s]+/).map(s => s.trim()).filter(Boolean);
      const remaining = tokens.filter(id => id !== targetId);
      if (remaining.length !== tokens.length) {
        removed = true;
        if (remaining.length === 0) {
          const rowKey = String(data[i][0] || '').trim();
          if (rowKey && i + 1 < data.length && adminFlags[i + 1] && !String(data[i + 1][0] || '').trim()) {
            sheet.getRange(i + 2, 1).setValue(rowKey);
            data[i + 1][0] = rowKey;
          }
          if (sheet.getMaxRows() > 1) {
            sheet.deleteRow(i + 1);
            data.splice(i, 1);
            adminFlags.splice(i, 1);
          } else {
            sheet.getRange(i + 1, 1, 1, Math.max(sheet.getLastColumn(), 3)).clearContent();
            data[i][0] = '';
            data[i][1] = '';
            adminFlags[i] = false;
          }
        } else {
          sheet.getRange(i + 1, 2).setValue(remaining.join(', '));
          data[i][1] = remaining.join(', ');
        }
      }
    }

    const ids = getAdminGroupIds(ss);
    return { removed: removed, total: ids.length, ids: ids };
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

// Legacy aliases kept for backward compatibility
function saveAdminGroupId(ss, groupId) {
  return addAdminGroupId(ss, groupId, '');
}

function getAdminGroupId(ss) {
  const ids = getAdminGroupIds(ss);
  return ids.length > 0 ? ids[0] : null;
}

function parseCleanInt(val) {
  if (val === null || val === undefined || val === '') return 0;
  const parsed = parseInt(val.toString().replace(/,/g, '').trim(), 10);
  return isNaN(parsed) ? 0 : parsed;
}

function getBranchInfo(ss, code) {
  const sheet = ss.getSheetByName(SHEET_BRANCH_NAME);
  if (!sheet) return { found: false, name: code, salesOffice: '-' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toUpperCase() === code) {
      return { found: true, name: data[i][1], salesOffice: data[i][2] || '-' };
    }
  }
  return { found: false, name: code, salesOffice: '-' };
}

function calculateTimeDiff(start, end) {
  if (!start || !end || start === '-' || end === '-') return '-';
  const s = start.split(':').map(Number);
  const e = end.split(':').map(Number);
  let d = (e[0] * 60 + e[1]) - (s[0] * 60 + s[1]);
  return d < 0 ? d + 1440 : d;
}

function formatDuration(startStr, endStr) {
  if (!startStr || !endStr || startStr === '-' || endStr === '-') return '-';
  const s = startStr.split(':').map(Number);
  const e = endStr.split(':').map(Number);
  let diffMins = (e[0] * 60 + e[1]) - (s[0] * 60 + s[1]);
  if (diffMins < 0) diffMins += 1440;
  const h = Math.floor(diffMins / 60);
  const m = diffMins % 60;
  return `${h}:${m.toString().padStart(2, '0')}:00`;
}

const BKK_DTF = new Intl.DateTimeFormat('en-GB', {
  timeZone: 'Asia/Bangkok',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  hourCycle: 'h23'
});

function toBangkokParts(dateObj) {
  // Uses IANA Asia/Bangkok tzdata so both 1899-12-30 time-only cells (LMT UTC+06:42:04)
  // and modern dates (UTC+07:00) match Google Sheets display time exactly with zero Java RPC calls.
  const parts = BKK_DTF.formatToParts(dateObj);
  let dd = '00', mm = '00', yyyy = '0000', HH = '00', MI = '00';
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (p.type === 'day') dd = p.value;
    else if (p.type === 'month') mm = p.value;
    else if (p.type === 'year') yyyy = p.value;
    else if (p.type === 'hour') HH = p.value === '24' ? '00' : p.value;
    else if (p.type === 'minute') MI = p.value;
  }
  return {
    dd: dd,
    mm: mm,
    yyyy: yyyy,
    yy: yyyy.slice(-2),
    HH: HH,
    MI: MI
  };
}

function formatTimeOnly(v) {
  if (!v) return '-';
  if (v instanceof Date) {
    const p = toBangkokParts(v);
    return `${p.HH}:${p.MI}`;
  }
  return v.toString().substring(0, 5);
}

function formatTimeForWeb(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const p = toBangkokParts(v);
    return `${p.HH}:${p.MI}`;
  }
  const s = v.toString().trim();
  if (!s || s === '-') return '';
  return s.substring(0, 5);
}

function formatDateForComparison(v) {
  if (v instanceof Date) {
    const p = toBangkokParts(v);
    return `${p.dd}/${p.mm}/${p.yy}`;
  }
  const str = String(v || '').trim();
  const m = str.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})/);
  if (m) {
    const dd = m[1].padStart(2, '0');
    const mm = m[2].padStart(2, '0');
    let y = parseInt(m[3], 10);
    if (y >= 2500) y -= 543;
    else if (y >= 50 && y <= 99) y -= 43;
    const yy = String(y % 100).padStart(2, '0');
    return `${dd}/${mm}/${yy}`;
  }
  return str;
}

function formatDateFullYear(v) {
  if (!v) return '-';
  if (v instanceof Date) {
    const p = toBangkokParts(v);
    return `${p.dd}/${p.mm}/${p.yyyy}`;
  }
  return v.toString().split(' ')[0];
}

function getThaiDateTime(d) {
  const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const p = toBangkokParts(d);
  return `${parseInt(p.dd, 10)} ${thaiMonths[parseInt(p.mm, 10) - 1]} ${parseInt(p.yyyy, 10) + 543} ${p.HH}:${p.MI} น.`;
}

function replyText(replyToken, message, ss) {
  sendToLine('reply', { replyToken: replyToken, messages: [{ type: 'text', text: message }] }, ss);
}

function sendToLine(type, payload, ss) {
  const token = getProperty('CHANNEL_ACCESS_TOKEN', '');
  if (!token) {
    logSystem('ERROR', 'MISSING_LINE_TOKEN', 'ยังไม่ได้ตั้งค่า Script Property: CHANNEL_ACCESS_TOKEN', '', ss);
    return;
  }
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + type, {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 400) {
      logSystem('ERROR', 'REPLY_ERROR', `LINE API HTTP ${code}`, res.getContentText(), ss);
    }
  } catch (e) {
    logSystem('ERROR', 'REPLY_CRASH', 'ส่งข้อความกลับล้มเหลว', e.message, ss);
  }
}

function verifyLineIdToken(idToken) {
  const channelId = getProperty('LINE_CHANNEL_ID', '');
  if (!channelId) return null;
  try {
    const response = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
      method: 'post',
      payload: { id_token: idToken, client_id: channelId },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return null;
    const claims = JSON.parse(response.getContentText());
    if (!claims || !claims.sub || String(claims.aud) !== String(channelId)) return null;
    if (!claims.exp || claims.exp * 1000 <= Date.now()) return null;
    return { sub: String(claims.sub), name: claims.name ? String(claims.name) : '' };
  } catch (_e) {
    return null;
  }
}

// ==========================================
// 9. WEB APP HELPERS (DSO-liff)
// ==========================================
function getDataFromSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonResponse({ status: 'success', data: [] });
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  const formattedData = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (!row[1]) continue;
    formattedData.push({
      rowIndex: i + 2,
      timestamp: formatDateFull(row[0]),
      branch: row[1],
      targetDate: formatDateForComparison(row[2]),
      initialAmount: row[3],
      timeBill: formatTimeForWeb(row[5]),
      currentAmount: row[6] || row[3],
      timeConf: formatTimeForWeb(row[8]),
      timeShip: formatTimeForWeb(row[12]),
      shipmentAmount: row[13],
      diff2: row[14],
      sequence: row[15],
      editCount: parseCleanInt(row[16]),
      note: row[17] ? String(row[17]).trim() : ''
    });
  }
  return jsonResponse({ status: 'success', data: formattedData });
}

function formatDateFull(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    const p = toBangkokParts(val);
    return `${p.dd}/${p.mm}/${p.yy} ${p.HH}:${p.MI}`;
  }
  return val.toString();
}

function getBranchesList() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BRANCH_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonResponse({ status: 'success', data: [] });
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const formattedList = data
    .filter(row => row[0])
    .map(row => ({ id: row[0], name: row[1] }));
  return jsonResponse({ status: 'success', data: formattedList });
}
