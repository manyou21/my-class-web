// ============================================
// 📘 ระบบเว็บการบ้าน & เก็บเงินห้อง + Redeem Code
// ============================================
const SPREADSHEET_ID = '1cT-N8AHw613xstQ7FUNdOk-QnavT9TRsZh1SiJJfYWA';
const DRIVE_FOLDER_ID = '1bv_ZpRpY5diAW9Sd_7Hkl8UCTeHU8YCn';
const QR_FOLDER_ID = '190MePEweP8hVXejCx5yCU2B2KUbB_Xey';

var ss; // Global spreadsheet instance

const SHEETS = {
  USERS: 'Users',
  HOMEWORK: 'Homework',
  HOMEWORK_STATUS: 'HomeworkStatus',
  TREASURY: 'Treasury',
  TREASURY_PAYMENTS: 'TreasuryPayments',
  ABSENCE: 'Absence',
  STUDENT_CODES: 'StudentCodes',
  LEAVE_REQUESTS: 'LeaveRequests',
  REDEEM_CODES: 'RedeemCodes',
  TIMETABLE: 'Timetable',
  SEAT_META: 'SeatMeta',
  SEAT_BOOKINGS: 'SeatBookings',
  SEAT_EDIT_CODES: 'SeatEditCodes',
  SEAT_EDIT_SESSIONS: 'SeatEditSessions',
  LOANS: 'Loans'
};

// ============================================
// ⚡ PERFORMANCE & CACHE CONFIGURATION
// ============================================
const CACHE_UPDATE_KEY = 'dashboard_last_update_v1';
const CACHE_COUNTS_KEY = 'dashboard_counts_v1';
const CACHE_DASHBOARD_KEY = 'dashboard_data_v1';
const CACHE_TTL = 300;           // 5 นาที สำหรับ counts/lastUpdate
const CACHE_DASHBOARD_TTL = 60;  // 60 วินาที สำหรับ full data (invalidate ได้บ่อยกว่า)

function getSpreadsheet() {
  if (ss) return ss;
  ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss;
}

function getSheet_(name) {
  const activeSs = getSpreadsheet();
  let sheet = activeSs.getSheetByName(name);
  if (!sheet) {
    ensureSheetsExist();
    sheet = activeSs.getSheetByName(name);
  }
  return sheet;
}

function getSheetFromSs_(activeSs, name) {
  let sheet = activeSs.getSheetByName(name);
  if (!sheet) {
    ensureSheetsExist();
    sheet = activeSs.getSheetByName(name);
  }
  return sheet;
}

function clearDashboardCaches_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([CACHE_UPDATE_KEY, CACHE_COUNTS_KEY, CACHE_DASHBOARD_KEY]);
  } catch (e) {
    Logger.log('Cache clear error: ' + e.toString());
  }
}

const STUDENTS = [
  { no: 1,  code: '47535', name: 'คีตภัทร ชัยปรัชย์' },    { no: 2,  code: '47550', name: 'จิรายุ วงกต' },
  { no: 3,  code: '47559', name: 'ชนาธิป ลาสิบสี' },        { no: 4,  code: '47594', name: 'ณัฐกรณ์ ปวงจันทร์' },
  { no: 5,  code: '47611', name: 'เดชณรงค์ มนตรี' },        { no: 6,  code: '47627', name: 'เทวาพิทักษ์ วังแก้ว' },
  { no: 7,  code: '47648', name: 'ธนวัฒน์ ทะฤาษี' },        { no: 8,  code: '47654', name: 'ธนาธิป เชียงว้อง' },
  { no: 9,  code: '47655', name: 'ธนาธิป อัฐวงศ์' },        { no: 10, code: '47660', name: 'ธีรวัฒน์ ค้ำชู' },
  { no: 11, code: '47670', name: 'นรภัทร อุระวัง' },        { no: 12, code: '47673', name: 'นวพล แก้วบุญมา' },
  { no: 13, code: '47681', name: 'เนติภูมิ หงษ์กัน' },      { no: 14, code: '47685', name: 'ปรมินทร์ พิยะ' },
  { no: 15, code: '47691', name: 'ปริญญา ประสม' },          { no: 16, code: '47700', name: 'ปุณยกร เขื่อนหนึ่ง' },
  { no: 17, code: '47704', name: 'พงศธร ไชยเมฆา' },        { no: 18, code: '47718', name: 'พาทิศ ปิมลื้อ' },
  { no: 19, code: '47721', name: 'พิชรัช พันธากุล' },       { no: 20, code: '47733', name: 'ภัทรบดี เรือนเครือ' },
  { no: 21, code: '47793', name: 'ศรณ์คุณัชญ์ หนองกาวี' }, { no: 22, code: '47825', name: 'อัฑฒกร ทุ่งสง' },
  { no: 23, code: '48811', name: 'ณัฐภัทร บัวบานแย้ม' },   { no: 24, code: '47836', name: 'กมลพร นูนเมือง' },
  { no: 25, code: '47838', name: 'กมลวรรณ กลั่นสกุล' },    { no: 26, code: '47850', name: 'กันต์ฐณิชา ภักดีรัตนมิตร' },
  { no: 27, code: '47883', name: 'ณัฐณิชา เวียงนาค' },     { no: 28, code: '47890', name: 'ณัฐภัทร ใจยวญ' },
  { no: 29, code: '47897', name: 'ณิชนันทน์ พลอยเพ็ชร' },  { no: 30, code: '47905', name: 'ธนัชชา กิติธันยพงศ์' },
  { no: 31, code: '47921', name: 'นวพร เครือพาน' },         { no: 32, code: '47937', name: 'ปวริศา นางวงศ์' },
  { no: 33, code: '47970', name: 'พิมลวรรณ สุขทรัพย์' },   { no: 34, code: '47989', name: 'มลธิดาภรณ์ เปลี่ยมแพร' },
  { no: 35, code: '48000', name: 'วชิรญาย์ ฆ้องคำ' },      { no: 36, code: '48004', name: 'วราศินี ปั้นแพทย์' },
  { no: 37, code: '48007', name: 'ศจีนาฏ วิทยา' },         { no: 38, code: '48011', name: 'ศศิวิมล ถานะวุฒิพงศ์' },
  { no: 39, code: '48012', name: 'ศิรประภา สุตาถี' },       { no: 40, code: '48019', name: 'สิริชญา ใจนวล' }
];

const SUBJECTS = [
  'ไทยหลัก', 'ไทยเสริม', 'คณิตหลัก', 'คณิตเสริม', 'วิทย์หลัก', 'วิทย์เสริม',
  'อังกฤษหลัก', 'อังกฤษเสริม Joshua', 'อังกฤษเสริม จิรารัตน์', 'IS', 'ประวัติ', 'สังคม',
  'ป้องกันการทุจริต', 'วิทยาการคำนวณ', 'มัลติมีเดีย', 'แนะแนว', 'นาฏศิลป์', 'ทัศนศิลป์',
  'การงาน', 'สุขศึกษา', 'พลศึกษา', 'อื่นๆ'
];

const ROLES = {
  OWNER:        { name: 'เจ้าของเว็บ',    canManageHomework: true,  canManageTreasury: true,  canApproveLeave: true,  canManageCodes: true  },
  CLASS_LEADER: { name: 'หัวหน้าห้อง',    canManageHomework: true,  canManageTreasury: true,  canApproveLeave: false, canManageCodes: false },
  SECRETARY:    { name: 'เลขานุการ',      canManageHomework: true,  canManageTreasury: true,  canApproveLeave: false, canManageCodes: false },
  TREASURER:    { name: 'เหรัญญิก',       canManageHomework: true,  canManageTreasury: true,  canApproveLeave: false, canManageCodes: false },
  TEACHER:      { name: 'ครูที่ปรึกษา',   canManageHomework: false, canManageTreasury: false, canApproveLeave: true,  canManageCodes: false },
  STUDENT:      { name: 'นักเรียน',       canManageHomework: false, canManageTreasury: false, canApproveLeave: false, canManageCodes: false },
  GUEST:        { name: 'ผู้ใช้ชั่วคราว', canManageHomework: false, canManageTreasury: false, canApproveLeave: false, canManageCodes: false }
};

const ROLE_MAPPING = {
  'พาทิศ ปิมลื้อ':  'OWNER',
  'อัฑฒกร ทุ่งสง':  'CLASS_LEADER',
  'ศิรประภา สุตาถี': 'SECRETARY',
  'สิริชญา ใจนวล':   'TREASURER',
  'ปทิตตา พิจารณ์':  'TEACHER',
  'กีรติ บุญทวี':    'TEACHER'
};

const SUBJECT_COLORS = [
  '#fd959f', '#f598b8', '#e3aded', '#D1C4E9', '#C5CAE9', '#BBDEFB',
  '#B2EBF2', '#B2DFDB', '#C8E6C9', '#DCEDC8', '#F0F4C3', '#FFF9C4',
  '#FFECB3', '#FFE0B2', '#FFCCBC', '#e1ff9c', '#a0daf3', '#f2cbf8'
];

// ============================================
// 🔥 API Endpoint for GitHub Pages (CORS)
// ============================================

// Whitelist: เฉพาะ function เหล่านี้เท่านั้นที่เรียกจาก API ได้
const ALLOWED_ACTIONS = new Set([
  'loginUser', 'registerUser', 'changePassword', 'getPasswordHint',
  'getStudents', 'getSubjects',
  'addHomework', 'getHomework', 'updateHomeworkStatus', 'deleteHomework',
  'addTreasuryItem', 'getTreasuryItems', 'updatePayment', 'deleteTreasuryItem',
  'submitLeaveRequest', 'getLeaveRequests', 'confirmLeaveRequest', 'updateLeaveStatus',
  'getLastUpdate', 'getDashboardData', 'getCounts',
  'generateCode', 'redeemCode',
  'clearAllWebsiteData',
  'createGuestAccount', 'deleteGuestAccount',
  'getTimetable', 'setTimetable',
  'seatGetSnapshot', 'seatSetBookingWindow', 'seatSetFrontBand',
  'seatSaveLayout', 'seatBook', 'seatCancelBooking',
  'seatCreateEditCode', 'seatValidateEditCode', 'seatListEditCodes',
  'seatRevokeEditCode', 'seatRevokeSession', 'seatListActiveSessions',
  'addLoan', 'getLoans', 'updateLoanStatus', 'deleteLoan'
]);

function doPost(e) {
  try {
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    // ป้องกัน payload ใหญ่เกินไป (max 2MB)
    const raw = e.postData.contents;
    if (!raw || raw.length > 2 * 1024 * 1024) {
      output.setContent(JSON.stringify({ status: 'error', message: 'Request too large' }));
      return output;
    }

    const data = JSON.parse(raw);
    const action = String(data.action || '');
    const args = Array.isArray(data.args) ? data.args : [];

    // Whitelist check — ป้องกันการเรียก function ที่ไม่ได้อนุญาต
    if (!ALLOWED_ACTIONS.has(action)) {
      output.setContent(JSON.stringify({ status: 'error', message: 'Action not allowed' }));
      return output;
    }

    // จำกัดจำนวน args ป้องกัน prototype pollution
    if (args.length > 10) {
      output.setContent(JSON.stringify({ status: 'error', message: 'Too many arguments' }));
      return output;
    }

    // Dispatch table — ปลอดภัยกว่า eval
    const DISPATCH = {
      loginUser, registerUser, changePassword, getPasswordHint,
      getStudents, getSubjects,
      addHomework, getHomework, updateHomeworkStatus, deleteHomework,
      addTreasuryItem, getTreasuryItems, updatePayment, deleteTreasuryItem,
      submitLeaveRequest, getLeaveRequests, confirmLeaveRequest, updateLeaveStatus,
      getLastUpdate, getDashboardData, getCounts,
      generateCode, redeemCode,
      clearAllWebsiteData,
      createGuestAccount, deleteGuestAccount,
      getTimetable, setTimetable,
      seatGetSnapshot, seatSetBookingWindow, seatSetFrontBand,
      seatSaveLayout, seatBook, seatCancelBooking,
      seatCreateEditCode, seatValidateEditCode, seatListEditCodes,
      seatRevokeEditCode, seatRevokeSession, seatListActiveSessions,
      addLoan, getLoans, updateLoanStatus, deleteLoan
    };
    if (!DISPATCH[action]) {
      output.setContent(JSON.stringify({ status: 'error', message: 'Action not found' }));
      return output;
    }
    const result = DISPATCH[action].apply(null, args);
    output.setContent(JSON.stringify({ status: 'success', result: result }));
    return output;
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Server error: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput('API is running. Please use POST requests.');
}

// ============================================
// 🔧 HELPER FUNCTIONS
// ============================================

// Hash password พร้อม salt ป้องกัน rainbow table attack
// salt คงที่ต่อระบบ (เพิ่มความปลอดภัยกว่า plain SHA-256)
const PW_SALT = 'CLS_SALT_2025_xK9#mP';

function byteDigestToHex_(digest) {
  var hexString = '';
  for (var i = 0; i < digest.length; i++) {
    hexString += ('0' + (digest[i] & 0xFF).toString(16)).slice(-2);
  }
  return hexString;
}

function hashPassword(p) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, PW_SALT + p + PW_SALT);
  return byteDigestToHex_(digest);
}

// Sanitize string input — ตัด whitespace และจำกัดความยาว
function sanitizeStr(val, maxLen) {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  return maxLen ? s.slice(0, maxLen) : s;
}

// ตรวจสอบว่า userId มีรูปแบบถูกต้อง (UUID หรือ GUEST_XXXX)
function isValidUserId(id) {
  if (!id) return false;
  const s = String(id);
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s)
      || /^GUEST_[A-Z0-9]{8}$/i.test(s);
}

// Rate limiting สำหรับ login — ป้องกัน brute force
// เก็บจำนวนครั้งที่ login ผิดใน ScriptProperties
const RATE_LIMIT_MAX = 10;       // ผิดได้สูงสุด 10 ครั้ง
const RATE_LIMIT_WINDOW_MS = 15 * 60 * 1000; // ใน 15 นาที

function checkLoginRateLimit(identifier) {
  try {
    const props = PropertiesService.getScriptProperties();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(identifier));
    const key = 'rl_' + byteDigestToHex_(digest).slice(0, 16);
    const raw = props.getProperty(key);
    const now = Date.now();
    let record = raw ? JSON.parse(raw) : { count: 0, windowStart: now };
    if (now - record.windowStart > RATE_LIMIT_WINDOW_MS) {
      record = { count: 0, windowStart: now };
    }
    if (record.count >= RATE_LIMIT_MAX) {
      const waitMin = Math.ceil((RATE_LIMIT_WINDOW_MS - (now - record.windowStart)) / 60000);
      return { blocked: true, message: `พยายามเข้าสู่ระบบผิดพลาดบ่อยเกินไป กรุณารอ ${waitMin} นาที` };
    }
    return { blocked: false, record: record, key: key };
  } catch (e) {
    return { blocked: false }; // ถ้า rate limit พัง ให้ผ่านไปก่อน
  }
}

function recordLoginFailure(key, record) {
  try {
    const props = PropertiesService.getScriptProperties();
    record.count = (record.count || 0) + 1;
    props.setProperty(key, JSON.stringify(record));
  } catch (e) {}
}

function clearLoginRateLimit(key) {
  try {
    PropertiesService.getScriptProperties().deleteProperty(key);
  } catch (e) {}
}

function validatePassword(p) {
  if (!p || p.length < 4) return { valid: false, m: 'รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร' };
  return { valid: true };
}

function getSubjectColor(subjectName) {
  if (!subjectName) return '#EEEEEE';
  let hash = 0;
  for (let i = 0; i < subjectName.length; i++) {
    hash = subjectName.charCodeAt(i) + ((hash << 5) - hash);
  }
  return SUBJECT_COLORS[Math.abs(hash) % SUBJECT_COLORS.length];
}

function uploadImageToDrive(base64Data, fileName) {
  try {
    if (!base64Data) return null;
    const split = base64Data.split(',');
    const type = split[0].split(';')[0].split(':')[1];
    const data = Utilities.base64Decode(split[1]);
    const blob = Utilities.newBlob(data, type, fileName);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch (e) {
    console.error('Drive Upload Error: ' + e.toString());
    return null;
  }
}

function uploadToDrive(base64Data, fileName, folderName = "Classroom_Images") {
  const folder = getOrCreateFolder(folderName);
  const contentType = base64Data.substring(5, base64Data.indexOf(';'));
  const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(SpreadsheetApp.Access.ANYONE_WITH_LINK, SpreadsheetApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function convertDriveToDirect(url) {
  if (!url) return '';
  const m = url.match(/[-\w]{25,}/);
  return m ? 'https://drive.google.com/uc?export=view&id=' + m[0] : url;
}

function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const existingSheets = {};
  ss.getSheets().forEach(function(s) {
    existingSheets[s.getName()] = s;
  });
  
  let createdAny = false;

  if (!existingSheets[SHEETS.USERS]) {
    ss.insertSheet(SHEETS.USERS).appendRow(['ID','DisplayName','Email','PasswordHash','Hint','StudentNo','Role','CreatedAt','LastLogin','TempRoleExpiry','HwCredits']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.STUDENT_CODES]) {
    const sc = ss.insertSheet(SHEETS.STUDENT_CODES);
    sc.appendRow(['StudentNo','StudentCode','StudentName','IsRegistered']);
    STUDENTS.forEach(s => sc.appendRow([s.no, s.code, s.name, false]));
    createdAny = true;
  }
  if (!existingSheets[SHEETS.HOMEWORK]) {
    ss.insertSheet(SHEETS.HOMEWORK).appendRow(['ID','Subject','Description','AssignedDate','DueDate','NoDueDate','CreatedBy','CreatedAt','SubjectColor']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.HOMEWORK_STATUS]) {
    ss.insertSheet(SHEETS.HOMEWORK_STATUS).appendRow(['HomeworkID','StudentNo','Status','ImagePath','CompletedAt','Notes','SubjectColor']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.TREASURY]) {
    ss.insertSheet(SHEETS.TREASURY).appendRow(['ID','Title','AmountPerPerson','CreatedBy','CreatedAt','Status','Color']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.TREASURY_PAYMENTS]) {
    ss.insertSheet(SHEETS.TREASURY_PAYMENTS).appendRow(['TreasuryID','StudentNo','AmountPaid','PaidAt','Notes','Color','ReceiptNo']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.LEAVE_REQUESTS]) {
    ss.insertSheet(SHEETS.LEAVE_REQUESTS).appendRow(['ID','StudentNo','StudentName','Type','Date','Reason','Status','ProofImage','Confirmed','Timestamp']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.REDEEM_CODES]) {
    ss.insertSheet(SHEETS.REDEEM_CODES).appendRow(['Code','ActionType','Value','Details','MaxUses','UsesCount','CreatedBy','CreatedAt','QrFileId','QrUrl','CreatedAtUTC']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.TIMETABLE]) {
    const tt = ss.insertSheet(SHEETS.TIMETABLE);
    tt.appendRow(['ImageUrl', 'LinkUrl', 'UpdatedAt', 'UpdatedBy']);
    tt.appendRow(['', '', '', '']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.SEAT_META]) {
    const sm = ss.insertSheet(SHEETS.SEAT_META);
    sm.appendRow(['LayoutJSON', 'BookingStart', 'BookingEnd', 'FrontBandRows', 'Version', 'UpdatedAt']);
    const defaultLayout = JSON.stringify({
      grid: { cols: 22, rows: 16, cell: 24 },
      seats: [],
      frontBand: 2
    });
    sm.appendRow([defaultLayout, '', '', 2, 0, new Date().toISOString()]);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.SEAT_BOOKINGS]) {
    ss.insertSheet(SHEETS.SEAT_BOOKINGS).appendRow(['SeatId', 'TargetStudentNo', 'TargetStudentName', 'BookedByUserId', 'CreatedAt']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.SEAT_EDIT_CODES]) {
    ss.insertSheet(SHEETS.SEAT_EDIT_CODES).appendRow(['Id', 'CodeHash', 'ExpiresAt', 'CreatedByUserId', 'Label', 'Revoked']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.SEAT_EDIT_SESSIONS]) {
    ss.insertSheet(SHEETS.SEAT_EDIT_SESSIONS).appendRow(['Token', 'CodeId', 'ExpiresAt']);
    createdAny = true;
  }
  if (!existingSheets[SHEETS.LOANS]) {
    ss.insertSheet(SHEETS.LOANS).appendRow(['ID', 'BorrowerNo', 'LenderNo', 'Amount', 'Note', 'ProofUrl', 'Status', 'CreatedAt', 'CreatedBy']);
    createdAny = true;
  }

  if (createdAny) {
    formatSheets();
  }
  return 'OK';
}

function formatSheets(specificSheetName = null) {
  try {
    const ss = getSpreadsheet();
    const sheets = specificSheetName ? [ss.getSheetByName(specificSheetName)] : ss.getSheets();
    
    // สร้าง Color Map สำหรับเติมข้อมูลย้อนหลัง
    const colorMap = {};
    try {
      const hwSheet = ss.getSheetByName(SHEETS.HOMEWORK);
      if (hwSheet) {
        const hwData = hwSheet.getDataRange().getValues();
        hwData.slice(1).forEach(r => { if(r[0]) colorMap[r[0]] = r[8]; });
      }
      const trSheet = ss.getSheetByName(SHEETS.TREASURY);
      if (trSheet) {
        const trData = trSheet.getDataRange().getValues();
        trData.slice(1).forEach(r => { if(r[0]) colorMap[r[0]] = r[6]; });
      }
    } catch(e) { Logger.log('ColorMap Error: ' + e.message); }

    sheets.forEach(sheet => {
      if (!sheet) return;
      const name = sheet.getName();
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      if (lastCol === 0) return;
      
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setBackground('#1e293b') // Slate 800
                 .setFontColor('#ffffff')
                 .setFontWeight('bold')
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle')
                 .setFontSize(10)
                 .setFontFamily('Prompt');
      
      try { sheet.setFrozenRows(1); } catch(e) {}
      
      // 2. ซ่อนคอลัมน์ที่เป็นข้อมูลทางเทคนิค
      const headers = headerRange.getValues()[0];
      const techKeywords = ['ID', 'JSON', 'HASH', 'TOKEN', 'UUID'];
      
      headers.forEach((header, index) => {
        const hStr = String(header || '').toUpperCase();
        const shouldHide = techKeywords.some(key => hStr.includes(key)) || (index === 0 && (hStr.includes('ID') || hStr.includes('CODE') || hStr.length < 4) && name !== SHEETS.STUDENT_CODES);
        
        try {
          const isHidden = sheet.isColumnHiddenByUser(index + 1);
          if (shouldHide && !isHidden) {
            sheet.hideColumns(index + 1);
          } else if (!shouldHide && isHidden) {
            sheet.showColumns(index + 1);
          }
        } catch(e) {}
      });

      // 3. ปรับความกว้างคอลัมน์ (ข้าม autoResizeColumn ถ้าเกิดรันแบบเจาะจงแผ่นงานเพื่อให้ทำงานเร็วขึ้น)
      for (let i = 1; i <= lastCol; i++) {
        try {
          if (sheet.isColumnHiddenByUser(i)) continue;
          if (!specificSheetName) { // รันเฉพาะการฟอร์แมตตั้งต้นทั้งหมด
            sheet.autoResizeColumn(i);
          }
          let width = sheet.getColumnWidth(i);
          if (width > 300) sheet.setColumnWidth(i, 300);
          if (width < 80) sheet.setColumnWidth(i, 80);
        } catch(e) {}
      }

      // 4. จัดกลุ่มและสี
      if (lastRow > 1) {
        const contentRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        contentRange.setFontFamily('Prompt').setFontSize(9).setVerticalAlignment('middle');
        
        const colorSheets = [SHEETS.HOMEWORK, SHEETS.HOMEWORK_STATUS, SHEETS.TREASURY, SHEETS.TREASURY_PAYMENTS];
        if (colorSheets.includes(name)) {
          const data = contentRange.getValues();
          let colorColIdx = -1;
          if (name === SHEETS.HOMEWORK) colorColIdx = 8;
          if (name === SHEETS.HOMEWORK_STATUS) colorColIdx = 6;
          if (name === SHEETS.TREASURY) colorColIdx = 6;
          if (name === SHEETS.TREASURY_PAYMENTS) colorColIdx = 5;

          if (colorColIdx !== -1) {
            const bgColors = data.map((row, i) => {
              let rowColor = row[colorColIdx];
              if (!rowColor || !/^#[0-9A-F]{6}$/i.test(rowColor)) {
                rowColor = colorMap[row[0]] || '#ffffff';
                if (rowColor !== '#ffffff' && colorColIdx < lastCol) try { sheet.getRange(i + 2, colorColIdx + 1).setValue(rowColor); } catch(e) {}
              }
              return Array(lastCol).fill(rowColor);
            });
            try { contentRange.getBandings().forEach(b => b.remove()); } catch(e) {}; try { contentRange.setBackgrounds(bgColors); } catch(e) { contentRange.setBackground('#ffffff'); }
          }
        } else {
          // ล้าง Banding เก่าออกก่อนเพื่อป้องกัน Error "Banding already exists"
          try { contentRange.getBandings().forEach(b => b.remove()); } catch(e) {}
          const rowColors = [];
          for (let r = 2; r <= lastRow; r++) {
            rowColors.push(Array(lastCol).fill(r % 2 === 0 ? '#f8fafc' : '#ffffff'));
          }
          if (rowColors.length > 0) contentRange.setBackgrounds(rowColors);
        }
      }
    });

    if (!specificSheetName || specificSheetName === SHEETS.USERS) {
      const userSheet = ss.getSheetByName(SHEETS.USERS);
      if (userSheet && userSheet.getLastRow() > 1) {
        try { userSheet.getRange(2, 4, userSheet.getLastRow() - 1, 1).setFontColor('#94a3b8').setFontSize(8); } catch(e) {}
      }
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('Critical Error in formatSheets: ' + err.message);
  }
}
// ============================================
// 🔐 AUTHENTICATION
// ============================================
function registerUser(name, email, pw, cpw, code, hint) {
  try {
    const uSheet = getSheet_(SHEETS.USERS); 
    const cSheet = getSheet_(SHEETS.STUDENT_CODES);

    // Sanitize inputs
    const cleanName  = sanitizeStr(name, 100);
    const cleanEmail = sanitizeStr(email, 200);
    const cleanCode  = sanitizeStr(code, 20);
    const cleanHint  = sanitizeStr(hint, 200);

    if (!cleanName || !pw) return { success: false, message: 'กรอกข้อมูลไม่ครบ' }; 
    if (pw !== cpw) return { success: false, message: 'รหัสผ่านไม่ตรงกัน' }; 
    const vP = validatePassword(pw); 
    if (!vP.valid) return { success: false, message: vP.m };

    // ตรวจ email format ถ้ากรอกมา
    if (cleanEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail)) {
      return { success: false, message: 'รูปแบบ Email ไม่ถูกต้อง' };
    }

    const uData = uSheet.getDataRange().getValues(); 
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === cleanName) return { success: false, message: 'มีผู้ใช้ชื่อนี้แล้ว' }; 
    }
    
    let role = 'STUDENT', studentNo = null;
    if (!cleanCode) { 
      if (ROLE_MAPPING[cleanName]) role = ROLE_MAPPING[cleanName]; 
      else return { success: false, message: 'ชื่อนี้ไม่อยู่ในทะเบียน กรุณากรอกรหัสนักเรียน' }; 
    } else {
      // ตรวจ format รหัสนักเรียน (5 หลักตัวเลข)
      if (!/^\d{5}$/.test(cleanCode)) return { success: false, message: 'รหัสนักเรียนต้องเป็นตัวเลข 5 หลัก' };

      const cData = cSheet.getDataRange().getValues(); 
      let found = false;
      for (let i = 1; i < cData.length; i++) { 
        if (String(cData[i][1]).trim() === cleanCode) { 
          if (cData[i][3] === true) return { success: false, message: 'รหัสนักเรียนนี้ถูกใช้งานแล้ว' }; 
          if (String(cData[i][2]).trim() !== cleanName.trim()) return { success: false, message: 'ชื่อไม่ตรงกับรหัสนักเรียน' }; 
          studentNo = cData[i][0]; 
          cSheet.getRange(i + 1, 4).setValue(true); 
          found = true; 
          if (ROLE_MAPPING[cleanName]) role = ROLE_MAPPING[cleanName]; 
          break; 
        } 
      }
      if (!found) return { success: false, message: 'ไม่พบรหัสนักเรียนนี้' }; 
    }
    
    uSheet.appendRow([Utilities.getUuid(), cleanName, cleanEmail, hashPassword(pw), cleanHint, studentNo, role, new Date(), new Date(), '', 0]);
    return { success: true, message: 'สมัครสมาชิกสำเร็จ! กรุณา Login' };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด' }; 
  }
}

// hash แบบเก่า (ไม่มี salt) — ใช้สำหรับ migration เท่านั้น
function hashPasswordLegacy_(p) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p);
  return byteDigestToHex_(digest);
}

function loginUser(id, pw) {
  try {
    const uSheet = getSheet_(SHEETS.USERS);
    if (!uSheet) return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };

    // Sanitize inputs
    const cleanId = sanitizeStr(id, 200);
    if (!cleanId || !pw) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };

    // Rate limit check
    const rl = checkLoginRateLimit(cleanId);
    if (rl.blocked) return { success: false, message: rl.message };
    
    const uData = uSheet.getDataRange().getValues(); 
    const pHashNew    = hashPassword(pw);          // hash ใหม่ (มี salt)
    const pHashLegacy = hashPasswordLegacy_(pw);   // hash เก่า (ไม่มี salt)
    
    for (let i = 1; i < uData.length; i++) {
      const row = uData[i];
      const nameMatch  = row[1] === cleanId;
      const emailMatch = row[2] === cleanId && row[2] !== '';
      if (!nameMatch && !emailMatch) continue;

      const storedHash = row[3];
      const matchNew    = storedHash === pHashNew;
      const matchLegacy = storedHash === pHashLegacy;

      if (matchNew || matchLegacy) {
        // Login สำเร็จ — ล้าง rate limit
        if (rl.key) clearLoginRateLimit(rl.key);

        // Auto-migrate: ถ้ายังเป็น hash เก่า → อัปเดตเป็น hash ใหม่ทันที
        if (matchLegacy && !matchNew) {
          uSheet.getRange(i + 1, 4).setValue(pHashNew);
        }

        let rK = row[6];
        if (row.length > 9 && row[9]) {
          const expiry = new Date(row[9]);
          if (expiry > new Date()) rK = 'TEACHER';
          else uSheet.getRange(i + 1, 10).setValue('');
        }
        const credits = (row.length > 10 && row[10]) ? Number(row[10]) : 0;
        const roleObj = ROLES[rK] || ROLES.STUDENT;
        uSheet.getRange(i + 1, 9).setValue(new Date());
        return {
          success: true,
          user: {
            id: row[0],
            displayName: row[1],
            email: row[2],
            role: roleObj.name,
            roleKey: rK,
            studentNo: row[5],
            canManageHomework: roleObj.canManageHomework || credits > 0,
            canManageTreasury: roleObj.canManageTreasury,
            canApproveLeave: roleObj.canApproveLeave,
            canManageCodes: roleObj.canManageCodes,
            hwCredits: credits
          }
        };
      } 
    }

    // Login ล้มเหลว — บันทึก rate limit
    if (rl.record && rl.key) recordLoginFailure(rl.key, rl.record);
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) { 
    return { success: false, message: 'Server Error: ' + e.toString() }; 
  }
}

function changePassword(uid, curr, newP, conf) {
  try {
    const uSheet = getSheet_(SHEETS.USERS); 
    const uData = uSheet.getDataRange().getValues();

    const cleanUid = sanitizeStr(uid, 100);
    if (!isValidUserId(cleanUid)) return { success: false, message: 'ข้อมูลผู้ใช้ไม่ถูกต้อง' };

    for (let i = 1; i < uData.length; i++) {
      if (uData[i][0] === cleanUid) {
        const storedHash = uData[i][3];
        // รองรับทั้ง hash เก่าและใหม่
        const currMatchNew    = storedHash === hashPassword(curr);
        const currMatchLegacy = storedHash === hashPasswordLegacy_(curr);
        if (!currMatchNew && !currMatchLegacy) return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
        if (newP !== conf) return { success: false, message: 'รหัสผ่านใหม่ไม่ตรงกัน' };
        const v = validatePassword(newP); 
        if (!v.valid) return { success: false, message: v.m };
        uSheet.getRange(i + 1, 4).setValue(hashPassword(newP)); // บันทึกเป็น hash ใหม่เสมอ
        return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด' }; 
  }
}

function getPasswordHint(username) {
  const data = getSheet_(SHEETS.USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { 
    if (data[i][1] === username) return { success: true, hint: data[i][4] || 'ไม่มีคำใบ้' }; 
  }
  return { success: false, message: 'ไม่พบผู้ใช้นี้' };
}

// ============================================
// 📦 MASTER DATA
// ============================================
function getStudents() { return { success: true, students: STUDENTS, subjects: SUBJECTS }; }
function getSubjects() { return { success: true, subjects: SUBJECTS }; }

// ============================================
// 📝 HOMEWORK
// ============================================
function addHomework(sub, desc, ad, dd, nd, by) {
  try {
    const uSheet = getSheet_(SHEETS.USERS); 
    const hwS = getSheet_(SHEETS.HOMEWORK); 
    const stS = getSheet_(SHEETS.HOMEWORK_STATUS);

    // Sanitize inputs
    const cleanSub  = sanitizeStr(sub, 100);
    const cleanDesc = sanitizeStr(desc, 1000);
    const cleanBy   = sanitizeStr(by, 200);

    if (!cleanSub) return { success: false, message: 'กรุณาระบุวิชา' };
    if (!cleanBy)  return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' };

    // ตรวจสอบ subject อยู่ใน whitelist
    if (!SUBJECTS.includes(cleanSub)) return { success: false, message: 'วิชาไม่ถูกต้อง' };

    const uData = uSheet.getDataRange().getValues(); 
    let hasPerm = false;
    let creatorName = '';

    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === cleanBy) {
        creatorName = uData[i][1];
        const rK = uData[i][6];
        if (ROLES[rK] && ROLES[rK].canManageHomework) {
          hasPerm = true;
        } else {
          const credits = (uData[i].length > 10 && uData[i][10]) ? Number(uData[i][10]) : 0;
          if (credits > 0) {
            hasPerm = true;
            uSheet.getRange(i + 1, 11).setValue(credits - 1);
          }
        }
        // Check temp access
        if (!hasPerm && uData[i].length > 9 && uData[i][9]) {
          const expiry = new Date(uData[i][9]);
          if (expiry > new Date()) hasPerm = true;
        }
        break;
      }
    }

    if (!hasPerm) return { success: false, message: 'คุณไม่มีสิทธิ์เพิ่มการบ้าน หรือ Credit หมดแล้ว' };

    const id = Utilities.getUuid();
    const color = getSubjectColor(cleanSub);

    hwS.appendRow([id, cleanSub, cleanDesc, ad, dd, nd, creatorName, new Date(), color]);
    hwS.getRange(hwS.getLastRow(), 1, 1, 9).setBackground(color);

    const rows = STUDENTS.map(s => [id, s.no, 'pending', '', '', '', color]);
    if (rows.length > 0) {
      const startRow = stS.getLastRow() + 1;
      stS.getRange(startRow, 1, rows.length, 7).setValues(rows);
      stS.getRange(startRow, 1, rows.length, 7).setBackground(color);
    }

    clearDashboardCaches_();
    return { success: true, message: 'เพิ่มการบ้านเรียบร้อยแล้ว' };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() }; 
  }
}

function getHomework(ss = null) {
  try {
    const activeSs = ss || getSpreadsheet();
    const hwSheet = getSheetFromSs_(activeSs, SHEETS.HOMEWORK); 
    const stSheet = getSheetFromSs_(activeSs, SHEETS.HOMEWORK_STATUS);

    const hwData = hwSheet.getDataRange().getValues(); 
    const stData = stSheet.getDataRange().getValues();

    // Optimize: Create a status map for O(1) lookup
    const statusMap = {};
    for (let j = 1; j < stData.length; j++) {
      const hwId = stData[j][0];
      if (!hwId) continue;
      if (!statusMap[hwId]) statusMap[hwId] = {};
      statusMap[hwId][stData[j][1]] = {
        status: stData[j][2] || 'pending',
        imagePath: stData[j][3] || ''
      };
    }

    const list = [];
    for (let i = 1; i < hwData.length; i++) { 
      const row = hwData[i];
      if (!row || !row[0]) continue;

      const item = {
        id: row[0],
        subject: row[1] || '(ไม่มีวิชา)',
        description: row[2] || '',
        assignedDate: row[3] ? new Date(row[3]).toISOString() : null,
        dueDate: row[4] ? new Date(row[4]).toISOString() : null,
        noDueDate: !!row[5],
        createdBy: row[6] || '',
        color: row[8] || getSubjectColor(row[1]),
        statuses: statusMap[row[0]] || {}
      };
      list.push(item);
    }
    return { success: true, homework: list };
  } catch (e) { 
    Logger.log('getHomework Error: ' + e.message); 
    return { success: false, message: e.message, homework: [] }; 
  }
}

function updateHomeworkStatus(hid, sno, stat, imgBase64) {
  try {
    const stS = getSheet_(SHEETS.HOMEWORK_STATUS);
    const d = stS.getDataRange().getValues();

    // Validate inputs
    const cleanHid  = sanitizeStr(hid, 100);
    const cleanSno  = parseInt(sno);
    const cleanStat = sanitizeStr(stat, 20);
    const VALID_STATUSES = ['pending', 'completed'];
    if (cleanStat && !VALID_STATUSES.includes(cleanStat)) return { success: false, message: 'สถานะไม่ถูกต้อง' };
    if (!cleanHid || isNaN(cleanSno) || cleanSno < 1) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === cleanHid && Number(d[i][1]) === cleanSno) {
        if (cleanStat) stS.getRange(i + 1, 3).setValue(cleanStat);
        if (imgBase64) {
          const url = uploadImageToDrive(imgBase64, 'HW_' + cleanSno + '_' + cleanHid);
          if (url) stS.getRange(i + 1, 4).setValue(url);
        }
        if (cleanStat === 'completed') stS.getRange(i + 1, 5).setValue(new Date());
        clearDashboardCaches_();
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรายการที่จะอัปเดต' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

function deleteHomework(id) {
  try {
    const hS = getSheet_(SHEETS.HOMEWORK);
    const sS = getSheet_(SHEETS.HOMEWORK_STATUS);

    const cleanId = sanitizeStr(id, 100);
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

    const hd = hS.getDataRange().getValues();
    for (let i = hd.length - 1; i >= 1; i--) {
      if (hd[i][0] === cleanId) hS.deleteRow(i + 1);
    }
    const sd = sS.getDataRange().getValues();
    for (let i = sd.length - 1; i >= 1; i--) {
      if (sd[i][0] === cleanId) sS.deleteRow(i + 1);
    }
    clearDashboardCaches_();
    return { success: true };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

// ============================================
// 🚪 LEAVE REQUESTS
// ============================================
function submitLeaveRequest(studentNo, studentName, type, date, reason, base64Image) {
  try {
    const sheet = getSheet_(SHEETS.LEAVE_REQUESTS); 

    // Sanitize inputs
    const cleanName   = sanitizeStr(studentName, 100);
    const cleanType   = sanitizeStr(type, 50);
    const cleanDate   = sanitizeStr(date, 20);
    const cleanReason = sanitizeStr(reason, 500);
    const cleanNo     = parseInt(studentNo) || 0;

    if (!cleanNo || cleanNo < 1 || cleanNo > 100) return { success: false, message: 'เลขที่นักเรียนไม่ถูกต้อง' };
    if (!cleanName) return { success: false, message: 'กรุณาระบุชื่อ' };
    if (!cleanType) return { success: false, message: 'กรุณาระบุประเภทการลา' };
    if (!cleanDate || !/^\d{4}-\d{2}-\d{2}$/.test(cleanDate)) return { success: false, message: 'รูปแบบวันที่ไม่ถูกต้อง' };

    const id = Utilities.getUuid(); 
    let imageUrl = '';

    if (base64Image) {
      try { 
        const split = base64Image.split(','); 
        const mimeMatch = split[0].match(/:(.*?);/); 
        const contentType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
        // ตรวจ MIME type ว่าเป็นรูปภาพเท่านั้น
        if (!['image/jpeg','image/png','image/gif','image/webp'].includes(contentType)) {
          return { success: false, message: 'ไฟล์ต้องเป็นรูปภาพเท่านั้น' };
        }
        const blob = Utilities.newBlob(Utilities.base64Decode(split[1]), contentType, 'leave_' + cleanNo + '_' + id); 
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID); 
        const file = folder.createFile(blob); 
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
        imageUrl = file.getUrl(); 
      } catch (imgErr) { 
        Logger.log('Image upload error: ' + imgErr.toString()); 
      }
    }

    sheet.appendRow([id, cleanNo, cleanName, cleanType, cleanDate, cleanReason, 'PENDING', imageUrl, false, new Date()]); 
    clearDashboardCaches_();
    return { success: true };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด' }; 
  }
}

function getLeaveRequests(ss = null) {
  try {
    const activeSs = ss || getSpreadsheet();
    const sheet = getSheetFromSs_(activeSs, SHEETS.LEAVE_REQUESTS); 
    const data = sheet.getDataRange().getValues(); 
    if (data.length <= 1) return [];
    return data.slice(1).map(row => ({
      id: row[0],
      studentNo: row[1],
      studentName: row[2],
      type: row[3],
      date: row[4],
      reason: row[5],
      status: row[6],
      proofImage: convertDriveToDirect(row[7]),
      confirmed: row[8] || false,
      timestamp: row[9] || ''
    }));
  } catch (e) { 
    Logger.log('getLeaveRequests Error: ' + e.toString()); 
    return []; 
  }
}

function confirmLeaveRequest(id, imgBase64) {
  try {
    const sheet = getSheet_(SHEETS.LEAVE_REQUESTS);
    const data = sheet.getDataRange().getValues();

    const cleanId = sanitizeStr(id, 100);
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === cleanId) {
        if (imgBase64) {
          // ตรวจ MIME type ว่าเป็นรูปภาพเท่านั้น
          const mimeMatch = imgBase64.match(/^data:(image\/[a-z]+);base64,/);
          if (!mimeMatch || !['image/jpeg','image/png','image/gif','image/webp'].includes(mimeMatch[1])) {
            return { success: false, message: 'ไฟล์ต้องเป็นรูปภาพเท่านั้น' };
          }
          const url = uploadImageToDrive(imgBase64, 'LeaveConfirm_' + cleanId);
          if (url) sheet.getRange(i + 1, 8).setValue(url);
        }
        sheet.getRange(i + 1, 9).setValue(true);
        clearDashboardCaches_();
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบคำขอนี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

function updateLeaveStatus(id, status) {
  try {
    const sheet = getSheet_(SHEETS.LEAVE_REQUESTS);
    const data = sheet.getDataRange().getValues();

    const cleanId = sanitizeStr(id, 100);
    const VALID_STATUSES = ['APPROVED', 'REJECTED', 'PENDING'];
    const cleanStatus = sanitizeStr(status, 20).toUpperCase();
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };
    if (!VALID_STATUSES.includes(cleanStatus)) return { success: false, message: 'สถานะไม่ถูกต้อง' };

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === cleanId) {
        sheet.getRange(i + 1, 7).setValue(cleanStatus);
        sheet.getRange(i + 1, 9).setValue(cleanStatus === 'APPROVED');
        clearDashboardCaches_();
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบคำขอนี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

// ============================================
// 💰 TREASURY
// ============================================
function addTreasuryItem(title, amt, by, targetStudents = null) {
  try {
    const ss = getSpreadsheet();
    const uSheet = getSheetFromSs_(ss, SHEETS.USERS);
    const uData = uSheet.getDataRange().getValues();
    let hasPerm = false;
    const cleanBy = sanitizeStr(by, 200);
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === cleanBy) {
        const rK = uData[i][6];
        if (ROLES[rK] && ROLES[rK].canManageTreasury) hasPerm = true;
        break;
      }
    }
    if (!hasPerm) return { success: false, message: 'คุณไม่มีสิทธิ์จัดการเงินห้อง' };

    const cleanTitle = sanitizeStr(title, 200);
    const cleanAmt   = parseFloat(amt);

    if (!cleanTitle) return { success: false, message: 'กรุณาระบุชื่อรายการ' };
    if (isNaN(cleanAmt) || cleanAmt < 0 || cleanAmt > 100000) return { success: false, message: 'จำนวนเงินไม่ถูกต้อง' };

    const tS = getSheetFromSs_(ss, SHEETS.TREASURY); 
    const pS = getSheetFromSs_(ss, SHEETS.TREASURY_PAYMENTS); 
    const id = Utilities.getUuid(); 
    const color = getSubjectColor(cleanTitle); 
    
    tS.appendRow([id, cleanTitle, cleanAmt, cleanBy, new Date(), 'active', color]); 
    
    let targetNos = [];
    if (targetStudents) {
      if (Array.isArray(targetStudents)) {
        // ตรวจว่า student numbers ถูกต้อง
        targetNos = targetStudents.map(n => parseInt(n)).filter(n => n > 0 && n <= 100);
      } else {
        targetNos = String(targetStudents).split(',').map(n => parseInt(n.trim())).filter(n => n > 0 && n <= 100);
      }
    } else {
      targetNos = STUDENTS.map(s => s.no);
    }

    const rows = targetNos.map(no => [id, no, 0, '', '', color, '']);
    if (rows.length > 0) pS.getRange(pS.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    
    clearDashboardCaches_();
    return { success: true, treasuryId: id };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

function getTreasuryItems(ss = null) {
  try {
    const activeSs = ss || getSpreadsheet();
    const tSheet = getSheetFromSs_(activeSs, SHEETS.TREASURY); 
    const pSheet = getSheetFromSs_(activeSs, SHEETS.TREASURY_PAYMENTS);

    const tData = tSheet.getDataRange().getValues(); 
    const pData = pSheet ? pSheet.getDataRange().getValues() : [];

    const payMap = {};
    for (let r = 1; r < pData.length; r++) {
      const tId = pData[r][0];
      if (!tId) continue;
      if (!payMap[tId]) payMap[tId] = {};
      payMap[tId][pData[r][1]] = {
        amountPaid: parseFloat(pData[r][2]) || 0,
        paidAt: pData[r][3] ? new Date(pData[r][3]).toISOString() : '',
        notes: pData[r][4] || '',
        color: pData[r][5] || '',
        receiptNo: pData[r][6] || ''
      };
    }

    const list = [];
    for (let r = 1; r < tData.length; r++) {
      const row = tData[r];
      if (!row[0]) continue;
      const tId = row[0];
      const amtPerPerson = parseFloat(row[2]) || 0;
      const itemPay = payMap[tId] || {};
      let countDone = 0;
      const payments = {};

      for (const sNo in itemPay) {
        const paid = itemPay[sNo].amountPaid;
        payments[sNo] = { amountPaid: paid, isComplete: paid >= amtPerPerson };
        if (paid >= amtPerPerson) countDone++;
      }

      list.push({
        id: tId,
        title: row[1] || '-',
        amountPerPerson: amtPerPerson,
        payments: payments,
        summary: { completedCount: countDone },
        color: row[6] || getSubjectColor(row[1] || '-') // ดึงค่าสี (หรือสร้างใหม่ถ้ารายการเก่าไม่มี)
      });
    }
    return { success: true, treasury: list };
  } catch (e) { 
    Logger.log('getTreasuryItems Error: ' + e); 
    return { success: false, message: e.toString(), treasury: [] }; 
  }
}

function updatePayment(tid, sno, paid) {
  try {
    const pS = getSheet_(SHEETS.TREASURY_PAYMENTS);
    const d = pS.getDataRange().getValues();

    const cleanTid  = sanitizeStr(tid, 100);
    const cleanSno  = parseInt(sno);
    const cleanPaid = parseFloat(paid);
    if (!cleanTid || isNaN(cleanSno) || cleanSno < 1) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };
    if (isNaN(cleanPaid) || cleanPaid < 0 || cleanPaid > 100000) return { success: false, message: 'จำนวนเงินไม่ถูกต้อง' };

    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === cleanTid && Number(d[i][1]) === cleanSno) {
        const paidAt = cleanPaid > 0 ? new Date() : '';
        // Generate server-side receipt number
        const receiptNo = 'RC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().split('-')[0];
        pS.getRange(i + 1, 3).setValue(cleanPaid);
        pS.getRange(i + 1, 4).setValue(paidAt);
        // Store receiptNo in column 7 (ReceiptNo) - creates column if missing
        try { pS.getRange(i + 1, 7).setValue(receiptNo); } catch(e) { /* ignore if cannot write */ }
        const props = PropertiesService.getScriptProperties();
        const count = parseInt(props.getProperty('tr_pay_counter') || '0');
        props.setProperty('tr_pay_counter', (count + 1).toString());
        clearDashboardCaches_();
        return { success: true, paidAt: paidAt ? paidAt.toISOString() : '', receiptNo: receiptNo };
      }
    }
    return { success: false, message: 'ไม่พบรายการ' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

function deleteTreasuryItem(id) {
  try {
    const tS = getSheet_(SHEETS.TREASURY);
    const pS = getSheet_(SHEETS.TREASURY_PAYMENTS);

    const cleanId = sanitizeStr(id, 100);
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

    const td = tS.getDataRange().getValues();
    for (let i = td.length - 1; i >= 1; i--) {
      if (td[i][0] === cleanId) tS.deleteRow(i + 1);
    }
    const pd = pS.getDataRange().getValues();
    for (let i = pd.length - 1; i >= 1; i--) {
      if (pd[i][0] === cleanId) pS.deleteRow(i + 1);
    }
    clearDashboardCaches_();
    return { success: true };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

// ============================================
// 💸 LOANS (เงินยืม)
// ============================================
function addLoan(borrowerNo, lenderNo, amount, note, proofBase64, createdBy) {
  try {
    const cleanBorrower = parseInt(borrowerNo);
    const cleanLender   = parseInt(lenderNo);
    const cleanAmount   = parseFloat(amount);
    const cleanNote     = sanitizeStr(note, 500);
    const cleanBy       = sanitizeStr(createdBy, 200);

    if (isNaN(cleanBorrower) || cleanBorrower < 1) return { success: false, message: 'เลขที่ผู้ยืมไม่ถูกต้อง' };
    if (isNaN(cleanLender)   || cleanLender < 1)   return { success: false, message: 'เลขที่ผู้ให้ยืมไม่ถูกต้อง' };
    if (cleanBorrower === cleanLender) return { success: false, message: 'ผู้ยืมและผู้ให้ยืมต้องไม่เป็นคนเดียวกัน' };
    if (isNaN(cleanAmount) || cleanAmount <= 0 || cleanAmount > 1000000) return { success: false, message: 'จำนวนเงินไม่ถูกต้อง' };

    // Upload proof image if provided
    let proofUrl = '';
    if (proofBase64 && proofBase64.startsWith('data:image')) {
      proofUrl = uploadImageToDrive(proofBase64, 'loan_proof_' + Date.now() + '.jpg') || '';
    }

    const id = Utilities.getUuid();
    const sheet = getSheet_(SHEETS.LOANS);
    sheet.appendRow([id, cleanBorrower, cleanLender, cleanAmount, cleanNote, proofUrl, 'pending', new Date().toISOString(), cleanBy]);

    return { success: true, loanId: id };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

function getLoans() {
  try {
    ensureSheetsExist();
    const sheet = getSheet_(SHEETS.LOANS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, loans: [] };

    const data = sheet.getDataRange().getValues();
    const loans = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      loans.push({
        id:         row[0],
        borrowerNo: row[1],
        lenderNo:   row[2],
        amount:     parseFloat(row[3]) || 0,
        note:       row[4] || '',
        proofUrl:   row[5] || '',
        status:     row[6] || 'pending',
        createdAt:  row[7] ? new Date(row[7]).toISOString() : '',
        createdBy:  row[8] || ''
      });
    }
    // Most recent first
    loans.reverse();
    return { success: true, loans: loans };
  } catch (e) {
    return { success: false, message: e.toString(), loans: [] };
  }
}

function updateLoanStatus(id, status) {
  try {
    const cleanId     = sanitizeStr(id, 100);
    const cleanStatus = sanitizeStr(status, 50);
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };
    const allowed = ['pending', 'returned'];
    if (!allowed.includes(cleanStatus)) return { success: false, message: 'สถานะไม่ถูกต้อง' };

    const sheet = getSheet_(SHEETS.LOANS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === cleanId) {
        sheet.getRange(i + 1, 7).setValue(cleanStatus);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรายการ' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

function deleteLoan(id) {
  try {
    const cleanId = sanitizeStr(id, 100);
    if (!cleanId) return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };

    const sheet = getSheet_(SHEETS.LOANS);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === cleanId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบรายการ' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด' };
  }
}

// ============================================
// 🔄 REAL-TIME POLLING
// ============================================
function getLastUpdate() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_UPDATE_KEY);
    if (cached) {
      return JSON.parse(cached);
    }

    const ss = getSpreadsheet();
    const hw = getSheetFromSs_(ss, SHEETS.HOMEWORK);
    const tr = getSheetFromSs_(ss, SHEETS.TREASURY);
    const lv = getSheetFromSs_(ss, SHEETS.LEAVE_REQUESTS);

    let pCount = 0;
    let lastPendingType = '';

    if (lv && lv.getLastRow() > 1) {
      const lvData = lv.getDataRange().getValues();
      const pendingRows = lvData.slice(1).filter(function(r) { return r[6] === 'PENDING'; });
      pCount = pendingRows.length;
      const lastRowVals = lvData[lvData.length - 1];
      if (lastRowVals[6] === 'PENDING') {
        lastPendingType = lastRowVals[3] || 'รายการ';
      }
    }

    const props = PropertiesService.getScriptProperties();
    const trPayCounter = parseInt(props.getProperty('tr_pay_counter') || '0');

    let seatVersion = 0;
    const seatMeta = getSheetFromSs_(ss, SHEETS.SEAT_META);
    if (seatMeta && seatMeta.getLastRow() >= 2) {
      seatVersion = Number(seatMeta.getRange(2, 5).getValue()) || 0;
    }

    const result = {
      success: true,
      hwCount: hw ? Math.max(0, hw.getLastRow() - 1) : 0,
      trCount: tr ? Math.max(0, tr.getLastRow() - 1) : 0,
      pendingLeaveCount: pCount,
      lastPendingType: lastPendingType,
      trPayCounter: trPayCounter,
      seatVersion: seatVersion
    };

    cache.put(CACHE_UPDATE_KEY, JSON.stringify(result), CACHE_TTL);
    return result;
  } catch (e) { 
    return { success: false, hwCount: 0, trCount: 0, pendingLeaveCount: 0, lastPendingType: '', trPayCounter: 0 }; 
  }
}

// ============================================
// 🎟️ REDEEM CODE
// ============================================
function generateCode(actionType, value, details) {
  try {
    const sheet = getSheet_(SHEETS.REDEEM_CODES);
    const code = Utilities.getUuid().split('-')[0].toUpperCase();
    const qr = createQrInDrive(code);
    const now = new Date();
    sheet.appendRow([code, actionType, value, JSON.stringify(details || {}), 1, 0, 'Owner', now, qr.fileId || '', qr.url || '', now.toISOString()]);
    clearDashboardCaches_();
    return { success: true, code: code, qrUrl: qr.url };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function createQrInDrive(code) {
  try {
    const url = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(code)}&ecc=H`;
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob().setName(`Redeem_${code}.png`);
    const folder = DriveApp.getFolderById(QR_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { fileId: file.getId(), url: `https://drive.google.com/uc?export=view&id=${file.getId()}` };
  } catch (e) {
    console.error('QR create error:', e.toString());
    return { fileId: '', url: '' };
  }
}

function cleanupOldQrFiles() {
  try {
    const folder = DriveApp.getFolderById(QR_FOLDER_ID);
    const files = folder.getFiles();
    const now = new Date().getTime();
    const oneDay = 24 * 60 * 60 * 1000;
    while (files.hasNext()) {
      const f = files.next();
      if (now - f.getDateCreated().getTime() > oneDay) f.setTrashed(true);
    }
  } catch (e) {
    console.error('cleanupOldQrFiles error:', e.toString());
  }
}

// ============================================
// 🔑 GUEST ACCOUNT SYSTEM
// createGuestAccount: ใส่โค้ด TEMP_ACCESS → สร้างบัญชี GUEST อัตโนมัติ → login ทันที
// deleteGuestAccount: ลบบัญชี GUEST เมื่อหมดเวลาหรือออกจากระบบ
// ============================================

function createGuestAccount(code) {
  try {
    const ss = getSpreadsheet();
    const codeSheet = getSheetFromSs_(ss, SHEETS.REDEEM_CODES);
    const userSheet = getSheetFromSs_(ss, SHEETS.USERS);

    if (!code) return { success: false, message: 'กรุณากรอกโค้ด' };

    // ค้นหาโค้ด
    const codeData = codeSheet.getDataRange().getValues();
    let codeRow = null, codeRowIndex = -1;
    for (let i = 1; i < codeData.length; i++) {
      if (String(codeData[i][0]).toUpperCase() === String(code).toUpperCase()) {
        codeRow = codeData[i];
        codeRowIndex = i;
        break;
      }
    }

    if (!codeRow) return { success: false, message: 'ไม่พบโค้ดนี้ กรุณาตรวจสอบอีกครั้ง' };

    // ต้องเป็น TEMP_ACCESS เท่านั้น
    if (codeRow[1] !== 'TEMP_ACCESS') {
      return { success: false, message: 'โค้ดนี้ไม่ใช่โค้ดสิทธิ์ชั่วคราว กรุณาล็อกอินก่อนใช้งาน' };
    }

    // ตรวจจำนวนการใช้
    const maxUses = Number(codeRow[4]) || 1;
    const usesCount = Number(codeRow[5]) || 0;
    if (usesCount >= maxUses) return { success: false, message: 'โค้ดนี้ถูกใช้งานครบแล้ว' };

    // คำนวณเวลาหมดอายุ
    const val = String(codeRow[2]);
    const parts = val.split('_');
    const num = parseInt(parts[0]) || 1;
    const unit = (parts[1] || 'DAYS').toUpperCase();
    const expiresAt = new Date();
    if (unit === 'DAYS')        expiresAt.setDate(expiresAt.getDate() + num);
    else if (unit === 'WEEKS')  expiresAt.setDate(expiresAt.getDate() + (num * 7));
    else if (unit === 'MONTHS') expiresAt.setMonth(expiresAt.getMonth() + num);
    else if (unit === 'YEARS')  expiresAt.setFullYear(expiresAt.getFullYear() + num);

    // สร้าง ID และชื่อบัญชีชั่วคราว
    const guestId   = 'GUEST_' + Utilities.getUuid().split('-')[0].toUpperCase();
    const guestName = 'Guest_' + Utilities.getUuid().split('-')[0].slice(0, 5).toUpperCase();

    // บันทึกบัญชีใหม่ใน Users sheet
    userSheet.appendRow([
      guestId, guestName, '', '', '', '', 'GUEST',
      new Date(), new Date(), expiresAt, 0
    ]);

    // Mark code as used
    codeSheet.getRange(codeRowIndex + 1, 6).setValue(usesCount + 1);

    const unitTh = unit === 'DAYS' ? 'วัน' : unit === 'WEEKS' ? 'อาทิตย์' : unit === 'MONTHS' ? 'เดือน' : 'ปี';
    const expiresAtISO = expiresAt.toISOString();

    clearDashboardCaches_();
    return {
      success: true,
      expiresAt: expiresAtISO,
      durationText: num + ' ' + unitTh,
      user: {
        id: guestId,
        displayName: guestName,
        email: '',
        role: ROLES.GUEST.name,
        roleKey: 'GUEST',
        studentNo: null,
        isGuest: true,
        canManageHomework: false,
        canManageTreasury: false,
        canApproveLeave: false,
        canManageCodes: false,
        hwCredits: 0,
        expiresAt: expiresAtISO
      }
    };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

function deleteGuestAccount(userId) {
  try {
    if (!userId || !String(userId).startsWith('GUEST_')) {
      return { success: false, message: 'ไม่ใช่บัญชีชั่วคราว' };
    }
    const userSheet = getSheet_(SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(userId) && data[i][6] === 'GUEST') {
        userSheet.deleteRow(i + 1);
        clearDashboardCaches_();
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบบัญชีชั่วคราวนี้' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Time-based trigger: ลบบัญชี GUEST ที่หมดอายุทั้งหมด (ตั้ง trigger ทุก 1 ชม. ใน GAS)
function cleanupExpiredGuestAccounts() {
  try {
    const userSheet = getSheet_(SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();
    const now = new Date();
    let deleted = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][6] === 'GUEST' && data[i][9]) {
        if (new Date(data[i][9]) < now) {
          userSheet.deleteRow(i + 1);
          deleted++;
        }
      }
    }
    if (deleted > 0) {
      clearDashboardCaches_();
    }
    return { success: true, deleted: deleted };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// FIX: ADD_MONEY now creates TreasuryPayments rows for all students
// FIX: Returns addedCredits for ADD_HW so client can update local state
function redeemCode(code, userId) {
  try {
    const ss = getSpreadsheet();
    const codeSheet = getSheetFromSs_(ss, SHEETS.REDEEM_CODES); 
    const userSheet = getSheetFromSs_(ss, SHEETS.USERS);

    const codeData = codeSheet.getDataRange().getValues();

    for (let i = 1; i < codeData.length; i++) {
      if (String(codeData[i][0]).toUpperCase() === String(code).toUpperCase()) {
        const maxUses = codeData[i][4];
        const usesCount = codeData[i][5];
        if (usesCount >= maxUses) return { success: false, message: 'โค้ดนี้ถูกใช้งานครบแล้ว' };

        const action = codeData[i][1];
        const val = String(codeData[i][2]);
        const details = JSON.parse(codeData[i][3] || '{}');

        // Mark code as used
        codeSheet.getRange(i + 1, 6).setValue(usesCount + 1);
        cleanupOldQrFiles();

        let resultMsg = 'ใช้โค้ดสำเร็จ';
        const extraData = { action: action };
        const uData = userSheet.getDataRange().getValues();

        for (let j = 1; j < uData.length; j++) {
          if (uData[j][0] === userId) {

            if (action === 'TEMP_ACCESS') {
              // FIX: Parse "qty_UNIT" format (e.g. "7_DAYS", "2_WEEKS")
              const parts = val.split('_');
              const num = parseInt(parts[0]) || 1;
              const unit = (parts[1] || 'DAYS').toUpperCase();
              const expiry = new Date();
              if (unit === 'DAYS')   expiry.setDate(expiry.getDate() + num);
              else if (unit === 'WEEKS')  expiry.setDate(expiry.getDate() + (num * 7));
              else if (unit === 'MONTHS') expiry.setMonth(expiry.getMonth() + num);
              else if (unit === 'YEARS')  expiry.setFullYear(expiry.getFullYear() + num);
              userSheet.getRange(j + 1, 10).setValue(expiry);
              resultMsg = `ได้รับสิทธิ์ผู้ใช้ชั่วคราว ${num} ${unit === 'DAYS' ? 'วัน' : unit === 'WEEKS' ? 'อาทิตย์' : unit === 'MONTHS' ? 'เดือน' : 'ปี'}`;
            }
            else if (action === 'ADD_HW') {
              const addAmt = parseInt(val) || 1;
              const cur = (uData[j].length > 10 && uData[j][10]) ? Number(uData[j][10]) : 0;
              userSheet.getRange(j + 1, 11).setValue(cur + addAmt);
              resultMsg = `ได้รับเครดิตตั้งการบ้าน ${addAmt} ครั้ง`;
              extraData.addedCredits = addAmt;
            }
            else if (action === 'ADD_MONEY') {
              const tSheet = getSheetFromSs_(ss, SHEETS.TREASURY);
              const pSheet = getSheetFromSs_(ss, SHEETS.TREASURY_PAYMENTS);
              const tId = Utilities.getUuid();
              const title = details.title || 'รายการพิเศษ';
              const amt = parseFloat(details.amount) || 0;
              const color = getSubjectColor(title); // เพิ่มสี

              // เพิ่ม color เข้าไปในแถว
              tSheet.appendRow([tId, title, amt, 'Code:' + code, new Date(), 'active', color]); 
              
              const rows = STUDENTS.map(s => [tId, s.no, 0, '', '', '']);
              if (rows.length > 0) {
                pSheet.getRange(pSheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
              }
              resultMsg = `สร้างรายการเก็บเงิน "${title}" (${amt} บาท/คน) สำเร็จ`;
              extraData.treasuryId = tId;
              extraData.title = title;
              extraData.amount = amt;
            }
            break;
          }
        }
        clearDashboardCaches_();
        return { success: true, message: resultMsg, ...extraData };
      }
    }
    return { success: false, message: 'ไม่พบโค้ดนี้ หรือโค้ดไม่ถูกต้อง' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  }
}

// ============================================
// 📷 TIMETABLE (Owner: image + link in Sheet)
// ============================================
function getTimetable() {
  try {
    const sh = getSheet_(SHEETS.TIMETABLE);
    if (!sh || sh.getLastRow() < 2) return { success: true, imageUrl: '', linkUrl: '', updatedAt: '' };
    const r = sh.getRange(2, 1, 1, 4).getValues()[0];
    return {
      success: true,
      imageUrl: convertDriveToDirect(String(r[0] || '')),
      linkUrl: String(r[1] || ''),
      updatedAt: r[2] ? new Date(r[2]).toISOString() : '',
      updatedBy: String(r[3] || '')
    };
  } catch (e) {
    return { success: false, message: e.toString(), imageUrl: '', linkUrl: '' };
  }
}

function setTimetable(userId, imageBase64, linkUrl) {
  try {
    const row = getUserRow_(userId);
    if (!row) return { success: false, message: 'ไม่พบผู้ใช้' };
    const rk = row.data[6];
    if (rk !== 'OWNER') return { success: false, message: 'เฉพาะเจ้าของเว็บเท่านั้นที่แก้ไขตารางเรียนได้' };

    const sh = getSheet_(SHEETS.TIMETABLE);
    const cur = sh.getRange(2, 1, 1, 4).getValues()[0];
    let imageUrl = cur[0] || '';
    let link = cur[1] || '';
    if (imageBase64) {
      const up = uploadImageToDrive(imageBase64, 'Timetable_' + new Date().getTime() + '.png');
      if (up) imageUrl = up;
    }
    // linkUrl === undefined → ไม่เปลี่ยนลิงก์ (ใช้ตอนอัปโหลดรูปอย่างเดียว)
    if (linkUrl !== undefined && linkUrl !== null) {
      link = String(linkUrl).trim();
    }
    sh.getRange(2, 1, 1, 4).setValues([[imageUrl, link, new Date(), row.data[1]]]);
    clearDashboardCaches_();
    return { success: true, imageUrl: convertDriveToDirect(String(imageUrl)), linkUrl: link };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================
// 🪑 SEAT MAP + BOOKING (Sheets backend)
// ============================================
function getUserRow_(userId) {
  if (!userId) return null;
  const sh = getSheet_(SHEETS.USERS);
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(userId)) return { row: i + 1, data: d[i] };
  }
  return null;
}

function clearSheetDataRows_(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  sheet.deleteRows(2, lastRow - 1);
  return lastRow - 1;
}

function clearGuestUsers_() {
  const userSheet = getSheet_(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] === 'GUEST' || String(data[i][0]).startsWith('GUEST_')) {
      userSheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}

function clearAllWebsiteData(userId) {
  try {
    const userRow = getUserRow_(userId);
    if (!userRow || userRow.data[6] !== 'OWNER') {
      return { success: false, message: 'เฉพาะเจ้าของเว็บเท่านั้นที่ลบข้อมูลทั้งหมดได้' };
    }

    const activeSs = getSpreadsheet();
    const cleared = {};
    [
      SHEETS.HOMEWORK,
      SHEETS.HOMEWORK_STATUS,
      SHEETS.TREASURY,
      SHEETS.TREASURY_PAYMENTS,
      SHEETS.LEAVE_REQUESTS,
      SHEETS.REDEEM_CODES,
      SHEETS.SEAT_BOOKINGS,
      SHEETS.SEAT_EDIT_CODES,
      SHEETS.SEAT_EDIT_SESSIONS,
      SHEETS.LOANS
    ].forEach(function(sheetName) {
      cleared[sheetName] = clearSheetDataRows_(getSheetFromSs_(activeSs, sheetName));
    });

    const timetableSheet = getSheetFromSs_(activeSs, SHEETS.TIMETABLE);
    if (timetableSheet) {
      clearSheetDataRows_(timetableSheet);
      timetableSheet.appendRow(['', '', new Date(), userId]);
      cleared[SHEETS.TIMETABLE] = 1;
    }

    const seatMetaSheet = getSheetFromSs_(activeSs, SHEETS.SEAT_META);
    if (seatMetaSheet) {
      clearSheetDataRows_(seatMetaSheet);
      const defaultLayout = JSON.stringify({
        grid: { cols: 22, rows: 16, cell: 24 },
        seats: [],
        frontBand: 2
      });
      seatMetaSheet.appendRow([defaultLayout, '', '', 2, 0, new Date().toISOString()]);
      cleared[SHEETS.SEAT_META] = 1;
    }

    cleared.GuestUsers = clearGuestUsers_();

    const props = PropertiesService.getScriptProperties();
    props.setProperty('tr_pay_counter', '0');
    clearDashboardCaches_();

    return {
      success: true,
      message: 'ลบข้อมูลทดลองทั้งหมดแล้ว',
      cleared: cleared
    };
  } catch (e) {
    return { success: false, message: 'ลบข้อมูลไม่สำเร็จ: ' + e.toString() };
  }
}

function roleIsSeatAdmin_(roleKey) {
  return roleKey === 'OWNER' || roleKey === 'TEACHER';
}

function parseLayout_(raw) {
  try {
    const o = raw ? JSON.parse(String(raw)) : {};
    return {
      grid: Object.assign({ cols: 22, rows: 16, cell: 24 }, o.grid || {}),
      seats: Array.isArray(o.seats) ? o.seats : [],
      frontBand: Number(o.frontBand) >= 0 ? Number(o.frontBand) : 2
    };
  } catch (e) {
    return { grid: { cols: 22, rows: 16, cell: 24 }, seats: [], frontBand: 2 };
  }
}

function normalizeSeatId_(id) {
  const seat = String(id || '').trim();
  return /^[A-Za-z0-9_-]{1,40}$/.test(seat) ? seat : '';
}

function normalizeLayout_(layout) {
  const raw = layout || {};
  const grid = raw.grid || {};
  const cols = Math.max(6, Math.min(60, Number(grid.cols) || 22));
  const rows = Math.max(6, Math.min(60, Number(grid.rows) || 16));
  const cell = Math.max(20, Math.min(48, Number(grid.cell) || 24));
  const frontBand = Math.max(0, Math.min(rows, Number(raw.frontBand) || 0));
  const seen = {};
  const seats = [];

  (Array.isArray(raw.seats) ? raw.seats : []).slice(0, 200).forEach(function(seat, index) {
    const id = normalizeSeatId_(seat && seat.id) || ('S_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8));
    if (seen[id]) return;
    seen[id] = true;
    const gw = Math.max(1, Math.min(10, Number(seat && seat.gw) || 3));
    const gh = Math.max(1, Math.min(10, Number(seat && seat.gh) || 2));
    const gx = Math.max(0, Math.min(cols - gw, Number(seat && seat.gx) || 0));
    const gy = Math.max(0, Math.min(rows - gh, Number(seat && seat.gy) || 0));
    seats.push({ id, gx, gy, gw, gh, label: sanitizeStr(seat && seat.label, 40) || String(index + 1), lock: !!(seat && seat.lock) });
  });

  return { grid: { cols, rows, cell }, seats, frontBand };
}

function withSeatLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function checkRateLimit_(scope, identifier, maxAttempts, windowMs) {
  try {
    const props = PropertiesService.getScriptProperties();
    const keySeed = scope + '|' + String(identifier || '');
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, keySeed);
    const key = 'rlx_' + byteDigestToHex_(digest).slice(0, 24);
    const now = Date.now();
    const raw = props.getProperty(key);
    let record = raw ? JSON.parse(raw) : { count: 0, windowStart: now };
    if (now - record.windowStart > windowMs) record = { count: 0, windowStart: now };
    if (record.count >= maxAttempts) return { blocked: true };
    record.count += 1;
    props.setProperty(key, JSON.stringify(record));
    return { blocked: false };
  } catch (e) {
    return { blocked: false };
  }
}

function bumpSeatVersion_(sh) {
  const v = Number(sh.getRange(2, 5).getValue()) || 0;
  sh.getRange(2, 5).setValue(v + 1);
  sh.getRange(2, 6).setValue(new Date().toISOString());
  return v + 1;
}

function seatSessionValid_(token) {
  if (!token) return null;
  const sh = getSheet_(SHEETS.SEAT_EDIT_SESSIONS);
  const d = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(token) && new Date(d[i][2]) > now) {
      return { row: i + 1, codeId: d[i][1] };
    }
  }
  return null;
}

function findStudentByCode_(code5) {
  const c = String(code5 || '').trim();
  if (!/^\d{5}$/.test(c)) return null;
  for (let i = 0; i < STUDENTS.length; i++) {
    if (String(STUDENTS[i].code) === c) return STUDENTS[i];
  }
  return null;
}

/** Public + editor snapshot */
function seatGetSnapshot(userId, guestEditToken) {
  try {
    const ss = getSpreadsheet();
    const meta = getSheetFromSs_(ss, SHEETS.SEAT_META);
    const book = getSheetFromSs_(ss, SHEETS.SEAT_BOOKINGS);
    const row = userId ? getUserRow_(userId) : null;
    const roleKey = row ? row.data[6] : '';
    const layout = normalizeLayout_(parseLayout_(meta.getRange(2, 1).getValue()));
    const bs = meta.getRange(2, 2).getValue();
    const be = meta.getRange(2, 3).getValue();
    const bookingStart = bs ? new Date(bs).toISOString() : '';
    const bookingEnd = be ? new Date(be).toISOString() : '';
    const now = new Date();
    let bookingOpen = true;
    if (bs && new Date(bs) > now) bookingOpen = false;
    if (be && new Date(be) < now) bookingOpen = false;

    const bd = book.getDataRange().getValues();
    const bookings = [];
    for (let i = 1; i < bd.length; i++) {
      if (!bd[i][0]) continue;
      bookings.push({
        seatId: bd[i][0],
        studentNo: bd[i][1],
        studentName: bd[i][2],
        bookedByUserId: bd[i][3]
      });
    }

    const ver = Number(meta.getRange(2, 5).getValue()) || 0;
    const sess = guestEditToken ? seatSessionValid_(guestEditToken) : null;
    const canManageSettings = roleIsSeatAdmin_(roleKey);
    const canEditLayout = canManageSettings || !!sess;

    return {
      success: true,
      layout: layout,
      bookings: bookings,
      bookingStart: bookingStart,
      bookingEnd: bookingEnd,
      bookingOpen: bookingOpen,
      version: ver,
      canEditLayout: canEditLayout,
      canManageSettings: canManageSettings,
      guestEditActive: !!sess,
      roleKey: roleKey || ''
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatSetBookingWindow(userId, startISO, endISO) {
  try {
    const row = getUserRow_(userId);
    if (!row) return { success: false, message: 'ไม่พบผู้ใช้' };
    if (!roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์ตั้งเวลาเปิดจอง' };
    const meta = getSheet_(SHEETS.SEAT_META);
    meta.getRange(2, 2).setValue(startISO ? new Date(startISO) : '');
    meta.getRange(2, 3).setValue(endISO ? new Date(endISO) : '');
    const v = bumpSeatVersion_(meta);
    clearDashboardCaches_();
    return { success: true, version: v };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatSetFrontBand(userId, frontBandRows, guestEditToken) {
  try {
    return withSeatLock_(function() {
      const snap = seatGetSnapshot(userId, guestEditToken);
      if (!snap.success || !snap.canEditLayout) return { success: false, message: 'Permission denied' };
      const meta = getSheet_(SHEETS.SEAT_META);
      const layout = normalizeLayout_(parseLayout_(meta.getRange(2, 1).getValue()));
      layout.frontBand = Math.max(0, Math.min(Number(frontBandRows) || 0, layout.grid.rows));
      meta.getRange(2, 4).setValue(layout.frontBand);
      meta.getRange(2, 1).setValue(JSON.stringify(layout));
      const v = bumpSeatVersion_(meta);
      clearDashboardCaches_();
      return { success: true, layout: layout, version: v };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatSaveLayout(userId, layoutJson, guestEditToken) {
  try {
    return withSeatLock_(function() {
      const snap = seatGetSnapshot(userId, guestEditToken);
      if (!snap.success || !snap.canEditLayout) return { success: false, message: 'Permission denied' };
      const layout = typeof layoutJson === 'string' ? parseLayout_(layoutJson) : parseLayout_(JSON.stringify(layoutJson));
      const normalized = normalizeLayout_(layout);
      const meta = getSheet_(SHEETS.SEAT_META);
      meta.getRange(2, 1).setValue(JSON.stringify(normalized));
      meta.getRange(2, 4).setValue(normalized.frontBand);
      const v = bumpSeatVersion_(meta);
      clearDashboardCaches_();
      return { success: true, version: v };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatBook(userId, seatId, studentCode5) {
  try {
    return withSeatLock_(function() {
      const row = getUserRow_(userId);
      if (!row) return { success: false, message: 'Login required' };
      const normalizedSeatId = normalizeSeatId_(seatId);
      if (!normalizedSeatId) return { success: false, message: 'Invalid seat id' };
      const st = findStudentByCode_(studentCode5);
      if (!st) return { success: false, message: 'Invalid 5-digit student code' };

      const snap = seatGetSnapshot(userId, null);
      if (!snap.bookingOpen) return { success: false, message: 'Booking is not open right now' };
      const seat = (snap.layout.seats || []).find(function(s) { return String(s.id) === String(normalizedSeatId); });
      if (!seat) return { success: false, message: 'Seat not found' };
      if (seat.lock) return { success: false, message: 'Seat is locked for teachers/front row' };

      const ss = getSpreadsheet();
      const book = getSheetFromSs_(ss, SHEETS.SEAT_BOOKINGS);
      const d = book.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][1]) === String(st.no)) return { success: false, message: 'This student number already booked another seat' };
      }
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][0]) === String(normalizedSeatId)) return { success: false, message: 'Seat already booked' };
      }

      book.appendRow([normalizedSeatId, st.no, st.name, userId, new Date()]);
      bumpSeatVersion_(getSheetFromSs_(ss, SHEETS.SEAT_META));
      clearDashboardCaches_();
      return { success: true, message: 'Booked successfully' };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatCancelBooking(userId, seatId) {
  try {
    return withSeatLock_(function() {
      const row = getUserRow_(userId);
      if (!row) return { success: false, message: 'Login required' };
      const normalizedSeatId = normalizeSeatId_(seatId);
      if (!normalizedSeatId) return { success: false, message: 'Invalid seat id' };
      const rk = row.data[6];
      const admin = roleIsSeatAdmin_(rk);

      const ss = getSpreadsheet();
      const book = getSheetFromSs_(ss, SHEETS.SEAT_BOOKINGS);
      const d = book.getDataRange().getValues();
      const myNo = row.data[5];
      for (let i = d.length - 1; i >= 1; i--) {
        if (String(d[i][0]) !== String(normalizedSeatId)) continue;
        const bookedBy = String(d[i][3]);
        const targetNo = d[i][1];
        if (admin || bookedBy === String(userId) || String(myNo) === String(targetNo)) {
          book.deleteRow(i + 1);
          bumpSeatVersion_(getSheetFromSs_(ss, SHEETS.SEAT_META));
          clearDashboardCaches_();
          return { success: true, message: 'Booking canceled' };
        }
        return { success: false, message: 'You cannot cancel this booking' };
      }
      return { success: false, message: 'Booking not found' };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatCreateEditCode(userId, plainCode, durationMinutes, label) {
  try {
    const row = getUserRow_(userId);
    if (!row) return { success: false, message: 'ไม่พบผู้ใช้' };
    if (!roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์สร้างโค้ดแก้ไข' };
    const pc = String(plainCode || '').trim();
    if (pc.length < 4) return { success: false, message: 'โค้ดสั้นเกินไป' };
    const mins = Math.max(5, Math.min(Number(durationMinutes) || 60, 24 * 60));
    const id = Utilities.getUuid();
    const hash = hashPassword(pc + '_SEAT_EDIT_SALT');
    const exp = new Date();
    exp.setMinutes(exp.getMinutes() + mins);

    const sh = getSheet_(SHEETS.SEAT_EDIT_CODES);
    sh.appendRow([id, hash, exp, userId, String(label || ''), false]);
    clearDashboardCaches_();
    return { success: true, codeId: id, expiresAt: exp.toISOString() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatValidateEditCode(plainCode) {
  try {
    const pc = String(plainCode || '').trim();
    if (!pc) return { success: false, message: 'Code is required' };
    const rl = checkRateLimit_('seatEditCode', pc.toUpperCase(), 12, 15 * 60 * 1000);
    if (rl.blocked) return { success: false, message: 'Too many attempts. Please wait and try again later' };
    const hash = hashPassword(pc + '_SEAT_EDIT_SALT');
    const ss = getSpreadsheet();
    const sh = getSheetFromSs_(ss, SHEETS.SEAT_EDIT_CODES);
    const d = sh.getDataRange().getValues();
    const now = new Date();

    for (let i = 1; i < d.length; i++) {
      if (String(d[i][5]) === 'true' || d[i][5] === true) continue;
      if (String(d[i][1]) !== hash) continue;
      if (new Date(d[i][2]) < now) return { success: false, message: 'โค้ดหมดอายุแล้ว' };

      const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
      const exp = new Date(d[i][2]);
      const sess = getSheetFromSs_(ss, SHEETS.SEAT_EDIT_SESSIONS);
      sess.appendRow([token, d[i][0], exp]);
      clearDashboardCaches_();
      return { success: true, token: token, expiresAt: exp.toISOString() };
    }
    return { success: false, message: 'โค้ดไม่ถูกต้อง' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatListEditCodes(userId) {
  try {
    const row = getUserRow_(userId);
    if (!row || !roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์' };
    const sh = getSheet_(SHEETS.SEAT_EDIT_CODES);
    const d = sh.getDataRange().getValues();
    const list = [];
    const now = new Date();
    for (let i = 1; i < d.length; i++) {
      const revoked = d[i][5] === true || String(d[i][5]) === 'true';
      list.push({
        id: d[i][0],
        expiresAt: d[i][2] ? new Date(d[i][2]).toISOString() : '',
        label: d[i][4] || '',
        revoked: revoked,
        active: !revoked && d[i][2] && new Date(d[i][2]) > now
      });
    }
    return { success: true, codes: list };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatRevokeEditCode(userId, codeId) {
  try {
    const row = getUserRow_(userId);
    if (!row || !roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์' };
    const ss = getSpreadsheet();
    const sh = getSheetFromSs_(ss, SHEETS.SEAT_EDIT_CODES);
    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]) === String(codeId)) {
        sh.getRange(i + 1, 6).setValue(true);
        const sess = getSheetFromSs_(ss, SHEETS.SEAT_EDIT_SESSIONS);
        const sd = sess.getDataRange().getValues();
        for (let j = sd.length - 1; j >= 1; j--) {
          if (String(sd[j][1]) === String(codeId)) sess.deleteRow(j + 1);
        }
        clearDashboardCaches_();
        return { success: true };
      }
    }
    return { success: false, message: 'ไม่พบโค้ด' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatRevokeSession(userId, token) {
  try {
    const row = getUserRow_(userId);
    if (!row || !roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์' };
    const sess = getSheet_(SHEETS.SEAT_EDIT_SESSIONS);
    const sd = sess.getDataRange().getValues();
    for (let i = sd.length - 1; i >= 1; i--) {
      if (String(sd[i][0]) === String(token)) sess.deleteRow(i + 1);
    }
    clearDashboardCaches_();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function seatListActiveSessions(userId) {
  try {
    const row = getUserRow_(userId);
    if (!row || !roleIsSeatAdmin_(row.data[6])) return { success: false, message: 'ไม่มีสิทธิ์' };
    const sess = getSheet_(SHEETS.SEAT_EDIT_SESSIONS);
    const sd = sess.getDataRange().getValues();
    const now = new Date();
    const out = [];
    for (let i = 1; i < sd.length; i++) {
      if (new Date(sd[i][2]) > now) {
        out.push({ token: sd[i][0], codeId: sd[i][1], expiresAt: new Date(sd[i][2]).toISOString() });
      }
    }
    return { success: true, sessions: out };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getDashboardData() {
  try {
    const ss = getSpreadsheet();
    const hw = getHomework(ss);
    const tr = getTreasuryItems(ss);
    const lv = getLeaveRequests(ss);
    return { success: true, homework: hw.homework || [], treasury: tr.treasury || [], leaveRequests: lv || [] };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getCounts() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_COUNTS_KEY);
    if (cached) {
      return JSON.parse(cached);
    }

    const ss = getSpreadsheet();
    const hwCount = getSheetFromSs_(ss, SHEETS.HOMEWORK).getLastRow() - 1;
    const trCount = getSheetFromSs_(ss, SHEETS.TREASURY).getLastRow() - 1;
    const lvCount = getSheetFromSs_(ss, SHEETS.LEAVE_REQUESTS).getLastRow() - 1;
    const trPayCounter = getSheetFromSs_(ss, SHEETS.TREASURY_PAYMENTS).getLastRow() - 1;
    const seatMeta = getSheetFromSs_(ss, SHEETS.SEAT_META);
    let seatVersion = 0;
    if (seatMeta && seatMeta.getLastRow() >= 2) {
      seatVersion = Number(seatMeta.getRange(2, 5).getValue()) || 0;
    }
    const result = { success: true, hwCount, trCount, lvCount, trPayCounter, seatVersion };
    
    cache.put(CACHE_COUNTS_KEY, JSON.stringify(result), CACHE_TTL);
    return result;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

