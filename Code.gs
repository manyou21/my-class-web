// ============================================
// 📘 ระบบเว็บการบ้าน & เก็บเงินห้อง + Redeem Code
// ============================================
const SPREADSHEET_ID = '1cT-N8AHw613xstQ7FUNdOk-QnavT9TRsZh1SiJJfYWA';
const DRIVE_FOLDER_ID = '1bv_ZpRpY5diAW9Sd_7Hkl8UCTeHU8YCn';

const SHEETS = {
  USERS: 'Users',
  HOMEWORK: 'Homework',
  HOMEWORK_STATUS: 'HomeworkStatus',
  TREASURY: 'Treasury',
  TREASURY_PAYMENTS: 'TreasuryPayments',
  ABSENCE: 'Absence',
  STUDENT_CODES: 'StudentCodes',
  LEAVE_REQUESTS: 'LeaveRequests',
  REDEEM_CODES: 'RedeemCodes'
};

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
  CLASS_LEADER: { name: 'หัวหน้าห้อง',    canManageHomework: true,  canManageTreasury: false, canApproveLeave: false, canManageCodes: false },
  SECRETARY:    { name: 'เลขานุการ',      canManageHomework: true,  canManageTreasury: false, canApproveLeave: false, canManageCodes: false },
  TREASURER:    { name: 'เหรัญญิก',       canManageHomework: true,  canManageTreasury: true,  canApproveLeave: false, canManageCodes: false },
  TEACHER:      { name: 'ครูที่ปรึกษา',   canManageHomework: false, canManageTreasury: false, canApproveLeave: true,  canManageCodes: false },
  STUDENT:      { name: 'นักเรียน',       canManageHomework: false, canManageTreasury: false, canApproveLeave: false, canManageCodes: false }
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
// 🔥 NEW: API Endpoint for GitHub Pages
// ============================================
function doPost(e) {
  try {
    // อนุญาตให้ทุกโดเมนเรียกใช้งานได้ (CORS)
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);
    
    // แก้ไข CORS Preflight (OPTIONS request) กรณีพิเศษ
    if (e.method === 'OPTIONS') {
      return output;
    }

    // Parse ข้อมูลที่ส่งมา
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const args = data.args || [];

    // ตรวจสอบว่ามีฟังก์ชันนั้นๆ ในระบบหรือไม่
    if (typeof this[action] === 'function') {
      const result = this[action](...args);
      output.setContent(JSON.stringify({ status: 'success', result: result }));
    } else {
      output.setContent(JSON.stringify({ status: 'error', message: 'Function not found: ' + action }));
    }
    return output;
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ต้องมี doGet เพื่อรองรับการทดสอบหรือ Error handling
function doGet(e) {
  return HtmlService.createHtmlOutput('API is running. Please use POST requests.');
}

// ============================================
// 🔧 HELPER FUNCTIONS
// ============================================
function hashPassword(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
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
    let split = base64Data.split(',');
    let type = split[0].split(';')[0].split(':')[1];
    let data = Utilities.base64Decode(split[1]);
    let blob = Utilities.newBlob(data, type, fileName);
    let folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    let file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch (e) {
    console.error('Drive Upload Error: ' + e.toString());
    return null;
  }
}

function convertDriveToDirect(url) {
  if (!url) return '';
  const m = url.match(/[-\w]{25,}/);
  return m ? 'https://drive.google.com/uc?export=view&id=' + m[0] : url;
}

function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss.getSheetByName(SHEETS.USERS)) ss.insertSheet(SHEETS.USERS).appendRow(['ID','DisplayName','Email','PasswordHash','Hint','StudentNo','Role','CreatedAt','LastLogin','TempRoleExpiry','HwCredits']);
  if (!ss.getSheetByName(SHEETS.STUDENT_CODES)) { 
    let sc = ss.insertSheet(SHEETS.STUDENT_CODES); 
    sc.appendRow(['StudentNo','StudentCode','StudentName','IsRegistered']); 
    STUDENTS.forEach(s => sc.appendRow([s.no, s.code, s.name, false])); 
  }
  if (!ss.getSheetByName(SHEETS.HOMEWORK)) ss.insertSheet(SHEETS.HOMEWORK).appendRow(['ID','Subject','Description','AssignedDate','DueDate','NoDueDate','CreatedBy','CreatedAt','SubjectColor']);
  if (!ss.getSheetByName(SHEETS.HOMEWORK_STATUS)) ss.insertSheet(SHEETS.HOMEWORK_STATUS).appendRow(['HomeworkID','StudentNo','Status','ImagePath','CompletedAt','Notes','SubjectColor']);
  if (!ss.getSheetByName(SHEETS.TREASURY)) ss.insertSheet(SHEETS.TREASURY).appendRow(['ID','Title','AmountPerPerson','CreatedBy','CreatedAt','Status']);
  if (!ss.getSheetByName(SHEETS.TREASURY_PAYMENTS)) ss.insertSheet(SHEETS.TREASURY_PAYMENTS).appendRow(['TreasuryID','StudentNo','AmountPaid','PaidAt','Notes']);
  if (!ss.getSheetByName(SHEETS.LEAVE_REQUESTS)) ss.insertSheet(SHEETS.LEAVE_REQUESTS).appendRow(['ID','StudentNo','StudentName','Type','Date','Reason','Status','ProofImage','Confirmed','Timestamp']);
  if (!ss.getSheetByName(SHEETS.REDEEM_CODES)) ss.insertSheet(SHEETS.REDEEM_CODES).appendRow(['Code','ActionType','Value','Details','MaxUses','UsesCount','CreatedBy','CreatedAt']);
  return 'OK';
}

// ============================================
// 🔐 AUTHENTICATION
// ============================================
function registerUser(name, email, pw, cpw, code, hint) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const uSheet = ss.getSheetByName(SHEETS.USERS); 
    const cSheet = ss.getSheetByName(SHEETS.STUDENT_CODES);

    if (!name || !pw) return { success: false, message: 'กรอกข้อมูลไม่ครบ' }; 
    if (pw !== cpw) return { success: false, message: 'รหัสผ่านไม่ตรงกัน' }; 
    const vP = validatePassword(pw); 
    if (!vP.valid) return { success: false, message: vP.m };

    const uData = uSheet.getDataRange().getValues(); 
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === name) return { success: false, message: 'มีผู้ใช้ชื่อนี้แล้ว' }; 
    }
    
    let role = 'STUDENT', studentNo = null;
    if (!code) { 
      if (ROLE_MAPPING[name]) role = ROLE_MAPPING[name]; 
      else return { success: false, message: 'ชื่อนี้ไม่อยู่ในทะเบียน กรุณากรอกรหัสนักเรียน' }; 
    } else {
      const cData = cSheet.getDataRange().getValues(); 
      let found = false;
      for (let i = 1; i < cData.length; i++) { 
        if (String(cData[i][1]).trim() === String(code).trim()) { 
          if (cData[i][3] === true) return { success: false, message: 'รหัสนักเรียนนี้ถูกใช้งานแล้ว' }; 
          if (String(cData[i][2]).trim() !== name.trim()) return { success: false, message: 'ชื่อไม่ตรงกับรหัสนักเรียน' }; 
          studentNo = cData[i][0]; 
          cSheet.getRange(i + 1, 4).setValue(true); 
          found = true; 
          if (ROLE_MAPPING[name]) role = ROLE_MAPPING[name]; 
          break; 
        } 
      }
      if (!found) return { success: false, message: 'ไม่พบรหัสนักเรียนนี้' }; 
    }
    
    uSheet.appendRow([Utilities.getUuid(), name, email || '', hashPassword(pw), hint || '', studentNo, role, new Date(), new Date(), '', 0]);
    return { success: true, message: 'สมัครสมาชิกสำเร็จ! กรุณา Login' };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() }; 
  }
}

function loginUser(id, pw) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const uSheet = ss.getSheetByName(SHEETS.USERS);
    if (!uSheet) return { success: false, message: 'ไม่พบข้อมูลผู้ใช้' }; 
    
    const uData = uSheet.getDataRange().getValues(); 
    const pHash = hashPassword(pw);
    
    for (let i = 1; i < uData.length; i++) {
      const row = uData[i];
      const nameMatch = row[1] === id;
      const emailMatch = row[2] === id && row[2] !== '';
      const pwMatch = row[3] === pHash;

      if ((nameMatch || emailMatch) && pwMatch) {
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
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) { 
    return { success: false, message: 'Server Error: ' + e.toString() }; 
  }
}

function changePassword(uid, curr, newP, conf) {
  try {
    ensureSheetsExist();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const uSheet = ss.getSheetByName(SHEETS.USERS); 
    const uData = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][0] === uid) {
        if (uData[i][3] !== hashPassword(curr)) return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
        if (newP !== conf) return { success: false, message: 'รหัสผ่านใหม่ไม่ตรงกัน' };
        const v = validatePassword(newP); 
        if (!v.valid) return { success: false, message: v.m };
        uSheet.getRange(i + 1, 4).setValue(hashPassword(newP));
        return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function getPasswordHint(username) {
  ensureSheetsExist();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  const data = ss.getSheetByName(SHEETS.USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { 
    if (data[i][1] === username) { 
      return { success: true, hint: data[i][4] || 'ไม่มีคำใบ้' }; 
    } 
  }
  return { success: false, message: 'ไม่พบผู้ใช้นี้' };
}

// ============================================
// 📦 MASTER DATA
// ============================================
function getStudents() { return { success: true, students: STUDENTS }; }
function getSubjects() { return { success: true, subjects: SUBJECTS }; }

// ============================================
// 📝 HOMEWORK
// ============================================
function addHomework(sub, desc, ad, dd, nd, by) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const uSheet = ss.getSheetByName(SHEETS.USERS); 
    const hwS = ss.getSheetByName(SHEETS.HOMEWORK); 
    const stS = ss.getSheetByName(SHEETS.HOMEWORK_STATUS);

    const uData = uSheet.getDataRange().getValues(); 
    let hasPerm = false;

    for (let i = 1; i < uData.length; i++) {
      if (uData[i][1] === by) {
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
        if (!hasPerm && uData[i].length > 9 && uData[i][9]) {
          const expiry = new Date(uData[i][9]);
          if (expiry > new Date()) hasPerm = true;
        }
        break;
      }
    }

    if (!hasPerm) return { success: false, message: 'คุณไม่มีสิทธิ์เพิ่มการบ้าน หรือ Credit หมดแล้ว' };

    const id = Utilities.getUuid();
    const color = getSubjectColor(sub);

    hwS.appendRow([id, sub, desc, ad, dd, nd, by, new Date(), color]);
    hwS.getRange(hwS.getLastRow(), 1, 1, 9).setBackground(color);

    const rows = STUDENTS.map(s => [id, s.no, 'pending', '', '', '', color]);
    if (rows.length > 0) {
      const startRow = stS.getLastRow() + 1;
      stS.getRange(startRow, 1, rows.length, 7).setValues(rows);
      stS.getRange(startRow, 1, rows.length, 7).setBackground(color);
    }

    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function getHomework() {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const hwSheet = ss.getSheetByName(SHEETS.HOMEWORK); 
    const stSheet = ss.getSheetByName(SHEETS.HOMEWORK_STATUS);

    const hwData = hwSheet.getDataRange().getValues(); 
    const stData = stSheet.getDataRange().getValues();

    const list = [];
    for (let i = 1; i < hwData.length; i++) {
      const row = hwData[i];
      if (!row[0]) continue;

      const item = {
        id: row[0],
        subject: row[1] || '(ไม่มีวิชา)',
        description: row[2] || '',
        assignedDate: row[3] ? new Date(row[3]).toISOString() : null,
        dueDate: row[4] ? new Date(row[4]).toISOString() : null,
        noDueDate: !!row[5],
        createdBy: row[6] || '',
        color: row[8] || getSubjectColor(row[1]),
        statuses: {}
      };

      for (let j = 1; j < stData.length; j++) {
        if (stData[j][0] === row[0]) {
          item.statuses[stData[j][1]] = {
            status: stData[j][2] || 'pending',
            imagePath: stData[j][3] || ''
          };
        }
      }
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const stS = ss.getSheetByName(SHEETS.HOMEWORK_STATUS); 
    const d = stS.getDataRange().getValues();

    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === hid && String(d[i][1]) === String(sno)) {
        if (stat) stS.getRange(i + 1, 3).setValue(stat);
        if (imgBase64) {
          var url = uploadImageToDrive(imgBase64, 'HW_' + sno + '_' + hid);
          if (url) stS.getRange(i + 1, 4).setValue(url);
        }
        if (stat === 'completed') stS.getRange(i + 1, 5).setValue(new Date());
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function deleteHomework(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const hS = ss.getSheetByName(SHEETS.HOMEWORK); 
    const sS = ss.getSheetByName(SHEETS.HOMEWORK_STATUS);

    var hd = hS.getDataRange().getValues();
    for (let i = hd.length - 1; i >= 1; i--) {
      if (hd[i][0] === id) hS.deleteRow(i + 1);
    }
    var sd = sS.getDataRange().getValues();
    for (let i = sd.length - 1; i >= 1; i--) {
      if (sd[i][0] === id) sS.deleteRow(i + 1);
    }
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

// ============================================
// 🚪 LEAVE REQUESTS
// ============================================
function submitLeaveRequest(studentNo, studentName, type, date, reason, base64Image) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEETS.LEAVE_REQUESTS); 
    const id = Utilities.getUuid(); 
    let imageUrl = '';

    if (base64Image) {
      try { 
        var split = base64Image.split(','); 
        var mimeMatch = split[0].match(/:(.*?);/); 
        var contentType = mimeMatch ? mimeMatch[1] : 'image/jpeg'; 
        var base64Data = split[1]; 
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, 'leave_' + studentNo + '_' + id); 
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID); 
        var file = folder.createFile(blob); 
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
        imageUrl = file.getUrl(); 
      } catch (imgErr) { 
        Logger.log('Image upload error: ' + imgErr.toString()); 
      }
    }

    sheet.appendRow([id, studentNo, studentName, type, date, reason, 'PENDING', imageUrl, false, new Date()]); 
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function getLeaveRequests() {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEETS.LEAVE_REQUESTS); 
    const data = sheet.getDataRange().getValues(); 
    if (data.length <= 1) return [];
    return data.slice(1).map(row => {
      return {
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
      };
    });
  } catch (e) { 
    Logger.log('getLeaveRequests Error: ' + e.toString()); 
    return []; 
  }
}

function confirmLeaveRequest(id, imgBase64) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEETS.LEAVE_REQUESTS); 
    const data = sheet.getDataRange().getValues(); 
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        var url = uploadImageToDrive(imgBase64, 'LeaveConfirm_' + id);
        if (url) sheet.getRange(i + 1, 8).setValue(url);
        sheet.getRange(i + 1, 9).setValue(true);
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function updateLeaveStatus(id, status) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEETS.LEAVE_REQUESTS); 
    const data = sheet.getDataRange().getValues(); 
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 7).setValue(status);
        sheet.getRange(i + 1, 9).setValue(status === 'APPROVED');
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

// ============================================
// 💰 TREASURY
// ============================================
function addTreasuryItem(title, amt, by) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const tS = ss.getSheetByName(SHEETS.TREASURY); 
    const pS = ss.getSheetByName(SHEETS.TREASURY_PAYMENTS); 
    const id = Utilities.getUuid(); 
    tS.appendRow([id, title, amt, by, new Date(), 'active']);
    const rows = STUDENTS.map(s => [id, s.no, 0, '', '']);
    if (rows.length > 0) pS.getRange(pS.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    return { success: true, treasuryId: id };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function getTreasuryItems() {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const tSheet = ss.getSheetByName(SHEETS.TREASURY); 
    const pSheet = ss.getSheetByName(SHEETS.TREASURY_PAYMENTS);

    const tData = tSheet.getDataRange().getValues(); 
    const pData = pSheet ? pSheet.getDataRange().getValues() : [];

    const payMap = {};
    for (let r = 1; r < pData.length; r++) {
      const tId = pData[r][0];
      if (!tId) continue;
      if (!payMap[tId]) payMap[tId] = {};
      payMap[tId][pData[r][1]] = { amountPaid: parseFloat(pData[r][2]) || 0 };
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
        summary: { completedCount: countDone }
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const pS = ss.getSheetByName(SHEETS.TREASURY_PAYMENTS); 
    const d = pS.getDataRange().getValues();

    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === tid && String(d[i][1]) === String(sno)) {
        pS.getRange(i + 1, 3).setValue(paid);
        pS.getRange(i + 1, 4).setValue(paid > 0 ? new Date() : '');
        const props = PropertiesService.getScriptProperties();
        const count = parseInt(props.getProperty('tr_pay_counter') || '0');
        props.setProperty('tr_pay_counter', (count + 1).toString());
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function setTreasuryMembers(treasuryId, studentNos, payerStudentNo, isPaid) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tSheet = ss.getSheetByName(SHEETS.TREASURY);
    const pSheet = ss.getSheetByName(SHEETS.TREASURY_PAYMENTS);
    
    const tData = tSheet.getDataRange().getValues();
    let amount = 0;
    for (let i = 1; i < tData.length; i++) {
      if (tData[i][0] === treasuryId) {
        amount = parseFloat(tData[i][2]) || 0;
        break;
      }
    }
    
    const pData = pSheet.getDataRange().getValues();
    for (let i = pData.length - 1; i >= 1; i--) {
      if (pData[i][0] === treasuryId) {
        pSheet.deleteRow(i + 1);
      }
    }
    
    if (studentNos && studentNos.length > 0) {
      const rows = studentNos.map(sNo => {
        const paid = (String(sNo) === String(payerStudentNo) && isPaid) ? amount : 0;
        return [treasuryId, sNo, paid, paid > 0 ? new Date() : '', ''];
      });
      pSheet.getRange(pSheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteTreasuryItem(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const tS = ss.getSheetByName(SHEETS.TREASURY); 
    const pS = ss.getSheetByName(SHEETS.TREASURY_PAYMENTS);

    var td = tS.getDataRange().getValues();
    for (let i = td.length - 1; i >= 1; i--) {
      if (td[i][0] === id) tS.deleteRow(i + 1);
    }
    var pd = pS.getDataRange().getValues();
    for (let i = pd.length - 1; i >= 1; i--) {
      if (pd[i][0] === id) pS.deleteRow(i + 1);
    }
    return { success: true };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

// ============================================
// 🔄 REAL-TIME POLLING
// ============================================
function getLastUpdate() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hw = ss.getSheetByName(SHEETS.HOMEWORK);
    const tr = ss.getSheetByName(SHEETS.TREASURY);
    const lv = ss.getSheetByName(SHEETS.LEAVE_REQUESTS);

    let pCount = 0;
    let lastPendingType = '';

    if (lv && lv.getLastRow() > 1) {
      const vals = lv.getRange(2, 7, lv.getLastRow() - 1, 1).getValues();
      pCount = vals.filter(r => r[0] === 'PENDING').length;
      const lastRow = lv.getLastRow();
      const lastVals = lv.getRange(lastRow, 4, 1, 4).getValues()[0];
      if (lastVals[3] === 'PENDING') lastPendingType = lastVals[0] || 'รายการ';
    }

    const props = PropertiesService.getScriptProperties();
    const trPayCounter = parseInt(props.getProperty('tr_pay_counter') || '0');

    return {
      success: true,
      hwCount: hw ? Math.max(0, hw.getLastRow() - 1) : 0,
      trCount: tr ? Math.max(0, tr.getLastRow() - 1) : 0,
      pendingLeaveCount: pCount,
      lastPendingType: lastPendingType,
      trPayCounter: trPayCounter
    };
  } catch (e) { 
    return { success: false, hwCount: 0, trCount: 0, pendingLeaveCount: 0, lastPendingType: '', trPayCounter: 0 }; 
  }
}

// ============================================
// 🎟️ REDEEM CODE
// ============================================
function generateCode(actionType, value, details) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEETS.REDEEM_CODES); 
    const code = Utilities.getUuid().split('-')[0].toUpperCase();
    sheet.appendRow([code, actionType, value, JSON.stringify(details || {}), 1, 0, 'Owner', new Date()]);
    return { success: true, code: code };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function redeemCode(code, userId) {
  try {
    ensureSheetsExist(); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const codeSheet = ss.getSheetByName(SHEETS.REDEEM_CODES); 
    const userSheet = ss.getSheetByName(SHEETS.USERS);

    const codeData = codeSheet.getDataRange().getValues();

    for (let i = 1; i < codeData.length; i++) {
      if (String(codeData[i][0]).toUpperCase() === String(code).toUpperCase()) {
        const maxUses = codeData[i][4];
        const usesCount = codeData[i][5];
        if (usesCount >= maxUses) return { success: false, message: 'โค้ดนี้ถูกใช้งานครบแล้ว' };

        const action = codeData[i][1];
        const val = codeData[i][2];
        const details = JSON.parse(codeData[i][3] || '{}');

        codeSheet.getRange(i + 1, 6).setValue(usesCount + 1);

        let resultMsg = 'ใช้โค้ดสำเร็จ';
        let extraData = {};
        const uData = userSheet.getDataRange().getValues();

        for (let j = 1; j < uData.length; j++) {
          if (uData[j][0] === userId) {

            if (action === 'TEMP_ACCESS') {
              const parts = String(val).split('_');
              const num = parseInt(parts[0]) || 1;
              const unit = parts[1] || 'DAYS';
              const expiry = new Date();
              if (unit === 'DAYS') expiry.setDate(expiry.getDate() + num);
              else if (unit === 'WEEKS') expiry.setDate(expiry.getDate() + (num * 7));
              else if (unit === 'MONTHS') expiry.setMonth(expiry.getMonth() + num);
              else if (unit === 'YEARS') expiry.setFullYear(expiry.getFullYear() + num);
              userSheet.getRange(j + 1, 10).setValue(expiry);
              resultMsg = 'ได้รับสิทธิ์ผู้ใช้ชั่วคราว ' + num + ' ' + unit;
            }
            else if (action === 'ADD_HW') {
              const cur = (uData[j].length > 10 && uData[j][10]) ? Number(uData[j][10]) : 0;
              userSheet.getRange(j + 1, 11).setValue(cur + parseInt(val));
              resultMsg = 'ได้รับเครดิตตั้งการบ้าน ' + val + ' ครั้ง';
            }
            else if (action === 'ADD_MONEY') {
              const tSheet = ss.getSheetByName(SHEETS.TREASURY);
              const tId = Utilities.getUuid();
              const title = details.title || 'รายการพิเศษ';
              const amt = parseFloat(details.amount) || 0;
              tSheet.appendRow([tId, title, amt, 'Code', new Date(), 'active']);
              resultMsg = 'สร้างรายการเก็บเงินกลุ่มแล้ว: ' + title;
              extraData.treasuryId = tId;
              extraData.title = title;
              extraData.amount = amt;
            }
            break;
          }
        }
        return { success: true, message: resultMsg, data: extraData };
      }
    }
    return { success: false, message: 'ไม่พบโค้ดนี้' };
  } catch (e) { 
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() }; 
  }
}