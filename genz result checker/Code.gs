
const GOOGLE_SHEET_ID = '1zyjmnpFIgtpIhrYhJvqo668yqjIQbGCbaxJhvK1LWzQ';

const SHEET_NAMES = {
  SETTINGS: 'SETTINGS',
  ADMINS: 'ADMINS',
  STUDENTS: 'STUDENTS',
  RESULTS: 'RESULTS',
  REMARKS: 'REMARKS'
};

function getSpreadsheet_() {
  return SpreadsheetApp.openById(GOOGLE_SHEET_ID);
}

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  setupSystem_();
  let params = {};

  try {
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (postErr) {
        params = e.parameter || {};
      }
    } else {
      params = (e && e.parameter) ? e.parameter : {};
    }

    const action = String(params.action || 'ping').trim();
    const callback = String(params.callback || '').trim();
    const payload = parseJson_(params.payload);
    let result = { success: false, message: 'Invalid action.' };

    switch (action) {
      case 'ping':
        result = {
          success: true,
          message: 'Backend is live.',
          timestamp: isoNow_()
        };
        break;
      case 'adminLogin':
        result = adminLogin_(String(params.username || payload.username || ''), String(params.password || payload.password || ''));
        break;
      case 'getAdminBootstrap':
        result = getAdminBootstrap_(String(params.token || payload.token || ''));
        break;
      case 'saveSettings':
        result = saveSettings_(String(params.token || payload.token || ''), payload);
        break;
      case 'saveStudent':
        result = saveStudent_(String(params.token || payload.token || ''), payload);
        break;
      case 'saveResult':
        result = saveResult_(String(params.token || payload.token || ''), payload);
        break;
      case 'studentSignup':
        result = studentSignup_(payload);
        break;
      case 'studentLogin':
        result = getStudentReport_(String(params.regId || payload.regId || payload.RegID || ''), String(params.password || payload.password || ''));
        break;
      default:
        result = { success: false, message: 'Unknown action.' };
    }

    return respond_(result, callback);
  } catch (err) {
    return respond_({
      success: false,
      message: err && err.message ? err.message : 'Server error.'
    }, params.callback || '');
  }
}

function respond_(obj, callback) {
  const text = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJson_(value) {
  if (!value) return {};
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(value);
  } catch (err) {
    return {};
  }
}

function setupSystem() {
  return setupSystem_();
}

function setupSystem_() {
  ensureSheet_(SHEET_NAMES.SETTINGS, ['Key', 'Value']);
  ensureSheet_(SHEET_NAMES.ADMINS, ['Username', 'PasswordHash', 'DisplayName', 'Active', 'CreatedAt']);
  ensureSheet_(SHEET_NAMES.STUDENTS, [
    'RegID', 'PasswordHash', 'FullName', 'Age', 'Gender', 'DOB',
    'ParentName', 'ParentPhone', 'ParentEmail',
    'SchoolName', 'State', 'CityLGA', 'ClassLevel', 'Category',
    'Address', 'PassportUrl', 'Active', 'CreatedAt', 'UpdatedAt'
  ]);
  ensureSheet_(SHEET_NAMES.RESULTS, [
    'ResultID', 'RegID', 'ExamTitle', 'ExamDate', 'MaxScore',
    'PassMarkNumber', 'PassMarkPercentage', 'StudentScore', 'Percentage',
    'Position', 'Grade', 'Remark', 'TeacherComment',
    'AcademicSession', 'Term', 'Published', 'CreatedAt', 'UpdatedAt'
  ]);
  ensureSheet_(SHEET_NAMES.REMARKS, ['MinPercent', 'MaxPercent', 'BandLabel', 'Remark']);

  bootstrapSettings_();
  bootstrapAdmins_();
  bootstrapRemarks_();

  return { success: true, message: 'System setup completed.' };
}

function ensureSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    let changed = false;
    headers.forEach(function(h, i) {
      if (current[i] !== h) changed = true;
    });
    if (changed) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0ff');
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function bootstrapSettings_() {
  const defaults = {
    BRAND_NAME: 'Genz Edutech Innovations',
    HEAD_OFFICE_ADDRESS: 'Update head office address in admin settings',
    SCHOOL_NAME: 'Genz Result Portal',
    SCHOOL_PHONE: '',
    SCHOOL_EMAIL: '',
    PRINCIPAL_NAME: '',
    ACADEMIC_SESSION: '2025/2026',
    TERM: 'First Term',
    EXAM_TITLE: 'End of Term Examination',
    PASS_MARK_NUMBER: '50',
    PASS_MARK_PERCENTAGE: '50',
    SHOW_POSITION_ON_REPORT: 'false'
  };
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  const existing = getSettings_();
  Object.keys(defaults).forEach(function(key) {
    if (!(key in existing)) sheet.appendRow([key, defaults[key]]);
  });
}

function bootstrapAdmins_() {
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  if (sh.getLastRow() <= 1) {
    sh.appendRow(['admin', hashText_('admin12345'), 'Main Admin', true, isoNow_()]);
  }
}

function bootstrapRemarks_() {
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.REMARKS);
  if (sh.getLastRow() > 1) return;
  const rows = [
    [0, 29, 'F', 'Very poor performance. Serious improvement is required.'],
    [30, 39, 'E', 'Poor performance. More effort and guidance are needed.'],
    [40, 49, 'D', 'Below average performance. Keep working harder.'],
    [50, 59, 'C', 'Average performance. There is room for growth.'],
    [60, 69, 'B', 'Good performance. Keep improving steadily.'],
    [70, 79, 'B+', 'Very good performance. Strong effort shown.'],
    [80, 89, 'A', 'Excellent performance. Keep it up.'],
    [90, 100, 'A+', 'Outstanding performance. Exceptional work.']
  ];
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

function getSettings_() {
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  const lastRow = sh.getLastRow();
  const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const out = {};
  data.forEach(function(r) {
    if (r[0]) out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function upsertSetting_(key, value) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const idx = data.findIndex(function(r) { return String(r[0]).trim() === key; });
  if (idx >= 0) sh.getRange(idx + 2, 2).setValue(value);
  else sh.appendRow([key, value]);
}

function adminLogin_(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();
  if (!username || !password) return { success: false, message: 'Username and password are required.' };

  const rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS));
  const found = rows.find(function(r) {
    return String(r.Username || '').toLowerCase() === username.toLowerCase() &&
      String(r.PasswordHash || '') === hashText_(password) &&
      normalizeBoolean_(r.Active, true);
  });
  if (!found) return { success: false, message: 'Invalid admin login details.' };

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_SESSION_' + token, username, 21600);
  return {
    success: true,
    message: 'Login successful.',
    token: token,
    displayName: found.DisplayName || username
  };
}

function requireAdmin_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Admin session not found.');
  const user = CacheService.getScriptCache().get('ADMIN_SESSION_' + token);
  if (!user) throw new Error('Admin session expired. Please log in again.');
  return user;
}

function getAdminBootstrap_(token) {
  requireAdmin_(token);
  const settings = getSettings_();
  const students = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS)).map(cleanStudent_);
  const results = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS)).map(cleanResult_);
  results.sort(function(a,b) {
    return new Date(b.UpdatedAt || b.CreatedAt || 0) - new Date(a.UpdatedAt || a.CreatedAt || 0);
  });
  return {
    success: true,
    settings: settings,
    students: students,
    results: results,
    categories: getCategoryOptions_(),
    summary: {
      studentCount: students.length,
      resultCount: results.length,
      publishedCount: results.filter(function(r) { return r.Published; }).length
    }
  };
}

function saveSettings_(token, payload) {
  requireAdmin_(token);
  const keys = ['BRAND_NAME','HEAD_OFFICE_ADDRESS','SCHOOL_NAME','SCHOOL_PHONE','SCHOOL_EMAIL','PRINCIPAL_NAME','ACADEMIC_SESSION','TERM','EXAM_TITLE','PASS_MARK_NUMBER','PASS_MARK_PERCENTAGE','SHOW_POSITION_ON_REPORT'];
  keys.forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(payload, key)) {
      upsertSetting_(key, payload[key]);
    }
  });
  return { success: true, message: 'Settings saved successfully.', settings: getSettings_() };
}

function studentSignup_(payload) {
  const regId = String(payload.RegID || payload.regId || '').trim();
  const password = String(payload.Password || payload.password || '').trim();
  if (!regId || !password) throw new Error('Registration ID and password are required.');
  const required = ['FullName','Age','Gender','ParentEmail','ParentPhone','SchoolName','State','CityLGA','ClassLevel','Category'];
  required.forEach(function(key) {
    if (!String(payload[key] || '').trim()) throw new Error(key + ' is required.');
  });

  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  const existing = findRowByValue_(sh, 1, regId);
  if (existing) throw new Error('This Registration ID already exists.');

  sh.appendRow([
    regId,
    hashText_(password),
    sanitizeValue_(payload.FullName),
    sanitizeValue_(payload.Age),
    sanitizeValue_(payload.Gender),
    sanitizeValue_(payload.DOB),
    sanitizeValue_(payload.ParentName),
    sanitizeValue_(payload.ParentPhone),
    sanitizeValue_(payload.ParentEmail),
    sanitizeValue_(payload.SchoolName),
    sanitizeValue_(payload.State),
    sanitizeValue_(payload.CityLGA),
    sanitizeValue_(payload.ClassLevel),
    sanitizeValue_(payload.Category),
    sanitizeValue_(payload.Address),
    '',
    true,
    isoNow_(),
    isoNow_()
  ]);

  return { success: true, message: 'Signup successful. You can now log in to check your result.' };
}

function saveStudent_(token, payload) {
  requireAdmin_(token);
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  const headers = getHeaders_(sh);
  const regId = String(payload.RegID || '').trim();
  const password = String(payload.Password || '').trim();
  if (!regId) throw new Error('Registration ID is required.');

  const existingRow = findRowByValue_(sh, 1, regId);
  if (!existingRow && !password) throw new Error('Password is required for new student.');

  const currentHash = existingRow ? sh.getRange(existingRow, headers.indexOf('PasswordHash') + 1).getValue() : '';
  const createdAt = existingRow ? sh.getRange(existingRow, headers.indexOf('CreatedAt') + 1).getValue() : isoNow_();

  const record = {
    RegID: regId,
    PasswordHash: password ? hashText_(password) : currentHash,
    FullName: sanitizeValue_(payload.FullName),
    Age: sanitizeValue_(payload.Age),
    Gender: sanitizeValue_(payload.Gender),
    DOB: sanitizeValue_(payload.DOB),
    ParentName: sanitizeValue_(payload.ParentName),
    ParentPhone: sanitizeValue_(payload.ParentPhone),
    ParentEmail: sanitizeValue_(payload.ParentEmail),
    SchoolName: sanitizeValue_(payload.SchoolName),
    State: sanitizeValue_(payload.State),
    CityLGA: sanitizeValue_(payload.CityLGA),
    ClassLevel: sanitizeValue_(payload.ClassLevel),
    Category: sanitizeValue_(payload.Category),
    Address: sanitizeValue_(payload.Address),
    PassportUrl: sanitizeValue_(payload.PassportUrl),
    Active: normalizeBoolean_(payload.Active, true),
    CreatedAt: createdAt,
    UpdatedAt: isoNow_()
  };

  const values = headers.map(function(h) { return record[h]; });
  if (existingRow) sh.getRange(existingRow, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);

  return { success: true, message: existingRow ? 'Student updated successfully.' : 'Student created successfully.' };
}

function saveResult_(token, payload) {
  requireAdmin_(token);
  const settings = getSettings_();
  const sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  const headers = getHeaders_(sh);

  const regId = String(payload.RegID || '').trim();
  if (!regId) throw new Error('Registration ID is required.');
  const student = getStudentByRegId_(regId);
  if (!student) throw new Error('Student not found.');

  const resultId = String(payload.ResultID || '').trim() || Utilities.getUuid();
  const examTitle = String(payload.ExamTitle || settings.EXAM_TITLE || '').trim();
  const examDate = String(payload.ExamDate || '').trim();
  const maxScore = asNumber_(payload.MaxScore);
  const studentScore = asNumber_(payload.StudentScore);
  const passMarkNumber = payload.PassMarkNumber === '' || payload.PassMarkNumber == null ? asNumber_(settings.PASS_MARK_NUMBER) : asNumber_(payload.PassMarkNumber);
  const passMarkPercentage = payload.PassMarkPercentage === '' || payload.PassMarkPercentage == null ? asNumber_(settings.PASS_MARK_PERCENTAGE) : asNumber_(payload.PassMarkPercentage);
  if (!examTitle || !examDate) throw new Error('Exam title and exam date are required.');
  if (maxScore <= 0) throw new Error('Overall exam score must be greater than zero.');

  const percentage = round2_((studentScore / maxScore) * 100);
  const grade = getGradeFromPercentage_(percentage);
  const remark = getRemarkFromPercentage_(percentage);
  const existingRow = findRowByValue_(sh, 1, resultId);
  const createdAt = existingRow ? sh.getRange(existingRow, headers.indexOf('CreatedAt') + 1).getValue() : isoNow_();

  const record = {
    ResultID: resultId,
    RegID: regId,
    ExamTitle: examTitle,
    ExamDate: examDate,
    MaxScore: maxScore,
    PassMarkNumber: passMarkNumber,
    PassMarkPercentage: passMarkPercentage,
    StudentScore: studentScore,
    Percentage: percentage,
    Position: sanitizeValue_(payload.Position),
    Grade: grade,
    Remark: remark,
    TeacherComment: sanitizeValue_(payload.TeacherComment),
    AcademicSession: sanitizeValue_(payload.AcademicSession || settings.ACADEMIC_SESSION || ''),
    Term: sanitizeValue_(payload.Term || settings.TERM || ''),
    Published: normalizeBoolean_(payload.Published, true),
    CreatedAt: createdAt,
    UpdatedAt: isoNow_()
  };

  const values = headers.map(function(h) { return record[h]; });
  if (existingRow) sh.getRange(existingRow, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);

  return { success: true, message: existingRow ? 'Result updated successfully.' : 'Result saved successfully.' };
}

function getStudentReport_(regId, password) {
  regId = String(regId || '').trim();
  password = String(password || '').trim();
  if (!regId || !password) return { success: false, message: 'Registration ID and password are required.' };

  const rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS));
  const student = rows.find(function(r) {
    return String(r.RegID || '').trim() === regId &&
      String(r.PasswordHash || '') === hashText_(password) &&
      normalizeBoolean_(r.Active, true);
  });
  if (!student) return { success: false, message: 'Invalid student login details.' };

  const settings = getSettings_();
  const results = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS))
    .filter(function(r) {
      return String(r.RegID || '').trim() === regId && normalizeBoolean_(r.Published, false);
    })
    .map(cleanResult_);

  results.sort(function(a,b) {
    return new Date(b.ExamDate || b.UpdatedAt || 0) - new Date(a.ExamDate || a.UpdatedAt || 0);
  });

  return {
    success: true,
    message: results.length ? 'Result loaded successfully.' : 'Signup/login successful, but no published result was found yet.',
    settings: settings,
    student: cleanStudent_(student),
    results: results
  };
}

function getStudentByRegId_(regId) {
  const rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS));
  return rows.find(function(r) { return String(r.RegID || '').trim() === regId; }) || null;
}

function cleanStudent_(r) {
  return {
    RegID: r.RegID || '',
    FullName: r.FullName || '',
    Age: r.Age || '',
    Gender: r.Gender || '',
    DOB: r.DOB || '',
    ParentName: r.ParentName || '',
    ParentPhone: r.ParentPhone || '',
    ParentEmail: r.ParentEmail || '',
    SchoolName: r.SchoolName || '',
    State: r.State || '',
    CityLGA: r.CityLGA || '',
    ClassLevel: r.ClassLevel || '',
    Category: r.Category || '',
    Address: r.Address || '',
    PassportUrl: r.PassportUrl || '',
    Active: normalizeBoolean_(r.Active, true),
    CreatedAt: r.CreatedAt || '',
    UpdatedAt: r.UpdatedAt || ''
  };
}

function cleanResult_(r) {
  return {
    ResultID: r.ResultID || '',
    RegID: r.RegID || '',
    ExamTitle: r.ExamTitle || '',
    ExamDate: r.ExamDate || '',
    MaxScore: asNumber_(r.MaxScore),
    PassMarkNumber: asNumber_(r.PassMarkNumber),
    PassMarkPercentage: asNumber_(r.PassMarkPercentage),
    StudentScore: asNumber_(r.StudentScore),
    Percentage: asNumber_(r.Percentage),
    Position: r.Position || '',
    Grade: r.Grade || '',
    Remark: r.Remark || '',
    TeacherComment: r.TeacherComment || '',
    AcademicSession: r.AcademicSession || '',
    Term: r.Term || '',
    Published: normalizeBoolean_(r.Published, false),
    CreatedAt: r.CreatedAt || '',
    UpdatedAt: r.UpdatedAt || '',
    Status: getResultStatus_(asNumber_(r.StudentScore), asNumber_(r.Percentage), asNumber_(r.PassMarkNumber), asNumber_(r.PassMarkPercentage))
  };
}

function getCategoryOptions_() {
  return ['Lower Primary', 'Upper Primary', 'Junior High School', 'Senior High School', 'Undergraduate'];
}

function getRemarkFromPercentage_(percentage) {
  const bands = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.REMARKS)).map(function(r) {
    return { MinPercent: asNumber_(r.MinPercent), MaxPercent: asNumber_(r.MaxPercent), Remark: r.Remark || '' };
  });
  const band = bands.find(function(item) {
    return percentage >= item.MinPercent && percentage <= item.MaxPercent;
  });
  return band ? band.Remark : 'No remark configured for this score band.';
}

function getGradeFromPercentage_(percentage) {
  if (percentage >= 90) return 'A+';
  if (percentage >= 80) return 'A';
  if (percentage >= 70) return 'B+';
  if (percentage >= 60) return 'B';
  if (percentage >= 50) return 'C';
  if (percentage >= 40) return 'D';
  if (percentage >= 30) return 'E';
  return 'F';
}

function getResultStatus_(score, percentage, passMarkNumber, passMarkPercentage) {
  return (score >= passMarkNumber && percentage >= passMarkPercentage) ? 'PASS' : 'FAIL';
}

function getSheetObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return rows.filter(function(row) {
    return row.some(function(cell) { return cell !== ''; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function findRowByValue_(sheet, col, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues().flat();
  const idx = values.findIndex(function(v) { return String(v).trim() === String(value).trim(); });
  return idx >= 0 ? idx + 2 : null;
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function hashText_(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''));
  return digest.map(function(byte) {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function asNumber_(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function normalizeBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  if (value === 'TRUE' || value === 'true' || value === 1 || value === '1') return true;
  if (value === 'FALSE' || value === 'false' || value === 0 || value === '0') return false;
  return defaultValue;
}

function sanitizeValue_(value) {
  return value == null ? '' : String(value).trim();
}

function isoNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}
