const GOOGLE_SHEET_ID = '1eOq3_XOmL8NqMbXIC3ysa0d3K8B9jQdmT9z6y-iBiOw';
const SESSION_TTL_SECONDS = 60 * 60 * 6;
const REQUEST_TTL_SECONDS = 60 * 5;
const RESET_CODE_TTL_MINUTES = 15;
const PASSWORD_HASH_VERSION = 'v2';
const PASSWORD_HASH_ROUNDS = 4000;

const SHEET_NAMES = {
  SETTINGS: 'SETTINGS',
  ADMINS: 'ADMINS',
  STUDENTS: 'STUDENTS',
  RESULTS: 'RESULTS',
  REMARKS: 'REMARKS',
  RESET_CODES: 'RESET_CODES',
  AUDIT_LOGS: 'AUDIT_LOGS'
};

const HEADERS = {
  SETTINGS: ['Key', 'Value'],
  ADMINS: ['Username', 'PasswordHash', 'PasswordSalt', 'PasswordVersion', 'DisplayName', 'Email', 'Phone', 'IsPrincipal', 'Active', 'Archived', 'Deleted', 'CreatedAt', 'UpdatedAt', 'LastLoginAt'],
  STUDENTS: ['RegID', 'PasswordHash', 'PasswordSalt', 'PasswordVersion', 'FullName', 'Age', 'Gender', 'DOB', 'ParentName', 'ParentPhone', 'ParentEmail', 'SchoolName', 'State', 'CityLGA', 'ClassLevel', 'Category', 'Address', 'PassportUrl', 'Active', 'Archived', 'Deleted', 'CreatedAt', 'UpdatedAt', 'LastLoginAt'],
  RESULTS: ['ResultID', 'RegID', 'ExamCode', 'ExamTitle', 'Subject', 'ExamDate', 'MaxScore', 'PassMarkNumber', 'PassMarkPercentage', 'StudentScore', 'Percentage', 'Position', 'Grade', 'Remark', 'TeacherComment', 'AcademicSession', 'Term', 'Published', 'ViewActive', 'PublishedAt', 'SignatureUrl', 'Archived', 'Deleted', 'CreatedAt', 'UpdatedAt'],
  REMARKS: ['MinPercent', 'MaxPercent', 'BandLabel', 'Remark'],
  RESET_CODES: ['Role', 'Identifier', 'CodeHash', 'ExpiresAt', 'Consumed', 'CreatedAt'],
  AUDIT_LOGS: ['Timestamp', 'ActorRole', 'ActorId', 'Action', 'Status', 'Details']
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
  var params = extractParams_(e);
  var action = sanitizeValue_(params.action || 'ping');
  try {
    validateAction_(action);
    if (!shouldBypassFreshRequest_(action)) {
      validateFreshRequest_(params, action);
    }

    var payload = buildPayload_(params);
    var result;
    switch (action) {
      case 'ping':
        result = ok_('Backend is live.', { timestamp: isoNow_() });
        break;
      case 'getPublicBootstrap':
        result = getPublicBootstrap_();
        break;
      case 'adminLogin':
        result = adminLogin_(payload);
        break;
      case 'adminSignup':
        result = adminSignup_(payload);
        break;
      case 'studentSignup':
        result = studentSignup_(payload);
        break;
      case 'studentLogin':
        result = studentLogin_(payload);
        break;
      case 'validateSession':
        result = validateSession_(payload);
        break;
      case 'getAdminBootstrap':
        result = getAdminBootstrap_(payload);
        break;
      case 'saveSettings':
        result = saveSettings_(payload);
        break;
      case 'saveStudent':
        result = saveStudent_(payload);
        break;
      case 'uploadStudentPassport':
        result = uploadStudentPassport_(payload);
        break;
      case 'uploadBrandingAsset':
        result = uploadBrandingAsset_(payload);
        break;
      case 'setStudentState':
        result = setStudentState_(payload);
        break;
      case 'importStudentsCsv':
        result = importStudentsCsv_(payload);
        break;
      case 'exportStudentsCsv':
        result = exportStudentsCsv_(payload);
        break;
      case 'bulkUpdatePassports':
        result = bulkUpdatePassports_(payload);
        break;
      case 'importPassportsCsv':
        result = importPassportsCsv_(payload);
        break;
      case 'saveResult':
        result = saveResultEntries_({ token: payload.token, entries: [payload] });
        break;
      case 'saveResultEntries':
        result = saveResultEntries_(payload);
        break;
      case 'importResultsCsv':
        result = importResultsCsv_(payload);
        break;
      case 'recalculateResults':
        result = recalculateResults_(payload);
        break;
      case 'setResultState':
        result = setResultState_(payload);
        break;
      case 'bulkSetResultState':
        result = bulkSetResultState_(payload);
        break;
      case 'exportResultsCsv':
        result = exportResultsCsv_(payload);
        break;
      case 'loadStudentExamCodes':
        result = loadStudentExamCodes_(payload);
        break;
      case 'loadStudentResults':
        result = loadStudentResults_(payload);
        break;
      case 'getImageDataUrl':
        result = getImageDataUrl_(payload);
        break;
      case 'generateStudentPdf':
        result = generateStudentPdf_(payload);
        break;
      case 'sendBulkEmail':
        result = sendBulkEmail_(payload);
        break;
      case 'authorizeMailStatus':
        result = authorizeMailStatus_(payload);
        break;
      case 'exportSmsContacts':
        result = exportSmsContacts_(payload);
        break;
      case 'requestPasswordReset':
        result = requestPasswordReset_(payload);
        break;
      case 'resetPassword':
        result = resetPassword_(payload);
        break;
      case 'setAdminState':
        result = setAdminState_(payload);
        break;
      default:
        result = fail_('Unknown action.');
    }
    return respond_(result);
  } catch (err) {
    logAuditSafe_(action, params, 'ERROR', err);
    return respond_(fail_(err && err.message ? err.message : 'Server error.'));
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ok_(message, data) {
  return { ok: true, message: message || 'Done.', data: data || {} };
}

function fail_(message, data) {
  return { ok: false, message: message || 'Request failed.', data: data || {} };
}

function extractParams_(e) {
  var params = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function(key) {
      params[key] = e.parameter[key];
    });
  }
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && typeof body === 'object') {
        Object.keys(body).forEach(function(key) { params[key] = body[key]; });
      }
    } catch (err) {}
  }
  return params;
}

function buildPayload_(params) {
  var payload = parseJson_(params.payload);
  if (!payload || typeof payload !== 'object') payload = {};
  ['token', 'clientId', 'requestTs', 'requestNonce'].forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(payload, key) && params[key] != null) payload[key] = params[key];
  });
  return payload;
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

function validateAction_(action) {
  var allowed = {
    ping: true, getPublicBootstrap: true, adminLogin: true, adminSignup: true, studentSignup: true, studentLogin: true,
    validateSession: true, getAdminBootstrap: true, saveSettings: true, saveStudent: true, uploadStudentPassport: true, uploadBrandingAsset: true, setStudentState: true,
    importStudentsCsv: true, exportStudentsCsv: true, bulkUpdatePassports: true, importPassportsCsv: true,
    saveResult: true, saveResultEntries: true, importResultsCsv: true, recalculateResults: true, setResultState: true, bulkSetResultState: true,
    exportResultsCsv: true, loadStudentExamCodes: true, loadStudentResults: true, getImageDataUrl: true, generateStudentPdf: true, sendBulkEmail: true,
    authorizeMailStatus: true, exportSmsContacts: true, requestPasswordReset: true, resetPassword: true,
    setAdminState: true
  };
  if (!allowed[action]) throw new Error('Invalid or missing action.');
}

function shouldBypassFreshRequest_(action) {
  return action === 'ping' || action === 'getPublicBootstrap';
}

function validateFreshRequest_(params, action) {
  var payload = parseJson_(params.payload);
  var ts = Number(params.requestTs || (payload && payload.requestTs) || 0);
  var nonce = sanitizeValue_(params.requestNonce || (payload && payload.requestNonce));
  if (!ts || !nonce) throw new Error('Security check failed. Refresh and try again.');
  if (Math.abs(Date.now() - ts) > REQUEST_TTL_SECONDS * 1000) throw new Error('This request expired. Refresh and try again.');
  var cache = CacheService.getScriptCache();
  var nonceKey = 'REQ_NONCE_' + nonce;
  if (cache.get(nonceKey)) throw new Error('Duplicate request blocked.');
  cache.put(nonceKey, action, REQUEST_TTL_SECONDS);
}

function setupSystem() {
  return setupSystem_();
}

function setupSystem_() {
  ensureSheet_(SHEET_NAMES.SETTINGS, HEADERS.SETTINGS);
  ensureSheet_(SHEET_NAMES.ADMINS, HEADERS.ADMINS);
  ensureSheet_(SHEET_NAMES.STUDENTS, HEADERS.STUDENTS);
  ensureSheet_(SHEET_NAMES.RESULTS, HEADERS.RESULTS);
  ensureSheet_(SHEET_NAMES.REMARKS, HEADERS.REMARKS);
  ensureSheet_(SHEET_NAMES.RESET_CODES, HEADERS.RESET_CODES);
  ensureSheet_(SHEET_NAMES.AUDIT_LOGS, HEADERS.AUDIT_LOGS);
  bootstrapSettings_();
  bootstrapAdmins_();
  bootstrapRemarks_();
  return ok_('System setup completed.');
}

function ensureSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var lastCol = sh.getLastColumn();
  var existing = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (sh.getLastRow() === 0 || existing.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var missing = headers.filter(function(h) { return existing.indexOf(h) === -1; });
    if (missing.length) {
      sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#e8f0ff');
  return sh;
}

function bootstrapSettings_() {
  var defaults = {
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
    SHOW_POSITION_ON_REPORT: 'true',
    BRAND_LOGO_URL: '',
    FAVICON_URL: '',
    SIGNATURE_NAME: 'Authorized Signatory',
    SIGNATURE_URL: '',
    PORTAL_NOTICE: 'Leave exam code blank to load all published results, or enter an exam code to load that particular published result.',
    CLASS_OPTIONS: 'Year 1,Year 2,Year 3,Year 4,Year 5,Year 6,JSS 1,JSS 2,JSS 3,SS 1,SS 2,SS 3,Undergraduate',
    CATEGORY_OPTIONS: 'Lower Primary,Upper Primary,Junior High School,Senior High School,Undergraduate',
    SUBJECT_OPTIONS: 'English Language,Mathematics,Basic Science,Social Studies',
    STUDENT_SIGNUP_ENABLED: 'false',
    PUBLIC_ADMIN_SIGNUP_ENABLED: 'false',
    RESULT_FOOTER_NOTE: 'This result was generated from the official result portal.'
  };
  var settings = getSettings_();
  Object.keys(defaults).forEach(function(key) {
    if (!(key in settings)) upsertSetting_(key, defaults[key]);
  });
}

function bootstrapAdmins_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  if (countNonBlankRows_(sh) > 0) return;
  var creds = createPasswordRecord_('admin12345');
  appendObjectRow_(sh, {
    Username: 'admin',
    PasswordHash: creds.hash,
    PasswordSalt: creds.salt,
    PasswordVersion: creds.version,
    DisplayName: 'Principal Admin',
    Email: '',
    Phone: '',
    IsPrincipal: true,
    Active: true,
    Archived: false,
    Deleted: false,
    CreatedAt: isoNow_(),
    UpdatedAt: isoNow_(),
    LastLoginAt: ''
  });
}

function bootstrapRemarks_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.REMARKS);
  if (countNonBlankRows_(sh) > 0) return;
  var rows = [
    [0, 29, 'F', 'Very poor performance. Serious improvement is required.'],
    [30, 39, 'E', 'Poor performance. More effort and guidance are needed.'],
    [40, 49, 'D', 'Below average performance. Keep working harder.'],
    [50, 59, 'C', 'Average performance. There is room for growth.'],
    [60, 69, 'B', 'Good performance. Keep improving steadily.'],
    [70, 79, 'B+', 'Very good performance. Strong effort shown.'],
    [80, 89, 'A', 'Excellent performance. Keep it up.'],
    [90, 100, 'A+', 'Outstanding performance. Exceptional work.']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function getSettings_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  var rows = getSheetObjects_(sh);
  var out = {};
  rows.forEach(function(row) {
    if (row.Key) out[String(row.Key).trim()] = row.Value;
  });
  return out;
}

function upsertSetting_(key, value) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  var row = findRowByValue_(sh, 'Key', key);
  if (row) {
    updateObjectRow_(sh, row, { Key: key, Value: value });
  } else {
    appendObjectRow_(sh, { Key: key, Value: value });
  }
}

function getPublicBootstrap_() {
  var settings = getSettings_();
  return ok_('Public portal settings loaded.', {
    settings: settings,
    classOptions: parseListSetting_(settings.CLASS_OPTIONS),
    categoryOptions: parseListSetting_(settings.CATEGORY_OPTIONS),
    subjectOptions: parseListSetting_(settings.SUBJECT_OPTIONS),
    allowStudentSignup: normalizeBoolean_(settings.STUDENT_SIGNUP_ENABLED, false),
    allowAdminSignup: normalizeBoolean_(settings.PUBLIC_ADMIN_SIGNUP_ENABLED, false)
  });
}


function ensurePrincipalAdminConsistency_() {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  var rows = getSheetObjectsWithIndex_(sh).filter(function(item){ return !normalizeBoolean_(item.obj.Deleted,false); });
  if (!rows.length) return null;
  var principals = rows.filter(function(item){ return normalizeBoolean_(item.obj.IsPrincipal,false); });
  if (principals.length === 1) return principals[0];
  var candidate = null;
  if (principals.length > 1) {
    candidate = principals[0];
  } else {
    candidate = rows.filter(function(item){ return normalizeBoolean_(item.obj.Active,true) && !normalizeBoolean_(item.obj.Archived,false); })[0] || rows[0];
  }
  rows.forEach(function(item){
    var shouldPrincipal = item.rowIndex === candidate.rowIndex;
    if (normalizeBoolean_(item.obj.IsPrincipal,false) !== shouldPrincipal) {
      updateObjectRow_(sh, item.rowIndex, { IsPrincipal: shouldPrincipal, UpdatedAt: isoNow_() });
      item.obj.IsPrincipal = shouldPrincipal;
    }
  });
  return candidate;
}

function getAdminRowByUsername_(username) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  return getSheetObjectsWithIndex_(sh).find(function(item) {
    return String(item.obj.Username || '').toLowerCase() === String(username || '').toLowerCase();
  });
}

function adminLogin_(payload) {
  var username = sanitizeValue_(payload.username);
  var password = String(payload.password || '');
  var clientId = sanitizeValue_(payload.clientId);
  if (!username || !password) throw new Error('Username and password are required.');
  requireClientId_(clientId);
  enforceRateLimit_('admin-login:' + username.toLowerCase(), 8, 300, 'Too many admin login attempts. Try again later.');

  ensurePrincipalAdminConsistency_();
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  var row = getSheetObjectsWithIndex_(sh).find(function(item) {
    return String(item.obj.Username || '').toLowerCase() === username.toLowerCase();
  });
  if (!row || !normalizeBoolean_(row.obj.Active, true) || normalizeBoolean_(row.obj.Archived, false) || normalizeBoolean_(row.obj.Deleted, false)) {
    throw new Error('Invalid admin login details.');
  }
  if (!verifyPasswordAndUpgrade_(sh, row, password)) {
    throw new Error('Invalid admin login details.');
  }
  updateObjectRow_(sh, row.rowIndex, { LastLoginAt: isoNow_(), UpdatedAt: isoNow_() });
  var session = createSession_('admin', row.obj.Username, clientId, {
    displayName: row.obj.DisplayName || row.obj.Username,
    isPrincipal: normalizeBoolean_(row.obj.IsPrincipal, false)
  });
  logAudit_('admin', row.obj.Username, 'adminLogin', 'OK', 'Admin logged in.');
  return ok_('Login successful.', {
    token: session.token,
    displayName: row.obj.DisplayName || row.obj.Username,
    isPrincipal: normalizeBoolean_(row.obj.IsPrincipal, false)
  });
}

function adminSignup_(payload) {
  var username = sanitizeValue_(payload.username).toLowerCase();
  var displayName = sanitizeValue_(payload.displayName);
  var email = sanitizeEmail_(payload.email);
  var phone = sanitizeValue_(payload.phone);
  var password = String(payload.password || '');
  var settings = getSettings_();
  ensurePrincipalAdminConsistency_();
  var admins = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS)).filter(function(a) {
    return !normalizeBoolean_(a.Deleted, false);
  });
  var caller = null;
  if (payload.token) {
    caller = requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId));
  }
  if (admins.length > 0) {
    if (!caller) throw new Error('Only the principal admin can create sub-admin accounts.');
    caller = requirePrincipalAdmin_(payload.token, sanitizeValue_(payload.clientId));
  } else if (!normalizeBoolean_(settings.PUBLIC_ADMIN_SIGNUP_ENABLED, false) && !caller) {
    throw new Error('Public admin signup is disabled.');
  }
  validatePasswordStrength_(password, 'Password');
  if (!username || !displayName || !email) throw new Error('Display name, username, and email are required.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  if (findObjectRow_(sh, function(obj) { return String(obj.Username || '').toLowerCase() === username; })) {
    throw new Error('This username already exists.');
  }
  if (findObjectRow_(sh, function(obj) { return sanitizeEmail_(obj.Email) === email; })) {
    throw new Error('This email is already in use.');
  }
  var creds = createPasswordRecord_(password);
  appendObjectRow_(sh, {
    Username: username,
    PasswordHash: creds.hash,
    PasswordSalt: creds.salt,
    PasswordVersion: creds.version,
    DisplayName: displayName,
    Email: email,
    Phone: phone,
    IsPrincipal: admins.length === 0,
    Active: true,
    Archived: false,
    Deleted: false,
    CreatedAt: isoNow_(),
    UpdatedAt: isoNow_(),
    LastLoginAt: ''
  });
  logAudit_(caller ? 'admin' : 'public', caller ? caller.id : 'public', 'adminSignup', 'OK', 'Created admin ' + username);
  return ok_(admins.length === 0 ? 'Principal admin created successfully.' : 'Sub-admin created successfully.');
}

function studentSignup_(payload) {
  var settings = getSettings_();
  if (!normalizeBoolean_(settings.STUDENT_SIGNUP_ENABLED, false)) {
    throw new Error('Student signup is disabled. Contact the school admin.');
  }
  var regId = sanitizeRegId_(payload.regId || payload.RegID);
  var password = String(payload.password || payload.Password || '');
  if (!regId) throw new Error('Registration ID is required.');
  validatePasswordStrength_(password, 'Password');
  var required = ['fullName','parentEmail','parentPhone','schoolName','classLevel','category'];
  required.forEach(function(key) {
    if (!sanitizeValue_(payload[key] || payload[toPascalCase_(key)])) throw new Error(formLabel_(key) + ' is required.');
  });
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  if (findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === regId; })) {
    throw new Error('This Registration ID already exists.');
  }
  var creds = createPasswordRecord_(password);
  appendObjectRow_(sh, normalizeStudentPayload_({
    RegID: regId,
    PasswordHash: creds.hash,
    PasswordSalt: creds.salt,
    PasswordVersion: creds.version,
    FullName: payload.fullName || payload.FullName,
    Age: payload.age || payload.Age,
    Gender: payload.gender || payload.Gender,
    DOB: payload.dob || payload.DOB,
    ParentName: payload.parentName || payload.ParentName,
    ParentPhone: payload.parentPhone || payload.ParentPhone,
    ParentEmail: payload.parentEmail || payload.ParentEmail,
    SchoolName: payload.schoolName || payload.SchoolName,
    State: payload.state || payload.State,
    CityLGA: payload.cityLGA || payload.CityLGA,
    ClassLevel: payload.classLevel || payload.ClassLevel,
    Category: payload.category || payload.Category,
    Address: payload.address || payload.Address,
    PassportUrl: payload.passportUrl || payload.PassportUrl,
    Active: true,
    Archived: false,
    Deleted: false,
    CreatedAt: isoNow_(),
    UpdatedAt: isoNow_(),
    LastLoginAt: ''
  }));
  return ok_('Signup successful. You can now log in to check your result.');
}

function studentLogin_(payload) {
  var regId = sanitizeRegId_(payload.regId || payload.RegID);
  var password = String(payload.password || payload.Password || '');
  var clientId = sanitizeValue_(payload.clientId);
  if (!regId || !password) throw new Error('Registration ID and password are required.');
  requireClientId_(clientId);
  enforceRateLimit_('student-login:' + regId, 8, 300, 'Too many student login attempts. Try again later.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var row = getSheetObjectsWithIndex_(sh).find(function(item) {
    return sanitizeRegId_(item.obj.RegID) === regId;
  });
  if (!row || !normalizeBoolean_(row.obj.Active, true) || normalizeBoolean_(row.obj.Archived, false) || normalizeBoolean_(row.obj.Deleted, false)) {
    throw new Error('Invalid student login details.');
  }
  if (!verifyPasswordAndUpgrade_(sh, row, password)) {
    throw new Error('Invalid student login details.');
  }
  updateObjectRow_(sh, row.rowIndex, { LastLoginAt: isoNow_(), UpdatedAt: isoNow_() });
  var session = createSession_('student', regId, clientId, { fullName: row.obj.FullName || regId });
  return ok_('Login successful.', { token: session.token, regId: regId, fullName: row.obj.FullName || regId });
}

function validateSession_(payload) {
  var session = requireSession_(payload.token, null, sanitizeValue_(payload.clientId));
  return ok_('Session is valid.', {
    role: session.role,
    username: session.role === 'admin' ? session.id : '',
    regId: session.role === 'student' ? session.id : '',
    displayName: session.displayName || session.fullName || session.id,
    isPrincipal: !!session.isPrincipal
  });
}

function getAdminBootstrap_(payload) {
  ensurePrincipalAdminConsistency_();
  var session = requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId));
  var settings = getSettings_();
  var admins = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS)).map(cleanAdmin_);
  var students = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS)).map(cleanStudent_);
  students.sort(function(a, b) { return String(a.fullName || '').localeCompare(String(b.fullName || '')); });
  var results = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS)).map(cleanResult_);
  results.sort(function(a, b) { return String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')); });
  return ok_('Dashboard loaded.', {
    settings: settings,
    admins: admins,
    students: students,
    results: results,
    classOptions: parseListSetting_(settings.CLASS_OPTIONS),
    categoryOptions: parseListSetting_(settings.CATEGORY_OPTIONS),
    subjectOptions: parseListSetting_(settings.SUBJECT_OPTIONS),
    examCodes: uniqueList_(results.map(function(r) { return r.examCode; }).filter(Boolean)),
    currentAdmin: admins.find(function(a) { return a.username === session.id; }) || { username: session.id, displayName: session.displayName || session.id, isPrincipal: !!session.isPrincipal },
    summary: {
      studentCount: students.length,
      activeStudentCount: students.filter(function(s) { return s.active && !s.archived && !s.deleted; }).length,
      resultCount: results.length,
      publishedCount: results.filter(function(r) { return r.published; }).length,
      viewActiveCount: results.filter(function(r) { return r.viewActive; }).length
    }
  });
}

function saveSettings_(payload) {
  requirePrincipalAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var keys = ['BRAND_NAME','HEAD_OFFICE_ADDRESS','SCHOOL_NAME','SCHOOL_PHONE','SCHOOL_EMAIL','PRINCIPAL_NAME','ACADEMIC_SESSION','TERM','EXAM_TITLE','PASS_MARK_NUMBER','PASS_MARK_PERCENTAGE','SHOW_POSITION_ON_REPORT','SIGNATURE_NAME','SIGNATURE_URL','BRAND_LOGO_URL','FAVICON_URL','PORTAL_NOTICE','CLASS_OPTIONS','CATEGORY_OPTIONS','SUBJECT_OPTIONS','STUDENT_SIGNUP_ENABLED','PUBLIC_ADMIN_SIGNUP_ENABLED','RESULT_FOOTER_NOTE','RESULT_DETAIL_FIELDS'];
  keys.forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(payload, key)) upsertSetting_(key, sanitizeSettingValue_(key, payload[key]));
  });
  logAudit_('admin', requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId)).id, 'saveSettings', 'OK', 'Portal settings updated.');
  return ok_('Settings saved successfully.', { settings: getSettings_() });
}

function saveStudent_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var regId = sanitizeRegId_(payload.regId || payload.RegID);
  if (!regId) throw new Error('Registration ID is required.');
  var row = findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === regId; });
  var password = String(payload.password || payload.Password || '');
  if (!row && !password) throw new Error('Password is required for new student.');
  var passwordRecord = row ? {
    hash: row.obj.PasswordHash || '',
    salt: row.obj.PasswordSalt || '',
    version: row.obj.PasswordVersion || ''
  } : createPasswordRecord_(password);
  if (password) {
    validatePasswordStrength_(password, 'Password');
    passwordRecord = createPasswordRecord_(password);
  }
  var base = row ? row.obj : {};
  var record = normalizeStudentPayload_({
    RegID: regId,
    PasswordHash: passwordRecord.hash,
    PasswordSalt: passwordRecord.salt,
    PasswordVersion: passwordRecord.version,
    FullName: payload.fullName || payload.FullName || base.FullName,
    Age: payload.age || payload.Age || base.Age,
    Gender: payload.gender || payload.Gender || base.Gender,
    DOB: payload.dob || payload.DOB || base.DOB,
    ParentName: payload.parentName || payload.ParentName || base.ParentName,
    ParentPhone: payload.parentPhone || payload.ParentPhone || base.ParentPhone,
    ParentEmail: payload.parentEmail || payload.ParentEmail || base.ParentEmail,
    SchoolName: payload.schoolName || payload.SchoolName || base.SchoolName,
    State: payload.state || payload.State || base.State,
    CityLGA: payload.cityLGA || payload.CityLGA || base.CityLGA,
    ClassLevel: payload.classLevel || payload.ClassLevel || base.ClassLevel,
    Category: payload.category || payload.Category || base.Category,
    Address: payload.address || payload.Address || base.Address,
    PassportUrl: normalizeImageUrl_(payload.passportUrl || payload.PassportUrl || base.PassportUrl),
    Active: payload.active != null ? payload.active : (base.Active !== undefined ? base.Active : true),
    Archived: payload.archived != null ? payload.archived : (base.Archived !== undefined ? base.Archived : false),
    Deleted: payload.deleted != null ? payload.deleted : (base.Deleted !== undefined ? base.Deleted : false),
    CreatedAt: base.CreatedAt || isoNow_(),
    UpdatedAt: isoNow_(),
    LastLoginAt: base.LastLoginAt || ''
  });
  if (row) updateObjectRow_(sh, row.rowIndex, record); else appendObjectRow_(sh, record);
  return ok_(row ? 'Student updated successfully.' : 'Student created successfully.');
}

function setStudentState_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var regId = sanitizeRegId_(payload.regId);
  var state = sanitizeValue_(payload.state).toLowerCase();
  if (!regId || !state) throw new Error('Student and state are required.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var row = findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === regId; });
  if (!row) throw new Error('Student not found.');

  if (state === 'harddelete') {
    var resultsSh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
    if (resultsSh) {
      var resultRows = getSheetObjectsWithIndex_(resultsSh)
        .filter(function(item) { return sanitizeRegId_(item.obj.RegID) === regId; })
        .map(function(item) { return item.rowIndex; })
        .sort(function(a, b) { return b - a; });
      resultRows.forEach(function(rowIndex) {
        resultsSh.deleteRow(rowIndex);
      });
    }

    sh.deleteRow(row.rowIndex);
    return ok_('Student deleted forever.');
  }

  var patch = studentStatePatch_(state);
  patch.UpdatedAt = isoNow_();
  updateObjectRow_(sh, row.rowIndex, patch);
  return ok_('Student state updated.');
}

function importStudentsCsv_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) throw new Error('No student rows were provided.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var created = 0;
  var updated = 0;
  rows.forEach(function(src) {
    var regId = sanitizeRegId_(src.regId || src.RegID);
    if (!regId) return;
    var existing = findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === regId; });
    var password = String(src.password || src.Password || '');
    var current = existing ? existing.obj : {};
    var passwordRecord = existing ? {
      hash: current.PasswordHash || '',
      salt: current.PasswordSalt || '',
      version: current.PasswordVersion || ''
    } : createPasswordRecord_(password || regId + '123');
    if (password) passwordRecord = createPasswordRecord_(password);
    var record = normalizeStudentPayload_({
      RegID: regId,
      PasswordHash: passwordRecord.hash,
      PasswordSalt: passwordRecord.salt,
      PasswordVersion: passwordRecord.version,
      FullName: src.fullName || src.FullName || current.FullName,
      Age: src.age || src.Age || current.Age,
      Gender: src.gender || src.Gender || current.Gender,
      DOB: src.dob || src.DOB || current.DOB,
      ParentName: src.parentName || src.ParentName || current.ParentName,
      ParentPhone: src.parentPhone || src.ParentPhone || current.ParentPhone,
      ParentEmail: src.parentEmail || src.ParentEmail || current.ParentEmail,
      SchoolName: src.schoolName || src.SchoolName || current.SchoolName,
      State: src.state || src.State || current.State,
      CityLGA: src.cityLGA || src.CityLGA || current.CityLGA,
      ClassLevel: src.classLevel || src.ClassLevel || current.ClassLevel,
      Category: src.category || src.Category || current.Category,
      Address: src.address || src.Address || current.Address,
      PassportUrl: normalizeImageUrl_(src.passportUrl || src.PassportUrl || src.imageUrl || current.PassportUrl),
      Active: src.active != null && src.active !== '' ? src.active : (current.Active !== undefined ? current.Active : true),
      Archived: current.Archived || false,
      Deleted: current.Deleted || false,
      CreatedAt: current.CreatedAt || isoNow_(),
      UpdatedAt: isoNow_(),
      LastLoginAt: current.LastLoginAt || ''
    });
    if (existing) {
      updateObjectRow_(sh, existing.rowIndex, record);
      updated++;
    } else {
      appendObjectRow_(sh, record);
      created++;
    }
  });
  return ok_('Student CSV import completed.', { created: created, updated: updated });
}

function exportStudentsCsv_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS)).map(cleanStudent_);
  var headers = ['regId','fullName','age','gender','dob','parentName','parentPhone','parentEmail','schoolName','state','cityLGA','classLevel','category','address','passportUrl','active','archived','deleted','createdAt','updatedAt'];
  return ok_('Student export prepared.', {
    filename: 'students_export.csv',
    csv: objectsToCsv_(rows, headers)
  });
}

function bulkUpdatePassports_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var regIds = Array.isArray(payload.regIds) ? payload.regIds : [];
  var passportUrl = normalizeImageUrl_(payload.passportUrl);
  if (!regIds.length) throw new Error('Select at least one student first.');
  if (!passportUrl) throw new Error('Passport URL is required.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var updated = 0;
  regIds.forEach(function(id) {
    var row = findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === sanitizeRegId_(id); });
    if (!row) return;
    updateObjectRow_(sh, row.rowIndex, { PassportUrl: passportUrl, UpdatedAt: isoNow_() });
    updated++;
  });
  return ok_('Passport URLs updated.', { updated: updated });
}

function importPassportsCsv_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) throw new Error('No passport rows were provided.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
  var updated = 0;
  rows.forEach(function(src) {
    var regId = sanitizeRegId_(src.regId || src.RegID);
    var passportUrl = normalizeImageUrl_(src.passportUrl || src.PassportUrl || src.imageUrl);
    if (!regId || !passportUrl) return;
    var row = findObjectRow_(sh, function(obj) { return sanitizeRegId_(obj.RegID) === regId; });
    if (!row) return;
    updateObjectRow_(sh, row.rowIndex, { PassportUrl: passportUrl, UpdatedAt: isoNow_() });
    updated++;
  });
  return ok_('Passport CSV import completed.', { updated: updated });
}

function saveResultEntries_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var entries = Array.isArray(payload.entries) ? payload.entries : [];
  if (!entries.length) throw new Error('No result entries were provided.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  var settings = getSettings_();
  var saved = 0;
  entries.forEach(function(entry) {
    var normalized = normalizeResultEntry_(entry, settings);
    if (!normalized.RegID || !normalized.ExamCode || !normalized.Subject) return;
    if (!getStudentByRegId_(normalized.RegID)) throw new Error('Student not found: ' + normalized.RegID);
    var existing = findObjectRow_(sh, function(obj) {
      return sanitizeRegId_(obj.RegID) === normalized.RegID &&
        sanitizeValue_(obj.ExamCode).toLowerCase() === normalized.ExamCode.toLowerCase() &&
        sanitizeValue_(obj.Subject).toLowerCase() === normalized.Subject.toLowerCase() &&
        sanitizeValue_(obj.AcademicSession).toLowerCase() === normalized.AcademicSession.toLowerCase() &&
        sanitizeValue_(obj.Term).toLowerCase() === normalized.Term.toLowerCase();
    });
    if (existing) {
      normalized.ResultID = existing.obj.ResultID || normalized.ResultID;
      normalized.CreatedAt = existing.obj.CreatedAt || normalized.CreatedAt;
      normalized.Archived = existing.obj.Archived || false;
      normalized.Deleted = existing.obj.Deleted || false;
      if (normalized.Published && !existing.obj.PublishedAt) normalized.PublishedAt = isoNow_();
      if (existing.obj.PublishedAt && normalized.Published) normalized.PublishedAt = existing.obj.PublishedAt;
      updateObjectRow_(sh, existing.rowIndex, normalized);
    } else {
      appendObjectRow_(sh, normalized);
    }
    saved++;
  });
  recalculateRankings_();
  return ok_('Result entries saved successfully.', { saved: saved });
}

function importResultsCsv_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) throw new Error('No result rows were provided.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  var settings = getSettings_();
  var created = 0;
  var updated = 0;
  rows.forEach(function(src) {
    var normalized = normalizeResultEntry_(src, settings);
    if (!normalized.RegID || !normalized.ExamCode || !normalized.Subject) return;
    var existing = findObjectRow_(sh, function(obj) {
      return sanitizeRegId_(obj.RegID) === normalized.RegID &&
        sanitizeValue_(obj.ExamCode).toLowerCase() === normalized.ExamCode.toLowerCase() &&
        sanitizeValue_(obj.Subject).toLowerCase() === normalized.Subject.toLowerCase() &&
        sanitizeValue_(obj.AcademicSession).toLowerCase() === normalized.AcademicSession.toLowerCase() &&
        sanitizeValue_(obj.Term).toLowerCase() === normalized.Term.toLowerCase();
    });
    if (existing) {
      normalized.ResultID = existing.obj.ResultID || normalized.ResultID;
      normalized.CreatedAt = existing.obj.CreatedAt || normalized.CreatedAt;
      normalized.Archived = existing.obj.Archived || false;
      normalized.Deleted = existing.obj.Deleted || false;
      if (normalized.Published && !existing.obj.PublishedAt) normalized.PublishedAt = isoNow_();
      if (existing.obj.PublishedAt && normalized.Published) normalized.PublishedAt = existing.obj.PublishedAt;
      updateObjectRow_(sh, existing.rowIndex, normalized);
      updated++;
    } else {
      appendObjectRow_(sh, normalized);
      created++;
    }
  });
  recalculateRankings_();
  return ok_('Result CSV import completed.', { created: created, updated: updated });
}

function recalculateResults_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var affected = recalculateRankings_(payload.examCode, payload.academicSession, payload.term);
  return ok_('Ranking recalculation completed.', { affected: affected });
}

function setResultState_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var resultId = sanitizeValue_(payload.resultId);
  var state = sanitizeValue_(payload.state).toLowerCase();
  if (!resultId || !state) throw new Error('Result and state are required.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  var row = findObjectRow_(sh, function(obj) { return sanitizeValue_(obj.ResultID) === resultId; });
  if (!row) throw new Error('Result not found.');
  if (state === 'harddelete') {
    sh.deleteRow(row.rowIndex);
    return ok_('Result deleted forever.');
  }
  var patch = resultStatePatch_(state, row.obj);
  patch.UpdatedAt = isoNow_();
  updateObjectRow_(sh, row.rowIndex, patch);
  return ok_('Result state updated.');
}

function exportResultsCsv_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS)).map(cleanResult_);
  var headers = ['resultId','regId','examCode','examTitle','subject','examDate','studentScore','maxScore','percentage','positionText','grade','remark','teacherComment','academicSession','term','published','viewActive','publishedAt','signatureUrl','archived','deleted'];
  return ok_('Result export prepared.', {
    filename: 'results_export.csv',
    csv: objectsToCsv_(rows, headers)
  });
}

function loadStudentExamCodes_(payload) {
  var session = requireSession_(payload.token, 'student', sanitizeValue_(payload.clientId));
  var regId = session.id;
  var rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS)).map(cleanResult_).filter(function(r) {
    return r.regId === regId && r.published && r.viewActive && !r.archived && !r.deleted;
  });
  var seen = {};
  var list = [];
  rows.forEach(function(r) {
    var key = [r.examCode, r.academicSession, r.term].join('|');
    if (seen[key]) return;
    seen[key] = true;
    list.push({ examCode: r.examCode, examTitle: r.examTitle, academicSession: r.academicSession, term: r.term, publishedAt: r.publishedAt });
  });
  list.sort(function(a, b) { return String(b.publishedAt || '').localeCompare(String(a.publishedAt || '')); });
  return ok_('Available exam codes loaded.', list);
}

function parseExamCodeList_(value) {
  return uniqueList_(String(value || '').split(/[;,\n]+/).map(function(code) {
    return sanitizeValue_(code).toLowerCase();
  }).filter(Boolean));
}

function matchesAnyExamCodeFilter_(resultRow, examCodes) {
  if (!examCodes || !examCodes.length) return true;
  var rowExamCode = sanitizeValue_(resultRow.examCode).toLowerCase();
  return examCodes.some(function(code) {
    return rowExamCode === sanitizeValue_(code).toLowerCase();
  });
}

function getStudentResultBundle_(session, payload) {
  var examCodeInput = sanitizeValue_(payload.examCode);
  var examCodes = parseExamCodeList_(examCodeInput);
  var academicSession = sanitizeValue_(payload.academicSession);
  var term = sanitizeValue_(payload.term);
  var student = cleanStudent_(getStudentByRegId_(session.id));
  if (!student) throw new Error('Student account not found.');
  var results = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS))
    .map(cleanResult_)
    .map(function(r) { r.classLevel = student.classLevel || ''; return r; })
    .filter(function(r) {
      return r.regId === session.id &&
        matchesAnyExamCodeFilter_(r, examCodes) &&
        r.published && r.viewActive && !r.archived && !r.deleted &&
        (!academicSession || r.academicSession === academicSession) &&
        (!term || r.term === term);
    })
    .sort(function(a, b) {
      var aDate = String(a.publishedAt || a.updatedAt || a.createdAt || '');
      var bDate = String(b.publishedAt || b.updatedAt || b.createdAt || '');
      var byDate = bDate.localeCompare(aDate);
      if (byDate !== 0) return byDate;
      var bySession = String(b.academicSession || '').localeCompare(String(a.academicSession || ''));
      if (bySession !== 0) return bySession;
      var byTerm = String(b.term || '').localeCompare(String(a.term || ''));
      if (byTerm !== 0) return byTerm;
      var byExam = String(a.examCode || '').localeCompare(String(b.examCode || ''));
      if (byExam !== 0) return byExam;
      return String(a.subject || '').localeCompare(String(b.subject || ''));
    });
  if (!results.length) {
    throw new Error(examCodes.length ? 'No published result was found for the supplied exam code.' : 'No published result is currently available for your account.');
  }
  return {
    examCode: examCodeInput || '',
    examCodeList: examCodes,
    loadMode: examCodes.length > 1 ? 'multi' : (examCodes.length ? 'single' : 'all'),
    student: student,
    settings: getSettings_(),
    results: results,
    sessions: uniqueList_(results.map(function(r) { return r.academicSession; }).filter(Boolean)),
    terms: uniqueList_(results.map(function(r) { return r.term; }).filter(Boolean))
  };
}

function loadStudentResults_(payload) {
  var session = requireSession_(payload.token, 'student', sanitizeValue_(payload.clientId));
  var data = getStudentResultBundle_(session, payload);
  return ok_(data.examCodeList.length ? 'Published result(s) for the supplied exam code loaded successfully.' : 'All published subjects loaded successfully.', data);
}

function normalizeSubjectKey_(value) {
  return sanitizeValue_(value).toLowerCase();
}

function applySelectedSubjectFilter_(rows, selectedSubjects) {
  if (!Array.isArray(selectedSubjects) || !selectedSubjects.length) return rows;
  var allowed = {};
  selectedSubjects.forEach(function(value) {
    var key = normalizeSubjectKey_(value);
    if (key) allowed[key] = true;
  });
  var filtered = (rows || []).filter(function(row) {
    return !!allowed[normalizeSubjectKey_(row && row.subject)];
  });
  return filtered.length ? filtered : rows;
}

function loadImageBlobFromSource_(value) {
  var raw = sanitizeValue_(value);
  if (!raw) throw new Error('Provide an image URL or file ID.');
  var driveId = extractDriveFileId_(raw);
  var blob = null;
  var source = '';
  var mimeType = '';
  var safeMimePattern = /^(image\/(png|jpe?g|webp|gif|bmp|svg\+xml))$/i;

  if (driveId) {
    var driveCandidates = buildPublicImageCandidates_(driveId);
    for (var d = 0; d < driveCandidates.length; d += 1) {
      var driveCandidate = driveCandidates[d];
      if (!driveCandidate) continue;
      try {
        var driveFetched = UrlFetchApp.fetch(driveCandidate, {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8'
          }
        });
        var driveCode = Number(driveFetched.getResponseCode() || 0);
        if (driveCode >= 200 && driveCode < 300) {
          var driveFetchedBlob = driveFetched.getBlob();
          var driveFetchedType = sanitizeValue_(driveFetchedBlob.getContentType()) || sanitizeValue_(driveFetched.getHeaders()['Content-Type']);
          if (safeMimePattern.test(driveFetchedType)) {
            blob = driveFetchedBlob;
            mimeType = driveFetchedType;
            source = 'drive_public';
            break;
          }
        }
      } catch (driveFetchErr) {}
    }
  }

  if (!blob) {
    try {
      if (driveId) {
        var file = DriveApp.getFileById(driveId);
        blob = file.getBlob();
        source = 'drive';
        if ((!blob.getContentType() || !/^image\//i.test(blob.getContentType())) && sanitizeValue_(file.getName())) {
          blob = blob.setContentTypeFromExtension();
        }
        mimeType = sanitizeValue_(blob.getContentType());
      }
    } catch (driveErr) {}
  }

  if (!blob) {
    var candidates = buildPublicImageCandidates_(raw);
    for (var i = 0; i < candidates.length; i += 1) {
      var candidate = candidates[i];
      if (!candidate) continue;
      try {
        var fetched = UrlFetchApp.fetch(candidate, {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8'
          }
        });
        var code = Number(fetched.getResponseCode() || 0);
        if (code >= 200 && code < 300) {
          var fetchedBlob = fetched.getBlob();
          var contentType = sanitizeValue_(fetchedBlob.getContentType()) || sanitizeValue_(fetched.getHeaders()['Content-Type']);
          if (/^image\//i.test(contentType)) {
            blob = fetchedBlob;
            mimeType = contentType;
            source = 'url';
            break;
          }
        }
      } catch (fetchErr) {}
    }
  }

  if (!blob) throw new Error('Unable to load that image for PDF export.');

  mimeType = sanitizeValue_(mimeType || blob.getContentType());
  if (!/^image\//i.test(mimeType)) {
    mimeType = guessMimeTypeFromName_(raw);
    try { blob = blob.setContentType(mimeType); } catch (setErr) {}
  }

  if (!safeMimePattern.test(mimeType) && driveId) {
    var safeCandidates = buildPublicImageCandidates_(driveId);
    for (var j = 0; j < safeCandidates.length; j += 1) {
      var safeCandidate = safeCandidates[j];
      if (!safeCandidate) continue;
      try {
        var safeFetched = UrlFetchApp.fetch(safeCandidate, {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8'
          }
        });
        var safeCode = Number(safeFetched.getResponseCode() || 0);
        if (safeCode >= 200 && safeCode < 300) {
          var safeBlob = safeFetched.getBlob();
          var safeType = sanitizeValue_(safeBlob.getContentType()) || sanitizeValue_(safeFetched.getHeaders()['Content-Type']);
          if (safeMimePattern.test(safeType)) {
            blob = safeBlob;
            mimeType = safeType;
            source = 'drive_public_safe';
            break;
          }
        }
      } catch (safeErr) {}
    }
  }

  if (!/^image\//i.test(mimeType)) throw new Error('The selected file is not an image.');
  return { blob: blob, mimeType: mimeType, source: source || (driveId ? 'drive' : 'url') };
}

function safeLoadImageBlob_(value) {
  try {
    return loadImageBlobFromSource_(value);
  } catch (err) {
    return null;
  }
}

function formatDisplayDate_(value) {
  var raw = sanitizeValue_(value);
  if (!raw) return '—';
  var dt = new Date(raw);
  if (isNaN(dt.getTime())) return raw;
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'MMM d, yyyy');
}

function formatNumberLikeStudent_(value) {
  if (value === '' || value == null) return '—';
  var num = Number(value);
  if (isNaN(num)) return sanitizeValue_(value);
  if (Math.abs(num - Math.round(num)) < 0.0000001) return String(Math.round(num));
  return String(Math.round(num * 100) / 100);
}

function calculateAveragePercent_(rows) {
  var list = (rows || []).map(function(row) { return Number(row && row.percentage); }).filter(function(n) { return !isNaN(n); });
  if (!list.length) return 0;
  var total = list.reduce(function(sum, value) { return sum + value; }, 0);
  return Math.round((total / list.length) * 100) / 100;
}

function makeResultPdfFileName_(student, rows, examCodeInput) {
  var first = (rows && rows[0]) || {};
  var code = sanitizeValue_(examCodeInput || first.examCode || 'all_published_results');
  var totalSubjects = (rows || []).length;
  return (safeName_(student && student.fullName || 'student') + '_' + safeName_(code) + '_' + totalSubjects + '_published_subjects_report_card.pdf').replace(/_+/g, '_');
}

function appendStyledParagraph_(body, text, options) {
  var p = body.appendParagraph(sanitizeValue_(text));
  options = options || {};
  if (options.heading) p.setHeading(options.heading);
  if (options.alignment) p.setAlignment(options.alignment);
  if (options.bold) p.setBold(true);
  if (options.italic) p.setItalic(true);
  if (options.spacingBefore != null) p.setSpacingBefore(options.spacingBefore);
  if (options.spacingAfter != null) p.setSpacingAfter(options.spacingAfter);
  if (options.fontSize) p.setFontSize(options.fontSize);
  if (options.foregroundColor) p.setForegroundColor(options.foregroundColor);
  if (options.backgroundColor) p.setBackgroundColor(options.backgroundColor);
  return p;
}

function setTableHeaderStyle_(row, palette) {
  if (!row) return;
  palette = palette || getPdfThemePalette_('royal');
  for (var i = 0; i < row.getNumCells(); i += 1) {
    var cell = row.getCell(i);
    try { cell.setBackgroundColor(palette.header); } catch (err) {}
    try {
      var text = cell.editAsText();
      text.setBold(true);
      text.setForegroundColor(palette.headerText);
    } catch (err2) {}
  }
}

function styleTableBodyRows_(table, palette) {
  if (!table) return;
  palette = palette || getPdfThemePalette_('royal');
  for (var r = 1; r < table.getNumRows(); r += 1) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c += 1) {
      var cell = row.getCell(c);
      try { cell.setBackgroundColor(r % 2 ? '#ffffff' : palette.light); } catch (err) {}
      try { cell.editAsText().setForegroundColor(palette.text); } catch (err2) {}
    }
  }
}

function createTableWithRows_(body, rows, palette) {
  var table = body.appendTable(rows);
  if (rows && rows.length) setTableHeaderStyle_(table.getRow(0), palette);
  styleTableBodyRows_(table, palette);
  return table;
}

function getPdfThemePalette_(designName) {
  var name = sanitizeValue_(designName || 'royal').toLowerCase();
  var base = name.split('-')[0] || 'royal';
  var isAlt = /-2$/.test(name);
  var map = {
    royal:   { header:'#2648c4', accent:'#12a2ff', light:'#f1f6ff', line:'#ccdafa', text:'#172640', sub:'#56708c', headerText:'#ffffff', lightAlt:'#fbfdff', lineAlt:'#d9e5ff' },
    emerald: { header:'#10815b', accent:'#1fc98c', light:'#effcf6', line:'#c7ead8', text:'#173428', sub:'#4f6d60', headerText:'#ffffff', lightAlt:'#fbfffd', lineAlt:'#d7efe2' },
    sunset:  { header:'#cf611a', accent:'#ff944d', light:'#fff5ec', line:'#f5d6bc', text:'#442519', sub:'#7f5a4c', headerText:'#ffffff', lightAlt:'#fffaf5', lineAlt:'#f6dfc9' },
    violet:  { header:'#6549be', accent:'#bc75ff', light:'#f7f2ff', line:'#dcd2f4', text:'#251c45', sub:'#6f6490', headerText:'#ffffff', lightAlt:'#fcf9ff', lineAlt:'#e6dcf8' },
    carbon:  { header:'#26344d', accent:'#6884ab', light:'#f4f7fc', line:'#ced6e4', text:'#1b2331', sub:'#5f6d81', headerText:'#ffffff', lightAlt:'#fafcff', lineAlt:'#d8e0ee' },
    rose:    { header:'#c03e61', accent:'#ff7ea8', light:'#fff2f7', line:'#f5cddd', text:'#441d2a', sub:'#866271', headerText:'#ffffff', lightAlt:'#fff8fb', lineAlt:'#f6dbe8' },
    ocean:   { header:'#0a74b8', accent:'#45c6e9', light:'#eef9ff', line:'#c7e6f2', text:'#153042', sub:'#567286', headerText:'#ffffff', lightAlt:'#f9fdff', lineAlt:'#d5eef6' },
    gold:    { header:'#a37314', accent:'#f2b12b', light:'#fff8ea', line:'#f1dfb2', text:'#4a3815', sub:'#7a6841', headerText:'#ffffff', lightAlt:'#fffdf7', lineAlt:'#f3e7c4' },
    forest:  { header:'#2a723a', accent:'#5aba6e', light:'#f1fcf3', line:'#cce5d1', text:'#18311f', sub:'#57705c', headerText:'#ffffff', lightAlt:'#fbfffc', lineAlt:'#d9ecdd' },
    candy:   { header:'#bf3f84', accent:'#568bff', light:'#fff3fa', line:'#ecd0e0', text:'#421c32', sub:'#7f6580', headerText:'#ffffff', lightAlt:'#fff9fc', lineAlt:'#f1dce8' }
  };
  var palette = map[base] || map.royal;
  if (!isAlt) return palette;
  return {
    header: palette.accent,
    accent: palette.header,
    light: palette.lightAlt || palette.light,
    line: palette.lineAlt || palette.line,
    text: palette.text,
    sub: palette.sub,
    headerText: '#ffffff'
  };
}

function setCellTextStyle_(cell, options) {
  if (!cell) return;
  options = options || {};
  try {
    var text = cell.editAsText();
    if (options.foregroundColor) text.setForegroundColor(options.foregroundColor);
    if (options.bold) text.setBold(true);
    if (options.fontSize) text.setFontSize(options.fontSize);
  } catch (err) {}
}

function appendParagraphInCell_(cell, text, options) {
  var p = cell.appendParagraph(sanitizeValue_(text));
  options = options || {};
  if (options.alignment) p.setAlignment(options.alignment);
  if (options.bold) p.setBold(true);
  if (options.fontSize) p.setFontSize(options.fontSize);
  if (options.foregroundColor) p.setForegroundColor(options.foregroundColor);
  if (options.backgroundColor) p.setBackgroundColor(options.backgroundColor);
  if (options.spacingAfter != null) p.setSpacingAfter(options.spacingAfter);
  if (options.spacingBefore != null) p.setSpacingBefore(options.spacingBefore);
  return p;
}


function appendCenteredImageInCell_(cell, blob, width, height, topSpace, bottomSpace) {
  if (!cell || !blob) return null;
  var p = cell.appendParagraph('');
  try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (err) {}
  if (topSpace != null) {
    try { p.setSpacingBefore(topSpace); } catch (err2) {}
  }
  if (bottomSpace != null) {
    try { p.setSpacingAfter(bottomSpace); } catch (err3) {}
  }
  var img = null;
  try {
    img = p.appendInlineImage(blob);
    if (width) img.setWidth(width);
    if (height) img.setHeight(height);
  } catch (err4) {}
  return img;
}

function resolveCurrentSignatureUrl_(settings, firstRow) {
  var current = normalizeImageUrl_(settings && settings.SIGNATURE_URL ? settings.SIGNATURE_URL : '');
  if (current) return current;
  return normalizeImageUrl_(firstRow && firstRow.signatureUrl ? firstRow.signatureUrl : '');
}

function createSummaryCardsTable_(body, cards, palette) {
  var table = body.appendTable([cards.map(function(card) { return sanitizeValue_(card.label || ''); })]);
  var row = table.getRow(0);
  for (var i = 0; i < cards.length; i += 1) {
    var cell = row.getCell(i);
    try { cell.clear(); } catch (err) {}
    try { cell.setBackgroundColor(palette.light); } catch (err2) {}
    appendParagraphInCell_(cell, sanitizeValue_(cards[i].label || ''), { bold: true, fontSize: 9, foregroundColor: palette.sub, spacingAfter: 4 });
    appendParagraphInCell_(cell, sanitizeValue_(cards[i].value || ''), { bold: true, fontSize: 12, foregroundColor: palette.text });
  }
  return table;
}

function createStudentPdfBlob_(bundle) {
  var rows = applySelectedSubjectFilter_(bundle.results || [], bundle.selectedSubjects || []);
  if (!rows.length) throw new Error('No published result is available for PDF generation.');
  var settings = bundle.settings || {};
  var student = bundle.student || {};
  var first = rows[0] || {};
  var average = calculateAveragePercent_(rows);
  var passed = rows.filter(function(row) { return sanitizeValue_(row.resultStatus).toUpperCase() === 'PASS'; }).length;
  var total = rows.length;
  var status = passed === total ? 'PASS' : (passed > 0 ? 'PARTIAL PASS' : 'NEEDS IMPROVEMENT');
  var palette = getPdfThemePalette_(bundle.resultDesign || settings.RESULT_DESIGN || 'royal');
  var signatureUrl = resolveCurrentSignatureUrl_(settings, first);
  var logoAsset = safeLoadImageBlob_(settings.BRAND_LOGO_URL || '');
  var passportAsset = safeLoadImageBlob_(student.passportUrl || '');
  var signatureAsset = safeLoadImageBlob_(signatureUrl);
  var fileName = makeResultPdfFileName_(student, rows, bundle.examCode);
  var statusColor = status === 'PASS' ? '#1b7f39' : (status === 'PARTIAL PASS' ? '#9b6a00' : '#b23a48');

  var doc = DocumentApp.create(fileName.replace(/\.pdf$/i, ''));
  var file = DriveApp.getFileById(doc.getId());
  try {
    var body = doc.getBody();
    body.clear();
    try { body.setMarginTop(28).setMarginBottom(28).setMarginLeft(28).setMarginRight(28); } catch (marginErr) {}

    var header = body.appendTable([['', '', '']]);
    var headerRow = header.getRow(0);
    var logoCell = headerRow.getCell(0);
    var textCell = headerRow.getCell(1);
    var passportCell = headerRow.getCell(2);
    try { logoCell.setBackgroundColor(palette.header); textCell.setBackgroundColor(palette.header); passportCell.setBackgroundColor(palette.header); } catch (err) {}
    try { logoCell.clear(); textCell.clear(); passportCell.clear(); } catch (clearErr) {}

    appendParagraphInCell_(logoCell, '', { spacingAfter: 2 });
    if (logoAsset && logoAsset.blob) {
      appendCenteredImageInCell_(logoCell, logoAsset.blob, 68, 68, 4, 2);
    } else {
      appendParagraphInCell_(logoCell, 'LOGO', { bold: true, fontSize: 12, foregroundColor: palette.headerText, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 4 });
    }
    appendParagraphInCell_(logoCell, settings.BRAND_NAME || 'School Brand', { bold: true, fontSize: 8, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER });

    appendParagraphInCell_(textCell, settings.BRAND_NAME || 'Genz Edutech Innovations', { bold: true, fontSize: 16, foregroundColor: palette.headerText, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 2, spacingBefore: 3 });
    appendParagraphInCell_(textCell, settings.SCHOOL_NAME || student.schoolName || 'Genz Result Portal', { bold: true, fontSize: 19, foregroundColor: palette.headerText, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 3 });
    if (sanitizeValue_(settings.HEAD_OFFICE_ADDRESS)) appendParagraphInCell_(textCell, settings.HEAD_OFFICE_ADDRESS, { fontSize: 9, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 1 });
    var headerContact = [settings.SCHOOL_PHONE, settings.SCHOOL_EMAIL].filter(function(v) { return sanitizeValue_(v); }).join(' • ');
    if (headerContact) appendParagraphInCell_(textCell, headerContact, { fontSize: 9, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 2 });
    appendParagraphInCell_(textCell, 'Official Student Result Portal', { fontSize: 8, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER });

    appendParagraphInCell_(passportCell, status, { bold: true, fontSize: 11, foregroundColor: palette.header, alignment: DocumentApp.HorizontalAlignment.CENTER, backgroundColor: '#ffffff', spacingAfter: 5, spacingBefore: 4 });
    if (passportAsset && passportAsset.blob) {
      appendCenteredImageInCell_(passportCell, passportAsset.blob, 78, 96, 2, 2);
    } else {
      appendParagraphInCell_(passportCell, 'No Passport', { fontSize: 9, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER, spacingBefore: 10, spacingAfter: 5 });
    }
    appendParagraphInCell_(passportCell, 'Student Passport', { fontSize: 8, foregroundColor: '#eef5ff', alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 2 });

    appendStyledParagraph_(body, 'RESULT', { alignment: DocumentApp.HorizontalAlignment.CENTER, bold: true, fontSize: 19, spacingBefore: 8, spacingAfter: 1, foregroundColor: palette.header });
    appendStyledParagraph_(body, status, { alignment: DocumentApp.HorizontalAlignment.CENTER, bold: true, fontSize: 11, spacingAfter: 8, foregroundColor: statusColor });

    var summaryTable = body.appendTable([['', '']]);
    var summaryRow = summaryTable.getRow(0);
    var summaryLeft = summaryRow.getCell(0);
    var summaryRight = summaryRow.getCell(1);
    try { summaryLeft.setBackgroundColor('#ffffff'); summaryRight.setBackgroundColor(palette.light); } catch (sumBgErr) {}
    try { summaryLeft.clear(); summaryRight.clear(); } catch (sumClearErr) {}
    appendParagraphInCell_(summaryLeft, 'Student Performance Summary', { bold: true, fontSize: 13, foregroundColor: palette.header, spacingAfter: 4 });
    appendParagraphInCell_(summaryLeft, [sanitizeValue_(student.fullName || 'Student'), sanitizeValue_(student.regId || '—'), sanitizeValue_(student.classLevel || '—')].filter(Boolean).join(' • '), { fontSize: 10, foregroundColor: palette.text, spacingAfter: 2 });
    appendParagraphInCell_(summaryLeft, 'Exam: ' + (sanitizeValue_(first.examTitle || first.examCode || '—')) + ' • Session: ' + ([sanitizeValue_(first.academicSession), sanitizeValue_(first.term)].filter(Boolean).join(' / ') || '—'), { fontSize: 10, foregroundColor: palette.text, spacingAfter: 2 });
    appendParagraphInCell_(summaryLeft, 'Published: ' + formatDisplayDate_(first.publishedAt) + ' • Average: ' + formatNumberLikeStudent_(average) + '% • Passed: ' + passed + '/' + total, { fontSize: 10, foregroundColor: palette.text, spacingAfter: 4 });
    appendParagraphInCell_(summaryLeft, 'Selection: ' + total + '/' + total + ' subject(s) included in this bulk report card', { fontSize: 9, foregroundColor: palette.sub });
    appendParagraphInCell_(summaryRight, 'Verified Academic Record', { bold: true, fontSize: 10, foregroundColor: palette.header, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 4 });
    appendParagraphInCell_(summaryRight, 'This PDF includes the official brand logo, student passport, and admin signature.', { fontSize: 9, foregroundColor: palette.sub, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 4 });
    appendParagraphInCell_(summaryRight, 'Student Reg ID', { bold: true, fontSize: 9, foregroundColor: palette.sub, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 2 });
    appendParagraphInCell_(summaryRight, sanitizeValue_(student.regId || '—'), { bold: true, fontSize: 13, foregroundColor: palette.text, alignment: DocumentApp.HorizontalAlignment.CENTER });

    createSummaryCardsTable_(body, [
      { label: 'Student Reg ID', value: sanitizeValue_(student.regId || '—') },
      { label: 'Subjects', value: String(total) },
      { label: 'Average %', value: formatNumberLikeStudent_(average) + '%' },
      { label: 'Published', value: formatDisplayDate_(first.publishedAt) }
    ], palette);

    appendStyledParagraph_(body, 'Published Subject Results', { bold: true, fontSize: 13, spacingBefore: 10, spacingAfter: 5, foregroundColor: palette.header });
    var resultTableRows = [['Exam Code', 'Session', 'Term', 'Subject', 'Score', 'Percent', 'Grade', 'Rank', 'Status']];
    rows.forEach(function(row) {
      resultTableRows.push([
        sanitizeValue_(row.examCode || '—'),
        sanitizeValue_(row.academicSession || '—'),
        sanitizeValue_(row.term || '—'),
        sanitizeValue_(row.subject || '—'),
        formatNumberLikeStudent_(row.studentScore) + ' / ' + formatNumberLikeStudent_(row.maxScore),
        formatNumberLikeStudent_(row.percentage) + '%',
        sanitizeValue_(row.grade || '—'),
        sanitizeValue_(row.positionText || '—'),
        sanitizeValue_(row.resultStatus || '—')
      ]);
    });
    createTableWithRows_(body, resultTableRows, palette);

    appendStyledParagraph_(body, 'Remarks and Teacher Comments', { bold: true, fontSize: 13, spacingBefore: 10, spacingAfter: 5, foregroundColor: palette.accent });
    var remarkRows = [['Exam / Subject', 'Remark', 'Teacher Comment']];
    rows.forEach(function(row) {
      remarkRows.push([
        sanitizeValue_((row.examCode || '—') + ' • ' + (row.subject || '—')),
        sanitizeValue_(row.remark || '—'),
        sanitizeValue_(row.teacherComment || 'No teacher comment')
      ]);
    });
    createTableWithRows_(body, remarkRows, { header: palette.accent, light: palette.light, text: palette.text, headerText: '#ffffff' });

    var footerTable = body.appendTable([['', '']]);
    var footerRow = footerTable.getRow(0);
    var noteCell = footerRow.getCell(0);
    var signCell = footerRow.getCell(1);
    try { noteCell.clear(); signCell.clear(); } catch (footErr) {}
    try { noteCell.setBackgroundColor(palette.light); signCell.setBackgroundColor('#ffffff'); } catch (footBgErr) {}
    appendParagraphInCell_(noteCell, 'Official Verification Note', { bold: true, fontSize: 12, foregroundColor: palette.header, spacingAfter: 4 });
    appendParagraphInCell_(noteCell, 'This result was generated from the official student result portal on ' + formatDisplayDate_(new Date()) + '. ' + total + ' selected subject(s) were included in this bulk report card. The brand logo, student passport, published result rows, and admin signature form part of this verified academic record.', { fontSize: 9, foregroundColor: palette.sub });
    appendParagraphInCell_(signCell, 'Authorized Signatory', { bold: true, fontSize: 11, foregroundColor: palette.header, alignment: DocumentApp.HorizontalAlignment.CENTER, spacingAfter: 2, spacingBefore: 2 });
    appendParagraphInCell_(signCell, '', { spacingAfter: 1 });
    if (signatureAsset && signatureAsset.blob) {
      appendCenteredImageInCell_(signCell, signatureAsset.blob, 150, 42, 2, 2);
    } else {
      appendParagraphInCell_(signCell, 'No signature added', { alignment: DocumentApp.HorizontalAlignment.CENTER, fontSize: 10, foregroundColor: palette.sub, spacingAfter: 2, spacingBefore: 8 });
    }
    appendParagraphInCell_(signCell, settings.SIGNATURE_NAME || settings.PRINCIPAL_NAME || 'Authorized Admin', { alignment: DocumentApp.HorizontalAlignment.CENTER, bold: true, fontSize: 11, foregroundColor: palette.text, spacingAfter: 1 });
    appendParagraphInCell_(signCell, 'Authorized Admin Signature', { alignment: DocumentApp.HorizontalAlignment.CENTER, fontSize: 9, foregroundColor: palette.sub });

    doc.saveAndClose();
    var pdfBlob = file.getAs(MimeType.PDF).setName(fileName);
    return { blob: pdfBlob, fileName: fileName, mimeType: 'application/pdf', subjectCount: total };
  } finally {
    try { doc.saveAndClose(); } catch (saveErr) {}
    try { file.setTrashed(true); } catch (trashErr) {}
  }
}

function generateStudentPdf_(payload) {
  var session = requireSession_(payload.token, 'student', sanitizeValue_(payload.clientId));
  var bundle = getStudentResultBundle_(session, payload);
  bundle.selectedSubjects = Array.isArray(payload.selectedSubjects) ? payload.selectedSubjects : [];
  bundle.resultDesign = sanitizeValue_(payload.resultDesign || bundle.resultDesign || 'royal');
  var generated = createStudentPdfBlob_(bundle);
  var bytes = generated.blob.getBytes();
  return ok_('Student PDF generated successfully.', {
    fileName: generated.fileName,
    mimeType: generated.mimeType,
    subjectCount: generated.subjectCount,
    pdfBase64: Utilities.base64Encode(bytes)
  });
}

function getImageDataUrl_(payload) {
  requireSession_(payload.token, '', sanitizeValue_(payload.clientId));
  var loaded = loadImageBlobFromSource_(sanitizeValue_((payload && (payload.url || payload.imageUrl || payload.src || payload.fileId)) || ''));
  var bytes = loaded.blob.getBytes();
  if (!bytes || !bytes.length) throw new Error('Image data is empty.');
  return ok_('Image data prepared.', {
    source: loaded.source,
    mimeType: loaded.mimeType,
    byteLength: bytes.length,
    dataUrl: 'data:' + loaded.mimeType + ';base64,' + Utilities.base64Encode(bytes)
  });
}

function sendBulkEmail_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var target = sanitizeValue_(payload.target || 'students').toLowerCase();
  var subject = sanitizeValue_(payload.subject);
  var message = sanitizeValue_(payload.message);
  if (!subject || !message) throw new Error('Subject and message are required.');
  var recipients = [];
  if (target === 'admins') {
    recipients = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS)).filter(function(a) {
      return sanitizeEmail_(a.Email) && normalizeBoolean_(a.Active, true) && !normalizeBoolean_(a.Archived, false) && !normalizeBoolean_(a.Deleted, false);
    }).map(function(a) { return sanitizeEmail_(a.Email); });
  } else {
    var selected = Array.isArray(payload.selected) ? payload.selected.map(sanitizeRegId_) : [];
    recipients = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS)).filter(function(s) {
      if (selected.length && selected.indexOf(sanitizeRegId_(s.RegID)) === -1) return false;
      return sanitizeEmail_(s.ParentEmail) && normalizeBoolean_(s.Active, true) && !normalizeBoolean_(s.Archived, false) && !normalizeBoolean_(s.Deleted, false);
    }).map(function(s) { return sanitizeEmail_(s.ParentEmail); });
  }
  recipients = uniqueList_(recipients);
  if (!recipients.length) throw new Error('No email recipients were found.');
  var quota = MailApp.getRemainingDailyQuota();
  var limit = Math.min(quota, 50, recipients.length);
  if (limit <= 0) throw new Error('Mail quota is exhausted for today.');
  recipients.slice(0, limit).forEach(function(email) {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: message.replace(/\n/g, '<br>') });
  });
  return ok_(limit < recipients.length ? 'Bulk email sent to the allowed quota limit.' : 'Bulk email sent successfully.', {
    recipients: limit,
    totalMatched: recipients.length
  });
}

function authorizeMailStatus_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var quota = MailApp.getRemainingDailyQuota();
  return ok_('Mail permission is available.', { authorized: true, remainingDailyQuota: quota });
}

function exportSmsContacts_(payload) {
  requireAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var target = sanitizeValue_(payload.target || 'students').toLowerCase();
  var rows = [];
  if (target === 'admins') {
    rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS)).filter(function(a) {
      return sanitizeValue_(a.Phone) && !normalizeBoolean_(a.Deleted, false);
    }).map(function(a) {
      return { name: a.DisplayName || a.Username, phone: a.Phone || '', email: a.Email || '' };
    });
  } else {
    rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS)).filter(function(s) {
      return sanitizeValue_(s.ParentPhone) && !normalizeBoolean_(s.Deleted, false);
    }).map(function(s) {
      return { name: s.FullName || s.RegID, phone: s.ParentPhone || '', email: s.ParentEmail || '', regId: s.RegID || '' };
    });
  }
  return ok_('SMS contacts export prepared.', {
    filename: target + '_sms_contacts.csv',
    csv: objectsToCsv_(rows, Object.keys(rows[0] || { name:'', phone:'', email:'' }))
  });
}


function bulkSetResultState_(payload) {
  var session = requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId));
  var ids = Array.isArray(payload.resultIds) ? payload.resultIds : [];
  var state = sanitizeValue_(payload.state).toLowerCase();
  if (!ids.length) throw new Error('Select at least one result first.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  var rows = getSheetObjectsWithIndex_(sh);
  var changed = 0;
  ids.forEach(function(id) {
    var row = rows.find(function(item){ return String(item.obj.ResultID || '') === String(id || ''); });
    if (!row) return;
    if (state === 'harddelete') {
      sh.deleteRow(row.rowIndex);
      rows = getSheetObjectsWithIndex_(sh);
    } else {
      var patch = resultStatePatch_(state, row.obj);
      patch.UpdatedAt = isoNow_();
      updateObjectRow_(sh, row.rowIndex, patch);
    }
    changed++;
  });
  logAudit_('admin', session.id, 'bulkSetResultState', 'OK', state + ' x' + changed);
  return ok_(changed + ' result(s) updated.');
}

function requestPasswordReset_(payload) {
  var role = sanitizeValue_(payload.role).toLowerCase();
  var identifier = sanitizeValue_(payload.identifier);
  if (!identifier || (role !== 'admin' && role !== 'student')) throw new Error('Invalid password reset request.');
  enforceRateLimit_('password-reset:' + role + ':' + identifier.toLowerCase(), 4, 900, 'Too many reset requests. Try again later.');
  var account = findAccountForReset_(role, identifier);
  if (account && account.email) {
    var code = generateResetCode_();
    storeResetCode_(role, account.id, code);
    MailApp.sendEmail({
      to: account.email,
      subject: 'Password Reset Code',
      htmlBody: 'Your reset code is <strong>' + code + '</strong>. It will expire in ' + RESET_CODE_TTL_MINUTES + ' minutes.'
    });
  }
  return ok_('If the account exists, a reset code has been sent to the registered email address.');
}

function resetPassword_(payload) {
  var role = sanitizeValue_(payload.role).toLowerCase();
  var identifier = sanitizeValue_(payload.identifier);
  var code = sanitizeValue_(payload.code);
  var newPassword = String(payload.newPassword || '');
  if (!identifier || !code || !newPassword) throw new Error('Identifier, reset code, and new password are required.');
  validatePasswordStrength_(newPassword, 'New password');
  var account = findAccountForReset_(role, identifier);
  if (!account) throw new Error('Invalid reset request.');
  if (!consumeResetCode_(role, account.id, code)) throw new Error('Invalid or expired reset code.');
  var creds = createPasswordRecord_(newPassword);
  updateObjectRow_(account.sheet, account.rowIndex, {
    PasswordHash: creds.hash,
    PasswordSalt: creds.salt,
    PasswordVersion: creds.version,
    UpdatedAt: isoNow_()
  });
  return ok_('Password reset successful.');
}

function setAdminState_(payload) {
  var session = requirePrincipalAdmin_(payload.token, sanitizeValue_(payload.clientId));
  var username = sanitizeValue_(payload.username).toLowerCase();
  var state = sanitizeValue_(payload.state).toLowerCase();
  if (!username || !state) throw new Error('Admin and state are required.');
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  var row = findObjectRow_(sh, function(obj) { return String(obj.Username || '').toLowerCase() === username; });
  if (!row) throw new Error('Admin not found.');
  if (normalizeBoolean_(row.obj.IsPrincipal, false)) throw new Error('The principal admin cannot be changed here.');
  if (state === 'harddelete') {
    sh.deleteRow(row.rowIndex);
    return ok_('Admin deleted forever.');
  }
  updateObjectRow_(sh, row.rowIndex, Object.assign(adminStatePatch_(state), { UpdatedAt: isoNow_() }));
  logAudit_('admin', session.id, 'setAdminState', 'OK', username + ' -> ' + state);
  return ok_('Admin state updated.');
}

function authorizeMailAccess() {
  return MailApp.getRemainingDailyQuota();
}

function requireAdmin_(token, clientId) {
  return requireSession_(token, 'admin', clientId);
}

function requirePrincipalAdmin_(token, clientId) {
  ensurePrincipalAdminConsistency_();
  var session = requireSession_(token, 'admin', clientId);
  var row = getAdminRowByUsername_(session.id);
  if (row && normalizeBoolean_(row.obj.IsPrincipal, false)) {
    session.isPrincipal = true;
    touchSession_(session.token, session);
    return session;
  }

  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS);
  var activeAdmins = getSheetObjectsWithIndex_(sh).filter(function(item) {
    return !normalizeBoolean_(item.obj.Deleted, false) && !normalizeBoolean_(item.obj.Archived, false) && normalizeBoolean_(item.obj.Active, true);
  });
  var principalRow = activeAdmins.find(function(item) { return normalizeBoolean_(item.obj.IsPrincipal, false); }) || null;

  var canPromoteCaller = !!row && normalizeBoolean_(row.obj.Active, true) && !normalizeBoolean_(row.obj.Archived, false) && !normalizeBoolean_(row.obj.Deleted, false) && (
    !principalRow ||
    principalRow.rowIndex === row.rowIndex ||
    (
      String(principalRow.obj.Username || '').toLowerCase() === 'admin' &&
      !sanitizeEmail_(principalRow.obj.Email) &&
      !sanitizeValue_(principalRow.obj.Phone) &&
      !sanitizeValue_(principalRow.obj.LastLoginAt)
    )
  );

  if (canPromoteCaller) {
    activeAdmins.forEach(function(item) {
      var shouldPrincipal = item.rowIndex === row.rowIndex;
      if (normalizeBoolean_(item.obj.IsPrincipal, false) !== shouldPrincipal) {
        updateObjectRow_(sh, item.rowIndex, { IsPrincipal: shouldPrincipal, UpdatedAt: isoNow_() });
      }
    });
    session.isPrincipal = true;
    touchSession_(session.token, session);
    return session;
  }

  if (!session.isPrincipal) throw new Error('Only the principal admin can perform this action.');
  return session;
}

function requireSession_(token, expectedRole, clientId) {
  token = sanitizeValue_(token);
  if (!token) throw new Error('Session expired. Please log in again.');
  var raw = CacheService.getScriptCache().get('SESSION_' + token);
  if (!raw) throw new Error('Session expired. Please log in again.');
  var session = parseJson_(raw);
  if (!session || !session.role) throw new Error('Session expired. Please log in again.');
  if (expectedRole && session.role !== expectedRole) throw new Error('You are not authorized for this action.');
  if (clientId && session.clientId && session.clientId !== clientId) throw new Error('This session belongs to another browser. Please log in again.');
  return session;
}

function createSession_(role, id, clientId, extra) {
  var token = Utilities.getUuid();
  var session = Object.assign({ role: role, id: id, clientId: clientId, createdAt: isoNow_(), token: token }, extra || {});
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), SESSION_TTL_SECONDS);
  return session;
}


function touchSession_(token, session) {
  token = sanitizeValue_(token);
  if (!token || !session) return;
  session.token = token;
  session.touchedAt = isoNow_();
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), SESSION_TTL_SECONDS);
}

function requireClientId_(clientId) {
  if (!clientId) throw new Error('Browser security token missing. Refresh and try again.');
}

function enforceRateLimit_(key, maxCount, ttlSeconds, message) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('RATE_' + key);
  var count = raw ? Number(raw) : 0;
  count++;
  cache.put('RATE_' + key, String(count), ttlSeconds);
  if (count > maxCount) throw new Error(message || 'Too many requests. Please try again later.');
}

function findAccountForReset_(role, identifier) {
  identifier = sanitizeValue_(identifier);
  if (role === 'admin') {
    return findInSheetForReset_(getSpreadsheet_().getSheetByName(SHEET_NAMES.ADMINS), function(obj) {
      return String(obj.Username || '').toLowerCase() === identifier.toLowerCase() || sanitizeEmail_(obj.Email) === sanitizeEmail_(identifier);
    }, 'Username', 'Email');
  }
  return findInSheetForReset_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS), function(obj) {
    return sanitizeRegId_(obj.RegID) === sanitizeRegId_(identifier) || sanitizeEmail_(obj.ParentEmail) === sanitizeEmail_(identifier);
  }, 'RegID', 'ParentEmail');
}

function findInSheetForReset_(sheet, matcher) {
  var found = getSheetObjectsWithIndex_(sheet).find(function(item) { return matcher(item.obj); });
  if (!found) return null;
  return {
    sheet: sheet,
    rowIndex: found.rowIndex,
    id: sanitizeValue_(found.obj.Username || found.obj.RegID),
    email: sanitizeEmail_(found.obj.Email || found.obj.ParentEmail)
  };
}

function storeResetCode_(role, identifier, code) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESET_CODES);
  appendObjectRow_(sh, {
    Role: role,
    Identifier: identifier,
    CodeHash: simpleHash_(code),
    ExpiresAt: isoDateOffsetMinutes_(RESET_CODE_TTL_MINUTES),
    Consumed: false,
    CreatedAt: isoNow_()
  });
}

function consumeResetCode_(role, identifier, code) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESET_CODES);
  var rows = getSheetObjectsWithIndex_(sh).filter(function(item) {
    return String(item.obj.Role || '') === role && String(item.obj.Identifier || '') === identifier && !normalizeBoolean_(item.obj.Consumed, false);
  });
  var valid = rows.find(function(item) {
    return String(item.obj.CodeHash || '') === simpleHash_(code) && new Date(item.obj.ExpiresAt).getTime() >= Date.now();
  });
  if (!valid) return false;
  updateObjectRow_(sh, valid.rowIndex, { Consumed: true });
  return true;
}

function getStudentByRegId_(regId) {
  var rows = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS));
  return rows.find(function(r) { return sanitizeRegId_(r.RegID) === sanitizeRegId_(regId); }) || null;
}

function normalizeStudentPayload_(obj) {
  return {
    RegID: sanitizeRegId_(obj.RegID),
    PasswordHash: sanitizeValue_(obj.PasswordHash),
    PasswordSalt: sanitizeValue_(obj.PasswordSalt),
    PasswordVersion: sanitizeValue_(obj.PasswordVersion || PASSWORD_HASH_VERSION),
    FullName: sanitizeValue_(obj.FullName),
    Age: sanitizeValue_(obj.Age),
    Gender: sanitizeValue_(obj.Gender),
    DOB: sanitizeValue_(obj.DOB),
    ParentName: sanitizeValue_(obj.ParentName),
    ParentPhone: sanitizeValue_(obj.ParentPhone),
    ParentEmail: sanitizeEmail_(obj.ParentEmail),
    SchoolName: sanitizeValue_(obj.SchoolName),
    State: sanitizeValue_(obj.State),
    CityLGA: sanitizeValue_(obj.CityLGA),
    ClassLevel: sanitizeValue_(obj.ClassLevel),
    Category: sanitizeValue_(obj.Category),
    Address: sanitizeValue_(obj.Address),
    PassportUrl: normalizeImageUrl_(obj.PassportUrl),
    Active: normalizeBoolean_(obj.Active, true),
    Archived: normalizeBoolean_(obj.Archived, false),
    Deleted: normalizeBoolean_(obj.Deleted, false),
    CreatedAt: sanitizeValue_(obj.CreatedAt || isoNow_()),
    UpdatedAt: sanitizeValue_(obj.UpdatedAt || isoNow_()),
    LastLoginAt: sanitizeValue_(obj.LastLoginAt)
  };
}

function normalizeResultEntry_(src, settings) {
  var examCode = sanitizeValue_(src.examCode || src.ExamCode).toUpperCase();
  var maxScore = maybeNumber_(src.maxScore || src.MaxScore, '');
  var studentScore = maybeNumber_(src.studentScore || src.StudentScore, '');
  var passMarkNumber = maybeNumber_(src.passMarkNumber || src.PassMarkNumber || settings.PASS_MARK_NUMBER, '');
  var passMarkPercentage = maybeNumber_(src.passMarkPercentage || src.PassMarkPercentage || settings.PASS_MARK_PERCENTAGE, '');
  var percentage = (studentScore !== '' && maxScore !== '' && Number(maxScore) > 0) ? round2_((Number(studentScore) / Number(maxScore)) * 100) : '';
  return {
    ResultID: sanitizeValue_(src.resultId || src.ResultID || Utilities.getUuid()),
    RegID: sanitizeRegId_(src.regId || src.RegID),
    ExamCode: examCode,
    ExamTitle: sanitizeValue_(src.examTitle || src.ExamTitle || settings.EXAM_TITLE),
    Subject: sanitizeValue_(src.subject || src.Subject),
    ExamDate: sanitizeValue_(src.examDate || src.ExamDate),
    MaxScore: maxScore,
    PassMarkNumber: passMarkNumber,
    PassMarkPercentage: passMarkPercentage,
    StudentScore: studentScore,
    Percentage: percentage,
    Position: maybeNumber_(src.position || src.Position, ''),
    Grade: percentage === '' ? '' : getGradeFromPercentage_(Number(percentage)),
    Remark: percentage === '' ? '' : getRemarkFromPercentage_(Number(percentage)),
    TeacherComment: sanitizeValue_(src.teacherComment || src.TeacherComment),
    AcademicSession: sanitizeValue_(src.academicSession || src.AcademicSession || settings.ACADEMIC_SESSION),
    Term: sanitizeValue_(src.term || src.Term || settings.TERM),
    Published: normalizeBoolean_(src.published || src.Published, false),
    ViewActive: normalizeBoolean_(src.viewActive || src.ViewActive, false),
    PublishedAt: normalizeBoolean_(src.published || src.Published, false) ? sanitizeValue_(src.publishedAt || src.PublishedAt || isoNow_()) : '',
    SignatureUrl: normalizeImageUrl_(src.signatureUrl || src.SignatureUrl || settings.SIGNATURE_URL),
    Archived: normalizeBoolean_(src.archived || src.Archived, false),
    Deleted: normalizeBoolean_(src.deleted || src.Deleted, false),
    CreatedAt: sanitizeValue_(src.createdAt || src.CreatedAt || isoNow_()),
    UpdatedAt: isoNow_()
  };
}

function recalculateRankings_(examCode, academicSession, term) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.RESULTS);
  var rows = getSheetObjectsWithIndex_(sh);
  var filtered = rows.filter(function(item) {
    var obj = item.obj;
    if (normalizeBoolean_(obj.Deleted, false) || normalizeBoolean_(obj.Archived, false)) return false;
    if (examCode && sanitizeValue_(obj.ExamCode).toLowerCase() !== sanitizeValue_(examCode).toLowerCase()) return false;
    if (academicSession && sanitizeValue_(obj.AcademicSession) !== sanitizeValue_(academicSession)) return false;
    if (term && sanitizeValue_(obj.Term) !== sanitizeValue_(term)) return false;
    return true;
  });
  var groups = {};
  filtered.forEach(function(item) {
    var key = [sanitizeValue_(item.obj.ExamCode).toLowerCase(), sanitizeValue_(item.obj.Subject).toLowerCase(), sanitizeValue_(item.obj.AcademicSession), sanitizeValue_(item.obj.Term)].join('|');
    groups[key] = groups[key] || [];
    groups[key].push(item);
  });
  var affected = 0;
  Object.keys(groups).forEach(function(key) {
    var group = groups[key].slice().sort(function(a, b) {
      return Number(b.obj.StudentScore || -1) - Number(a.obj.StudentScore || -1);
    });
    var position = 0;
    var lastScore = null;
    group.forEach(function(item, idx) {
      var score = item.obj.StudentScore === '' ? null : Number(item.obj.StudentScore);
      var maxScore = item.obj.MaxScore === '' ? null : Number(item.obj.MaxScore);
      if (score == null || maxScore == null || maxScore <= 0) {
        updateObjectRow_(sh, item.rowIndex, {
          Position: '',
          Percentage: '',
          Grade: '',
          Remark: '',
          UpdatedAt: isoNow_()
        });
        affected++;
        return;
      }
      if (lastScore === null || score < lastScore) position = idx + 1;
      lastScore = score;
      var percentage = round2_((score / maxScore) * 100);
      updateObjectRow_(sh, item.rowIndex, {
        Position: position,
        Percentage: percentage,
        Grade: getGradeFromPercentage_(percentage),
        Remark: getRemarkFromPercentage_(percentage),
        UpdatedAt: isoNow_()
      });
      affected++;
    });
  });
  return affected;
}

function getRemarkFromPercentage_(percentage) {
  var bands = getSheetObjects_(getSpreadsheet_().getSheetByName(SHEET_NAMES.REMARKS));
  var band = bands.find(function(item) {
    return Number(percentage) >= Number(item.MinPercent) && Number(percentage) <= Number(item.MaxPercent);
  });
  return band ? String(band.Remark || '') : 'No remark configured for this score band.';
}

function getGradeFromPercentage_(percentage) {
  percentage = Number(percentage);
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
  var hasScore = !(score === '' || score == null);
  var hasPct = !(percentage === '' || percentage == null);
  if (!hasScore && !hasPct) return 'PENDING';
  var numberOk = hasScore ? Number(score) >= Number(passMarkNumber || 0) : true;
  var pctOk = hasPct ? Number(percentage) >= Number(passMarkPercentage || 0) : true;
  return (numberOk && pctOk) ? 'PASS' : 'FAIL';
}

function cleanAdmin_(r) {
  return {
    username: sanitizeValue_(r.Username),
    displayName: sanitizeValue_(r.DisplayName),
    email: sanitizeValue_(r.Email),
    phone: sanitizeValue_(r.Phone),
    isPrincipal: normalizeBoolean_(r.IsPrincipal, false),
    active: normalizeBoolean_(r.Active, true),
    archived: normalizeBoolean_(r.Archived, false),
    deleted: normalizeBoolean_(r.Deleted, false),
    createdAt: sanitizeValue_(r.CreatedAt),
    updatedAt: sanitizeValue_(r.UpdatedAt),
    lastLoginAt: sanitizeValue_(r.LastLoginAt)
  };
}

function cleanStudent_(r) {
  return {
    regId: sanitizeRegId_(r.RegID),
    fullName: sanitizeValue_(r.FullName),
    age: sanitizeValue_(r.Age),
    gender: sanitizeValue_(r.Gender),
    dob: sanitizeValue_(r.DOB),
    parentName: sanitizeValue_(r.ParentName),
    parentPhone: sanitizeValue_(r.ParentPhone),
    parentEmail: sanitizeValue_(r.ParentEmail),
    schoolName: sanitizeValue_(r.SchoolName),
    state: sanitizeValue_(r.State),
    cityLGA: sanitizeValue_(r.CityLGA),
    classLevel: sanitizeValue_(r.ClassLevel),
    category: sanitizeValue_(r.Category),
    address: sanitizeValue_(r.Address),
    passportUrl: normalizeImageUrl_(r.PassportUrl),
    active: normalizeBoolean_(r.Active, true),
    archived: normalizeBoolean_(r.Archived, false),
    deleted: normalizeBoolean_(r.Deleted, false),
    createdAt: sanitizeValue_(r.CreatedAt),
    updatedAt: sanitizeValue_(r.UpdatedAt),
    lastLoginAt: sanitizeValue_(r.LastLoginAt)
  };
}

function cleanResult_(r) {
  var studentScore = maybeNumber_(r.StudentScore, '');
  var maxScore = maybeNumber_(r.MaxScore, '');
  var percentage = maybeNumber_(r.Percentage, '');
  var position = maybeNumber_(r.Position, '');
  return {
    resultId: sanitizeValue_(r.ResultID),
    regId: sanitizeRegId_(r.RegID),
    examCode: sanitizeValue_(r.ExamCode),
    examTitle: sanitizeValue_(r.ExamTitle),
    subject: sanitizeValue_(r.Subject),
    examDate: sanitizeValue_(r.ExamDate),
    maxScore: maxScore,
    passMarkNumber: maybeNumber_(r.PassMarkNumber, ''),
    passMarkPercentage: maybeNumber_(r.PassMarkPercentage, ''),
    studentScore: studentScore,
    percentage: percentage,
    position: position,
    positionText: position === '' ? '' : ordinal_(Number(position)),
    grade: sanitizeValue_(r.Grade),
    remark: sanitizeValue_(r.Remark),
    teacherComment: sanitizeValue_(r.TeacherComment),
    academicSession: sanitizeValue_(r.AcademicSession),
    term: sanitizeValue_(r.Term),
    published: normalizeBoolean_(r.Published, false),
    viewActive: normalizeBoolean_(r.ViewActive, false),
    publishedAt: sanitizeValue_(r.PublishedAt),
    signatureUrl: normalizeImageUrl_(r.SignatureUrl),
    archived: normalizeBoolean_(r.Archived, false),
    deleted: normalizeBoolean_(r.Deleted, false),
    createdAt: sanitizeValue_(r.CreatedAt),
    updatedAt: sanitizeValue_(r.UpdatedAt),
    resultStatus: getResultStatus_(studentScore, percentage, r.PassMarkNumber, r.PassMarkPercentage)
  };
}

function getSheetObjects_(sheet) {
  return getSheetObjectsWithIndex_(sheet).map(function(item) { return item.obj; });
}

function getSheetObjectsWithIndex_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || !lastCol) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return rows.map(function(row, idx) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return { rowIndex: idx + 2, obj: obj };
  }).filter(function(item) {
    return Object.keys(item.obj).some(function(key) { return item.obj[key] !== ''; });
  });
}

function countNonBlankRows_(sheet) {
  return getSheetObjects_(sheet).length;
}

function getHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  return lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
}

function appendObjectRow_(sheet, obj) {
  var headers = getHeaders_(sheet);
  var row = headers.map(function(h) { return Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ''; });
  sheet.appendRow(row);
}

function updateObjectRow_(sheet, rowIndex, patch) {
  var headers = getHeaders_(sheet);
  var existing = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var data = {};
  headers.forEach(function(h, i) { data[h] = existing[i]; });
  Object.keys(patch).forEach(function(key) {
    data[key] = patch[key];
  });
  var row = headers.map(function(h) { return data[h]; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function findRowByValue_(sheet, headerName, value) {
  var headers = getHeaders_(sheet);
  var col = headers.indexOf(headerName);
  if (col === -1) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  var values = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues().flat();
  var idx = values.findIndex(function(v) { return String(v).trim() === String(value).trim(); });
  return idx >= 0 ? idx + 2 : null;
}

function findObjectRow_(sheet, matcher) {
  return getSheetObjectsWithIndex_(sheet).find(function(item) { return matcher(item.obj); }) || null;
}

function sanitizeValue_(value) {
  return value == null ? '' : String(value).trim();
}

function sanitizeEmail_(value) {
  return sanitizeValue_(value).toLowerCase();
}

function sanitizeRegId_(value) {
  return sanitizeValue_(value).toUpperCase();
}

function sanitizeSettingValue_(key, value) {
  if (key === 'BRAND_LOGO_URL' || key === 'SIGNATURE_URL' || key === 'FAVICON_URL') return normalizeImageUrl_(value);
  if (key === 'SHOW_POSITION_ON_REPORT' || key === 'STUDENT_SIGNUP_ENABLED' || key === 'PUBLIC_ADMIN_SIGNUP_ENABLED') return String(normalizeBoolean_(value, false));
  return sanitizeValue_(value);
}

function normalizeImageUrl_(url) {
  url = sanitizeValue_(url);
  if (!url) return '';
  var imgMatch = url.match(/<img[^>]+src=["']?([^"' >]+)["']?/i);
  if (imgMatch && imgMatch[1]) url = sanitizeValue_(imgMatch[1]);
  var cssMatch = url.match(/url\((['"]?)(.*?)\1\)/i);
  if (cssMatch && cssMatch[2]) url = sanitizeValue_(cssMatch[2]);
  var match = extractDriveFileId_(url);
  if (match) return 'https://drive.google.com/thumbnail?id=' + match + '&sz=w2000';
  if (/dropbox\.com/i.test(url)) {
    var next = url.replace(/\?dl=0$/i, '?raw=1').replace(/\?dl=1$/i, '?raw=1');
    if (!/[?&]raw=1/i.test(next)) next += (next.indexOf('?') >= 0 ? '&' : '?') + 'raw=1';
    return next;
  }
  var gh = url.match(/^https:\/\/github\.com\/([^\/]+)\/([^\/]+)\/blob\/([^\/]+)\/(.+)$/i);
  if (gh) return 'https://raw.githubusercontent.com/' + gh[1] + '/' + gh[2] + '/' + gh[3] + '/' + gh[4];
  return url;
}

function extractDriveFileId_(value) {
  var raw = sanitizeValue_(value);
  if (!raw) return '';
  var match = raw.match(/drive\.google\.com\/file\/d\/([^\/?#]+)/i);
  if (!match) match = raw.match(/drive\.google\.com\/open\?id=([^&]+)/i);
  if (!match) match = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/i);
  if (!match) match = raw.match(/thumbnail\?[^#]*id=([a-zA-Z0-9_-]+)/i);
  if (!match) match = raw.match(/uc\?export=(?:view|download)&id=([a-zA-Z0-9_-]+)/i);
  if (!match) match = raw.match(/drive\.usercontent\.google\.com\/(?:download|u\/\d+\/uc)\?[^#]*id=([a-zA-Z0-9_-]+)/i);
  if (!match && /^[A-Za-z0-9_-]{20,}$/.test(raw)) match = [raw, raw];
  return match && match[1] ? sanitizeValue_(match[1]) : '';
}

function buildPublicImageCandidates_(value) {
  var raw = sanitizeValue_(value);
  if (!raw) return [];
  var out = [];
  function push(nextValue) {
    var next = sanitizeValue_(nextValue);
    if (next && out.indexOf(next) === -1) out.push(next);
  }
  push(normalizeImageUrl_(raw));
  var driveId = extractDriveFileId_(raw);
  if (driveId) {
    push('https://drive.google.com/thumbnail?id=' + driveId + '&sz=w2000');
    push('https://drive.google.com/uc?export=view&id=' + driveId);
    push('https://drive.google.com/uc?id=' + driveId);
    push('https://drive.usercontent.google.com/download?id=' + driveId + '&export=view&authuser=0');
  }
  push(raw);
  return out;
}

function guessMimeTypeFromName_(name) {
  var lower = sanitizeValue_(name).toLowerCase();
  if (/\.png$/i.test(lower)) return 'image/png';
  if (/\.(jpg|jpeg)$/i.test(lower)) return 'image/jpeg';
  if (/\.webp$/i.test(lower)) return 'image/webp';
  if (/\.gif$/i.test(lower)) return 'image/gif';
  if (/\.svg$/i.test(lower)) return 'image/svg+xml';
  if (/\.bmp$/i.test(lower)) return 'image/bmp';
  return 'application/octet-stream';
}

function buildDriveImageUrls_(fileId) {
  var id = sanitizeValue_(fileId);
  return {
    fileId: id,
    thumbnailUrl: id ? ('https://drive.google.com/thumbnail?id=' + id + '&sz=w1600') : '',
    viewUrl: id ? ('https://drive.google.com/uc?export=view&id=' + id) : '',
    previewUrl: id ? ('https://drive.google.com/file/d/' + id + '/preview') : ''
  };
}

function safeName_(value) {
  return sanitizeValue_(value).replace(/[^a-zA-Z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'item';
}

function getRootStorageFolder_() {
  var folderName = 'Genz Result Checker Storage';
  var iter = DriveApp.getFoldersByName(folderName);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);
}

function getNestedFolder_(parts) {
  var folder = getRootStorageFolder_();
  (parts || []).forEach(function(part) {
    var name = sanitizeValue_(part);
    if (!name) return;
    var iter = folder.getFoldersByName(name);
    folder = iter.hasNext() ? iter.next() : folder.createFolder(name);
  });
  return folder;
}

function createPublicImageFile_(folderParts, fileName, mimeType, base64Data) {
  var bytes = Utilities.base64Decode(sanitizeValue_(base64Data));
  if (!bytes || !bytes.length) throw new Error('Invalid image data.');
  if (bytes.length > 8 * 1024 * 1024) throw new Error('Image is too large. Keep it under 8 MB.');
  var folder = getNestedFolder_(folderParts);
  var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName || ('image_' + Utilities.getUuid()));
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    try { file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW); } catch (innerErr) {}
  }
  return file;
}

function uploadStudentPassport_(payload) {
  var session = requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId));
  var fileName = sanitizeValue_(payload.fileName || payload.originalName);
  var mimeType = sanitizeValue_(payload.mimeType) || 'application/octet-stream';
  var base64Data = sanitizeValue_(payload.fileData || payload.base64Data);
  var regId = sanitizeValue_(payload.regId || payload.studentRegId || '');
  var fullName = sanitizeValue_(payload.fullName || payload.studentName || '');
  if (!fileName || !base64Data) throw new Error('Choose a passport image to upload.');
  if (!/^image\//i.test(mimeType) && !/icon/i.test(mimeType)) throw new Error('Only image files can be uploaded here.');
  var folderParts = ['Student Passports', safeName_(regId || fullName || 'General')];
  var stampedName = safeName_(regId || fullName || 'passport') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + safeName_(fileName);
  var file = createPublicImageFile_(folderParts, stampedName, mimeType, base64Data);
  var links = buildDriveImageUrls_(file.getId());
  var updated = false;
  if (regId) {
    var sh = getSpreadsheet_().getSheetByName(SHEET_NAMES.STUDENTS);
    var row = findRowByValue_(sh, 'RegID', regId);
    if (row) {
      updateObjectRow_(sh, row, { PassportUrl: links.thumbnailUrl || links.viewUrl || file.getUrl(), UpdatedAt: isoNow_() });
      updated = true;
    }
  }
  logAudit_('admin', session.id, 'uploadStudentPassport', 'OK', 'Passport upload for ' + (regId || fullName || fileName));
  return ok_('Student passport uploaded successfully.', {
    regId: regId,
    fullName: fullName,
    fileId: file.getId(),
    savedUrl: links.thumbnailUrl || links.viewUrl || file.getUrl(),
    thumbnailUrl: links.thumbnailUrl,
    viewUrl: links.viewUrl,
    previewUrl: links.previewUrl,
    updatedStudent: updated
  });
}

function uploadBrandingAsset_(payload) {
  var session = requireSession_(payload.token, 'admin', sanitizeValue_(payload.clientId));
  var settingKey = sanitizeValue_(payload.settingKey || payload.key).toUpperCase();
  var allowed = { BRAND_LOGO_URL: true, SIGNATURE_URL: true };
  if (!allowed[settingKey]) throw new Error('Unsupported branding asset type.');
  var fileName = sanitizeValue_(payload.fileName || payload.originalName);
  var mimeType = sanitizeValue_(payload.mimeType) || 'application/octet-stream';
  var base64Data = sanitizeValue_(payload.fileData || payload.base64Data);
  if (!fileName || !base64Data) throw new Error('Choose an image to upload.');
  if (!/^image\//i.test(mimeType) && !/icon/i.test(mimeType)) throw new Error('Only image files can be uploaded here.');
  var labelMap = { BRAND_LOGO_URL: 'Brand Logo', SIGNATURE_URL: 'Signature' };
  var folderParts = ['Branding Assets', labelMap[settingKey] || settingKey];
  var stampedName = safeName_(labelMap[settingKey] || settingKey) + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + safeName_(fileName);
  var file = createPublicImageFile_(folderParts, stampedName, mimeType, base64Data);
  var links = buildDriveImageUrls_(file.getId());
  var savedUrl = links.thumbnailUrl || links.viewUrl || file.getUrl();
  upsertSetting_(settingKey, savedUrl);
  logAudit_('admin', session.id, 'uploadBrandingAsset', 'OK', settingKey + ' uploaded');
  return ok_((labelMap[settingKey] || 'Image') + ' uploaded successfully.', {
    settingKey: settingKey,
    fileId: file.getId(),
    savedUrl: savedUrl,
    thumbnailUrl: links.thumbnailUrl,
    viewUrl: links.viewUrl,
    previewUrl: links.previewUrl
  });
}

function parseListSetting_(value) {
  return uniqueList_(sanitizeValue_(value).split(',').map(function(part) { return sanitizeValue_(part); }).filter(Boolean));
}

function uniqueList_(arr) {
  var seen = {};
  return arr.filter(function(item) {
    var key = String(item);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function normalizeBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  if (value === 1 || value === '1' || String(value).toLowerCase() === 'true' || String(value).toLowerCase() === 'yes') return true;
  if (value === 0 || value === '0' || String(value).toLowerCase() === 'false' || String(value).toLowerCase() === 'no') return false;
  return defaultValue;
}

function maybeNumber_(value, defaultValue) {
  if (value === '' || value == null) return defaultValue;
  var n = Number(value);
  return isNaN(n) ? defaultValue : n;
}

function round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function isoNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function isoDateOffsetMinutes_(minutes) {
  return Utilities.formatDate(new Date(Date.now() + minutes * 60000), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function ordinal_(n) {
  n = Number(n || 0);
  var j = n % 10, k = n % 100;
  if (j === 1 && k !== 11) return n + 'st';
  if (j === 2 && k !== 12) return n + 'nd';
  if (j === 3 && k !== 13) return n + 'rd';
  return n + 'th';
}

function formLabel_(key) {
  return key.replace(/([A-Z])/g, ' $1').replace(/^./, function(ch) { return ch.toUpperCase(); });
}

function toPascalCase_(key) {
  return key.charAt(0).toUpperCase() + key.slice(1);
}

function objectsToCsv_(rows, headers) {
  var safeHeaders = headers || Object.keys(rows[0] || {});
  var lines = [safeHeaders.join(',')];
  rows.forEach(function(row) {
    lines.push(safeHeaders.map(function(h) { return csvCell_(row[h]); }).join(','));
  });
  return lines.join('\n');
}

function csvCell_(value) {
  var text = String(value == null ? '' : value);
  if (/[,"\n]/.test(text)) return '"' + text.replace(/"/g, '""') + '"';
  return text;
}

function simpleHash_(text) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''));
  return digest.map(function(byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function createPasswordRecord_(password) {
  var salt = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  return { salt: salt, version: PASSWORD_HASH_VERSION, hash: hashPasswordV2_(password, salt) };
}

function hashPasswordV2_(password, salt) {
  var value = salt + '|' + String(password || '');
  for (var i = 0; i < PASSWORD_HASH_ROUNDS; i++) {
    value = simpleHash_(value + '|' + i);
  }
  return value;
}

function verifyPasswordAndUpgrade_(sheet, row, password) {
  var obj = row.obj;
  var version = sanitizeValue_(obj.PasswordVersion);
  var ok = false;
  if (version === PASSWORD_HASH_VERSION && obj.PasswordSalt) {
    ok = hashPasswordV2_(password, obj.PasswordSalt) === String(obj.PasswordHash || '');
  } else {
    ok = simpleHash_(password) === String(obj.PasswordHash || '');
    if (ok) {
      var creds = createPasswordRecord_(password);
      updateObjectRow_(sheet, row.rowIndex, {
        PasswordHash: creds.hash,
        PasswordSalt: creds.salt,
        PasswordVersion: creds.version,
        UpdatedAt: isoNow_()
      });
    }
  }
  return ok;
}

function validatePasswordStrength_(password, label) {
  label = label || 'Password';
  if (String(password || '').length < 8) throw new Error(label + ' must be at least 8 characters long.');
}

function generateResetCode_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function studentStatePatch_(state) {
  switch (state) {
    case 'activate': return { Active: true };
    case 'deactivate': return { Active: false };
    case 'archive': return { Archived: true, Active: false };
    case 'delete': return { Deleted: true, Active: false };
    case 'restore': return { Deleted: false, Archived: false, Active: true };
    default: throw new Error('Unsupported student state.');
  }
}

function adminStatePatch_(state) {
  switch (state) {
    case 'activate': return { Active: true };
    case 'deactivate': return { Active: false };
    case 'archive': return { Archived: true, Active: false };
    case 'delete': return { Deleted: true, Active: false };
    case 'restore': return { Deleted: false, Archived: false, Active: true };
    default: throw new Error('Unsupported admin state.');
  }
}

function resultStatePatch_(state, current) {
  current = current || {};
  switch (state) {
    case 'publish': return { Published: true, PublishedAt: current.PublishedAt || isoNow_() };
    case 'unpublish': return { Published: false, PublishedAt: '' };
    case 'viewactive': return { ViewActive: true };
    case 'viewinactive': return { ViewActive: false };
    case 'archive': return { Archived: true };
    case 'delete': return { Deleted: true, Published: false, ViewActive: false };
    case 'restore': return { Deleted: false, Archived: false };
    default: throw new Error('Unsupported result state.');
  }
}

function logAudit_(actorRole, actorId, action, status, details) {
  try {
    appendObjectRow_(getSpreadsheet_().getSheetByName(SHEET_NAMES.AUDIT_LOGS), {
      Timestamp: isoNow_(),
      ActorRole: actorRole,
      ActorId: actorId,
      Action: action,
      Status: status,
      Details: sanitizeValue_(details)
    });
  } catch (err) {}
}

function logAuditSafe_(action, params, status, err) {
  try {
    logAudit_('unknown', sanitizeValue_((params && (params.username || params.regId || params.identifier)) || ''), action, status, err && err.message ? err.message : String(err || ''));
  } catch (e) {}
}
