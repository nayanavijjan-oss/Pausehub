// ============================================================
//  Cuemath — Step Up Classroom Training Management System
//  Google Apps Script — SELF-HOSTED (no GitHub needed!)
//  
//  HOW TO USE:
//  1. Paste this entire file into Apps Script editor
//  2. Also create a file called "Page.html" and paste index_template.html there
//  3. Deploy as Web App → Execute as Me → Anyone can access
//  4. Open the URL — done! No GitHub needed.
// ============================================================

var SHEET_ID     = '1GQH39RfmmbyqoKIrZRyHRucNw3rMk2pnTT_4YCcCy44';
var MASTER_SHEET = 'Master';
var DATA_SHEET   = 'Requests';
var ADMIN_SHEET  = 'AdminConfig';

// ── doGet — serves the HTML dashboard ────────────────────────
function doGet(e) {
  var scriptUrl = ScriptApp.getService().getUrl();
  var template  = HtmlService.createTemplateFromFile('Page');
  template.scriptUrl = scriptUrl;
  var html = template.evaluate()
    .setTitle('Cuemath Training Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// ── doPost — ALL API actions ──────────────────────────────────
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'login')          return makeResponse(loginUser(body.username, body.password));
    if (action === 'lookupMobile')   return makeResponse(lookupMobile(body.mobile));
    if (action === 'getAll')         return makeResponse(getAllRequests(body.mentorFilter));
    if (action === 'getStats')       return makeResponse(getStats(body.mentorFilter));
    if (action === 'getUsers')       return makeResponse(getUsers());
    if (action === 'submitRequest')  return makeResponse(submitRequest(body.data));
    if (action === 'approveRequest') return makeResponse(approveRequest(body.data));
    if (action === 'rejectRequest')  return makeResponse(rejectRequest(body.data));
    if (action === 'queryRequest')   return makeResponse(queryRequest(body.data));
    if (action === 'addRemark')      return makeResponse(addRemark(body.data));
    if (action === 'addUser')        return makeResponse(addUser(body.data));
    if (action === 'deleteRequest')  return makeResponse(deleteRequest(body.data));

    return makeResponse({ success: false, error: 'Unknown action: ' + action });
  } catch(err) {
    return makeResponse({ success: false, error: err.message });
  }
}

// ── JSON response helper ──────────────────────────────────────
function makeResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Sheet helpers ─────────────────────────────────────────────
function getMasterSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(MASTER_SHEET);
}

function getDataSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s  = ss.getSheetByName(DATA_SHEET);
  if (!s) s = ss.insertSheet(DATA_SHEET);
  return s;
}

function getAdminSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s  = ss.getSheetByName(ADMIN_SHEET);
  if (!s) {
    s = ss.insertSheet(ADMIN_SHEET);
    s.getRange('A1:E1').setValues([['username','password','mentorName','','role']]);
    s.getRange('A1:E1').setFontWeight('bold').setBackground('#FFA500').setFontColor('#FFFFFF');
    s.getRange('A2:E2').setValues([['admin','admin123','','','admin']]);
    s.getRange('A3:E3').setValues([['ops','ops123','','','admin']]);
    s.setFrozenRows(1);
  }
  return s;
}

// ── Ensure Requests sheet headers ─────────────────────────────
function ensureHeaders() {
  var sheet = getDataSheet();
  if (sheet.getLastRow() === 0 || sheet.getRange('A1').getValue() !== 'RequestID') {
    var h = [
      'RequestID','SubmittedAt','DBID','TeacherName','Mobile','Email',
      'MentorName','SubmittedBy','TrainingStage','Status',
      'PauseStartDate','PauseReason','ResumeDate','EmailConfirmed',
      'ResignReason','ResignScreenshot',
      'LastResponseDate','FollowUpAttempts',
      'FailReason','DisqualifyReason',
      'BackToTrainingStep','BackToTrainingDate',
      'Notes','ApprovalStatus','ApprovedBy','ApprovedAt',
      'RejectedBy','RejectedAt','RejectionReason',
      'Remarks','LastUpdated'
    ];
    sheet.getRange(1,1,1,h.length).setValues([h]);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#FF8C00').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

// ── Generate Request ID ───────────────────────────────────────
function generateRequestId() {
  var sheet = getDataSheet();
  var last  = Math.max(sheet.getLastRow(), 1);
  return 'REQ-' + String(last).padStart(4,'0');
}

// ── LOGIN ─────────────────────────────────────────────────────
function loginUser(username, password) {
  if (!username || !password)
    return { success: false, error: 'Username and password required.' };

  var s    = getAdminSheet();
  var last = s.getLastRow();
  if (last >= 2) {
    var rows = s.getRange(2, 1, last-1, 5).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase() &&
          String(rows[i][1]).trim() === String(password).trim()) {
        return { success: true, user: {
          username: String(rows[i][0]).trim(),
          role:     String(rows[i][4]).trim() || 'admin',
          mentor:   String(rows[i][2]).trim() || String(rows[i][0]).trim()
        }};
      }
    }
  }

  // Mentor auto-login from Master sheet
  var ms    = getMasterSheet();
  var mlast = ms.getLastRow();
  if (mlast >= 2) {
    var mrows = ms.getRange(2,1,mlast-1,5).getValues();
    for (var j = 0; j < mrows.length; j++) {
      var mname = String(mrows[j][4]).trim();
      if (mname.toLowerCase() === String(username).trim().toLowerCase()
          && String(password).trim() === 'mentor123') {
        return { success: true, user: { username: mname, role: 'mentor', mentor: mname }};
      }
    }
  }
  return { success: false, error: 'Invalid username or password. Please try again.' };
}

// ── MOBILE LOOKUP ─────────────────────────────────────────────
function lookupMobile(mobile) {
  if (!mobile) return { success: false, error: 'Mobile number required.' };
  mobile = String(mobile).trim().replace(/\D/g,'');
  if (mobile.length !== 10)
    return { success: false, error: 'Please provide a valid 10-digit registered mobile number.' };

  var sheet = getMasterSheet();
  var last  = sheet.getLastRow();
  if (last < 2) return { success: false, error: 'Master data not found.' };

  var data = sheet.getRange(2,1,last-1,5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][2]).trim().replace(/\D/g,'') === mobile) {
      return { success: true, teacher: {
        dbid:   data[i][0], name:  data[i][1],
        mobile: data[i][2], email: data[i][3], mentor: data[i][4]
      }};
    }
  }
  return { success: false, error: 'Mobile number not registered. Please provide a valid 10-digit registered mobile number.' };
}

// ── GET USERS ─────────────────────────────────────────────────
function getUsers() {
  var s    = getAdminSheet();
  var last = s.getLastRow();
  if (last < 2) return { success: true, data: [] };
  var rows = s.getRange(2,1,last-1,5).getValues();
  return { success: true, data: rows
    .filter(function(r){ return String(r[0]).trim(); })
    .map(function(r){ return { username: r[0], role: r[4] || 'admin' }; }) };
}

function addUser(d) {
  getAdminSheet().appendRow([d.username, d.password, d.mentorName||'', '', d.role||'mentor']);
  return { success: true };
}

// ── SUBMIT REQUEST ────────────────────────────────────────────
function submitRequest(d) {
  ensureHeaders();
  var sheet = getDataSheet();
  var id    = generateRequestId();
  var now   = new Date();
  sheet.appendRow([
    id, now,
    d.dbid||'', d.teacherName||'', d.mobile||'', d.email||'',
    d.mentorName||'', d.submittedBy||'', d.trainingStage||'', d.status||'',
    d.pauseStartDate||'', d.pauseReason||'', d.resumeDate||'', d.emailConfirmed||'',
    d.resignReason||'', d.resignScreenshot||'',
    d.lastResponseDate||'', d.followUpAttempts||'',
    d.failReason||'', d.disqualifyReason||'',
    d.backToTrainingStep||'', d.backToTrainingDate||'',
    d.notes||'',
    'Pending','','','','','','',now
  ]);
  colorRow(sheet, sheet.getLastRow(), 'Pending');
  return { success: true, id: id };
}

// ── APPROVE ───────────────────────────────────────────────────
function approveRequest(d) {
  var sheet = getDataSheet();
  var rows  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      var now = new Date();
      var ts  = Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm');
      sheet.getRange(i+1,24).setValue('Approved');
      sheet.getRange(i+1,25).setValue(d.approvedBy);
      sheet.getRange(i+1,26).setValue(now);
      sheet.getRange(i+1,31).setValue(now);
      var existing = String(rows[i][29]||'');
      var remark   = '['+ts+' — '+d.approvedBy+']: ✅ APPROVED. '+(d.note||'');
      sheet.getRange(i+1,30).setValue(existing ? existing+'\n'+remark : remark);
      colorRow(sheet, i+1, 'Approved');
      return { success: true };
    }
  }
  return { success: false, error: 'Request not found.' };
}

// ── REJECT ────────────────────────────────────────────────────
function rejectRequest(d) {
  var sheet = getDataSheet();
  var rows  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      var now = new Date();
      var ts  = Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm');
      sheet.getRange(i+1,24).setValue('Rejected');
      sheet.getRange(i+1,27).setValue(d.rejectedBy);
      sheet.getRange(i+1,28).setValue(now);
      sheet.getRange(i+1,29).setValue(d.reason||'');
      sheet.getRange(i+1,31).setValue(now);
      var existing = String(rows[i][29]||'');
      var remark   = '['+ts+' — '+d.rejectedBy+']: ❌ REJECTED. Reason: '+(d.reason||'');
      sheet.getRange(i+1,30).setValue(existing ? existing+'\n'+remark : remark);
      colorRow(sheet, i+1, 'Rejected');
      return { success: true };
    }
  }
  return { success: false, error: 'Request not found.' };
}

// ── QUERY ─────────────────────────────────────────────────────
function queryRequest(d) {
  var sheet = getDataSheet();
  var rows  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      var now = new Date();
      var ts  = Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm');
      sheet.getRange(i+1,24).setValue('Needs Clarification');
      sheet.getRange(i+1,31).setValue(now);
      var existing = String(rows[i][29]||'');
      var remark   = '['+ts+' — '+d.by+']: 💬 QUERY: '+(d.question||'');
      sheet.getRange(i+1,30).setValue(existing ? existing+'\n'+remark : remark);
      colorRow(sheet, i+1, 'Needs Clarification');
      return { success: true };
    }
  }
  return { success: false, error: 'Request not found.' };
}

// ── ADD REMARK ────────────────────────────────────────────────
function addRemark(d) {
  var sheet = getDataSheet();
  var rows  = sheet.getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      var now = new Date();
      var ts  = Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm');
      var existing = String(rows[i][29]||'');
      var remark   = '['+ts+' — '+(d.by||'User')+']: '+d.remark;
      sheet.getRange(i+1,30).setValue(existing ? existing+'\n'+remark : remark);
      sheet.getRange(i+1,31).setValue(now);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found.' };
}

// ── DELETE ────────────────────────────────────────────────────
function deleteRequest(d) {
  var sheet = getDataSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(d.id)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found.' };
}

// ── GET ALL REQUESTS ──────────────────────────────────────────
function getAllRequests(mentorFilter) {
  ensureHeaders();
  var sheet = getDataSheet();
  var last  = sheet.getLastRow();
  if (last <= 1) return { success: true, data: [] };

  var rows = sheet.getRange(2,1,last-1,31).getValues();
  var tz   = Session.getScriptTimeZone();

  function fmt(v) {
    if (!v || v === '') return '';
    try { return Utilities.formatDate(new Date(v), tz, 'yyyy-MM-dd HH:mm'); }
    catch(e) { return String(v); }
  }

  var data = rows
    .filter(function(r){ return String(r[0]).trim() !== ''; })
    .map(function(r) {
      return {
        id:r[0], submittedAt:fmt(r[1]),
        dbid:r[2], teacherName:r[3], mobile:r[4], email:r[5],
        mentorName:r[6], submittedBy:r[7], trainingStage:r[8], status:r[9],
        pauseStartDate:r[10], pauseReason:r[11], resumeDate:r[12], emailConfirmed:r[13],
        resignReason:r[14], resignScreenshot:r[15],
        lastResponseDate:r[16], followUpAttempts:r[17],
        failReason:r[18], disqualifyReason:r[19],
        backToTrainingStep:r[20], backToTrainingDate:r[21],
        notes:r[22], approvalStatus:r[23],
        approvedBy:r[24], approvedAt:fmt(r[25]),
        rejectedBy:r[26], rejectedAt:fmt(r[27]), rejectionReason:r[28],
        remarks:r[29], lastUpdated:fmt(r[30])
      };
    });

  if (mentorFilter) {
    data = data.filter(function(r){
      return String(r.mentorName).toLowerCase() === String(mentorFilter).toLowerCase();
    });
  }
  return { success: true, data: data };
}

// ── STATS ─────────────────────────────────────────────────────
function getStats(mentorFilter) {
  var all   = getAllRequests(mentorFilter).data;
  var today = new Date();
  var stats = {
    success:true, total:all.length,
    pending:0, approved:0, rejected:0, needsClarity:0,
    nonResponsiveFlag:0, byStatus:{}, byMentor:{}, byStage:{}
  };
  all.forEach(function(r) {
    if (r.approvalStatus==='Pending')             stats.pending++;
    if (r.approvalStatus==='Approved')            stats.approved++;
    if (r.approvalStatus==='Rejected')            stats.rejected++;
    if (r.approvalStatus==='Needs Clarification') stats.needsClarity++;
    if (r.status==='Non-Responsive' && r.lastResponseDate) {
      try { if ((today-new Date(r.lastResponseDate))/86400000 >= 7) stats.nonResponsiveFlag++; }
      catch(e){}
    }
    stats.byStatus[r.status]       = (stats.byStatus[r.status]||0)+1;
    stats.byMentor[r.mentorName]   = (stats.byMentor[r.mentorName]||0)+1;
    stats.byStage[r.trainingStage] = (stats.byStage[r.trainingStage]||0)+1;
  });
  return stats;
}

// ── ROW COLORING ──────────────────────────────────────────────
function colorRow(sheet, rowNum, approvalStatus) {
  var colors = {
    'Pending':'#FFF8E7','Approved':'#E8F5E9',
    'Rejected':'#FFEBEE','Needs Clarification':'#E3F2FD'
  };
  var lastCol = sheet.getLastColumn() || 31;
  sheet.getRange(rowNum,1,1,lastCol).setBackground(colors[approvalStatus]||'#FFFFFF');
}
