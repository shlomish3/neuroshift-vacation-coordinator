// ── Submission lock & archive helpers ────────────────────────────────────────
function getSubmissionStatus() {
  const props = PropertiesService.getScriptProperties();
  const today = new Date();
  const months = [];
  for (let i = 1; i <= 3; i++) {
    let m = today.getMonth() + 1 + i;
    let y = today.getFullYear();
    if (m > 12) { m -= 12; y++; }
    const key = `${m}-${y}`;
    months.push({
      month: m,
      year: y,
      key: key,
      locked: props.getProperty(`SUBMISSIONS_LOCKED_${key}`) === 'true',
      lockedAt: props.getProperty(`SUBMISSIONS_LOCKED_AT_${key}`) || ''
    });
  }
  return { months: months };
}

function setSubmissionLock(monthKey, lock) {
  if (!monthKey) return { success: false };
  const props = PropertiesService.getScriptProperties();
  if (lock) {
    props.setProperty(`SUBMISSIONS_LOCKED_${monthKey}`, 'true');
    props.setProperty(`SUBMISSIONS_LOCKED_AT_${monthKey}`, new Date().toISOString());
  } else {
    props.deleteProperty(`SUBMISSIONS_LOCKED_${monthKey}`);
    props.deleteProperty(`SUBMISSIONS_LOCKED_AT_${monthKey}`);
  }
  return { success: true };
}

function archiveMonth() {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName('תגובות לטופס 1');
    if (!src) return { success: false, message: 'Sheet not found' };

    const hebrewMonths = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

    // Use ROSTER_MONTH as authoritative source for archive name
    const rm        = getRosterMonth();
    const archiveM  = rm.month;
    const archiveY  = rm.year;
    const archiveName = hebrewMonths[archiveM - 1] + ' ' + String(archiveY).slice(2);

    // Copy sheet, rename, move to end
    const lastRow = src.getLastRow();
    const copy = src.copyTo(ss);
    copy.setName(archiveName);
    ss.setActiveSheet(copy);
    ss.moveActiveSheet(ss.getNumSheets());

    // Clear data rows in source (keep header)
    if (lastRow > 1) src.deleteRows(2, lastRow - 1);

    // Advance ROSTER_MONTH to next month, then repopulate
    let newM = archiveM + 1;
    let newY = archiveY;
    if (newM > 12) { newM -= 12; newY++; }
    const newKey = newM + '-' + newY;
    PropertiesService.getScriptProperties().setProperty('ROSTER_MONTH', newKey);
    repopulateRosterSheet(newKey);

    return { success: true, sheetName: archiveName, newRosterMonth: newKey };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ── Roster Month management ────────────────────────────────────────────────

function getRosterMonth() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('ROSTER_MONTH');
  const today  = new Date();
  const todayM = today.getMonth() + 1;
  const todayY = today.getFullYear();
  let m, y;

  if (stored && stored.indexOf('-') > -1) {
    const parts = stored.split('-');
    m = parseInt(parts[0], 10);
    y = parseInt(parts[1], 10);
    // Auto-advance: if roster month <= current month, advance to next month
    if (!isNaN(m) && !isNaN(y) && (y < todayY || (y === todayY && m <= todayM))) {
      m = todayM + 1;
      y = todayY;
      if (m > 12) { m -= 12; y++; }
      props.setProperty('ROSTER_MONTH', m + '-' + y);
    }
  } else {
    m = todayM + 1;
    y = todayY;
    if (m > 12) { m -= 12; y++; }
  }
  return { key: m + '-' + y, month: m, year: y };
}

function setRosterMonth(monthKey) {
  if (!monthKey || monthKey.indexOf('-') === -1) return { success: false };
  PropertiesService.getScriptProperties().setProperty('ROSTER_MONTH', monthKey);
  return { success: true, key: monthKey };
}

function shiftRosterMonth(delta) {
  const rm = getRosterMonth();
  let m = rm.month + delta;
  let y = rm.year;
  if (m > 12) { m -= 12; y++; }
  if (m < 1)  { m += 12; y--; }
  const newKey = m + '-' + y;
  PropertiesService.getScriptProperties().setProperty('ROSTER_MONTH', newKey);
  repopulateRosterSheet(newKey);
  return { success: true, key: newKey, month: m, year: y };
}

function repopulateRosterSheet(monthKey) {
  try {
    const parts  = monthKey.split('-');
    const tgtM   = parseInt(parts[0], 10);
    const tgtY   = parseInt(parts[1], 10);
    const ss     = SpreadsheetApp.getActiveSpreadsheet();

    // Ensure legacy sheet exists (create with headers if missing)
    let legacySheet = ss.getSheetByName('\u05EA\u05D2\u05D5\u05D1\u05D5\u05EA \u05DC\u05D8\u05D5\u05E4\u05E1 1');
    if (!legacySheet) {
      legacySheet = ss.insertSheet('\u05EA\u05D2\u05D5\u05D1\u05D5\u05EA \u05DC\u05D8\u05D5\u05E4\u05E1 1');
      legacySheet.appendRow(['חותמת זמן', 'כתובת אימייל', 'שם', 'לילות בהם אינכם יכולים לעשות תורנות/כוננות', 'בחירת ימי חופש - כולל ימי שישי', 'התייחסות חופשית + תאריכים מועדפים', 'תאריכים מועדפים']);
      legacySheet.getRange('A1:G1').setFontWeight('bold');
    }
    // Clear data rows (keep header)
    const lastLeg = legacySheet.getLastRow();
    if (lastLeg > 1) legacySheet.deleteRows(2, lastLeg - 1);

    // Read Vacation_Requests_Clean
    const cleanSheet = ss.getSheetByName('Vacation_Requests_Clean');
    if (!cleanSheet) return;
    const cleanData = cleanSheet.getDataRange().getValues();
    if (cleanData.length < 2) return;

    // Exact type strings matching appendDates() in processForm
    const TYPE_CANNOT   = '\u05DC\u05D0 \u05D9\u05DB\u05D5\u05DC \u05EA\u05D5\u05E8\u05E0\u05D5\u05EA'; // לא יכול תורנות
    const TYPE_VACATION = '\u05D7\u05D5\u05E4\u05E9';                           // חופש
    const TYPE_PREF     = '\u05DE\u05D5\u05E2\u05D3\u05E3';                     // מועדף
    const TYPE_PREF_IMP = '\u05DE\u05D5\u05E2\u05D3\u05E3 (\u05D7\u05E9\u05D5\u05D1)'; // מועדף (חשוב)

    // Group by email: split dates by type to match legacy columns
    const map = {}; // email -> { name, cannot[], vacation[], preferred[], reason, ts, editId }
    cleanData.slice(1).forEach(function(row) {
      const ts     = row[0];
      const name   = String(row[2]).trim();
      const email  = String(row[3]).trim().toLowerCase();
      const dtVal  = row[4];
      const dtType = String(row[5]).trim();
      const reason = String(row[6]).trim();
      const editId = String(row[7]).trim();
      const status = String(row[8]).trim();

      if (!email || email === 'email') return;
      if (status === 'Revoked' || status === 'Deleted') return;

      let d;
      if (dtVal instanceof Date) {
        d = dtVal;
      } else {
        const s = String(dtVal).trim();
        const sp = s.split('/');
        if (sp.length === 3) {
          d = new Date(parseInt(sp[2]), parseInt(sp[1]) - 1, parseInt(sp[0]));
        } else {
          d = new Date(s);
        }
      }
      if (!d || isNaN(d.getTime())) return;
      if ((d.getMonth() + 1) !== tgtM || d.getFullYear() !== tgtY) return;

      const dd  = String(d.getDate()).padStart(2, '0');
      const mm2 = String(d.getMonth() + 1).padStart(2, '0');
      const ds  = dd + '/' + mm2 + '/' + d.getFullYear();

      if (!map[email]) map[email] = { name: name, cannot: [], vacation: [], preferred: [], reason: '', ts: ts, editId: editId };
      if (reason && !map[email].reason) map[email].reason = reason;

      if (dtType === TYPE_CANNOT) {
        map[email].cannot.push(ds);
      } else if (dtType === TYPE_VACATION) {
        map[email].vacation.push(ds);
      } else if (dtType === TYPE_PREF || dtType === TYPE_PREF_IMP) {
        map[email].preferred.push(ds);
      } else {
        // Fallback: unknown types go to preferred (safer than cannot)
        map[email].preferred.push(ds);
      }
    });

    const sortDates = function(arr) {
      return arr.sort(function(a, b) {
        const pa = a.split('/'); const pb = b.split('/');
        return new Date(pa[2], pa[1]-1, pa[0]) - new Date(pb[2], pb[1]-1, pb[0]);
      });
    };
    const rows = [];
    Object.keys(map).forEach(function(email) {
      const d2  = map[email];
      const row = new Array(26).fill('');
      row[0]  = d2.ts;
      row[1]  = email;
      row[2]  = d2.name;
      row[3]  = sortDates(d2.cannot).join(', ');    // col D: cannot-work dates
      row[4]  = sortDates(d2.vacation).join(', ');  // col E: vacation dates
      row[5]  = d2.reason;                           // col F: free text
      row[6]  = sortDates(d2.preferred).join(', '); // col G: preferred dates
      row[7]  = monthKey;
      row[25] = d2.editId;
      rows.push(row);
    });

    if (rows.length > 0) {
      legacySheet.getRange(2, 1, rows.length, 26).setValues(rows);
    }
  } catch(e) {
    Logger.log('repopulateRosterSheet error: ' + e.message);
  }
}

function getEmailWorkerList() {
  const db = getTeamDatabase();
  const result = [];
  const seen = {};
  Object.keys(db).forEach(function(email) {
    if (!email || email === 'email') return;
    if (db[email].primaryEmail) return; // skip alias
    if (seen[email]) return;
    seen[email] = true;
    result.push({ name: db[email].name || '', email: email, role: db[email].role || '' });
  });
  return result.sort(function(a, b) { return a.name.localeCompare(b.name, 'he'); });
}

function sendEmailBlast(params) {
  const recipients = params.recipients || [];
  const subject    = params.subject    || '';
  const body       = params.body       || '';
  if (!subject || !body || recipients.length === 0) return { success: false };

  const htmlBody = '<div dir="rtl" style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.7;">' +
    body.replace(/\n/g, '<br>') + '</div>';

  let sent = 0, failed = 0, failedList = [];
  recipients.forEach(function(r) {
    try {
      GmailApp.sendEmail(r.email, subject, body, {
        name:      'Neuro Shift',
        from:      'shamir.neuroshift@gmail.com',
        replyTo:   'shamir.neuroshift@gmail.com',
        htmlBody:  htmlBody
      });
      sent++;
    } catch(e) {
      try {
        MailApp.sendEmail({ to: r.email, subject: subject, body: body, htmlBody: htmlBody });
        sent++;
      } catch(e2) {
        failed++;
        failedList.push(r.email);
      }
    }
  });
  return { success: true, sent: sent, failed: failed, failedList: failedList };
}


function processForm(formData) {
  try {
    const rawEmail = formData.email || "";
    const email = rawEmail.toLowerCase().trim();

    // ── Submission lock check (per-month) ─────────────────────────────────────
    // (Month scope is parsed next, so we do a quick pre-check here)
    const _preMonth = parseInt(formData.targetMonth, 10) || (new Date().getMonth() + 2);
    const _preYear  = parseInt(formData.targetYear,  10) || new Date().getFullYear();
    const _preKey   = `${_preMonth}-${_preYear}`;
    const _props    = PropertiesService.getScriptProperties();
    if (_props.getProperty(`SUBMISSIONS_LOCKED_${_preKey}`) === 'true') {
      return { status: 'Locked', message: 'הגשות לחודש זה נעולות כרגע. פנה למנהל לפרטים.' };
    }

    const db = getTeamDatabase();
    if (!db[email]) {
      return "משתמש לא זוהה. אנא הזן מייל מוכר או פנה למנהל המערכת.";
    }

    const name = db[email].name;
    const role = db[email].role;
    const freeText = formData.freeText || "";

    // ── Month scope ──────────────────────────────────────────────────────────
    const targetMonth = parseInt(formData.targetMonth, 10) || (new Date().getMonth() + 2); // default: next month
    const targetYear  = parseInt(formData.targetYear,  10) || new Date().getFullYear();
    const firstDay = new Date(targetYear, targetMonth - 1, 1).getTime();
    const lastDay  = new Date(targetYear, targetMonth,     0, 23, 59, 59).getTime();
    const monthKey = `${targetMonth}-${targetYear}`; // used in legacy sheet col H

    const sheetName = 'Vacation_Requests_Clean';
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.appendRow(['Timestamp', 'Role', 'Name', 'Email', 'Date', 'Date Type', 'Free Text', 'EditID', 'Status']);
      sheet.getRange("A1:I1").setFontWeight("bold");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    // Preserve manager statuses for this month before deleting
    const existingStatuses = {};
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][3]).trim().toLowerCase() !== email) continue;
      const odStr = formatDateValue(values[i][4]);
      if (!odStr) continue;
      const parts = odStr.split('/');
      if (parts.length === 3) {
        const dT = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0])).getTime();
        if (dT >= firstDay && dT <= lastDay) {
          const oldType   = String(values[i][5]).trim();
          const oldStatus = String(values[i][8]).trim();
          if (oldStatus) existingStatuses[`${odStr}_${oldType}`] = oldStatus;
        }
      }
    }

    // Delete all existing rows for this email in the target month
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][3]).toLowerCase().trim() !== email) continue;
      const dStr = formatDateValue(values[i][4]);
      if (!dStr) continue;
      const parts = dStr.split('/');
      if (parts.length === 3) {
        const dT = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0])).getTime();
        if (dT >= firstDay && dT <= lastDay) sheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();

    const editId = Utilities.getUuid(); // fresh editId per month-submission
    const timestamp = new Date();

    function appendDates(datesString, dateType) {
      if (!datesString) return;
      const datesArray = datesString.split(', ');
      for (let i = 0; i < datesArray.length; i++) {
        let dateValue = datesArray[i].trim();
        if (dateValue) {
          const finalStatus = existingStatuses[`${dateValue}_${dateType}`] || "Approved";
          sheet.appendRow([timestamp, role, name, email, dateValue, dateType, freeText, editId, finalStatus]);
        }
      }
    }
    
    let hasData = false;
    
    if (formData.datesCannot) { appendDates(formData.datesCannot, "לא יכול תורנות"); hasData = true; }
    if (formData.datesVacation) { appendDates(formData.datesVacation, "חופש"); hasData = true; }
    
    const rawImportant = formData.datesImportant || "";
    let prefArr = formData.datesPreferred ? formData.datesPreferred.split(', ') : [];
    let impArr = rawImportant ? rawImportant.split(', ') : [];
    
    let regularPref = prefArr.filter(d => !impArr.includes(d));
    
    if (regularPref.length > 0) { appendDates(regularPref.join(', '), "מועדף"); hasData = true; }
    if (impArr.length > 0) { appendDates(impArr.join(', '), "מועדף (חשוב)"); hasData = true; }
    
    if (!hasData && freeText) {
      sheet.appendRow([timestamp, role, name, email, "", "טקסט חופשי", freeText, editId, "Approved"]);
    }
    
    // --- LEGACY DUAL-WRITE LOGIC (active ROSTER_MONTH only) ---
    if (monthKey === getRosterMonth().key) {
  // --- LEGACY DUAL-WRITE LOGIC (תגובות לטופס 1) ---
      const legacySheetName = "תגובות לטופס 1";
      let legacySheet = spreadsheet.getSheetByName(legacySheetName);
      if (!legacySheet) {
        legacySheet = spreadsheet.insertSheet(legacySheetName);
        legacySheet.appendRow(['חותמת זמן', 'כתובת אימייל', 'שם', 'לילות בהם אינכם יכולים לעשות תורנות/כוננות', 'בחירת ימי חופש - כולל ימי שישי', 'התייחסות חופשית + תאריכים מועדפים', 'תאריכים מועדפים']);
        legacySheet.getRange("A1:G1").setFontWeight("bold");
      }
  
      // Delete previous row for this email + this month (col H = monthKey)
      const legacyVals = legacySheet.getDataRange().getValues();
      for (let i = legacyVals.length - 1; i >= 1; i--) {
        if (String(legacyVals[i][1]).trim().toLowerCase() === email &&
            String(legacyVals[i][7]).trim() === monthKey) {
          legacySheet.deleteRow(i + 1);
        }
      }
      SpreadsheetApp.flush();
  
      const legacyRow = new Array(26).fill("");
      legacyRow[0] = timestamp;
      legacyRow[1] = email;
      legacyRow[2] = name;
      legacyRow[3] = formatLegacyDates(formData.datesCannot);
      legacyRow[4] = formatLegacyDates(formData.datesVacation);
      legacyRow[5] = freeText;
      legacyRow[7] = monthKey; // col H: month-year key for future re-submission lookup
      
      let legacyPrefFormatted = [];
      if (regularPref.length > 0) legacyPrefFormatted.push(formatLegacyDates(regularPref.join(', '), false));
      if (impArr.length > 0) legacyPrefFormatted.push(formatLegacyDates(impArr.join(', '), true));
      
      legacyRow[6] = legacyPrefFormatted.filter(Boolean).join(', ');
      legacyRow[25] = editId; // Col Z
      
      legacySheet.appendRow(legacyRow);
      // --------------------------------
    } // end ROSTER_MONTH gate
    
    // Send email receipt
    try {
      const SITE_URL = 'https://sites.google.com/view/shamir-neuroshift/';
      const isEditReceipt = Object.keys(existingStatuses).length > 0;
      const mLabel = monthKey
        ? (function() {
            const mp = monthKey.split('-');
            const mN = ['\u05D9\u05E0\u05D5\u05D0\u05E8','\u05E4\u05D1\u05E8\u05D5\u05D0\u05E8','\u05DE\u05E8\u05E5','\u05D0\u05E4\u05E8\u05D9\u05DC','\u05DE\u05D0\u05D9','\u05D9\u05D5\u05E0\u05D9','\u05D9\u05D5\u05DC\u05D9','\u05D0\u05D5\u05D2\u05D5\u05E1\u05D8','\u05E1\u05E4\u05D8\u05DE\u05D1\u05E8','\u05D0\u05D5\u05E7\u05D8\u05D5\u05D1\u05E8','\u05E0\u05D5\u05D1\u05DE\u05D1\u05E8','\u05D3\u05E6\u05DE\u05D1\u05E8'];
            return (mN[parseInt(mp[0],10)-1] || mp[0]) + ' \'' + (mp[1]||'').slice(2);
          })()
        : '';

      const subject = (isEditReceipt ? '[\u05E2\u05D3\u05DB\u05D5\u05DF] ' : '') + '\u05D0\u05D9\u05E9\u05D5\u05E8 \u05D1\u05E7\u05E9\u05D4 \u05E0\u05D9\u05D5\u05E8\u05D5\u05DC\u05D5\u05D2\u05D9\u05D4 \u2014 ' + mLabel;

      let wb = "<div dir='rtl' style='font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;'>";
      wb += "<h3 style='color:#1a5276;margin-bottom:4px;'>" + (isEditReceipt ? '\u270F\uFE0F \u05E2\u05D3\u05DB\u05D5\u05DF \u05D1\u05E7\u05E9\u05D4' : '\u2705 \u05D1\u05E7\u05E9\u05D4 \u05D4\u05EA\u05E7\u05D1\u05DC\u05D4 \u05D1\u05D4\u05E6\u05DC\u05D7\u05D4') + "</h3>";
      wb += "<p style='margin-top:4px;'>\u05E9\u05DC\u05D5\u05DD " + name + ", \u05D1\u05E7\u05E9\u05EA\u05DA \u05DC\u05D7\u05D5\u05D3\u05E9 <strong>" + mLabel + "</strong> \u05E0\u05E7\u05DC\u05D8\u05D4 \u05D1\u05DE\u05E2\u05E8\u05DB\u05EA.</p>";

      wb += "<table style='border-collapse:collapse;width:100%;margin:12px 0;'>";
      wb += "<tr><td style='padding:5px 12px;font-weight:bold;width:180px;background:#eaf2fb;'>\u05D7\u05D5\u05D3\u05E9</td><td style='padding:5px 12px;background:#eaf2fb;'>" + mLabel + "</td></tr>";
      if (formData.datesCannot)   wb += "<tr><td style='padding:5px 12px;font-weight:bold;background:#fdf2f8;'>\u05DC\u05D0 \u05D9\u05DB\u05D5\u05DC \u05EA\u05D5\u05E8\u05E0\u05D5\u05EA</td><td style='padding:5px 12px;background:#fdf2f8;'>" + formData.datesCannot + "</td></tr>";
      if (formData.datesVacation) wb += "<tr><td style='padding:5px 12px;font-weight:bold;background:#fefefe;'>\u05D9\u05DE\u05D9 \u05D7\u05D5\u05E4\u05E9</td><td style='padding:5px 12px;background:#fefefe;'>" + formData.datesVacation + "</td></tr>";
      if (formData.datesPreferred) {
        const dispP = prefArr.map(function(d){ return impArr.includes(d) ? d + ' (\u05D7\u05E9\u05D5\u05D1)' : d; });
        wb += "<tr><td style='padding:5px 12px;font-weight:bold;background:#f0fdf4;'>\u05DE\u05D5\u05E2\u05D3\u05E4\u05D9\u05DD</td><td style='padding:5px 12px;background:#f0fdf4;'>" + dispP.join(', ') + "</td></tr>";
      }
      if (formData.freeText) wb += "<tr><td style='padding:5px 12px;font-weight:bold;background:#fffde7;'>\u05D4\u05E2\u05E8\u05D5\u05EA</td><td style='padding:5px 12px;background:#fffde7;'>" + formData.freeText + "</td></tr>";
      wb += "</table>";

      wb += "<p style='margin-top:14px;'>\u05E0\u05D9\u05EA\u05DF \u05DC\u05E2\u05D3\u05DB\u05DF \u05D0\u05D5 \u05DC\u05E9\u05E0\u05D5\u05EA \u05D0\u05EA \u05D4\u05D1\u05E7\u05E9\u05D4 \u05D1\u05DB\u05DC \u05E2\u05EA \u05D3\u05E8\u05DA \u05D0\u05EA\u05E8 \u05D4\u05DE\u05E2\u05E8\u05DB\u05EA:<br>";
      wb += "<a href='" + SITE_URL + "' style='color:#1a5276;'>" + SITE_URL + "</a></p>";

      wb += "<p style='margin-top:18px;color:#555;'>\u05D1\u05D1\u05E8\u05DB\u05D4,<br><strong>\u05E6\u05D5\u05D5\u05EA \u05E0\u05D9\u05D4\u05D5\u05DC \u05EA\u05D5\u05E8\u05E0\u05D5\u05D9\u05D5\u05EA \u05E0\u05D9\u05D5\u05E8\u05D5\u05DC\u05D5\u05D2\u05D9\u05D4</strong></p>";
      wb += "</div>";

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: wb,
        name: 'Neuro Shift'
      });

        // ── Manager notification ──────────────────────────────────────────
        const isEdit = Object.keys(existingStatuses).length > 0;
        const monthLabel = monthKey
          ? (function() {
              const mp = monthKey.split('-');
              const mNames = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
              return (mNames[parseInt(mp[0],10)-1] || mp[0]) + ' \'' + (mp[1]||'').slice(2);
            })()
          : '';
        const mgrSubject = (isEdit ? '[עדכון] ' : '[חדש] ') + name + ' — ' + monthLabel;
        let mgrBody = "<div dir='rtl' style='font-family:Arial,sans-serif;font-size:14px;color:#222;'>";
        mgrBody += "<h3 style='color:#1a5276;'>" + (isEdit ? '✏️ עדכון בקשה' : '🆕 בקשה חדשה') + "</h3>";
        mgrBody += "<table style='border-collapse:collapse;width:100%;'>";
        mgrBody += "<tr><td style='padding:4px 10px;font-weight:bold;width:160px;'>עובד</td><td style='padding:4px 10px;'>" + name + "</td></tr>";
        mgrBody += "<tr style='background:#f2f6fc;'><td style='padding:4px 10px;font-weight:bold;'>מייל</td><td style='padding:4px 10px;direction:ltr;'>" + email + "</td></tr>";
        mgrBody += "<tr><td style='padding:4px 10px;font-weight:bold;'>חודש</td><td style='padding:4px 10px;'>" + monthLabel + "</td></tr>";
        if (formData.datesCannot)   mgrBody += "<tr style='background:#fdf2f8;'><td style='padding:4px 10px;font-weight:bold;'>לא יכול תורנות</td><td style='padding:4px 10px;'>" + formData.datesCannot + "</td></tr>";
        if (formData.datesVacation) mgrBody += "<tr><td style='padding:4px 10px;font-weight:bold;'>ימי חופש</td><td style='padding:4px 10px;'>" + formData.datesVacation + "</td></tr>";
        if (formData.datesPreferred) {
          const dispPref = prefArr.map(function(d){ return impArr.includes(d) ? d + ' (חשוב)' : d; });
          mgrBody += "<tr style='background:#f0fdf4;'><td style='padding:4px 10px;font-weight:bold;'>מועדפים</td><td style='padding:4px 10px;'>" + dispPref.join(', ') + "</td></tr>";
        }
        if (formData.freeText) mgrBody += "<tr><td style='padding:4px 10px;font-weight:bold;'>הערות</td><td style='padding:4px 10px;'>" + formData.freeText + "</td></tr>";
        mgrBody += "</table>";
        mgrBody += "<p style='margin-top:14px;font-size:12px;color:#666;'>נשלח אוטומטית ע\"י מערכת Neuro Shift</p>";
        mgrBody += "</div>";

        MailApp.sendEmail({
          to: 'shamir.neuroshift@gmail.com',
          subject: mgrSubject,
          htmlBody: mgrBody,
          name: 'Neuro Shift'
        });
    } catch(emailErr) {
      console.error("Failed to send email receipt:", emailErr);
    }
    
    return { status: "Success", editId: editId };
    
  } catch (e) {
    return "שגיאה: " + e.message;
  }
}

function getPreviousData(editId) {
  if (!editId) return null;
  const sheetName = 'Vacation_Requests_Clean'; 
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return null;
  
  const values = sheet.getDataRange().getValues();
  const result = {
    email: "",
    datesCannot: [],
    datesVacation: [],
    datesPreferred: [],
    datesImportant: [],
    freeText: ""
  };
  
  for (let i = 1; i < values.length; i++) {
    // Check if EditID in Column H (index 7) matches
      if (values[i][7] === editId) {
      result.email = values[i][3];
      const dateType = values[i][5];
      const dateVal = formatDateValue(values[i][4]);
      const status = String(values[i][8] || "Approved").trim();
      
      if (status === 'Revoked') continue;
      
      if (dateVal) {
        if (dateType === "לא יכול תורנות") result.datesCannot.push(dateVal);
        else if (dateType === "חופש") result.datesVacation.push(dateVal);
        else if (dateType === "מועדף") result.datesPreferred.push(dateVal);
        else if (dateType === "מועדף (חשוב)") {
          result.datesPreferred.push(dateVal);
          result.datesImportant.push(dateVal);
        }
      }
      if (values[i][6]) {
        result.freeText = values[i][6];
      }
    }
  }
  
  if (!result.email) return null;
  
  result.datesCannot = result.datesCannot.join(', ');
  result.datesVacation = result.datesVacation.join(', ');
  result.datesPreferred = result.datesPreferred.join(', ');
  result.datesImportant = result.datesImportant.join(', ');
  
  return result;
}

function getPreviousDataByEmail(email, targetMonth, targetYear) {
  if (!email) return null;
  email = String(email).trim().toLowerCase();

  // Default to next month if no month/year provided
  const today = new Date();
  if (!targetMonth || !targetYear) {
    targetMonth = today.getMonth() + 2;
    targetYear  = today.getFullYear();
    if (targetMonth > 12) { targetMonth = 1; targetYear++; }
  }
  targetMonth = parseInt(targetMonth, 10);
  targetYear  = parseInt(targetYear,  10);

  const firstDay = new Date(targetYear, targetMonth - 1, 1).getTime();
  const lastDay  = new Date(targetYear, targetMonth,     0, 23, 59, 59).getTime();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Vacation_Requests_Clean');
  if (!sheet) return null;

  const values = sheet.getDataRange().getValues();

  const result = {
    email: email,
    datesCannot:   [],
    datesVacation: [],
    datesPreferred:[],
    datesImportant:[],
    freeText: ''
  };

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][3]).trim().toLowerCase() !== email) continue;
    const dateVal = formatDateValue(values[i][4]);
    if (!dateVal) continue;
    const parts = dateVal.split('/');
    if (parts.length !== 3) continue;
    const dT = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0])).getTime();
    if (dT < firstDay || dT > lastDay) continue;

    const dateType = String(values[i][5]).trim();
    const status   = String(values[i][8] || 'Approved').trim();
    if (status === 'Revoked') continue;

    if      (dateType === 'לא יכול תורנות') result.datesCannot.push(dateVal);
    else if (dateType === 'חופש')           result.datesVacation.push(dateVal);
    else if (dateType === 'מועדף')          result.datesPreferred.push(dateVal);
    else if (dateType === 'מועדף (חשוב)') {
      result.datesPreferred.push(dateVal);
      result.datesImportant.push(dateVal);
    }
    if (values[i][6]) result.freeText = values[i][6];
  }

  result.datesCannot    = result.datesCannot.join(', ');
  result.datesVacation  = result.datesVacation.join(', ');
  result.datesPreferred = result.datesPreferred.join(', ');
  result.datesImportant = result.datesImportant.join(', ');

  return result;
}

function sendBugReport(text, email) {
  MailApp.sendEmail({
    to: "shamir.neuroshift@gmail.com",
    subject: "Bug Report from Neuro Shift Dashboard",
    body: "Bug Report received:\n\n" + text + "\n\nFrom user: " + email
  });
}

function testGetVacations() {
  Logger.log(JSON.stringify(getExistingVacations()));
}

function toggleVacationStatus(email, dateStr, newStatus, reqType, reason) {
  if (!email || !dateStr || !reqType) return { success: false, message: "Missing email, date, or type." };
  email = String(email).trim().toLowerCase();
  dateStr = String(dateStr).trim();
  
  const sheetName = 'Vacation_Requests_Clean';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "No database found." };
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][3]).toLowerCase().trim();
    const rowDate = formatDateValue(values[i][4]);
    const rowType = String(values[i][5]).trim();
    
    if (rowEmail === email && rowDate === dateStr && rowType === reqType) {
      sheet.getRange(i + 1, 9).setValue(newStatus);
      
      if (newStatus === 'Revoked') {
        const db = getTeamDatabase();
        if (db[email]) {
          const name = db[email].name;
          const subject = "עדכון בנוגע לבקשת אילוצים / חופשים";
          let htmlBody = "<div dir='rtl' style='font-family: Arial, sans-serif; font-size: 14px;'>";
          htmlBody += "<p>שלום " + name + ",</p>";
          htmlBody += "<p>מנהל המערכת <b>דחה</b> את בקשתך (" + reqType + ") לתאריך <b>" + dateStr + "</b>.</p>";
          if (reason) {
            htmlBody += "<p><b>הודעת המנהל:</b> " + reason + "</p>";
          }
          htmlBody += "<p>בברכה,<br>צוות ניהול תורנויות נוירולוגיה</p>";
          htmlBody += "</div>";
          
          try {
            MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
            if (db[email].email2) MailApp.sendEmail({ to: db[email].email2, subject: subject, htmlBody: htmlBody });
          } catch(e) {}
        }
      }
      return { success: true };
    }
  }
  return { success: false, message: "Request not found in database." };
}
// Trigger update V109

// Trigger update V112


function bulkManageRequests(email, dates, action, reason) {
  if (!email || !dates || dates.length === 0) return { success: false, message: "Missing email or dates." };
  email = String(email).trim().toLowerCase();
  
  const sheetName = 'Vacation_Requests_Clean';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "No database found." };
  
  const db = getTeamDatabase();
  const worker = db[email];
  if (!worker) return { success: false, message: "Worker not found in database." };
  
  const lastRow = sheet.getLastRow();
  const values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 9).getValues() : [];
  
  const timestamp = new Date();
  const editId = Utilities.getUuid();

  // ── Fast bulk-write approach ──────────────────────────────────────────────
  // Instead of calling setValue / appendRow once per date (O(3N) API calls),
  // we: 1) scan the entire data in memory, 2) collect all changes, then
  // 3) flush with two bulk setValues calls regardless of N.

  const dateSet = new Set(dates);
  const datesFound = new Set(); // dates that already have a חופש row

  // Snapshot the status and reason columns so we can bulk-write them back
  // (index 0 = header row so indices align with 1-based sheet rows)
  const statusCol = values.map(r => [r[8] !== undefined ? r[8] : '']); // col 9
  const reasonCol = values.map(r => [r[6] !== undefined ? r[6] : '']); // col 7
  let anyStatusChanged = false;
  let anyReasonChanged = false;

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][3]).toLowerCase().trim();
    const rowDate  = formatDateValue(values[i][4]);
    const rowType  = String(values[i][5]).trim();
    if (rowEmail === email && dateSet.has(rowDate)) {
      // ForcedLeave / Revoked only touch vacation (חופש) rows – never cannot-work
      if ((action === 'ForcedLeave' || action === 'Revoked') && rowType !== 'חופש') continue;
      statusCol[i][0] = action;
      anyStatusChanged = true;
      datesFound.add(rowDate);
      if (reason) { reasonCol[i][0] = reason; anyReasonChanged = true; }
    }
  }

  // Flush status / reason updates in two single calls (skip if nothing changed)
  if (anyStatusChanged && values.length > 1) {
    sheet.getRange(1, 9, values.length, 1).setValues(statusCol);
  }
  if (anyReasonChanged && values.length > 1) {
    sheet.getRange(1, 7, values.length, 1).setValues(reasonCol);
  }

  const processedDates = Array.from(datesFound);

  // ── Batch-append new rows (dates with no existing חופש row) ────────────────
  const newDates = [];
  if (action !== 'Revoked') {
    const newRowsData = [];
    dates.forEach(dateStr => {
      if (!datesFound.has(dateStr)) {
        newRowsData.push([timestamp, worker.role, worker.name, email, dateStr, 'חופש', reason || '', editId, action]);
        newDates.push(dateStr);
        processedDates.push(dateStr);
      }
    });
    if (newRowsData.length > 0) {
      // ONE setValues call for all new rows — vastly faster than N appendRow calls
      const appendAt = sheet.getLastRow() + 1;
      sheet.getRange(appendAt, 1, newRowsData.length, 9).setValues(newRowsData);
    }
  }
  // ─────────────────────────────────────────────────────────────────────────

  if (newDates.length > 0) {
    const legacySheetName = "תגובות לטופס 1";
    let legacySheet = spreadsheet.getSheetByName(legacySheetName);
    if (!legacySheet) {
      legacySheet = spreadsheet.insertSheet(legacySheetName);
      legacySheet.appendRow(['חותמת זמן', 'כתובת אימייל', 'שם', 'לילות בהם אינכם יכולים לעשות תורנות/כוננות', 'בחירת ימי חופש - כולל ימי שישי', 'התייחסות חופשית + תאריכים מועדפים', 'תאריכים מועדפים']);
      legacySheet.getRange("A1:G1").setFontWeight("bold");
    }
    legacySheet.appendRow([
      timestamp,
      email,
      worker.name,
      "", 
      newDates.join(', '), 
      reason || "", 
      ""  
    ]);
  }

  
  if (action === 'Revoked' && processedDates.length > 0) {
    const subject = "עדכון בנוגע לבקשת אילוצים / חופשים";
    let htmlBody = "<div dir='rtl' style='font-family: Arial, sans-serif; font-size: 14px;'>";
    htmlBody += "<p>שלום " + worker.name + ",</p>";
    htmlBody += "<p>מנהל המערכת <b>דחה</b> את בקשותיך לתאריכים הבאים:</p>";
    htmlBody += "<p><b>" + processedDates.join(', ') + "</b></p>";
    if (reason) {
      htmlBody += "<p><b>הודעת המנהל:</b> " + reason + "</p>";
    }
    htmlBody += "<p>בברכה,<br>צוות ניהול תורנויות נוירולוגיה</p>";
    htmlBody += "</div>";
    
    try {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
      if (worker.email2) MailApp.sendEmail({ to: worker.email2, subject: subject, htmlBody: htmlBody });
    } catch(e) {}
  }
  
  return { success: true };
}

// ── God Mode User Management ──────────────────────────────────────────────────

const SUPER_ADMIN        = 'shlomi.shmuel3@gmail.com';
const DEFAULT_GOD_MODE   = 'shlomi.shmuel3@gmail.com,shlomip@shamir.gov.il';

function getGodModeUsers() {
  const stored = PropertiesService.getScriptProperties().getProperty('GOD_MODE_USERS');
  const raw    = stored || DEFAULT_GOD_MODE;
  return raw.split(',').map(function(e) { return e.trim().toLowerCase(); }).filter(Boolean);
}

function setGodModeUsers(emailList, callerEmail) {
  try {
    const caller = (callerEmail || '').toLowerCase().trim();
    if (caller !== SUPER_ADMIN) return { success: false, message: 'No permission' };
    let list = (emailList || []).map(function(e){ return e.trim().toLowerCase(); }).filter(function(e){ return e.includes('@'); });
    if (!list.includes(SUPER_ADMIN)) list.unshift(SUPER_ADMIN);
    PropertiesService.getScriptProperties().setProperty('GOD_MODE_USERS', list.join(','));
    return { success: true, users: list };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
// ── God Mode Login Notification ───────────────────────────────────────────────

function logGodModeLogin(managerEmail) {
  try {
    const now       = new Date();
    const pad       = function(n){ return String(n).padStart(2,'0'); };
    const timeStr   = pad(now.getDate()) + '/' + pad(now.getMonth()+1) + '/' + now.getFullYear() +
                      ' ' + pad(now.getHours()) + ':' + pad(now.getMinutes());
    const subject   = '\uD83D\uDD10 \u05DB\u05E0\u05D9\u05E1\u05D4 \u05DC\u05DE\u05E6\u05D1 \u05DE\u05E0\u05D4\u05DC: ' + (managerEmail || 'unknown');
    const body      = "<div dir='rtl' style='font-family:Arial,sans-serif;font-size:14px;color:#222;'>" +
      "<h3 style='color:#6c3483;'>\uD83D\uDD10 \u05DB\u05E0\u05D9\u05E1\u05D4 \u05DC\u05DE\u05E6\u05D1 \u05DE\u05E0\u05D4\u05DC</h3>" +
      "<table style='border-collapse:collapse;'>" +
      "<tr><td style='padding:4px 12px;font-weight:bold;'>\u05DE\u05E9\u05EA\u05DE\u05E9</td>" +
           "<td style='padding:4px 12px;direction:ltr;'>" + (managerEmail || 'unknown') + "</td></tr>" +
      "<tr style='background:#f5eef8;'><td style='padding:4px 12px;font-weight:bold;'>\u05D6\u05DE\u05DF</td>" +
           "<td style='padding:4px 12px;'>" + timeStr + "</td></tr>" +
      "</table>" +
      "<p style='margin-top:14px;font-size:12px;color:#888;'>\u05E0\u05E9\u05DC\u05D7 \u05D0\u05D5\u05D8\u05D5\u05DE\u05D8\u05D9\u05EA \u05E2\"" +
      "\u05D9 \u05DE\u05E2\u05E8\u05DB\u05EA Neuro Shift</p></div>";
    MailApp.sendEmail({
      to: 'shamir.neuroshift@gmail.com',
      subject: subject,
      htmlBody: body,
      name: 'Neuro Shift'
    });
  } catch(e) {
    console.error('logGodModeLogin error:', e.message);
  }
}