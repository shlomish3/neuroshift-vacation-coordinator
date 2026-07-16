function getTeamDatabase() {
  mergeDuplicateNames();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let teamSheet = ss.getSheetByName('רשימת צוות');
  
  if (!teamSheet) {
    teamSheet = ss.insertSheet('רשימת צוות');
    teamSheet.appendRow(['Email', 'Name', 'Role']);
    const defaultData = {
      "avigailbartal@hotmail.com": { name: "ד\"ר אביגיל ברטל", role: "בכיר" },
      "gandelman@shamir.gov.il": { name: "פרופ' רויטל גנדלמן-מרטון", role: "בכיר" },
      "nirhersh@gmail.com": { name: "ד\"ר ניר הרש", role: "בכיר" },
      "coheno@shamir.gov.il": { name: "פרופ' אורן כהן", role: "בכיר" },
      "oren.s.cohen@gmail.com": { name: "פרופ' אורן כהן", role: "בכיר" },
      "shlomip@shamir.gov.il": { name: "ד\"ר שלומי פרץ", role: "בכיר" },
      "itzhakk@shamir.gov.il": { name: "ד\"ר יצחק קימיאגר", role: "בכיר" },
      "giladankori@gmail.com": { name: "ד\"ר גלעד קינן", role: "בכיר" },
      "nitai.shimon@gmail.com": { name: "ד\"ר נתאי שמעון", role: "בכיר" },
      "nettastr@gmail.com": { name: "ד\"ר נטע אגאג'ני", role: "מתמחה" },
      "ardash.nat@gmail.com": { name: "ד\"ר נטליה ארדשירוב", role: "מתמחה" },
      "berg.assaf@gmail.com": { name: "ד\"ר אסף ברג", role: "מתמחה" },
      "assaf.berg@mail.huji.ac.il": { name: "ד\"ר אסף ברג", role: "מתמחה" },
      "sofimdneuro@gmail.com": { name: "ד\"ר סופיה גלינסקיה", role: "מתמחה" },
      "liordekel3@gmail.com": { name: "ד\"ר ליאור דקל", role: "מתמחה" },
      "eladhaser7@gmail.com": { name: "ד\"ר אלעד הסר", role: "מתמחה" },
      "ahmad3x@hotmail.com": { name: "ד\"ר אחמד חדיג'ה", role: "מתמחה" },
      "albasantclinic@gmail.com": { name: "ד\"ר אחמד חדיג'ה", role: "מתמחה" },
      "lakinsheli@gmail.com": { name: "ד\"ר שלי לקן", role: "מתמחה" },
      "hossensaoub@gmail.com": { name: "ד\"ר חוסין סעוב", role: "מתמחה" },
      "aya_asly@hotmail.com": { name: "ד\"ר איה עסלי", role: "מתמחה" },
      "aviranpriante93@gmail.com": { name: "ד\"ר אבירן פריאנטה", role: "מתמחה" },
      "shlomi.shmuel3@gmail.com": { name: "ד\"ר שלומי שמואל", role: "מתמחה" },
      "miniovitcha@shamir.gov.il": { name: "ד\"ר אלה מיניוביץ'", role: "אחר" },
      "resident1@g.com": { name: "מתמחה טסט 1", role: "מתמחה" },
      "resident2@g.com": { name: "מתמחה טסט 2", role: "מתמחה" },
      "senior1@g.com": { name: "בכיר טסט 1", role: "בכיר" },
      "senior2@g.com": { name: "בכיר טסט 2", role: "בכיר" }
    };
    
    for (const email in defaultData) {
      teamSheet.appendRow([email, defaultData[email].name, defaultData[email].role]);
    }
  }

  const data = teamSheet.getDataRange().getValues();
  if (data.length < 1) return {};
  
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  let eIdx = headers.indexOf('email');
  let nIdx = headers.indexOf('name');
  let rIdx = headers.indexOf('role');
  let e2Idx = headers.indexOf('email2');
  
  // Fallback if headers are missing or renamed
  if (eIdx === -1) eIdx = 0;
  if (nIdx === -1) nIdx = 1;
  if (rIdx === -1) rIdx = 2;
  if (e2Idx === -1) e2Idx = 3;

  const db = {};
  for (let i = 1; i < data.length; i++) {
    const email1 = String(data[i][eIdx] || "").trim().toLowerCase();
    const name = String(data[i][nIdx] || "").trim();
    const role = String(data[i][rIdx] || "").trim();
    const email2 = String(data[i][e2Idx] || "").trim().toLowerCase();
    
    if (email1 === 'email' || name.toLowerCase() === 'name') continue;
    
    if (email1) {
      db[email1] = { name: name, role: role, email2: email2 };
      if (email2) {
         db[email2] = { name: name, role: role, primaryEmail: email1 };
      }
    }
  }
  return db;
}

function getTeamOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamSheet = ss.getSheetByName('רשימת צוות');
  if (!teamSheet) return [];
  const data = teamSheet.getDataRange().getValues();
  if (data.length < 1) return [];
  
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  let eIdx = headers.indexOf('email');
  if (eIdx === -1) eIdx = 0;
  
  const order = [];
  for (let i = 1; i < data.length; i++) {
    const email1 = String(data[i][eIdx] || "").trim().toLowerCase();
    if (email1 && email1 !== 'email' && email1 !== 'name') {
      order.push(email1);
    }
  }
  return order;
}

function saveTeamDatabase(newRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let teamSheet = ss.getSheetByName('רשימת צוות');
  if (!teamSheet) {
    teamSheet = ss.insertSheet('רשימת צוות');
  }
  const data = teamSheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : ['#', 'Email', 'Name', 'Role', 'Email2'];
  
  const hLower = headers.map(h => String(h).toLowerCase().trim());
  let eIdx = hLower.indexOf('email');
  let nIdx = hLower.indexOf('name');
  let rIdx = hLower.indexOf('role');
  let e2Idx = hLower.indexOf('email2');
  let numIdx = hLower.indexOf('#');
  
  if (eIdx === -1) eIdx = 1;
  if (nIdx === -1) nIdx = 2;
  if (rIdx === -1) rIdx = 3;
  if (e2Idx === -1) e2Idx = 4;
  if (numIdx === -1) numIdx = 0;
  
  teamSheet.clear();
  teamSheet.appendRow(headers);
  
  for (let i = 0; i < newRows.length; i++) {
    const rowObj = newRows[i];
    let newRow = new Array(headers.length).fill('');
    newRow[numIdx] = i + 1;
    newRow[eIdx] = rowObj.email1 || '';
    newRow[nIdx] = rowObj.name || '';
    newRow[rIdx] = rowObj.role || '';
    newRow[e2Idx] = rowObj.email2 || '';
    teamSheet.appendRow(newRow);
  }
  return true;
}

function getExistingVacations() {
  const sheetName = 'Vacation_Requests_Clean'; 
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return {};
  
  const values = sheet.getDataRange().getValues();
  const vacations = {};
  const notesObj = {};
  const doctorDatabase = getTeamDatabase();
  
  for (let i = 1; i < values.length; i++) {
    const email = String(values[i][3]).toLowerCase().trim();
    const name = String(values[i][2]).trim();
    let role = String(values[i][1]).trim();
    const notesText = String(values[i][6]).trim();
    
    if (!role && doctorDatabase[email]) {
      role = doctorDatabase[email].role;
    }
    
    if (role.includes('@')) {
      const legacyEmail = role.toLowerCase();
      role = doctorDatabase[legacyEmail] ? doctorDatabase[legacyEmail].role : "אחר";
      
      const legacyDateVal = formatDateValue(values[i][3]);
      const legacyDateType = String(values[i][4]).trim();
      
      if (notesText && legacyDateVal && name) {
        const parts = legacyDateVal.split('/');
        if (parts.length === 3) {
          const monthKey = parts[1] + '-' + parts[2];
          if (!notesObj[name]) notesObj[name] = {};
          if (!notesObj[name][monthKey]) notesObj[name][monthKey] = [];
          if (!notesObj[name][monthKey].includes(notesText)) notesObj[name][monthKey].push(notesText);
        }
      }
      
      if ((legacyDateType === "חופש" || legacyDateType === "לא יכול תורנות" || legacyDateType === "מועדף" || legacyDateType === "מועדף (חשוב)") && legacyDateVal) {
        const finalLegacyType = legacyDateType === "מועדף (חשוב)" ? "מועדף" : legacyDateType;
        if (!vacations[legacyDateVal]) vacations[legacyDateVal] = [];
        if (!vacations[legacyDateVal].some(v => v.name === name && v.type === finalLegacyType)) {
          vacations[legacyDateVal].push({ name: name, role: role, type: finalLegacyType });
        }
      }
      continue;
    }
    
    const dateVal = formatDateValue(values[i][4]);
    const dateType = String(values[i][5]).trim();
    const status = String(values[i][8] || "Approved").trim();
    
    if (notesText && dateVal && name) {
      const parts = dateVal.split('/');
      if (parts.length === 3) {
        const monthKey = parts[1] + '-' + parts[2];
        if (!notesObj[name]) notesObj[name] = {};
        if (!notesObj[name][monthKey]) notesObj[name][monthKey] = [];
        if (!notesObj[name][monthKey].includes(notesText)) notesObj[name][monthKey].push(notesText);
      }
    }
    
    if ((dateType === "חופש" || dateType === "לא יכול תורנות" || dateType === "מועדף" || dateType === "מועדף (חשוב)") && dateVal) {
      const finalDateType = dateType === "מועדף (חשוב)" ? "מועדף" : dateType;
      if (!vacations[dateVal]) vacations[dateVal] = [];
      if (!vacations[dateVal].some(v => v.name === name && v.type === finalDateType)) {
        vacations[dateVal].push({ name: name, role: role, type: finalDateType, email: email, status: status });
      }
    }
  }
  const holidays = {};
  try {
    const templateSs = SpreadsheetApp.openById('10azqdWsqw_E2K1lRRd7UTMFAba9zYDyRMY89DcSxt5A');
    const holidaysSheet = templateSs.getSheetByName("חגים");
    if (holidaysSheet) {
      const hValues = holidaysSheet.getDataRange().getValues();
      for (let i = 1; i < hValues.length; i++) {
        let dateVal = hValues[i][0];
        let name = hValues[i][1];
        if (dateVal instanceof Date) {
          let d = ("0" + dateVal.getDate()).slice(-2);
          let m = ("0" + (dateVal.getMonth() + 1)).slice(-2);
          let y = dateVal.getFullYear();
          holidays[`${d}/${m}/${y}`] = name;
        }
      }
    }
  } catch (e) {
    Logger.log("Error fetching holidays: " + e.toString());
    holidays["error"] = e.toString();
  }
  
  return { dates: vacations, notes: notesObj, holidays: holidays };
}

function migrateAllLegacyData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const cleanSheet = spreadsheet.getSheetByName('Vacation_Requests_Clean');
  if (!cleanSheet) {
    SpreadsheetApp.getUi().alert('שגיאה: גיליון Vacation_Requests_Clean לא קיים.');
    return;
  }
  
  const sheets = spreadsheet.getSheets();
  const doctorDatabase = getTeamDatabase();
  let addedRows = 0;
  
  // Get existing signatures to prevent duplicates
  const existingValues = cleanSheet.getDataRange().getValues();
  const existingSet = new Set();
  for (let i = 1; i < existingValues.length; i++) {
    const row = existingValues[i];
    const dVal = formatDateValue(row[4]);
    const signature = `${row[2]}_${dVal}_${row[5]}`; // Name_Date_Type
    existingSet.add(signature);
  }
  
  const rowsToAppend = [];
  
  for (let sheet of sheets) {
    const sName = String(sheet.getName()).trim();
    // Skip system sheets and the superfluous 'יוליל' sheet
    if (sName === 'Vacation_Requests_Clean' || sName === 'רשימת צוות' || sName.toLowerCase().includes('dashboard') || sName === 'תורנויות וחופשים נוירולוגיה' || sName.includes('יוליל')) {
      continue;
    }
    
    let overrideYear = null;
    const yearMatch = sName.match(/\b(20\d{2}|\d{2})\b/);
    if (yearMatch) {
       let yStr = yearMatch[0];
       if (yStr.length === 2) yStr = "20" + yStr;
       overrideYear = yStr;
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) continue;
    
    const headers = values[0].map(h => String(h).trim());
    let emailIdx = headers.findIndex(h => h.includes("אימייל") || h.includes("email") || h.includes("דוא\"ל"));
    let nameIdx = headers.findIndex(h => h.includes("שם") && !h.includes("משפחה"));
    let cannotIdx = headers.findIndex(h => h.includes("לא יכול") || h.includes("אינני יכול"));
    let vacIdx = headers.findIndex(h => h.includes("חופש") && !h.includes("מועדף") && !h.includes("מועדפ"));
    let prefIdx = headers.findIndex(h => (h.includes("מועדף") || h.includes("מועדפ") || h.includes("העדפ")) && !h.includes("התייחסות") && !h.includes("הערות"));
    let notesIdx = headers.findIndex(h => h.includes("הערות") || h.includes("התייחסות חופשית"));
    
    if (emailIdx === -1) emailIdx = 1;
    if (nameIdx === -1) nameIdx = 2;
    if (cannotIdx === -1) cannotIdx = 3;
    if (vacIdx === -1) vacIdx = 4;
    if (notesIdx === -1) notesIdx = 5;
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const timestamp = row[0] instanceof Date ? row[0] : new Date(0); // 1970 timestamp so it doesn't override real user submissions
      const email = emailIdx >= 0 && row[emailIdx] ? String(row[emailIdx]).trim().toLowerCase() : "";
      let name = nameIdx >= 0 && row[nameIdx] ? String(row[nameIdx]).trim() : "";
      
      let role = "אחר";
      let finalEmail = email;
      if (email && doctorDatabase[email]) {
         name = doctorDatabase[email].name;
         role = doctorDatabase[email].role;
         finalEmail = doctorDatabase[email].primaryEmail || email;
      } else if (name) {
          for (let k of Object.keys(doctorDatabase)) {
            const dbName = doctorDatabase[k].name.replace(/(ד"ר|ד״ר|ד''ר|פרופ'|פרופסור)\s*/, '').trim();
            const legacyName = name.replace(/(ד"ר|ד״ר|ד''ר|פרופ'|פרופסור)\s*/, '').trim();
            if (dbName === legacyName || legacyName.includes(dbName) || dbName.includes(legacyName)) {
              name = doctorDatabase[k].name;
              role = doctorDatabase[k].role;
              finalEmail = doctorDatabase[k].primaryEmail || k;
              break;
            }
          }
      }
      if (!name) continue;
      
      const notes = notesIdx >= 0 && row[notesIdx] ? String(row[notesIdx]) : "";
      const cannotStr = cannotIdx >= 0 && row[cannotIdx] ? String(row[cannotIdx]) : "";
      const vacStr = vacIdx >= 0 && row[vacIdx] ? String(row[vacIdx]) : "";
      const prefStr = prefIdx >= 0 && row[prefIdx] ? String(row[prefIdx]) : "";
      
      const editId = Utilities.getUuid();
      
      const parseAndAppend = (str, type) => {
        if (!str) return;
        const parts = str.split(/[,;\n]+/).map(s => s.trim()).filter(Boolean);
        for (let p of parts) {
          let dVal = p;
          if (p.includes(' ')) dVal = p.split(' ')[0];
          dVal = formatDateValue(dVal, overrideYear);
          if (dVal) {
            let finalType = type;
            if (type === "מועדף" && p.includes("חשוב")) finalType = "מועדף (חשוב)";
            
            const signature = `${name}_${dVal}_${finalType}_${finalEmail}`;
            if (!existingSet.has(signature)) {
              rowsToAppend.push([timestamp, role, name, finalEmail, dVal, finalType, notes, editId]);
              existingSet.add(signature);
            }
          }
        }
      };
      
      parseAndAppend(cannotStr, "לא יכול תורנות");
      parseAndAppend(vacStr, "חופש");
      parseAndAppend(prefStr, "מועדף");
    }
  }
  
  if (rowsToAppend.length > 0) {
    rowsToAppend.sort((a, b) => {
      const timeA = new Date(a[0]).getTime() || 0;
      const timeB = new Date(b[0]).getTime() || 0;
      return timeA - timeB;
    });
    cleanSheet.getRange(cleanSheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
  }
  
  if (cleanSheet.getLastRow() > 1) {
    cleanSheet.getRange(2, 1, cleanSheet.getLastRow() - 1, cleanSheet.getLastColumn()).sort(1);
  }
  
  SpreadsheetApp.getUi().alert('ייבוא הסתיים בהצלחה! הרשומות מויינו באופן כרונולוגי.');
}

function mergeDuplicateNames() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('mergedTeams3') === 'true') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let teamSheet = ss.getSheetByName('רשימת צוות');
  if (!teamSheet) return;
  
  const data = teamSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const nameToEmails = {};
  for (let i = 1; i < data.length; i++) {
    const email1 = String(data[i][0]).trim().toLowerCase();
    const name = String(data[i][1]).trim();
    const role = String(data[i][2]).trim();
    const email2 = String(data[i][3] || "").trim().toLowerCase();
    if (!name || email1 === 'email' || name.toLowerCase() === 'name') continue;

    if (!nameToEmails[name]) {
      nameToEmails[name] = { emails: new Set(), role: role };
    }
    if (email1) nameToEmails[name].emails.add(email1);
    if (email2) nameToEmails[name].emails.add(email2);
  }

  teamSheet.clear();
  teamSheet.appendRow(['Email', 'Name', 'Role', 'Email2']);
  for (const name in nameToEmails) {
    const emails = Array.from(nameToEmails[name].emails);
    const email1 = emails[0] || "";
    const email2 = emails[1] || "";
    teamSheet.appendRow([email1, name, nameToEmails[name].role, email2]);
  }
  
  props.setProperty('mergedTeams3', 'true');
}