function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function formatDateValue(val, overrideYear) {
  if (!val) return "";
  if (val instanceof Date) {
    let d = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (overrideYear) {
      d = d.substring(0, 6) + overrideYear;
    }
    return d;
  }
  let str = String(val).trim();
  const parts = str.split(/[\/\-.]/);
  if (parts.length >= 2) {
    const d = parts[0].padStart(2, '0');
    const m = parts[1].padStart(2, '0');
    let y = parts[2] || new Date().getFullYear().toString();
    if (y.length === 2) {
      y = "20" + y;
    }
    if (overrideYear) y = overrideYear;
    return `${d}/${m}/${y}`;
  }
  return str;
}

function formatLegacyDates(dateStr, isImportant) {
  if (!dateStr) return "";
  const days = ["יום א'", "יום ב'", "יום ג'", "יום ד'", "יום ה'", "יום ו'", "שבת"];
  const dateArr = dateStr.split(', ');
  const formattedArr = dateArr.map(d => {
    const parts = d.split('/');
    if (parts.length === 3) {
      const dateObj = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      return d + " " + days[dateObj.getDay()] + (isImportant ? " (חשוב)" : "");
    }
    return d;
  });
  return formattedArr.join(', ');
}