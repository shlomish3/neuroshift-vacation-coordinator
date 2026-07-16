function doGet(e) {
  if (e && e.parameter && e.parameter.debug === 'legacy') {
    const output = debugLegacySheets();
    return ContentService.createTextOutput(output);
  }
  if (e && e.parameter && e.parameter.debug) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let result = [];
    for (let s of sheets) {
      const data = s.getDataRange().getValues();
      if (data.length > 0) {
        let firstRow = data[0];
        let name = s.getName();
        result.push(name + ": " + firstRow.join(" | "));
      }
    }
    return ContentService.createTextOutput(result.join("\n"));
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.doctorDbJson = JSON.stringify(getTeamDatabase());
  template.teamOrderJson   = JSON.stringify(getTeamOrder());
  template.godModeUsersJson = JSON.stringify(getGodModeUsers());
  return template.evaluate()
    .setTitle('תורנויות וחופשים נוירולוגיה')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ניהול מערכת - נוירו')
    .addItem('ייבוא נתוני עבר לגיליון הנקי', 'migrateAllLegacyData')
    .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}