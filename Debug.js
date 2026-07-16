function restoreFeb26() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('פברואר 26');
  if (!sheet) {
    sheet = ss.insertSheet('פברואר 26');
  }
  sheet.clear();
  sheet.appendRow(['חותמת זמן', 'כתובת אימייל', 'שם', 'לילות בהם אינכם יכולים לעשות תורנות/כוננות', 'בחירת ימי חופש - כולל ימי שישי', 'התייחסות חופשית + תאריכים מועדפים']);
  
  const data = [
    { email: "ardash.nat@gmail.com", name: "ד\"ר נטליה ארדשיר", dates: "1, 2, 3, 4, 5, 6, 7" },
    { email: "assaf.berg@mail.huji.ac.il", name: "ד\"ר אסף ברג", dates: "1, 2, 3, 4, 5" },
    { email: "gandelman@shamir.gov.il", name: "פרופ' רויטל גנדלמן", dates: "2, 15, 16, 23" },
    { email: "lakinsheli@gmail.com", name: "ד\"ר שלי לקן", dates: "2, 16, 17" },
    { email: "aya_asly@hotmail.com", name: "ד\"ר איה עסלי", dates: "4, 18, 19, 20, 21, 22, 25, 27, 28" },
    { email: "eladhaser7@gmail.com", name: "ד\"ר אלעד הסר", dates: "10, 11, 22, 23" },
    { email: "coheno@shamir.gov.il", name: "פרופ' אורן כהן", dates: "13, 19, 20, 27" },
    { email: "shlomi.shmuel3@gmail.com", name: "ד\"ר שלומי שמואל", dates: "13, 20" },
    { email: "aviranpriante93@gmail.com", name: "ד\"ר אבירן פריאנטה", dates: "17" },
    { email: "hossensaoub@gmail.com", name: "ד\"ר חוסיין סעוב", dates: "18, 19, 20, 21, 28" },
    { email: "shlomip@shamir.gov.il", name: "ד\"ר שלומי פרץ", dates: "18" },
    { email: "liordekel3@gmail.com", name: "ד\"ר ליאור דקל", dates: "20" },
    { email: "giladankori@gmail.com", name: "ד\"ר גלעד קינן", dates: "20" },
    { email: "nirhersh@gmail.com", name: "ד\"ר ניר הרש", dates: "22" },
    { email: "nitai.shimon@gmail.com", name: "ד\"ר נתאי שמעון", dates: "25" }
  ];
  
  const daysHe = ["יום א'", "יום ב'", "יום ג'", "יום ד'", "יום ה'", "יום ו'", "שבת"];
  
  const formatDates = (dateStr) => {
    if (!dateStr) return "";
    return dateStr.split(',').map(d => {
      let dayNum = parseInt(d.trim(), 10);
      let dateObj = new Date(2026, 1, dayNum); // Month 1 is February
      let dayOfWeek = daysHe[dateObj.getDay()];
      let dayStr = dayNum < 10 ? "0" + dayNum : dayNum;
      return `${dayStr}/02/2026 ${dayOfWeek}`;
    }).join(', ');
  };
  
  const timestamp = "1/2/2026 10:00:00";
  for (const row of data) {
    sheet.appendRow([timestamp, row.email, row.name, "", formatDates(row.dates), ""]);
  }
  return "Successfully restored פברואר 26";
}
