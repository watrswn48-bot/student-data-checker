function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('فاحص بيانات الطلاب')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('افتح المشروع من داخل Google Sheets المرتبط بالبيانات.');
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {sheetName:sheet.getName(), total:0, incomplete:0, missingCells:0, students:[]};

  const headers = sheet.getRange(2,1,1,lastCol).getDisplayValues()[0].map(x => String(x||'').trim());
  const rows = lastRow >= 3 ? sheet.getRange(3,1,lastRow-2,lastCol).getDisplayValues() : [];
  const candidates = ['اسم الطالب','اسم التلميذ','الاسم','اسم'];
  const normalized = headers.map(h => h.replace(/\s+/g,'').toLowerCase());
  let nameCol = -1;
  for (const candidate of candidates) {
    const c = candidate.replace(/\s+/g,'').toLowerCase();
    nameCol = normalized.findIndex(h => h === c || h.includes(c));
    if (nameCol >= 0) break;
  }
  if (nameCol < 0) nameCol = 0;

  const students = [];
  let missingCells = 0;
  rows.forEach((row, i) => {
    if (row.every(v => String(v||'').trim() === '')) return;
    const missing = [];
    headers.forEach((h, j) => { if (h && String(row[j]||'').trim() === '') missing.push(h); });
    if (missing.length) {
      missingCells += missing.length;
      students.push({rowNumber:i+3, name:String(row[nameCol]||'بدون اسم').trim() || 'بدون اسم', missing});
    }
  });
  return {sheetName:sheet.getName(), total:rows.filter(r=>r.some(v=>String(v||'').trim()!=='')).length,
    incomplete:students.length, missingCells, students, updatedAt:new Date().toISOString()};
}