const SPREADSHEET_ID = '1bTRSe5l7RTMk1taHNtYaAUMcFBEIwGUf6Yz0icPtp2M';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const allowed = ['index', 'facilities', 'visits', 'reports', 'employees'];
  const target = allowed.indexOf(page) !== -1 ? page : 'index';
  return HtmlService.createHtmlOutputFromFile(target).setTitle('営業管理');
}

function addRecord(data) {
  return addFacility(data);
}

function setupSheets() {
  const ss = getSpreadsheet();
  const names = ss.getSheets().map(s => s.getName());
  if (!names.includes('Facilities')) {
    ss.insertSheet('Facilities').appendRow(['id','name','address','phone','contact','notes','createdAt','createdBy']);
  }
  if (!names.includes('Visits')) {
    ss.insertSheet('Visits').appendRow(['id','facilityId','visitDate','visitorName','visitorEmail','purpose','notes','createdAt','createdBy']);
  }
  if (!names.includes('Reports')) {
    ss.insertSheet('Reports').appendRow(['id','facilityId','reportDate','reporterName','reporterEmail','summary','details','followUp','createdAt','createdBy']);
  }
  if (!names.includes('Employees')) {
    ss.insertSheet('Employees').appendRow(['id','name','email','phone','role','createdAt','createdBy']);
  }
}

function nowIso() {
  return new Date().toISOString();
}

function makeId(prefix) {
  const seed = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  return prefix + '-' + seed + '-' + Math.floor(Math.random() * 1000);
}

function activeUserEmail() {
  try {
    const user = Session.getActiveUser();
    return user && user.getEmail ? user.getEmail() : '';
  } catch (err) {
    return '';
  }
}

function normalizeDate(value, fallback) {
  if (!value) return fallback || nowIso();
  const dt = new Date(value);
  if (isNaN(dt.getTime())) return fallback || nowIso();
  return dt.toISOString();
}

function addFacility(data) {
  if (!data || !data.name) throw new Error('施設名は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) throw new Error('Facilities シートが見つかりません。setupSheets() を実行してください。');
  const id = makeId('FAC');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  sheet.appendRow([id, data.name || '', data.address || '', data.phone || '', data.contact || '', data.notes || '', createdAt, createdBy]);
  return { id, createdAt };
}

function getFacilities() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    list.push({ id: r[0], name: r[1], address: r[2], phone: r[3], contact: r[4], notes: r[5] });
  }
  return list;
}

function addVisit(data) {
  if (!data || !data.facilityId) throw new Error('facilityId は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) throw new Error('Visits シートが見つかりません。setupSheets() を実行してください。');
  const id = makeId('VIS');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  const visitDate = normalizeDate(data.visitDate, nowIso());
  sheet.appendRow([
    id,
    data.facilityId,
    visitDate,
    data.visitorName || '',
    data.visitorEmail || '',
    data.purpose || '',
    data.notes || '',
    createdAt,
    createdBy
  ]);
  return { id, createdAt };
}

function getVisits(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[0],
      facilityId: r[1],
      visitDate: r[2],
      visitorName: r[3],
      visitorEmail: r[4],
      purpose: r[5],
      notes: r[6],
      createdAt: r[7],
      createdBy: r[8]
    };
    if (params.facilityId && item.facilityId !== params.facilityId) continue;
    if (params.from && item.visitDate < params.from) continue;
    if (params.to && item.visitDate > params.to) continue;
    list.push(item);
  }
  list.sort((a, b) => (b.visitDate || '').localeCompare(a.visitDate || ''));
  return list;
}

function addReport(data) {
  if (!data || !data.facilityId) throw new Error('facilityId は必須です');
  if (!data.summary) throw new Error('summary は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');
  if (!sheet) throw new Error('Reports シートが見つかりません。setupSheets() を実行してください。');
  const id = makeId('RPT');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  const reportDate = normalizeDate(data.reportDate, createdAt);
  sheet.appendRow([
    id,
    data.facilityId,
    reportDate,
    data.reporterName || '',
    data.reporterEmail || '',
    data.summary || '',
    data.details || '',
    data.followUp || '',
    createdAt,
    createdBy
  ]);
  return { id, createdAt };
}

function getReports(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[0],
      facilityId: r[1],
      reportDate: r[2],
      reporterName: r[3],
      reporterEmail: r[4],
      summary: r[5],
      details: r[6],
      followUp: r[7],
      createdAt: r[8],
      createdBy: r[9]
    };
    if (params.facilityId && item.facilityId !== params.facilityId) continue;
    if (params.from && item.reportDate < params.from) continue;
    if (params.to && item.reportDate > params.to) continue;
    if (params.q) {
      const q = params.q.toLowerCase();
      const text = (item.summary || '') + ' ' + (item.details || '') + ' ' + (item.followUp || '');
      if (!text.toLowerCase().includes(q)) continue;
    }
    list.push(item);
  }
  list.sort((a, b) => (b.reportDate || '').localeCompare(a.reportDate || ''));
  return list;
}

function exportReportsCsv(params) {
  const reports = getReports(params);
  const headers = ['id','facilityId','reportDate','reporterName','reporterEmail','summary','details','followUp','createdAt','createdBy'];
  const body = reports.map(r => headers.map(h => (r[h] || '').toString().replace(/\r?\n/g, ' ').replace(/"/g, '""')));
  const csv = [headers.join(',')].concat(body.map(row => '"' + row.join('","') + '"')).join('\n');
  return csv;
}

function addEmployee(data) {
  if (!data || !data.name) throw new Error('name は必須です');
  if (!data.email) throw new Error('email は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees シートが見つかりません。setupSheets() を実行してください。');
  const id = makeId('EMP');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  sheet.appendRow([
    id,
    data.name || '',
    data.email || '',
    data.phone || '',
    data.role || '',
    createdAt,
    createdBy
  ]);
  return { id, createdAt };
}

function getEmployees() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    list.push({
      id: r[0],
      name: r[1],
      email: r[2],
      phone: r[3],
      role: r[4],
      createdAt: r[5],
      createdBy: r[6]
    });
  }
  return list;
}
