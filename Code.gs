// 利用するスプレッドシートID（環境に合わせて差し替え可能）
const SPREADSHEET_ID = '1bTRSe5l7RTMk1taHNtYaAUMcFBEIwGUf6Yz0icPtp2M';

// 共通でSpreadsheetを取得するヘルパー
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Webアプリのエントリポイント。URLの ?page=xxx に応じてHTMLテンプレートを振り分けます。
// 不正な page の場合は index を返し、読み込みエラー時は簡易エラーページを返します。
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const allowed = ['index', 'facilities', 'visits', 'reports', 'employees'];
  Logger.log('[doGet] raw page param=%s', page);
  // 必要なシートが揃っているか確認（無ければ作成・マイグレーション）
  try { setupSheets(); } catch (err) { Logger.log('[setupSheets][WARN] %s', err && err.message); }
  let target = allowed.indexOf(page) !== -1 ? page : 'index';
  if (target !== page) {
    Logger.log('[doGet] page "%s" は許可リストに無いため "%s" を使用', page, target);
  }
  try {
    // テンプレート評価を使い、HTML内の <?= ... ?> スクリプトレットを有効にする
    const out = HtmlService.createTemplateFromFile(target).evaluate().setTitle('営業管理');
    Logger.log('[doGet] served file=%s', target);
    return out;
  } catch (err) {
    Logger.log('[doGet][ERROR] target=%s message=%s', target, err && err.message);
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:24px">'
      + '<h2>表示エラー</h2>'
      + '<p>ページ "' + sanitize(page) + '" の読み込み中に問題が発生しました。</p>'
      + '<pre style="white-space:pre-wrap;background:#f5f5f5;padding:12px;border:1px solid #ccc">' + sanitize(err && err.message) + '</pre>'
      + '<p><a href="?page=index">メニューへ戻る</a></p>'
      + '</body></html>'
    ).setTitle('表示エラー');
  }
}

// HTMLに埋め込む文字列をサニタイズ（XSS/表示崩れ防止）
function sanitize(str) {
  if (str == null) return '';
  return String(str).replace(/[&<>"']/g, function(ch) {
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[ch]);
  });
}

function addRecord(data) {
  return addFacility(data);
}

// 初期セットアップ：必要なシートが存在しない場合はヘッダ行付きで作成
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
    ss.insertSheet('Reports').appendRow(['id','facilityId','reportDate','reporterName','reporterEmail','channel','summary','details','followUp','createdAt','createdBy']);
  } else {
    // 既存Reportsに channel 列が無ければ追加（summaryの前）
    const rs = ss.getSheetByName('Reports');
    const headers = rs.getRange(1,1,1,rs.getLastColumn()).getValues()[0];
    if (headers.indexOf('channel') === -1) {
      // 6列目に挿入（1始まり）: id(1) facilityId(2) reportDate(3) reporterName(4) reporterEmail(5) channel(6)
      rs.insertColumnBefore(6);
      rs.getRange(1,6).setValue('channel');
    }
  }
  if (!names.includes('Employees')) {
    ss.insertSheet('Employees').appendRow(['id','name','email','phone','role','createdAt','createdBy']);
  }
  if (!names.includes('FacilityContacts')) {
    ss.insertSheet('FacilityContacts').appendRow(['id','facilityId','contactName','contactPhone','contactNotes','createdAt','createdBy']);
  } else {
    // 既存のFacilityContactsにcontactNotes列がなければ追加（contactPhoneの後ろ）
    const cs = ss.getSheetByName('FacilityContacts');
    const headers = cs.getRange(1,1,1,cs.getLastColumn()).getValues()[0];
    if (headers.indexOf('contactNotes') === -1) {
      const phoneIdx = headers.indexOf('contactPhone');
      if (phoneIdx !== -1) {
        cs.insertColumnAfter(phoneIdx + 1); // 1-based index
        cs.getRange(1, phoneIdx + 2).setValue('contactNotes');
      } else {
        cs.insertColumnAfter(cs.getLastColumn());
        cs.getRange(1, cs.getLastColumn()).setValue('contactNotes');
      }
    }
    // 名刺画像のファイルID列がなければ追加（contactNotes の後ろ）
    const headers2 = cs.getRange(1,1,1,cs.getLastColumn()).getValues()[0];
    if (headers2.indexOf('cardFileId') === -1) {
      const notesIdx = headers2.indexOf('contactNotes');
      if (notesIdx !== -1) {
        cs.insertColumnAfter(notesIdx + 1);
        cs.getRange(1, notesIdx + 2).setValue('cardFileId');
      } else {
        cs.insertColumnAfter(cs.getLastColumn());
        cs.getRange(1, cs.getLastColumn()).setValue('cardFileId');
      }
    }
  }
  // 施設の旧 contact 列から FacilityContacts へ自動移行（必要な場合のみ）
  try { migrateFacilityContactsFromFacilities(); } catch (e) { Logger.log('[migrate][WARN] %s', e && e.message); }
}

// Facilities シートの contact 列にある値を FacilityContacts に移行し、元セルは空にします
function migrateFacilityContactsFromFacilities() {
  const ss = getSpreadsheet();
  const fac = ss.getSheetByName('Facilities');
  const con = ss.getSheetByName('FacilityContacts');
  if (!fac || !con) return;
  const fRows = fac.getDataRange().getValues();
  if (fRows.length <= 1) return;
  const fHeaders = fRows[0];
  const idIdx = fHeaders.indexOf('id');
  const contactIdx = fHeaders.indexOf('contact');
  if (idIdx === -1 || contactIdx === -1) return; // 移行対象無し
  const now = nowIso();
  const by = activeUserEmail();
  let wrote = false;

  // 既存の FacilityContacts を取得して重複を避ける
  const cRows = con.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < cRows.length; i++) {
    const r = cRows[i];
    existing.add(String(r[1]) + '::' + String(r[2])); // facilityId::contactName
  }

  for (let i = 1; i < fRows.length; i++) {
    const r = fRows[i];
    const facilityId = String(r[idIdx] || '');
    const contactName = String(r[contactIdx] || '').trim();
    if (!facilityId || !contactName) continue;
    const key = facilityId + '::' + contactName;
    if (existing.has(key)) {
      // すでにある場合は Facilities 側だけ空にする
      fac.getRange(i + 1, contactIdx + 1).setValue('');
      continue;
    }
    // 追記
    con.appendRow([
      makeId('FC'),
      facilityId,
      contactName,
      '', // phone 不明
      '', // notes 空
      now,
      by
    ]);
    // 元の contact を空に
    fac.getRange(i + 1, contactIdx + 1).setValue('');
    wrote = true;
  }
  if (wrote) Logger.log('[migrate] Facilities.contact を FacilityContacts へ移行しました');
}

function nowIso() {
  return new Date().toISOString();
}

// 簡易ユニークID生成（日時＋乱数）
function makeId(prefix) {
  const seed = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  return prefix + '-' + seed + '-' + Math.floor(Math.random() * 1000);
}

// 実行ユーザーのメール（取得できない場合は空文字）
function activeUserEmail() {
  try {
    const user = Session.getActiveUser();
    return user && user.getEmail ? user.getEmail() : '';
  } catch (err) {
    return '';
  }
}

// 入力された日付文字列が不正な場合は fallback/現在日時に置換
function normalizeDate(value, fallback) {
  if (!value) return fallback || nowIso();
  const dt = new Date(value);
  if (isNaN(dt.getTime())) return fallback || nowIso();
  return dt.toISOString();
}

// 施設追加
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

// 施設一覧取得（簡易構造）
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

// 訪問記録追加
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

// 訪問記録一覧取得（facilityId / from / to で絞り込み可）
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

// 営業報告追加
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
  // ヘッダ確認と不足列追加（channel / contactId / contactName）
  const ensureHeaders = () => {
    let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('channel') === -1) {
      sheet.insertColumnBefore(6);
      sheet.getRange(1,6).setValue('channel');
      headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    }
    // followUp の直後に contactId, contactName を並べる
    let updated = headers.slice();
    const followIdx = updated.indexOf('followUp');
    if (updated.indexOf('contactId') === -1) {
      if (followIdx !== -1) {
        sheet.insertColumnAfter(followIdx + 1);
        sheet.getRange(1, followIdx + 2).setValue('contactId');
      } else {
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue('contactId');
      }
      updated = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    }
    if (updated.indexOf('contactName') === -1) {
      const idxContactId = updated.indexOf('contactId');
      if (idxContactId !== -1) {
        sheet.insertColumnAfter(idxContactId + 1);
        sheet.getRange(1, idxContactId + 2).setValue('contactName');
      } else if (followIdx !== -1) {
        sheet.insertColumnAfter(followIdx + 1);
        sheet.getRange(1, followIdx + 2).setValue('contactName');
      } else {
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue('contactName');
      }
    }
  };
  ensureHeaders();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const row = new Array(headers.length).fill('');
  function set(h, val) { if (idx[h] != null) row[idx[h]] = val; }
  set('id', id);
  set('facilityId', data.facilityId);
  set('reportDate', reportDate);
  set('reporterName', data.reporterName || '');
  set('reporterEmail', data.reporterEmail || '');
  set('channel', data.channel || '');
  set('summary', data.summary || '');
  set('details', data.details || '');
  set('followUp', data.followUp || '');
  set('contactId', data.contactId || '');
  set('contactName', data.contactName || '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  return { id, createdAt };
}

// 営業報告一覧取得（facilityId / from / to / キーワード検索 q 対応）
function getReports(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Reports');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  const idx = {};
  headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id] || r[0],
      facilityId: r[idx.facilityId] || r[1],
      reportDate: r[idx.reportDate] || r[2],
      reporterName: r[idx.reporterName] || r[3],
      reporterEmail: r[idx.reporterEmail] || r[4],
      channel: idx.channel != null ? r[idx.channel] : '',
      summary: r[idx.summary] || r[5],
      details: r[idx.details] || r[6],
      followUp: r[idx.followUp] || r[7],
      createdAt: r[idx.createdAt] || r[8],
      createdBy: r[idx.createdBy] || r[9],
      contactId: idx.contactId != null ? r[idx.contactId] : '',
      contactName: idx.contactName != null ? r[idx.contactName] : ''
    };
    if (params.facilityId && item.facilityId !== params.facilityId) continue;
    if (params.from && item.reportDate < params.from) continue;
    if (params.to && item.reportDate > params.to) continue;
    if (params.q) {
      const q = params.q.toLowerCase();
      const text = (item.summary || '') + ' ' + (item.details || '') + ' ' + (item.followUp || '') + ' ' + (item.channel || '') + ' ' + (item.contactName || '');
      if (!text.toLowerCase().includes(q)) continue;
    }
    list.push(item);
  }
  list.sort((a, b) => (b.reportDate || '').localeCompare(a.reportDate || ''));
  return list;
}

// 営業報告をCSV文字列としてエクスポート
function exportReportsCsv(params) {
  const reports = getReports(params);
  const headers = ['id','facilityId','reportDate','reporterName','reporterEmail','channel','summary','details','followUp','contactId','contactName','createdAt','createdBy'];
  const body = reports.map(r => headers.map(h => (r[h] || '').toString().replace(/\r?\n/g, ' ').replace(/"/g, '""')));
  const csv = [headers.join(',')].concat(body.map(row => '"' + row.join('","') + '"')).join('\n');
  return csv;
}

// 施設担当者の更新
function updateFacilityContact(data) {
  if (!data || !data.id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) throw new Error('FacilityContacts シートがありません');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  let target = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx.id]) === String(data.id)) { target = i; break; }
  }
  if (target === -1) throw new Error('指定IDの担当者が見つかりません');
  if (idx.contactName != null) sheet.getRange(target+1, idx.contactName+1).setValue(data.contactName || '');
  if (idx.contactPhone != null) sheet.getRange(target+1, idx.contactPhone+1).setValue(data.contactPhone || '');
  if (idx.contactNotes != null) sheet.getRange(target+1, idx.contactNotes+1).setValue(data.contactNotes || '');
  return { id: data.id };
}

// 施設担当者の削除
function deleteFacilityContact(id) {
  if (!id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) throw new Error('FacilityContacts シートがありません');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx.id]) === String(id)) {
      sheet.deleteRow(i+1);
      return { id: id };
    }
  }
  throw new Error('指定IDの担当者が見つかりません');
}

// 施設担当者追加
function addFacilityContact(data) {
  if (!data || !data.facilityId) throw new Error('facilityId は必須です');
  if (!data.contactName) throw new Error('contactName は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) throw new Error('FacilityContacts シートがありません');
  const id = makeId('FC');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  // ヘッダ動的取得
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const row = new Array(headers.length).fill('');
  function set(h,val){ if (idx[h]!=null) row[idx[h]] = val; }
  set('id', id);
  set('facilityId', data.facilityId);
  set('contactName', data.contactName || '');
  set('contactPhone', data.contactPhone || '');
  set('contactNotes', data.contactNotes || '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  return { id, createdAt };
}

// 施設担当者一覧取得（facilityId で絞り込み）
function getFacilityContacts(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id] || r[0],
      facilityId: r[idx.facilityId] || r[1],
      contactName: r[idx.contactName] || r[2],
      contactPhone: idx.contactPhone!=null ? r[idx.contactPhone] : '',
      contactNotes: idx.contactNotes!=null ? r[idx.contactNotes] : '',
      cardFileId: idx.cardFileId!=null ? r[idx.cardFileId] : '',
      createdAt: r[idx.createdAt] || r[4],
      createdBy: r[idx.createdBy] || r[5]
    };
    if (params.facilityId && item.facilityId !== params.facilityId) continue;
    list.push(item);
  }
  return list;
}

// 名刺画像アップロード用フォルダ取得/作成
function getFacilityCardsFolder_() {
  const name = 'FacilityContactCards';
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

// 施設担当者の名刺画像をアップロードし、cardFileId を更新
function uploadFacilityContactCard(data) {
  if (!data || !data.contactId) throw new Error('contactId は必須です');
  if (!data.dataUrl && !data.base64) throw new Error('画像データがありません');
  let contentType = data.contentType || '';
  let b64 = data.base64 || '';
  if (data.dataUrl) {
    const m = String(data.dataUrl).match(/^data:(.*?);base64,(.*)$/);
    if (!m) throw new Error('dataUrl の形式が不正です');
    contentType = contentType || m[1];
    b64 = m[2];
  }
  if (!contentType) contentType = 'image/png';
  const bytes = Utilities.base64Decode(b64);
  const ext = (function(mt){
    if (mt.indexOf('jpeg') !== -1 || mt.indexOf('jpg') !== -1) return '.jpg';
    if (mt.indexOf('png') !== -1) return '.png';
    if (mt.indexOf('gif') !== -1) return '.gif';
    if (mt.indexOf('heic') !== -1) return '.heic';
    return '';
  })(contentType.toLowerCase());
  const fname = (data.filename && String(data.filename).trim()) || ('card_' + data.contactId + '_' + Date.now() + ext);
  const folder = getFacilityCardsFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, fname));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const fileId = file.getId();

  // シート更新
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) throw new Error('FacilityContacts シートがありません');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  // cardFileId 列が無ければ追加
  if (idx.cardFileId == null) {
    const notesIdx = headers.indexOf('contactNotes');
    if (notesIdx !== -1) {
      sheet.insertColumnAfter(notesIdx + 1);
      sheet.getRange(1, notesIdx + 2).setValue('cardFileId');
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('cardFileId');
    }
  }
  // 再取得
  const headers2 = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx2 = {}; headers2.forEach((h,i)=> idx2[h]=i);
  let target = -1;
  for (let i = 1; i < rows.length; i++) {
    const idVal = rows[i][idx.id || 0];
    if (String(idVal) === String(data.contactId)) { target = i; break; }
  }
  if (target === -1) throw new Error('指定IDの担当者が見つかりません');
  sheet.getRange(target+1, idx2.cardFileId+1).setValue(fileId);
  return { fileId: fileId, viewUrl: 'https://drive.google.com/uc?export=view&id=' + fileId };
}

// 社員レコード追加
function addEmployee(data) {
  if (!data || !data.name) throw new Error('name は必須です');
  const name = String(data.name).trim();
  if (!name) throw new Error('name は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees シートが見つかりません。setupSheets() を実行してください。');

  // 重複登録防止：名前（大文字小文字無視、前後空白除去）で既存チェック
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const existingName = String(r[1] || '').trim().toLowerCase();
    if (existingName && existingName === name.toLowerCase()) {
      return { id: r[0], createdAt: r[5] }; // 既存レコードを返す
    }
  }

  const id = makeId('EMP');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  sheet.appendRow([
    id,
    name,
    (data.email || ''), // 互換列。今は空で保存
    data.phone || '',
    data.role || '',
    createdAt,
    createdBy
  ]);
  return { id, createdAt };
}

// 社員一覧取得
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

// 社員更新（id 必須）
function updateEmployee(data) {
  if (!data || !data.id) throw new Error('id は必須です');
  if (!data.name) throw new Error('name は必須です');
  const id = String(data.id);
  const name = String(data.name).trim();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees シートが見つかりません。setupSheets() を実行してください。');

  const values = sheet.getDataRange().getValues();
  let targetRow = -1; // 0-based（ヘッダ含む）
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[0]) === id) {
      targetRow = i; break;
    }
  }
  if (targetRow === -1) throw new Error('指定された社員IDが見つかりません: ' + id);

  // 同名重複チェック（自身は除外）
  for (let i = 1; i < values.length; i++) {
    if (i === targetRow) continue;
    const r = values[i];
    const existingName = String(r[1] || '').trim().toLowerCase();
    if (existingName && existingName === name.toLowerCase()) {
      throw new Error('同じ氏名の社員が既に存在します');
    }
  }

  // 更新：列順は [id, name, email, phone, role, createdAt, createdBy]
  sheet.getRange(targetRow + 1, 2).setValue(name); // name
  sheet.getRange(targetRow + 1, 3).setValue(data.email || ''); // email（互換）
  sheet.getRange(targetRow + 1, 4).setValue(data.phone || ''); // phone
  sheet.getRange(targetRow + 1, 5).setValue(data.role || ''); // role
  return { id: id };
}

// 社員削除（id 必須）
function deleteEmployee(id) {
  if (!id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees シートが見つかりません。setupSheets() を実行してください。');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[0]) === String(id)) {
      sheet.deleteRow(i + 1); // 1-based index, ヘッダ含む
      return { id: id };
    }
  }
  throw new Error('指定された社員IDが見つかりません: ' + id);
}
