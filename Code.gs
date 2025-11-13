// 利用するスプレッドシートID（環境に合わせて差し替え可能）
const SPREADSHEET_ID = '1bTRSe5l7RTMk1taHNtYaAUMcFBEIwGUf6Yz0icPtp2M';
const CONFIG_SHEET_NAME = 'Config';

// 共通でSpreadsheetを取得するヘルパー
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Webアプリのエントリポイント。URLの ?page=xxx に応じてHTMLテンプレートを振り分けます。
// 不正な page の場合は index を返し、読み込みエラー時は簡易エラーページを返します。
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const allowed = ['index', 'facilities', 'visits', 'reports', 'employees', 'manage', 'targets', 'partners'];
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

// 施設検索（prefecture / municipality / facilityType / q）
function searchFacilities(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  const limit = Number(params.limit || 0);
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id] || r[0],
      name: r[idx.name] || r[1],
      prefecture: idx.prefecture!=null ? r[idx.prefecture] : '',
      municipality: idx.municipality!=null ? r[idx.municipality] : '',
      facilityType: idx.facilityType!=null ? r[idx.facilityType] : '',
      contact: idx.contact!=null ? r[idx.contact] : '',
      address: r[idx.address] != null ? r[idx.address] : r[2],
      phone: r[idx.phone] != null ? r[idx.phone] : r[3],
      notes: idx.notes!=null ? r[idx.notes] : ''
    };
    if (params.prefecture && item.prefecture !== params.prefecture) continue;
    if (params.municipality && item.municipality !== params.municipality) continue;
    if (params.facilityType) {
      const sel = String(params.facilityType).trim();
      const src = String(item.facilityType || '');
      if (!src) continue;
      const tokens = src.split(/[／,、]/).map(s=>String(s).trim()).filter(Boolean);
      const match = sel === 'その他'
        ? tokens.some(t => t === 'その他' || t.indexOf('その他:') === 0)
        : tokens.indexOf(sel) !== -1;
      if (!match) continue;
    }
    if (params.q) {
      const q = String(params.q).toLowerCase();
      const hay = (item.name+' '+item.address+' '+(item.municipality||'')+' '+(item.notes||'')+' '+(item.facilityType||'')).toLowerCase();
      if (!hay.includes(q)) continue;
    }
    list.push(item);
    if (limit > 0 && list.length >= limit) break;
  }
  // 名前でソート
  list.sort((a,b)=> String(a.name||'').localeCompare(String(b.name||'')));
  return list;
}

// 住所単位で施設を統合表示（同一住所の複数サービスを1行化）
// params: { prefecture?, municipality?, facilityType?, q? }
function searchFacilitiesGrouped(params) {
  params = params || {};
  const all = searchFacilities({}); // グループ化の完全性のため全件を取得
  const groups = new Map();
  all.forEach(f => {
    // フィルタ判定（施設タイプはグループに少なくとも1つ一致すれば採用する方針）
    if (params.prefecture && f.prefecture !== params.prefecture) return;
    if (params.municipality && f.municipality !== params.municipality) return;
    if (params.q) {
      const q = String(params.q).toLowerCase();
      const hay = (f.name+' '+f.address+' '+(f.notes||'')+' '+(f.facilityType||'')).toLowerCase();
      if (!hay.includes(q)) return;
    }
    const addrKey = String(f.address || '').trim();
    // 空住所は個別扱い（IDをキーにして重複防止）
    const key = addrKey ? 'ADDR::'+addrKey : 'ID::'+f.id;
    if (!groups.has(key)) {
      groups.set(key, {
        address: f.address || '',
        prefecture: f.prefecture || '',
        municipality: f.municipality || '',
        names: new Set(),
        types: new Set(),
        phones: new Set(),
        notesArr: [],
        ids: new Set()
      });
    }
    const g = groups.get(key);
    g.names.add(f.name || '');
    if (f.facilityType) {
      // 複数種別が既にまとめられているケース（区切り）も分割統合
      String(f.facilityType).split(/[\/、,\s]+/).filter(Boolean).forEach(t=> g.types.add(t));
    }
    if (f.phone) g.phones.add(f.phone);
    if (f.notes) g.notesArr.push(f.notes);
    g.ids.add(f.id);
  });
  const out = [];
  groups.forEach(g => {
    // facilityType で絞り込み（グループ内に対象タイプが含まれないなら除外）
    if (params.facilityType) {
      if (![...g.types].some(t => t === params.facilityType)) return;
    }
    const nameJoined = [...g.names].slice(0,3).join('／');
    const typeJoined = [...g.types].sort().join('／');
    const phoneJoined = [...g.phones].join(' / ');
    // notes は代表的なものだけ（長文化防止）
    const notes = g.notesArr.length ? g.notesArr[0] : '';
    out.push({
      idList: [...g.ids],
      name: nameJoined,
      prefecture: g.prefecture,
      municipality: g.municipality,
      facilityType: typeJoined,
      address: g.address,
      phone: phoneJoined,
      notes: notes,
      facilityCount: g.ids.size
    });
  });
  // ソート：prefecture > municipality > address
  out.sort((a,b)=> {
    const kA = (a.prefecture||'')+'\t'+(a.municipality||'')+'\t'+(a.address||'');
    const kB = (b.prefecture||'')+'\t'+(b.municipality||'')+'\t'+(b.address||'');
    return kA.localeCompare(kB);
  });
  const limit = Number(params.limit || 0);
  if (limit > 0 && out.length > limit) return out.slice(0, limit);
  return out;
}

// 施設更新（id 必須）
function updateFacility(data) {
  if (!data || !data.id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) throw new Error('Facilities シートが見つかりません');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) throw new Error('更新対象がありません');
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  let target = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx.id]) === String(data.id)) { target = i; break; }
  }
  if (target === -1) throw new Error('指定IDの施設が見つかりません');
  // include contact in updatable columns
  const setters = ['name','prefecture','municipality','facilityType','address','phone','contact','notes'];
  setters.forEach(function(key){
    if (data[key] != null && idx[key] != null) {
      sheet.getRange(target+1, idx[key]+1).setValue(data[key]);
    }
  });
  return { id: data.id };
}

// 施設削除（関連する FacilityContacts をカスケード削除）
function deleteFacility(id) {
  if (!id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const fac = ss.getSheetByName('Facilities');
  if (!fac) throw new Error('Facilities シートが見つかりません');
  const fvals = fac.getDataRange().getValues();
  const fheaders = fvals[0];
  const fidx = {}; fheaders.forEach((h,i)=> fidx[h]=i);
  let frow = -1;
  for (let i = 1; i < fvals.length; i++) {
    if (String(fvals[i][fidx.id]) === String(id)) { frow = i; break; }
  }
  if (frow === -1) throw new Error('指定IDの施設が見つかりません');
  // 先に担当者を削除
  const con = ss.getSheetByName('FacilityContacts');
  if (con) {
    const cvals = con.getDataRange().getValues();
    const cheaders = cvals[0];
    const cidx = {}; cheaders.forEach((h,i)=> cidx[h]=i);
    for (let i = cvals.length - 1; i >= 1; i--) {
      if (String(cvals[i][cidx.facilityId]) === String(id)) {
        con.deleteRow(i + 1);
      }
    }
  }
  // 施設本体を削除
  fac.deleteRow(frow + 1);
  return { id: id };
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
    // ヘッダ順を固定: contact 列を notes の前に配置
    ss.insertSheet('Facilities').appendRow(['id','name','prefecture','municipality','facilityType','address','phone','contact','notes','createdAt','createdBy']);
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
    // パートナー営業用の追加列（不足のみ追加）
  let headers3 = cs.getRange(1,1,1,cs.getLastColumn()).getValues()[0];
  const needPartnerCols = ['partnerFlag','partnerType','stage','referralCount','lastReferralDate','partnerSince','nextAction','nextActionDate','lastActivityDate'];
    needPartnerCols.forEach(function(col){
      if (headers3.indexOf(col) === -1) {
        cs.insertColumnAfter(cs.getLastColumn());
        cs.getRange(1, cs.getLastColumn()).setValue(col);
        headers3 = cs.getRange(1,1,1,cs.getLastColumn()).getValues()[0];
      }
    });
  }
  // パートナー活動ログシート
  if (!names.includes('PartnerActivities')) {
    ss.insertSheet('PartnerActivities').appendRow(['id','contactId','activityDate','activityType','summary','nextActionDate','createdAt','createdBy']);
  }
  // 市区町村マスタ
  if (!names.includes('Municipalities')) {
    ss.insertSheet('Municipalities').appendRow(['prefecture','municipality']);
  }
  if (!names.includes(CONFIG_SHEET_NAME)) {
    ss.insertSheet(CONFIG_SHEET_NAME).appendRow(['key','value']);
  }
  // 関東（+山梨）市区町村の初期投入（空の場合のみ）
  try { seedMunicipalitiesIfEmpty(); } catch (e) { Logger.log('[seedMunicipalities][WARN] %s', e && e.message); }
  // 既存Facilitiesに新列が無ければ挿入（名前の後ろに順番）
  const facSheet = ss.getSheetByName('Facilities');
  if (facSheet) {
    const facHeaders = facSheet.getRange(1,1,1,facSheet.getLastColumn()).getValues()[0];
    const ensureCol = (colName, afterName) => {
      if (facHeaders.indexOf(colName) === -1) {
        const afterIdx = facHeaders.indexOf(afterName);
        if (afterIdx !== -1) {
          facSheet.insertColumnAfter(afterIdx + 1);
          facSheet.getRange(1, afterIdx + 2).setValue(colName);
        } else {
          facSheet.insertColumnAfter(facSheet.getLastColumn());
          facSheet.getRange(1, facSheet.getLastColumn()).setValue(colName);
        }
        // ヘッダ再取得
        facHeaders.splice(0, facHeaders.length, ...facSheet.getRange(1,1,1,facSheet.getLastColumn()).getValues()[0]);
      }
    };
    ensureCol('prefecture', 'name');
    ensureCol('municipality', 'prefecture');
    ensureCol('facilityType', 'municipality');
    // contact を phone の後、notes の前に確実に挿入
    ensureCol('contact', 'phone');
  }
  // 施設の旧 contact 列から FacilityContacts へ自動移行（必要な場合のみ）
  try { migrateFacilityContactsFromFacilities(); } catch (e) { Logger.log('[migrate][WARN] %s', e && e.message); }
}

// Municipalities シートがヘッダのみなら関東（+山梨）の市区町村を投入
function municipalitySeedMap_() {
  return {
    '東京都': [
      '千代田区','中央区','港区','新宿区','文京区','台東区','墨田区','江東区','品川区','目黒区','大田区','世田谷区','渋谷区','中野区','杉並区','豊島区','北区','荒川区','板橋区','練馬区','足立区','葛飾区','江戸川区',
      '八王子市','立川市','武蔵野市','三鷹市','青梅市','府中市','昭島市','調布市','町田市','小金井市','小平市','日野市','東村山市','国分寺市','国立市','福生市','狛江市','東大和市','清瀬市','東久留米市','武蔵村山市','多摩市','稲城市','羽村市','あきる野市','西東京市',
      '西多摩郡瑞穂町','西多摩郡日の出町','西多摩郡檜原村','西多摩郡奥多摩町',
      '大島町','利島村','新島村','神津島村','三宅村','御蔵島村','八丈町','青ヶ島村','小笠原村'
    ],
    '神奈川県': [
      '横浜市','川崎市','相模原市','横須賀市','平塚市','鎌倉市','藤沢市','小田原市','茅ヶ崎市','逗子市','三浦市','秦野市','厚木市','大和市','伊勢原市','海老名市','座間市','南足柄市','綾瀬市',
      '三浦郡葉山町',
      '高座郡寒川町',
      '中郡大磯町','中郡二宮町',
      '足柄上郡中井町','足柄上郡大井町','足柄上郡松田町','足柄上郡山北町','足柄上郡開成町',
      '足柄下郡箱根町','足柄下郡真鶴町','足柄下郡湯河原町',
      '愛甲郡愛川町','愛甲郡清川村'
    ],
    '千葉県': [
      '千葉市','銚子市','市川市','船橋市','館山市','木更津市','松戸市','野田市','茂原市','成田市','佐倉市','東金市','旭市','習志野市','柏市','勝浦市','市原市','流山市','八千代市','我孫子市','鴨川市','鎌ケ谷市','君津市','富津市','浦安市','四街道市','袖ケ浦市','八街市','印西市','白井市','富里市','南房総市','匝瑳市','香取市','山武市','いすみ市','大網白里市',
      '印旛郡酒々井町','印旛郡栄町',
      '香取郡神崎町','香取郡多古町','香取郡東庄町',
      '山武郡九十九里町','山武郡芝山町','山武郡横芝光町',
      '長生郡一宮町','長生郡睦沢町','長生郡長生村','長生郡白子町','長生郡長柄町','長生郡長南町',
      '夷隅郡大多喜町','夷隅郡御宿町',
      '安房郡鋸南町'
    ],
    '埼玉県': [
      'さいたま市','川越市','熊谷市','川口市','行田市','秩父市','所沢市','飯能市','加須市','本庄市','東松山市','春日部市','狭山市','羽生市','鴻巣市','深谷市','上尾市','草加市','越谷市','蕨市','戸田市','入間市','朝霞市','志木市','和光市','新座市','桶川市','久喜市','北本市','八潮市','富士見市','三郷市','蓮田市','坂戸市','幸手市','鶴ヶ島市','日高市','吉川市','ふじみ野市','白岡市',
      '北足立郡伊奈町',
      '入間郡三芳町','入間郡毛呂山町','入間郡越生町',
      '比企郡滑川町','比企郡嵐山町','比企郡小川町','比企郡川島町','比企郡吉見町','比企郡鳩山町','比企郡ときがわ町',
      '秩父郡横瀬町','秩父郡皆野町','秩父郡長瀞町','秩父郡小鹿野町','秩父郡東秩父村',
      '児玉郡上里町','児玉郡美里町','児玉郡神川町',
      '大里郡寄居町',
      '南埼玉郡宮代町',
      '北葛飾郡杉戸町','北葛飾郡松伏町'
    ],
    '茨城県': [
      '水戸市','日立市','土浦市','古河市','石岡市','結城市','龍ケ崎市','下妻市','常総市','常陸太田市','高萩市','北茨城市','笠間市','取手市','牛久市','つくば市','ひたちなか市','鹿嶋市','潮来市','守谷市','常陸大宮市','那珂市','筑西市','坂東市','稲敷市','かすみがうら市','桜川市','神栖市','行方市','鉾田市','つくばみらい市','小美玉市',
      '東茨城郡茨城町','東茨城郡大洗町','東茨城郡城里町',
      '那珂郡東海村',
      '久慈郡大子町',
      '稲敷郡阿見町','稲敷郡河内町',
      '結城郡八千代町',
      '猿島郡五霞町','猿島郡境町',
      '北相馬郡利根町'
    ],
    '栃木県': [
      '宇都宮市','足利市','栃木市','佐野市','鹿沼市','日光市','小山市','真岡市','大田原市','矢板市','那須塩原市','さくら市','那須烏山市','下野市',
      '河内郡上三川町',
      '芳賀郡益子町','芳賀郡茂木町','芳賀郡市貝町','芳賀郡芳賀町',
      '下都賀郡壬生町','下都賀郡野木町',
      '塩谷郡塩谷町','塩谷郡高根沢町',
      '那須郡那須町','那須郡那珂川町'
    ],
    '群馬県': [
      '前橋市','高崎市','桐生市','伊勢崎市','太田市','沼田市','館林市','渋川市','藤岡市','富岡市','安中市','みどり市',
      '北群馬郡榛東村','北群馬郡吉岡町',
      '多野郡上野村','多野郡神流町',
      '甘楽郡下仁田町','甘楽郡南牧村','甘楽郡甘楽町',
      '吾妻郡中之条町','吾妻郡長野原町','吾妻郡嬬恋村','吾妻郡草津町','吾妻郡高山村','吾妻郡東吾妻町',
      '利根郡片品村','利根郡川場村','利根郡昭和村','利根郡みなかみ町',
      '佐波郡玉村町',
      '邑楽郡板倉町','邑楽郡明和町','邑楽郡千代田町','邑楽郡大泉町','邑楽郡邑楽町'
    ],
    '山梨県': [
      '甲府市','富士吉田市','都留市','山梨市','大月市','韮崎市','南アルプス市','北杜市','甲斐市','笛吹市','上野原市','甲州市','中央市',
      '西八代郡市川三郷町',
      '南巨摩郡早川町','南巨摩郡身延町','南巨摩郡南部町','南巨摩郡富士川町',
      '中巨摩郡昭和町',
      '南都留郡道志村','南都留郡西桂町','南都留郡忍野村','南都留郡山中湖村','南都留郡鳴沢村','南都留郡富士河口湖町',
      '北都留郡小菅村','北都留郡丹波山村'
    ]
  };
}

function seedMunicipalitiesIfEmpty() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Municipalities');
  if (!sheet) return;
  if (sheet.getLastRow() > 1) return; // 既にデータあり
  const dataMap = municipalitySeedMap_();
  const rows = [];
  Object.keys(dataMap).forEach(pref => {
    dataMap[pref].forEach(m => rows.push([pref, m]));
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 2).setValues(rows);
    Logger.log('[seedMunicipalities] %s 行を投入', rows.length);
  }
}

// 既存データに対して欠けている市区町村だけを追加（再試行用）
function seedMunicipalitiesMerge() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Municipalities');
  if (!sheet) throw new Error('Municipalities シートが見つかりません');
  const current = sheet.getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < current.length; i++) {
    const r = current[i];
    if (!r[0] || !r[1]) continue;
    set.add(String(r[0]) + '||' + String(r[1]));
  }
  const map = municipalitySeedMap_();
  const rows = [];
  Object.keys(map).forEach(pref => {
    map[pref].forEach(m => {
      const key = pref + '||' + m;
      if (!set.has(key)) rows.push([pref, m]);
    });
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 2).setValues(rows);
  }
  return { inserted: rows.length };
}

// 市区町村マスタ取得（prefecture で絞り込み）
function getMunicipalities(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Municipalities');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = { prefecture: r[0], municipality: r[1] };
    if (params.prefecture && item.prefecture !== params.prefecture) continue;
    list.push(item);
  }
  // 重複除去・ソート
  const seen = new Set();
  const uniq = [];
  list.forEach(it => { const k = it.municipality; if (!seen.has(k)) { seen.add(k); uniq.push(it); } });
  uniq.sort((a,b)=> a.municipality.localeCompare(b.municipality));
  return uniq;
}

// 市区町村マスタに追加（単純追加・重複はスキップ）
function addMunicipality(data) {
  if (!data || !data.prefecture || !data.municipality) throw new Error('prefecture, municipality は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Municipalities');
  if (!sheet) throw new Error('Municipalities シートが見つかりません');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.prefecture) && String(rows[i][1]) === String(data.municipality)) {
      return { ok: true, existed: true };
    }
  }
  sheet.appendRow([data.prefecture, data.municipality]);
  return { ok: true, existed: false };
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
  // ヘッダ駆動で行を構築
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const row = new Array(headers.length).fill('');
  function set(h, v){ if (idx[h]!=null) row[idx[h]] = v; }
  set('id', id);
  set('name', data.name || '');
  set('prefecture', data.prefecture || '');
  set('municipality', data.municipality || '');
  set('facilityType', data.facilityType || '');
  set('address', data.address || '');
  set('phone', data.phone || '');
  set('contact', data.contact || '');
  set('notes', data.notes || '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  return { id, createdAt };
}

// 施設一覧取得（簡易構造）
function getFacilities(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  const limit = Number(params.limit || 0);
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    list.push({
      id: r[idx.id] || r[0],
      name: r[idx.name] || r[1],
      prefecture: idx.prefecture!=null ? r[idx.prefecture] : '',
      municipality: idx.municipality!=null ? r[idx.municipality] : '',
      facilityType: idx.facilityType!=null ? r[idx.facilityType] : '',
      address: r[idx.address] != null ? r[idx.address] : r[2],
      phone: r[idx.phone] != null ? r[idx.phone] : r[3],
      notes: idx.notes!=null ? r[idx.notes] : r[5]
    });
    if (limit > 0 && list.length >= limit) break;
  }
  return list;
}

// 問い合わせ記録追加（旧訪問記録を包含）
function addVisit(data) {
  if (!data || !data.facilityId) throw new Error('facilityId は必須です');
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Visits');
  if (!sheet) {
    sheet = ss.insertSheet('Visits');
    sheet.appendRow(['id','facilityId','visitDate','inquiryType','visitorName','visitorRelation','diagnosis','disabilityCategory','wantsGroupHome','wantsHomeNursing','visitPurpose','notes','createdAt','createdBy']);
  }
  // ヘッダ拡張（旧visitType互換、staffName削除／inquiryType採用）
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const needed = ['id','facilityId','visitDate','inquiryType','visitorName','visitorRelation','callerPhone','callerEmail','subjectName','subjectAge','subjectGender','diagnosis','disabilityCategory','careLevel','residenceMunicipality','wantsGroupHome','wantsHomeNursing','desiredStartDate','visitPurpose','urgency','preferredContactMethod','preferredContactTime','referralSource','consentFlag','notes','createdAt','createdBy'];
  needed.forEach(function(h){
    if (headers.indexOf(h) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(h);
      headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    }
  });
  // 旧 visitType 列が存在し inquiryType が空なら値移行（単回）
  const visitTypeIdx = headers.indexOf('visitType');
  const inquiryTypeIdx = headers.indexOf('inquiryType');
  if (visitTypeIdx !== -1 && inquiryTypeIdx !== -1) {
    const rng = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let migrated = false;
    for (let i=0;i<rng.length;i++) {
      const row = rng[i];
      if (row[inquiryTypeIdx] === '' && row[visitTypeIdx] !== '') {
        sheet.getRange(i+2, inquiryTypeIdx+1).setValue(row[visitTypeIdx]);
        migrated = true;
      }
    }
    if (migrated) Logger.log('[addVisit] migrated legacy visitType -> inquiryType');
  }
  const id = data.id || makeId('VIS');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  const visitDate = normalizeDate(data.visitDate, createdAt);
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const row = new Array(headers.length).fill('');
  function set(h,v){ if (idx[h]!=null) row[idx[h]] = v; }
  set('id', id);
  set('facilityId', data.facilityId);
  set('visitDate', visitDate);
  set('inquiryType', data.inquiryType || data.visitType || ''); // 電話 / 来訪 / 訪問 等
  set('visitorName', data.visitorName || '');
  set('visitorRelation', data.visitorRelation || '');
  set('callerPhone', data.callerPhone || '');
  set('callerEmail', data.callerEmail || '');
  set('subjectName', data.subjectName || '');
  set('subjectAge', data.subjectAge || '');
  set('subjectGender', data.subjectGender || '');
  set('diagnosis', data.diagnosis || '');
  set('disabilityCategory', data.disabilityCategory || '');
  set('careLevel', data.careLevel || '');
  set('residenceMunicipality', data.residenceMunicipality || '');
  set('wantsGroupHome', data.wantsGroupHome ? '1' : '');
  set('wantsHomeNursing', data.wantsHomeNursing ? '1' : '');
  set('desiredStartDate', data.desiredStartDate || '');
  set('visitPurpose', data.visitPurpose || '');
  set('urgency', data.urgency || '');
  set('preferredContactMethod', data.preferredContactMethod || '');
  set('preferredContactTime', data.preferredContactTime || '');
  set('referralSource', data.referralSource || '');
  set('consentFlag', data.consentFlag ? '1' : '');
  set('notes', data.notes || '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  return { id, createdAt };
}

function getVisits(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  for (let i=1;i<rows.length;i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id] || r[0],
      facilityId: r[idx.facilityId] || r[1],
      visitDate: r[idx.visitDate] || r[2],
      inquiryType: idx.inquiryType!=null ? r[idx.inquiryType] : (idx.visitType!=null ? r[idx.visitType] : ''),
      visitorName: idx.visitorName!=null ? r[idx.visitorName] : '',
      visitorRelation: idx.visitorRelation!=null ? r[idx.visitorRelation] : '',
      callerPhone: idx.callerPhone!=null ? r[idx.callerPhone] : '',
      callerEmail: idx.callerEmail!=null ? r[idx.callerEmail] : '',
      subjectName: idx.subjectName!=null ? r[idx.subjectName] : '',
      subjectAge: idx.subjectAge!=null ? r[idx.subjectAge] : '',
      subjectGender: idx.subjectGender!=null ? r[idx.subjectGender] : '',
      diagnosis: idx.diagnosis!=null ? r[idx.diagnosis] : '',
      disabilityCategory: idx.disabilityCategory!=null ? r[idx.disabilityCategory] : '',
      careLevel: idx.careLevel!=null ? r[idx.careLevel] : '',
      residenceMunicipality: idx.residenceMunicipality!=null ? r[idx.residenceMunicipality] : '',
      wantsGroupHome: idx.wantsGroupHome!=null ? r[idx.wantsGroupHome] : '',
      wantsHomeNursing: idx.wantsHomeNursing!=null ? r[idx.wantsHomeNursing] : '',
      desiredStartDate: idx.desiredStartDate!=null ? r[idx.desiredStartDate] : '',
      visitPurpose: idx.visitPurpose!=null ? r[idx.visitPurpose] : '',
      urgency: idx.urgency!=null ? r[idx.urgency] : '',
      preferredContactMethod: idx.preferredContactMethod!=null ? r[idx.preferredContactMethod] : '',
      preferredContactTime: idx.preferredContactTime!=null ? r[idx.preferredContactTime] : '',
      referralSource: idx.referralSource!=null ? r[idx.referralSource] : '',
      consentFlag: idx.consentFlag!=null ? r[idx.consentFlag] === '1' : false,
      notes: idx.notes!=null ? r[idx.notes] : '',
      createdAt: idx.createdAt!=null ? r[idx.createdAt] : '',
      createdBy: idx.createdBy!=null ? r[idx.createdBy] : ''
    };
    if (params.facilityId && String(item.facilityId) !== String(params.facilityId)) continue;
    if (params.inquiryType && params.inquiryType !== '') {
      if (String(item.inquiryType) !== String(params.inquiryType)) continue;
    }
    if (params.disabilityCategory && params.disabilityCategory !== '') {
      if (!item.disabilityCategory || String(item.disabilityCategory) !== String(params.disabilityCategory)) continue;
    }
    list.push(item);
  }
  list.sort((a,b)=> (b.visitDate||'').localeCompare(a.visitDate||''));
  return list;
}

// 施設の活動サマリー（訪問回数/最終訪問日、報告回数/最終報告日）
function getFacilityActivitySummary(facilityId) {
  if (!facilityId) return { visitCount: 0, lastVisitDate: '', reportCount: 0, lastReportDate: '' };
  const ss = getSpreadsheet();
  // Visits（問い合わせ記録）
  let visitCount = 0, lastVisitDate = '';
  const vs = ss.getSheetByName('Visits');
  if (vs) {
    const rows = vs.getDataRange().getValues();
    if (rows.length > 1) {
      const headers = rows[0];
      const vidx = {}; headers.forEach((h,i)=> vidx[h]=i);
      for (let i=1;i<rows.length;i++) {
        const r = rows[i];
        const fid = vidx.facilityId!=null ? r[vidx.facilityId] : r[1];
        if (String(fid) === String(facilityId)) {
          visitCount++;
          const d = vidx.visitDate!=null ? r[vidx.visitDate] : r[2];
          if (d && String(d) > String(lastVisitDate)) lastVisitDate = String(d);
        }
      }
    }
  }
  // Reports
  let reportCount = 0, lastReportDate = '';
  const rs = ss.getSheetByName('Reports');
  if (rs) {
    const rows = rs.getDataRange().getValues();
    const headers = rows[0] || [];
    const idx = {}; headers.forEach((h,i)=> idx[h]=i);
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const fid = idx.facilityId != null ? r[idx.facilityId] : r[1];
      if (String(fid) === String(facilityId)) {
        reportCount++;
        const d = (idx.reportDate != null ? r[idx.reportDate] : r[2]) || (idx.createdAt != null ? r[idx.createdAt] : r[8]);
        if (d && String(d) > String(lastReportDate)) lastReportDate = String(d);
      }
    }
  }
  return { visitCount, lastVisitDate, reportCount, lastReportDate };
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
    // フォローアップ強化用の列を追加
    let h2 = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const need = ['nextAction','nextActionDate','nextActionOwner','priority','statusStage','reminderFlag'];
    need.forEach(function(col){
      if (h2.indexOf(col) === -1) {
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue(col);
        h2 = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      }
    });
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
  set('nextAction', data.nextAction || '');
  set('nextActionDate', data.nextActionDate || '');
  set('nextActionOwner', data.nextActionOwner || '');
  set('priority', data.priority || '');
  set('statusStage', data.statusStage || '');
  set('reminderFlag', data.reminderFlag ? '1' : '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  // Webhook 通知（LINE WORKS など）
  try {
    const facilityName = (function(){
      try { const fac = getFacilities().find(f=> String(f.id) === String(data.facilityId)); return fac ? fac.name : ''; } catch(e){ return ''; }
    })();
    const msg = '【営業報告】\n' +
      (facilityName ? '施設: ' + facilityName + '\n' : '') +
      (data.reportDate ? '日付: ' + String(data.reportDate) + '\n' : '') +
      (data.reporterName ? '報告者: ' + String(data.reporterName) + '\n' : '') +
      (data.summary ? '概要: ' + String(data.summary) + '\n' : '') +
      (data.followUp ? 'フォロー: ' + String(data.followUp) + '\n' : '');
    try { sendLineWorksWebhook(msg); } catch (e) { Logger.log('[sendLineWorksWebhook][ERROR] %s', e && e.message); }
  } catch (e) {
    Logger.log('[addReport][WEBHOOK-ERR] %s', e && e.message);
  }

  return { id, createdAt };
}

// ----- LINE WORKS Webhook helper -----
function setLineWorksWebhookUrl(url) {
  if (!isAdminUser_()) throw new Error('権限がありません');
  if (!url) throw new Error('url は必須です');
  writeConfigValue('LINE_WORKS_WEBHOOK_URL', String(url));
  return { ok: true };
}

function getLineWorksWebhookUrl() {
  return readConfigValue('LINE_WORKS_WEBHOOK_URL') || '';
}

function sendLineWorksWebhook(message) {
  try {
    const url = getLineWorksWebhookUrl();
    if (!url) { Logger.log('[sendLineWorksWebhook] no webhook url configured'); return { ok: false, reason: 'no_url' }; }
    const payload = { content: message };
    const opts = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    Logger.log('[sendLineWorksWebhook] code=%s body=%s', code, body);
    return { ok: code >= 200 && code < 300, code: code, body: body };
  } catch (e) {
    Logger.log('[sendLineWorksWebhook][ERROR] %s', e && e.message);
    return { ok: false, error: e && e.message };
  }
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
      contactName: idx.contactName != null ? r[idx.contactName] : '',
      nextAction: idx.nextAction != null ? r[idx.nextAction] : '',
      nextActionDate: idx.nextActionDate != null ? r[idx.nextActionDate] : '',
      nextActionOwner: idx.nextActionOwner != null ? r[idx.nextActionOwner] : '',
      priority: idx.priority != null ? r[idx.priority] : '',
      statusStage: idx.statusStage != null ? r[idx.statusStage] : '',
      reminderFlag: idx.reminderFlag != null ? r[idx.reminderFlag] : ''
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
  const limit = Number(params.limit || 0);
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id] || r[0],
      facilityId: r[idx.facilityId] || r[1],
      contactName: r[idx.contactName] || r[2],
      contactPhone: idx.contactPhone!=null ? r[idx.contactPhone] : '',
      contactNotes: idx.contactNotes!=null ? r[idx.contactNotes] : '',
      cardFileId: idx.cardFileId!=null ? r[idx.cardFileId] : '',
  stage: idx.stage!=null ? r[idx.stage] : '',
      partnerFlag: idx.partnerFlag!=null ? (r[idx.partnerFlag] === '1') : false,
      partnerType: idx.partnerType!=null ? r[idx.partnerType] : '',
      referralCount: idx.referralCount!=null ? Number(r[idx.referralCount]||0) : 0,
      lastReferralDate: idx.lastReferralDate!=null ? r[idx.lastReferralDate] : '',
      partnerSince: idx.partnerSince!=null ? r[idx.partnerSince] : '',
      nextAction: idx.nextAction!=null ? r[idx.nextAction] : '',
      nextActionDate: idx.nextActionDate!=null ? r[idx.nextActionDate] : '',
      lastActivityDate: idx.lastActivityDate!=null ? r[idx.lastActivityDate] : '',
      createdAt: r[idx.createdAt] || r[4],
      createdBy: r[idx.createdBy] || r[5]
    };
    if (params.facilityId && item.facilityId !== params.facilityId) continue;
    if (params.partnerFlag === true && !item.partnerFlag) continue;
    list.push(item);
    if (limit > 0 && list.length >= limit) break;
  }
  return list;
}

// パートナー（紹介元）抽出。Facility名も付与。
function getPartnerContacts(params) {
  params = params || {}; params.partnerFlag = true;
  const contacts = getFacilityContacts(params);
  if (!contacts.length) return [];
  const facilities = getFacilities();
  const map = {}; facilities.forEach(f=> map[f.id]=f.name);
  return contacts.map(c => ({
    id: c.id,
    facilityId: c.facilityId,
    facilityName: map[c.facilityId] || '',
    contactName: c.contactName,
    partnerType: c.partnerType,
    stage: c.stage,
    referralCount: c.referralCount,
    lastReferralDate: c.lastReferralDate,
    nextAction: c.nextAction,
    nextActionDate: c.nextActionDate,
    lastActivityDate: c.lastActivityDate,
    partnerSince: c.partnerSince
  }));
}

// パートナー設定/更新
function updatePartnerContact(data) {
  if (!data || !data.id) throw new Error('id は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('FacilityContacts');
  if (!sheet) throw new Error('FacilityContacts シートがありません');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) throw new Error('更新対象がありません');
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  let target = -1;
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][idx.id]) === String(data.id)) { target = i; break; }
  }
  if (target === -1) throw new Error('指定IDが見つかりません');
  function set(col, val){ if (idx[col]!=null) sheet.getRange(target+1, idx[col]+1).setValue(val); }
  if (data.partnerFlag != null) set('partnerFlag', data.partnerFlag ? '1' : '');
  if (data.partnerType != null) set('partnerType', data.partnerType || '');
  if (data.stage != null) set('stage', data.stage || '');
  if (data.nextAction != null) set('nextAction', data.nextAction || '');
  if (data.nextActionDate != null) set('nextActionDate', data.nextActionDate || '');
  // partnerSince 初期設定
  if (data.partnerFlag && idx.partnerSince!=null) {
    const cur = rows[target][idx.partnerSince];
    if (!cur) sheet.getRange(target+1, idx.partnerSince+1).setValue(nowIso());
  }
  return { id: data.id };
}

// パートナー活動ログ追加
function addPartnerActivity(data) {
  if (!data || !data.contactId) throw new Error('contactId は必須です');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('PartnerActivities');
  if (!sheet) throw new Error('PartnerActivities シートがありません');
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const id = makeId('PACT');
  const createdAt = nowIso();
  const createdBy = activeUserEmail();
  const activityDate = normalizeDate(data.activityDate, createdAt);
  const row = new Array(headers.length).fill('');
  function set(h,v){ if (idx[h]!=null) row[idx[h]] = v; }
  set('id', id);
  set('contactId', data.contactId);
  set('activityDate', activityDate);
  set('activityType', data.activityType || '');
  set('summary', data.summary || '');
  set('nextActionDate', data.nextActionDate || '');
  set('createdAt', createdAt);
  set('createdBy', createdBy);
  sheet.appendRow(row);
  // FacilityContacts 側の lastActivityDate 更新 & nextActionDate 反映
  const fcSheet = ss.getSheetByName('FacilityContacts');
  if (fcSheet) {
    const fcRows = fcSheet.getDataRange().getValues();
    const h2 = fcRows[0];
    const ix = {}; h2.forEach((h,i)=> ix[h]=i);
    for (let i=1;i<fcRows.length;i++) {
      if (String(fcRows[i][ix.id]) === String(data.contactId)) {
        if (ix.lastActivityDate!=null) fcSheet.getRange(i+1, ix.lastActivityDate+1).setValue(activityDate);
        if (data.nextActionDate && ix.nextActionDate!=null) fcSheet.getRange(i+1, ix.nextActionDate+1).setValue(data.nextActionDate);
        break;
      }
    }
  }
  return { id, createdAt };
}

// パートナー活動取得（contactId で絞り込み、指定なしで全件）
function getPartnerActivities(params) {
  params = params || {};
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('PartnerActivities');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const list = [];
  for (let i=1;i<rows.length;i++) {
    const r = rows[i];
    const item = {
      id: r[idx.id],
      contactId: r[idx.contactId],
      activityDate: r[idx.activityDate],
      activityType: r[idx.activityType],
      summary: r[idx.summary],
      nextActionDate: idx.nextActionDate!=null ? r[idx.nextActionDate] : '',
      createdAt: r[idx.createdAt],
      createdBy: r[idx.createdBy]
    };
    if (params.contactId && String(item.contactId) !== String(params.contactId)) continue;
    list.push(item);
  }
  list.sort((a,b)=> (b.activityDate||'').localeCompare(a.activityDate||''));
  return list;
}

// 期限が近いパートナーアクション（nextActionDate が今日から horizon 日以内）
function getUpcomingPartnerActions(horizonDays) {
  const days = Number(horizonDays || 14);
  const now = new Date();
  const contacts = getFacilityContacts({ partnerFlag: true });
  const upcoming = [];
  contacts.forEach(c => {
    if (!c.nextActionDate) return;
    const dt = new Date(c.nextActionDate);
    if (isNaN(dt.getTime())) return;
    const diff = (dt.getTime() - now.getTime()) / 86400000;
    if (diff >= -0.5 && diff <= days) {
      upcoming.push({
        id: c.id,
        facilityId: c.facilityId,
        contactName: c.contactName,
        partnerType: c.partnerType,
        nextAction: c.nextAction,
        nextActionDate: c.nextActionDate
      });
    }
  });
  upcoming.sort((a,b)=> (a.nextActionDate||'').localeCompare(b.nextActionDate||''));
  return upcoming;
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

// ===== Facilities 住所単位 破壊的マージ（C） =====
function isAdminUser_() {
  try {
    const email = String(activeUserEmail() || '').toLowerCase();
    const prop = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS');
    if (prop && String(prop).trim()) {
      const allowed = String(prop).split(/[\s,;]+/).map(s=>String(s).toLowerCase()).filter(Boolean);
      return email && allowed.indexOf(email) !== -1;
    }
    // 設定が無い場合は許可（開発初期運用）
    return true;
  } catch (e) {
    return true; // 取得不可時は許可（必要に応じて false に変更）
  }
}
// options: {
//   keepRepresentativeId: true,     // 代表行（最初の行）の id を残す
//   joinPhones: true,                // phone をユニークに結合（' / '）
//   maxNames: 3,                     // name の代表表示に含める件数
//   joinAllNotes: false              // notes を全結合するか（false の場合は代表行の notes を維持）
// }
function mergeFacilitiesByAddressDestructive(options) {
  options = options || {};
  const keepRepresentativeId = options.keepRepresentativeId !== false; // default true
  const joinPhones = options.joinPhones !== false; // default true
  const maxNames = Math.max(1, Number(options.maxNames || 3));
  const joinAllNotes = options.joinAllNotes === true; // default false

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Facilities');
  if (!sheet) throw new Error('Facilities シートが見つかりません');

  // バックアップ作成
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupName = 'Facilities_backup_' + stamp;
  const backup = ss.insertSheet(backupName);
  const values = sheet.getDataRange().getValues();
  if (values.length) backup.getRange(1,1,values.length,values[0].length).setValues(values);

  if (values.length <= 1) return { backupSheet: backupName, mergedGroups: 0, deletedRows: 0, updatedRows: 0 };

  const headers = values[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);
  const idCol = idx.id != null ? idx.id : 0;
  const nameCol = idx.name != null ? idx.name : 1;
  const prefectureCol = idx.prefecture;
  const municipalityCol = idx.municipality;
  const typeCol = idx.facilityType;
  const addressCol = idx.address != null ? idx.address : 2;
  const phoneCol = idx.phone != null ? idx.phone : 3;
  const notesCol = idx.notes != null ? idx.notes : null;

  // グルーピング（住所完全一致、空白はトリム）
  const groups = new Map(); // key -> array of row indices (1-based for sheet, excluding header)
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const addr = String(r[addressCol] || '').trim();
    const key = addr ? addr : '__EMPTY__::' + (r[idCol] || i);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i + 1); // シート上は1始まり
  }

  let mergedGroups = 0;
  let deletedRows = 0;
  let updatedRows = 0;
  const rowsToDelete = [];

  // 住所が空のグループ（__EMPTY__）は事実上1件ずつなので処理不要。2件以上あってもマージしない方針。
  groups.forEach((rowIdxList, key) => {
    if (key.startsWith('__EMPTY__')) return; // スキップ
    if (!rowIdxList || rowIdxList.length <= 1) return;
    // 代表は最初の行（小さい行番号）。
    rowIdxList.sort((a,b)=> a-b);
    const repRow = rowIdxList[0];
    const others = rowIdxList.slice(1);
    // 集計
    const nameSet = new Set();
    const typeSet = new Set();
    const phoneSet = new Set();
    let repNotes = '';

    rowIdxList.forEach(rn => {
      const r = sheet.getRange(rn, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (nameCol != null && r[nameCol]) nameSet.add(String(r[nameCol]));
      if (typeCol != null && r[typeCol]) {
        String(r[typeCol]).split(/[\/、,\s]+/).filter(Boolean).forEach(t => typeSet.add(t));
      }
      if (joinPhones && phoneCol != null && r[phoneCol]) phoneSet.add(String(r[phoneCol]));
      if (!joinAllNotes && rn === repRow && notesCol != null) repNotes = r[notesCol] || repNotes;
    });
    const repRange = sheet.getRange(repRow, 1, 1, sheet.getLastColumn());
    const rep = repRange.getValues()[0];
    // 更新内容を構築
    const newName = [...nameSet].slice(0, maxNames).join('／') || rep[nameCol];
    const newType = [...typeSet].join('／');
    const newPhone = joinPhones ? [...phoneSet].join(' / ') : rep[phoneCol];
    const newNotes = joinAllNotes && notesCol != null ? (function(){
      const collected = [];
      rowIdxList.forEach(rn => {
        const r = sheet.getRange(rn,1,1,sheet.getLastColumn()).getValues()[0];
        if (notesCol != null && r[notesCol]) collected.push(String(r[notesCol]));
      });
      return collected.join(' / ');
    })() : repNotes;

    // シートへ反映（代表行）。id はそのまま。
    if (nameCol != null) rep[nameCol] = newName;
    if (typeCol != null) rep[typeCol] = newType;
    if (phoneCol != null && newPhone !== undefined) rep[phoneCol] = newPhone;
    if (notesCol != null && newNotes !== undefined) rep[notesCol] = newNotes;
    repRange.setValues([rep]);
    updatedRows++;

    // 代表以外を削除予定に追加
    rowsToDelete.push(...others);
    mergedGroups++;
  });

  // 行削除（降順）
  rowsToDelete.sort((a,b)=> b-a);
  rowsToDelete.forEach(rn => { sheet.deleteRow(rn); deletedRows++; });

  return {
    backupSheet: backupName,
    mergedGroups: mergedGroups,
    deletedRows: deletedRows,
    updatedRows: updatedRows
  };
}

// 既定オプションで実行するラッパー（clasp run用）
function runMergeFacilitiesByAddress() {
  if (!isAdminUser_()) throw new Error('権限がありません（ADMIN_EMAILS 未許可）');
  return mergeFacilitiesByAddressDestructive({
    keepRepresentativeId: true,
    joinPhones: true,
    maxNames: 3,
    joinAllNotes: false
  });
}

function getConfigSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(['key','value']);
  }
  return sheet;
}

function readConfigValue(key) {
  if (!key) return null;
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(key)) {
      return rows[i][1];
    }
  }
  return null;
}

function writeConfigValue(key, value) {
  if (!key) throw new Error('key は必須です');
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}
