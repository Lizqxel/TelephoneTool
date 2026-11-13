/**
 * TelephoneTool performance endpoint (independent from the existing Form flow).
 *
 * Script properties required before setupPerformanceTracking():
 *   SPREADSHEET_ID = target spreadsheet id
 *   SHARED_TOKEN   = long random token used by TelephoneTool
 *   ACTIVE         = true (setup sets this automatically)
 */

const TT_STATS_SHEET = 'TT_実績受信';
const TT_CORRECTIONS_SHEET = 'TT_実績修正';
const TT_DETAILS_SHEET = 'TT_架電明細';
const TT_SUMMARY_SHEET = 'TT_時間帯集計';

const TT_STATS_HEADERS = [
  '一意キー', '対象日', '時間帯', 'スタッフID', '担当者名',
  'コール', '対話', 'トス', '対話秒数', '端末ID', '更新日時'
];
const TT_CORRECTION_HEADERS = [
  '対象日', '時間帯', 'スタッフID', 'コール補正', '対話補正',
  'トス補正', '対話秒数補正', '修正理由', '修正者', '有効'
];
const TT_DETAIL_HEADERS = [
  'イベントID', '対象日', '時間帯', 'スタッフID', '担当者名',
  '架電日時', '架電結果', '対話', 'トス', 'IP電話通話秒数', '更新日時'
];
const TT_SUMMARY_HEADERS = [
  '対象日', '時間帯', 'スタッフID', '担当者名', 'コール', '対話',
  'トス', '対話秒数', '管理者修正', '更新日時'
];

function setupPerformanceTracking() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const token = props.getProperty('SHARED_TOKEN');
  if (!spreadsheetId || !token) {
    throw new Error('SPREADSHEET_ID と SHARED_TOKEN をスクリプトプロパティへ設定してください。');
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  ensureSheet_(spreadsheet, TT_STATS_SHEET, TT_STATS_HEADERS);
  ensureSheet_(spreadsheet, TT_CORRECTIONS_SHEET, TT_CORRECTION_HEADERS);
  ensureSheet_(spreadsheet, TT_DETAILS_SHEET, TT_DETAIL_HEADERS);
  ensureSheet_(spreadsheet, TT_SUMMARY_SHEET, TT_SUMMARY_HEADERS);
  props.setProperty('ACTIVE', 'true');
  return 'TelephoneTool実績シートを作成しました。';
}

function backupPerformanceSheets() {
  const spreadsheet = getSpreadsheet_();
  const suffix = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  [TT_STATS_SHEET, TT_CORRECTIONS_SHEET, TT_DETAILS_SHEET, TT_SUMMARY_SHEET].forEach(name => {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) sheet.copyTo(spreadsheet).setName(`${name}_backup_${suffix}`);
  });
}

function setPerformanceEnabled(enabled) {
  PropertiesService.getScriptProperties().setProperty('ACTIVE', enabled ? 'true' : 'false');
}

function doPost(e) {
  try {
    assertActive_();
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    assertToken_(body.token);
    if (body.action === 'query') {
      const day = String(body.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
      return json_({ok: true, dashboard: buildDashboard_(day)});
    }
    if (body.action !== 'sync') throw new Error('未対応のactionです。');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const spreadsheet = getSpreadsheet_();
      const statsSheet = ensureSheet_(spreadsheet, TT_STATS_SHEET, TT_STATS_HEADERS);
      const detailSheet = ensureSheet_(spreadsheet, TT_DETAILS_SHEET, TT_DETAIL_HEADERS);
      const clientId = String(body.clientId || '');
      const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");

      const statsRows = (body.stats || []).map(item => {
        const day = String(item.business_date || body.businessDate || '');
        const hour = number_(item.hour);
        const staffId = String(item.staff_id || '');
        return [
          `${day}|${hour}|${staffId}`, day, hour, staffId, String(item.staff_name || ''),
          number_(item.calls), number_(item.dialogues), number_(item.tosses),
          number_(item.talk_seconds), clientId, now
        ];
      }).filter(row => row[1] && row[3]);

      const detailRows = (body.details || []).map(item => [
        String(item.event_id || ''), String(item.business_date || body.businessDate || ''),
        number_(item.hour), String(item.staff_id || ''), String(item.staff_name || ''),
        String(item.called_at || ''), String(item.result || ''), item.is_dialogue ? 1 : 0,
        item.is_toss ? 1 : 0, number_(item.talk_seconds), now
      ]).filter(row => row[0] && row[1] && row[3]);

      upsertRows_(statsSheet, statsRows, 0);
      upsertRows_(detailSheet, detailRows, 0);
      writeSummary_(spreadsheet, buildDashboard_(String(body.businessDate || '')));
      return json_({ok: true, stats: statsRows.length, details: detailRows.length});
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return json_({ok: false, error: String(error && error.message || error)});
  }
}

function doGet(e) {
  try {
    assertActive_();
    assertToken_(e.parameter.token);
    const day = String(e.parameter.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
    return json_({ok: true, dashboard: buildDashboard_(day)});
  } catch (error) {
    return json_({ok: false, error: String(error && error.message || error)});
  }
}

function buildDashboard_(day) {
  const spreadsheet = getSpreadsheet_();
  const statsSheet = ensureSheet_(spreadsheet, TT_STATS_SHEET, TT_STATS_HEADERS);
  const correctionsSheet = ensureSheet_(spreadsheet, TT_CORRECTIONS_SHEET, TT_CORRECTION_HEADERS);
  const detailsSheet = ensureSheet_(spreadsheet, TT_DETAILS_SHEET, TT_DETAIL_HEADERS);
  const stats = dataRows_(statsSheet).filter(row => dateText_(row[1]) === day);
  const corrections = dataRows_(correctionsSheet).filter(row => {
    const enabled = String(row[9] === undefined ? '' : row[9]).toLowerCase();
    return dateText_(row[0]) === day && enabled !== 'false' && enabled !== '無効' && enabled !== '0';
  });
  const details = dataRows_(detailsSheet).filter(row => dateText_(row[1]) === day);
  const byStaff = {};

  stats.forEach(row => {
    const staffId = String(row[3] || '');
    if (!staffId) return;
    const staff = byStaff[staffId] || (byStaff[staffId] = newStaff_(staffId, String(row[4] || '')));
    const hour = number_(row[2]);
    const hourRow = staff.hours.find(item => item.hour === hour) || addHour_(staff, hour);
    hourRow.calls = number_(row[5]);
    hourRow.dialogues = number_(row[6]);
    hourRow.tosses = number_(row[7]);
    hourRow.talk_seconds = number_(row[8]);
  });

  corrections.forEach(row => {
    const staffId = String(row[2] || '');
    if (!staffId) return;
    const staff = byStaff[staffId] || (byStaff[staffId] = newStaff_(staffId, staffId));
    const hour = number_(row[1]);
    const hourRow = staff.hours.find(item => item.hour === hour) || addHour_(staff, hour);
    hourRow.calls += number_(row[3]);
    hourRow.dialogues += number_(row[4]);
    hourRow.tosses += number_(row[5]);
    hourRow.talk_seconds += number_(row[6]);
    hourRow.calls = Math.max(0, hourRow.calls);
    hourRow.dialogues = Math.max(0, hourRow.dialogues);
    hourRow.tosses = Math.max(0, hourRow.tosses);
    hourRow.talk_seconds = Math.max(0, hourRow.talk_seconds);
    staff.corrected = true;
  });

  details.forEach(row => {
    const staffId = String(row[3] || '');
    const staff = byStaff[staffId];
    if (!staff) return;
    staff.details.push({
      event_id: String(row[0] || ''), called_at: String(row[5] || ''),
      hour: number_(row[2]), result: String(row[6] || ''),
      is_dialogue: number_(row[7]) > 0, is_toss: number_(row[8]) > 0,
      talk_seconds: number_(row[9])
    });
  });

  const totals = {calls: 0, dialogues: 0, tosses: 0, talk_seconds: 0};
  const staff = Object.keys(byStaff).map(key => {
    const item = byStaff[key];
    item.hours.sort((a, b) => a.hour - b.hour);
    item.details.sort((a, b) => a.called_at.localeCompare(b.called_at));
    ['calls', 'dialogues', 'tosses', 'talk_seconds'].forEach(metric => {
      item[metric] = item.hours.reduce((sum, hour) => sum + number_(hour[metric]), 0);
      totals[metric] += item[metric];
    });
    return item;
  }).sort((a, b) => a.staff_name.localeCompare(b.staff_name, 'ja'));
  return {date: day, staff: staff, totals: totals, source: 'google-sheets'};
}

function writeSummary_(spreadsheet, dashboard) {
  const sheet = ensureSheet_(spreadsheet, TT_SUMMARY_SHEET, TT_SUMMARY_HEADERS);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  const existing = dataRows_(sheet);
  const retained = existing.filter(row => dateText_(row[0]) !== dashboard.date);
  const current = [];
  (dashboard.staff || []).forEach(staff => {
    (staff.hours || []).forEach(hour => current.push([
      dashboard.date, number_(hour.hour), staff.staff_id, staff.staff_name,
      number_(hour.calls), number_(hour.dialogues), number_(hour.tosses),
      number_(hour.talk_seconds), staff.corrected ? 'あり' : '', now
    ]));
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, TT_SUMMARY_HEADERS.length).setValues([TT_SUMMARY_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, TT_SUMMARY_HEADERS.length).setFontWeight('bold').setBackground('#D9EAF7');
  const rows = retained.concat(current);
  if (rows.length) sheet.getRange(2, 1, rows.length, TT_SUMMARY_HEADERS.length).setValues(rows);
}

function newStaff_(id, name) {
  return {staff_id: id, staff_name: name, calls: 0, dialogues: 0, tosses: 0,
    talk_seconds: 0, corrected: false, hours: [], details: []};
}

function addHour_(staff, hour) {
  const item = {hour: hour, calls: 0, dialogues: 0, tosses: 0, talk_seconds: 0};
  staff.hours.push(item);
  return item;
}

function upsertRows_(sheet, rows, keyColumn) {
  if (!rows.length) return;
  const existing = dataRows_(sheet);
  const index = {};
  existing.forEach((row, offset) => { index[String(row[keyColumn] || '')] = offset + 2; });
  const append = [];
  rows.forEach(row => {
    const key = String(row[keyColumn] || '');
    if (index[key]) sheet.getRange(index[key], 1, 1, row.length).setValues([row]);
    else append.push(row);
  });
  if (append.length) sheet.getRange(sheet.getLastRow() + 1, 1, append.length, append[0].length).setValues(append);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D9EAF7');
  }
  return sheet;
}

function dataRows_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_IDが未設定です。');
  return SpreadsheetApp.openById(id);
}

function assertToken_(value) {
  const expected = PropertiesService.getScriptProperties().getProperty('SHARED_TOKEN');
  if (!expected || String(value || '') !== expected) throw new Error('認証に失敗しました。');
}

function assertActive_() {
  if (PropertiesService.getScriptProperties().getProperty('ACTIVE') === 'false') {
    throw new Error('実績同期は停止中です。');
  }
}

function number_(value) {
  const result = Number(value || 0);
  return Number.isFinite(result) ? Math.trunc(result) : 0;
}

function dateText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10).replace(/\//g, '-');
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

