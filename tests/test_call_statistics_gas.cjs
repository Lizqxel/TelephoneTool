// Googleへ接続せず、GAS本体をV8相当のVMとシート/フォームのモックで検証する。
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../gas/call_statistics.gs'), 'utf8');
const integration = fs.readFileSync(path.join(__dirname, '../gas/auto_refresh_integration.gs'), 'utf8');
const legacySource = fs.readFileSync(path.join(__dirname, '../gas/reference/legacy_receiver.gs.txt'), 'utf8');
const clone = value => JSON.parse(JSON.stringify(value));

class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = clone(rows); this.hidden = false; this.maxRows = 10; this.failWrites = false; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map(row => row.length)); }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(after, count) { this.maxRows += count; }
  setFrozenRows(count) { this.frozen = count; }
  hideSheet() { this.hidden = true; }
  isSheetHidden() { return this.hidden; }
  getRange(row, column, height, width) {
    const sheet = this;
    return {
      getValues() { return Array.from({length: height}, (_, r) => Array.from({length: width}, (_, c) => sheet.rows[row - 1 + r]?.[column - 1 + c] ?? '')); },
      getDisplayValues() { return this.getValues().map(row => row.map(String)); },
      setNumberFormat() { return this; },
      setValues(values) {
        if (sheet.failWrites) throw new Error('temporary write failure');
        assert.equal(values.length, height);
        for (let r = 0; r < height; r++) {
          assert.equal(values[r].length, width);
          sheet.rows[row - 1 + r] ||= [];
          for (let c = 0; c < width; c++) sheet.rows[row - 1 + r][column - 1 + c] = values[r][c];
        }
        return this;
      }
    };
  }
}

function environment(script = source) {
  const legacy = new Sheet('コール集計打ち込み用_11月', [
    ['', '2026年11月'], ['', '【11/13（金）】'],
    ['', '名前', '', '', '日計', '', '', '10時台', '', '', '11時台', '', ''],
    ['', '', '部署', 'GPユニット', 'コール', '対話', 'トス', 'コール', '対話', 'トス', 'コール', '対話', 'トス'],
    ['', '姓 名', '部署A', 'ユニットA', 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['', '別人', '部署B', 'ユニットB', 0, 0, 0, 0, 0, 0, 0, 0, 0]
  ]);
  const sheets = new Map([[legacy.name, legacy]]);
  const spreadsheet = {getSheetByName: name => sheets.get(name), insertSheet(name) {
    assert(!sheets.has(name), 'duplicate sheet creation');
    const sheet = new Sheet(name); sheets.set(name, sheet); return sheet;
  }};
  const form = {responses: [], deleted: [], getResponses() { return this.responses; }, deleteResponse(id) { this.deleted.push(id); this.responses = this.responses.filter(response => response.getId() !== id); }};
  for (const name of ['setCollectEmail', 'setLimitOneResponsePerUser', 'setAllowResponseEdits', 'setShowLinkToRespondAgain', 'setRequireLogin', 'setAcceptingResponses']) form[name] = () => {};
  form.getItems = () => [{asTextItem: () => ({createResponse: () => ({})})}, {asParagraphTextItem: () => ({createResponse: () => ({})})}];
  form.createResponse = () => ({withItemResponse() { return this; }, toPrefilledUrl() { return 'https://example.invalid?entry.1=TOKENVALUE&entry.2=DATAVALUE'; }});
  form.getPublishedUrl = () => 'https://example.invalid/viewform';
  form.requiresLogin = () => false;
  const state = {locked: false, canLock: true};
  const logs = [];
  const properties = {TOKEN: 'test-token', FORM_ID: 'form'};
  const triggers = [{getHandlerFunction: () => 'processSubmissions'}];
  const createdTriggers = [];
  const scriptApp = {
    getScriptId: () => 'receiver-test-project',
    getProjectTriggers: () => triggers,
    newTrigger(name) {
      const trigger = {name, minutes: null};
      return {timeBased() { return this; }, everyMinutes(minutes) { trigger.minutes = minutes; return this; },
        create() { createdTriggers.push(trigger); triggers.push({getHandlerFunction: () => name}); }};
    }
  };
  const context = vm.createContext({
    console: {log(message) { logs.push(String(message)); }, warn(message) { logs.push(String(message)); }, error(message) { logs.push(String(message)); }},
    SpreadsheetApp: {openById: () => spreadsheet},
    PropertiesService: {getScriptProperties: () => ({getProperty: key => properties[key], setProperty(key, value) { properties[key] = value; }})},
    FormApp: {openById: () => form},
    ScriptApp: scriptApp,
    LockService: {getScriptLock: () => ({tryLock() { if (!state.canLock) return false; state.locked = true; return true; }, waitLock() { state.locked = true; }, releaseLock() { state.locked = false; }})},
    Utilities: {formatDate: () => '2026-11'}
  });
  vm.runInContext(script, context);
  return {context, sheets, legacy, spreadsheet, form, state, logs, properties, triggers, createdTriggers};
}

function payload(stamp = '2026-11-13T10:45:02+09:00') {
  const jst = new Date(Date.parse(stamp) + 9 * 3600000);
  const length = Math.floor((jst.getUTCHours() * 60 + jst.getUTCMinutes()) / 30) + 1;
  return {schemaVersion: 2, date: '2026-11-13', staffId: 'bb1', staffName: '姓 名', observedAt: stamp,
    hours: [{hour: 10, calls: 4, dialogues: 3, toss: 1}],
    intervals: Array.from({length}, (_, i) => ({startMinute: i * 30, calls: i === 20 ? 3 : i === 21 ? 1 : 0,
      dialogues: i === 20 ? 2 : i === 21 ? 1 : 0, toss: i === 20 ? 1 : 0, dialogueSeconds: i === 20 ? 90.5 : i === 21 ? 30 : 0}))};
}

const dataSheet = env => env.sheets.get('通話集計データ_2026-11');
function response(id, body, sentAt = '2026-11-13T11:00:00+09:00', token = 'test-token') {
  return {getId: () => id, getTimestamp: () => new Date(sentAt), getItemResponses: () => [
    {getItem: () => ({getTitle: () => 'token'}), getResponse: () => token},
    {getItem: () => ({getTitle: () => 'data'}), getResponse: () => JSON.stringify(body)}
  ]};
}

if (process.argv[2] === '--payload') {
  const env = environment();
  env.context.validateStatistics(JSON.parse(fs.readFileSync(process.argv[3], 'utf8')));
}
if (process.argv[2] === '--roundtrip') {
  const env = environment();
  env.context.writeStatistics(JSON.parse(fs.readFileSync(process.argv[3], 'utf8')));
  fs.writeFileSync(process.argv[4], JSON.stringify(dataSheet(env).rows));
}

test('creates monthly hidden sheet, stores seconds and metadata, preserves daily formulas', () => {
  const env = environment();
  const result = env.context.writeStatistics(payload());
  assert.deepEqual(clone(result.written), [10]);
  assert.equal(dataSheet(env).hidden, true);
  assert.equal(dataSheet(env).rows.length, 23);
  assert.equal(dataSheet(env).rows[21][9], 90.5);
  assert.deepEqual(dataSheet(env).rows[21].slice(0, 6), ['2026-11-13', 'bb1', '姓 名', '部署A', 'ユニットA', 600]);
  assert.deepEqual(env.legacy.rows[4].slice(7, 10), [4, 3, 1]);
  assert.deepEqual(env.legacy.rows[4].slice(4, 7), [0, 0, 0]);
  assert(dataSheet(env).maxRows >= 23);
});

test('same submission replaces rows; staff IDs separate people; next capture adds only new slots', () => {
  const env = environment();
  env.context.writeStatistics(payload());
  env.context.writeStatistics(payload());
  assert.equal(dataSheet(env).rows.length, 23);
  env.context.writeStatistics({...payload(), staffName: '別人', staffId: 'bb2'});
  assert.equal(dataSheet(env).rows.length, 45);
  env.context.writeStatistics(payload('2026-11-13T11:05:02+09:00'));
  assert.equal(dataSheet(env).rows.length, 46);
});

test('older capture cannot overwrite either legacy table or hidden values', () => {
  const env = environment();
  const newer = payload('2026-11-13T11:00:02+09:00');
  newer.hours[0].calls = 99;
  env.context.writeStatistics(newer);
  const before = clone(dataSheet(env).rows);
  assert.equal(env.context.writeStatistics(payload()).stale, true);
  assert.deepEqual(dataSheet(env).rows, before);
  assert.equal(env.legacy.rows[4][7], 99);
});

test('unknown hidden values stay blank; known zero stays zero; legacy null preserves the cell', () => {
  const env = environment();
  env.context.writeStatistics(payload());
  const data = payload();
  data.intervals[20].dialogues = data.intervals[20].dialogueSeconds = null;
  data.hours[0].dialogues = null;
  env.context.writeStatistics(data);
  assert.equal(dataSheet(env).rows[21][7], '');
  assert.equal(dataSheet(env).rows[21][9], '');
  assert.equal(dataSheet(env).rows[1][9], 0);
  assert.equal(env.legacy.rows[4][8], 3);
});

test('old clients keep working without manufacturing 30 minute data', () => {
  const env = environment();
  const data = payload(); delete data.schemaVersion; delete data.intervals; delete data.observedAt;
  env.context.writeStatistics(data);
  assert.equal(dataSheet(env), undefined);
  assert.equal(env.legacy.rows[4][7], 4);
});

const invalid = [
  ['negative calls', data => data.intervals[0].calls = -1],
  ['fractional counts', data => data.intervals[0].calls = 1.5],
  ['numeric string', data => data.intervals[0].calls = '1'],
  ['negative seconds', data => data.intervals[0].dialogueSeconds = -1],
  ['nonfinite seconds', data => data.intervals[0].dialogueSeconds = Infinity],
  ['missing seconds', data => delete data.intervals[0].dialogueSeconds],
  ['unknown dialogues with known time', data => data.intervals[20].dialogues = null],
  ['dialogues exceed calls', data => data.intervals[0].dialogues = 1],
  ['duplicate intervals', data => data.intervals[1].startMinute = 0],
  ['not a half hour', data => data.intervals[1].startMinute = 31],
  ['future slot', data => data.intervals[1].startMinute = 1410],
  ['incomplete day snapshot', data => data.intervals.pop()],
  ['unknown version', data => data.schemaVersion = 9],
  ['invalid date', data => data.date = '2026-02-30'],
  ['different day', data => data.observedAt = '2026-11-12T10:45:02+09:00'],
  ['ambiguous timestamp', data => data.observedAt = '2026-11-13T10:45:02'],
  ['missing staff ID', data => data.staffId = ''],
  ['formula as staff ID', data => data.staffId = '=IMPORTXML("url")'],
  ['duplicate hour', data => data.hours.push({...data.hours[0]})],
];
for (const [name, alter] of invalid) test(`rejects ${name} before any sheet mutation`, () => {
  const env = environment(), data = payload(), before = clone(env.legacy.rows);
  alter(data);
  assert.throws(() => env.context.writeStatistics(data));
  assert.deepEqual(env.legacy.rows, before);
  assert.equal(env.sheets.size, 1);
});

test('missing or ambiguous roster member does not write hidden data', () => {
  const env = environment();
  assert.throws(() => env.context.writeStatistics({...payload(), staffName: 'unknown'}));
  assert.equal(dataSheet(env), undefined);
  env.legacy.rows.push(clone(env.legacy.rows[4]));
  assert.throws(() => env.context.writeStatistics(payload()));
  assert.equal(dataSheet(env), undefined);
});

test('wrong header does not overwrite legacy results or existing hidden contents', () => {
  const env = environment();
  env.sheets.set('通話集計データ_2026-11', new Sheet('通話集計データ_2026-11', [['old schema']]));
  const before = clone(env.legacy.rows);
  assert.throws(() => env.context.writeStatistics(payload()), /見出し/);
  assert.deepEqual(env.legacy.rows, before);
  assert.deepEqual(dataSheet(env).rows, [['old schema']]);
});

test('setup is idempotent and never clears stored data', () => {
  const env = environment();
  env.context.setupStatisticsData();
  env.context.writeStatistics(payload());
  const before = clone(dataSheet(env).rows);
  env.context.setupStatisticsData();
  assert.deepEqual(dataSheet(env).rows, before);
  assert.equal(env.state.locked, false);
});

test('queue keeps newest observation even when older capture arrives later; old client does not erase details', () => {
  const env = environment();
  const newer = payload('2026-11-13T11:00:02+09:00');
  env.form.responses = [response('newer', newer, '2026-11-13T11:01:00+09:00'), response('older', payload(), '2026-11-13T11:02:00+09:00')];
  env.context.processSubmissions();
  assert.equal(dataSheet(env).rows.length, 24);
  assert(env.form.deleted.includes('newer') && env.form.deleted.includes('older'));
  assert.equal(env.state.locked, false);
});

test('invalid shared token never writes data', () => {
  const env = environment();
  env.form.responses = [response('bad', payload(), undefined, 'wrong-token')];
  env.context.processSubmissions();
  assert.equal(dataSheet(env), undefined);
  assert.deepEqual(env.form.deleted, ['bad']);
});

test('transient write failure retains response; retry repairs without double counting', () => {
  const env = environment();
  env.context.setupStatisticsData();
  dataSheet(env).failWrites = true;
  const recent = new Date().toISOString();
  env.form.responses = [response('retry', payload(), recent)];
  env.context.processSubmissions();
  assert.deepEqual(env.form.deleted, []);
  dataSheet(env).failWrites = false;
  env.context.processSubmissions();
  assert.deepEqual(env.form.deleted, ['retry']);
  assert.equal(dataSheet(env).rows.length, 23);
});

test('concurrent processing returns without changing data', () => {
  const env = environment();
  env.state.canLock = false;
  env.form.responses = [response('locked', payload())];
  env.context.processSubmissions();
  assert.equal(dataSheet(env), undefined);
  assert.deepEqual(env.form.deleted, []);
});

test('dashboard autoRefresh preserves its original stages and never opens the form', () => {
  const events = [];
  const context = vm.createContext({console, FormApp: {openById() { throw new Error('dashboard must not open forms'); }},
    processSubmissions() { throw new Error('dashboard must not consume responses'); },
    autoRefreshBase_() { events.push('dashboard'); }, rebuildClosedMonths_() { events.push('closed'); }, aclRefresh_() { events.push('acl'); }});
  vm.runInContext(integration, context);
  context.autoRefresh();
  assert.deepEqual(events, ['dashboard', 'closed', 'acl']);
});

test('dashboard keeps original error handling for historical rebuild', () => {
  const events = [];
  const context = vm.createContext({console: {log() {}}, autoRefreshBase_() { events.push('dashboard'); },
    rebuildClosedMonths_() { throw new Error('rebuild failed'); }, aclRefresh_() { events.push('acl'); }});
  vm.runInContext(integration, context);
  context.autoRefresh();
  assert.deepEqual(events, ['dashboard', 'acl']);
});

test('receiver setup reuses its existing trigger and never deletes dashboard triggers', () => {
  const env = environment();
  env.triggers.push({getHandlerFunction: () => 'autoRefresh'});
  const info = env.context.setup();
  assert.equal(info.tokenEntry, 'entry.1');
  assert.equal(info.dataEntry, 'entry.2');
  assert.equal(dataSheet(env).hidden, true);
  assert.equal(env.createdTriggers.length, 0);
  assert.equal(env.triggers.length, 2);
  assert.equal(env.properties.FORM_ID, 'form');
});

test('receiver setup creates only one missing processSubmissions trigger at one minute', () => {
  const env = environment();
  env.triggers.splice(0);
  env.triggers.push({getHandlerFunction: () => 'autoRefresh'});
  env.context.ensureStatisticsReceiverTrigger();
  env.context.ensureStatisticsReceiverTrigger();
  assert.deepEqual(env.createdTriggers, [{name: 'processSubmissions', minutes: 1}]);
  assert.equal(env.triggers[0].getHandlerFunction(), 'autoRefresh');
});

test('30 minute trigger can retry a submission after 45 minutes', () => {
  const env = environment();
  env.context.setupStatisticsData();
  dataSheet(env).failWrites = true;
  env.form.responses = [response('delayed', payload(), new Date(Date.now() - 45 * 60000).toISOString())];
  env.context.processSubmissions();
  assert.deepEqual(env.form.deleted, []);
  dataSheet(env).failWrites = false;
  env.context.processSubmissions();
  assert.deepEqual(env.form.deleted, ['delayed']);
});

test('dashboard export compiles, preserves APIs, and contains no receiver or FormApp', () => {
  const combined = fs.readFileSync(path.join(__dirname, '../gas/dashboard_with_call_statistics.gs.txt'), 'utf8');
  new vm.Script(combined);
  assert.equal((combined.match(/^function autoRefresh\(/gm) || []).length, 1);
  for (const name of ['doGet', 'getCallHours', 'autoRefreshBase_', 'rebuildClosedMonths_', 'aclRefresh_']) assert(combined.includes(`function ${name}(`), name);
  assert(!combined.includes('FormApp'));
  assert(!combined.includes('function processSubmissions('));
  const fullContext = vm.createContext({console});
  vm.runInContext(combined, fullContext);
  const events = [];
  for (const name of ['autoRefreshBase_', 'rebuildClosedMonths_', 'aclRefresh_']) fullContext[name] = () => events.push(name);
  fullContext.processSubmissions = () => { throw new Error('must not consume forms'); };
  fullContext.autoRefresh();
  assert.deepEqual(events, ['autoRefreshBase_', 'rebuildClosedMonths_', 'aclRefresh_']);
});

test('reproduces the legacy receiver discarding dialogue time before the second script can read it', () => {
  const old = environment(legacySource);
  old.form.responses = [response('already-consumed', payload())];
  old.context.processSubmissions();
  assert.equal(old.legacy.rows[4][7], 4);
  assert.equal(dataSheet(old), undefined);
  assert.deepEqual(old.form.deleted, ['already-consumed']);
  const newer = environment();
  newer.context.FormApp.openById = () => old.form;
  newer.context.processSubmissions();
  assert.equal(dataSheet(newer), undefined);
});

test('updated receiver saves dialogue time before deleting the shared form response', () => {
  const env = environment();
  const remove = env.form.deleteResponse.bind(env.form);
  env.form.deleteResponse = id => { assert.equal(dataSheet(env).rows[21][9], 90.5); remove(id); };
  env.form.responses = [response('saved-first', payload())];
  env.context.processSubmissions();
  assert.equal(env.form.responses.length, 0);
  const log = env.logs.map(line => { try { return JSON.parse(line); } catch { return null; } }).find(item => item && item.intervalCount);
  assert.equal(log.intervalSheet, '通話集計データ_2026-11');
  assert.equal(log.intervalCount, 22);
});

for (const key of ['TOKEN', 'FORM_ID']) test(`missing ${key} never consumes or deletes responses`, () => {
  const env = environment();
  delete env.properties[key];
  env.form.responses = [response('keep', payload())];
  assert.throws(() => env.context.processSubmissions(), new RegExp(key));
  assert.equal(env.form.responses.length, 1);
  assert.deepEqual(env.form.deleted, []);
  assert.equal(env.state.locked, false);
});

test('receiver diagnostics report configuration and positive dialogue time without mutation or secrets', () => {
  const env = environment();
  env.form.responses = [response('v2', payload()), response('v1', {date: '2026-11-13', staffName: '姓 名', hours: []})];
  const result = env.context.diagnoseStatisticsReceiver();
  assert.equal(result.pendingResponses, 2);
  assert.equal(result.v2Responses, 1);
  assert.equal(result.responsesWithDialogueSeconds, 1);
  assert.equal(result.dataSheet.exists, false);
  assert.equal(env.sheets.size, 1);
  assert.deepEqual(env.form.deleted, []);
  assert(!env.logs.join('\n').includes('test-token'));
  assert(!env.logs.join('\n').includes('姓 名'));
});

test('paste-ready txt exports match their sources exactly and remain separated', () => {
  const receiver = fs.readFileSync(path.join(__dirname, '../gas/exports/01_receiver_save.gs.txt'), 'utf8');
  const dashboard = fs.readFileSync(path.join(__dirname, '../gas/exports/02_dashboard.gs.txt'), 'utf8');
  const compat = fs.readFileSync(path.join(__dirname, '../gas/dashboard_with_call_statistics.gs.txt'), 'utf8');
  assert.equal(receiver.trimEnd(), source.trimEnd());
  assert.equal(dashboard, compat);
  new vm.Script(receiver);
  new vm.Script(dashboard);
  assert(receiver.includes('function processSubmissions('));
  assert(!receiver.includes('function autoRefresh('));
  assert(dashboard.includes('function autoRefresh('));
  assert(!dashboard.includes('function processSubmissions('));
});
