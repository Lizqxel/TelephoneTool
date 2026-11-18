/**
 * TelephoneTool → 「CC第二コール管理」の「コール集計打ち込み用_M月」へ、1時間ごとのコール・対話・トスを書き込む。
 *
 * 会社の設定でウェブアプリを「全員」に公開できないため、ログイン不要の送信用Googleフォームで受け取る。
 * このスプレッドシートは開くだけで数十秒〜100秒近くかかるため、送信のたびには開かず、
 * 1分ごとの時間トリガーで、届いた送信をまとめて1回開いて書き込む（送信が無い回は開かない）。
 * TelephoneTool は30分ごとに「本人・当日」の取得済み全時間帯を送る。
 * 旧来の1時間集計表へはコール・対話・トスを反映し、30分集計・対話時間は
 * 非表示の「通話集計データ_YYYY-MM」へ保存する。対話時間は秒数。
 * 書き込むのは 日付の表 × 名前の行 × 「H時台」の3列（コール／対話／トス）だけで、
 * 名前・部署・GPユニット・日計の列や表の構成には触れない。何度書いても重複しない。
 *
 * 初回の設定：
 *   1. プロジェクトの設定 → スクリプト プロパティに TOKEN（TelephoneToolの共有トークンと同じ値）を登録する。
 *   2. エディタで setup を実行して権限を承認する。送信用フォームと1分ごとのトリガーが作られ、
 *      TelephoneTool に設定する formResponseUrl・tokenEntry・dataEntry が実行ログに出る。
 * 書き込み結果やエラーは「実行数」の画面（processSubmissions のログ）で確認できる。
 */

const SPREADSHEET_ID = '1Asghaq_Hp3tqWQpirbBoBnQpXsKuWiTi33Hzfq-0qRU';
const SHEET_NAME_PREFIX = 'コール集計打ち込み用_';
const LOOKUP_COLUMNS = 4; // 名前の列と日付見出しを探す範囲（A〜D列）
const FORM_TITLE = 'TelephoneTool 通話集計（自動送信用・手入力しないでください）';
const HANDLER = 'processSubmissions';
const RETRY_MINUTES = 30; // シートを開けないなど一時的な失敗は、この時間まで1分ごとに再処理する
const INTERVAL_MINUTES = 1; // 送信が無い回はシートを開かないので、短くしても負担は小さい
const STAT_DATA_PREFIX = '通話集計データ_';
const STAT_DATA_HEADERS = ['日付', 'スタッフID', '氏名', '部署', 'GPユニット', '開始分', 'コール数', '対話数', 'トス数', '対話時間（秒）', '取得日時'];

/** 初回だけエディタから実行する。何度実行しても、フォームとトリガーは1つだけになる。 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('FORM_ID');
  let form = formId ? FormApp.openById(formId) : null;
  if (!form) {
    form = FormApp.create(FORM_TITLE);
    form.setDescription('TelephoneTool が30分ごとに通話集計を送るためのフォームです。手で回答しないでください。');
    form.addTextItem().setTitle('token').setRequired(true);
    form.addParagraphTextItem().setTitle('data').setRequired(true);
    props.setProperty('FORM_ID', form.getId());
  }
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(false);
  form.setRequireLogin(false);
  if (typeof form.setPublished === 'function') form.setPublished(true);
  form.setAcceptingResponses(true);

  // 以前の版のフォーム送信トリガー（onStatisticsSubmit）も含め、作り直す。
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return ['onStatisticsSubmit', HANDLER].indexOf(t.getHandlerFunction()) >= 0; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(HANDLER).timeBased().everyMinutes(INTERVAL_MINUTES).create();

  const items = form.getItems();
  const prefilled = form.createResponse()
    .withItemResponse(items[0].asTextItem().createResponse('TOKENVALUE'))
    .withItemResponse(items[1].asParagraphTextItem().createResponse('DATAVALUE'))
    .toPrefilledUrl();
  const entry = function (marker) {
    const found = new RegExp('(entry\\.\\d+)=' + marker).exec(prefilled);
    return found ? found[1] : '';
  };
  // requiresLogin() は組織の設定によっては読めないため、確認できなくても設定値の出力は続ける。
  let requiresLogin;
  try {
    requiresLogin = form.requiresLogin();
  } catch (err) {
    requiresLogin = '確認できません（' + err.message + '）';
  }
  const info = {
    formResponseUrl: form.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse'),
    tokenEntry: entry('TOKENVALUE'),
    dataEntry: entry('DATAVALUE'),
    requiresLogin: requiresLogin,
    tokenRegistered: Boolean(props.getProperty('TOKEN')),
  };
  console.log('TelephoneTool設定: ' + JSON.stringify(info));
  setupStatisticsData();
  return info;
}

/** 管理者が初回に実行する。翌月以降のシートは受信時に自動作成する。 */
function setupStatisticsData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const book = { spreadsheet: SpreadsheetApp.openById(SPREADSHEET_ID), sheets: {} };
    const ym = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    statisticsDataSheet(book, ym);
    console.log('30分集計用の非表示シートを準備しました: ' + STAT_DATA_PREFIX + ym);
  } finally {
    lock.releaseLock();
  }
}

/** 1分ごとの時間トリガー。届いた送信をまとめてシートへ書き込み、処理した回答（トークンを含む）は消す。 */
function processSubmissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // 前回の処理がまだ続いている
  try {
    const props = PropertiesService.getScriptProperties();
    const form = FormApp.openById(props.getProperty('FORM_ID'));
    const responses = form.getResponses();
    if (!responses.length) return;
    const token = props.getProperty('TOKEN');
    const latest = {};
    const done = []; // 処理を終えた回答（書き込めた・やり直しても無駄・古すぎる）。トークンを含むので消す
    responses.forEach(function (response) {
      const answers = {};
      response.getItemResponses().forEach(function (r) { answers[r.getItem().getTitle()] = String(r.getResponse()); });
      try {
        if (!token || answers.token !== token) throw new Error('認証に失敗しました（トークンが一致しません）');
        const body = JSON.parse(answers.data || '{}');
        validateStatistics(body);
        // 毎回その日の全時間帯が送られてくるので、同じ人・同じ日は最後の送信だけ書く（古い分は同じ結果で片付ける）。
        // 旧クライアントの送信で、新形式の30分データを取りこぼさない。
        const key = body.date + '|' + normalize(body.staffName) + '|' + (body.schemaVersion === 2 ? 'v2' : 'v1');
        const sentAt = response.getTimestamp().getTime();
        const sourceAt = body.schemaVersion === 2 ? Date.parse(body.observedAt) : sentAt;
        const entry = latest[key] || (latest[key] = { body: body, sentAt: sentAt, sourceAt: sourceAt, responses: [] });
        entry.responses.push(response);
        if (entry.sourceAt < sourceAt || (entry.sourceAt === sourceAt && entry.sentAt <= sentAt)) {
          entry.body = body;
          entry.sentAt = sentAt;
          entry.sourceAt = sourceAt;
        }
      } catch (err) {
        console.error('通話集計の送信を読み取れません: ' + ((err && err.message) || err));
        done.push(response);
      }
    });
    const book = { spreadsheet: null, sheets: {} };
    Object.keys(latest).forEach(function (key) {
      const entry = latest[key];
      const body = entry.body;
      try {
        const result = writeStatistics(body, book);
        console.log(JSON.stringify({ staff: body.staffName, date: body.date, row: result.row, written: result.written, skipped: result.skipped }));
        Array.prototype.push.apply(done, entry.responses);
      } catch (err) {
        const message = (err && err.message) || err;
        const expired = Date.now() - entry.sentAt > RETRY_MINUTES * 60 * 1000;
        if (err && err.permanent) {
          // 名前が表に無いなど、やり直しても同じ結果になるもの
          console.error('通話集計の書き込みに失敗（' + body.staffName + ' ' + body.date + '）: ' + message);
          Array.prototype.push.apply(done, entry.responses);
        } else if (expired) {
          console.error('通話集計の書き込みを' + RETRY_MINUTES + '分やり直しても失敗したため破棄（' + body.staffName + ' ' + body.date + '）: ' + message);
          Array.prototype.push.apply(done, entry.responses);
        } else {
          // シートを開けないなど一時的な失敗は、回答を残して次の回（1分後）にやり直す
          console.warn('通話集計の書き込みを次の回にやり直します（' + body.staffName + ' ' + body.date + '）: ' + message);
        }
      }
    });
    done.forEach(function (response) {
      try {
        form.deleteResponse(response.getId());
      } catch (err) {
        console.warn('回答の削除に失敗: ' + err);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

/** やり直しても同じ結果になる失敗（表や名前が無いなど）。回答は残さず片付ける。 */
function permanentError(message) {
  const error = new Error(message);
  error.permanent = true;
  return error;
}

/** 全角・半角と空白の違いを無視して比べる（CTIの「岩田 悠河」とシートの「岩田悠河」を同一視）。 */
function normalize(value) {
  return String(value === null || value === undefined ? '' : value).normalize('NFKC').replace(/\s+/g, '');
}

/** 月のシートと、名前・日付を探すA〜D列。1回の処理の中では一度だけ読む。 */
function monthSheet(book, year, month) {
  const sheetName = SHEET_NAME_PREFIX + month + '月';
  if (!book.sheets[sheetName]) {
    // 開けなかったら、同じ回の残りの人では開き直さない（1回で数十秒かかるため）。次の回にやり直す
    if (book.openError) throw book.openError;
    if (!book.spreadsheet) {
      try {
        book.spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      } catch (err) {
        book.openError = err;
        throw err;
      }
    }
    const sheet = book.spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw permanentError('シート「' + sheetName + '」がありません');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    book.sheets[sheetName] = {
      name: sheetName, sheet: sheet, lastColumn: lastColumn, headers: {},
      left: sheet.getRange(1, 1, lastRow, Math.min(LOOKUP_COLUMNS, lastColumn)).getDisplayValues(),
    };
  }
  const data = book.sheets[sheetName];
  if (!normalize(data.left[0].join('')).includes(year + '年' + month + '月')) {
    throw permanentError('シート「' + sheetName + '」が' + year + '年' + month + '月の表ではありません');
  }
  return data;
}

function writeStatistics(body, book) {
  validateStatistics(body); // 検証に失敗した送信は、どのセルにも書かない。
  book = book || { spreadsheet: null, sheets: {} };
  const date = /^(\d{4})-(\d{2})-(\d{2})$/.exec(body.date || '');
  if (!date) throw permanentError('日付の形式が不正です');
  const year = Number(date[1]);
  const month = Number(date[2]);
  const day = Number(date[3]);
  const staffName = normalize(body.staffName);
  if (!staffName) throw permanentError('名前がありません');
  const data = monthSheet(book, year, month);
  const left = data.left;

  // 「【11/1（日）】」のような日付の見出し行を探す。「【11/1(」は「【11/10(」に一致しない。
  const marker = normalize('【' + month + '/' + day + '（');
  const isBlock = function (row) { return row.some(function (cell) { return /^【\d+\/\d+\(/.test(normalize(cell)); }); };
  const blockIndex = left.findIndex(function (row) {
    return row.some(function (cell) { return normalize(cell).indexOf(marker) === 0; });
  });
  if (blockIndex < 0) throw permanentError(month + '/' + day + 'の表が見つかりません');

  if (!data.headers[blockIndex]) {
    data.headers[blockIndex] = data.sheet.getRange(blockIndex + 2, 1, 2, data.lastColumn).getDisplayValues();
  }
  const hourRow = data.headers[blockIndex][0];
  const labelRow = data.headers[blockIndex][1];
  const nameColumn = hourRow.slice(0, LOOKUP_COLUMNS).findIndex(function (cell) { return normalize(cell) === '名前'; });
  if (nameColumn < 0) throw permanentError(month + '/' + day + 'の表に「名前」の列が見つかりません');

  const matches = [];
  for (let r = blockIndex + 3; r < left.length && !isBlock(left[r]); r++) {
    if (normalize(left[r][nameColumn]) === staffName) matches.push(r);
  }
  if (matches.length !== 1) {
    throw permanentError(matches.length
      ? '「' + body.staffName + '」が' + month + '/' + day + 'の表に複数あります'
      : '「' + body.staffName + '」が' + month + '/' + day + 'の表の名前欄にありません');
  }
  const sheetRow = matches[0] + 1;
  const store = body.schemaVersion === 2 ? statisticsDataSheet(book, body.date.slice(0, 7)) : null;
  const firstKey = body.date + '|' + body.staffId + '|0';
  if (store && store.index[firstKey] !== undefined &&
      Date.parse(store.rows[store.index[firstKey]][10]) > Date.parse(body.observedAt)) {
    return { ok: true, sheet: data.name, row: sheetRow, written: [], skipped: [], stale: true };
  }

  const written = [];
  const skipped = [];
  (body.hours || []).forEach(function (item) {
    const column = hourRow.findIndex(function (cell) { return normalize(cell) === item.hour + '時台'; });
    const labels = column < 0 ? [] : labelRow.slice(column, column + 3).map(normalize);
    if (labels.join('/') !== 'コール/対話/トス') {
      skipped.push(item.hour);
      return;
    }
    const range = data.sheet.getRange(sheetRow, column + 1, 1, 3);
    const current = range.getValues()[0];
    // 未取得（null）の値は上書きせず、シートの値を残す。
    const next = [item.calls, item.dialogues, item.toss].map(function (value, i) {
      return value === null || value === undefined ? current[i] : value;
    });
    range.setValues([next]);
    written.push(item.hour);
  });
  if (store) {
    // 名前・部署・ユニットは既存の当日表と対応づける。
    writeStatisticsData(store, body, left[matches[0]][nameColumn + 1] || '', left[matches[0]][nameColumn + 2] || '');
  }
  return { ok: true, sheet: data.name, row: sheetRow, written: written, skipped: skipped };
}

/** 型・範囲・重複・取得日時をすべて確認してから書き込む。旧形式も受け付ける。 */
function validateStatistics(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) throw permanentError('送信データが不正です');
  const date = typeof body.date === 'string' && /^(\d{4})-(\d{2})-(\d{2})$/.exec(body.date);
  if (!date || Number(date[2]) < 1 || Number(date[2]) > 12 || Number(date[3]) < 1 ||
      Number(date[3]) > new Date(Date.UTC(Number(date[1]), Number(date[2]), 0)).getUTCDate()) throw permanentError('日付の形式が不正です');
  if (typeof body.staffName !== 'string' || !normalize(body.staffName) || /[\r\n\t]/.test(body.staffName)) throw permanentError('名前が不正です');
  if (!Array.isArray(body.hours) || body.hours.length > 24) throw permanentError('1時間集計の形式が不正です');
  const validNumber = function (value, integer) {
    return value === null || (typeof value === 'number' && isFinite(value) && value >= 0 && (!integer || Number.isInteger(value)));
  };
  const seen = {};
  body.hours.forEach(function (item) {
    if (!item || !Number.isInteger(item.hour) || item.hour < 0 || item.hour > 23 || seen[item.hour]) throw permanentError('時間帯が不正または重複しています');
    seen[item.hour] = true;
    ['calls', 'dialogues', 'toss'].forEach(function (key) {
      if (!validNumber(item[key], true)) throw permanentError('1時間集計の件数が不正です');
    });
  });
  if (body.schemaVersion === undefined) return;
  if (body.schemaVersion !== 2) throw permanentError('対応していない送信バージョンです');
  if (typeof body.staffId !== 'string' || !body.staffId.trim() || body.staffId.length > 128 || /[|=\r\n\t]/.test(body.staffId)) throw permanentError('スタッフIDが不正です');
  if (typeof body.observedAt !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,6})?(Z|[+-]\d{2}:\d{2})$/.test(body.observedAt) || !isFinite(Date.parse(body.observedAt))) throw permanentError('取得日時が不正です');
  const jst = new Date(Date.parse(body.observedAt) + 9 * 60 * 60 * 1000);
  if (jst.toISOString().slice(0, 10) !== body.date) throw permanentError('取得日時と対象日が一致しません');
  const length = Math.floor((jst.getUTCHours() * 60 + jst.getUTCMinutes()) / 30) + 1;
  if (!Array.isArray(body.intervals) || body.intervals.length !== length) throw permanentError('取得済みの全30分枠が必要です');
  const minutes = {};
  body.intervals.forEach(function (item) {
    if (!item || !Number.isInteger(item.startMinute) || item.startMinute < 0 || item.startMinute >= length * 30 || item.startMinute % 30 || minutes[item.startMinute]) throw permanentError('30分枠が不正または重複しています');
    minutes[item.startMinute] = true;
    ['calls', 'dialogues', 'toss'].forEach(function (key) {
      if (!validNumber(item[key], true)) throw permanentError('30分集計の件数が不正です');
    });
    if (!validNumber(item.dialogueSeconds, false) || (item.dialogues === null && item.dialogueSeconds !== null)) throw permanentError('対話時間が不正です');
    if (item.calls !== null && item.dialogues !== null && item.dialogues > item.calls) throw permanentError('対話数がコール数を超えています');
  });
}

/** 月別に保存するため、読み取り時に過去の全期間を走査しない。 */
function statisticsDataSheet(book, ym) {
  const name = STAT_DATA_PREFIX + ym;
  book.statistics = book.statistics || {};
  if (book.statistics[name]) return book.statistics[name];
  let sh = book.spreadsheet.getSheetByName(name);
  if (!sh) sh = book.spreadsheet.insertSheet(name);
  if (!sh.getLastRow()) {
    sh.getRange(1, 1, 1, STAT_DATA_HEADERS.length).setValues([STAT_DATA_HEADERS]);
    sh.setFrozenRows(1);
  }
  const values = sh.getRange(1, 1, sh.getLastRow(), STAT_DATA_HEADERS.length).getValues();
  if (JSON.stringify(values[0]) !== JSON.stringify(STAT_DATA_HEADERS)) throw new Error('30分集計シートの見出しが一致しません: ' + name);
  sh.hideSheet();
  const rows = values.slice(1), index = Object.create(null);
  rows.forEach(function (row, i) {
    const key = row[0] + '|' + row[1] + '|' + row[5];
    if (index[key] !== undefined) throw new Error('30分集計シートに重複行があります: ' + name);
    index[key] = i;
  });
  return book.statistics[name] = { sheet: sh, rows: rows, index: index };
}

/** 加算せず置換する。途中で失敗しても、同じ送信の再処理で復旧する。 */
function writeStatisticsData(store, body, department, unit) {
  const changed = [];
  body.intervals.forEach(function (item) {
    const key = body.date + '|' + body.staffId + '|' + item.startMinute;
    let index = store.index[key];
    if (index === undefined) { index = store.rows.length; store.index[key] = index; store.rows.push([]); }
    store.rows[index] = [body.date, body.staffId, body.staffName, department, unit, item.startMinute,
                         item.calls, item.dialogues, item.toss, item.dialogueSeconds, body.observedAt]
      .map(function (value) { return value === null ? '' : value; });
    changed.push(index);
  });
  changed.sort(function (a, b) { return a - b; });
  if (store.sheet.getMaxRows() < store.rows.length + 1) store.sheet.insertRowsAfter(store.sheet.getMaxRows(), store.rows.length + 1 - store.sheet.getMaxRows());
  // 連続する行をまとめて書く。他の人の行は変更しない。
  for (let i = 0; i < changed.length;) {
    const start = changed[i];
    let end = start;
    while (++i < changed.length && changed[i] === end + 1) end = changed[i];
    const count = end - start + 1;
    store.sheet.getRange(start + 2, 1, count, 5).setNumberFormat('@');
    store.sheet.getRange(start + 2, 11, count, 1).setNumberFormat('@');
    // シートへ式として解釈される文字列は、文字列指定の先頭記号を付ける。
    const values = store.rows.slice(start, end + 1).map(function (row) {
      return row.map(function (value) { return typeof value === 'string' && /^[=+@-]/.test(value) ? "'" + value : value; });
    });
    store.sheet.getRange(start + 2, 1, count, STAT_DATA_HEADERS.length).setValues(values);
  }
}
