/**
 * TelephoneTool → 「CC第二コール管理」の「コール集計打ち込み用_M月」へ、1時間ごとのコール・対話・トスを書き込む。
 *
 * 会社の設定でウェブアプリを「全員」に公開できないため、ログイン不要の送信用Googleフォームで受け取る。
 * このスプレッドシートは開くだけで数十秒〜100秒近くかかるため、送信のたびには開かず、
 * 1分ごとの時間トリガーで、届いた送信をまとめて1回開いて書き込む（送信が無い回は開かない）。
 * TelephoneTool は毎時「本人・当日」の全時間帯を送るので、同じ人・同じ日は最新の送信だけを書けばよい。
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
const INTERVAL_MINUTES = 1; // 送信が無い回はシートを開かないので、短くしても負担は小さい

/** 初回だけエディタから実行する。何度実行しても、フォームとトリガーは1つだけになる。 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('FORM_ID');
  let form = formId ? FormApp.openById(formId) : null;
  if (!form) {
    form = FormApp.create(FORM_TITLE);
    form.setDescription('TelephoneTool が1時間ごとに通話集計を送るためのフォームです。手で回答しないでください。');
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
  return info;
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
    responses.forEach(function (response) {
      const answers = {};
      response.getItemResponses().forEach(function (r) { answers[r.getItem().getTitle()] = String(r.getResponse()); });
      try {
        if (!token || answers.token !== token) throw new Error('認証に失敗しました（トークンが一致しません）');
        const body = JSON.parse(answers.data || '{}');
        // 毎回その日の全時間帯が送られてくるので、同じ人・同じ日は最後の送信だけ書く。
        const key = body.date + '|' + normalize(body.staffName);
        const sentAt = response.getTimestamp().getTime();
        if (!latest[key] || latest[key].sentAt <= sentAt) latest[key] = { body: body, sentAt: sentAt };
      } catch (err) {
        console.error('通話集計の送信を読み取れません: ' + ((err && err.message) || err));
      }
    });
    const book = { spreadsheet: null, sheets: {} };
    Object.keys(latest).forEach(function (key) {
      const body = latest[key].body;
      try {
        const result = writeStatistics(body, book);
        console.log(JSON.stringify({ staff: body.staffName, date: body.date, row: result.row, written: result.written, skipped: result.skipped }));
      } catch (err) {
        console.error('通話集計の書き込みに失敗（' + body.staffName + ' ' + body.date + '）: ' + ((err && err.message) || err));
      }
    });
    // 失敗した分もTelephoneToolが次の時間に全時間帯を送り直すので、回答は残さない。
    responses.forEach(function (response) {
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

/** 全角・半角と空白の違いを無視して比べる（CTIの「岩田 悠河」とシートの「岩田悠河」を同一視）。 */
function normalize(value) {
  return String(value === null || value === undefined ? '' : value).normalize('NFKC').replace(/\s+/g, '');
}

/** 月のシートと、名前・日付を探すA〜D列。1回の処理の中では一度だけ読む。 */
function monthSheet(book, year, month) {
  const sheetName = SHEET_NAME_PREFIX + month + '月';
  if (!book.sheets[sheetName]) {
    if (!book.spreadsheet) book.spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = book.spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error('シート「' + sheetName + '」がありません');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    book.sheets[sheetName] = {
      name: sheetName, sheet: sheet, lastColumn: lastColumn, headers: {},
      left: sheet.getRange(1, 1, lastRow, Math.min(LOOKUP_COLUMNS, lastColumn)).getDisplayValues(),
    };
  }
  const data = book.sheets[sheetName];
  if (!normalize(data.left[0].join('')).includes(year + '年' + month + '月')) {
    throw new Error('シート「' + sheetName + '」が' + year + '年' + month + '月の表ではありません');
  }
  return data;
}

function writeStatistics(body, book) {
  const date = /^(\d{4})-(\d{2})-(\d{2})$/.exec(body.date || '');
  if (!date) throw new Error('日付の形式が不正です');
  const year = Number(date[1]);
  const month = Number(date[2]);
  const day = Number(date[3]);
  const staffName = normalize(body.staffName);
  if (!staffName) throw new Error('名前がありません');
  const data = monthSheet(book || { spreadsheet: null, sheets: {} }, year, month);
  const left = data.left;

  // 「【11/1（日）】」のような日付の見出し行を探す。「【11/1(」は「【11/10(」に一致しない。
  const marker = normalize('【' + month + '/' + day + '（');
  const isBlock = function (row) { return row.some(function (cell) { return /^【\d+\/\d+\(/.test(normalize(cell)); }); };
  const blockIndex = left.findIndex(function (row) {
    return row.some(function (cell) { return normalize(cell).indexOf(marker) === 0; });
  });
  if (blockIndex < 0) throw new Error(month + '/' + day + 'の表が見つかりません');

  if (!data.headers[blockIndex]) {
    data.headers[blockIndex] = data.sheet.getRange(blockIndex + 2, 1, 2, data.lastColumn).getDisplayValues();
  }
  const hourRow = data.headers[blockIndex][0];
  const labelRow = data.headers[blockIndex][1];
  const nameColumn = hourRow.slice(0, LOOKUP_COLUMNS).findIndex(function (cell) { return normalize(cell) === '名前'; });
  if (nameColumn < 0) throw new Error(month + '/' + day + 'の表に「名前」の列が見つかりません');

  const matches = [];
  for (let r = blockIndex + 3; r < left.length && !isBlock(left[r]); r++) {
    if (normalize(left[r][nameColumn]) === staffName) matches.push(r);
  }
  if (matches.length !== 1) {
    throw new Error(matches.length
      ? '「' + body.staffName + '」が' + month + '/' + day + 'の表に複数あります'
      : '「' + body.staffName + '」が' + month + '/' + day + 'の表の名前欄にありません');
  }
  const sheetRow = matches[0] + 1;

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
  return { ok: true, sheet: data.name, row: sheetRow, written: written, skipped: skipped };
}
