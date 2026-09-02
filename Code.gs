// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//  Bemolle タイムカード — Google Apps Script バックエンド
//
//  セットアップ手順:
//  1. Google スプレッドシートを新規作成
//  2. 拡張機能 → Apps Script を開く
//  3. このファイルの内容を全コピーして貼り付け
//  4. 保存（Ctrl+S）
//  5. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
//     ・次のユーザーとして実行: 自分
//     ・アクセスできるユーザー: 全員
//  6. デプロイ → URLをコピー → index.html の API_URL に貼り付け
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const SH_RECORDS  = '打刻記録';
const SH_MEMOS    = '日報メモ';
const SH_STAFF    = 'スタッフ';
const SH_SETTINGS = '設定';

const NAKATA_NAME   = '中田有加';   // 日当あり・メモ連動
const TAKEUCHI_NAME = '竹内美佳';   // 勤務時間のみ

// 現在月のシート名を返す（例: "中田有加_2026-05"）
function nakataSheetName()   { return NAKATA_NAME   + '_' + monthKey(); }
function takeuchiSheetName() { return TAKEUCHI_NAME + '_' + monthKey(); }
function monthKey() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

// ─── エントリポイント ───────────────────────────────
function doGet(e) {
  return handle(e.parameter, null);
}

function doPost(e) {
  const body = e.postData ? JSON.parse(e.postData.contents) : {};
  return handle(e.parameter, body);
}

// ━━ セキュリティの前提と受容リスク（2026-07-14 Sol敵対的レビュー2巡の結論）━━
//  この端末は「ログインなしの公開キオスク」。最も危険な保存型XSS(→管理者PIN窃取)は
//  index.html側で全描画値エスケープにより封鎖済み。以下は"サロン運営上の現実的リスクは低い"
//  として意図的に受容（対策するにはGoogleアカウント認証等の設計変更が必要・別途）：
//   - clockActionは本人確認まではしない（キオゆえ他人になりすまし打刻は原理的に可能）
//   - saveMemoは登録スタッフ名+文字数のみ検証（任意日メモの上書き/削除の余地）
//   - 読み取り総当たり/PINロックのDoSの理論的余地
//  管理者PINは1234以外に設定済み（初期値1234はinitSheets内のみ・実機は独自PIN）。

// ── 管理者PIN門番（2026-07-13追加）─────────────────────────────
// このWebアプリのURLは誰でも到達できるため、危険なアクションはサーバー側で
// PINを検証する（画面のPINモーダルは飾りで、直接APIを叩けば素通りだった）。
// getRecords/getMemos/getAllMemos/clockAction/getStaff/checkPin/keepWarm は
// スタッフ画面が使うため開放のまま。
const ADMIN_ACTIONS = ['init','deleteRecord','updateRecordTimestamp','adminClockAction',
  'clearRecords','addStaff','removeStaff','changePin','setupTriggers','getDailyAllowance',
  'deleteMemo'];
// ⚠️ getStartupData は一般スタッフの起動API（返すのは getStaff/getRecords/getAllMemos と同じ公開データ）。
//    2026-07-14に誤って管理専用へ入れられ、2026-09-02の本番反映で起動が unauthorized → 打刻ボタンがロックされる事故になった。
//    公開に戻す。init も本番では不要なため起動時に呼ばない（index.html側）。

function getStoredPin() {
  const rows = sh(SH_SETTINGS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'pin') return rows[i][1].toString();
  }
  return '';  // 2026-07-14 既知の初期値1234を廃止。未設定時は空＝どのPINとも一致せず管理操作は全拒否(fail closed)
}

// 2026-07-14 リクエストのPINが正しいか（管理者か）。PIN未設定(空)なら常にfalse
function isAdmin(params, d) {
  const stored = getStoredPin();
  if (!stored) return false;
  const pin = ((d && d.pin != null ? d.pin : (params && params.pin)) || '').toString();
  return pin === stored;
}

// 2026-07-14 PIN総当たり対策：連続失敗をCacheServiceで数え、一定回数でロック
function pinRateGuard_(pin) {
  const cache = CacheService.getScriptCache();
  const stored = getStoredPin();
  const failsKey = 'pin_fails';
  var fails = Number(cache.get(failsKey) || 0);
  if (fails >= 10) throw new Error('locked');  // 10回失敗で15分ロック
  const ok = !!stored && pin.toString() === stored;
  if (!ok) cache.put(failsKey, fails + 1, 900); else cache.remove(failsKey);
  return ok;
}

function handle(params, body) {
  const d      = body || {};
  const action = params.action || d.action;
  try {
    if (ADMIN_ACTIONS.indexOf(action) !== -1) {
      const pin = ((d.pin != null ? d.pin : params.pin) || '').toString();
      if (!pinRateGuard_(pin)) throw new Error('unauthorized');
    }
    // 2026-07-14 読み取り系(getRecords等)を非管理者は「今日」だけに制限するための判定
    params.__admin = isAdmin(params, d);
    let result;
    switch (action) {
      case 'init':              result = initSheets();              break;
      case 'clockAction':       result = clockAction(d);            break;
      case 'getRecords':        result = getRecords(params);        break;
      case 'deleteRecord':      result = deleteRecord(d);           break;
      case 'updateRecordTimestamp': result = updateRecordTimestamp(d); break;
      case 'adminClockAction':  result = adminClockAction(d);       break;
      case 'clearRecords':      result = clearRecords();            break;
      case 'getMemos':          result = getMemos(params);          break;
      case 'getAllMemos':        result = getAllMemos(params);       break;
      case 'saveMemo':          result = saveMemo(d);               break;
      case 'deleteMemo':        result = saveMemo({date:d.date, staff:d.staff, text:''}); break;
      case 'getDailyAllowance': result = getDailyAllowance(params); break;
      case 'getStaff':          result = getStaff();                break;
      case 'addStaff':          result = addStaff(d);               break;
      case 'removeStaff':       result = removeStaff(d);            break;
      case 'checkPin':          result = checkPin(d);                break;
      case 'changePin':         result = changePin(d);              break;
      case 'getStartupData':   result = getStartupData(params);   break;
      case 'setupTriggers':    result = setupTriggers();          break;
      case 'keepWarm':         result = { ok: true };             break;
      default: throw new Error('Unknown action: ' + action);
    }
    return ok(result);
  } catch (err) {
    return err_(err.message);
  }
}

function ok(data)  { return json({ ok: true,  data: data }); }
function err_(msg) { return json({ ok: false, error: msg }); }
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── シート初期化 ────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function getOrCreate(name, headers) {
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.appendRow(headers);
    }
    return s;
  }

  getOrCreate(SH_RECORDS, ['id', 'staff', 'type', 'timestamp']);
  getOrCreate(SH_MEMOS,   ['date', 'staff', 'memo']);

  // 今月の個人シートを作成
  const nSh = getOrCreate(nakataSheetName(),   ['日付', '出勤', '退勤', '日当(円)']);
  nSh.getRange('D2:D200').setNumberFormat('¥#,##0');
  getOrCreate(takeuchiSheetName(), ['日付', '打刻時刻', '出勤', '退勤', '勤務時間', '実働時間（休憩1.5h引き）']);

  const staffSh = getOrCreate(SH_STAFF, ['name']);
  if (staffSh.getLastRow() <= 1) {
    staffSh.appendRow([NAKATA_NAME]);
    // 竹内美佳は2026-07-01より運用対象外（新規登録しない／過去データは保持）
  }

  const setsSh = getOrCreate(SH_SETTINGS, ['key', 'value']);
  if (setsSh.getLastRow() <= 1) {
    setsSh.appendRow(['pin', '1234']);
  }

  refreshNakataSheet();
  refreshTakeuchiSheet();
  return { initialized: true };
}

function sh(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) { initSheets(); return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  return s;
}

// ─── 打刻 ────────────────────────────────────────────
function clockAction(d) {
  // 2026-07-14 入力検証：登録済みスタッフ名＋type厳格化（任意名の注入・XSS素材化を防ぐ）
  if (getStaff().indexOf(d.staff) === -1) throw new Error('unknown staff');
  // 同時リクエスト（302後の再送・二重タップ）で同じ打刻が2件入るのを防ぐ（Sol指摘・2026-09-02）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('busy: しばらくしてからもう一度押してください');
  try {
    return clockActionLocked_(d);
  } finally {
    lock.releaseLock();
  }
}

function clockActionLocked_(d) {
  if (d.type !== 'in' && d.type !== 'out') throw new Error('invalid type');
  // 重複打刻チェック（サーバー側絶対防御）
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const rows  = sh(SH_RECORDS).getDataRange().getValues();
  const todayRecsForStaff = rows.slice(1)
    .filter(r => r[1] === d.staff && Utilities.formatDate(new Date(tsStr(r[3])), tz, 'yyyy-MM-dd') === today)
    .sort((a, b) => tsStr(a[3]).localeCompare(tsStr(b[3])));
  const lastType = todayRecsForStaff.length > 0
    ? todayRecsForStaff[todayRecsForStaff.length - 1][2]
    : null;

  if (d.type === 'in' && lastType === 'in') {
    throw new Error('既に出勤打刻されています');
  }
  if (d.type === 'in' && lastType === 'out') {
    throw new Error('本日は既に退勤打刻されています');
  }
  if (d.type === 'out' && lastType !== 'in') {
    throw new Error('出勤打刻されていません');
  }

  const id        = Date.now().toString();
  const timestamp = new Date().toISOString();
  sh(SH_RECORDS).appendRow([id, d.staff, d.type, timestamp]);
  if (d.staff === NAKATA_NAME)   refreshNakataSheet();
  if (d.staff === TAKEUCHI_NAME) refreshTakeuchiSheet();
  if (isMobileUA(d.userAgent)) sendUnauthorizedAlert(d.staff, d.type, d.userAgent, timestamp);
  return { id, staff: d.staff, type: d.type, timestamp };
}

function isMobileUA(ua) {
  if (!ua) return false;
  return /iPhone|Android|Mobile|iPad(?!.*Mac)/i.test(ua);
}

function sendUnauthorizedAlert(staff, type, ua, timestamp) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) return;

  const tz      = Session.getScriptTimeZone();
  const timeStr = Utilities.formatDate(new Date(timestamp), tz, 'yyyy-MM-dd HH:mm:ss');
  const kind    = type === 'in' ? '出勤' : '退勤';
  const device  = /iPhone/i.test(ua) ? 'iPhone' : /Android/i.test(ua) ? 'Android' : 'モバイル';

  const msg = `⚠️ 不正打刻の疑い\n\n${staff}さんが ${kind} 打刻しました。\n\n📱 デバイス: ${device}\n🕐 時刻: ${timeStr}\n\nUA: ${ua.slice(0, 100)}`;

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ messages: [{ type: 'text', text: msg }] }),
    muteHttpExceptions: true,
  });
}

// ─── 退勤打ち忘れチェック（毎日20時トリガー） ───────────
function checkForgottenClockOut() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const rec = ss.getSheetByName(SH_RECORDS);
  if (!rec) return;

  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data  = rec.getDataRange().getValues(); // [id, staff, type, timestamp]

  // 今日の打刻をスタッフ別に集計（最後の type を記録）
  const lastType  = {};
  const inTimeMap = {};

  data.slice(1).forEach(row => {
    const ts    = row[3];
    if (!ts) return;
    const tsDate = Utilities.formatDate(new Date(ts), tz, 'yyyy-MM-dd');
    if (tsDate !== today) return;

    const staff = row[1];
    const type  = row[2];
    lastType[staff] = type;
    if (type === 'in') {
      inTimeMap[staff] = Utilities.formatDate(new Date(ts), tz, 'HH:mm');
    }
  });

  // 退勤していないスタッフを抽出
  const forgotten = Object.keys(lastType).filter(s => lastType[s] === 'in');
  if (forgotten.length === 0) return;

  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) return;

  const lines = forgotten.map(s => `・${s}（出勤 ${inTimeMap[s]}）`).join('\n');
  const msg   = `⚠️ 退勤打ち忘れの可能性\n\n${lines}\n\n確認してください。`;

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ messages: [{ type: 'text', text: msg }] }),
    muteHttpExceptions: true,
  });
}

// ─── 起動データ一括取得（スタッフ＋今日の記録＋今日のメモ）────
function getStartupData(params) {
  const tz    = Session.getScriptTimeZone();
  const today = params.today || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const from  = params.from  || today;
  const to    = params.to    || today;
  return {
    staff:   getStaff(),
    records: getRecords({ from, to }),
    memos:   getAllMemos({ from: today, to: today }),
  };
}

// ─── keep-warm（コールドスタート防止）────────────────
function keepWarm() {
  // 10分ごとのトリガーで呼ばれる。GASを起動状態に保つだけ。
}

// ─── トリガー設定 ────────────────────────────────────
function setupTriggers() {
  // checkForgottenClockOut: 毎日20時
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForgottenClockOut')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForgottenClockOut')
    .timeBased().everyDays(1).atHour(20).create();

  // keepWarm: 10分ごと
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'keepWarm')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('keepWarm')
    .timeBased().everyMinutes(10).create();

  return { message: 'トリガーを設定しました（退勤チェック20時・keepWarm10分毎）' };
}

// ─── 打刻記録 取得 ───────────────────────────────────
function getRecords(params) {
  if (!params.__admin) { const _t = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); params.from = _t; params.to = _t; }  // 2026-07-14 非管理者は当日のみ(全期間漏洩防止)
  const rows = sh(SH_RECORDS).getDataRange().getValues();
  if (rows.length <= 1) return [];

  const tz = Session.getScriptTimeZone();

  let recs = rows.slice(1)
    .map(r => ({ id: r[0].toString(), staff: r[1], type: r[2], timestamp: tsStr(r[3]) }))
    .filter(r => r.id && r.staff && r.timestamp);

  if (params.from) recs = recs.filter(r => Utilities.formatDate(new Date(r.timestamp), tz, 'yyyy-MM-dd') >= params.from);
  if (params.to)   recs = recs.filter(r => Utilities.formatDate(new Date(r.timestamp), tz, 'yyyy-MM-dd') <= params.to);
  return recs;
}

function deleteRecord(d) {
  const sheet = sh(SH_RECORDS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === d.id.toString()) {
      const staffName = rows[i][1];
      sheet.deleteRow(i + 1);
      if (staffName === NAKATA_NAME)   refreshNakataSheet();
      if (staffName === TAKEUCHI_NAME) refreshTakeuchiSheet();
      return { deleted: true };
    }
  }
  return { deleted: false };
}

const SH_ARCHIVE  = '打刻アーカイブ';  // 2026-09-02 clearRecords の退避先（監査指摘：全削除に退避が無かった）

function clearRecords() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('busy: 他の処理中です。しばらくしてから再実行してください');
  try {
    return clearRecordsLocked_();
  } finally {
    lock.releaseLock();
  }
}

function clearRecordsLocked_() {
  const sheet = sh(SH_RECORDS);
  const last  = sheet.getLastRow();
  let archived = 0;
  if (last > 1) {
    // 全削除の前に必ずアーカイブへ退避する（給料の元データが一発で消えるのを防ぐ）。
    // 退避に失敗したら削除しない（fail closed）。
    const rows = sheet.getRange(2, 1, last - 1, 4).getValues();
    const ssApp = SpreadsheetApp.getActiveSpreadsheet();
    let arc = ssApp.getSheetByName(SH_ARCHIVE);
    if (!arc) { arc = ssApp.insertSheet(SH_ARCHIVE); }
    if (arc.getLastRow() === 0) arc.appendRow(['id', 'staff', 'type', 'timestamp', 'archived_at']);  // 名前だけの空シートでもヘッダを作る（Sol指摘）
    const at = new Date().toISOString();
    arc.getRange(arc.getLastRow() + 1, 1, rows.length, 5).setValues(rows.map(r => r.concat([at])));
    archived = rows.length;
    sheet.deleteRows(2, last - 1);
  }
  refreshNakataSheet();
  refreshTakeuchiSheet();
  return { cleared: true, archived };
}

// ─── 打刻時刻の修正（管理者操作）───────────────────
function updateRecordTimestamp(d) {
  const sheet = sh(SH_RECORDS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === d.id.toString()) {
      sheet.getRange(i + 1, 4).setValue(d.timestamp);
      const staffName = rows[i][1];
      if (staffName === NAKATA_NAME)   refreshNakataSheet();
      if (staffName === TAKEUCHI_NAME) refreshTakeuchiSheet();
      return { updated: true, id: d.id, timestamp: d.timestamp };
    }
  }
  return { updated: false };
}

// ─── 代理打刻（管理者操作・重複チェック無効）─────────
function adminClockAction(d) {
  const id = Date.now().toString();
  sh(SH_RECORDS).appendRow([id, d.staff, d.type, d.timestamp]);
  if (d.staff === NAKATA_NAME)   refreshNakataSheet();
  if (d.staff === TAKEUCHI_NAME) refreshTakeuchiSheet();
  return { id, staff: d.staff, type: d.type, timestamp: d.timestamp };
}

// ─── 中田有加 月次シート 自動更新 ────────────────────
// シート名: "中田有加_YYYY-MM"（当月のみ更新）
function refreshNakataSheet() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName   = nakataSheetName();
  const currentMonth = monthKey();
  const tz          = Session.getScriptTimeZone();

  let nakSh = ss.getSheetByName(sheetName);
  if (!nakSh) {
    nakSh = ss.insertSheet(sheetName);
    nakSh.appendRow(['日付', '出勤', '退勤', '日当(円)']);
    nakSh.getRange('D2:D200').setNumberFormat('¥#,##0');
  }

  // 日報メモ読み込み（メモあり → 日当空白、メモなし → ¥5,000）
  const memoSheet = ss.getSheetByName(SH_MEMOS);
  const memoMap   = {};
  if (memoSheet) {
    memoSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] && r[1] === NAKATA_NAME && r[2]) memoMap[r[0].toString()] = true;
    });
  }

  // 打刻記録から当月分を集計（JST基準で日付判定）
  const recSheet = ss.getSheetByName(SH_RECORDS);
  const dayMap   = {};
  if (recSheet) {
    recSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[1] !== NAKATA_NAME || !r[3]) return;
      const ts   = tsStr(r[3]);
      const date = Utilities.formatDate(new Date(ts), tz, 'yyyy-MM-dd');
      if (date.slice(0, 7) !== currentMonth) return;
      if (!dayMap[date]) dayMap[date] = { in: null, out: null };
      if (r[2] === 'in'  && !dayMap[date].in)  dayMap[date].in  = ts;
      if (r[2] === 'out')                       dayMap[date].out = ts;
    });
  }

  const dates = Object.keys(dayMap).sort();
  const lastRow = nakSh.getLastRow();
  if (lastRow > 1) nakSh.deleteRows(2, lastRow - 1);
  if (dates.length === 0) return;

  const rows = dates.map(date => {
    const { in: inTs, out: outTs } = dayMap[date];
    const inTime  = inTs  ? Utilities.formatDate(new Date(inTs),  tz, 'HH:mm') : '';
    const outTime = outTs ? Utilities.formatDate(new Date(outTs), tz, 'HH:mm') : '';
    const allow   = memoMap[date] ? '' : 5000;
    return [date, inTime, outTime, allow];
  });

  nakSh.getRange(2, 1, rows.length, 4).setValues(rows);
  nakSh.getRange(2, 4, rows.length, 1).setNumberFormat('¥#,##0');
}

// ─── 竹内美佳 月次シート 自動更新 ────────────────────
// シート名: "竹内美佳_YYYY-MM"、D列は勤務時間（H:mm）
function refreshTakeuchiSheet() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName    = takeuchiSheetName();
  const currentMonth = monthKey();
  const tz           = Session.getScriptTimeZone();

  const HEADERS = ['日付', '打刻時刻', '出勤', '退勤', '勤務時間', '実働時間（休憩1.5h引き）'];
  let tkSh = ss.getSheetByName(sheetName);
  if (!tkSh) {
    tkSh = ss.insertSheet(sheetName);
    tkSh.appendRow(HEADERS);
  } else {
    // ヘッダーを最新に揃える（既存シートで打刻時刻列が無い場合に対応）
    tkSh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  // 打刻記録から当月分を集計（JST基準で日付判定）
  const recSheet = ss.getSheetByName(SH_RECORDS);
  const dayMap   = {};
  if (recSheet) {
    recSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[1] !== TAKEUCHI_NAME || !r[3]) return;
      const ts   = tsStr(r[3]);
      const date = Utilities.formatDate(new Date(ts), tz, 'yyyy-MM-dd');
      if (date.slice(0, 7) !== currentMonth) return;
      if (!dayMap[date]) dayMap[date] = { in: null, out: null };
      if (r[2] === 'in'  && !dayMap[date].in)  dayMap[date].in  = ts;
      if (r[2] === 'out')                       dayMap[date].out = ts;
    });
  }

  const dates = Object.keys(dayMap).sort();
  const lastRow = tkSh.getLastRow();
  if (lastRow > 1) tkSh.deleteRows(2, lastRow - 1);
  if (dates.length === 0) return;

  const rows = dates.map(date => {
    const { in: rawInTs, out: outTs } = dayMap[date];
    const rawInTime = rawInTs ? Utilities.formatDate(new Date(rawInTs), tz, 'HH:mm') : '';

    // 2026年6月に限り、9時前出勤は9時スタートとして計算
    let inTs = rawInTs;
    if (inTs && date.slice(0, 7) === '2026-06') {
      const inHM = Utilities.formatDate(new Date(inTs), tz, 'HH:mm');
      if (inHM < '09:00') {
        // JST 09:00 = UTC 00:00 同日
        inTs = date + 'T00:00:00.000Z';
      }
    }

    // 2026-06-05 特別ルール: 12:30 スタートとして計算（6月限定の一時ルール）
    if (inTs && date === '2026-06-05') {
      const inHM = Utilities.formatDate(new Date(inTs), tz, 'HH:mm');
      if (inHM < '12:30') {
        // JST 12:30 = UTC 03:30 同日
        inTs = date + 'T03:30:00.000Z';
      }
    }

    const inTime  = inTs  ? Utilities.formatDate(new Date(inTs),  tz, 'HH:mm') : '';
    const outTime = outTs ? Utilities.formatDate(new Date(outTs), tz, 'HH:mm') : '';
    let duration = '';
    let actual   = '';
    if (inTs && outTs) {
      const mins = Math.round((new Date(outTs) - new Date(inTs)) / 60000);
      if (mins > 0) {
        duration = Math.floor(mins / 60) + ':' + String(mins % 60).padStart(2, '0');
        const actualMins = mins - 90; // 休憩1時間30分を引く
        if (actualMins > 0) {
          actual = Math.floor(actualMins / 60) + ':' + String(actualMins % 60).padStart(2, '0');
        }
      }
    }
    return [date, rawInTime, inTime, outTime, duration, actual];
  });

  tkSh.getRange(2, 1, rows.length, 6).setValues(rows);
}

// ISO文字列 or Date オブジェクトを ISO文字列に統一
function tsStr(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString();
  return val.toString();
}

// ─── 日当 取得（全月の中田有加シートから読み取り）────────
// 返値: [{ date, staff, amount }]
function getDailyAllowance(params) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheets  = ss.getSheets().filter(s => s.getName().startsWith(NAKATA_NAME + '_'));
  const results = [];

  sheets.forEach(sheet => {
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    rows.slice(1).filter(r => r[0]).forEach(r => {
      const date = r[0].toString();
      if (params.from && date < params.from) return;
      if (params.to   && date > params.to)   return;
      results.push({ date, staff: NAKATA_NAME, amount: r[3] || '' });
    });
  });

  return results.sort((a, b) => a.date.localeCompare(b.date));
}

// ─── 日報メモ ────────────────────────────────────────
function getMemos(params) {
  if (!params.__admin) { const _t = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); params.from = _t; params.to = _t; }  // 2026-07-14 非管理者は当日のみ
  const rows  = sh(SH_MEMOS).getDataRange().getValues();
  const memos = {};
  rows.slice(1).forEach(r => {
    if (!r[0]) return;
    const date  = r[0].toString();
    const staff = r[1] || '';
    const memo  = r[2] || '';
    if (params.staff && params.staff !== staff) return;
    if (params.from  && date < params.from)     return;
    if (params.to    && date > params.to)       return;
    memos[date] = memo;
  });
  return memos;
}

function getAllMemos(params) {
  if (!params.__admin) { const _t = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); params.from = _t; params.to = _t; }  // 2026-07-14 非管理者は当日のみ
  const rows = sh(SH_MEMOS).getDataRange().getValues();
  const list = [];
  rows.slice(1).forEach(r => {
    if (!r[0]) return;
    const date  = r[0].toString();
    const staff = r[1] || '';
    const memo  = r[2] || '';
    if (params.from && date < params.from) return;
    if (params.to   && date > params.to)   return;
    list.push({ date, staff, memo });
  });
  return list.sort((a, b) => b.date.localeCompare(a.date) || a.staff.localeCompare(b.staff));
}

function saveMemo(d) {
  // 2026-07-14 入力検証：登録済みスタッフのみ・本文長制限（未認証書き込み悪用の緩和）
  if (getStaff().indexOf(d.staff) === -1) throw new Error('unknown staff');
  if (!/^\d{4}-\d{2}-\d{2}$/.test((d.date||'').toString())) throw new Error('bad date');  // 2026-07-14 日付形式を強制(注入源の封鎖)
  if (d.text && d.text.toString().length > 1000) throw new Error('memo too long');
  const sheet = sh(SH_MEMOS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === d.date && rows[i][1] === d.staff) {
      if (d.text && d.text.trim()) {
        sheet.getRange(i + 1, 3).setValue(d.text);
      } else {
        sheet.deleteRow(i + 1);
      }
      if (d.staff === NAKATA_NAME) refreshNakataSheet();
      return { saved: true };
    }
  }
  if (d.text && d.text.trim()) sheet.appendRow([d.date, d.staff, d.text]);
  if (d.staff === NAKATA_NAME) refreshNakataSheet();
  return { saved: true };
}

// ─── スタッフ ────────────────────────────────────────
function getStaff() {
  const rows = sh(SH_STAFF).getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => r[0]).filter(Boolean);
}

function addStaff(d) {
  const current = getStaff();
  if (current.includes(d.name)) return { added: false, reason: '同名のスタッフが既に存在します' };
  sh(SH_STAFF).appendRow([d.name]);
  return { added: true };
}

function removeStaff(d) {
  const sheet = sh(SH_STAFF);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.name) { sheet.deleteRow(i + 1); return { removed: true }; }
  }
  return { removed: false };
}

// ─── PIN ─────────────────────────────────────────────
function checkPin(d) {
  return { valid: pinRateGuard_((d.pin || '').toString()) };  // 2026-07-14 総当たり対策のレートガード経由
}

function changePin(d) {
  const np = (d.newPin || '').toString();
  if (!/^[0-9]{4}$/.test(np)) throw new Error('新PINは4桁の数字にしてください');  // 2026-07-14 サーバー側検証
  d = { newPin: np };
  const sheet = sh(SH_SETTINGS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'pin') { sheet.getRange(i + 1, 2).setValue(d.newPin); return { changed: true }; }
  }
  sheet.appendRow(['pin', d.newPin]);
  return { changed: true };
}
