/**
 * セルクエスト（細胞の構造ゲーム）記録用 Apps Script
 *
 * 記録用スプレッドシートの［拡張機能］→［Apps Script］に貼り付けて使う。
 * index.html を同じプロジェクトに「index」という名前のHTMLファイルとして追加し、
 * ウェブアプリとしてデプロイすると、
 *   - …/exec を開くとゲームが表示され（doGet）、結果は google.script.run で記録される
 *   - HTMLファイルを直接配る場合は、index.html の CONFIG.GAS_URL に …/exec を入れると doPost で記録される
 *
 * 作られるシート
 *   記録      … 1プレイ1行（得点・正答率・誤答した問題）
 *   誤答ログ  … 誤答1件1行（問題ごとの集計用）
 *   誤答集計  … メニュー［セルクエスト］→［誤答集計を更新］で作り直す
 *
 * このファイルの関数名・定数名は CQ_ で始まるもの以外に doGet / doPost / saveResults / onOpen を使う。
 * 既存のスクリプトに追加する場合は、onOpen が重ならないよう CQ_addMenu() を既存の onOpen から呼ぶ。
 */

var CQ_SHEET_LOG = '記録';
var CQ_SHEET_WRONG = '誤答ログ';
var CQ_SHEET_SUMMARY = '誤答集計';
var CQ_LOG_HEADER = ['受信日時', '端末の日時', '講座・クラス', '番号', '氏名', 'ゲーム', 'モード', '得点', '正答数', '出題数', '正答率(%)', '所要時間(秒)', '誤答数', '誤答した問題', 'Googleアカウント'];
var CQ_WRONG_HEADER = ['受信日時', '講座・クラス', '番号', '氏名', 'ゲーム', 'モード', '問題ID', '問題', '生徒の答え', '正解'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('セルクエスト 細胞の構造')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** HTMLファイル版（CONFIG.GAS_URL）からの送信。本文は {results:[…]} のJSON。 */
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var n = CQ_save_(body.results || [], '');
  return ContentService.createTextOutput(JSON.stringify({ ok: true, saved: n }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GAS配信版（google.script.run）からの送信。 */
function saveResults(results) {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (err) { /* 取得できない設定では空欄 */ }
  return CQ_save_(results || [], email);
}

function onOpen() {
  CQ_addMenu();
}

function CQ_addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('セルクエスト')
    .addItem('誤答集計を更新', 'CQ_buildSummary')
    .addToUi();
}

function CQ_save_(results, email) {
  if (!Array.isArray(results)) results = [results];
  results = results.slice(0, 50);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = CQ_sheet_(ss, CQ_SHEET_LOG, CQ_LOG_HEADER);
    var wrongSheet = CQ_sheet_(ss, CQ_SHEET_WRONG, CQ_WRONG_HEADER);
    var now = new Date();
    var logRows = [], wrongRows = [];
    results.forEach(function (r) {
      if (!r || typeof r !== 'object') return;
      var wrong = Array.isArray(r.wrong) ? r.wrong.slice(0, 60) : [];
      var who = [CQ_txt_(r.course), CQ_txt_(r.number), CQ_txt_(r.name)];
      logRows.push([
        now, CQ_txt_(r.clientTime), who[0], who[1], who[2], CQ_txt_(r.game), CQ_txt_(r.mode),
        CQ_num_(r.score), CQ_num_(r.correct), CQ_num_(r.total), CQ_num_(r.rate), CQ_num_(r.seconds),
        wrong.length,
        CQ_txt_(wrong.map(function (w) { return w.id + ' ' + w.q + '（答：' + w.your + '／正：' + w.answer + '）'; }).join('\n'), 5000),
        CQ_txt_(email)
      ]);
      wrong.forEach(function (w) {
        wrongRows.push([now, who[0], who[1], who[2], CQ_txt_(r.game), CQ_txt_(r.mode), CQ_txt_(w.id), CQ_txt_(w.q), CQ_txt_(w.your), CQ_txt_(w.answer)]);
      });
    });
    if (logRows.length) log.getRange(log.getLastRow() + 1, 1, logRows.length, CQ_LOG_HEADER.length).setValues(logRows);
    if (wrongRows.length) wrongSheet.getRange(wrongSheet.getLastRow() + 1, 1, wrongRows.length, CQ_WRONG_HEADER.length).setValues(wrongRows);
    return logRows.length;
  } finally {
    lock.releaseLock();
  }
}

/** 誤答ログを問題IDごとに数え、誤答の多い順に並べる。 */
function CQ_buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(CQ_SHEET_WRONG);
  if (!src || src.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('誤答ログがまだありません。');
    return;
  }
  var values = src.getRange(2, 1, src.getLastRow() - 1, CQ_WRONG_HEADER.length).getValues();
  var map = {};
  values.forEach(function (v) {
    var id = v[6];
    if (!map[id]) map[id] = { game: v[4], q: v[7], answer: v[9], count: 0, students: {}, answers: {} };
    var m = map[id];
    m.count++;
    m.students[v[1] + '|' + v[2] + '|' + v[3]] = true;
    m.answers[v[8]] = (m.answers[v[8]] || 0) + 1;
  });
  var rows = Object.keys(map).map(function (id) {
    var m = map[id];
    var top = Object.keys(m.answers).sort(function (a, b) { return m.answers[b] - m.answers[a]; })
      .slice(0, 3).map(function (a) { return a + '（' + m.answers[a] + '）'; }).join('、');
    return [id, m.game, m.q, m.answer, m.count, Object.keys(m.students).length, top];
  }).sort(function (a, b) { return b[4] - a[4]; });

  var out = ss.getSheetByName(CQ_SHEET_SUMMARY) || ss.insertSheet(CQ_SHEET_SUMMARY);
  out.clear();
  var header = ['問題ID', 'ゲーム', '問題', '正解', '誤答数', '誤答した生徒数', 'よくある誤答（回数）'];
  out.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#dbeafe');
  if (rows.length) out.getRange(2, 1, rows.length, header.length).setValues(rows);
  out.setFrozenRows(1);
  out.setColumnWidth(3, 420);
  out.setColumnWidth(7, 320);
  out.getRange(1, 1, rows.length + 1, header.length).setWrap(true).setVerticalAlignment('top');
  ss.setActiveSheet(out);
}

function CQ_sheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#dbeafe');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 文字列に直し、長さを制限し、= + - @ で始まる値は数式として解釈されないようにする。 */
function CQ_txt_(v, max) {
  var s = (v === null || v === undefined) ? '' : String(v);
  s = s.slice(0, max || 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function CQ_num_(v) {
  var n = Number(v);
  return isFinite(n) ? n : '';
}
