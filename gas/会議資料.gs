/**
 * 単位認定会議資料（成績確認システム連携）
 *
 * 「KS_成績確認システム」の Apps Script プロジェクトに、このファイルを追加して使う。
 * 既存の onOpen() の中に次の1行を足すと、メニュー「単位認定会議資料」が出る。
 *     KG_addMenu();
 *
 * 使うデータ
 *   - 設定シート「取込フォルダURL」内のクラス評定一覧（.xlsx）… 評定・出欠日数
 *   - 実行ログの最新の成績確認結果ブック … 履修不認定・単位不認定（欠時は授業時数の1/3以上）
 *   - 異動者シート … 転学・退学
 *   - 設定シート「推薦生徒一覧URL」のスプレッドシート … クラブ・コース
 * 作ったブックは設定シート「結果フォルダURL」に保存する。
 *
 * このファイルの関数名はすべて KG_ で始まり、既存のコードとは名前がぶつからない。
 */

var KG_DEFAULTS = [
  // [設定シートの項目名, 既定値, 説明]
  ['推薦生徒一覧URL', '', '単位認定会議資料: クラブ推薦生徒一覧（年・組・番・所属）のスプレッドシート'],
  ['登校不調_欠席日数', 10, '単位認定会議資料: 欠席がこの日数以上なら登校不調'],
  ['登校不調_遅刻回数', 10, '単位認定会議資料: 遅刻がこの回数以上なら登校不調'],
  ['評定平均上位の人数', 30, '単位認定会議資料: 学年ごとの上位者一覧の人数（同順位は全員）'],
  ['推薦生徒_要確認の評定平均', 3.0, '単位認定会議資料: 推薦生徒の評定平均値がこれ未満なら要確認']
];

function KG_addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('単位認定会議資料')
    .addItem('会議資料を作成', 'KG_createMeetingDoc')
    .addToUi();
}

function KG_createMeetingDoc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var warnings = [];
  var conf = KG_readSettings_(ss);

  ss.toast('クラス評定一覧を読み込んでいます…', '単位認定会議資料', -1);
  var students = KG_loadClassFiles_(conf.importFolderId, warnings);
  if (!students.length) {
    ui.alert('取込フォルダにクラス評定一覧（.xlsx）が見つかりませんでした。');
    return;
  }
  ss.toast('成績確認結果を読み込んでいます…', '単位認定会議資料', -1);
  var result = KG_loadLatestResult_(ss, conf, warnings);
  var moves = KG_readMoves_(ss);
  var recs = [];
  if (conf.recommendUrl) {
    recs = KG_readRecommend_(SpreadsheetApp.openByUrl(conf.recommendUrl).getSheets()[0].getDataRange().getValues(), warnings);
  } else {
    warnings.push('設定シートの「推薦生徒一覧URL」が空欄のため、推薦生徒・所属別の集計は行いません');
  }

  ss.toast('集計しています…', '単位認定会議資料', -1);
  var model = KG_buildModel_(students, result.rows, moves, recs, conf, warnings);
  model.meta = {
    year: conf.year, term: conf.term, resultUrl: result.url, resultName: result.name,
    files: students.files, created: new Date()
  };
  var out = KG_writeWorkbook_(model, conf, warnings);
  ss.toast('完了しました', '単位認定会議資料', 5);
  ui.alert('単位認定会議資料を作成しました。\n\n' + out.getUrl() +
    (warnings.length ? '\n\n注意（' + warnings.length + '件）は「概要」シートの下にあります。' : ''));
}

// ---------------------------------------------------------------- 設定・入出力（Apps Script）

function KG_readSettings_(ss) {
  var sh = ss.getSheetByName('設定');
  var values = sh.getDataRange().getValues();
  var map = {};
  values.forEach(function (r) { if (r[0] !== '') map[String(r[0]).trim()] = r[1]; });
  // 会議資料用の項目が無ければ既定値で追記する
  var add = KG_DEFAULTS.filter(function (d) { return !(d[0] in map); });
  if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
    add.forEach(function (d) { map[d[0]] = d[1]; });
  }
  var num = function (k, def) { var v = Number(map[k]); return isNaN(v) || map[k] === '' ? def : v; };
  return {
    year: map['年度'],
    term: map['判定時期'],
    importFolderId: KG_idFromUrl_(map['取込フォルダURL']),
    resultFolderId: KG_idFromUrl_(map['結果フォルダURL']),
    recommendUrl: String(map['推薦生徒一覧URL'] || '').trim(),
    absentLimit: num('登校不調_欠席日数', 10),
    lateLimit: num('登校不調_遅刻回数', 10),
    topN: num('評定平均上位の人数', 30),
    recMinAvg: num('推薦生徒_要確認の評定平均', 3.0)
  };
}

function KG_idFromUrl_(url) {
  var m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

/** 取込フォルダの .xlsx（クラス評定一覧）をすべて読む。同じクラスが複数あれば新しいほう。 */
function KG_loadClassFiles_(folderId, warnings) {
  var it = DriveApp.getFolderById(folderId).getFiles();
  var byClass = {};
  var files = [];
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (!/\.xlsx$/i.test(name) || name.indexOf('~$') === 0) continue;
    var parsed;
    try {
      parsed = KG_parseClassRows_(KG_readXlsx_(f.getBlob()), name);
    } catch (e) {
      if (/評定/.test(name)) warnings.push(name + ': 読み込めませんでした（' + e.message + '）');
      continue;
    }
    if (!parsed) continue;
    var prev = byClass[parsed.label];
    if (prev && prev.updated > f.getLastUpdated()) {
      warnings.push(parsed.label + ': クラス評定一覧が複数あります。新しいほう（' + prev.name + '）を使います');
      continue;
    }
    if (prev) warnings.push(parsed.label + ': クラス評定一覧が複数あります。新しいほう（' + name + '）を使います');
    byClass[parsed.label] = { name: name, updated: f.getLastUpdated(), students: parsed.students };
  }
  var students = [];
  Object.keys(byClass).forEach(function (k) {
    files.push(byClass[k].name);
    students = students.concat(byClass[k].students);
  });
  students.files = files.sort();
  return students;
}

/** xlsx の Blob を2次元配列（1枚目のシート）にする。Drive の変換は使わない。 */
function KG_readXlsx_(blob) {
  var parts = {};
  Utilities.unzip(blob.setContentType('application/zip')).forEach(function (b) {
    parts[b.getName()] = b.getDataAsString('UTF-8');
  });
  return KG_xlsxToRows_(parts);
}

/** 実行ログの最新の結果ブック（なければ結果フォルダの最新）から「科目別 不認定者一覧」を読む。 */
function KG_loadLatestResult_(ss, conf, warnings) {
  var url = '';
  var log = ss.getSheetByName('実行ログ');
  if (log && log.getLastRow() > 1) {
    var v = log.getRange(2, 1, log.getLastRow() - 1, log.getLastColumn()).getValues();
    for (var i = v.length - 1; i >= 0 && !url; i--) {
      for (var j = v[i].length - 1; j >= 0; j--) {
        if (/^https:\/\/docs\.google\.com\/spreadsheets\//.test(String(v[i][j]))) { url = String(v[i][j]); break; }
      }
    }
  }
  var book;
  if (url) {
    book = SpreadsheetApp.openByUrl(url);
  } else {
    var it = DriveApp.getFolderById(conf.resultFolderId).searchFiles(
      "title contains '成績確認結果' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    var latest = null;
    while (it.hasNext()) { var f = it.next(); if (!latest || f.getDateCreated() > latest.getDateCreated()) latest = f; }
    if (!latest) {
      warnings.push('成績確認結果ブックが見つかりません。先に成績確認を実行してください（不認定の一覧は空になります）');
      return { rows: [], url: '', name: '' };
    }
    book = SpreadsheetApp.openById(latest.getId());
  }
  var sh = book.getSheetByName('科目別 不認定者一覧');
  if (!sh) {
    warnings.push(book.getName() + ' に「科目別 不認定者一覧」シートがありません');
    return { rows: [], url: book.getUrl(), name: book.getName() };
  }
  return { rows: sh.getDataRange().getDisplayValues(), url: book.getUrl(), name: book.getName() };
}

function KG_readMoves_(ss) {
  var sh = ss.getSheetByName('異動者');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getDisplayValues().slice(1).map(function (r) {
    return { cls: r[0], id: r[1], name: r[2], kind: r[3], date: r[4] };
  }).filter(function (m) { return m.name; });
}

// ---------------------------------------------------------------- xlsx 読み取り（純粋関数）

function KG_xmlText_(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCodePoint(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(Number(d)); })
    .replace(/&amp;/g, '&');
}

function KG_colIndex_(ref) {
  var letters = ref.replace(/\d+/g, '');
  var n = 0;
  for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

/** parts: {パス: XML文字列} → 1枚目のシートの2次元配列（0始まり） */
function KG_xlsxToRows_(parts) {
  var shared = [];
  var ss = parts['xl/sharedStrings.xml'];
  if (ss) {
    (ss.match(/<si>[\s\S]*?<\/si>|<si\/>/g) || []).forEach(function (si) {
      si = si.replace(/<rPh[\s\S]*?<\/rPh>/g, '');   // ふりがなは除く
      var t = '';
      (si.match(/<t(?:\s[^>]*)?>[\s\S]*?<\/t>/g) || []).forEach(function (x) {
        t += KG_xmlText_(x.replace(/^<t[^>]*>/, '').replace(/<\/t>$/, ''));
      });
      shared.push(t);
    });
  }
  // 1枚目のシートのファイル名を workbook.xml と rels から求める
  var sheetPath = 'xl/worksheets/sheet1.xml';
  var wb = parts['xl/workbook.xml'], rels = parts['xl/_rels/workbook.xml.rels'];
  if (wb && rels) {
    var m = wb.match(/<sheet\b[^>]*r:id="([^"]+)"/);
    if (m) {
      var rm = rels.match(new RegExp('<Relationship\\b[^>]*Id="' + m[1] + '"[^>]*>'));
      var tm = rm && rm[0].match(/Target="([^"]+)"/);
      if (tm) sheetPath = tm[1].charAt(0) === '/' ? tm[1].slice(1) : 'xl/' + tm[1].replace(/^\.\//, '');
    }
  }
  var xml = parts[sheetPath];
  if (!xml) throw new Error('シートが見つかりません');
  var rows = [];
  var re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, mm;
  while ((mm = re.exec(xml))) {
    var attrs = mm[1], inner = mm[2] || '';
    var r = attrs.match(/\br="([A-Z]+)(\d+)"/);
    if (!r) continue;
    var row = Number(r[2]) - 1, col = KG_colIndex_(r[1]);
    var type = (attrs.match(/\bt="([^"]+)"/) || [])[1];
    var vm = inner.match(/<v>([\s\S]*?)<\/v>/);
    var val = '';
    if (type === 's') val = vm ? shared[Number(vm[1])] : '';
    else if (type === 'inlineStr') {
      val = '';
      (inner.match(/<t(?:\s[^>]*)?>[\s\S]*?<\/t>/g) || []).forEach(function (x) {
        val += KG_xmlText_(x.replace(/^<t[^>]*>/, '').replace(/<\/t>$/, ''));
      });
    } else if (type === 'str' || type === 'e') val = vm ? KG_xmlText_(vm[1]) : '';
    else if (type === 'b') val = vm ? vm[1] === '1' : '';
    else val = vm ? Number(vm[1]) : '';
    if (val === '') continue;
    while (rows.length <= row) rows.push([]);
    rows[row][col] = val;
  }
  var width = rows.reduce(function (w, r) { return Math.max(w, r.length); }, 0);
  return rows.map(function (r) {
    var out = [];
    for (var i = 0; i < width; i++) out.push(r[i] === undefined ? '' : r[i]);
    return out;
  });
}

// ---------------------------------------------------------------- クラス評定一覧の解析（純粋関数）

var KG_ATTEND = ['授業日数', '出停・忌引日数', '出席すべき日数', '欠席日数', '出席日数', '遅刻', '早退', '学期末評定平均'];

function KG_nfkc_(s) { return String(s === null || s === undefined ? '' : s).normalize('NFKC').trim(); }

function KG_num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  var s = KG_nfkc_(v);
  if (s === '' || isNaN(Number(s))) return null;
  return Number(s);
}

/** 「1」「1組」「FA１」→ "1" / "FA1" */
function KG_normCls_(v) { return KG_nfkc_(v).replace(/組$/, ''); }

function KG_nameKey_(s) { return KG_nfkc_(s).replace(/[\s　]/g, ''); }

/** 帳票の2次元配列 → {label, grade, cls, students[]}。クラス評定一覧でなければ null。 */
function KG_parseClassRows_(rows, fileName) {
  var headerRows = [];
  rows.forEach(function (r, i) { if (r[1] === '科目名') headerRows.push(i); });
  if (!headerRows.length) return null;
  var grade = null, cls = null;
  for (var i = 0; i < Math.min(15, rows.length) && grade === null; i++) {
    for (var j = 0; j < rows[i].length; j++) {
      var m = KG_nfkc_(rows[i][j]).match(/^(\d+)年(\S+?)組$/);
      if (m) { grade = Number(m[1]); cls = KG_normCls_(m[2]); break; }
    }
  }
  if (grade === null) {
    var fm = KG_nfkc_(fileName).match(/第(\d+)学年(.+?)組/);
    if (!fm) throw new Error('クラスが分かりません');
    grade = Number(fm[1]); cls = KG_normCls_(fm[2]);
  }
  var label = grade + '年' + cls + '組';
  var byNo = {}, order = [];
  headerRows.forEach(function (h) {
    var hdr = rows[h];
    var subj = [];
    for (var c = 3; c < Math.min(61, hdr.length); c += 3) if (hdr[c] !== '') subj.push([c, String(hdr[c]).trim()]);
    var att = {};
    hdr.forEach(function (v, c) { var s = String(v).trim(); if (KG_ATTEND.indexOf(s) >= 0) att[s] = c; });
    var i = h + 1;
    while (i < rows.length && rows[i][1] !== 'No.') i++;
    for (i++; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[2]).indexOf('学期末評定') === 0 || r[1] === '科目名') break;
      var no = KG_num_(r[1]);
      if (no === null || r[2] === '') continue;
      var st = byNo[no];
      if (!st) {
        st = byNo[no] = { grade: grade, cls: cls, label: label, no: no, name: String(r[2]).trim(), subjects: [], att: {} };
        order.push(no);
      }
      subj.forEach(function (p) {
        var c = p[0];
        var ev = KG_num_(r[c]), kan = String(r[c + 1] === undefined ? '' : r[c + 1]).trim(), ab = KG_num_(r[c + 2]);
        if (ev === null && kan === '' && ab === null) return;   // 履修していない
        st.subjects.push({ name: String(p[1]).trim(), grade: ev, kanten: kan, absent: ab });
      });
      Object.keys(att).forEach(function (k) { if (st.att[k] === undefined || st.att[k] === null) st.att[k] = KG_num_(r[att[k]]); });
    }
  });
  return { label: label, grade: grade, cls: cls, students: order.map(function (n) { return byNo[n]; }) };
}

// ---------------------------------------------------------------- 推薦生徒一覧（純粋関数）

/** 推薦生徒一覧での組の書き方をクラス評定一覧に合わせる（Nu1→N1、G1→GC1） */
var KG_RECOMMEND_CLASS_ALIAS = [[/^Nu(\d+)$/i, 'N$1'], [/^G(\d+)$/, 'GC$1']];

function KG_recCls_(v) {
  var c = KG_normCls_(v);
  KG_RECOMMEND_CLASS_ALIAS.forEach(function (a) { c = c.replace(a[0], a[1]); });
  return c;
}

/** 推薦生徒一覧の値（見出し: 年・組・番・所属）→ [{grade, label, no, raw, org}] */
function KG_readRecommend_(values, warnings) {
  var alias = { grade: ['年', '学年'], cls: ['組', 'クラス'], no: ['番', '番号', '出席番号'], org: ['所属', 'クラブ', '部活動', 'コース'] };
  var hi = -1, col = {};
  for (var i = 0; i < Math.min(20, values.length) && hi < 0; i++) {
    var cells = values[i].map(function (v) { return String(v).trim(); });
    var found = {};
    Object.keys(alias).forEach(function (k) {
      for (var a = 0; a < alias[k].length; a++) { var idx = cells.indexOf(alias[k][a]); if (idx >= 0) { found[k] = idx; break; } }
    });
    if (Object.keys(found).length === 4) { hi = i; col = found; }
  }
  if (hi < 0) { warnings.push('推薦生徒一覧に「年・組・番・所属」の見出しが見つかりません'); return []; }
  var out = [];
  values.slice(hi + 1).forEach(function (r) {
    var raw = String(r[col.org]).trim();
    if (!raw) return;
    var g = KG_num_(String(r[col.grade]).replace('年', '')), n = KG_num_(r[col.no]);
    if (g === null || n === null) return;
    raw.split(/[、,，]/).forEach(function (part) {
      part = part.trim();
      if (!part) return;
      out.push({ grade: g, label: g + '年' + KG_recCls_(r[col.cls]) + '組', no: n, raw: part });
    });
  });
  // 表記ゆれの統一: 印を除き、「女子サッカー」→「女子サッカー部」のように既存の部名へ寄せる
  var strip = function (s) { return KG_nfkc_(s).replace(/^[◎○●☆★◇◆]+/, '').trim(); };
  var canon = {};
  out.forEach(function (x) { var s = strip(x.raw); if (/部$/.test(s)) canon[s] = true; });
  var names = Object.keys(canon);
  out.forEach(function (x) {
    var s = strip(x.raw);
    var org = s;
    if (!canon[s]) {
      if (canon[s + '部']) org = s + '部';
      else {
        var hits = names.filter(function (c) { return c.indexOf(s) === 0; });
        if (hits.length === 1) org = hits[0];
      }
    }
    x.org = org;
    x.kind = /^FA\//i.test(org) ? 'コース' : 'クラブ';
  });
  return out;
}

// ---------------------------------------------------------------- 集計（純粋関数）

function KG_round_(x, d) { var p = Math.pow(10, d); return Math.round(x * p) / p; }

function KG_mean_(a) { return a.length ? a.reduce(function (s, x) { return s + x; }, 0) / a.length : null; }

function KG_buildModel_(students, failRows, moves, recs, conf, warnings) {
  var key = function (label, no) { return label + '-' + no; };
  var st = {};
  students.forEach(function (s) {
    var g = s.subjects.filter(function (x) { return x.grade !== null; }).map(function (x) { return x.grade; });
    s.key = key(s.label, s.no);
    s.course = KG_course_(s.cls);
    s.count = g.length;
    s.sum = g.reduce(function (a, b) { return a + b; }, 0);
    s.avg = g.length ? KG_round_(s.sum / g.length, 4) : null;
    s.n1 = g.filter(function (x) { return x === 1; }).length;
    s.n2 = g.filter(function (x) { return x === 2; }).length;
    s.fails = [];
    s.orgs = [];
    var a = s.att;
    s.absent = a['欠席日数'] || 0; s.late = a['遅刻'] || 0; s.early = a['早退'] || 0;
    st[s.key] = s;
  });

  // 異動者（氏名で照合）
  var moveByName = {};
  moves.forEach(function (m) { moveByName[KG_nameKey_(m.name)] = m; });
  students.forEach(function (s) {
    var m = moveByName[KG_nameKey_(s.name)];
    s.move = m ? m.kind + '（' + m.date + '）' : '';
    s.status = s.move ? s.move : (s.count === 0 ? '評定なし（要確認）' : '');
  });

  // 不認定（成績確認結果の「科目別 不認定者一覧」）
  var fails = [];
  if (failRows.length) {
    var h = failRows[0];
    var ci = function (name) { return h.indexOf(name); };
    var c = { subj: ci('科目名'), kind: ci('区分'), cls: ci('クラス'), no: ci('番号'), name: ci('氏名'), abs: ci('欠時数／上限'), grade: ci('評定') };
    failRows.slice(1).forEach(function (r) {
      if (!r[c.cls]) return;
      var label = KG_nfkc_(r[c.cls]);
      var f = { label: label, no: KG_num_(r[c.no]), name: r[c.name], kind: r[c.kind], subj: r[c.subj], abs: r[c.abs], grade: r[c.grade] };
      var s = st[key(label, f.no)];
      if (s) s.fails.push(f);
      f.student = s || null;
      fails.push(f);
    });
  }

  // 推薦生徒
  var unmatched = [];
  recs.forEach(function (r) {
    var s = st[key(r.label, r.no)];
    if (!s) { unmatched.push(r.label + r.no + '番（' + r.org + '）'); return; }
    if (s.orgs.indexOf(r.org) < 0) s.orgs.push(r.org);
    r.student = s;
  });
  var loadedClasses = {};
  students.forEach(function (s) { loadedClasses[s.label] = true; });
  var unmatchedLoaded = unmatched.filter(function (u) { return loadedClasses[u.replace(/\d+番（.*$/, '')]; });
  if (unmatchedLoaded.length) warnings.push('推薦生徒一覧のうち、クラス評定一覧に見つからない生徒: ' + unmatchedLoaded.join('、'));

  // 学年ごとの順位・平均
  var grades = {};
  students.forEach(function (s) { (grades[s.grade] = grades[s.grade] || []).push(s); });
  var gradeKeys = Object.keys(grades).map(Number).sort();
  var gradeAvg = {};
  gradeKeys.forEach(function (g) {
    var list = grades[g].filter(function (s) { return s.avg !== null && !s.move; });
    gradeAvg[g] = KG_mean_(list.map(function (s) { return s.avg; }));
    list.forEach(function (s) {
      s.rank = list.filter(function (o) { return o.avg > s.avg; }).length + 1;
      s.rankOf = list.length;
    });
  });

  students.forEach(function (s) {
    s.perfect = s.count > 0 && !s.move && s.absent === 0 && s.late === 0 && s.early === 0;
    s.poor = s.absent >= conf.absentLimit || s.late >= conf.lateLimit;
    s.poorWhy = s.poor ? (s.absent >= conf.absentLimit && s.late >= conf.lateLimit ? '欠席・遅刻' : s.absent >= conf.absentLimit ? '欠席' : '遅刻') : '';
    s.nFail = s.fails.length;
    s.nRishu = s.fails.filter(function (f) { return f.kind === '履修不認定'; }).length;
    s.nTani = s.fails.filter(function (f) { return f.kind === '単位不認定'; }).length;
    if (s.orgs.length) {
      var why = [];
      if (s.avg === null) why.push('評定なし');
      else if (s.avg < conf.recMinAvg) why.push('評定平均' + conf.recMinAvg.toFixed(1) + '未満');
      if (s.n1) why.push('評定1が' + s.n1 + '科目');
      if (s.nFail) why.push('不認定' + s.nFail + '科目');
      if (s.poor) why.push('登校不調（' + s.poorWhy + '）');
      s.recCheck = why.join('、');
    }
  });

  // 所属別（クラブ・コース）
  var orgRows = [];
  var orgs = {};
  recs.forEach(function (r) { if (r.student) { orgs[r.org] = orgs[r.org] || { kind: r.kind, list: [] }; if (orgs[r.org].list.indexOf(r.student) < 0) orgs[r.org].list.push(r.student); } });
  var orgNames = Object.keys(orgs).sort(function (a, b) {
    return (orgs[a].kind === orgs[b].kind ? 0 : orgs[a].kind === 'クラブ' ? -1 : 1) || (a < b ? -1 : a > b ? 1 : 0);
  });
  orgNames.forEach(function (o) {
    var groups = [];
    gradeKeys.forEach(function (g) {
      var l = orgs[o].list.filter(function (s) { return s.grade === g; });
      if (l.length) groups.push([g, l]);
    });
    groups.forEach(function (gr) {
      var l = gr[1], withAvg = l.filter(function (s) { return s.avg !== null; });
      var avgs = withAvg.map(function (s) { return s.avg; });
      var diffs = withAvg.map(function (s) { return s.avg - gradeAvg[s.grade]; });
      orgRows.push({
        kind: orgs[o].kind, org: o, grade: gr[0], n: l.length,
        avg: KG_mean_(avgs), max: avgs.length ? Math.max.apply(null, avgs) : null,
        min: avgs.length ? Math.min.apply(null, avgs) : null, diff: KG_mean_(diffs),
        below: withAvg.filter(function (s) { return s.avg < gradeAvg[s.grade]; }).length,
        check: l.filter(function (s) { return s.recCheck; }).length,
        perfect: l.filter(function (s) { return s.perfect; }).length,
        poor: l.filter(function (s) { return s.poor; }).length
      });
    });
  });
  var aliasRows = [];
  var seen = {};
  recs.forEach(function (r) {
    var s = KG_nfkc_(r.raw);
    if (s !== r.org && !seen[s]) { seen[s] = true; aliasRows.push([r.raw, r.org]); }
  });

  // 科目別評定分布（学年別）
  var subjDist = [];
  gradeKeys.forEach(function (g) {
    var bySubj = {}, order = [];
    grades[g].forEach(function (s) {
      s.subjects.forEach(function (x) {
        if (!bySubj[x.name]) { bySubj[x.name] = { n: 0, c: [0, 0, 0, 0, 0, 0], sum: 0, blank: 0 }; order.push(x.name); }
        var b = bySubj[x.name];
        if (x.grade === null) { b.blank++; return; }
        b.n++; b.sum += x.grade;
        if (x.grade >= 1 && x.grade <= 5) b.c[x.grade]++;
      });
    });
    order.forEach(function (name) {
      var b = bySubj[name];
      if (!b.n) return;   // 評定を付けない科目
      subjDist.push([g + '年', name, b.n, b.c[5], b.c[4], b.c[3], b.c[2], b.c[1], b.sum / b.n,
        b.c[5] / b.n, (b.c[1] + b.c[2]) / b.n, b.blank]);
    });
  });

  return { students: students, grades: grades, gradeKeys: gradeKeys, gradeAvg: gradeAvg, fails: fails,
    recs: recs, orgRows: orgRows, aliasRows: aliasRows, subjDist: subjDist, conf: conf };
}

/** コース（組の表記から判定） */
var KG_COURSES = [
  ['自己創造／ウォラーレ', /^\d+$/],
  ['GC／CBE', /^(GC|CBE)\d*$/],
  ['SS', /^SS\d*$/],
  ['FA', /^FA\d*$/],
  ['看護科', /^N\d*$/]
];

function KG_course_(cls) {
  for (var i = 0; i < KG_COURSES.length; i++) if (KG_COURSES[i][1].test(String(cls))) return KG_COURSES[i][0];
  return 'その他';
}

function KG_courseOrder_(c) {
  for (var i = 0; i < KG_COURSES.length; i++) if (KG_COURSES[i][0] === c) return i;
  return KG_COURSES.length;
}

/** 会議資料の各表を作る。学年別の表は {学年: [見出し, 行...]} の形。 */
function KG_tables_(model) {
  var conf = model.conf;
  var sortSt = function (a, b) { return a.grade - b.grade || KG_clsOrder_(a.cls, b.cls) || a.no - b.no; };
  var all = model.students.slice().sort(sortSt);
  var orgsOf = function (s) { return s.orgs.join('、'); };
  var avgOf = function (s) { return s.avg === null ? '' : s.avg; };
  var rankOf = function (s) { return s.rank ? s.rank + '／' + s.rankOf : ''; };
  var byGrade = function (header, rowsOf) {
    var out = {};
    model.gradeKeys.forEach(function (g) { out[g] = [header].concat(rowsOf(g)); });
    return out;
  };
  var inGrade = function (g) { return all.filter(function (s) { return s.grade === g; }); };
  var T = {};

  // 概要（学年×コース、学年計、全体計）
  var sumRow = function (g, course, list) {
    var graded = list.filter(function (s) { return s.avg !== null && !s.move; });
    var avg = KG_mean_(graded.map(function (s) { return s.avg; }));
    var rec = list.filter(function (s) { return s.orgs.length; });
    var recAvg = KG_mean_(rec.filter(function (s) { return s.avg !== null; }).map(function (s) { return s.avg; }));
    var ga = typeof g === 'number' ? model.gradeAvg[g] : null;
    var cnt = function (kind) {
      return list.reduce(function (n, s) { return n + s.fails.filter(function (f) { return f.kind === kind; }).length; }, 0);
    };
    return [typeof g === 'number' ? g + '年' : g, course, list.length, graded.length, avg === null ? '' : avg,
      avg === null || ga === null || course === '学年計' ? '' : avg - ga,
      cnt('履修不認定'), cnt('単位不認定'), list.filter(function (s) { return s.nFail; }).length,
      list.filter(function (s) { return s.perfect; }).length, list.filter(function (s) { return s.poor; }).length,
      rec.length, recAvg === null ? '' : recAvg, rec.filter(function (s) { return s.recCheck; }).length];
  };
  T.summary = [['学年', 'コース', '在籍（帳票）', '評定のある生徒', '評定平均値', '学年平均との差', '履修不認定（件）',
    '単位不認定（件）', '不認定のある生徒', '皆勤者', '登校不調者', '推薦生徒', '推薦生徒の評定平均値', '推薦生徒 要確認']];
  T.summaryTotals = [];
  model.gradeKeys.forEach(function (g) {
    var l = inGrade(g);
    var courses = [];
    l.forEach(function (s) { if (courses.indexOf(s.course) < 0) courses.push(s.course); });
    courses.sort(function (a, b) { return KG_courseOrder_(a) - KG_courseOrder_(b); });
    courses.forEach(function (c) { T.summary.push(sumRow(g, c, l.filter(function (s) { return s.course === c; }))); });
    T.summary.push(sumRow(g, '学年計', l));
    T.summaryTotals.push(T.summary.length - 1);
  });
  T.summary.push(sumRow('全学年', '合計', all));
  T.summaryTotals.push(T.summary.length - 1);

  // 不認定一覧（成績確認結果より）
  T.fails = byGrade(['コース', 'クラス', '番号', '氏名', '区分', '科目名', '欠時数／上限', '評定', '推薦（所属）', '備考'], function (g) {
    return model.fails.filter(function (f) { return KG_gradeOfLabel_(f.label) === g; }).sort(function (a, b) {
      return KG_labelOrder_(a.label, b.label) || (a.no - b.no) || (a.kind < b.kind ? -1 : a.kind > b.kind ? 1 : 0);
    }).map(function (f) {
      var s = f.student;
      return [KG_course_(KG_clsOfLabel_(f.label)), f.label, f.no, f.name, f.kind, f.subj, f.abs, f.grade,
        s ? orgsOf(s) : '', s ? s.status : ''];
    });
  });

  // 評定平均上位（不認定科目のある生徒を除いて順位を付ける）
  T.top = byGrade(['順位', 'コース', 'クラス', '番号', '氏名', '評定平均値', '科目数', '推薦（所属）'], function (g) {
    var elig = inGrade(g).filter(function (s) { return s.avg !== null && !s.move && !s.nFail; });
    elig.forEach(function (s) { s.topRank = elig.filter(function (o) { return o.avg > s.avg; }).length + 1; });
    return elig.filter(function (s) { return s.topRank <= conf.topN; })
      .sort(function (a, b) { return a.topRank - b.topRank || sortSt(a, b); })
      .map(function (s) { return [s.topRank, s.course, s.label, s.no, s.name, s.avg, s.count, orgsOf(s)]; });
  });

  T.perfect = byGrade(['コース', 'クラス', '番号', '氏名', '評定平均値', '学年内順位', '推薦（所属）'], function (g) {
    return inGrade(g).filter(function (s) { return s.perfect; }).map(function (s) {
      return [s.course, s.label, s.no, s.name, avgOf(s), rankOf(s), orgsOf(s)];
    });
  });

  T.poor = byGrade(['コース', 'クラス', '番号', '氏名', '出席すべき日数', '欠席', '遅刻', '早退', '該当理由', '評定平均値',
    '不認定', '推薦（所属）', '備考'], function (g) {
    return inGrade(g).filter(function (s) { return s.poor; }).map(function (s) {
      return [s.course, s.label, s.no, s.name, s.att['出席すべき日数'], s.absent, s.late, s.early, s.poorWhy,
        avgOf(s), s.nFail || '', orgsOf(s), s.status];
    });
  });

  // 推薦生徒の成績確認（個人）
  T.rec = byGrade(['種別', '所属', 'コース', 'クラス', '番号', '氏名', '評定平均値', '学年平均との差', '学年内順位', '評定1', '評定2',
    '履修不認定', '単位不認定', '欠席', '遅刻', '要確認の理由'], function (g) {
    return model.recs.filter(function (r) { return r.student && r.student.grade === g; }).slice().sort(function (a, b) {
      return (a.kind === b.kind ? 0 : a.kind === 'クラブ' ? -1 : 1) || (a.org < b.org ? -1 : a.org > b.org ? 1 : 0) ||
        ((b.student.avg || 0) - (a.student.avg || 0));
    }).map(function (r) {
      var s = r.student;
      return [r.kind, r.org, s.course, s.label, s.no, s.name, avgOf(s),
        s.avg === null ? '' : s.avg - model.gradeAvg[s.grade], rankOf(s),
        s.n1 || '', s.n2 || '', s.nRishu || '', s.nTani || '', s.absent, s.late, s.recCheck || ''];
    });
  });

  T.org = byGrade(['種別', '所属', '人数', '評定平均値', '最高', '最低', '学年平均との差', '学年平均未満', '要確認', '皆勤', '登校不調'],
    function (g) {
      return model.orgRows.filter(function (o) { return o.grade === g; }).map(function (o) {
        return [o.kind, o.org, o.n, o.avg === null ? '' : o.avg, o.max === null ? '' : o.max,
          o.min === null ? '' : o.min, o.diff === null ? '' : o.diff, o.below, o.check, o.perfect, o.poor];
      });
    });
  T.alias = [['一覧での表記', '集計に使った所属名']].concat(model.aliasRows);

  T.subj = byGrade(['科目', '評定人数', '評定5', '評定4', '評定3', '評定2', '評定1', '評定平均', '5の割合', '1・2の割合', '評定空欄'],
    function (g) {
      return model.subjDist.filter(function (r) { return r[0] === g + '年'; }).map(function (r) { return r.slice(1); });
    });

  T.cls = byGrade(['コース', 'クラス', '在籍（帳票）', '評定平均値', '履修不認定（件）', '単位不認定（件）', '皆勤者', '登校不調者',
    '推薦生徒', '評定なし・異動'], function (g) {
    var l = inGrade(g);
    var labels = [];
    l.forEach(function (s) { if (labels.indexOf(s.label) < 0) labels.push(s.label); });
    return labels.map(function (lb) {
      var c = l.filter(function (s) { return s.label === lb; });
      var f = model.fails.filter(function (x) { return x.label === lb; });
      var avg = KG_mean_(c.filter(function (s) { return s.avg !== null && !s.move; }).map(function (s) { return s.avg; }));
      return [c[0].course, lb, c.length, avg === null ? '' : avg,
        f.filter(function (x) { return x.kind === '履修不認定'; }).length, f.filter(function (x) { return x.kind === '単位不認定'; }).length,
        c.filter(function (s) { return s.perfect; }).length, c.filter(function (s) { return s.poor; }).length,
        c.filter(function (s) { return s.orgs.length; }).length, c.filter(function (s) { return s.status; }).length];
    });
  });

  T.all = byGrade(['コース', 'クラス', '番号', '氏名', '評定平均値', '評定合計', '科目数', '学年内順位', '評定1', '評定2', '履修不認定',
    '単位不認定', '授業日数', '出席すべき日数', '欠席', '遅刻', '早退', '皆勤', '登校不調', '推薦（所属）', '状態'], function (g) {
    return inGrade(g).map(function (s) {
      return [s.course, s.label, s.no, s.name, avgOf(s), s.sum, s.count, rankOf(s),
        s.n1, s.n2, s.nRishu, s.nTani, s.att['授業日数'], s.att['出席すべき日数'], s.absent, s.late, s.early,
        s.perfect ? '○' : '', s.poor ? '○' : '', orgsOf(s), s.status];
    });
  });
  return T;
}

function KG_gradeOfLabel_(label) { var m = String(label).match(/^(\d+)年/); return m ? Number(m[1]) : null; }

function KG_clsOfLabel_(label) { var m = String(label).match(/^\d+年(.+)組$/); return m ? m[1] : ''; }

function KG_clsOrder_(a, b) {
  var na = /^\d+$/.test(a), nb = /^\d+$/.test(b);
  if (na && nb) return Number(a) - Number(b);
  var ca = KG_courseOrder_(KG_course_(a)), cb = KG_courseOrder_(KG_course_(b));
  if (ca !== cb) return ca - cb;
  return a < b ? -1 : a > b ? 1 : 0;
}

function KG_labelOrder_(a, b) {
  return ((KG_gradeOfLabel_(a) || 0) - (KG_gradeOfLabel_(b) || 0)) || KG_clsOrder_(KG_clsOfLabel_(a), KG_clsOfLabel_(b));
}

// ---------------------------------------------------------------- 出力（Apps Script）

function KG_writeWorkbook_(model, conf, warnings) {
  var T = KG_tables_(model);
  var meta = model.meta;
  var tz = Session.getScriptTimeZone();
  var book = SpreadsheetApp.create(meta.year + '年度_' + meta.term + '_単位認定会議資料_' +
    Utilities.formatDate(meta.created, tz, 'yyyyMMdd-HHmm'));
  if (conf.resultFolderId) DriveApp.getFileById(book.getId()).moveTo(DriveApp.getFolderById(conf.resultFolderId));
  var first = book.getSheets()[0];

  var put = function (sh, row, table, fmt) {
    var w = table[0].length;
    sh.getRange(row, 1, table.length, w).setValues(table);
    sh.getRange(row, 1, 1, w).setFontWeight('bold').setBackground('#d9e1f2').setWrap(true)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 1, table.length, w).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    if (table.length > 1) {
      Object.keys(fmt || {}).forEach(function (c) {
        sh.getRange(row + 1, Number(c), table.length - 1, 1).setNumberFormat(fmt[c]);
      });
    }
    return row + table.length;
  };
  /** 学年ごとに「■ ○年」の見出しと表を縦に並べる。該当なしの学年も見出しは出す。 */
  var putByGrade = function (sh, row, tables, fmt, after) {
    model.gradeKeys.forEach(function (g) {
      var t = tables[g];
      sh.getRange(row, 1).setValue('■ ' + g + '年（' + (t.length - 1) + '件）').setFontWeight('bold').setFontSize(12);
      if (t.length === 1) {
        put(sh, row + 1, t);
        sh.getRange(row + 2, 1).setValue('該当なし');
        row += 4;
        return;
      }
      var end = put(sh, row + 1, t, fmt);
      if (after) after(sh, row + 2, t);
      row = end + 2;
    });
    return row;
  };
  var title = function (sh, text, note) {
    sh.getRange(1, 1).setValue(text).setFontSize(14).setFontWeight('bold');
    if (note) sh.getRange(2, 1).setValue(note).setFontColor('#595959').setFontSize(9);
  };
  var finish = function (sh, width1) {
    sh.autoResizeColumns(1, sh.getLastColumn());
    sh.setColumnWidth(1, width1 || 100);   // A1 の見出し文字で A 列が広がりすぎないように
  };
  var sheet = function (name, heading, note, tables, fmt, after) {
    var sh = book.insertSheet(name);
    title(sh, heading, note);
    var r = putByGrade(sh, 4, tables, fmt, after);
    finish(sh);
    return { sh: sh, row: r };
  };

  // 概要
  var s0 = first.setName('概要');
  title(s0, '単位認定会議資料（' + meta.year + '年度 ' + meta.term + '）',
    '作成: ' + Utilities.formatDate(meta.created, tz, 'yyyy/MM/dd HH:mm') + '　取扱注意（個人情報）');
  s0.getRange(3, 1).setValue('学年・コース別 集計').setFontWeight('bold').setFontSize(12);
  var r = put(s0, 4, T.summary, { 5: '0.00', 6: '+0.00;-0.00;0.00', 13: '0.00' });
  T.summaryTotals.forEach(function (i) { s0.getRange(4 + i, 1, 1, T.summary[0].length).setFontWeight('bold').setBackground('#f2f2f2'); });
  var notes = [
    ['判定基準'],
    ['・コース: 自己創造／ウォラーレ＝1～5組、GC／CBE＝GC1組・CBE1組、SS＝SS1組、FA＝FA1組・FA2組、看護科＝N1組'],
    ['・履修不認定・単位不認定: 成績確認システムの判定（欠時数が授業時数の1/3以上＝履修不認定）。元: ' + (meta.resultName || '（なし）')],
    ['・評定平均値: 評定のついた科目の単純平均。評定のない生徒・異動者は順位と平均から除く'],
    ['・評定平均上位: 不認定科目（履修・単位）のある生徒を除いて順位を付ける'],
    ['・皆勤: 欠席・遅刻・早退がすべて0　／　登校不調: 欠席' + conf.absentLimit + '日以上 または 遅刻' + conf.lateLimit + '回以上'],
    ['・推薦生徒の要確認: 評定平均値' + conf.recMinAvg.toFixed(1) + '未満、評定1、不認定、登校不調のいずれか（基準は設定シートで変更可）'],
    ['・クラス評定一覧: ' + meta.files.length + 'ファイル']
  ];
  s0.getRange(r + 1, 1, notes.length, 1).setValues(notes);
  s0.getRange(r + 1, 1).setFontWeight('bold');
  r += notes.length + 2;
  if (meta.resultUrl) s0.getRange(r++, 1).setValue('成績確認結果: ' + meta.resultUrl);
  if (warnings.length) {
    s0.getRange(r + 1, 1).setValue('注意（' + warnings.length + '件）').setFontWeight('bold').setFontColor('#c00000');
    s0.getRange(r + 2, 1, warnings.length, 1).setValues(warnings.map(function (w) { return [w]; }));
  }
  s0.setFrozenRows(4);
  finish(s0, 60);

  sheet('不認定一覧', '履修不認定・単位不認定一覧（学年別）', '成績確認システムの最新の結果から。クラス・番号順', T.fails);
  sheet('評定平均上位', '評定平均値 上位' + conf.topN + '名（学年別）',
    '不認定科目のある生徒・評定のない生徒・異動者を除く。同順位は全員掲載', T.top, { 6: '0.00' });
  sheet('皆勤者一覧', '皆勤者一覧（学年別・評定平均値付き）', '欠席・遅刻・早退がすべて0', T.perfect, { 5: '0.00' });
  sheet('登校不調者一覧', '登校不調者一覧（学年別）',
    '欠席' + conf.absentLimit + '日以上 または 遅刻' + conf.lateLimit + '回以上', T.poor, { 10: '0.00' });
  sheet('推薦生徒 成績確認', '推薦生徒の成績確認（学年別）',
    '所属ごと・評定平均値の高い順。「要確認の理由」がある行は色付き', T.rec, { 7: '0.00', 8: '+0.00;-0.00;0.00' },
    function (sh, row, t) {
      var bg = t.slice(1).map(function (x) {
        var c = x[x.length - 1] ? '#fce4e4' : null;
        return x.map(function () { return c; });
      });
      sh.getRange(row, 1, bg.length, t[0].length).setBackgrounds(bg);
    });
  var so = sheet('所属別 評定平均', '所属別（クラブ・コース区分）評定平均値（学年別）',
    '推薦生徒一覧の所属で集計。学年平均との差は、各生徒の（評定平均値－その学年の平均）の平均', T.org,
    { 4: '0.00', 5: '0.00', 6: '0.00', 7: '+0.00;-0.00;0.00' });
  if (T.alias.length > 1) {
    so.sh.getRange(so.row, 1).setValue('所属名の統一（◎の有無・表記ゆれ）').setFontWeight('bold');
    put(so.sh, so.row + 1, T.alias);
  }
  sheet('科目別評定分布', '科目別 評定分布（学年別）', '評定を付けない科目は除く', T.subj, { 8: '0.00', 9: '0%', 10: '0%' });
  sheet('クラス別比較', 'クラス別比較（学年別）', '', T.cls, { 4: '0.00' });
  sheet('生徒一覧', '生徒一覧（学年別・基礎データ）', '', T.all, { 5: '0.00' });

  book.setActiveSheet(s0);
  return book;
}

// Node でのテスト用（Apps Script では module が無いので何もしない）
if (typeof module !== 'undefined') {
  module.exports = { KG_xlsxToRows_: KG_xlsxToRows_, KG_parseClassRows_: KG_parseClassRows_, KG_readRecommend_: KG_readRecommend_,
    KG_buildModel_: KG_buildModel_, KG_tables_: KG_tables_, KG_course_: KG_course_ };
}
