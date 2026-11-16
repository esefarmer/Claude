/**
 * 校内ポータル（掲載台帳連動）
 *
 * 掲載台帳のスプレッドシートに紐づけて使う Apps Script。
 * 「台帳」シートの行を読み、教職員用・生徒用のリンク集として表示する。
 * 設置手順は README.md を参照。
 */

const CONFIG = {
  // 台帳のシート名
  LEDGER_SHEET: '台帳',
  // 教職員のメールアドレスを A 列に並べたシート名（1 行目は見出し）
  STAFF_SHEET: '教職員',
  // 教職員の Google グループがあればそのアドレスを入れる（例: 'staff@example.ac.jp'）。
  // 空のままなら STAFF_SHEET の一覧で判定する。
  STAFF_GROUP: '',
  // 台帳の変更がポータルに反映されるまでの最長時間（秒）
  CACHE_SECONDS: 300,
  TITLE: '校内ポータル',
};

const CACHE_KEY = 'portal-items-v1';

/** ウェブアプリの入口。閲覧者が教職員か生徒かを判定し、見せてよい項目だけを渡す。 */
function doGet() {
  const email = Session.getActiveUser().getEmail();
  const staff = isStaff_(email);
  const items = loadItems_();

  const data = {
    staff: staff,
    // 生徒には生徒向けの項目だけを送る。教職員用の項目はページのデータにも含めない。
    items: staff ? items : items.filter(isForStudent_),
  };

  const template = HtmlService.createTemplateFromFile('index');
  // 「<」を置き換えて、台帳の文字列が HTML として解釈されないようにする
  template.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  return template
    .evaluate()
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 生徒に見せてよい項目か。校外秘の項目は、掲載先の指定にかかわらず生徒には出さない。 */
function isForStudent_(item) {
  return (item.aud === '生徒用' || item.aud === '両方') && !item.secret;
}

/** 教職員かどうか。グループの指定があればグループで、なければ「教職員」シートで判定する。 */
function isStaff_(email) {
  if (!email) return false;
  const address = email.toLowerCase();
  if (CONFIG.STAFF_GROUP) {
    try {
      return GroupsApp.getGroupByEmail(CONFIG.STAFF_GROUP).hasUser(address);
    } catch (e) {
      console.error('グループを確認できません: ' + e);
      return false;
    }
  }
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.STAFF_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return false;
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .some(function (row) { return row[0].trim().toLowerCase() === address; });
}

/** 台帳から「掲載中」の項目を読む。結果は CACHE_SECONDS の間キャッシュする。 */
function loadItems_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const items = readLedger_();
  try {
    cache.put(CACHE_KEY, JSON.stringify(items), CONFIG.CACHE_SECONDS);
  } catch (e) {
    // 件数が多くキャッシュに入らない場合は、毎回台帳を読む
    console.warn('キャッシュに保存できません: ' + e);
  }
  return items;
}

/** 見出し名で列を探すので、台帳の列の並びを変えても動く。 */
function readLedger_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.LEDGER_SHEET);
  if (!sheet) throw new Error('シート「' + CONFIG.LEDGER_SHEET + '」が見つかりません。');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const head = values[0].map(function (v) { return v.trim(); });
  const col = function (name) {
    return head.findIndex(function (h) { return h.indexOf(name) === 0; });
  };
  const c = {
    name: col('名称'), url: col('URL'), desc: col('一言説明'), cat: col('カテゴリ'),
    aud: col('掲載先'), secret: col('校外秘'), status: col('掲載状況'), tag: col('種別'),
  };
  ['name', 'cat', 'aud', 'status'].forEach(function (key) {
    if (c[key] < 0) throw new Error('台帳の見出し行に必要な列がありません（名称・カテゴリ・掲載先・掲載状況）。');
  });
  const get = function (row, i) { return i < 0 ? '' : String(row[i]).trim(); };

  return values.slice(1)
    .map(function (row) {
      const url = get(row, c.url);
      return {
        name: get(row, c.name),
        url: /^https?:\/\//i.test(url) ? url : '',
        desc: get(row, c.desc),
        cat: get(row, c.cat) || 'その他',
        aud: get(row, c.aud),
        secret: get(row, c.secret) !== '',
        tag: get(row, c.tag),
        status: get(row, c.status),
      };
    })
    .filter(function (item) { return item.name && item.status === '掲載中'; })
    .map(function (item) { delete item.status; return item; });
}

/** スプレッドシートにメニューを追加する。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ポータル')
    .addItem('台帳の変更を今すぐ反映', 'clearCache')
    .addItem('掲載内容を確認', 'showSummary')
    .addToUi();
}

/** キャッシュを消して、次の表示から最新の台帳を使う。 */
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  SpreadsheetApp.getActive().toast('次にポータルを開いたときから、最新の台帳が表示されます。', 'ポータル');
}

/** 掲載中の件数と、生徒用に指定されているのに校外秘で出ない項目を知らせる。 */
function showSummary() {
  const items = readLedger_();
  const staffCount = items.filter(function (i) { return i.aud === '教職員用' || i.aud === '両方'; }).length;
  const studentCount = items.filter(isForStudent_).length;
  const hidden = items.filter(function (i) {
    return (i.aud === '生徒用' || i.aud === '両方') && i.secret;
  });
  const noUrl = items.filter(function (i) { return !i.url; });

  let message = '教職員用: ' + staffCount + '件\n生徒用: ' + studentCount + '件';
  if (hidden.length) {
    message += '\n\n校外秘のため生徒用には出ない項目:\n' +
      hidden.map(function (i) { return '・' + i.name; }).join('\n');
  }
  if (noUrl.length) {
    message += '\n\nURL が空か http(s) で始まらない項目（押しても開きません）:\n' +
      noUrl.map(function (i) { return '・' + i.name; }).join('\n');
  }
  SpreadsheetApp.getUi().alert('掲載中の項目', message, SpreadsheetApp.getUi().ButtonSet.OK);
}
