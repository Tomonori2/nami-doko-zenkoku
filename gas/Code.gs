// 全国波情報 写真・動画の受け口
// アプリから送られた「どのポイント・何時・サイズ・ひとこと・写真（または動画）」を、このスプレッドシートに1行ずつ記録する。
// 届いたものはすべて「保留」。管理人が「公開」の列を「公開」に変えた行だけがアプリに出る（写真・動画もひとことも）。
// スマホのスプレッドシートアプリだけで確認できるように、シートの中に小さい画像と「開く」リンクを出し、
// 「公開」の列は 保留／公開／非公開 から選ぶだけにする。届いたら持ち主にメールで知らせる（30分に1通まで）。
// 動画はアプリの中で15秒まで・小さく作り直してから届く（位置情報と音は消えている）。最初の1コマを写真として一緒に受ける。
// 送った人の名前は受け取らない。合言葉は無し（誰でも送れる）。そのかわり1日の件数に上限をつける。

const SHEET = '写真';
const SPOTS = ['oarai', 'otake', 'katakai', 'shidashita', 'taito', 'kamogawa', 'chikura', 'kugenuma', 'tsujido', 'shichiri',
  'omain', 'melon', 'shizu', 'kata', 'long', 'loco', 'zen', 'shin', 'kou', 'isonoura', 'komatsu', 'ikumi', 'kizaki', 'okura'];
// 動画IDは後から足した列なので一番右（前からある行の並びを変えないため）
const HEAD = ['受けた日時', '公開', '写真', '開く', '日', '時刻', 'ポイント', 'サイズ', '乗りやすさ', 'ひとこと', '写真ID', '予想の印', '予想のサイズm', '版', '動画ID'];
const COL = {};
HEAD.forEach(function (h, i) { COL[h] = i; });
const CHOICES = ['保留', '公開', '非公開'];
const MAX_PER_DAY = 50;        // 1日に受ける件数の上限（いたずら対策）
const MAX_VID_PER_DAY = 10;    // そのうち動画の上限（ドライブの容量を守る）
const MAX_IMG_CHARS = 1500000; // 写真の大きさの上限（約1MB）
const MAX_VID_CHARS = 8000000; // 動画の大きさの上限（約6MB）
const VID_TYPES = { 'video/mp4': 'mp4', 'video/webm': 'webm' };
const NOTIFY = true;           // 届いたらメールで知らせる。いらなければ false
const NOTIFY_GAP_MIN = 30;     // お知らせメールの間かく（分）

// 最初に1回だけ、エディタの「実行」で動かす。シート・写真フォルダ・選ぶだけの欄を作り、必要な許可をまとめて取る
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); }
  head_(sh);
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL['写真'] + 1, 160);
  sh.getRange(2, COL['公開'] + 1, sh.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CHOICES, true).setAllowInvalid(false).build());
  folder_();
}

// 見出しが古い（列が足りない）ときは最新にそろえる
function head_(sh) {
  if (sh.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1]) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('全国波情報 写真');
  props.setProperty('FOLDER', f.getId());
  return f;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// 文字は長さを切り、先頭の = + - @ を外す（スプレッドシートの式として動かないように）
function text_(v, n) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').replace(/^[=+\-@\s]+/, '').slice(0, n); }
function int_(v, max) { const n = Math.floor(Number(v)); return n >= 0 && n <= max ? n : ''; }
const B64 = /^[A-Za-z0-9+/=]+$/;

// ドライブに保存して、リンクを知っている人だけ見られるようにする
function save_(b64, mime, name) {
  const file = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

// 持ち主にお知らせ。前のお知らせから30分たっていなければ、数だけためて次にまとめる
function notify_(ss) {
  if (!NOTIFY) return;
  const props = PropertiesService.getScriptProperties();
  const waiting = Number(props.getProperty('WAITING') || 0) + 1, last = Number(props.getProperty('LAST_MAIL') || 0);
  if (Date.now() - last < NOTIFY_GAP_MIN * 60000) { props.setProperty('WAITING', String(waiting)); return; }
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【全国波情報】写真・動画が' + waiting + '件届きました（確認待ち）',
      '確認待ちの写真・動画があります。スプレッドシートを開いて、「公開」の列を 公開 か 非公開 に変えてください。\n\n' + ss.getUrl() +
      '\n\n顔がはっきり分かる人・車のナンバー・家や表札が写っているものは「非公開」にしてください。動画は最後まで見てから決めてください。');
    props.setProperty('LAST_MAIL', String(Date.now()));
    props.setProperty('WAITING', '0');
  } catch (e) { props.setProperty('WAITING', String(waiting)); }   // メールが送れなくても受付は止めない
}

// アプリから写真・動画を受け取る（写真なしは受けない。動画のときも最初の1コマを写真として受ける）
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: 'format' }); }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SPOTS.indexOf(d.spot) < 0 || !/^\d{4}-\d{2}-\d{2}$/.test(d.date) || !/^\d{2}:\d{2}$/.test(d.time)) return json_({ ok: false, err: 'format' });
  const img = typeof d.img === 'string' ? d.img : '', vid = typeof d.vid === 'string' ? d.vid : '';
  if (!img || img.length > MAX_IMG_CHARS || !B64.test(img)) return json_({ ok: false, err: 'image' });
  if (vid && (vid.length > MAX_VID_CHARS || !B64.test(vid) || !VID_TYPES[d.mime])) return json_({ ok: false, err: 'video' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, err: 'busy' });  // 順番が回ってこなければ受けない（上限のすり抜け防止）
  try {
    const sh = ss.getSheetByName(SHEET);
    head_(sh);
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const todays = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues()
      .filter(function (r) { return r[0] && Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'yyyy-MM-dd') === today; });
    if (todays.length >= MAX_PER_DAY) return json_({ ok: false, err: 'limit' });
    if (vid && todays.filter(function (r) { return r[COL['動画ID']]; }).length >= MAX_VID_PER_DAY) return json_({ ok: false, err: 'vlimit' });
    const name = d.date + '_' + d.time.replace(':', '') + '_' + d.spot;
    const id = save_(img, 'image/jpeg', name + '.jpg'), vidId = vid ? save_(vid, d.mime, name + '.' + VID_TYPES[d.mime]) : '';
    const pred = d.pred || {};
    sh.appendRow([new Date(), '保留',
      '=IMAGE("https://drive.google.com/thumbnail?id=' + id + '&sz=w400")',
      '=HYPERLINK("https://drive.google.com/file/d/' + (vidId || id) + '/view","' + (vidId ? '動画を開く' : '開く') + '")',
      "'" + d.date, "'" + d.time, d.spot, int_(d.size, 9), int_(d.cond, 3),
      text_(d.memo, 60), id, int_(pred.mark + 1, 4) === '' ? '' : pred.mark, Number(pred.face) > 0 ? Math.round(pred.face * 100) / 100 : '', text_(d.ver, 8), vidId]);
    sh.setRowHeight(sh.getLastRow(), 110);   // 小さい画像が見える高さ
    notify_(ss);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

// アプリに、最近14日分の「公開」の行だけを返す（「保留」「非公開」は写真・動画もひとことも出ない）
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh || sh.getLastRow() < 2) return json_({ ok: true, list: [] });
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getDisplayValues()
    .map(function (r) { r[COL['日']] = r[COL['日']].replace(/^'/, ''); r[COL['時刻']] = r[COL['時刻']].replace(/^'/, ''); return r; });
  const since = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), 'Asia/Tokyo', 'yyyy-MM-dd');
  const okId = function (s) { return /^[A-Za-z0-9_\-]+$/.test(s); };
  const list = rows.filter(function (r) {
    return r[COL['公開']] === '公開' && r[COL['日']] >= since && SPOTS.indexOf(r[COL['ポイント']]) >= 0 && okId(r[COL['写真ID']]);
  }).slice(-80).map(function (r) {
    return { date: r[COL['日']], time: r[COL['時刻']], spot: r[COL['ポイント']], size: r[COL['サイズ']], cond: r[COL['乗りやすさ']], memo: r[COL['ひとこと']],
      img: r[COL['写真ID']], vid: okId(r[COL['動画ID']]) ? r[COL['動画ID']] : '' };
  });
  return json_({ ok: true, list: list });
}
