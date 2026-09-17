/**
 * おこづかいちょう — データ同期スクリプト(Google Apps Script)
 *
 * スプレッドシートに紐づけて「ウェブアプリ」としてデプロイして使います。
 * 設定手順は 同期セットアップ手順.md を参照してください。
 *
 * 【重要】下の SECRET を、自分で決めた合言葉に必ず書き換えてください。
 *         同じ文字列をアプリの保護者ページにも入力します。
 *
 * 【注意】書き換えるのは「スプレッドシートに貼り付けたあとの Apps Script 側」だけです。
 *         このファイルはウェブに公開されているため、実際の合言葉を書いて
 *         コミットしないでください(ここはプレースホルダのままにしておく)。
 */

var SECRET = 'kaeru-2026-himitsu';   // ←★Apps Script 側で自分の合言葉に書き換える(このファイルは変更しない)

var SHEET_TX   = '明細';
var SHEET_BANK = 'ぎんこう';
var SHEET_META = '設定';

var TX_HEADER   = ['id', '日付', '種類', 'カテゴリ', '金額', 'メモ'];
var BANK_HEADER = ['id', '日付', '種類', '金額', '後残高'];
var META_HEADER = ['キー', '値'];

// 文字列として保存する列(1=テキスト固定)。日付の自動変換と先頭ゼロの欠落を防ぐ。
var TX_TEXTCOLS   = [1, 1, 1, 1, 0, 1];
var BANK_TEXTCOLS = [1, 1, 1, 0, 0];
var META_TEXTCOLS = [1, 1];

/* ============================ エントリポイント ============================ */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return jsonOut_({ ok: false, reason: 'busy' });
  }
  try {
    var req = JSON.parse(e.postData.contents);
    if (String(req.secret || '') !== SECRET) {
      return jsonOut_({ ok: false, reason: 'unauthorized' });
    }
    if (req.action === 'ping') return jsonOut_(ping_());
    if (req.action === 'pull') return jsonOut_(pull_());
    if (req.action === 'push') return jsonOut_(push_(req));
    return jsonOut_({ ok: false, reason: 'unknown_action' });
  } catch (err) {
    return jsonOut_({ ok: false, reason: 'error', message: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

// ブラウザで URL を直接開いたときの動作確認用
function doGet() {
  return jsonOut_({ ok: true, info: 'okozukai sync endpoint' });
}

/* ============================== 各アクション ============================== */

function ping_() {
  var meta = readMeta_();
  return {
    ok: true,
    revision: num_(meta._revision, 0),
    hasData: num_(meta._revision, 0) > 0,
    sheetUrl: ss_().getUrl(),
    updatedAt: meta._updatedAt || null
  };
}

function pull_() {
  var meta = readMeta_();
  return { ok: true, revision: num_(meta._revision, 0), payload: readAll_(meta) };
}

function push_(req) {
  var meta = readMeta_();
  var cur = num_(meta._revision, 0);

  // 同じ opId の再送(応答が届かなかった場合の再試行)は、書き込まずに成功として返す
  if (req.opId && meta._lastOpId && String(req.opId) === String(meta._lastOpId)) {
    return { ok: true, revision: cur, duplicate: true };
  }
  // force は保護者が「この端末の内容で上書き」を選んだ場合のみ
  if (!req.force && num_(req.baseRevision, -1) !== cur) {
    return { ok: false, reason: 'conflict', revision: cur };
  }

  var next = cur + 1;
  applyPayload_(req.payload || {}, next, req.opId || '');
  return { ok: true, revision: next };
}

/* ============================== 読み込み ============================== */

function readAll_(meta) {
  meta = meta || readMeta_();
  return {
    wallet: { transactions: readTx_() },
    bank: {
      balance: num_(meta['bank.balance'], 0),
      annualRate: num_(meta['bank.annualRate'], 10),
      lastInterestYm: meta['bank.lastInterestYm'] || '',
      history: readBank_()
    },
    settings: {
      parentPin: meta['settings.parentPin'] || null,
      lastBackupAt: meta['settings.lastBackupAt'] || null
    }
  };
}

function readTx_() {
  var rows = readRows_(sheet_(SHEET_TX, TX_HEADER), TX_HEADER.length);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[1]) continue;
    out.push({
      id: r[0], date: r[1], type: r[2], category: r[3],
      amount: num_(r[4], 0), memo: r[5] || ''
    });
  }
  return out;
}

function readBank_() {
  var rows = readRows_(sheet_(SHEET_BANK, BANK_HEADER), BANK_HEADER.length);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[1]) continue;
    out.push({
      id: r[0], date: r[1], type: r[2],
      amount: num_(r[3], 0), balanceAfter: num_(r[4], 0)
    });
  }
  return out;
}

function readMeta_() {
  var rows = readRows_(sheet_(SHEET_META, META_HEADER), META_HEADER.length);
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0]) out[rows[i][0]] = rows[i][1];
  }
  return out;
}

// getDisplayValues を使い、日付が Date オブジェクトに変換されるのを避ける
function readRows_(sh, width) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, width).getDisplayValues();
}

/* ============================== 書き込み ============================== */

function applyPayload_(payload, revision, opId) {
  var tx    = (payload.wallet && payload.wallet.transactions) || [];
  var bank  = payload.bank || {};
  var hist  = bank.history || [];
  var st    = payload.settings || {};

  writeRows_(sheet_(SHEET_TX, TX_HEADER), TX_HEADER, tx.map(function (t) {
    return [s_(t.id), s_(t.date), s_(t.type), s_(t.category), num_(t.amount, 0), s_(t.memo)];
  }), TX_TEXTCOLS);

  writeRows_(sheet_(SHEET_BANK, BANK_HEADER), BANK_HEADER, hist.map(function (h) {
    return [s_(h.id), s_(h.date), s_(h.type), num_(h.amount, 0), num_(h.balanceAfter, 0)];
  }), BANK_TEXTCOLS);

  // 保護者が一目で分かるよう、計算済みの残高も書き出す(読み込み時は使わない)
  var walletBal = 0;
  for (var i = 0; i < tx.length; i++) {
    walletBal += (tx[i].type === 'in' ? 1 : -1) * num_(tx[i].amount, 0);
  }
  var bankBal = num_(bank.balance, 0);

  var meta = {};
  meta['おさいふ残高'] = walletBal;
  meta['ぎんこう残高'] = bankBal;
  meta['合計'] = walletBal + bankBal;
  meta['bank.balance'] = bankBal;
  meta['bank.annualRate'] = num_(bank.annualRate, 10);
  meta['bank.lastInterestYm'] = s_(bank.lastInterestYm);
  meta['settings.parentPin'] = s_(st.parentPin);
  meta['settings.lastBackupAt'] = s_(st.lastBackupAt);
  meta['_revision'] = revision;
  meta['_updatedAt'] = new Date().toISOString();
  meta['_lastOpId'] = s_(opId);

  var keys = ['おさいふ残高', 'ぎんこう残高', '合計', 'bank.balance', 'bank.annualRate',
              'bank.lastInterestYm', 'settings.parentPin', 'settings.lastBackupAt',
              '_revision', '_updatedAt', '_lastOpId'];
  var metaRows = keys.map(function (k) { return [k, String(meta[k])]; });
  writeRows_(sheet_(SHEET_META, META_HEADER), META_HEADER, metaRows, META_TEXTCOLS);
}

/**
 * 2行目以降を入れ替える。シート自体は削除しないため、
 * 他シートに作ったグラフや数式の参照は壊れない。
 */
function writeRows_(sh, header, rows, textCols) {
  var last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, header.length).clearContent();
  }
  if (!rows.length) return;

  var needed = rows.length + 1;
  if (sh.getMaxRows() < needed) {
    sh.insertRowsAfter(sh.getMaxRows(), needed - sh.getMaxRows());
  }
  // setValues の前に書式を確定させる(後からでは値の変換を防げない)
  for (var c = 0; c < textCols.length; c++) {
    if (textCols[c]) sh.getRange(2, c + 1, rows.length, 1).setNumberFormat('@');
  }
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

/* ================================ 補助 ================================ */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, header) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function num_(v, def) {
  if (v === '' || v === null || v === undefined) return def;
  var n = Number(v);
  return isFinite(n) ? n : def;
}

function s_(v) { return (v === null || v === undefined) ? '' : String(v); }

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
