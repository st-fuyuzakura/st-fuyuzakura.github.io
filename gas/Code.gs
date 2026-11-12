/**
 * Studio冬桜 依頼フォーム受付（Google Apps Script）
 *
 * できること
 *  - サイトのフォーム送信を受け取り、スプレッドシートに1行追加（受付番号・ステータス）
 *  - 依頼者へ受付メール（固定料金プランは金額と納期つき／相談・オプションありは「見積もりします」）
 *  - 本人へ通知メール
 *  - ステータスを「入金済み」にすると着手連絡、「納品」にすると納品メールを自動送信
 *  - 素材リンクが空のまま3日たつと、依頼者へ1回だけリマインド
 *
 * 設置手順は gas/README.md を参照。最初に setup() を1回だけ実行する。
 */

const CONFIG = {
  STUDIO_NAME: 'Studio冬桜',
  NOTIFY_TO: 'st.fuyuzakura@gmail.com',      // 本人への通知先
  SITE_URL: 'https://st-fuyuzakura.github.io/',
  SHEET_NAME: '依頼',
  ID_PREFIX: 'FZ',
  REMIND_DAYS: 3,
  PAYMENT_TEXT: 'お支払い方法は、別途ご連絡いたします。',   // 仮：支払い方法が決まったら差し替え
};

/* 固定料金プラン（サイトの選択肢の文字と一致させる） */
const PLANS = {
  'ショート ¥6,500':      { price: '¥6,500',  days: '3〜5日' },
  'ワンコーラス ¥10,500': { price: '¥10,500', days: '5〜7日' },
  'スタンダード ¥15,000': { price: '¥15,000', days: '7〜10日' },
  'プレミアム ¥20,000':   { price: '¥20,000', days: '7〜10日' },
};

/* 希望欄にこの言葉があれば「オプションあり」とみなし、手動見積もりにする */
const OPTION_WORDS = /特急|急ぎ|追加|ハモリ|紹介|エンコード|修正|セット|トラック|[4-9４-９]本/;

const STATUS = ['受付', '見積もり済み', '入金済み', '納品', '完了', 'キャンセル'];

const COLS = [
  '受付番号', '受付日時', 'ステータス', 'お名前', 'メール', '曲名', 'プラン', 'XのID',
  '素材リンク', 'イメージや希望', '見積もり', '金額', '納期目安', '見積もり日',
  'リマインド送信日', '納品リンク', '着手連絡日', '納品連絡日', 'メモ',
];
const C = Object.fromEntries(COLS.map((name, i) => [name, i + 1])); // 列番号（1始まり）

/* ───────── 初回設定：シート・プルダウン・トリガーを作る ───────── */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#eeeeee');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build();
  sh.getRange(2, C['ステータス'], sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.setColumnWidth(C['イメージや希望'], 260);
  sh.setColumnWidth(C['メモ'], 200);

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyRemind').timeBased().everyDays(1).atHour(10).create();
}

/* ───────── フォーム受信 ───────── */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return ok_();                       // ボット対策（見えない入力欄に何か入っていたら無視）
    const v = k => String(d[k] || '').trim().slice(0, 2000);
    const name = v('name'), mail = v('contact'), song = v('song'), plan = v('plan');
    if (!name || !song || !plan || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) return ok_();

    const sh = sheet_();
    const now = new Date();
    const id = nextId_(sh, now);
    const fixed = PLANS[plan] && !OPTION_WORDS.test(v('note'));
    const p = PLANS[plan] || {};

    const row = new Array(COLS.length).fill('');
    row[C['受付番号'] - 1] = id;
    row[C['受付日時'] - 1] = now;
    row[C['ステータス'] - 1] = '受付';
    row[C['お名前'] - 1] = name;
    row[C['メール'] - 1] = mail;
    row[C['曲名'] - 1] = song;
    row[C['プラン'] - 1] = plan;
    row[C['XのID'] - 1] = v('x');
    row[C['素材リンク'] - 1] = v('files');
    row[C['イメージや希望'] - 1] = v('note');
    row[C['見積もり'] - 1] = fixed ? '自動（固定料金）' : '要・手動見積もり';
    row[C['金額'] - 1] = fixed ? p.price : '';
    row[C['納期目安'] - 1] = fixed ? p.days : '';
    sh.appendRow(row.map(x => (typeof x === 'string' && /^[=+\-@]/.test(x)) ? "'" + x : x)); // 数式として解釈されないように

    const r = { id, name, mail, song, plan, files: v('files'), note: v('note'), x: v('x'), price: p.price, days: p.days };
    send_(mail, `【${CONFIG.STUDIO_NAME}】ご依頼を受け付けました（${id}）`, fixed ? mailFixed_(r) : mailManual_(r));
    MailApp.sendEmail({
      to: CONFIG.NOTIFY_TO,
      subject: `［新規依頼］${id} ${name} 様／${song}／${fixed ? '自動見積もり済み' : '要・見積もり'}`,
      body: [
        `受付番号：${id}`, `お名前：${name}`, `メール：${mail}`, `XのID：${r.x || '—'}`,
        `曲名：${song}`, `プラン：${plan}`, `素材リンク：${r.files || '未'}`, `イメージや希望：${r.note || '—'}`, '',
        fixed ? `受付メールに金額（${p.price}）と納期（素材受け取りから${p.days}）を記載して送信済みです。`
              : '見積もりメールを手動で送り、ステータスを「見積もり済み」にしてください。',
        '', SpreadsheetApp.getActiveSpreadsheet().getUrl(),
      ].join('\n'),
    });
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

/* 動作確認用：ブラウザでウェブアプリURLを開くと表示される */
function doGet() {
  return ContentService.createTextOutput('OK');
}

/* ───────── ステータス変更 → 自動メール ───────── */
function onStatusEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.SHEET_NAME || e.range.getColumn() !== C['ステータス'] || e.range.getRow() < 2) return;
  if (e.range.getNumRows() > 1) return;
  const rowNo = e.range.getRow();
  const r = readRow_(sh, rowNo);
  const status = String(e.value || '');
  const now = new Date();

  if (status === '見積もり済み' && !r['見積もり日']) {
    sh.getRange(rowNo, C['見積もり日']).setValue(now);
  }
  if (status === '入金済み' && !r['着手連絡日']) {
    send_(r['メール'], `【${CONFIG.STUDIO_NAME}】ご入金を確認しました（${r['受付番号']}）`, mailStart_(r));
    sh.getRange(rowNo, C['着手連絡日']).setValue(now);
  }
  if (status === '納品' && !r['納品連絡日']) {
    if (!r['納品リンク']) {
      sh.getRange(rowNo, C['ステータス']).setValue(e.oldValue || '入金済み');
      SpreadsheetApp.getActiveSpreadsheet().toast('先に「納品リンク」欄にURLを入れてください。メールは送っていません。', '納品メール', 8);
      return;
    }
    send_(r['メール'], `【${CONFIG.STUDIO_NAME}】納品のご案内（${r['受付番号']}）`, mailDeliver_(r));
    sh.getRange(rowNo, C['納品連絡日']).setValue(now);
  }
  if (status === '入金済み' || status === '納品') {
    SpreadsheetApp.getActiveSpreadsheet().toast(`${r['お名前']} 様にメールを送りました。`, status, 5);
  }
}

/* ───────── 毎日10時：素材リンクが無いままの依頼にリマインド（1回だけ） ───────── */
function dailyRemind() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const values = sh.getRange(2, 1, last - 1, COLS.length).getValues();
  const now = new Date();
  values.forEach((row, i) => {
    const r = Object.fromEntries(COLS.map((n, j) => [n, row[j]]));
    if (r['素材リンク'] || r['リマインド送信日']) return;
    let base = null;
    if (r['ステータス'] === '受付' && r['見積もり'] === '自動（固定料金）') base = r['受付日時'];
    if (r['ステータス'] === '見積もり済み') base = r['見積もり日'];
    if (!(base instanceof Date)) return;
    if (now - base < CONFIG.REMIND_DAYS * 86400000) return;
    send_(r['メール'], `【${CONFIG.STUDIO_NAME}】素材のご準備はいかがでしょうか（${r['受付番号']}）`, mailRemind_(r));
    sh.getRange(i + 2, C['リマインド送信日']).setValue(now);
  });
}

/* ───────── メール本文 ───────── */
function head_(name) { return `${name} 様\n\nこの度は ${CONFIG.STUDIO_NAME} にご依頼いただき、ありがとうございます。\n`; }
function foot_() {
  return `\nご不明な点は、このメールにそのまま返信してください。\n\n――――――――――\n${CONFIG.STUDIO_NAME}\n${CONFIG.SITE_URL}\n――――――――――`;
}
function summary_(r) {
  return [
    '■ ご依頼内容',
    `受付番号：${r.id || r['受付番号']}`,
    `曲名：${r.song || r['曲名']}`,
    `プラン：${r.plan || r['プラン']}`,
  ].join('\n');
}
function filesLine_(files) {
  return files
    ? `素材リンク：${files}\n（念のため、ダウンロードできるか確認いたします）`
    : '素材は、準備ができ次第このメールに返信する形でリンクをお送りください。\n（ギガファイル便・Googleドライブなど、どのサービスでも大丈夫です）';
}
function mailFixed_(r) {
  return [
    head_(r.name), '以下の内容でご依頼を受け付けました。', '', summary_(r),
    `金額：${r.price}（税込・ボーカル3トラックまで。補正・ノイズ処理・ハモリ作成・マスタリング込み）`,
    `納期の目安：素材受け取りから ${r.days}`, '',
    '■ 素材について', filesLine_(r.files), '',
    '■ お支払いについて', CONFIG.PAYMENT_TEXT, 'ご入金を確認でき次第、制作に取りかかり、着手のご連絡をお送りします。',
    foot_(),
  ].join('\n');
}
function mailManual_(r) {
  return [
    head_(r.name), '以下の内容でご依頼を受け付けました。', '', summary_(r),
    r.note ? `イメージや希望：${r.note}` : '', '',
    '内容を確認し、24時間以内に金額と納期をご連絡いたします。',
    'この時点では料金は発生していませんので、ご安心ください。', '',
    '■ 素材について', filesLine_(r.files),
    foot_(),
  ].join('\n');
}
function mailStart_(r) {
  return [
    `${r['お名前']} 様\n`, 'ご入金を確認いたしました。ありがとうございます。', '本日より制作に取りかかります。', '',
    `受付番号：${r['受付番号']}`, `曲名：${r['曲名']}`,
    r['納期目安'] ? `納期の目安：素材受け取りから ${r['納期目安']}` : '',
    r['素材リンク'] ? '' : '\nまだ素材をお預かりしていないため、準備ができ次第このメールに返信でリンクをお送りください。',
    '\n仕上がりましたら、こちらのメールアドレスへ納品のご案内をお送りします。',
    foot_(),
  ].join('\n');
}
function mailDeliver_(r) {
  return [
    `${r['お名前']} 様\n`, 'お待たせいたしました。MIXが仕上がりましたので納品いたします。', '',
    `受付番号：${r['受付番号']}`, `曲名：${r['曲名']}`, '',
    '■ 納品データ', String(r['納品リンク']), '（リンクの保存期限がある場合があります。お早めにダウンロードしてください）', '',
    '気になるところがあれば、このメールに返信でお知らせください。修正の範囲はプランの内容に沿って対応いたします。',
    '\n素敵な作品になりますように。ご依頼いただき、ありがとうございました。',
    foot_(),
  ].join('\n');
}
function mailRemind_(r) {
  return [
    `${r['お名前']} 様\n`, `先日は ${CONFIG.STUDIO_NAME} にご依頼いただき、ありがとうございます。`,
    `「${r['曲名']}」（受付番号：${r['受付番号']}）の素材のご準備はいかがでしょうか。`, '',
    '準備ができましたら、このメールに返信する形で素材のリンクをお送りください。',
    'ご予定が変わった場合や、ご不明な点がある場合も、お気軽にご返信ください。',
    '（このご案内は1回だけお送りしています）',
    foot_(),
  ].join('\n');
}

/* ───────── 共通 ───────── */
function send_(to, subject, body) {
  MailApp.sendEmail({ to, subject, body: body.replace(/\n{3,}/g, '\n\n'), name: CONFIG.STUDIO_NAME, replyTo: CONFIG.NOTIFY_TO });
}
function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('先に setup() を実行してください');
  return sh;
}
function readRow_(sh, rowNo) {
  const row = sh.getRange(rowNo, 1, 1, COLS.length).getValues()[0];
  return Object.fromEntries(COLS.map((n, j) => [n, row[j]]));
}
function nextId_(sh, now) {
  const day = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  const prefix = `${CONFIG.ID_PREFIX}-${day}-`;
  const last = sh.getLastRow();
  const ids = last < 2 ? [] : sh.getRange(2, C['受付番号'], last - 1, 1).getValues().flat();
  const n = ids.filter(x => String(x).startsWith(prefix)).length + 1;
  return prefix + String(n).padStart(2, '0');
}
function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
