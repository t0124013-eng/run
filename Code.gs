/**
 * 持久走ラボ — Google Apps Script（スプレッドシートに紐づけて使う）
 *
 * 使い方
 *  1. 新しいGoogleスプレッドシートを作る →「拡張機能」→「Apps Script」
 *  2. このファイルを Code.gs に、index.html を「ファイル＋ → HTML」で index という名前で貼り付け
 *  3. 関数 setup を一度実行（シートが自動で作られます）
 *  4. 「デプロイ」→「新しいデプロイ」→ 種類：ウェブアプリ
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：（学校ドメイン）内の全員
 *  5. 発行されたURLを生徒に配布（QRコードなど）
 */

const SH = { SETTINGS: '設定', SCORE: '得点表', ROSTER: '名簿', REC: '記録', HYP: '仮説' };

const HEAD = {
  '設定':  ['項目', '値', '説明'],
  '得点表': ['性別', '点数', '上限タイム秒', '表示（参考）'],
  '名簿':  ['email', '学年', '組', '番号', '氏名', '性別', '登録日時'],
  '記録':  ['id', '登録日時', '実施日', 'email', '氏名', '計測者', '距離m', 'タイム秒', 'ラップ秒',
            '睡眠h', '朝食', '体調', 'アップ分', '週練習分', 'きつさ', '作戦', '振り返り', '教員コメント', '入力方法'],
  '仮説':  ['id', '登録日時', 'email', '要因', '仮説', '理由', '検証方法', '結論', '考察', '更新日時']
};

const SETTING_DEFS = [
  ['アプリ名', 'title', '持久走ラボ', '画面上部に表示されます'],
  ['教員メール', 'teachers', '', '教員画面を使える人のメール（カンマ区切り）。スクリプトの所有者は自動で教員になります'],
  ['男子ラップ', 'lapM', '100,100,100,100,100,100,100,100,100,100,100,100,100,100,100', '男子1500mを周回ごとの距離で（合計1500）。1周100mなら100を15個'],
  ['女子ラップ', 'lapF', '100,100,100,100,100,100,100,100,100,100', '女子1000mを周回ごとの距離で（合計1000）。1周100mなら100を10個'],
  ['採点に使う記録', 'basis', 'best', 'best＝自己ベスト ／ latest＝最新の記録']
];

// 学校オリジナル得点表（サンプル。アプリの「採点」画面から変更できます）
const DEFAULT_SCORE = {
  M: [[10, 300], [9, 320], [8, 340], [7, 360], [6, 380], [5, 400], [4, 420], [3, 450], [2, 480], [1, 5999]],
  F: [[10, 215], [9, 230], [8, 245], [7, 260], [6, 275], [5, 290], [4, 305], [3, 325], [2, 345], [1, 5999]]
};

/* ───────── 初期設定 ───────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const h = HEAD[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#13315c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  // 文字列として保存したい列
  textCols_(SH.REC, ['id', '実施日', 'ラップ秒', 'email', '計測者']);
  textCols_(SH.ROSTER, ['email', '学年', '組']);
  textCols_(SH.HYP, ['id', 'email']);

  const st = ss.getSheetByName(SH.SETTINGS);
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, SETTING_DEFS.length, 3).setValues(SETTING_DEFS.map(d => [d[0], d[2], d[3]]));
    st.getRange('B:B').setNumberFormat('@');
  }
  const sc = ss.getSheetByName(SH.SCORE);
  if (sc.getLastRow() < 2) writeScoreRows_(toScoreRows_(DEFAULT_SCORE));
  const def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function textCols_(name, keys) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  keys.forEach(k => {
    const c = HEAD[name].indexOf(k) + 1;
    if (c > 0) sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  });
}

/* ───────── Web アプリ ───────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(getSettings_().title || '持久走ラボ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────── 共通 ───────── */

function me_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Googleアカウントのメールアドレスを取得できませんでした。学校のアカウントでログインしてから開いてください。');
  return email;
}
function isTeacher_(email) {
  const owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const list = String(getSettings_().teachers || '').split(/[,\s、]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
  return email === owner || list.indexOf(email) >= 0;
}
function requireTeacher_() {
  const e = me_();
  if (!isTeacher_(e)) throw new Error('教員のみ操作できます');
  return e;
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}
function rows_(name) {
  const h = HEAD[name];
  const v = sheet_(name).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    const o = { _row: i + 1 };
    h.forEach((k, j) => o[k] = v[i][j]);
    out.push(o);
  }
  return out;
}
function rowArr_(name, obj) { return HEAD[name].map(k => (obj[k] === undefined || obj[k] === null) ? '' : obj[k]); }
function append_(name, obj) { sheet_(name).appendRow(rowArr_(name, obj)); }
function writeRow_(name, row, obj) {
  const h = HEAD[name];
  sheet_(name).getRange(row, 1, 1, h.length).setValues([rowArr_(name, obj)]);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
const tz_ = () => Session.getScriptTimeZone() || 'Asia/Tokyo';
function d_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || ''); }
function dt_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm') : String(v || ''); }
function num_(v) { if (v === '' || v === null || v === undefined) return null; const n = Number(v); return isNaN(n) ? null : n; }
function r1_(n) { return Math.round(Number(n) * 10) / 10; }
function g_(v) { const s = String(v || ''); return (s.indexOf('女') >= 0 || s.toUpperCase() === 'F') ? 'F' : 'M'; }

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.SETTINGS);
  const o = {};
  SETTING_DEFS.forEach(d => o[d[1]] = d[2]);
  if (!sh) return o;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const d = SETTING_DEFS.find(x => x[0] === r[0]);
    if (d) o[d[1]] = String(r[1]);
  });
  return o;
}
function expandLap_(str) {
  const out = [];
  String(str || '').split(/[,、\s]+/).filter(Boolean).forEach(t => {
    const m = t.match(/^(\d+)\s*[x×*]\s*(\d+)$/i);
    if (m) for (let i = 0; i < Number(m[2]); i++) out.push(Number(m[1]));
    else if (Number(t) > 0) out.push(Number(t));
  });
  return out;
}
function settingsOut_(s, teacher) {
  const lap = str => expandLap_(str);
  const o = { title: s.title || '持久走ラボ', lapM: lap(s.lapM), lapF: lap(s.lapF), basis: s.basis === 'latest' ? 'latest' : 'best' };
  if (teacher) o.teachers = s.teachers || '';
  return o;
}
function rosterOut_() {
  return rows_(SH.ROSTER).map(r => ({
    email: String(r.email || '').trim().toLowerCase(), grade: String(r['学年'] || ''), cls: String(r['組'] || ''), no: num_(r['番号']),
    name: String(r['氏名'] || ''), gender: g_(r['性別']), _row: r._row
  })).filter(r => r.name);
}
function scoreOut_() {
  return rows_(SH.SCORE).map(r => ({ gender: g_(r['性別']), points: num_(r['点数']), limit: num_(r['上限タイム秒']) }))
    .filter(r => r.points !== null && r.limit !== null);
}
function recOut_(r) {
  const bf = r['朝食'];
  return {
    id: String(r.id), created: dt_(r['登録日時']), date: d_(r['実施日']), email: String(r.email).toLowerCase(),
    name: String(r['氏名'] || ''), recorder: String(r['計測者'] || ''), distance: num_(r['距離m']), time: num_(r['タイム秒']),
    laps: String(r['ラップ秒'] || '').split(/[,、\s]+/).filter(Boolean).map(Number).filter(n => !isNaN(n)),
    sleep: num_(r['睡眠h']), breakfast: bf === '' ? null : (bf === 'あり' || bf === true || bf === 'TRUE'),
    condition: num_(r['体調']), warmup: num_(r['アップ分']), weekly: num_(r['週練習分']), rpe: num_(r['きつさ']),
    strategy: String(r['作戦'] || ''), reflection: String(r['振り返り'] || ''), comment: String(r['教員コメント'] || ''),
    method: String(r['入力方法'] || '')
  };
}
function hypOut_(h) {
  return {
    id: String(h.id), created: dt_(h['登録日時']), email: String(h.email).toLowerCase(), factor: String(h['要因'] || ''),
    hypothesis: String(h['仮説'] || ''), reason: String(h['理由'] || ''), method: String(h['検証方法'] || ''),
    conclusion: String(h['結論'] || ''), discussion: String(h['考察'] || ''), updated: dt_(h['更新日時'])
  };
}
function newId_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

/* ───────── 読み込み ───────── */

function getBootstrap() {
  const email = me_();
  const teacher = isTeacher_(email);
  const roster = rosterOut_();
  const profile = roster.find(r => r.email === email) || null;
  const linked = roster.filter(r => r.email);
  const recs = rows_(SH.REC).map(recOut_);
  const hyps = rows_(SH.HYP).map(hypOut_);
  const res = {
    email: email, teacher: teacher, profile: profile,
    settings: settingsOut_(getSettings_(), teacher), scoreTable: teacher ? scoreOut_() : [],  // 得点は生徒に送らない
    myRecords: recs.filter(r => r.email === email),
    myHypotheses: hyps.filter(h => h.email === email),
    // 計測の相手として選べる生徒（生徒は同じ学年のみ、教員は全員）
    runners: (teacher ? linked : (profile ? linked.filter(r => r.grade === profile.grade) : []))
      .map(r => ({ email: r.email, grade: r.grade, cls: r.cls, no: r.no, name: r.name, gender: r.gender }))
  };
  if (teacher) res.all = { roster: linked.map(r => ({ email: r.email, grade: r.grade, cls: r.cls, no: r.no, name: r.name, gender: r.gender })), records: recs, hypotheses: hyps };
  return res;
}

/* ───────── 生徒の操作 ───────── */

const GRADES = ['中1', '中2', '中3', '高1', '高2', '高3'];

/** 初回登録：生徒が学年・組・番号・氏名・性別を自分で入力する */
function registerProfile(p) {
  const email = me_();
  const grade = String(p.grade || '').trim();
  const cls = String(p.cls || '').trim().replace(/組$/, '');
  const no = Number(p.no);
  const name = String(p.name || '').trim().replace(/\s+/g, '　');
  if (GRADES.indexOf(grade) < 0) throw new Error('学年を選んでください');
  if (!cls) throw new Error('組を入力してください');
  if (!(no >= 1 && no <= 99)) throw new Error('番号を正しく入力してください');
  if (!name) throw new Error('氏名を入力してください');
  return withLock_(() => {
    const roster = rosterOut_();
    if (roster.some(r => r.email === email)) return getBootstrap();
    const dup = roster.find(r => r.email && r.grade === grade && r.cls === cls && r.no === no);
    if (dup) throw new Error(grade + ' ' + cls + '組' + no + '番はすでに別のアカウントで登録されています。先生に伝えてください');
    const obj = { email: email, '学年': grade, '組': cls, '番号': no, '氏名': name, '性別': p.gender === 'F' ? '女' : '男', '登録日時': new Date() };
    // 先生が名簿を用意している場合は、同じ学年・組・番号の空き行に書き込む
    const pre = roster.find(r => !r.email && r.grade === grade && r.cls === cls && r.no === no);
    if (pre) writeRow_(SH.ROSTER, pre._row, obj); else append_(SH.ROSTER, obj);
    return getBootstrap();
  });
}

function saveRecord(p) {
  const me = me_();
  const teacher = isTeacher_(me);
  const roster = rosterOut_();
  const runnerEmail = String(p.runner || me).toLowerCase();
  const runner = roster.find(r => r.email && r.email === runnerEmail);
  if (!runner) throw new Error('走った人が名簿に見つかりません（最初に名簿登録が必要です）');
  if (!teacher && runnerEmail !== me) {
    const mine = roster.find(r => r.email && r.email === me);
    if (!mine || mine.grade !== runner.grade) throw new Error('同じ学年の生徒の記録だけ計測できます');
  }
  const t = Number(p.time);
  if (!(t > 60 && t < 3600)) throw new Error('タイムが正しくありません');
  const laps = (p.laps || []).map(Number).filter(n => n > 0).map(r1_);
  const rec = {
    id: newId_(), '登録日時': new Date(), '実施日': p.date || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
    email: runnerEmail, '氏名': runner.name, '計測者': me, '距離m': runner.gender === 'F' ? 1000 : 1500,
    'タイム秒': r1_(t), 'ラップ秒': laps.join(','), '入力方法': p.method === '手入力' ? '手入力' : 'ペア計測'
  };
  return withLock_(() => { append_(SH.REC, rec); return recOut_(rec); });
}

const REC_FIELDS = { sleep: '睡眠h', breakfast: '朝食', condition: '体調', warmup: 'アップ分', weekly: '週練習分', rpe: 'きつさ', strategy: '作戦', reflection: '振り返り' };

function updateRecord(id, f) {
  const me = me_();
  const teacher = isTeacher_(me);
  return withLock_(() => {
    const r = rows_(SH.REC).find(x => String(x.id) === String(id));
    if (!r) throw new Error('記録が見つかりません');
    if (String(r.email).toLowerCase() !== me && !teacher) throw new Error('自分の記録だけ編集できます');
    Object.keys(REC_FIELDS).forEach(k => {
      if (!(k in f)) return;
      let v = f[k];
      if (k === 'breakfast') v = v === null || v === '' ? '' : (v === true || v === 'true' ? 'あり' : 'なし');
      r[REC_FIELDS[k]] = v === null ? '' : v;
    });
    if (teacher) {
      if ('time' in f && Number(f.time) > 0) r['タイム秒'] = r1_(f.time);
      if ('date' in f && f.date) r['実施日'] = f.date;
    }
    writeRow_(SH.REC, r._row, r);
    return recOut_(r);
  });
}

function deleteRecord(id) {
  const me = me_();
  const teacher = isTeacher_(me);
  return withLock_(() => {
    const r = rows_(SH.REC).find(x => String(x.id) === String(id));
    if (!r) return true;
    const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    const okStudent = (String(r['計測者']).toLowerCase() === me || String(r.email).toLowerCase() === me) && d_(r['登録日時']).slice(0, 10) === today;
    if (!teacher && !okStudent) throw new Error('記録を消せるのは、計測した当日だけです。先生に相談してください');
    sheet_(SH.REC).deleteRow(r._row);
    return true;
  });
}

function saveHypothesis(h) {
  const me = me_();
  return withLock_(() => {
    const obj = {
      '要因': h.factor || '', '仮説': h.hypothesis || '', '理由': h.reason || '', '検証方法': h.method || '',
      '結論': h.conclusion || '', '考察': h.discussion || '', '更新日時': new Date()
    };
    if (h.id) {
      const ex = rows_(SH.HYP).find(x => String(x.id) === String(h.id));
      if (!ex || String(ex.email).toLowerCase() !== me) throw new Error('仮説が見つかりません');
      Object.assign(ex, obj);
      writeRow_(SH.HYP, ex._row, ex);
      return hypOut_(ex);
    }
    const n = Object.assign({ id: newId_(), '登録日時': new Date(), email: me }, obj);
    append_(SH.HYP, n);
    return hypOut_(n);
  });
}

function deleteHypothesis(id) {
  const me = me_();
  return withLock_(() => {
    const ex = rows_(SH.HYP).find(x => String(x.id) === String(id));
    if (ex && String(ex.email).toLowerCase() === me) sheet_(SH.HYP).deleteRow(ex._row);
    return true;
  });
}

/* ───────── 教員の操作 ───────── */


function saveSettings(s) {
  requireTeacher_();
  const sum = str => expandLap_(str).reduce((a, b) => a + b, 0);
  if (sum(s.lapM) !== 1500) throw new Error('男子ラップの合計が1500mになっていません');
  if (sum(s.lapF) !== 1000) throw new Error('女子ラップの合計が1000mになっていません');
  return withLock_(() => {
    const sh = sheet_(SH.SETTINGS);
    const vals = SETTING_DEFS.map(d => [d[0], String(s[d[1]] !== undefined ? s[d[1]] : d[2]), d[3]]);
    sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).clearContent();
    sh.getRange(2, 2, vals.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, vals.length, 3).setValues(vals);
    return settingsOut_(getSettings_(), true);
  });
}

function saveScoreTable(rows, basis) {
  requireTeacher_();
  const clean = (rows || []).map(r => ({ gender: r.gender === 'F' ? 'F' : 'M', points: Number(r.points), limit: Number(r.limit) }))
    .filter(r => !isNaN(r.points) && r.limit > 0);
  if (!clean.length) throw new Error('得点表が空です');
  return withLock_(() => {
    writeScoreRows_(clean);
    if (basis) {
      const s = getSettings_(); s.basis = basis === 'latest' ? 'latest' : 'best';
      const sh = sheet_(SH.SETTINGS);
      const vals = SETTING_DEFS.map(d => [d[0], String(s[d[1]]), d[3]]);
      sh.getRange(2, 1, vals.length, 3).setValues(vals);
    }
    return { scoreTable: scoreOut_(), settings: settingsOut_(getSettings_(), true) };
  });
}

function toScoreRows_(def) {
  const out = [];
  ['M', 'F'].forEach(g => def[g].forEach(p => out.push({ gender: g, points: p[0], limit: p[1] })));
  return out;
}
function writeScoreRows_(rows) {
  const sh = sheet_(SH.SCORE);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD[SH.SCORE].length).clearContent();
  const fmt = s => s >= 5999 ? 'それより遅い' : Math.floor(s / 60) + '分' + ('0' + Math.round(s % 60)).slice(-2) + '秒以内';
  const vals = rows.map(r => [r.gender === 'F' ? '女' : '男', r.points, r.limit, fmt(r.limit)]);
  sh.getRange(2, 1, vals.length, 4).setValues(vals);
}
