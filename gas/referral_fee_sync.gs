/**
 * 紹介料管理シート 自動生成スクリプト（Google Apps Script）
 *
 * 「従業員データ」シートを元に「紹介料管理」シートを丸ごと作り直します。
 *
 * ■ 導入手順
 *   1. スプレッドシートで「拡張機能 → Apps Script」を開く
 *   2. このファイルの内容を貼り付けて保存
 *   3. スプレッドシートを再読み込みすると、メニューに「紹介料管理」が出る
 *   4. 「紹介料管理 → 従業員データから更新」を実行（初回は権限の許可が必要）
 *
 * ■ 自動更新
 *   「従業員データ」シートを編集すると onEdit で自動的に再生成されます。
 *   不要なら onEdit 関数を削除してください。
 *
 * ■ 前提
 *   - 両シートとも 1 行目が見出し、2 行目がフィルタ行、3 行目からデータ
 *   - 列の位置は見出し名で探すので、列を移動しても動きます
 *   - 「従業員データ」の「紹介者」列は 氏名（例: 中村 尚史）でも 社員番号（例: 0001）でも可
 *
 * ■ 注意
 *   「紹介料管理」の 3 行目以降（対象列）は毎回上書きされます。手入力はしないでください。
 */

const CONFIG = {
  SOURCE_SHEET: '従業員データ',
  TARGET_SHEET: '紹介料管理',
  HEADER_ROW: 1,
  DATA_START_ROW: 3,

  // 「従業員データ」の見出し名
  SRC: {
    id: '社員番号',
    lastName: '氏名_姓',
    firstName: '氏名_名',
    hireDate: '入社年月日',
    referrer: '紹介者',
  },

  // 「紹介料管理」の見出し名（この順で出力）
  DST: {
    id: '社員番号',
    name: '氏名',
    hireDate: '入社年月日',
    referrerId: '紹介者社員番号',
    referrerName: '紹介者氏名',
    rank: '順位',
    fee: '1回額(税込)',
    months: ['支給月1', '支給月2', '支給月3', '支給月4'],
  },

  // 支給ルール
  FIRST_PAYMENT_OFFSET_MONTHS: 2, // 入社月の何か月後に 1 回目を支給するか
  PAYMENT_INTERVAL_MONTHS: 3,     // 2 回目以降の間隔（か月）
  FEE_BY_RANK: { 1: 30000 },      // 順位ごとの 1 回額。ここに無い順位は FEE_DEFAULT
  FEE_DEFAULT: 10000,

  // 制度対象外の社員番号（順位は数えるが、金額・支給月を空にする）
  EXCLUDED_IDS: [],

  // 紹介者がいない社員も一覧に載せるか
  INCLUDE_WITHOUT_REFERRER: true,

  // 当月以前の支給月セルをグレーにする（false なら背景は触らない）
  SHADE_PAST_MONTHS: true,
  PAST_MONTH_COLOR: '#d9d9d9',

  // 表示形式
  ID_DIGITS: 4,               // 社員番号が数値で入っている場合のゼロ埋め桁数
  NAME_SEPARATOR: ' ',        // 姓と名の間の区切り（全角にするなら '　'）
  DATE_FORMAT: 'yyyy/mm/dd',
  MONTH_FORMAT: 'yyyy/mm',
  FEE_FORMAT: '#,##0',
};

// ---------------------------------------------------------------- トリガー

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('紹介料管理')
    .addItem('従業員データから更新', 'syncReferralFees')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== CONFIG.SOURCE_SHEET) return;
  try {
    syncReferralFees();
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(err.message || err), '紹介料管理 更新エラー', 10);
  }
}

// ---------------------------------------------------------------- メイン

function syncReferralFees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const dst = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!src) throw new Error(`シート「${CONFIG.SOURCE_SHEET}」が見つかりません`);
  if (!dst) throw new Error(`シート「${CONFIG.TARGET_SHEET}」が見つかりません`);

  const employees = readEmployees_(src);
  const warnings = [];

  const byId = new Map(employees.map(e => [e.id, e]));
  const byName = new Map();
  employees.forEach(e => {
    const key = normalizeName_(e.fullName);
    if (!key) return;
    if (byName.has(key)) warnings.push(`同姓同名: ${e.fullName}（${byName.get(key).id} と ${e.id}）`);
    byName.set(key, e);
  });

  employees.forEach(e => {
    const r = resolveReferrer_(e.referrerRaw, byId, byName);
    e.referrerId = r.id;
    e.referrerName = r.name;
    if (r.unresolved) warnings.push(`紹介者が見つかりません: ${e.id} ${e.fullName} → 「${e.referrerRaw}」`);
  });

  assignRanks_(employees);

  const excluded = new Set(CONFIG.EXCLUDED_IDS.map(String));
  const rows = employees
    .filter(e => CONFIG.INCLUDE_WITHOUT_REFERRER || e.referrerId)
    .sort((a, b) => a.id.localeCompare(b.id))
    .map(e => buildRow_(e, excluded));

  writeRows_(dst, rows);

  const msg = `${rows.length} 件を更新しました` + (warnings.length ? `\n要確認 ${warnings.length} 件:\n` + warnings.join('\n') : '');
  ss.toast(msg, '紹介料管理', warnings.length ? 15 : 5);
  if (warnings.length) Logger.log(warnings.join('\n'));
}

// ---------------------------------------------------------------- 読み込み

function readEmployees_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW) return [];

  const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const col = name => {
    const i = headers.indexOf(name);
    if (i < 0) throw new Error(`「${CONFIG.SOURCE_SHEET}」に見出し「${name}」がありません`);
    return i;
  };
  const c = {
    id: col(CONFIG.SRC.id),
    last: col(CONFIG.SRC.lastName),
    first: col(CONFIG.SRC.firstName),
    hire: col(CONFIG.SRC.hireDate),
    ref: col(CONFIG.SRC.referrer),
  };

  const values = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, lastCol).getValues();
  return values
    .map(r => {
      const id = normalizeId_(r[c.id]);
      const last = String(r[c.last] || '').trim();
      const first = String(r[c.first] || '').trim();
      return {
        id,
        fullName: [last, first].filter(Boolean).join(CONFIG.NAME_SEPARATOR),
        hireDate: toDate_(r[c.hire]),
        referrerRaw: String(r[c.ref] == null ? '' : r[c.ref]).trim(),
        referrerId: '',
        referrerName: '',
        rank: null,
      };
    })
    .filter(e => e.id);
}

function resolveReferrer_(raw, byId, byName) {
  const s = String(raw || '').trim();
  if (!s || s === '-' || s === '－' || s === 'ー') return { id: '', name: '', unresolved: false };

  // 社員番号で指定されている場合
  const asId = normalizeId_(s);
  if (byId.has(asId)) return { id: asId, name: byId.get(asId).fullName, unresolved: false };

  // 氏名で指定されている場合
  const hit = byName.get(normalizeName_(s));
  if (hit) return { id: hit.id, name: hit.fullName, unresolved: false };

  return { id: '', name: s + '（未一致）', unresolved: true };
}

// ---------------------------------------------------------------- 計算

/** 同じ紹介者に紹介された人を入社日順（同日は社員番号順）に並べて順位を付ける */
function assignRanks_(employees) {
  const groups = new Map();
  employees.forEach(e => {
    if (!e.referrerId) return;
    if (!groups.has(e.referrerId)) groups.set(e.referrerId, []);
    groups.get(e.referrerId).push(e);
  });
  groups.forEach(list => {
    list.sort((a, b) => {
      const ta = a.hireDate ? a.hireDate.getTime() : Infinity;
      const tb = b.hireDate ? b.hireDate.getTime() : Infinity;
      if (ta !== tb) return ta - tb;
      return a.id.localeCompare(b.id);
    });
    list.forEach((e, i) => { e.rank = i + 1; });
  });
}

function feeForRank_(rank) {
  if (!rank) return '';
  return Object.prototype.hasOwnProperty.call(CONFIG.FEE_BY_RANK, rank) ? CONFIG.FEE_BY_RANK[rank] : CONFIG.FEE_DEFAULT;
}

function paymentMonths_(hireDate) {
  const n = CONFIG.DST.months.length;
  if (!hireDate) return Array(n).fill('');
  const out = [];
  for (let i = 0; i < n; i++) {
    const offset = CONFIG.FIRST_PAYMENT_OFFSET_MONTHS + i * CONFIG.PAYMENT_INTERVAL_MONTHS;
    out.push(new Date(hireDate.getFullYear(), hireDate.getMonth() + offset, 1));
  }
  return out;
}

function buildRow_(e, excluded) {
  const eligible = e.rank && !excluded.has(e.id);
  return {
    id: e.id,
    name: e.fullName,
    hireDate: e.hireDate || '',
    referrerId: e.referrerId,
    referrerName: e.referrerName,
    rank: e.rank || '',
    fee: eligible ? feeForRank_(e.rank) : '',
    months: eligible ? paymentMonths_(e.hireDate) : Array(CONFIG.DST.months.length).fill(''),
  };
}

// ---------------------------------------------------------------- 書き込み

function writeRows_(dst, rows) {
  const lastCol = dst.getLastColumn();
  const headers = dst.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const col = name => {
    const i = headers.indexOf(name);
    if (i < 0) throw new Error(`「${CONFIG.TARGET_SHEET}」に見出し「${name}」がありません`);
    return i + 1; // 1 始まり
  };
  const D = CONFIG.DST;
  const columns = [
    { c: col(D.id),           fmt: '@',                 get: r => r.id },
    { c: col(D.name),         fmt: null,                get: r => r.name },
    { c: col(D.hireDate),     fmt: CONFIG.DATE_FORMAT,  get: r => r.hireDate },
    { c: col(D.referrerId),   fmt: '@',                 get: r => r.referrerId },
    { c: col(D.referrerName), fmt: null,                get: r => r.referrerName },
    { c: col(D.rank),         fmt: '0',                 get: r => r.rank },
    { c: col(D.fee),          fmt: CONFIG.FEE_FORMAT,   get: r => r.fee },
  ];
  D.months.forEach((h, i) => columns.push({ c: col(h), fmt: CONFIG.MONTH_FORMAT, get: r => r.months[i], isMonth: true }));

  const start = CONFIG.DATA_START_ROW;
  const oldCount = Math.max(dst.getLastRow() - start + 1, 0);
  const newCount = rows.length;
  const clearCount = Math.max(oldCount, newCount);

  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  columns.forEach(({ c, fmt, get, isMonth }) => {
    if (clearCount > 0) {
      const clearRange = dst.getRange(start, c, clearCount, 1);
      clearRange.clearContent();
      if (isMonth && CONFIG.SHADE_PAST_MONTHS) clearRange.setBackground(null);
    }
    if (newCount === 0) return;

    const range = dst.getRange(start, c, newCount, 1);
    if (fmt) range.setNumberFormat(fmt);
    range.setValues(rows.map(r => [get(r)]));

    if (isMonth && CONFIG.SHADE_PAST_MONTHS) {
      range.setBackgrounds(rows.map(r => {
        const v = get(r);
        return [v instanceof Date && v <= todayStart ? CONFIG.PAST_MONTH_COLOR : null];
      }));
    }
  });
}

// ---------------------------------------------------------------- ユーティリティ

function normalizeId_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(v).padStart(CONFIG.ID_DIGITS, '0');
  const s = String(v).trim();
  return /^\d+$/.test(s) && s.length < CONFIG.ID_DIGITS ? s.padStart(CONFIG.ID_DIGITS, '0') : s;
}

/** 氏名の比較用: 全角/半角スペースを全て除去 */
function normalizeName_(s) {
  return String(s || '').replace(/[\s　]/g, '');
}

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v == null ? '' : v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
