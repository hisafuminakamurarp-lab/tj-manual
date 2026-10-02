/**
 * 代理店進捗管理シート：セットアップ・連動スクリプト（Google Apps Script）
 *
 *   - 担当者ログ（前田紘孝：代理店・浅沼雄太：代理店）… 代理店1社につき1行。代理店とのやりとりを入力し、
 *                       紹介件数・進行中・受注・対応漏れ・紹介料は営業管理から自動で数える
 *   - 代理店全体ログ  … 担当者ログを自動で合算。上部に担当者別・合計のステータス
 *   - 紹介・成約実績  … 営業管理のうち「紹介元の代理店」が入った案件を自動で一覧化
 *   - サマリー        … 代理店別集計（稼働率・紹介件数・受注・金額）
 *   - 代理店ID        … kintone『8. 代理店管理』へのリンク（入力すると自動でリンク化）
 *   - 営業管理（担当者ごと）… 顧客の営業進捗を入力。代理店からの紹介案件もここに入れる（営業管理.gs）
 *   - 営業管理全体ログ … 担当者ごとの営業管理を自動で合算。上部に担当者別・合計のステータス（営業管理.gs）
 *
 * 使い方：拡張機能 → Apps Script に貼り付け → fixAgencySheet を実行。
 * 何度実行しても同じ状態になる（入力済みのデータは消さない）。
 * 『使い方』『定例アクション計画』『設定』の説明を作り直すときだけ updateGuides を実行。
 */

// 自社の担当者。増えたらここに追加して fixAgencySheet を実行（「名前：代理店」のタブが自動で作られる）
const ASSIGNEES = ['前田紘孝', '浅沼雄太'];
const TAB = name => `${name}：代理店`;

const SH = {
  log: '代理店全体ログ',
  kpi: '紹介・成約実績',
  sum: 'サマリー',
  plan: '定例アクション計画',
  howto: '使い方',
  conf: '設定',
};

// 担当者ログ：3〜4行目＝ステータス、6行目＝見出し、7行目〜＝データ（代理店1社につき1行）
const P_HEAD = 6, P_FIRST = 7, P_ROWS = 300;
const P_LAST = P_FIRST + P_ROWS - 1;
// [見出し, 列幅, 自動なら 'auto']
const P_COLS = [
  ['代理店ID', 100], ['代理店名', 200], ['代理店担当者', 110],
  ['連絡ツール', 110], ['接触手段', 110], ['最終アクション日', 100], ['最終アクション内容', 260],
  ['次回アクション予定日', 110], ['次回アクション内容', 260],
  ['紹介件数', 65, 'auto'], ['進行中', 65, 'auto'], ['受注', 65, 'auto'], ['対応漏れ', 65, 'auto'],
  ['受注金額（円）', 110, 'auto'], ['紹介料（円）', 100, 'auto'], ['紹介料未払（円）', 100, 'auto'],
  ['備考', 220],
];
const PC = {  // 担当者ログの列
  id: 'A', name: 'B', tool: 'D', method: 'E', lastDate: 'F', lastText: 'G', nextDate: 'H', nextText: 'I',
  count: 'J', open: 'K', won: 'L', miss: 'M', amount: 'N', fee: 'O', unpaid: 'P',
};
const TOOL_OPTIONS = ['グループLINE', '個別LINE', 'Chatwork', 'メール', '電話'];
const METHOD_OPTIONS = ['電話', '訪問', 'オンライン面談', 'メール', 'LINE/チャット', '勉強会・セミナー', '会食'];
const STATUS_LABELS = ['担当代理店数', '要連絡（赤）', '期限間近（黄）', '今月の接触',
  '紹介件数', '進行中', '受注', '対応漏れ', '受注金額（円）', '紹介料未払（円）'];

// 代理店全体ログ：3行目〜＝担当者別ステータス表、その2行下が見出し、次の行から合算データ
const LOG_HEAD = ASSIGNEES.length + 6;
const LOG_FIRST = LOG_HEAD + 1;
const LOG_END = LOG_FIRST + ASSIGNEES.length * P_ROWS - 1;
const LOG_COLS = P_COLS.slice(0, 2).concat([['自社担当', 90]], P_COLS.slice(2));  // 自社担当をC列に挿入

// 紹介・成約実績：4行目＝見出し、5行目〜＝営業管理全体ログから自動（紹介元の代理店が入った案件だけ）
const KPI_HEADS = ['紹介日', '紹介元の代理店', '自社担当', '会社名', '単発／顧問', '提案商材', '進捗',
  '次回アポ日', '状況', '受注金額（円）', '紹介料（円）', '紹介料支払', 'メモ'];
const KPI_FIRST = 5;
const KPI_END = KPI_FIRST + ASSIGNEES.length * P_ROWS - 1;

const SUM_FIRST = 6;    // サマリーの代理店1行目（5行目は合計）
const SUM_END = 105;    // 代理店100社分

// kintone『8. 代理店管理』アプリのURLと、代理店IDのフィールドコード
const KINTONE_APP_URL = 'https://zeimukeeoer.cybozu.com/k/39/';
const KINTONE_ID_FIELD = '代理店ID';

// 成約商材の選択肢（kintoneのドロップダウンと同じ内容）。kintone側で増やしたらここにも追加して再実行
const ANKEN1_OPTIONS = [
  '節税ショット', '助成金', '補助金', '税務調査', '資金調達', '金融コンサル', '経営コンサル',
  '社会保険料削減スキーム', '蓄電池', '決算申告', '確定申告', '記帳代行',
];
const ANKEN2_OPTIONS = [
  'GPUサーバー', 'モバイルファンド', '大型ビジョン', 'デジタルサイネージ', 'Elan Media 広告',
  '税務キーパー決算書診断', 'ものづくり補助金', 'IT導入補助金', 'キャリアアップ助成金',
  '事業展開等リスキングコース', 'UTGL', 'ファミリーオフィス', 'ジンバブエ投資(ゴールド)',
  'エンジェル税制', 'キャプティブ', 'アイヌ', '同和', 'ドバイ不動産', '税務調査決算書診断',
  '税務調査立会', '模擬調査',
];
const OPT_FIRST = 14, OPT_END = 63;  // 『設定』の選択肢リスト行（50件まで）

const COLOR = {
  title: '#1F3864', input: '#FFF2CC', auto: '#D9D9D9', red: '#F8CBAD', redFont: '#9C0006',
  yellow: '#FFE699', total: '#DDEBF7', gray: '#E7E6E6', green: '#C6EFCE', border: '#BFBFBF',
};

function fixAgencySheet() {
  const ss = SpreadsheetApp.getActive();
  const need = [SH.log, SH.kpi, SH.sum, SH.conf].filter(n => !ss.getSheetByName(n));
  if (need.length) throw new Error('シートが見つかりません：' + need.join('、'));

  const notes = writeLists_(ss);
  ASSIGNEES.forEach(name => {
    if (ss.getSheetByName(TAB(name))) return;
    // 以前の名前（「名前ログ」「名前だけ」）のタブがあれば改名して使う
    const plain = ss.getSheetByName(`${name}ログ`) || ss.getSheetByName(name);
    if (plain) plain.setName(TAB(name)); else ss.insertSheet(TAB(name), 0);
  });
  if (!migrateOldLog_(ss)) return;  // 旧形式のデータ移行（担当者未入力なら中断）

  setupSales_(ss);                     // 営業管理を先に整える（代理店側の自動集計が参照するため）
  notes.push(...migrateReferrals_(ss));  // 担当者ログの紹介データを営業管理へ移す（1回だけ）
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(TAB(name));
    remapColumns_(sh, P_COLS.map(([h]) => h), { '備考': ['備考', 'メモ'] });
    setupPersonTab_(sh, name);
  });
  buildLog_(ss.getSheetByName(SH.log));
  notes.push(...buildKpi_(ss));
  fixSummary_(ss.getSheetByName(SH.sum));
  linkKintoneIds();
  SpreadsheetApp.flush();

  const broken = findBrokenFormulas_(ss);
  let msg = broken.length
    ? '修正しました。ただし参照切れの数式が残っています：\n' + broken.slice(0, 20).join('\n')
    : '完了しました。参照切れの数式はありません。';
  if (notes.length) msg += '\n\n' + notes.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

// ---------------------------------------------------------------- 旧形式からの移行
// 旧形式（4行目が見出し、D列＝自社担当）のデータを、自社担当ごとに担当者ログへ移す
function migrateOldLog_(ss) {
  const log = ss.getSheetByName(SH.log);
  const head = log.getRange('A4:J4').getDisplayValues()[0];
  if (head[0] !== '代理店ID' || head[3] !== '自社担当') return true;  // 移行済み

  const last = log.getLastRow();
  const rows = last >= 5 ? log.getRange(5, 1, last - 4, 10).getValues() : [];
  const data = rows.filter(r => String(r[0]).trim() !== '');
  const unassigned = data.filter(r => !ASSIGNEES.includes(String(r[3]).trim()));
  if (unassigned.length) {
    setListValidation_(log.getRange(5, 4, Math.max(last - 4, 1), 1),
      ss.getSheetByName(SH.conf).getRange(`G${OPT_FIRST}:G${OPT_FIRST + 9}`));
    SpreadsheetApp.getUi().alert(
      `担当者ログへの移行を止めました。\n\n『${SH.log}』D列（自社担当）が空欄、または担当者名以外の代理店があります：\n` +
      unassigned.map(r => `${r[0]}　${r[1]}`).join('\n') +
      `\n\nD列のプルダウンで「${ASSIGNEES.join('」か「')}」を選んでから、もう一度 fixAgencySheet を実行してください。`);
    return false;
  }

  if (!ss.getSheetByName(`${SH.log}（移行前）`)) {
    log.copyTo(ss).setName(`${SH.log}（移行前）`);
  }
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(TAB(name));
    // 旧：A ID / B 代理店名 / C 代理店担当者 / E 連絡ツール / F 接触手段 / G〜J 最終・次回アクション
    const mine = data.filter(r => String(r[3]).trim() === name)
      .map(r => [r[0], r[1], r[2], r[4], r[5], r[6], r[7], r[8], r[9]]);
    if (!mine.length) return;
    // 見出しを新しい形にしてから書き込む（remapColumns_ で組み替え対象にしない）
    sh.getRange(P_HEAD, 1, 1, P_COLS.length).setValues([P_COLS.map(([h]) => h)]);
    const filled = sh.getRange(P_FIRST, 1, P_ROWS, 1).getValues().filter(([v]) => v !== '').length;
    sh.getRange(P_FIRST + filled, 1, mine.length, mine[0].length).setValues(mine);
  });

  if (log.getFilter()) log.getFilter().remove();
  log.setConditionalFormatRules([]);
  log.getRange(3, 1, log.getMaxRows() - 2, log.getMaxColumns()).clear().clearDataValidations();
  return true;
}

// 見出しが今の列構成と違うとき、見出し名を手がかりにデータを新しい列へ並べ替える（担当者ログ・営業管理で共通）
// （使わなくなった列・自動計算の列は捨てる。alias＝旧名 → 新名）
function remapColumns_(sh, want, alias) {
  const norm = v => String(v).replace(/\s/g, '');
  const width = Math.max(sh.getLastColumn(), want.length);
  const head = sh.getRange(P_HEAD, 1, 1, width).getDisplayValues()[0].map(norm);
  if (want.every((h, i) => head[i] === norm(h))) return;
  if (head.every(v => v === '')) return;  // 新規タブ

  const rng = sh.getRange(P_FIRST, 1, P_ROWS, width);
  const values = rng.getValues();
  const formulas = rng.getFormulas();
  const src = want.map(h => {
    const names = ((alias || {})[h] || [h]).map(norm);
    return head.findIndex(v => names.includes(v));
  });
  const rows = values.map((row, r) => src.map(i => (i < 0 || formulas[r][i]) ? '' : row[i]));

  const ss = sh.getParent();
  const backup = `${sh.getName()}（組み替え前）`;
  if (!ss.getSheetByName(backup)) sh.copyTo(ss).setName(backup);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  sh.getRange(P_FIRST, 1, P_ROWS, sh.getMaxColumns()).clear().clearDataValidations();
  sh.getRange(P_FIRST, 1, P_ROWS, want.length).setValues(rows);
  sh.getRange(P_HEAD, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(P_HEAD, 1, 1, want.length).setValues([want]);
}

// 担当者ログに「紹介社名」などの列があった頃の形 → 紹介は担当者の営業管理へ1件1行で移し、
// 代理店の行は1社1行にまとめる（同じ代理店の行は、最終アクション日がいちばん新しい行の内容を残す）
function migrateReferrals_(ss) {
  const notes = [];
  const tz = ss.getSpreadsheetTimeZone();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy/MM/dd') : String(d);
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(TAB(name));
    const width = sh.getLastColumn();
    if (width < 1) return;
    const head = sh.getRange(P_HEAD, 1, 1, width).getDisplayValues()[0].map(v => v.replace(/\s/g, ''));
    if (!head.includes('紹介社名')) return;

    const backup = `${sh.getName()}（紹介移行前）`;
    if (!ss.getSheetByName(backup)) sh.copyTo(ss).setName(backup);
    const rows = sh.getRange(P_FIRST, 1, P_ROWS, width).getValues();
    const g = (row, h) => { const i = head.indexOf(h.replace(/\s/g, '')); return i < 0 ? '' : row[i]; };
    const has = v => String(v).trim() !== '';

    // 紹介 → 営業管理（1件1行）。代理店名が空なら、同じ代理店IDの別の行から補う
    const nameOf = row => {
      if (has(g(row, '代理店名'))) return g(row, '代理店名');
      const same = rows.find(x => has(g(x, '代理店名')) && has(g(row, '代理店ID')) &&
        String(g(x, '代理店ID')).trim() === String(g(row, '代理店ID')).trim());
      return same ? g(same, '代理店名') : '';
    };
    const deals = [];
    rows.filter(row => has(g(row, '紹介社名'))).forEach(row => {
      const komon = g(row, '成約商材（顧問種別）'), a1 = g(row, '成約商材（案件種別①）'), a2 = g(row, '成約商材（案件種別②）');
      const close = g(row, '成約日');
      const memo = [has(komon) ? `顧問種別：${komon}` : '', has(a1) && has(a2) ? `案件種別：${a1}／${a2}` : '',
        has(close) ? `成約日：${fmt(close)}` : ''].filter(has).join('　');
      deals.push({
        '会社名': g(row, '紹介社名'), '紹介元の代理店': nameOf(row), '紹介日': g(row, '紹介日'),
        '単発／顧問': has(komon) ? '顧問' : '', '提案商材': has(a2) ? a2 : a1, '進捗': has(close) ? '受注' : '',
        '受注金額（円）': g(row, '成約金額（円）'), '紹介料（円）': g(row, '代理店報酬（円）'), 'メモ': memo,
      });
    });
    const moved = appendDeals_(findSalesTab_(ss, name), deals);

    // 代理店の行 → 1社1行
    const groups = new Map();
    rows.forEach(row => {
      const key = String(g(row, '代理店ID')).trim() || String(g(row, '代理店名')).trim();
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(row);
    });
    const firstOf = (list, h) => { const row = list.find(x => has(g(x, h))); return row ? g(row, h) : ''; };
    const out = [...groups.values()].map(list => {
      const acted = list.filter(x => ['最終アクション日', '最終アクション内容', '次回アクション予定日', '次回アクション内容']
        .some(h => has(g(x, h))));
      const time = x => g(x, '最終アクション日') instanceof Date ? g(x, '最終アクション日').getTime() : 0;
      const latest = acted.reduce((a, x) => (!a || time(x) >= time(a)) ? x : a, null);
      const notesCol = [...new Set(list.map(x => String(g(x, '備考') || g(x, 'メモ')).trim()).filter(has))].join(' / ');
      return P_COLS.map(([h, , kind]) => {
        if (kind === 'auto') return '';
        if (h === '備考') return notesCol;
        if (/アクション/.test(h)) return latest ? g(latest, h) : '';
        return firstOf(list, h);
      });
    });

    if (sh.getFilter()) sh.getFilter().remove();
    sh.setConditionalFormatRules([]);
    sh.getRange(P_FIRST, 1, P_ROWS, sh.getMaxColumns()).clear().clearDataValidations();
    sh.getRange(P_HEAD, 1, 1, sh.getMaxColumns()).clearContent();
    sh.getRange(P_HEAD, 1, 1, P_COLS.length).setValues([P_COLS.map(([h]) => h)]);
    if (out.length) sh.getRange(P_FIRST, 1, out.length, P_COLS.length).setValues(out);
    const before = rows.filter(row => row.some(has)).length;
    notes.push(`※『${sh.getName()}』の紹介${deals.length}件を『${findSalesTab_(ss, name).getName()}』へ移しました` +
      `（追加${moved}件・同じ会社名はスキップ）。代理店の行は${before}行→${out.length}行（1社1行）にまとめました。元の内容は『${backup}』に保存。`);
  });
  return notes;
}

// ---------------------------------------------------------------- 担当者ログ
function setupPersonTab_(sh, name) {
  const ss = sh.getParent();
  const nCol = P_COLS.length;
  ensureSize_(sh, P_LAST, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(P_FIRST, 1, P_ROWS, sh.getMaxColumns()).clearDataValidations();
  titleRow_(sh, `${name}の担当代理店`, nCol,
    '代理店1社につき1行。接触したら「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。' +
    '紹介いただいた会社は営業管理に入力し「紹介元の代理店」を選ぶと、灰色の列（紹介件数〜紹介料未払）に自動で反映されます。');

  // 自動の列：営業管理全体ログのうち、紹介元の代理店がこの行の代理店名と同じ案件を数える
  const r = col => `${col}${P_FIRST}:${col}${P_LAST}`;
  const nm = r(PC.name);
  const ag = salesRange_(ss, '紹介元の代理店'), pr = salesRange_(ss, '進捗'), st = salesRange_(ss, '状況');
  const am = salesRange_(ss, '受注金額（円）'), fe = salesRange_(ss, '紹介料（円）'), pd = salesRange_(ss, '紹介料支払');
  const each = expr => `=ARRAYFORMULA(IF(${nm}="","",${expr}))`;
  sh.getRange(`${PC.count}${P_FIRST}:${PC.unpaid}${P_LAST}`).clearContent();
  sh.getRange(`${PC.count}${P_FIRST}:${PC.unpaid}${P_FIRST}`).setFormulas([[
    each(`COUNTIFS(${ag},${nm})`),
    each(`COUNTIFS(${ag},${nm},${pr},"<>受注",${pr},"<>失注")`),
    each(`COUNTIFS(${ag},${nm},${pr},"受注")`),
    each(`COUNTIFS(${ag},${nm},${st},"アポなし")+COUNTIFS(${ag},${nm},${st},"期限切れ")`),
    each(`SUMIFS(${am},${ag},${nm})`),
    each(`SUMIFS(${fe},${ag},${nm})`),
    each(`SUMIFS(${fe},${ag},${nm},${pd},"<>支払済")`),
  ]]);

  // ステータス（3〜4行目）
  sh.getRange(3, 1, 1, sh.getMaxColumns()).clear();
  sh.getRange(4, 1, 1, sh.getMaxColumns()).clear();
  sh.getRange(3, 1, 1, STATUS_LABELS.length).setValues([STATUS_LABELS])
    .setBackground(COLOR.auto).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setWrap(true);
  sh.getRange(4, 1, 1, STATUS_LABELS.length).setFormulas([[
    `=IFERROR(ROWS(UNIQUE(FILTER(${nm},${nm}<>""))),0)`,  // 空のときは0（COUNTUNIQUEだとエラーを1件と数える）
    `=COUNTIFS(${nm},"?*",${r(PC.nextDate)},"")+COUNTIFS(${nm},"?*",${r(PC.nextDate)},"<"&TODAY())`,
    `=COUNTIFS(${nm},"?*",${r(PC.nextDate)},">="&TODAY(),${r(PC.nextDate)},"<="&TODAY()+設定!$C$8)`,
    `=COUNTIFS(${nm},"?*",${r(PC.lastDate)},">="&設定!$C$5)`,
    `=SUM(${r(PC.count)})`,
    `=SUM(${r(PC.open)})`,
    `=SUM(${r(PC.won)})`,
    `=SUM(${r(PC.miss)})`,
    `=SUM(${r(PC.amount)})`,
    `=SUM(${r(PC.unpaid)})`,
  ]]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange(3, 1, 2, STATUS_LABELS.length)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(3, 30);

  headerRow_(sh, P_HEAD, P_COLS.map(([h, w, kind]) => [h, kind || 'in', w]));
  const body = sh.getRange(P_FIRST, 1, P_ROWS, nCol);
  body.setNumberFormat('General').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  [PC.lastDate, PC.nextDate].forEach(col => sh.getRange(r(col)).setNumberFormat('yyyy/mm/dd'));
  [PC.lastText, PC.nextText].forEach(col => sh.getRange(r(col)).setWrap(true));
  sh.getRange(`${PC.count}${P_FIRST}:${PC.miss}${P_LAST}`).setNumberFormat('0;-0;"-"').setHorizontalAlignment('center');
  sh.getRange(`${PC.amount}${P_FIRST}:${PC.unpaid}${P_LAST}`).setNumberFormat('#,##0;-#,##0;"-"');
  sh.getRange(`${PC.count}${P_FIRST}:${PC.unpaid}${P_LAST}`).setBackground('#F3F3F3');

  // プルダウン（選択肢は『設定』）
  const conf = ss.getSheetByName(SH.conf);
  const list = col => conf.getRange(`${col}${OPT_FIRST}:${col}${OPT_FIRST + 9}`);
  setListValidation_(sh.getRange(r(PC.tool)), list('K'));
  setListValidation_(sh.getRange(r(PC.method)), list('H'));

  // 色分け：次回アクション予定日が未入力・期限切れ＝D〜I列を赤、期限間近＝H〜I列を黄
  //         対応漏れがある＝赤、紹介料の未払がある＝オレンジ（ステータスも同じ）
  const f = P_FIRST;
  const status = label => sh.getRange(4, STATUS_LABELS.indexOf(label) + 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(COUNTA($A${f}:$I${f})>0,OR($H${f}="",$H${f}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(`D${f}:I${P_LAST}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($H${f}<>"",$H${f}>=TODAY(),$H${f}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow).setRanges([sh.getRange(`H${f}:I${P_LAST}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont).setBold(true)
      .setRanges([sh.getRange(r(PC.miss)), status('要連絡（赤）'), status('対応漏れ')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground('#FCE5CD').setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(r(PC.unpaid)), status('紹介料未払（円）')]).build(),
  ]);

  sh.setFrozenRows(P_HEAD);
  sh.setFrozenColumns(2);
  sh.getRange(P_HEAD, 1, P_ROWS + 1, nCol).createFilter();
}

// ---------------------------------------------------------------- 代理店全体ログ（合算）
function buildLog_(sh) {
  const nCol = LOG_COLS.length;
  ensureSize_(sh, LOG_END, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns()).clear().clearDataValidations();
  titleRow_(sh, `${SH.log}（全担当者の合算）`, nCol,
    `${ASSIGNEES.map(TAB).join('・')}の内容を自動で合算した一覧です（ここでは入力しません。修正は各担当者ログで）。上の表は担当者別と合計のステータス。`);

  // 担当者別ステータス表（3行目〜）
  const n = ASSIGNEES.length;
  sh.getRange(3, 1, 1, STATUS_LABELS.length + 1).setValues([['担当者'].concat(STATUS_LABELS)])
    .setBackground(COLOR.auto).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  const cols = STATUS_LABELS.map((_, i) => colLetter_(i + 1));
  // 名前・「合計」は文字として、数値は数式として別々に書く（setFormulas に文字を渡すと #NAME? になる）
  sh.getRange(4, 1, n, 1).setValues(ASSIGNEES.map(name => [name]));
  sh.getRange(4, 2, n, STATUS_LABELS.length).setFormulas(
    ASSIGNEES.map(name => cols.map(col => `='${TAB(name)}'!${col}4`)));
  const totalRow = 4 + n;
  sh.getRange(totalRow, 1).setValue('合計');
  sh.getRange(totalRow, 2, 1, STATUS_LABELS.length).setFormulas([
    STATUS_LABELS.map((_, i) => {
      const col = colLetter_(i + 2);
      return `=SUM(${col}4:${col}${totalRow - 1})`;
    })]);
  sh.getRange(totalRow, 1, 1, STATUS_LABELS.length + 1).setBackground(COLOR.total).setFontWeight('bold');
  sh.getRange(4, 2, n + 1, STATUS_LABELS.length).setHorizontalAlignment('center').setNumberFormat('#,##0');
  ['要連絡（赤）', '対応漏れ'].forEach(label =>
    sh.getRange(4, STATUS_LABELS.indexOf(label) + 2, n + 1, 1).setFontColor(COLOR.redFont).setFontWeight('bold'));
  sh.setRowHeight(3, 30);
  sh.getRange(3, 1, 1, STATUS_LABELS.length + 1).setWrap(true);
  sh.getRange(3, 1, n + 2, STATUS_LABELS.length + 1)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);

  headerRow_(sh, LOG_HEAD, LOG_COLS.map(([h, w]) => [h, 'auto', w]));

  // 担当者ログを縦に合算（自社担当の名前をC列に入れる）。代理店IDはkintoneリンク付き
  const rest = P_COLS.slice(2).map((_, i) => i + 3).join(',');
  const used = x => `((CHOOSECOLS(${x},1)<>"")+(CHOOSECOLS(${x},2)<>""))`;  // 代理店IDか代理店名が入っている行
  const parts = ASSIGNEES.map((name, i) =>
    `HSTACK(CHOOSECOLS(p${i},1,2),IF(${used(`p${i}`)},"${name}",""),CHOOSECOLS(p${i},${rest}))`);
  const lets = ASSIGNEES.map((name, i) =>
    `p${i},'${TAB(name)}'!A${P_FIRST}:${colLetter_(P_COLS.length)}${P_LAST}`).join(',');
  const others = LOG_COLS.slice(1).map((_, i) => i + 2).join(',');
  sh.getRange(LOG_FIRST, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(LET(${lets},all,VSTACK(${parts.join(',')}),f,FILTER(all,${used('all')}),` +
    `HSTACK(${idLink_('CHOOSECOLS(f,1)')},CHOOSECOLS(f,${others})))),"")`);

  // 書式（担当者ログの列が自社担当の分だけ右にずれる）
  const L = h => colLetter_(LOG_COLS.findIndex(([x]) => x === h) + 1);
  const rr = col => sh.getRange(`${col}${LOG_FIRST}:${col}${LOG_END}`);
  sh.getRange(LOG_FIRST, 1, LOG_END - LOG_FIRST + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['最終アクション日', '次回アクション予定日'].forEach(h => rr(L(h)).setNumberFormat('yyyy/mm/dd'));
  ['紹介件数', '進行中', '受注', '対応漏れ'].forEach(h => rr(L(h)).setNumberFormat('0;-0;"-"').setHorizontalAlignment('center'));
  ['受注金額（円）', '紹介料（円）', '紹介料未払（円）'].forEach(h => rr(L(h)).setNumberFormat('#,##0;-#,##0;"-"'));
  ['最終アクション内容', '次回アクション内容'].forEach(h => rr(L(h)).setWrap(true));
  sh.getRange(`C${LOG_FIRST}:C${LOG_END}`).setHorizontalAlignment('center');

  // 色分け（担当者ログと同じ基準）
  const f = LOG_FIRST, next = L('次回アクション予定日');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(OR($A${f}<>"",$B${f}<>""),OR($${next}${f}="",$${next}${f}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(`${L('連絡ツール')}${f}:${L('次回アクション内容')}${LOG_END}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${next}${f}<>"",$${next}${f}>=TODAY(),$${next}${f}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow)
      .setRanges([sh.getRange(`${next}${f}:${L('次回アクション内容')}${LOG_END}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont).setBold(true)
      .setRanges([rr(L('対応漏れ'))]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground('#FCE5CD').setFontColor(COLOR.redFont)
      .setRanges([rr(L('紹介料未払（円）'))]).build(),
  ]);
  sh.setFrozenRows(LOG_HEAD);
  sh.setFrozenColumns(2);
}

// ---------------------------------------------------------------- 紹介・成約実績（自動）
function buildKpi_(ss) {
  const notes = [];
  const sh = ss.getSheetByName(SH.kpi);
  // 以前の手入力データが残っていれば、消す前にタブごとバックアップ
  const a5 = sh.getRange(`A${KPI_FIRST}`);
  if (!a5.getFormula() && sh.getLastRow() >= KPI_FIRST &&
      sh.getRange(KPI_FIRST, 1, sh.getLastRow() - KPI_FIRST + 1, 1).getValues().some(([v]) => v !== '')) {
    const backup = `${SH.kpi}（移行前）`;
    if (!ss.getSheetByName(backup)) sh.copyTo(ss).setName(backup);
    notes.push(`※『${SH.kpi}』に手入力の紹介データがあったため『${backup}』に退避しました。紹介は各担当者の営業管理に入力してください。`);
  }

  const nCol = KPI_HEADS.length;
  ensureSize_(sh, KPI_END, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns()).clear().clearDataValidations();
  titleRow_(sh, `${SH.kpi}（紹介1件につき1行・自動）`, nCol,
    '営業管理で「紹介元の代理店」を選んだ案件が、ここに自動で並びます（ここでは入力しません。修正は各担当者の営業管理で）。');

  const widths = { '紹介元の代理店': 180, '自社担当': 90, '会社名': 200, '提案商材': 150, 'メモ': 260 };
  headerRow_(sh, 4, KPI_HEADS.map(h => [h, 'auto', widths[h] || 90]));

  // 営業管理全体ログの列（A列＝担当、B列以降＝営業管理の列）から必要な列だけ取り出す
  const src = KPI_HEADS.map(h => h === '自社担当' ? 1 : S_COLS.findIndex(([x]) => x === h) + 2);
  const agency = S_COLS.findIndex(([x]) => x === '紹介元の代理店') + 2;
  sh.getRange(KPI_FIRST, 1).setFormula(
    `=IFERROR(LET(d,${salesRange_(ss)},SORT(FILTER(CHOOSECOLS(d,${src.join(',')}),CHOOSECOLS(d,${agency})<>""),1,TRUE)),"")`);

  const K = h => colLetter_(KPI_HEADS.indexOf(h) + 1);
  const rr = h => sh.getRange(`${K(h)}${KPI_FIRST}:${K(h)}${KPI_END}`);
  sh.getRange(KPI_FIRST, 1, KPI_END - KPI_FIRST + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['紹介日', '次回アポ日'].forEach(h => rr(h).setNumberFormat('yyyy/mm/dd'));
  ['受注金額（円）', '紹介料（円）'].forEach(h => rr(h).setNumberFormat('#,##0'));
  ['自社担当', '単発／顧問', '進捗', '状況', '紹介料支払'].forEach(h => rr(h).setHorizontalAlignment('center'));
  rr('状況').setFontWeight('bold');

  // 状況：アポなし・期限切れ＝赤、今週アポ＝黄、受注＝緑。紹介料：金額があって未払＝オレンジ
  const st = `$${K('状況')}${KPI_FIRST}`;
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=OR(${st}="アポなし",${st}="期限切れ")`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont).setRanges([rr('状況')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('今週アポ')
      .setBackground(COLOR.yellow).setRanges([rr('状況')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('受注')
      .setBackground(COLOR.green).setRanges([rr('状況')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${K('紹介料（円）')}${KPI_FIRST}>0,$${K('紹介料支払')}${KPI_FIRST}<>"支払済")`)
      .setBackground('#FCE5CD').setFontColor(COLOR.redFont).setRanges([rr('紹介料支払')]).build(),
  ]);
  sh.setFrozenRows(4);
  sh.setFrozenColumns(2);
  return notes;
}

// ---------------------------------------------------------------- 『設定』の選択肢
function writeLists_(ss) {
  const notes = [];
  const conf = ss.getSheetByName(SH.conf);
  const last = OPT_FIRST + 9;
  conf.getRange('G12').setValue('入力リスト（担当者ログ・営業管理のプルダウン）').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange(`I12:J${last}`).clear();  // 使わなくなった「代理店ステータス」「紹介見込み」
  conf.getRange('G13:K13').setValues([['自社担当', '接触手段', '', '', '連絡ツール']])
    .setFontWeight('bold').setHorizontalAlignment('center');
  ['G13', 'H13', 'K13'].forEach(a => conf.getRange(a).setBackground(COLOR.input));
  writeOptions_(conf.getRange(`G${OPT_FIRST}:G${last}`), ASSIGNEES);
  // 接触手段・連絡ツールは空のときだけ初期値を入れる（手で追加した項目を消さない）
  const keep = (col, items) => {
    const rng = conf.getRange(`${col}${OPT_FIRST}:${col}${last}`);
    if (rng.getValues().every(([v]) => v === '')) writeOptions_(rng, items);
  };
  keep('H', METHOD_OPTIONS);
  keep('K', TOOL_OPTIONS);
  ['G', 'H', 'K'].forEach(col => conf.getRange(`${col}${OPT_FIRST}:${col}${last}`)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID));
  conf.getRange('G25').setValue('※ 自社担当はスクリプト上部の ASSIGNEES で管理。接触手段・連絡ツールは空欄に追記するとプルダウンに反映されます（各10件まで）。');

  // 成約商材（kintoneと同じ選択肢）。顧問種別は手入力の内容を残す
  conf.getRange('M12').setValue('商材の選択肢（kintoneと同じ内容。営業管理の提案商材＝S列の初期値）').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange('M13:O13').setValues([['顧問種別', '案件種別①', '案件種別②']])
    .setFontWeight('bold').setBackground(COLOR.input).setHorizontalAlignment('center');
  writeOptions_(conf.getRange(`N${OPT_FIRST}:N${OPT_END}`), ANKEN1_OPTIONS);
  writeOptions_(conf.getRange(`O${OPT_FIRST}:O${OPT_END}`), ANKEN2_OPTIONS);
  conf.setColumnWidths(13, 3, 180);
  return notes;
}

// ---------------------------------------------------------------- サマリー
function fixSummary_(sh) {
  const ss = sh.getParent();
  sh.getCharts().forEach(ch => sh.removeChart(ch));
  const lastRow = Math.max(sh.getMaxRows(), SUM_END);
  sh.getRange(5, 1, lastRow - 4, 7).clearContent().setBackground(null).setFontWeight('normal');

  sh.getRange('A1').setValue('サマリー（代理店別）');
  sh.getRange('A2').setValue(
    `すべて自動計算。代理店は担当者ログ（${SH.log}）から、紹介・受注は営業管理（紹介元の代理店）から集計します。` +
    '稼働率＝直近の対象期間（『設定』C6、初期値3ヶ月）のうち紹介があった月の割合。');
  sh.getRange('A4:G4').setValues([['代理店名', '自社担当', '稼働率', '紹介件数', '受注', '受注金額（円）', '紹介料（円）']]);

  const L = `'${SH.log}'`;
  const logName = `${L}!$B$${LOG_FIRST}:$B$${LOG_END}`, logWho = `${L}!$C$${LOG_FIRST}:$C$${LOG_END}`;
  const ag = salesRange_(ss, '紹介元の代理店'), intro = salesRange_(ss, '紹介日'), pr = salesRange_(ss, '進捗');
  const am = salesRange_(ss, '受注金額（円）'), fe = salesRange_(ss, '紹介料（円）');

  // 合計行
  sh.getRange('A5').setValue('合計');
  sh.getRange('B5:G5').setFormulas([[
    `=COUNTIF(A${SUM_FIRST}:A${SUM_END},"?*")&"社"`,
    `=IFERROR(AVERAGE(C${SUM_FIRST}:C${SUM_END}),"")`,
    `=SUM(D${SUM_FIRST}:D${SUM_END})`,
    `=SUM(E${SUM_FIRST}:E${SUM_END})`,
    `=SUM(F${SUM_FIRST}:F${SUM_END})`,
    `=SUM(G${SUM_FIRST}:G${SUM_END})`,
  ]]);
  sh.getRange('A5:G5').setBackground(COLOR.total).setFontWeight('bold');

  // 代理店名：代理店全体ログから重複なしで自動展開
  sh.getRange(`A${SUM_FIRST}`).setFormula(`=IFERROR(UNIQUE(FILTER(${logName},${logName}<>"")),"")`);

  const rows = [];
  for (let r = SUM_FIRST; r <= SUM_END; r++) {
    const A = `$A${r}`;
    const m = 'DATE(YEAR(設定!$C$5),MONTH(設定!$C$5)';
    const seq = 'SEQUENCE(設定!$C$6)';
    rows.push([
      `=IF(${A}="","",IFERROR(INDEX(${logWho},MATCH(${A},${logName},0)),""))`,
      `=IF(${A}="","",ARRAYFORMULA(SUMPRODUCT(--(COUNTIFS(${ag},${A},` +
        `${intro},">="&${m}+1-${seq},1),${intro},"<"&${m}+2-${seq},1))>0)))/設定!$C$6)`,
      `=IF(${A}="","",COUNTIFS(${ag},${A}))`,
      `=IF(${A}="","",COUNTIFS(${ag},${A},${pr},"受注"))`,
      `=IF(${A}="","",SUMIFS(${am},${ag},${A}))`,
      `=IF(${A}="","",SUMIFS(${fe},${ag},${A}))`,
    ]);
  }
  sh.getRange(SUM_FIRST, 2, rows.length, 6).setFormulas(rows);

  sh.getRange(5, 2, SUM_END - 4, 1).setHorizontalAlignment('center');
  sh.getRange(5, 3, SUM_END - 4, 1).setNumberFormat('0%');
  sh.getRange(5, 4, SUM_END - 4, 4).setNumberFormat('#,##0');
  sh.getRange(5, 1, SUM_END - 4, 7)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 200);
  sh.setFrozenRows(5);

  // 稼働率：0%＝グレー（休眠）、100%＝緑
  const target = sh.getRange(SUM_FIRST, 3, SUM_END - SUM_FIRST + 1, 1);
  const rules = sh.getConditionalFormatRules().filter(r => !overlaps_(r, target));
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${SUM_FIRST}<>"",$C${SUM_FIRST}=0)`)
      .setBackground(COLOR.gray).setFontColor('#7F7F7F').setRanges([target]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${SUM_FIRST}<>"",$C${SUM_FIRST}>=1)`)
      .setBackground(COLOR.green).setRanges([target]).build());
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- 説明文（updateGuides）
// 『設定』の不要部分を片付け、『定例アクション計画』『使い方』を今の構成に合わせて書き直す
function updateGuides() {
  const ss = SpreadsheetApp.getActive();
  tidySettings_(ss);
  writePlan_(ss.getSheetByName(SH.plan));
  writeHowto_(ss);
  SpreadsheetApp.getUi().alert('『設定』『定例アクション計画』『使い方』を更新しました。');
}

function tidySettings_(ss) {
  const conf = ss.getSheetByName(SH.conf);
  conf.getRange('B6').setValue('稼働率の対象期間（ヶ月）');
  conf.getRange('D6').setValue('『サマリー』の稼働率：直近この月数のうち紹介があった月の割合');
  conf.getRange('B8').setValue('期限間近の表示（日前）');
  conf.getRange('D8').setValue('担当者ログ・代理店全体ログで、次回アクション予定日のこの日数前から黄色表示');
  conf.getRange('B7:D7').clear();
  conf.getRange('B9:D9').clear();
  conf.getRange('B12:E19').clear();  // 使われなくなった代理店ランク定義
  const unused = ['list_ランク', 'list_接触目的', 'list_温度感', 'list_完了'];
  ss.getNamedRanges().filter(nr => unused.includes(nr.getName())).forEach(nr => nr.remove());
}

function writePlan_(sh) {
  sh.getRange('A1').setValue('定例アクション計画（代理店を動かし続けるための接点づくり）');
  sh.getRange('A2').setValue('いつ・どの代理店に・何をするかの運用ルール。対象の代理店は担当者ログ・『サマリー』の表示で判断します。');
  const plan = [
    ['毎週月曜', '自分の担当代理店',
     '自分の担当者ログで赤（次回予定日なし・期限切れ）→黄（期限間近）の順に連絡し、最終アクションと次回アクションを更新',
     '連絡の抜け漏れをゼロにする', '各担当', '担当者ログ'],
    ['毎週月曜', '全担当者',
     '『代理店全体ログ』上部のステータス表で、担当者ごとの要連絡件数・今月の接触数を確認', '担当者間の偏り・漏れを防ぐ', '責任者', '代理店全体ログ'],
    ['毎週', '紹介が多い代理店（『サマリー』で稼働率が高い）',
     '紹介予定・進捗の確認連絡', '紹介の流れを止めない', '各担当', '電話・LINE'],
    ['紹介を受けた当日中', '紹介元の代理店',
     '紹介受領のお礼と今後の流れを連絡し、自分の営業管理に会社を追加して「紹介元の代理店」「紹介日」を入力',
     '紹介して良かったと感じてもらう', '各担当', '電話・チャット'],
    ['初回面談から3営業日以内', '紹介元の代理店',
     '面談結果・次のステップを報告', '代理店が顧客にフォローしやすくする', '各担当', 'メール・チャット'],
    ['成約時（24時間以内）', '紹介元の代理店',
     '成約報告と紹介料の連絡。営業管理のその会社の行で進捗を「受注」にし、受注金額・紹介料を入力',
     '成功体験を共有し次の紹介につなげる', '各担当', '電話'],
    ['毎月初', '全代理店',
     '新商材・キャンペーン情報、紹介トークのポイントを配信', '紹介のきっかけを定期的に提供', 'マーケ担当', 'メール・チャット'],
    ['毎月末', '全代理店',
     '『サマリー』をもとに、代理店別の紹介・成約実績を報告', '実績を見える化して関係を深める', '各担当', 'メール'],
    ['毎月末', '稼働率0%の代理店（『サマリー』で灰色）',
     '再活性化の個別提案（勉強会招待、同行営業の提案など）', '休眠代理店を掘り起こす', '各担当', '電話・訪問'],
    ['四半期ごと', '全代理店',
     '代理店向け勉強会（税制改正・商材説明・成功事例共有）', '知識を更新し、紹介しやすい状態をつくる', '責任者', 'セミナー・オンライン'],
    ['半期ごと', '重点代理店（『サマリー』で紹介件数・受注が多い代理店）',
     '個別面談（目標のすり合わせ・インセンティブ案内）', '重点代理店のモチベーション維持', '責任者', '訪問・オンライン'],
    ['顧客の決算期が集中する時期の前', '全代理店',
     '節税・決算対策の提案キャンペーンを案内', '繁忙期前に紹介を集中させる', 'マーケ担当', 'メール・勉強会'],
    ['契約直後の1ヶ月', '新規代理店（担当者ログに追加した代理店）',
     '立上げ面談（商材説明・紹介フロー説明・紹介したい顧客像の洗い出し）', '最初の1件目の紹介を早く出す', '各担当', '訪問・オンライン'],
  ];
  sh.getRange(5, 1, 40, 6).clearContent();
  sh.getRange(5, 1, plan.length, 6).setValues(plan).setWrap(true).setVerticalAlignment('middle');
}

function writeHowto_(ss) {
  const how = ss.getSheetByName(SH.howto);
  how.getRange('A1').setValue(`${ss.getName()}　使い方（税務キーパー株式会社）`);
  how.getRange('A2').setValue('代理店とのコミュニケーションを切らさず、紹介を継続的に生み出すための管理シートです。');
  const lines = [
    ['■ シート構成', null],
    ['担当者ログ', `${ASSIGNEES.map(TAB).join('・')}。代理店1社につき1行。代理店とのやりとりを入力する。紹介件数〜紹介料未払は営業管理から自動。上部に本人のステータス。`],
    ['代理店全体ログ', '担当者ログを自動で合算した一覧（入力しない）。上部に担当者別・合計のステータス表。'],
    ['紹介・成約実績', '営業管理のうち「紹介元の代理店」が入った案件を自動で一覧化（入力しない）。'],
    ['サマリー', '代理店別の稼働率・紹介件数・受注・受注金額・紹介料を自動集計（入力しない）。'],
    ['定例アクション計画', 'いつ・どの代理店に・何をするかの運用ルール。'],
    ['営業管理（担当者ごと）', '顧客ごとに1行。代理店からの紹介案件もここに入れ、「紹介元の代理店」を選ぶ。紺色の列を入力し、緑色の列（決算まで・状況）は自動。案件IDはkintone『案件管理』へのリンクになる。上部に本人のステータス。'],
    ['営業管理全体ログ', '担当者ごとの営業管理を自動で合算（入力しない）。上部に担当者別・合計のステータス表。'],
    ['設定', '稼働率の対象期間（C6）、期限間近の日数（C8）、プルダウンの選択肢（接触手段＝H列、連絡ツール＝K列、商材＝M〜O列、営業管理＝Q〜S列）。'],
    ['', null],
    ['■ 色のルール', null],
    ['黄色の見出し', '入力する列。'],
    ['灰色の見出し', '自動の列・タブ（上書きしない）。担当者ログの紹介件数〜紹介料未払も自動。'],
    ['赤', '担当者ログ・代理店全体ログ：次回アクション予定日が未入力、または予定日を過ぎている。'],
    ['黄', '担当者ログ・代理店全体ログ：次回アクション予定日まで『設定』C8の日数以内。'],
    ['対応漏れの赤', '紹介案件のうち、営業管理の状況が「アポなし」「期限切れ」のものがある。'],
    ['紹介料未払のオレンジ', '紹介料が入っていて、紹介料支払が「支払済」になっていない。'],
    ['サマリーの灰色', '稼働率0%（対象期間に紹介がない）。'],
    ['サマリーの緑', '稼働率100%（対象期間の毎月に紹介がある）。'],
    ['', null],
    ['■ 運用の流れ', null],
    ['① 代理店を登録', '自分の担当者ログに1社1行で代理店ID・代理店名・代理店担当者を入力。代理店IDはkintone『8. 代理店管理』と同じIDを使う。'],
    ['② 接触したら更新', 'その代理店の行の「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。'],
    ['③ 紹介を受けたら', '自分の営業管理に会社を1行追加し、「紹介元の代理店」をプルダウンで選んで紹介日を入力。以降の進捗も営業管理で追う。'],
    ['④ 成約したら', '営業管理で進捗を「受注」にし、受注金額・紹介料を入力。紹介料を払ったら紹介料支払を「支払済」に。代理店の行・紹介・成約実績・サマリーに自動で反映される。'],
    ['⑤ 毎週月曜に確認', '担当者：自分のログで赤→黄の順に連絡し、対応漏れがあれば営業管理で次回アポを入れる。責任者：『代理店全体ログ』上部で担当者ごとの要連絡・対応漏れを確認。'],
    ['⑥ 月末に振り返り', '『サマリー』で稼働率0%（灰色）の代理店を洗い出し、『定例アクション計画』に沿って再活性化の手を打つ。'],
    ['担当替えのとき', 'その代理店の行を切り取り、相手の担当者ログに貼り付ける。'],
    ['', null],
    ['■ 自動計算の意味', null],
    ['担当代理店数', '担当者ログにある代理店名の数。'],
    ['稼働率', '直近の対象期間（初期値3ヶ月）のうち、紹介があった月の割合。3ヶ月中2ヶ月紹介あり→67%。'],
    ['紹介件数・進行中・受注', '営業管理で紹介元の代理店にその代理店名が入った案件の数（進行中＝受注・失注以外）。'],
    ['対応漏れ', 'そのうち、営業管理の状況が「アポなし」「期限切れ」の案件の数。'],
    ['紹介料未払', 'そのうち、紹介料支払が「支払済」でない案件の紹介料の合計。'],
    ['代理店IDのリンク', 'クリックするとkintone『8. 代理店管理』でその代理店に絞り込んだ一覧が開く（入力すると自動でリンク化）。'],
    ['', null],
    ['■ 営業管理の自動項目', null],
    ['決算まで', '決算月まであと何ヶ月か（今月／1ヶ月…）。2ヶ月以内はオレンジ。'],
    ['状況', '進捗が受注・失注ならそのまま。次回アポ日が空欄＝アポなし、過ぎている＝期限切れ（赤）、今週中＝今週アポ（黄）、それ以降＝予定あり。'],
    ['進行中・顧問見込み', '進行中＝受注・失注以外の顧客。顧問見込み＝そのうち単発／顧問が「顧問」。'],
    ['', null],
    ['■ 管理者向け', null],
    ['列を変えたとき', '拡張機能 → Apps Script で fixAgencySheet を実行すると、数式・プルダウン・色分けが整う。'],
    ['担当者が増えたとき', 'スクリプト上部の ASSIGNEES に名前を追加して fixAgencySheet と setupSales を実行（「名前：代理店」「名前：営業管理」のタブが自動で作られる）。'],
    ['商材を増やしたとき', '『設定』N・O列に追記し、スクリプト上部の ANKEN1_OPTIONS / ANKEN2_OPTIONS にも同じ項目を追加する。'],
  ];
  const colors = {
    '黄色の見出し': COLOR.input, '灰色の見出し': COLOR.auto, '赤': COLOR.red,
    '黄': COLOR.yellow, 'サマリーの灰色': COLOR.gray, 'サマリーの緑': COLOR.green,
    '対応漏れの赤': COLOR.red, '紹介料未払のオレンジ': '#FCE5CD',
  };
  how.getRange('B4:C70').clearContent().setBackground(null).setFontWeight('normal')
    .setFontColor('#000000').setFontSize(10);
  how.getRange(4, 2, lines.length, 2).setValues(lines.map(([k, v]) => [k, v || '']));
  lines.forEach(([k, v], i) => {
    const cell = how.getRange(4 + i, 2);
    cell.setFontWeight('bold');
    if (v === null) cell.setFontColor(COLOR.title).setFontSize(11);
    if (colors[k]) cell.setBackground(colors[k]);
  });
  how.getRange(4, 3, lines.length, 1).setWrap(true).setVerticalAlignment('middle');
  how.setColumnWidth(2, 190);
}

// ---------------------------------------------------------------- kintoneリンク
// 担当者ログの代理店IDを、kintoneへのリンクにする（既存データを一括変換）
function linkKintoneIds() {
  const ss = SpreadsheetApp.getActive();
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(TAB(name));
    if (sh && sh.getLastRow() >= P_FIRST) linkIds_(sh.getRange(P_FIRST, 1, sh.getLastRow() - P_FIRST + 1, 1));
  });
}

// 代理店IDを入力・貼り付けしたら自動でリンク化（シンプルトリガー）
function onEdit(e) {
  onEditSales_(e);  // 営業管理タブの案件IDのリンク化
  const sh = e.range.getSheet();
  if (!ASSIGNEES.map(TAB).includes(sh.getName()) || e.range.getColumn() !== 1) return;
  const top = Math.max(e.range.getRow(), P_FIRST);
  const bottom = e.range.getLastRow();
  if (bottom < top) return;
  linkIds_(sh.getRange(top, 1, bottom - top + 1, 1));
}

function linkIds_(range, toUrl) {
  const values = range.getDisplayValues();
  const rich = values.map(([v]) => {
    const text = String(v).trim();
    const url = (toUrl || kintoneUrl_)(text);
    const b = SpreadsheetApp.newRichTextValue().setText(text);
    return [url ? b.setLinkUrl(url).build() : b.build()];
  });
  range.setRichTextValues(rich);
}

function kintoneUrl_(id) {
  if (!id) return null;
  return KINTONE_APP_URL + '?query=' + encodeURIComponent(`${KINTONE_ID_FIELD} = "${id}"`);
}

// 数式内で代理店IDをkintoneリンクにする（expr は ID の列や配列）
function idLink_(expr) {
  return `IF(${expr}="","",HYPERLINK("${KINTONE_APP_URL}?query="&ENCODEURL("${KINTONE_ID_FIELD} = """&${expr}&""""),${expr}))`;
}

// 『設定』タブを消してしまったときの復旧用（基準日などの値と選択肢を作り直す）
function restoreSettings() {
  const ss = SpreadsheetApp.getActive();
  const conf = ss.getSheetByName(SH.conf) || ss.insertSheet(SH.conf);
  ensureSize_(conf, OPT_END, 19);
  conf.getRange('A1').setValue('設定・マスタ').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14);
  conf.getRange(1, 1, 1, 19).setBackground(COLOR.title);
  conf.getRange('A2').setValue('黄色セルを自社の運用に合わせて変更してください。他シートの判定・プルダウンはここを参照します。')
    .setFontColor('#595959').setFontSize(9);
  conf.getRange('B4:D8').setValues([
    ['基準日', '', '自動（今日の日付）'],
    ['今月初日', '', '自動'],
    ['稼働率の対象期間（ヶ月）', 3, '『サマリー』の稼働率：直近この月数のうち紹介があった月の割合'],
    ['', '', ''],
    ['期限間近の表示（日前）', 3, '担当者ログ・代理店全体ログで、次回アクション予定日のこの日数前から黄色表示'],
  ]);
  conf.getRange('C4').setFormula('=TODAY()');
  conf.getRange('C5').setFormula('=DATE(YEAR(C4),MONTH(C4),1)');
  conf.getRange('C4:C5').setNumberFormat('yyyy/mm/dd');
  conf.getRange('B4:B8').setFontWeight('bold');
  conf.getRange('C6').setBackground(COLOR.input);
  conf.getRange('C8').setBackground(COLOR.input);
  conf.getRange('D4:D8').setFontColor('#595959').setFontSize(9);
  conf.setColumnWidth(2, 200);
  const komon = conf.getRange(`M${OPT_FIRST}:M${OPT_END}`);
  if (komon.getValues().every(([v]) => v === '')) writeOptions_(komon, ['税務顧問', '節税顧問']);
  writeLists_(ss);
  writeSalesLists_(ss);
  SpreadsheetApp.getUi().alert('『設定』タブを作り直しました。続けて fixAgencySheet を実行してください。');
}

// ================================================================ 営業管理
// 担当者ごとの営業管理タブ（入力）と、営業管理全体ログ（自動合算）
// タブ名は「名前：営業管理」「営業管理全体ログ」。既に「営業」を含む同名系のタブがあればそれを使う
const SALES_TAB = name => `${name}：営業管理`;
// kintone『案件管理』アプリのURLと、案件IDのフィールドコード（案件IDで絞り込んだ一覧を開く）
const KINTONE_DEAL_URL = 'https://zeimukeeoer.cybozu.com/k/33/';
const KINTONE_DEAL_FIELD = '案件ID';
const SALES_LOG = '営業管理全体ログ';
// [見出し, 入力 or 自動, 列幅]
const S_COLS = [
  ['案件ID', 'in', 100], ['会社名', 'in', 200], ['氏名', 'in', 110], ['役職', 'in', 100],
  ['紹介元の代理店', 'in', 160], ['紹介日', 'in', 90], ['決算月', 'in', 65],
  ['決算まで', 'auto', 70], ['単発／顧問', 'in', 80], ['提案商材', 'in', 150], ['進捗', 'in', 110],
  ['最終接触日', 'in', 90], ['次回アポ日', 'in', 90], ['アポ方法', 'in', 90], ['状況', 'auto', 85],
  ['受注金額（円）', 'in', 110], ['紹介料（円）', 'in', 100], ['紹介料支払', 'in', 90],
  ['メモ', 'in', 320],
];
const SC = {  // 営業管理タブの列
  deal: 'A', company: 'B', position: 'D', agency: 'E', introDate: 'F', fiscal: 'G', untilFiscal: 'H', kind: 'I',
  product: 'J', progress: 'K', lastDate: 'L', nextDate: 'M', method: 'N', state: 'O', amount: 'P', fee: 'Q', paid: 'R',
};
const PAID_OPTIONS = ['未払', '支払済'];
const S_STATUS = ['進行中', '今週アポ', 'アポなし', '期限切れ', '決算間近', '顧問見込み', '受注'];
const POSITION_OPTIONS = ['代表取締役', '取締役', '部長', '担当者', 'その他'];
const PROGRESS_OPTIONS = ['アポ調整中', 'アポ確定', '提案中', '見積提出', 'クロージング', '受注', '失注', '保留'];
const FISCAL_NEAR = 2;  // 決算まで何ヶ月以内を「決算間近」とするか
const SCOLOR = { input: '#1F3864', auto: '#2B6F68', red: '#F4CCCC', redFont: '#990000', orange: '#FCE5CD' };

// 営業管理だけを作り直したいとき用
function setupSales() {
  setupSales_(SpreadsheetApp.getActive());
  SpreadsheetApp.getUi().alert('営業管理タブと営業管理全体ログを整えました。');
}

function setupSales_(ss) {
  writeSalesLists_(ss);
  const tabs = ASSIGNEES.map(name => [name, findSalesTab_(ss, name)]);
  tabs.forEach(([name, sh]) => setupSalesTab_(sh, name));
  buildSalesLog_(findSalesLog_(ss) || ss.insertSheet(SALES_LOG), tabs);
}

function findSalesLog_(ss) {
  return ss.getSheetByName(SALES_LOG)
    || ss.getSheets().find(sh => /営業/.test(sh.getName()) && /全体/.test(sh.getName()) && !/（/.test(sh.getName()));
}

// 営業管理全体ログの範囲（A列＝担当、B列以降＝営業管理の列）。head を省くと全列。代理店側の自動集計で使う
function salesRange_(ss, head) {
  const log = findSalesLog_(ss);
  const name = log ? log.getName() : SALES_LOG;
  const first = ASSIGNEES.length + 7, end = first + ASSIGNEES.length * P_ROWS - 1;
  if (!head) return `'${name}'!$A$${first}:$${colLetter_(S_COLS.length + 1)}$${end}`;
  const col = colLetter_(S_COLS.findIndex(([h]) => h === head) + 2);
  return `'${name}'!$${col}$${first}:$${col}$${end}`;
}

// 営業管理タブの末尾に案件を追加（同じ会社名が既にあれば追加しない）。追加した件数を返す
function appendDeals_(sh, deals) {
  const norm = v => String(v).replace(/\s/g, '');
  const companies = sh.getRange(`${SC.company}${P_FIRST}:${SC.company}${P_LAST}`).getValues().map(([v]) => norm(v));
  const add = deals.filter(d => !companies.includes(norm(d['会社名'])));
  if (!add.length) return 0;
  const lastUsed = companies.reduce((a, v, i) => v ? i : a, -1);
  const start = P_FIRST + lastUsed + 1;
  if (start + add.length - 1 > P_LAST) throw new Error(`『${sh.getName()}』に空き行が足りません。`);
  const rng = sh.getRange(start, 1, add.length, S_COLS.length);
  const formulas = rng.getFormulas();  // 自動の列（決算まで・状況）の数式はそのまま残す
  rng.setValues(add.map((d, r) => S_COLS.map(([h, kind], c) => kind === 'auto' ? formulas[r][c] : (d[h] ?? ''))));
  return add.length;
}

function findSalesTab_(ss, name) {
  const exact = ss.getSheetByName(SALES_TAB(name));
  if (exact) return exact;
  const surname = name.slice(0, 2);
  return ss.getSheets().find(sh => {
    const n = sh.getName();
    return /営業/.test(n) && !/全体|（/.test(n) && (n.includes(name) || n.includes(surname));
  }) || ss.insertSheet(SALES_TAB(name));
}

// 選択肢（『設定』Q〜S列）。手で編集した内容は残し、空のときだけ初期値を入れる
function writeSalesLists_(ss) {
  const conf = ss.getSheetByName(SH.conf) || ss.insertSheet(SH.conf);
  ensureSize_(conf, OPT_END, 19);
  if (conf.getRange(`H${OPT_FIRST}:H${OPT_END}`).getValues().every(([v]) => v === '')) {
    conf.getRange('H13').setValue('接触手段').setFontWeight('bold').setBackground(COLOR.input);
    writeOptions_(conf.getRange(`H${OPT_FIRST}:H${OPT_END}`), METHOD_OPTIONS);  // アポ方法の選択肢
  }
  conf.getRange('Q12').setValue('営業管理の選択肢').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange('Q13:S13').setValues([['役職', '進捗', '提案商材']])
    .setFontWeight('bold').setBackground(COLOR.input).setHorizontalAlignment('center');
  const keep = (col, items) => {
    const rng = conf.getRange(`${col}${OPT_FIRST}:${col}${OPT_END}`);
    if (rng.getValues().every(([v]) => v === '')) writeOptions_(rng, items);
  };
  keep('Q', POSITION_OPTIONS);
  keep('R', PROGRESS_OPTIONS);
  keep('S', ['未定'].concat(ANKEN1_OPTIONS, ANKEN2_OPTIONS));
  conf.setColumnWidths(17, 3, 150);
}

// ---------------------------------------------------------------- 営業管理（担当者ごと）
function setupSalesTab_(sh, name) {
  const nCol = S_COLS.length;
  // 列構成が変わっていたら、見出し名を手がかりに入力済みのデータを新しい列へ移す（元の内容は（組み替え前）に保存）
  ensureSize_(sh, P_LAST, nCol);
  remapColumns_(sh, S_COLS.map(([h]) => h));
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(P_FIRST, 1, P_ROWS, sh.getMaxColumns()).clearDataValidations();
  salesTitle_(sh, `${name}の営業管理`, nCol);

  // ステータス（3〜4行目）
  const r = col => `${col}${P_FIRST}:${col}${P_LAST}`;
  const open = `${r(SC.company)},"?*",${r(SC.progress)},"<>受注",${r(SC.progress)},"<>失注"`;
  sh.getRange(3, 1, 2, sh.getMaxColumns()).clear();
  sh.getRange(3, 1, 1, S_STATUS.length).setValues([S_STATUS]);
  sh.getRange(4, 1, 1, S_STATUS.length).setFormulas([[
    `=COUNTIFS(${open})`,
    `=COUNTIF(${r(SC.state)},"今週アポ")`,
    `=COUNTIF(${r(SC.state)},"アポなし")`,
    `=COUNTIF(${r(SC.state)},"期限切れ")`,
    `=COUNTIFS(${open},${r(SC.untilFiscal)},"<="&${FISCAL_NEAR})`,
    `=COUNTIFS(${open},${r(SC.kind)},"顧問")`,
    `=COUNTIF(${r(SC.progress)},"受注")`,
  ]]);
  salesStatusStyle_(sh, 3, 1, 1);

  // 見出し・本体
  salesHeader_(sh, P_HEAD, S_COLS);
  sh.getRange(P_FIRST, 1, P_ROWS, nCol).setNumberFormat('General').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  [SC.lastDate, SC.nextDate].forEach(col => sh.getRange(r(col)).setNumberFormat('m/d'));
  sh.getRange(r(SC.introDate)).setNumberFormat('yyyy/mm/dd');
  [SC.amount, SC.fee].forEach(col => sh.getRange(r(col)).setNumberFormat('#,##0'));
  sh.getRange(r(SC.paid)).setHorizontalAlignment('center');
  sh.getRange(r(SC.untilFiscal)).setNumberFormat('[=0]"今月";0"ヶ月"').setHorizontalAlignment('center');
  sh.getRange(r(SC.state)).setHorizontalAlignment('center').setFontWeight('bold');
  [SC.untilFiscal, SC.state].forEach(col => sh.getRange(r(col)).setBackground('#F3F3F3'));

  // 自動の列：決算まで（月数）・状況
  const f = [];
  const c = col => i => `${col}${i}`;
  const [co, fi, pr, nx] = [c(SC.company), c(SC.fiscal), c(SC.progress), c(SC.nextDate)];
  for (let i = P_FIRST; i <= P_LAST; i++) {
    f.push([`=IF(OR(${co(i)}="",${fi(i)}=""),"",MOD(VALUE(SUBSTITUTE(${fi(i)},"月",""))-MONTH(TODAY()),12))`]);
  }
  sh.getRange(r(SC.untilFiscal)).setFormulas(f);
  const g = [];
  for (let i = P_FIRST; i <= P_LAST; i++) {
    g.push([`=IF(${co(i)}="","",IF(${pr(i)}="受注","受注",IF(${pr(i)}="失注","失注",IF(${nx(i)}="","アポなし",` +
      `IF(${nx(i)}<TODAY(),"期限切れ",IF(${nx(i)}<=TODAY()-WEEKDAY(TODAY(),2)+7,"今週アポ","予定あり"))))))`]);
  }
  sh.getRange(r(SC.state)).setFormulas(g);

  // プルダウン
  const conf = sh.getParent().getSheetByName(SH.conf);
  const list = col => conf.getRange(`${col}${OPT_FIRST}:${col}${OPT_END}`);
  setListValidation_(sh.getRange(r(SC.position)), list('Q'));  // 役職
  setValueListValidation_(sh.getRange(r(SC.fiscal)), Array.from({ length: 12 }, (_, i) => `${i + 1}月`));
  setValueListValidation_(sh.getRange(r(SC.kind)), ['単発', '顧問']);
  setListValidation_(sh.getRange(r(SC.product)), list('S'));
  setListValidation_(sh.getRange(r(SC.progress)), list('R'));
  setListValidation_(sh.getRange(r(SC.method)), conf.getRange(`H${OPT_FIRST}:H${OPT_FIRST + 9}`));
  setValueListValidation_(sh.getRange(r(SC.paid)), PAID_OPTIONS);
  const agencyLog = sh.getParent().getSheetByName(SH.log);  // 紹介元の代理店＝担当者ログの代理店名
  if (agencyLog) setListValidation_(sh.getRange(r(SC.agency)), agencyLog.getRange(`B${LOG_FIRST}:B${LOG_END}`));

  salesColors_(sh, P_FIRST, P_LAST, SC.untilFiscal, SC.state, nCol, SC.fee, SC.paid);
  sh.setFrozenRows(P_HEAD);
  sh.setFrozenColumns(2);
  sh.getRange(P_HEAD, 1, P_ROWS + 1, nCol).createFilter();
  if (sh.getLastRow() >= P_FIRST) linkIds_(sh.getRange(P_FIRST, 1, sh.getLastRow() - P_FIRST + 1, 1), dealUrl_);
}

// 案件IDをkintone『案件管理』へのリンクにする
function dealUrl_(id) {
  if (!id) return null;
  return KINTONE_DEAL_URL + '?query=' + encodeURIComponent(`${KINTONE_DEAL_FIELD} = "${id}"`);
}

// 営業管理タブのA列（案件ID）を入力・貼り付けしたら自動でリンク化（onEdit から呼ばれる）
function onEditSales_(e) {
  const sh = e.range.getSheet();
  if (!ASSIGNEES.map(SALES_TAB).includes(sh.getName()) || e.range.getColumn() !== 1) return;
  const top = Math.max(e.range.getRow(), P_FIRST);
  if (e.range.getLastRow() < top) return;
  linkIds_(sh.getRange(top, 1, e.range.getLastRow() - top + 1, 1), dealUrl_);
}

// ---------------------------------------------------------------- 営業管理全体ログ（合算）
function buildSalesLog_(sh, tabs) {
  const cols = [['担当', 'auto', 80]].concat(S_COLS.map(([h, , w]) => [h, 'auto', w]));
  const nCol = cols.length;
  const n = tabs.length;
  const head = n + 6, first = head + 1, end = first + n * P_ROWS - 1;
  ensureSize_(sh, end, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns()).clear().clearDataValidations();
  salesTitle_(sh, '営業管理（全担当者の合算）', nCol,
    `${tabs.map(([, t]) => t.getName()).join('・')}を自動で合算（ここでは入力しない。修正は各担当者の営業管理タブで）`);

  // 担当者別ステータス表
  sh.getRange(3, 1, 1, S_STATUS.length + 1).setValues([['担当'].concat(S_STATUS)]);
  sh.getRange(4, 1, n, 1).setValues(tabs.map(([name]) => [name]));
  sh.getRange(4, 2, n, S_STATUS.length).setFormulas(
    tabs.map(([, t]) => S_STATUS.map((_, i) => `='${t.getName()}'!${colLetter_(i + 1)}4`)));
  sh.getRange(4 + n, 1).setValue('合計');
  sh.getRange(4 + n, 2, 1, S_STATUS.length).setFormulas([
    S_STATUS.map((_, i) => { const c = colLetter_(i + 2); return `=SUM(${c}4:${c}${3 + n})`; })]);
  salesStatusStyle_(sh, 3, 2, n);

  salesHeader_(sh, head, cols);
  const rest = S_COLS.map((_, i) => i + 1).join(',');
  const lets = tabs.map(([, t], i) => `p${i},'${t.getName()}'!A${P_FIRST}:${colLetter_(S_COLS.length)}${P_LAST}`).join(',');
  const company = S_COLS.findIndex(([h]) => h === '会社名') + 1;
  const parts = tabs.map(([name], i) =>
    `HSTACK(IF(CHOOSECOLS(p${i},${company})<>"","${name}",""),CHOOSECOLS(p${i},${rest}))`);
  const after = S_COLS.slice(1).map((_, i) => i + 3).join(',');
  const id = 'CHOOSECOLS(f,2)';
  sh.getRange(first, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(LET(${lets},all,VSTACK(${parts.join(',')}),f,FILTER(all,CHOOSECOLS(all,${company + 1})<>""),` +
    `HSTACK(CHOOSECOLS(f,1),IF(${id}="","",HYPERLINK("${KINTONE_DEAL_URL}?query="&ENCODEURL("${KINTONE_DEAL_FIELD} = """&${id}&""""),${id})),` +
    `CHOOSECOLS(f,${after})))),"")`);

  // 書式（営業管理タブの列が「担当」の分だけ右にずれる）
  const L = col => colLetter_(col.charCodeAt(0) - 64 + 1);
  const rr = col => sh.getRange(`${col}${first}:${col}${end}`);
  sh.getRange(first, 1, end - first + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  [SC.lastDate, SC.nextDate].forEach(col => rr(L(col)).setNumberFormat('m/d'));
  rr(L(SC.introDate)).setNumberFormat('yyyy/mm/dd');
  [SC.amount, SC.fee].forEach(col => rr(L(col)).setNumberFormat('#,##0'));
  rr(L(SC.paid)).setHorizontalAlignment('center');
  rr(L(SC.untilFiscal)).setNumberFormat('[=0]"今月";0"ヶ月"').setHorizontalAlignment('center');
  rr(L(SC.state)).setHorizontalAlignment('center').setFontWeight('bold');
  salesColors_(sh, first, end, L(SC.untilFiscal), L(SC.state), nCol, L(SC.fee), L(SC.paid));
  sh.setFrozenRows(head);
  sh.setFrozenColumns(3);
}

// ---------------------------------------------------------------- 営業管理の共通
function salesTitle_(sh, title, nCol, note) {
  sh.getRange(1, 1, 2, sh.getMaxColumns()).clear();
  sh.getRange('A1').setValue(title).setFontColor(SCOLOR.input).setFontWeight('bold').setFontSize(14);
  sh.getRange('D1').setFormula(
    `="基準日：" & TEXT(TODAY(),"yyyy/mm/dd") & "　／　${note || '紺色の列だけ入力（緑色の列は自動）'}"`)
    .setFontColor('#595959').setFontSize(9);
  sh.setRowHeight(1, 30);
}

function salesHeader_(sh, row, cols) {
  sh.getRange(row, 1, 1, cols.length)
    .setValues([cols.map(([h]) => h)])
    .setBackgrounds([cols.map(([, kind]) => kind === 'in' ? SCOLOR.input : SCOLOR.auto)])
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 36);
  cols.forEach(([, , w], i) => sh.setColumnWidth(i + 1, w));
}

// ステータス表：見出し行 headRow、数値は startCol 列から、行数 n（＋合計行）
function salesStatusStyle_(sh, headRow, startCol, n) {
  const width = S_STATUS.length + startCol - 1;
  const rows = startCol === 1 ? 1 : n + 1;
  sh.getRange(headRow, 1, 1, width).setBackground(SCOLOR.auto).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(headRow, 30);
  const vals = sh.getRange(headRow + 1, startCol, rows, S_STATUS.length);
  vals.setNumberFormat('0;-0;"-"').setHorizontalAlignment('center').setFontSize(startCol === 1 ? 14 : 10);
  if (startCol > 1) {
    sh.getRange(headRow + 1, 1, rows, 1).setFontWeight('bold');
    sh.getRange(headRow + rows, 1, 1, width).setBackground(COLOR.total).setFontWeight('bold');
  }
  sh.getRange(headRow, 1, rows + 1, width)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  // アポなし・期限切れ＝赤、決算間近＝オレンジ（0件のときは色なし）
  const col = label => colLetter_(S_STATUS.indexOf(label) + startCol);
  const rng = label => sh.getRange(`${col(label)}${headRow + 1}:${col(label)}${headRow + rows}`);
  const rules = sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(x => x.getRow() >= headRow && x.getRow() <= headRow + rows));
  ['アポなし', '期限切れ'].forEach(label => rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground(SCOLOR.red).setFontColor(SCOLOR.redFont).setBold(true)
    .setRanges([rng(label)]).build()));
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground(SCOLOR.orange).setFontColor(SCOLOR.redFont).setBold(true)
    .setRanges([rng('決算間近')]).build());
  sh.setConditionalFormatRules(rules);
}

// 決算まで（2ヶ月以内＝オレンジ）と状況（期限切れ・アポなし＝赤、今週アポ＝黄、受注＝緑、失注＝グレー）
// 紹介料支払（紹介料が入っていて支払済でない＝オレンジ）
function salesColors_(sh, first, last, fiscalCol, stateCol, nCol, feeCol, paidCol) {
  const fiscal = sh.getRange(`${fiscalCol}${first}:${fiscalCol}${last}`);
  const state = sh.getRange(`${stateCol}${first}:${stateCol}${last}`);
  const s = `$${stateCol}${first}`;
  const rules = sh.getConditionalFormatRules().filter(rule => rule.getRanges().every(x => x.getRow() < first));
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${fiscalCol}${first}<>"",$${fiscalCol}${first}<=${FISCAL_NEAR},${s}<>"受注",${s}<>"失注")`)
      .setBackground(SCOLOR.orange).setFontColor(SCOLOR.redFont).setBold(true).setRanges([fiscal]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('期限切れ')
      .setBackground(SCOLOR.red).setFontColor(SCOLOR.redFont).setRanges([state]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('アポなし')
      .setBackground(SCOLOR.red).setFontColor(SCOLOR.redFont).setRanges([state]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('今週アポ')
      .setBackground(COLOR.yellow).setRanges([state]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('受注')
      .setBackground(COLOR.green).setRanges([state]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${feeCol}${first}>0,$${paidCol}${first}<>"支払済")`)
      .setBackground(SCOLOR.orange).setFontColor(SCOLOR.redFont)
      .setRanges([sh.getRange(`${paidCol}${first}:${paidCol}${last}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=${s}="失注"`)
      .setFontColor('#A6A6A6').setRanges([sh.getRange(first, 1, last - first + 1, nCol)]).build());
  sh.setConditionalFormatRules(rules);
}

function setValueListValidation_(range, items) {
  range.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build());
}

// ---------------------------------------------------------------- 共通
function titleRow_(sh, title, nCol, note) {
  sh.getRange(1, 1, 1, nCol).setBackground(COLOR.title);
  sh.getRange('A1').setValue(title).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14);
  sh.setRowHeight(1, 32);
  sh.getRange('A2').setValue(note).setFontColor('#595959').setFontSize(9);
}

function headerRow_(sh, row, cols) {
  const rng = sh.getRange(row, 1, 1, cols.length);
  rng.setValues([cols.map(([h]) => h)])
    .setBackgrounds([cols.map(([, kind]) => kind === 'in' ? COLOR.input : COLOR.auto)])
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true).setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 36);
  cols.forEach(([, , w], i) => sh.setColumnWidth(i + 1, w));
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function writeOptions_(range, items) {
  const rows = range.getNumRows();
  const vals = [];
  for (let i = 0; i < rows; i++) vals.push([items[i] || '']);
  range.setValues(vals);
}

function setListValidation_(range, source) {
  range.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(source, true)
      .setAllowInvalid(true)
      .build());
}

function overlaps_(rule, target) {
  return rule.getRanges().some(r =>
    r.getSheet().getSheetId() === target.getSheet().getSheetId() &&
    r.getColumn() <= target.getLastColumn() && target.getColumn() <= r.getLastColumn() &&
    r.getRow() <= target.getLastRow() && target.getRow() <= r.getLastRow());
}

function findBrokenFormulas_(ss) {
  const out = [];
  ss.getSheets().forEach(sh => {
    if (/（(移行前|組み替え前|作成前|紹介移行前)）$/.test(sh.getName())) return;
    const rng = sh.getDataRange();
    const f = rng.getFormulas();
    for (let i = 0; i < f.length; i++) {
      for (let j = 0; j < f[i].length; j++) {
        if (f[i][j] && /#REF!|代理店一覧/.test(f[i][j])) {
          out.push(`${sh.getName()}!${rng.getCell(i + 1, j + 1).getA1Notation()}`);
        }
      }
    }
  });
  return out;
}
