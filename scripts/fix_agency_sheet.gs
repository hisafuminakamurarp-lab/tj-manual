/**
 * 代理店進捗管理シート：セットアップ・連動スクリプト（Google Apps Script）
 *
 *   - 担当者ログ（前田紘孝ログ・浅沼雄太ログ）… 入力はここだけ。代理店とのやりとり＋紹介いただいた企業の情報
 *   - 代理店全体ログ  … 担当者ログを自動で合算。上部に担当者別・合計のステータス
 *   - 紹介・成約実績  … 担当者ログのうち紹介社名が入った行を自動で一覧化
 *   - サマリー        … 代理店別集計（稼働率・紹介数・成約数・金額）
 *   - 代理店ID        … kintone『8. 代理店管理』へのリンク（入力すると自動でリンク化）
 *
 * 使い方：拡張機能 → Apps Script に貼り付け → fixAgencySheet を実行。
 * 何度実行しても同じ状態になる（入力済みのデータは消さない）。
 * 『使い方』『定例アクション計画』『設定』の説明を作り直すときだけ updateGuides を実行。
 */

// 自社の担当者。増えたらここに追加して fixAgencySheet を実行（「名前＋ログ」のタブが自動で作られる）
const ASSIGNEES = ['前田紘孝', '浅沼雄太'];
const TAB = name => `${name}ログ`;

const SH = {
  log: '代理店全体ログ',
  kpi: '紹介・成約実績',
  sum: 'サマリー',
  plan: '定例アクション計画',
  howto: '使い方',
  conf: '設定',
};

// 担当者ログ：3〜4行目＝ステータス、6行目＝見出し、7行目〜＝データ（1行＝代理店とのやりとり or 紹介1件）
const P_HEAD = 6, P_FIRST = 7, P_ROWS = 300;
const P_LAST = P_FIRST + P_ROWS - 1;
// [見出し, 列幅]
const P_COLS = [
  ['代理店ID', 100], ['代理店名', 200], ['代理店担当者', 110],
  ['連絡ツール', 110], ['接触手段', 110], ['最終アクション日', 100], ['最終アクション内容', 260],
  ['次回アクション予定日', 110], ['次回アクション内容', 260],
  ['紹介社名', 180], ['紹介日', 95], ['成約日', 95],
  ['成約商材（顧問種別）', 120], ['成約商材（案件種別①）', 130], ['成約商材（案件種別②）', 150],
  ['成約金額（円）', 110], ['代理店報酬（円）', 110], ['備考', 220],
];
const PC = {  // 担当者ログの列
  id: 'A', tool: 'D', method: 'E', lastDate: 'F', lastText: 'G', nextDate: 'H', nextText: 'I',
  company: 'J', intro: 'K', close: 'L', komon: 'M', anken1: 'N', anken2: 'O', amount: 'P', reward: 'Q',
};
const TOOL_OPTIONS = ['グループLINE', '個別LINE', 'Chatwork', 'メール', '電話'];
const METHOD_OPTIONS = ['電話', '訪問', 'オンライン面談', 'メール', 'LINE/チャット', '勉強会・セミナー', '会食'];
const STATUS_LABELS = ['担当代理店数', '要連絡（赤）', '期限間近（黄）', '今月の接触', '紹介数', '成約数', '成約金額（円）'];

// 代理店全体ログ：3行目〜＝担当者別ステータス表、その2行下が見出し、次の行から合算データ
const LOG_HEAD = ASSIGNEES.length + 6;
const LOG_FIRST = LOG_HEAD + 1;
const LOG_END = LOG_FIRST + ASSIGNEES.length * P_ROWS - 1;
const LOG_COLS = P_COLS.slice(0, 2).concat([['自社担当', 90]], P_COLS.slice(2));  // 自社担当をC列に挿入

// 紹介・成約実績：4行目＝見出し、5行目〜＝担当者ログから自動（紹介社名が入った行だけ）
const KPI_HEADS = ['代理店ID', '代理店名', '自社担当', '紹介社名', '紹介日', '成約日',
  '成約商材（顧問種別）', '成約商材（案件種別①）', '成約商材（案件種別②）', '成約金額（円）', '代理店報酬（円）', '備考'];
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
    const plain = ss.getSheetByName(name);  // 以前の「名前だけ」のタブがあれば改名して使う
    if (plain) plain.setName(TAB(name)); else ss.insertSheet(TAB(name), 0);
  });
  if (!migrateOldLog_(ss)) return;  // 旧形式のデータ移行（担当者未入力なら中断）

  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(TAB(name));
    remapPersonColumns_(sh);
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
    // 見出しを新しい形にしてから書き込む（remapPersonColumns_ で組み替え対象にしない）
    sh.getRange(P_HEAD, 1, 1, P_COLS.length).setValues([P_COLS.map(([h]) => h)]);
    const filled = sh.getRange(P_FIRST, 1, P_ROWS, 1).getValues().filter(([v]) => v !== '').length;
    sh.getRange(P_FIRST + filled, 1, mine.length, mine[0].length).setValues(mine);
  });

  if (log.getFilter()) log.getFilter().remove();
  log.setConditionalFormatRules([]);
  log.getRange(3, 1, log.getMaxRows() - 2, log.getMaxColumns()).clear().clearDataValidations();
  return true;
}

// 担当者ログの見出しが今の列構成と違うとき、見出し名を手がかりにデータを新しい列へ並べ替える
// （使わなくなった列・自動計算の列は捨てる。「メモ」は「備考」へ移す）
function remapPersonColumns_(sh) {
  const want = P_COLS.map(([h]) => h);
  const width = Math.max(sh.getLastColumn(), want.length);
  const head = sh.getRange(P_HEAD, 1, 1, width).getDisplayValues()[0].map(v => v.replace(/\s/g, ''));
  if (want.every((h, i) => head[i] === h.replace(/\s/g, ''))) return;
  if (head.every(v => v === '')) return;  // 新規タブ

  const rng = sh.getRange(P_FIRST, 1, P_ROWS, width);
  const values = rng.getValues();
  const formulas = rng.getFormulas();
  const alias = { '備考': ['備考', 'メモ'] };
  const src = want.map(h => {
    const names = (alias[h] || [h]).map(n => n.replace(/\s/g, ''));
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
}

// ---------------------------------------------------------------- 担当者ログ
function setupPersonTab_(sh, name) {
  const nCol = P_COLS.length;
  ensureSize_(sh, P_LAST, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(P_FIRST, 1, P_ROWS, sh.getMaxColumns()).clearDataValidations();
  titleRow_(sh, `${name}の担当代理店`, nCol,
    '代理店とのやりとりを1行ずつ記録し、紹介をいただいたら紹介社名から右に入力します（同じ代理店から2件目の紹介は、同じ代理店IDで行を追加）。' +
    `接触したら「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。内容は『${SH.log}』『${SH.kpi}』に自動で反映されます。`);

  // ステータス（3〜4行目）
  const r = col => `${col}${P_FIRST}:${col}${P_LAST}`;
  sh.getRange(3, 1, 1, sh.getMaxColumns()).clear();
  sh.getRange(4, 1, 1, sh.getMaxColumns()).clear();
  sh.getRange(3, 1, 1, STATUS_LABELS.length).setValues([STATUS_LABELS])
    .setBackground(COLOR.auto).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, STATUS_LABELS.length).setFormulas([[
    `=IFERROR(ROWS(UNIQUE(FILTER(${r(PC.id)},${r(PC.id)}<>""))),0)`,  // 空のときは0（COUNTUNIQUEだとエラーを1件と数える）
    `=COUNTIFS(${r(PC.id)},"?*",${r(PC.nextDate)},"")+COUNTIFS(${r(PC.id)},"?*",${r(PC.nextDate)},"<"&TODAY())`,
    `=COUNTIFS(${r(PC.id)},"?*",${r(PC.nextDate)},">="&TODAY(),${r(PC.nextDate)},"<="&TODAY()+設定!$C$8)`,
    `=COUNTIFS(${r(PC.id)},"?*",${r(PC.lastDate)},">="&設定!$C$5)`,
    `=COUNTIF(${r(PC.company)},"?*")`,
    `=COUNT(${r(PC.close)})`,
    `=SUM(${r(PC.amount)})`,
  ]]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange(3, 1, 2, STATUS_LABELS.length)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B4').setFontColor(COLOR.redFont);

  headerRow_(sh, P_HEAD, P_COLS.map(([h, w]) => [h, 'in', w]));
  // 紹介に関する列は見出しを少し濃くして、代理店とのやりとり部分と区別する
  sh.getRange(`${PC.company}${P_HEAD}:R${P_HEAD}`).setBackground('#FCE4D6');
  const body = sh.getRange(P_FIRST, 1, P_ROWS, nCol);
  body.setNumberFormat('General').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  [PC.lastDate, PC.nextDate, PC.intro, PC.close].forEach(col => sh.getRange(r(col)).setNumberFormat('yyyy/mm/dd'));
  [PC.amount, PC.reward].forEach(col => sh.getRange(r(col)).setNumberFormat('#,##0'));
  [PC.lastText, PC.nextText].forEach(col => sh.getRange(r(col)).setWrap(true));

  // プルダウン（選択肢は『設定』）
  const conf = sh.getParent().getSheetByName(SH.conf);
  const list = (col, end) => conf.getRange(`${col}${OPT_FIRST}:${col}${end || OPT_FIRST + 9}`);
  setListValidation_(sh.getRange(r(PC.tool)), list('K'));
  setListValidation_(sh.getRange(r(PC.method)), list('H'));
  setListValidation_(sh.getRange(r(PC.komon)), list('M', OPT_END));
  setListValidation_(sh.getRange(r(PC.anken1)), list('N', OPT_END));
  setListValidation_(sh.getRange(r(PC.anken2)), list('O', OPT_END));

  // 色分け：次回アクション予定日が未入力・期限切れ＝D〜I列を赤、期限間近＝H〜I列を黄
  const f = P_FIRST;
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(COUNTA($A${f}:$R${f})>0,OR($H${f}="",$H${f}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(`D${f}:I${P_LAST}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($H${f}<>"",$H${f}>=TODAY(),$H${f}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow).setRanges([sh.getRange(`H${f}:I${P_LAST}`)]).build(),
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
  sh.getRange(4, 3, n + 1, 1).setFontColor(COLOR.redFont).setFontWeight('bold');
  sh.getRange(3, 1, n + 2, STATUS_LABELS.length + 1)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);

  headerRow_(sh, LOG_HEAD, LOG_COLS.map(([h, w]) => [h, 'auto', w]));

  // 担当者ログを縦に合算（自社担当の名前をC列に入れる）。代理店IDはkintoneリンク付き
  const rest = P_COLS.slice(2).map((_, i) => i + 3).join(',');
  const parts = ASSIGNEES.map((name, i) =>
    `HSTACK(CHOOSECOLS(p${i},1,2),IF(CHOOSECOLS(p${i},1)<>"","${name}",""),CHOOSECOLS(p${i},${rest}))`);
  const lets = ASSIGNEES.map((name, i) =>
    `p${i},'${TAB(name)}'!A${P_FIRST}:${colLetter_(P_COLS.length)}${P_LAST}`).join(',');
  const others = LOG_COLS.slice(1).map((_, i) => i + 2).join(',');
  sh.getRange(LOG_FIRST, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(LET(${lets},all,VSTACK(${parts.join(',')}),f,FILTER(all,CHOOSECOLS(all,1)<>""),` +
    `HSTACK(${idLink_('CHOOSECOLS(f,1)')},CHOOSECOLS(f,${others})))),"")`);

  // 書式（担当者ログの列が自社担当の分だけ右にずれる）
  const L = h => colLetter_(LOG_COLS.findIndex(([x]) => x === h) + 1);
  const rr = col => sh.getRange(`${col}${LOG_FIRST}:${col}${LOG_END}`);
  sh.getRange(LOG_FIRST, 1, LOG_END - LOG_FIRST + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['最終アクション日', '次回アクション予定日', '紹介日', '成約日'].forEach(h => rr(L(h)).setNumberFormat('yyyy/mm/dd'));
  ['成約金額（円）', '代理店報酬（円）'].forEach(h => rr(L(h)).setNumberFormat('#,##0'));
  ['最終アクション内容', '次回アクション内容'].forEach(h => rr(L(h)).setWrap(true));
  sh.getRange(`C${LOG_FIRST}:C${LOG_END}`).setHorizontalAlignment('center');

  // 色分け（担当者ログと同じ基準）
  const f = LOG_FIRST, next = L('次回アクション予定日');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${f}<>"",OR($${next}${f}="",$${next}${f}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(`${L('連絡ツール')}${f}:${L('次回アクション内容')}${LOG_END}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($${next}${f}<>"",$${next}${f}>=TODAY(),$${next}${f}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow)
      .setRanges([sh.getRange(`${next}${f}:${L('次回アクション内容')}${LOG_END}`)]).build(),
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
    notes.push(`※『${SH.kpi}』に手入力の紹介データがあったため『${backup}』に退避しました。各担当者ログの紹介社名〜備考の列に入力し直してください。`);
  }

  const nCol = KPI_HEADS.length;
  ensureSize_(sh, KPI_END, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.setConditionalFormatRules([]);
  sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns()).clear().clearDataValidations();
  titleRow_(sh, `${SH.kpi}（紹介1件につき1行・自動）`, nCol,
    '担当者ログで紹介社名を入力した行が、ここに自動で並びます（ここでは入力しません。修正は各担当者ログで）。');

  const widths = { '代理店ID': 100, '代理店名': 200, '自社担当': 90, '紹介社名': 180, '備考': 220 };
  headerRow_(sh, 4, KPI_HEADS.map(h => [h, 'auto', widths[h] || 115]));

  const src = KPI_HEADS.map(h => LOG_COLS.findIndex(([x]) => x === h) + 1);
  const company = LOG_COLS.findIndex(([x]) => x === '紹介社名') + 1;
  const range = `'${SH.log}'!A${LOG_FIRST}:${colLetter_(LOG_COLS.length)}${LOG_END}`;
  const rest = KPI_HEADS.slice(1).map((_, i) => i + 2).join(',');
  sh.getRange(KPI_FIRST, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(LET(d,${range},f,FILTER(CHOOSECOLS(d,${src.join(',')}),CHOOSECOLS(d,${company})<>""),` +
    `HSTACK(${idLink_('CHOOSECOLS(f,1)')},CHOOSECOLS(f,${rest})))),"")`);

  const K = h => colLetter_(KPI_HEADS.indexOf(h) + 1);
  const rr = col => sh.getRange(`${col}${KPI_FIRST}:${col}${KPI_END}`);
  sh.getRange(KPI_FIRST, 1, KPI_END - KPI_FIRST + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['紹介日', '成約日'].forEach(h => rr(K(h)).setNumberFormat('yyyy/mm/dd'));
  ['成約金額（円）', '代理店報酬（円）'].forEach(h => rr(K(h)).setNumberFormat('#,##0'));
  rr(K('自社担当')).setHorizontalAlignment('center');
  sh.setFrozenRows(4);
  sh.setFrozenColumns(2);
  return notes;
}

// 紹介・成約実績の列を参照する範囲（サマリー用）
function kpiRange_(head) {
  const col = colLetter_(KPI_HEADS.indexOf(head) + 1);
  return `'${SH.kpi}'!$${col}$${KPI_FIRST}:$${col}$${KPI_END}`;
}

// ---------------------------------------------------------------- 『設定』の選択肢
function writeLists_(ss) {
  const notes = [];
  const conf = ss.getSheetByName(SH.conf);
  const last = OPT_FIRST + 9;
  conf.getRange('G12').setValue('入力リスト（担当者ログのプルダウン）').setFontWeight('bold').setFontColor(COLOR.title);
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
  conf.getRange('M12').setValue('成約商材の選択肢（kintoneと同じ内容）').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange('M13:O13').setValues([['顧問種別', '案件種別①', '案件種別②']])
    .setFontWeight('bold').setBackground(COLOR.input).setHorizontalAlignment('center');
  writeOptions_(conf.getRange(`N${OPT_FIRST}:N${OPT_END}`), ANKEN1_OPTIONS);
  writeOptions_(conf.getRange(`O${OPT_FIRST}:O${OPT_END}`), ANKEN2_OPTIONS);
  conf.setColumnWidths(13, 3, 180);
  if (conf.getRange(`M${OPT_FIRST}:M${OPT_END}`).getValues().every(([v]) => v === '')) {
    notes.push('※ 顧問種別の選択肢が未登録です。『設定』シートM14から下にkintoneの顧問種別の項目を入力してください。');
  }
  return notes;
}

// ---------------------------------------------------------------- サマリー
function fixSummary_(sh) {
  sh.getCharts().forEach(ch => sh.removeChart(ch));
  const lastRow = Math.max(sh.getMaxRows(), SUM_END);
  sh.getRange(5, 1, lastRow - 4, 7).clearContent().setBackground(null).setFontWeight('normal');

  sh.getRange('A1').setValue('サマリー（代理店別）');
  sh.getRange('A2').setValue(
    `すべて自動計算。代理店IDは担当者ログ（${SH.log}）から自動で並びます。` +
    '稼働率＝直近の対象期間（『設定』C6、初期値3ヶ月）のうち紹介があった月の割合。');
  sh.getRange('A4:G4').setValues([['代理店ID', '代理店名', '稼働率', '紹介数', '成約数', '成約金額（円）', '代理店報酬（円）']]);

  const L = `'${SH.log}'`;
  const logId = `${L}!$A$${LOG_FIRST}:$A$${LOG_END}`, logName = `${L}!$B$${LOG_FIRST}:$B$${LOG_END}`;
  const kId = kpiRange_('代理店ID'), kIntro = kpiRange_('紹介日'), kClose = kpiRange_('成約日');
  const kAmount = kpiRange_('成約金額（円）'), kReward = kpiRange_('代理店報酬（円）');

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

  // 代理店ID：代理店全体ログから重複なしで自動展開
  sh.getRange(`A${SUM_FIRST}`).setFormula(
    `=IFERROR(LET(ids,UNIQUE(FILTER(${logId},${logId}<>"")),ARRAYFORMULA(IFERROR(${idLink_('ids')},ids))),"")`);

  const rows = [];
  for (let r = SUM_FIRST; r <= SUM_END; r++) {
    const A = `$A${r}`;
    const m = 'DATE(YEAR(設定!$C$5),MONTH(設定!$C$5)';
    const seq = 'SEQUENCE(設定!$C$6)';
    rows.push([
      `=IF(${A}="","",IFERROR(INDEX(${logName},MATCH(${A},${logId},0)),""))`,
      `=IF(${A}="","",ARRAYFORMULA(SUMPRODUCT(--(COUNTIFS(${kId},${A},` +
        `${kIntro},">="&${m}+1-${seq},1),${kIntro},"<"&${m}+2-${seq},1))>0)))/設定!$C$6)`,
      `=IF(${A}="","",COUNTIFS(${kId},${A}))`,
      `=IF(${A}="","",COUNTIFS(${kId},${A},${kClose},"<>"))`,
      `=IF(${A}="","",SUMIFS(${kAmount},${kId},${A}))`,
      `=IF(${A}="","",SUMIFS(${kReward},${kId},${A}))`,
    ]);
  }
  sh.getRange(SUM_FIRST, 2, rows.length, 6).setFormulas(rows);

  sh.getRange(5, 1, SUM_END - 4, 2).setHorizontalAlignment('center');
  sh.getRange(5, 3, SUM_END - 4, 1).setNumberFormat('0%');
  sh.getRange(5, 4, SUM_END - 4, 4).setNumberFormat('#,##0');
  sh.getRange(5, 1, SUM_END - 4, 7)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
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
     '紹介受領のお礼と今後の流れを連絡し、担当者ログに紹介社名・紹介日を入力（2件目以降は同じ代理店IDで行を追加）',
     '紹介して良かったと感じてもらう', '各担当', '電話・チャット'],
    ['初回面談から3営業日以内', '紹介元の代理店',
     '面談結果・次のステップを報告', '代理店が顧客にフォローしやすくする', '各担当', 'メール・チャット'],
    ['成約時（24時間以内）', '紹介元の代理店',
     '成約報告と報酬見込みの連絡。担当者ログのその紹介の行に成約日・成約商材・成約金額・代理店報酬を入力',
     '成功体験を共有し次の紹介につなげる', '各担当', '電話'],
    ['毎月初', '全代理店',
     '新商材・キャンペーン情報、紹介トークのポイントを配信', '紹介のきっかけを定期的に提供', 'マーケ担当', 'メール・チャット'],
    ['毎月末', '全代理店',
     '『サマリー』をもとに、代理店別の紹介・成約実績を報告', '実績を見える化して関係を深める', '各担当', 'メール'],
    ['毎月末', '稼働率0%の代理店（『サマリー』で灰色）',
     '再活性化の個別提案（勉強会招待、同行営業の提案など）', '休眠代理店を掘り起こす', '各担当', '電話・訪問'],
    ['四半期ごと', '全代理店',
     '代理店向け勉強会（税制改正・商材説明・成功事例共有）', '知識を更新し、紹介しやすい状態をつくる', '責任者', 'セミナー・オンライン'],
    ['半期ごと', '重点代理店（『サマリー』で紹介数・成約数が多い代理店）',
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
    ['担当者ログ', `${ASSIGNEES.map(TAB).join('・')}。入力はここだけ。代理店とのやりとりと、紹介いただいた企業の情報を入力する。上部に本人のステータス。`],
    ['代理店全体ログ', '担当者ログを自動で合算した一覧（入力しない）。上部に担当者別・合計のステータス表。'],
    ['紹介・成約実績', '担当者ログのうち紹介社名が入った行を自動で一覧化（入力しない）。'],
    ['サマリー', '代理店別の稼働率・紹介数・成約数・成約金額・代理店報酬を自動集計（入力しない）。'],
    ['定例アクション計画', 'いつ・どの代理店に・何をするかの運用ルール。'],
    ['設定', '稼働率の対象期間（C6）、期限間近の日数（C8）、プルダウンの選択肢（接触手段＝H列、連絡ツール＝K列、成約商材＝M〜O列）。'],
    ['', null],
    ['■ 色のルール', null],
    ['黄色の見出し', '入力する列（オレンジの見出しは紹介に関する列）。'],
    ['灰色の見出し', '自動の列・タブ（上書きしない）。'],
    ['赤', '担当者ログ・代理店全体ログ：次回アクション予定日が未入力、または予定日を過ぎている。'],
    ['黄', '担当者ログ・代理店全体ログ：次回アクション予定日まで『設定』C8の日数以内。'],
    ['サマリーの灰色', '稼働率0%（対象期間に紹介がない）。'],
    ['サマリーの緑', '稼働率100%（対象期間の毎月に紹介がある）。'],
    ['', null],
    ['■ 運用の流れ', null],
    ['① 代理店を登録', '自分の担当者ログに代理店ID・代理店名・代理店担当者を入力。代理店IDはkintone『8. 代理店管理』と同じIDを使う。'],
    ['② 接触したら更新', 'その代理店の行の「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。'],
    ['③ 紹介を受けたら', '同じ行の紹介社名・紹介日を入力（紹介欄が埋まっている代理店の2件目以降は、同じ代理店IDで行を追加）。'],
    ['④ 成約したら', 'その紹介の行に成約日・成約商材・成約金額・代理店報酬を入力。紹介・成約実績とサマリーに自動で反映される。'],
    ['⑤ 毎週月曜に確認', '担当者：自分のログで赤→黄の順に連絡。責任者：『代理店全体ログ』上部で担当者ごとの要連絡件数を確認。'],
    ['⑥ 月末に振り返り', '『サマリー』で稼働率0%（灰色）の代理店を洗い出し、『定例アクション計画』に沿って再活性化の手を打つ。'],
    ['担当替えのとき', 'その代理店の行を切り取り、相手の担当者ログに貼り付ける。'],
    ['', null],
    ['■ 自動計算の意味', null],
    ['担当代理店数', '担当者ログにある代理店IDの数（同じ代理店の行が複数あっても1社）。'],
    ['稼働率', '直近の対象期間（初期値3ヶ月）のうち、紹介があった月の割合。3ヶ月中2ヶ月紹介あり→67%。'],
    ['紹介数・成約数', '紹介社名が入った行の数と、そのうち成約日が入った行の数。'],
    ['代理店IDのリンク', 'クリックするとkintone『8. 代理店管理』でその代理店に絞り込んだ一覧が開く（入力すると自動でリンク化）。'],
    ['', null],
    ['■ 管理者向け', null],
    ['列を変えたとき', '拡張機能 → Apps Script で fixAgencySheet を実行すると、数式・プルダウン・色分けが整う。'],
    ['担当者が増えたとき', 'スクリプト上部の ASSIGNEES に名前を追加して fixAgencySheet を実行（「名前＋ログ」のタブが自動で作られる）。'],
    ['商材を増やしたとき', '『設定』N・O列に追記し、スクリプト上部の ANKEN1_OPTIONS / ANKEN2_OPTIONS にも同じ項目を追加する。'],
  ];
  const colors = {
    '黄色の見出し': COLOR.input, '灰色の見出し': COLOR.auto, '赤': COLOR.red,
    '黄': COLOR.yellow, 'サマリーの灰色': COLOR.gray, 'サマリーの緑': COLOR.green,
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
  const sh = e.range.getSheet();
  if (!ASSIGNEES.map(TAB).includes(sh.getName()) || e.range.getColumn() !== 1) return;
  const top = Math.max(e.range.getRow(), P_FIRST);
  const bottom = e.range.getLastRow();
  if (bottom < top) return;
  linkIds_(sh.getRange(top, 1, bottom - top + 1, 1));
}

function linkIds_(range) {
  const values = range.getDisplayValues();
  const rich = values.map(([v]) => {
    const text = String(v).trim();
    const url = kintoneUrl_(text);
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
  return `HYPERLINK("${KINTONE_APP_URL}?query="&ENCODEURL("${KINTONE_ID_FIELD} = """&${expr}&""""),${expr})`;
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
    if (/（(移行前|組み替え前)）$/.test(sh.getName())) return;
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
