/**
 * 代理店進捗管理シート：セットアップ・連動スクリプト（Google Apps Script）
 *
 *   - 担当者タブ（前田紘孝・浅沼雄太）… 担当代理店を入力する場所。上部に本人のステータス
 *   - アクションログ  … 担当者タブを自動で合算（入力不要）。上部に担当者別・合計のステータス
 *   - 紹介・成約実績  … 代理店名を自動表示、成約商材はkintoneと同じ選択肢のプルダウン
 *   - サマリー        … 代理店別集計（IDはアクションログから自動抽出）
 *   - 代理店ID        … kintone『8. 代理店管理』へのリンク（入力すると自動でリンク化）
 *
 * 使い方：拡張機能 → Apps Script に貼り付け → fixAgencySheet を実行。
 * 何度実行しても同じ状態になる（入力済みのデータは消さない）。
 * 『使い方』『定例アクション計画』『設定』の説明を作り直すときだけ updateGuides を実行。
 */

// 自社の担当者。増えたらここに追加して fixAgencySheet を実行（同名のタブが自動で作られる）
const ASSIGNEES = ['前田紘孝', '浅沼雄太'];

const SH = {
  log: 'アクションログ',
  kpi: '紹介・成約実績',
  sum: 'サマリー',
  plan: '定例アクション計画',
  howto: '使い方',
  conf: '設定',
};

// 担当者タブ：3〜4行目＝ステータス、6行目＝見出し、7行目〜＝代理店
const P_HEAD = 6, P_FIRST = 7, P_ROWS = 300;
const P_LAST = P_FIRST + P_ROWS - 1;
// [見出し, 入力 or 自動, 列幅]（kintone『案件管理』の項目を代理店管理向けに抽出）
const P_COLS = [
  ['代理店ID', 'in', 100], ['代理店名', 'in', 200], ['代理店担当者', 'in', 110],
  ['代理店副担当者', 'in', 110], ['自社副担当', 'in', 100], ['代理店ステータス', 'in', 110],
  ['紹介見込み', 'in', 90], ['初回面談日', 'in', 95], ['連絡ツール', 'in', 110],
  ['接触手段', 'in', 110], ['最終アクション日', 'in', 100], ['最終アクション内容', 'in', 280],
  ['次回アクション予定日', 'in', 110], ['次回アクション内容', 'in', 280],
  ['紹介数', 'auto', 70], ['成約数', 'auto', 70], ['成約金額（円）', 'auto', 110], ['メモ', 'in', 240],
];
const PC = {  // 担当者タブの列
  id: 'A', name: 'B', agent: 'C', status: 'F', lastDate: 'K', nextDate: 'M', nextText: 'N',
  intro: 'O', close: 'P', amount: 'Q',
};
const STATUS_OPTIONS = ['立上げ中', '稼働中', 'フォロー強化', '休眠', '契約終了'];
const PROSPECT_OPTIONS = ['高', '中', '低', 'なし'];
const TOOL_OPTIONS = ['グループLINE', '個別LINE', 'Chatwork', 'メール', '電話'];
const STATUS_LABELS = ['担当代理店数', '要連絡（赤）', '期限間近（黄）', '今月の接触', '紹介数', '成約数', '成約金額（円）'];

// アクションログ：3行目〜＝担当者別ステータス表、その2行下が見出し、次の行から合算データ
const LOG_HEAD = ASSIGNEES.length + 6;
const LOG_FIRST = LOG_HEAD + 1;
const LOG_END = LOG_FIRST + ASSIGNEES.length * P_ROWS - 1;
const LOG_COLS = P_COLS.slice(0, 2).concat([['自社担当', 'auto', 90]], P_COLS.slice(2));  // 自社担当をC列に挿入

const KPI_END = 3003;   // 紹介・成約実績の最終行
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

  writeLists_(ss);
  ASSIGNEES.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name, 0);
  });
  if (!migrateOldLog_(ss)) return;  // 旧アクションログのデータ移行（担当者未入力なら中断）

  const c = kpiCols_(ss.getSheetByName(SH.kpi));
  ASSIGNEES.forEach(name => setupPersonTab_(ss.getSheetByName(name), c));
  buildLog_(ss.getSheetByName(SH.log));
  const notes = fixKpi_(ss.getSheetByName(SH.kpi), c);
  fixSummary_(ss.getSheetByName(SH.sum), c);
  linkKintoneIds();
  SpreadsheetApp.flush();

  const broken = findBrokenFormulas_(ss);
  let msg = broken.length
    ? '修正しました。ただし参照切れの数式が残っています：\n' + broken.slice(0, 20).join('\n')
    : '完了しました。参照切れの数式はありません。';
  if (notes.length) msg += '\n\n' + notes.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

// ---------------------------------------------------------------- 旧アクションログからの移行
// 旧形式（4行目が見出し、D列＝自社担当）のデータを、自社担当ごとに担当者タブへ移す
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
      '担当者タブへの移行を止めました。\n\n『アクションログ』D列（自社担当）が空欄、または担当者名以外の代理店があります：\n' +
      unassigned.map(r => `${r[0]}　${r[1]}`).join('\n') +
      `\n\nD列のプルダウンで「${ASSIGNEES.join('」か「')}」を選んでから、もう一度 fixAgencySheet を実行してください。`);
    return false;
  }

  if (!ss.getSheetByName(`${SH.log}（移行前）`)) {
    log.copyTo(ss).setName(`${SH.log}（移行前）`);
  }
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(name);
    const mine = data.filter(r => String(r[3]).trim() === name)
      // A:ID B:代理店名 C:代理店担当者 D〜H:新規項目 I:連絡ツール J:接触手段 K〜N:最終・次回アクション
      .map(r => [r[0], r[1], r[2], '', '', '', '', '', r[4], r[5], r[6], r[7], r[8], r[9]]);
    if (!mine.length) return;
    const filled = sh.getRange(P_FIRST, 1, P_ROWS, 1).getValues().filter(([v]) => v !== '').length;
    sh.getRange(P_FIRST + filled, 1, mine.length, mine[0].length).setValues(mine);
  });

  // 旧アクションログを空にして、合算表示用に作り直せる状態にする
  if (log.getFilter()) log.getFilter().remove();
  log.setConditionalFormatRules([]);
  log.getRange(3, 1, log.getMaxRows() - 2, log.getMaxColumns()).clear().clearDataValidations();
  return true;
}

// ---------------------------------------------------------------- 担当者タブ
function setupPersonTab_(sh, c) {
  const name = sh.getName();
  const nCol = P_COLS.length;
  ensureSize_(sh, P_LAST, nCol);
  titleRow_(sh, `${name}の担当代理店`, nCol,
    '自分が担当する代理店を1行ずつ管理します。接触したら「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。' +
    'ここに入力した内容は『アクションログ』に自動で合算されます。');

  // ステータス（3〜4行目）
  const r = col => `${col}${P_FIRST}:${col}${P_LAST}`;
  const live = `${r(PC.id)},"?*",${r(PC.status)},"<>契約終了"`;
  sh.getRange(3, 1, 1, STATUS_LABELS.length).setValues([STATUS_LABELS])
    .setBackground(COLOR.auto).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, STATUS_LABELS.length).setFormulas([[
    `=COUNTIF(${r(PC.id)},"?*")`,
    `=COUNTIFS(${live},${r(PC.nextDate)},"")+COUNTIFS(${live},${r(PC.nextDate)},"<"&TODAY())`,
    `=COUNTIFS(${live},${r(PC.nextDate)},">="&TODAY(),${r(PC.nextDate)},"<="&TODAY()+設定!$C$8)`,
    `=COUNTIFS(${r(PC.id)},"?*",${r(PC.lastDate)},">="&設定!$C$5)`,
    `=SUM(${r(PC.intro)})`,
    `=SUM(${r(PC.close)})`,
    `=SUM(${r(PC.amount)})`,
  ]]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange(3, 1, 2, STATUS_LABELS.length)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B4').setFontColor(COLOR.redFont);

  headerRow_(sh, P_HEAD, P_COLS);
  const body = sh.getRange(P_FIRST, 1, P_ROWS, nCol);
  body.setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['H', 'K', 'M'].forEach(col => sh.getRange(r(col)).setNumberFormat('yyyy/mm/dd'));
  sh.getRange(r('L')).setWrap(true);
  sh.getRange(r('N')).setWrap(true);

  // 紹介数・成約数・成約金額（『紹介・成約実績』から自動集計）
  const kId = kpiRange_(c.id), kClose = kpiRange_(c.close), kAmount = kpiRange_(c.amount);
  const f = [];
  for (let i = P_FIRST; i <= P_LAST; i++) {
    f.push([
      `=IF(A${i}="","",COUNTIF(${kId},A${i}))`,
      `=IF(A${i}="","",COUNTIFS(${kId},A${i},${kClose},"<>"))`,
      `=IF(A${i}="","",SUMIFS(${kAmount},${kId},A${i}))`,
    ]);
  }
  sh.getRange(`O${P_FIRST}:Q${P_LAST}`).setFormulas(f);
  sh.getRange(`O${P_FIRST}:P${P_LAST}`).setHorizontalAlignment('center');
  sh.getRange(`O${P_FIRST}:Q${P_LAST}`).setNumberFormat('#,##0');

  // プルダウン（選択肢は『設定』）
  const conf = sh.getParent().getSheetByName(SH.conf);
  const list = col => conf.getRange(`${col}${OPT_FIRST}:${col}${OPT_FIRST + 9}`);
  setListValidation_(sh.getRange(r('E')), list('G'));  // 自社副担当
  setListValidation_(sh.getRange(r('F')), list('I'));  // 代理店ステータス
  setListValidation_(sh.getRange(r('G')), list('J'));  // 紹介見込み
  setListValidation_(sh.getRange(r('I')), list('K'));  // 連絡ツール
  setListValidation_(sh.getRange(r('J')), list('H'));  // 接触手段

  // 代理店IDの重複チェック（他の担当者タブも含めて1回だけ）
  const others = ASSIGNEES.filter(n => n !== name)
    .map(n => `+COUNTIF(INDIRECT("'${n}'!A${P_FIRST}:A${P_LAST}"),A${P_FIRST})`).join('');
  sh.getRange(r('A')).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(`=COUNTIF($A$${P_FIRST}:$A$${P_LAST},A${P_FIRST})${others}<=1`)
      .setAllowInvalid(false)
      .setHelpText('この代理店IDは既にどこかの担当者タブに登録されています。')
      .build());

  // 色分け：次回アクション予定日が未入力・期限切れ＝D〜N列を赤、期限間近＝M〜N列を黄（契約終了は除く）
  const red = sh.getRange(`D${P_FIRST}:N${P_LAST}`);
  const yellow = sh.getRange(`M${P_FIRST}:N${P_LAST}`);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(COUNTA($A${P_FIRST}:$N${P_FIRST})>0,$F${P_FIRST}<>"契約終了",OR($M${P_FIRST}="",$M${P_FIRST}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont).setRanges([red]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($M${P_FIRST}<>"",$M${P_FIRST}>=TODAY(),$M${P_FIRST}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow).setRanges([yellow]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$F${P_FIRST}="契約終了"`)
      .setFontColor('#A6A6A6').setRanges([sh.getRange(P_FIRST, 1, P_ROWS, nCol)]).build(),
  ]);

  sh.setFrozenRows(P_HEAD);
  sh.setFrozenColumns(2);
  if (!sh.getFilter()) sh.getRange(P_HEAD, 1, P_ROWS + 1, nCol).createFilter();
}

// ---------------------------------------------------------------- アクションログ（合算）
function buildLog_(sh) {
  const nCol = LOG_COLS.length;
  ensureSize_(sh, LOG_END, nCol);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns()).clear().clearDataValidations();
  titleRow_(sh, 'アクションログ（全担当者の合算）', nCol,
    '各担当者タブの内容を自動で合算した一覧です（ここでは入力しません。修正は各担当者タブで）。上の表は担当者別と合計のステータス。');

  // 担当者別ステータス表（3行目〜）
  const n = ASSIGNEES.length;
  sh.getRange(3, 1, 1, STATUS_LABELS.length + 1).setValues([['担当者'].concat(STATUS_LABELS)])
    .setBackground(COLOR.auto).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  const cols = STATUS_LABELS.map((_, i) => colLetter_(i + 1));
  // 名前・「合計」は文字として、数値は数式として別々に書く（setFormulas に文字を渡すと #NAME? になる）
  sh.getRange(4, 1, n, 1).setValues(ASSIGNEES.map(name => [name]));
  sh.getRange(4, 2, n, STATUS_LABELS.length).setFormulas(
    ASSIGNEES.map(name => cols.map(col => `='${name}'!${col}4`)));
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

  headerRow_(sh, LOG_HEAD, LOG_COLS.map(([h, , w]) => [h, 'auto', w]));

  // 担当者タブを縦に合算（自社担当＝タブ名をC列に入れる）。代理店IDはkintoneリンク付き
  const rest = P_COLS.slice(2).map((_, i) => i + 3).join(',');
  const parts = ASSIGNEES.map((name, i) => {
    const v = `p${i}`;
    return `HSTACK(CHOOSECOLS(${v},1,2),IF(CHOOSECOLS(${v},1)<>"","${name}",""),CHOOSECOLS(${v},${rest}))`;
  });
  const lets = ASSIGNEES.map((name, i) =>
    `p${i},'${name}'!A${P_FIRST}:${colLetter_(P_COLS.length)}${P_LAST}`).join(',');
  const others = LOG_COLS.slice(1).map((_, i) => i + 2).join(',');
  sh.getRange(LOG_FIRST, 1).setFormula(
    `=IFERROR(ARRAYFORMULA(LET(${lets},all,VSTACK(${parts.join(',')}),f,FILTER(all,CHOOSECOLS(all,1)<>""),` +
    `HSTACK(HYPERLINK("${KINTONE_APP_URL}?query="&ENCODEURL("${KINTONE_ID_FIELD} = """&CHOOSECOLS(f,1)&""""),CHOOSECOLS(f,1)),` +
    `CHOOSECOLS(f,${others})))),"")`);

  const rr = col => sh.getRange(`${col}${LOG_FIRST}:${col}${LOG_END}`);
  sh.getRange(LOG_FIRST, 1, LOG_END - LOG_FIRST + 1, nCol).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  ['I', 'L', 'N'].forEach(col => rr(col).setNumberFormat('yyyy/mm/dd'));
  ['M', 'O'].forEach(col => rr(col).setWrap(true));
  rr('R').setNumberFormat('#,##0');
  sh.getRange(`P${LOG_FIRST}:Q${LOG_END}`).setHorizontalAlignment('center');

  // 色分け（担当者タブと同じ基準。列は自社担当の分だけ右にずれる）
  const f = LOG_FIRST;
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${f}<>"",$G${f}<>"契約終了",OR($N${f}="",$N${f}<TODAY()))`)
      .setBackground(COLOR.red).setFontColor(COLOR.redFont)
      .setRanges([sh.getRange(`E${f}:O${LOG_END}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($N${f}<>"",$N${f}>=TODAY(),$N${f}-TODAY()<=INDIRECT("設定!C8"))`)
      .setBackground(COLOR.yellow).setRanges([sh.getRange(`N${f}:O${LOG_END}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$G${f}="契約終了"`)
      .setFontColor('#A6A6A6').setRanges([sh.getRange(f, 1, LOG_END - f + 1, nCol)]).build(),
  ]);
  sh.setFrozenRows(LOG_HEAD);
  sh.setFrozenColumns(2);
}

// ---------------------------------------------------------------- 『設定』の選択肢
function writeLists_(ss) {
  const conf = ss.getSheetByName(SH.conf);
  conf.getRange('G12').setValue('入力リスト（担当者タブのプルダウン）').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange('G13:K13').setValues([['自社担当', '接触手段', '代理店ステータス', '紹介見込み', '連絡ツール']])
    .setFontWeight('bold').setBackground(COLOR.input).setHorizontalAlignment('center');
  const last = OPT_FIRST + 9;
  writeOptions_(conf.getRange(`G${OPT_FIRST}:G${last}`), ASSIGNEES);
  writeOptions_(conf.getRange(`I${OPT_FIRST}:I${last}`), STATUS_OPTIONS);
  writeOptions_(conf.getRange(`J${OPT_FIRST}:J${last}`), PROSPECT_OPTIONS);
  // 接触手段・連絡ツールは空のときだけ初期値を入れる（手で追加した項目を消さない）
  const keep = (col, items) => {
    const rng = conf.getRange(`${col}${OPT_FIRST}:${col}${last}`);
    if (rng.getValues().every(([v]) => v === '')) writeOptions_(rng, items);
  };
  keep('H', ['電話', '訪問', 'オンライン面談', 'メール', 'LINE/チャット', '勉強会・セミナー', '会食']);
  keep('K', TOOL_OPTIONS);
  conf.getRange(`G${OPT_FIRST}:K${last}`)
    .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
  conf.getRange('G25').setValue('※ 自社担当はスクリプト上部の ASSIGNEES で管理。接触手段・連絡ツールは空欄に追記するとプルダウンに反映されます（各10件まで）。');
}

// ---------------------------------------------------------------- 紹介・成約実績
// 見出し（4行目）の名前から列番号を探す。列を増やしても動くようにするため
function kpiCols_(sh) {
  const h = sh.getRange(4, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(v => v.replace(/\s/g, ''));
  const find = (label, re) => {
    const i = h.findIndex(v => re.test(v));
    if (i < 0) throw new Error(`『${SH.kpi}』の4行目に「${label}」の列が見つかりません`);
    return colLetter_(i + 1);
  };
  return {
    id: find('代理店ID', /代理店ID/),
    intro: find('紹介日', /紹介日/),
    close: find('成約日', /^成約日/),
    komon: find('成約商材（顧問種別）', /顧問種別/),
    anken1: find('成約商材（案件種別①）', /案件種別(①|1)/),
    anken2: find('成約商材（案件種別②）', /案件種別(②|2)/),
    amount: find('成約金額（円）', /成約金額/),
    reward: find('代理店報酬（円）', /代理店報酬/),
  };
}

function kpiRange_(col) {
  return `'${SH.kpi}'!$${col}$5:$${col}$${KPI_END}`;
}

function fixKpi_(sh, c) {
  const notes = [];
  sh.getRange('A1').setValue('紹介・成約実績（紹介1件につき1行）');
  sh.getRange('B4').setBackground(COLOR.auto);  // 代理店名は自動表示
  sh.getRange('A2').setValue(
    '紹介1件につき1行。代理店IDを選ぶと代理店名が自動表示されます（IDは担当者タブに登録したもの）。' +
    '成約したら成約日・成約商材・成約金額・代理店報酬を入力。');

  const n = KPI_END - 4;
  const f = [];
  for (let r = 5; r <= KPI_END; r++) {
    f.push([`=IF(A${r}="","",IFERROR(INDEX('${SH.log}'!$B$${LOG_FIRST}:$B$${LOG_END},` +
            `MATCH(A${r},'${SH.log}'!$A$${LOG_FIRST}:$A$${LOG_END},0)),"※未登録のID"))`]);
  }
  sh.getRange(5, 2, n, 1).setFormulas(f);
  sh.getRange(`C5:C${KPI_END}`).setNumberFormat('@');                    // 紹介社名
  sh.getRange(`${c.intro}5:${c.intro}${KPI_END}`).setNumberFormat('yyyy/mm/dd');
  sh.getRange(`${c.close}5:${c.close}${KPI_END}`).setNumberFormat('yyyy/mm/dd');
  [c.komon, c.anken1, c.anken2].forEach(col =>
    sh.getRange(`${col}5:${col}${KPI_END}`).setNumberFormat('@'));
  sh.getRange(`${c.amount}5:${c.amount}${KPI_END}`).setNumberFormat('#,##0');
  sh.getRange(`${c.reward}5:${c.reward}${KPI_END}`).setNumberFormat('#,##0');

  // 代理店IDはサマリーに並んだID（＝担当者タブ登録済み）から選択
  setListValidation_(sh.getRange(5, 1, n, 1),
    sh.getParent().getSheetByName(SH.sum).getRange(`A${SUM_FIRST}:A${SUM_END}`));

  // 成約商材のプルダウン（選択肢は『設定』M〜O列）
  const conf = sh.getParent().getSheetByName(SH.conf);
  conf.getRange('M12').setValue('成約商材の選択肢（kintoneと同じ内容）').setFontWeight('bold').setFontColor(COLOR.title);
  conf.getRange('M13:O13').setValues([['顧問種別', '案件種別①', '案件種別②']])
    .setFontWeight('bold').setBackground(COLOR.input).setHorizontalAlignment('center');
  writeOptions_(conf.getRange(`N${OPT_FIRST}:N${OPT_END}`), ANKEN1_OPTIONS);
  writeOptions_(conf.getRange(`O${OPT_FIRST}:O${OPT_END}`), ANKEN2_OPTIONS);
  conf.setColumnWidths(13, 3, 180);

  // 顧問種別：『設定』M列が空なら、既に設定済みのプルダウン項目を引き継ぐ
  const komonList = conf.getRange(`M${OPT_FIRST}:M${OPT_END}`);
  const komonCol = sh.getRange(`${c.komon}5:${c.komon}${KPI_END}`);
  if (komonList.getValues().every(([v]) => v === '')) {
    const dv = komonCol.getCell(1, 1).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      writeOptions_(komonList, dv.getCriteriaValues()[0]);
    }
  }
  if (komonList.getValues().some(([v]) => v !== '')) {
    setListValidation_(komonCol, komonList);
  } else {
    notes.push('※ 顧問種別の選択肢が未登録です。『設定』シートM14から下にkintoneの顧問種別の項目を入力し、もう一度実行してください。');
  }
  setListValidation_(sh.getRange(`${c.anken1}5:${c.anken1}${KPI_END}`), conf.getRange(`N${OPT_FIRST}:N${OPT_END}`));
  setListValidation_(sh.getRange(`${c.anken2}5:${c.anken2}${KPI_END}`), conf.getRange(`O${OPT_FIRST}:O${OPT_END}`));
  return notes;
}

// ---------------------------------------------------------------- サマリー
function fixSummary_(sh, c) {
  sh.getCharts().forEach(ch => sh.removeChart(ch));
  const lastRow = Math.max(sh.getMaxRows(), SUM_END);
  sh.getRange(5, 1, lastRow - 4, 7).clearContent().setBackground(null).setFontWeight('normal');

  sh.getRange('A1').setValue('サマリー（代理店別）');
  sh.getRange('A2').setValue(
    'すべて自動計算。代理店IDは担当者タブ（アクションログ）から自動で並びます。' +
    '稼働率＝直近の対象期間（『設定』C6、初期値3ヶ月）のうち紹介があった月の割合。');
  sh.getRange('A4:G4').setValues([['代理店ID', '代理店名', '稼働率', '紹介数', '成約数', '成約金額（円）', '代理店報酬（円）']]);

  const L = `'${SH.log}'`;
  const logId = `${L}!$A$${LOG_FIRST}:$A$${LOG_END}`, logName = `${L}!$B$${LOG_FIRST}:$B$${LOG_END}`;
  const kId = kpiRange_(c.id), kIntro = kpiRange_(c.intro), kClose = kpiRange_(c.close);
  const kAmount = kpiRange_(c.amount), kReward = kpiRange_(c.reward);

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

  // 代理店ID：アクションログから重複なしで自動展開
  sh.getRange(`A${SUM_FIRST}`).setFormula(
    `=IFERROR(LET(ids,UNIQUE(FILTER(${logId},${logId}<>"")),` +
    `ARRAYFORMULA(IFERROR(HYPERLINK("${KINTONE_APP_URL}?query="&ENCODEURL("${KINTONE_ID_FIELD} = """&ids&""""),ids),ids))),"")`);

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
  conf.getRange('D8').setValue('担当者タブ・アクションログで、次回アクション予定日のこの日数前から黄色表示');
  conf.getRange('B7:D7').clear();
  conf.getRange('B9:D9').clear();
  conf.getRange('B12:E19').clear();  // 使われなくなった代理店ランク定義
  const unused = ['list_ランク', 'list_接触目的', 'list_温度感', 'list_完了'];
  ss.getNamedRanges().filter(nr => unused.includes(nr.getName())).forEach(nr => nr.remove());
}

function writePlan_(sh) {
  sh.getRange('A1').setValue('定例アクション計画（代理店を動かし続けるための接点づくり）');
  sh.getRange('A2').setValue('いつ・どの代理店に・何をするかの運用ルール。対象の代理店は担当者タブ・『サマリー』の表示で判断します。');
  const plan = [
    ['毎週月曜', '自分の担当代理店',
     '自分の担当者タブで赤（次回予定日なし・期限切れ）→黄（期限間近）の順に連絡し、最終アクションと次回アクションを更新',
     '連絡の抜け漏れをゼロにする', '各担当', '担当者タブ'],
    ['毎週月曜', '全担当者',
     '『アクションログ』上部のステータス表で、担当者ごとの要連絡件数・今月の接触数を確認', '担当者間の偏り・漏れを防ぐ', '責任者', 'アクションログ'],
    ['毎週', '紹介見込み「高」・ステータス「稼働中」の代理店',
     '紹介予定・進捗の確認連絡', '紹介の流れを止めない', '各担当', '電話・LINE'],
    ['紹介を受けた当日中', '紹介元の代理店',
     '紹介受領のお礼と今後の流れを連絡し、『紹介・成約実績』に1行追加',
     '紹介して良かったと感じてもらう', '各担当', '電話・チャット'],
    ['初回面談から3営業日以内', '紹介元の代理店',
     '面談結果・次のステップを報告', '代理店が顧客にフォローしやすくする', '各担当', 'メール・チャット'],
    ['成約時（24時間以内）', '紹介元の代理店',
     '成約報告と報酬見込みの連絡。『紹介・成約実績』に成約日・成約商材・成約金額・代理店報酬を入力',
     '成功体験を共有し次の紹介につなげる', '各担当', '電話'],
    ['毎月初', '全代理店',
     '新商材・キャンペーン情報、紹介トークのポイントを配信', '紹介のきっかけを定期的に提供', 'マーケ担当', 'メール・チャット'],
    ['毎月末', '全代理店',
     '『サマリー』をもとに、代理店別の紹介・成約実績を報告', '実績を見える化して関係を深める', '各担当', 'メール'],
    ['毎月末', '稼働率0%（『サマリー』で灰色）・ステータス「休眠」の代理店',
     '再活性化の個別提案（勉強会招待、同行営業の提案など）。ステータスを「フォロー強化」に変更', '休眠代理店を掘り起こす', '各担当', '電話・訪問'],
    ['四半期ごと', '全代理店',
     '代理店向け勉強会（税制改正・商材説明・成功事例共有）', '知識を更新し、紹介しやすい状態をつくる', '責任者', 'セミナー・オンライン'],
    ['半期ごと', '重点代理店（『サマリー』で紹介数・成約数が多い代理店）',
     '個別面談（目標のすり合わせ・インセンティブ案内）', '重点代理店のモチベーション維持', '責任者', '訪問・オンライン'],
    ['顧客の決算期が集中する時期の前', '全代理店',
     '節税・決算対策の提案キャンペーンを案内', '繁忙期前に紹介を集中させる', 'マーケ担当', 'メール・勉強会'],
    ['契約直後の1ヶ月', 'ステータス「立上げ中」の代理店',
     '立上げ面談（商材説明・紹介フロー説明・紹介したい顧客像の洗い出し）。初回面談日を記録', '最初の1件目の紹介を早く出す', '各担当', '訪問・オンライン'],
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
    ['担当者タブ', `${ASSIGNEES.join('・')}のタブ。自分の担当代理店を1行ずつ入力する場所。上部に本人のステータス（要連絡件数・紹介数など）。`],
    ['アクションログ', '担当者タブを自動で合算した一覧（入力しない）。上部に担当者別・合計のステータス表。'],
    ['紹介・成約実績', '紹介1件につき1行。代理店IDを選ぶと代理店名が自動表示。成約したら成約日・成約商材・成約金額・代理店報酬を入力。'],
    ['サマリー', '代理店別の稼働率・紹介数・成約数・成約金額・代理店報酬を自動集計（入力不要）。'],
    ['定例アクション計画', 'いつ・どの代理店に・何をするかの運用ルール。'],
    ['設定', '稼働率の対象期間（C6）、期限間近の日数（C8）、プルダウンの選択肢（G〜K列、成約商材＝M〜O列）。'],
    ['', null],
    ['■ 色のルール', null],
    ['黄色の見出し', '入力する列。'],
    ['灰色の見出し', '自動計算の列（数式が入っているので上書きしない）。'],
    ['赤', '担当者タブ・アクションログ：次回アクション予定日が未入力、または予定日を過ぎている（契約終了は除く）。'],
    ['黄', '担当者タブ・アクションログ：次回アクション予定日まで『設定』C8の日数以内。'],
    ['サマリーの灰色', '稼働率0%（対象期間に紹介がない）。'],
    ['サマリーの緑', '稼働率100%（対象期間の毎月に紹介がある）。'],
    ['', null],
    ['■ 運用の流れ', null],
    ['① 代理店を登録', '自分の担当者タブに代理店ID・代理店名・代理店担当者・ステータスを入力。代理店IDはkintone『8. 代理店管理』と同じIDを使う。'],
    ['② 接触したら更新', 'その代理店の行の「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。'],
    ['③ 紹介を受けたら記録', '『紹介・成約実績』に1行追加（代理店ID・紹介社名・紹介日）。成約したら同じ行に成約日・成約商材・成約金額・代理店報酬を追記。'],
    ['④ 毎週月曜に確認', '担当者：自分のタブで赤→黄の順に連絡。責任者：『アクションログ』上部で担当者ごとの要連絡件数を確認。'],
    ['⑤ 月末に振り返り', '『サマリー』で稼働率0%（灰色）の代理店を洗い出し、『定例アクション計画』に沿って再活性化の手を打つ。'],
    ['担当替えのとき', '担当者タブの行を切り取り、相手の担当者タブに貼り付ける。'],
    ['', null],
    ['■ 項目の意味', null],
    ['代理店ステータス', '立上げ中（契約直後）／稼働中（紹介が出ている）／フォロー強化（紹介が止まり気味）／休眠／契約終了（赤表示・件数の対象外）。'],
    ['紹介見込み', '近いうちに紹介が出そうかの担当者の見立て（高／中／低／なし）。'],
    ['稼働率', '直近の対象期間（初期値3ヶ月）のうち、紹介があった月の割合。3ヶ月中2ヶ月紹介あり→67%。'],
    ['紹介数・成約数・成約金額', '『紹介・成約実績』から自動集計。'],
    ['代理店IDのリンク', 'クリックするとkintone『8. 代理店管理』でその代理店に絞り込んだ一覧が開く（入力すると自動でリンク化）。'],
    ['', null],
    ['■ 管理者向け', null],
    ['列を変えたとき', '拡張機能 → Apps Script で fixAgencySheet を実行すると、数式・プルダウン・色分けが今の列に合わせて整う。'],
    ['担当者が増えたとき', 'スクリプト上部の ASSIGNEES に名前を追加して fixAgencySheet を実行（担当者タブが自動で作られる）。'],
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
// 担当者タブ・紹介・成約実績の代理店IDを、kintoneへのリンクにする（既存データを一括変換）
function linkKintoneIds() {
  const ss = SpreadsheetApp.getActive();
  ASSIGNEES.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() >= P_FIRST) linkIds_(sh.getRange(P_FIRST, 1, sh.getLastRow() - P_FIRST + 1, 1));
  });
  const kpi = ss.getSheetByName(SH.kpi);
  if (kpi.getLastRow() >= 5) linkIds_(kpi.getRange(5, 1, kpi.getLastRow() - 4, 1));
}

// 代理店IDを入力・貼り付け・プルダウン選択したら自動でリンク化（シンプルトリガー）
function onEdit(e) {
  const sh = e.range.getSheet();
  const first = ASSIGNEES.includes(sh.getName()) ? P_FIRST : sh.getName() === SH.kpi ? 5 : 0;
  if (!first || e.range.getColumn() !== 1) return;
  const top = Math.max(e.range.getRow(), first);
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
    if (sh.getName().endsWith('（移行前）')) return;
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
