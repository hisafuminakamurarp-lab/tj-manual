/**
 * 代理店アクション管理シート：連動修正スクリプト（Google Apps Script）
 *
 * 「代理店一覧」シート削除で切れた数式を、今のシート構成のまま繋ぎ直す。
 *   - サマリー        … 代理店別集計（IDはアクションログから自動抽出）
 *   - 紹介・成約実績  … 代理店名をアクションログから自動表示
 *   - アクションログ  … 参照切れ数式の削除、プルダウン、色分け（予定日なし・期限切れ＝D〜J列赤）
 *   - 使い方 / 設定 / 定例アクション計画 … 説明文を今の構成に合わせて更新
 *   - 代理店ID … kintone『8. 代理店管理』のレコードへリンク（入力すると自動でリンク化）
 *
 * 使い方：拡張機能 → Apps Script に貼り付け → fixAgencySheet を実行。
 * 何度実行しても同じ状態になる（入力済みのデータは消さない）。
 */

const SH = {
  log: 'アクションログ',
  kpi: '紹介・成約実績',
  sum: 'サマリー',
  plan: '定例アクション計画',
  howto: '使い方',
  conf: '設定',
};
const LOG_END = 2002;   // アクションログの最終行
const KPI_END = 3003;   // 紹介・成約実績の最終行
const SUM_FIRST = 6;    // サマリーの代理店1行目（5行目は合計）
const SUM_END = 105;    // 代理店100社分

// kintone『8. 代理店管理』アプリのURLと、代理店IDのフィールドコード
// 代理店IDで一覧を絞り込んだページを開く（例：代理店ID = "P-0000016"）
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

function fixAgencySheet() {
  const ss = SpreadsheetApp.getActive();
  const need = Object.values(SH).filter(n => !ss.getSheetByName(n));
  if (need.length) throw new Error('シートが見つかりません：' + need.join('、'));

  const c = kpiCols_(ss.getSheetByName(SH.kpi));
  fixLog_(ss.getSheetByName(SH.log));
  const notes = fixKpi_(ss.getSheetByName(SH.kpi), c);
  fixSummary_(ss.getSheetByName(SH.sum), c);
  fixTexts_(ss);
  linkKintoneIds();
  SpreadsheetApp.flush();

  const broken = findBrokenFormulas_(ss);
  let msg = broken.length
    ? '修正しました。ただし参照切れの数式が残っています：\n' + broken.slice(0, 20).join('\n')
    : '修正が完了しました。参照切れの数式はありません。';
  if (notes.length) msg += '\n\n' + notes.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

// ---------------------------------------------------------------- アクションログ
function fixLog_(sh) {
  sh.getRange('A2').setValue(
    '代理店ごとに1行。接触したら「最終アクション日・内容」と「次回アクション予定日・内容」を更新します。' +
    '代理店ID・代理店名はここが元データになり、『紹介・成約実績』『サマリー』に自動反映されます。');

  // C列（代理店担当者）：削除済みシートを参照する数式だけを消し、入力済みの値は残す
  const rng = sh.getRange(5, 3, LOG_END - 4, 1);
  const formulas = rng.getFormulas();
  const values = rng.getValues();
  rng.setValues(formulas.map((f, i) => [f[0] ? '' : values[i][0]]));

  sh.getRange(5, 7, LOG_END - 4, 1).setNumberFormat('yyyy/mm/dd');  // 最終アクション日
  sh.getRange(5, 9, LOG_END - 4, 1).setNumberFormat('yyyy/mm/dd');  // 次回アクション予定日

  // プルダウン（選択肢は『設定』シート）
  const conf = sh.getParent().getSheetByName(SH.conf);
  setListValidation_(sh.getRange(5, 4, LOG_END - 4, 1), conf.getRange('G14:G23'));  // 自社担当
  setListValidation_(sh.getRange(5, 6, LOG_END - 4, 1), conf.getRange('H14:H23'));  // 接触手段

  // 代理店IDの重複チェック
  sh.getRange(5, 1, LOG_END - 4, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied(`=COUNTIF($A$5:$A$${LOG_END},A5)<=1`)
      .setAllowInvalid(false)
      .setHelpText('同じ代理店IDが既に登録されています。')
      .build());

  setLogColors_(sh);
}

// 次回アクション予定日が未入力・期限切れ＝D〜J列を赤、期限間近（設定C8の日数以内）＝I〜J列を黄
function setLogColors() {
  setLogColors_(SpreadsheetApp.getActive().getSheetByName(SH.log));
}

function setLogColors_(sh) {
  const red = sh.getRange(5, 4, LOG_END - 4, 7);     // D〜J
  const yellow = sh.getRange(5, 9, LOG_END - 4, 2);  // I〜J
  const rules = sh.getConditionalFormatRules().filter(r => !overlaps_(r, red));
  rules.unshift(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(COUNTA($A5:$J5)>0,OR($I5="",$I5<TODAY()))')
      .setBackground('#F8CBAD').setFontColor('#9C0006')
      .setRanges([red]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($I5<>"",$I5>=TODAY(),$I5-TODAY()<=INDIRECT("設定!C8"))')
      .setBackground('#FFE699')
      .setRanges([yellow]).build());
  sh.setConditionalFormatRules(rules);
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

function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function fixKpi_(sh, c) {
  const notes = [];
  sh.getRange('A2').setValue(
    '紹介1件につき1行。代理店IDを選ぶと代理店名が自動表示されます（IDは『アクションログ』に登録したもの）。' +
    '成約したら成約日・成約商材・成約金額・代理店報酬を入力。');

  const n = KPI_END - 4;
  const f = [];
  for (let r = 5; r <= KPI_END; r++) {
    f.push([`=IF(A${r}="","",IFERROR(INDEX('${SH.log}'!$B$5:$B$${LOG_END},` +
            `MATCH(A${r},'${SH.log}'!$A$5:$A$${LOG_END},0)),"※未登録のID"))`]);
  }
  sh.getRange(5, 2, n, 1).setFormulas(f);
  sh.getRange(`C5:C${KPI_END}`).setNumberFormat('@');                    // 紹介社名
  sh.getRange(`${c.intro}5:${c.intro}${KPI_END}`).setNumberFormat('yyyy/mm/dd');
  sh.getRange(`${c.close}5:${c.close}${KPI_END}`).setNumberFormat('yyyy/mm/dd');
  [c.komon, c.anken1, c.anken2].forEach(col =>
    sh.getRange(`${col}5:${col}${KPI_END}`).setNumberFormat('@'));
  sh.getRange(`${c.amount}5:${c.amount}${KPI_END}`).setNumberFormat('#,##0');
  sh.getRange(`${c.reward}5:${c.reward}${KPI_END}`).setNumberFormat('#,##0');

  // 代理店IDはサマリーに並んだID（＝アクションログ登録済み）から選択
  setListValidation_(sh.getRange(5, 1, n, 1),
    sh.getParent().getSheetByName(SH.sum).getRange(`A${SUM_FIRST}:A${SUM_END}`));

  // 成約商材のプルダウン（選択肢は『設定』M〜O列）
  const conf = sh.getParent().getSheetByName(SH.conf);
  conf.getRange('M12').setValue('成約商材の選択肢（kintoneと同じ内容）').setFontWeight('bold').setFontColor('#1F3864');
  conf.getRange('M13:O13').setValues([['顧問種別', '案件種別①', '案件種別②']])
    .setFontWeight('bold').setBackground('#FFF2CC').setHorizontalAlignment('center');
  writeOptions_(conf.getRange(`N${OPT_FIRST}:N${OPT_END}`), ANKEN1_OPTIONS);
  writeOptions_(conf.getRange(`O${OPT_FIRST}:O${OPT_END}`), ANKEN2_OPTIONS);
  conf.setColumnWidths(13, 3, 180);

  // 顧問種別：『設定』M列が空なら、F列に既に設定済みのプルダウン項目を引き継ぐ
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

function writeOptions_(range, items) {
  const rows = range.getNumRows();
  const vals = [];
  for (let i = 0; i < rows; i++) vals.push([items[i] || '']);
  range.setValues(vals);
}

// ---------------------------------------------------------------- サマリー
function fixSummary_(sh, c) {
  sh.getCharts().forEach(c => sh.removeChart(c));  // 月次前提のグラフは列構成と合わないため削除
  const lastRow = Math.max(sh.getMaxRows(), SUM_END);
  sh.getRange(5, 1, lastRow - 4, 7).clearContent().setBackground(null).setFontWeight('normal');

  sh.getRange('A1').setValue('サマリー（代理店別）');
  sh.getRange('A2').setValue(
    'すべて自動計算。代理店IDは『アクションログ』から自動で並びます。' +
    '稼働率＝直近の対象期間（『設定』C6、初期値3ヶ月）のうち紹介があった月の割合。');
  sh.getRange('A4:G4').setValues([['代理店ID', '代理店名', '稼働率', '紹介数', '成約数', '成約金額（円）', '代理店報酬（円）']]);

  const L = `'${SH.log}'`;
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
  sh.getRange('A5:G5').setBackground('#DDEBF7').setFontWeight('bold');

  // 代理店ID：アクションログから重複なしで自動展開
  sh.getRange(`A${SUM_FIRST}`).setFormula(
    `=IFERROR(LET(ids,UNIQUE(FILTER(${L}!A5:A${LOG_END},${L}!A5:A${LOG_END}<>"")),` +
    `ARRAYFORMULA(IFERROR(HYPERLINK("${KINTONE_APP_URL}?query="&ENCODEURL("${KINTONE_ID_FIELD} = """&ids&""""),ids),ids))),"")`);

  const rows = [];
  for (let r = SUM_FIRST; r <= SUM_END; r++) {
    const A = `$A${r}`;
    const m = 'DATE(YEAR(設定!$C$5),MONTH(設定!$C$5)';
    const seq = 'SEQUENCE(設定!$C$6)';
    rows.push([
      `=IF(${A}="","",IFERROR(INDEX(${L}!$B$5:$B$${LOG_END},MATCH(${A},${L}!$A$5:$A$${LOG_END},0)),""))`,
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
    .setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(5);

  // 稼働率：0%＝グレー（休眠）、100%＝緑
  const target = sh.getRange(SUM_FIRST, 3, SUM_END - SUM_FIRST + 1, 1);
  const rules = sh.getConditionalFormatRules().filter(r => !overlaps_(r, target));
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${SUM_FIRST}<>"",$C${SUM_FIRST}=0)`)
      .setBackground('#E7E6E6').setFontColor('#7F7F7F').setRanges([target]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${SUM_FIRST}<>"",$C${SUM_FIRST}>=1)`)
      .setBackground('#C6EFCE').setRanges([target]).build());
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- 説明文
function fixTexts_(ss) {
  ss.getSheetByName(SH.plan).getRange('C5').setValue(
    '『アクションログ』で赤（次回予定日なし・期限切れ）→黄（期限間近）の代理店に連絡し、次回アクションを入れる');

  const conf = ss.getSheetByName(SH.conf);
  conf.getRange('B6').setValue('稼働率の対象期間（ヶ月）');
  conf.getRange('D6').setValue('『サマリー』の稼働率：直近この月数のうち紹介があった月の割合');
  conf.getRange('B8').setValue('期限間近の表示（日前）');
  conf.getRange('D8').setValue('『アクションログ』で次回アクション予定日のこの日数前から黄色表示');
  conf.getRange('B7:D7').clearContent().setBackground(null);  // 旧「稼働中とみなす紹介数」（未使用）
  conf.getRange('B9:D9').clearContent().setBackground(null);  // 旧「年度開始月」（未使用）
  conf.getRange('B12').setValue('代理店ランク定義（参考：接触頻度の目安）');

  const lines = [
    ['■ シート構成', null],
    ['アクションログ', '代理店ごとに1行。代理店の基本情報と、最終アクション・次回アクションを管理する。代理店IDの元データ。'],
    ['紹介・成約実績', '紹介1件につき1行。代理店IDを選ぶと代理店名が自動表示。成約したら成約日・金額・報酬を入力。'],
    ['サマリー', '代理店別の稼働率・紹介数・成約数・成約金額・代理店報酬を自動集計（入力不要）。'],
    ['定例アクション計画', '週次・月次・四半期で行う代理店向けアクションの一覧（運用例）。'],
    ['設定', '稼働率の対象期間、期限間近の日数、プルダウンの選択肢（成約商材はM〜O列）を管理。'],
    ['', null],
    ['■ 色のルール', null],
    ['黄色の見出し', '入力する列。'],
    ['灰色の見出し／数式の列', '自動計算（数式が入っているので上書きしない）。'],
    ['アクションログ 赤（D〜J列）', '次回アクション予定日が未入力、または予定日を過ぎている。'],
    ['次回アクション 黄', '次回アクション予定日まで『設定』C8の日数以内。'],
    ['', null],
    ['■ 運用の流れ', null],
    ['① 代理店を登録', '『アクションログ』に代理店ID・代理店名・担当者を入力。代理店IDはkintone『マスタ｜代理店管理』と同じIDを使う。'],
    ['② 接触したら更新', 'その代理店の行の「最終アクション日・内容」を書き換え、「次回アクション予定日・内容」を必ず入れる。'],
    ['③ 紹介を受けたら記録', '『紹介・成約実績』に1行追加（代理店ID・紹介社名・紹介日）。成約したら同じ行に成約日・成約商材・金額・報酬を追記。'],
    ['④ 毎週月曜に確認', '『アクションログ』で赤の行（予定日なし・期限切れ）→黄の行（期限間近）の順に連絡し、次回アクションを入れる。'],
    ['⑤ 月末に振り返り', '『サマリー』で稼働率が0%（灰色）の代理店を洗い出し、再活性化の打ち手を決める。'],
    ['', null],
    ['■ 自動計算の意味', null],
    ['稼働率', '直近の対象期間（初期値3ヶ月）のうち、紹介があった月の割合。3ヶ月中2ヶ月紹介あり→67%。'],
    ['紹介数・成約数', '『紹介・成約実績』の行数と、そのうち成約日が入っている行数。'],
    ['代理店IDのリンク', '代理店IDはkintone『8. 代理店管理』のレコードへのリンク。クリックするとkintoneでその代理店に絞り込んだ一覧が開く（入力すると自動でリンク化）。'],
  ];
  const how = ss.getSheetByName(SH.howto);
  how.getRange('B4:C40').clearContent().setBackground(null).setFontWeight('normal');
  how.getRange(4, 2, lines.length, 2).setValues(lines.map(([k, v]) => [k, v || '']));
  lines.forEach(([k, v], i) => {
    const cell = how.getRange(4 + i, 2);
    cell.setFontWeight('bold');
    if (v === null) cell.setFontColor('#1F3864').setFontSize(11);
  });
  how.getRange(4, 3, lines.length, 1).setWrap(true);
  how.getRange('B12').setBackground('#FFF2CC');
  how.getRange('B13').setBackground('#D9D9D9');
  how.getRange('B14').setBackground('#F8CBAD');
  how.getRange('B15').setBackground('#FFE699');
}

// ---------------------------------------------------------------- kintoneリンク
// アクションログ・紹介・成約実績の代理店IDを、kintoneのレコードへのリンクにする（既存データを一括変換）
function linkKintoneIds() {
  const ss = SpreadsheetApp.getActive();
  [SH.log, SH.kpi].forEach(name => {
    const sh = ss.getSheetByName(name);
    const last = sh.getLastRow();
    if (last >= 5) linkIds_(sh.getRange(5, 1, last - 4, 1));
  });
}

// 代理店IDを入力・貼り付け・プルダウン選択したら自動でリンク化（シンプルトリガー）
function onEdit(e) {
  const sh = e.range.getSheet();
  if (![SH.log, SH.kpi].includes(sh.getName())) return;
  if (e.range.getColumn() !== 1) return;
  const top = Math.max(e.range.getRow(), 5);
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
