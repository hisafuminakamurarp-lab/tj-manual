/**
 * 代理店アクション管理シート：連動修正スクリプト（Google Apps Script）
 *
 * 「代理店一覧」シート削除で切れた数式を、今のシート構成のまま繋ぎ直す。
 *   - サマリー        … 代理店別集計（IDはアクションログから自動抽出）
 *   - 紹介・成約実績  … 代理店名をアクションログから自動表示
 *   - アクションログ  … 参照切れ数式の削除、プルダウン、色分け（予定日なし・期限切れ＝D〜J列赤）
 *   - 使い方 / 設定 / 定例アクション計画 … 説明文を今の構成に合わせて更新
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

function fixAgencySheet() {
  const ss = SpreadsheetApp.getActive();
  const need = Object.values(SH).filter(n => !ss.getSheetByName(n));
  if (need.length) throw new Error('シートが見つかりません：' + need.join('、'));

  fixLog_(ss.getSheetByName(SH.log));
  fixKpi_(ss.getSheetByName(SH.kpi));
  fixSummary_(ss.getSheetByName(SH.sum));
  fixTexts_(ss);
  SpreadsheetApp.flush();

  const broken = findBrokenFormulas_(ss);
  const msg = broken.length
    ? '修正しました。ただし参照切れの数式が残っています：\n' + broken.slice(0, 20).join('\n')
    : '修正が完了しました。参照切れの数式はありません。';
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
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(OR($A5<>"",$B5<>""),OR($I5="",$I5<TODAY()))')
      .setBackground('#F8CBAD').setFontColor('#9C0006')
      .setRanges([red]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($I5<>"",$I5>=TODAY(),$I5-TODAY()<=INDIRECT("設定!C8"))')
      .setBackground('#FFE699')
      .setRanges([yellow]).build());
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- 紹介・成約実績
function fixKpi_(sh) {
  sh.getRange('A2').setValue(
    '紹介1件につき1行。代理店IDを選ぶと代理店名が自動表示されます（IDは『アクションログ』に登録したもの）。' +
    '成約したら成約日・成約金額・代理店報酬を入力。');

  const n = KPI_END - 4;
  const f = [];
  for (let r = 5; r <= KPI_END; r++) {
    f.push([`=IF(A${r}="","",IFERROR(INDEX('${SH.log}'!$B$5:$B$${LOG_END},` +
            `MATCH(A${r},'${SH.log}'!$A$5:$A$${LOG_END},0)),"※未登録のID"))`]);
  }
  sh.getRange(5, 2, n, 1).setFormulas(f);
  sh.getRange(5, 3, n, 1).setNumberFormat('@');           // 紹介社名
  sh.getRange(5, 4, n, 2).setNumberFormat('yyyy/mm/dd');  // 紹介日・成約日
  sh.getRange(5, 6, n, 2).setNumberFormat('#,##0');       // 金額・報酬

  // 代理店IDはサマリーに並んだID（＝アクションログ登録済み）から選択
  setListValidation_(sh.getRange(5, 1, n, 1),
    sh.getParent().getSheetByName(SH.sum).getRange(`A${SUM_FIRST}:A${SUM_END}`));
}

// ---------------------------------------------------------------- サマリー
function fixSummary_(sh) {
  sh.getCharts().forEach(c => sh.removeChart(c));  // 月次前提のグラフは列構成と合わないため削除
  const lastRow = Math.max(sh.getMaxRows(), SUM_END);
  sh.getRange(5, 1, lastRow - 4, 7).clearContent().setBackground(null).setFontWeight('normal');

  sh.getRange('A1').setValue('サマリー（代理店別）');
  sh.getRange('A2').setValue(
    'すべて自動計算。代理店IDは『アクションログ』から自動で並びます。' +
    '稼働率＝直近の対象期間（『設定』C6、初期値3ヶ月）のうち紹介があった月の割合。');
  sh.getRange('A4:G4').setValues([['代理店ID', '代理店名', '稼働率', '紹介数', '成約数', '成約金額（円）', '代理店報酬（円）']]);

  const L = `'${SH.log}'`, K = `'${SH.kpi}'`;
  const kId = `${K}!$A$5:$A$${KPI_END}`, kIntro = `${K}!$D$5:$D$${KPI_END}`, kClose = `${K}!$E$5:$E$${KPI_END}`;

  // 合計行
  sh.getRange('A5:G5').setFormulas([[
    '合計',
    `=COUNTIF(A${SUM_FIRST}:A${SUM_END},"?*")&"社"`,
    `=IFERROR(AVERAGE(C${SUM_FIRST}:C${SUM_END}),"")`,
    `=SUM(D${SUM_FIRST}:D${SUM_END})`,
    `=SUM(E${SUM_FIRST}:E${SUM_END})`,
    `=SUM(F${SUM_FIRST}:F${SUM_END})`,
    `=SUM(G${SUM_FIRST}:G${SUM_END})`,
  ]]).setBackground('#DDEBF7').setFontWeight('bold');

  // 代理店ID：アクションログから重複なしで自動展開
  sh.getRange(`A${SUM_FIRST}`).setFormula(
    `=IFERROR(UNIQUE(FILTER(${L}!A5:A${LOG_END},${L}!A5:A${LOG_END}<>"")),"")`);

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
      `=IF(${A}="","",SUMIFS(${K}!$F$5:$F$${KPI_END},${kId},${A}))`,
      `=IF(${A}="","",SUMIFS(${K}!$G$5:$G$${KPI_END},${kId},${A}))`,
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
    ['設定', '稼働率の対象期間、期限間近の日数、プルダウンの選択肢を管理。'],
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
    ['③ 紹介を受けたら記録', '『紹介・成約実績』に1行追加（代理店ID・紹介社名・紹介日）。成約したら同じ行に成約日・金額・報酬を追記。'],
    ['④ 毎週月曜に確認', '『アクションログ』で赤の行（予定日なし・期限切れ）→黄の行（期限間近）の順に連絡し、次回アクションを入れる。'],
    ['⑤ 月末に振り返り', '『サマリー』で稼働率が0%（灰色）の代理店を洗い出し、再活性化の打ち手を決める。'],
    ['', null],
    ['■ 自動計算の意味', null],
    ['稼働率', '直近の対象期間（初期値3ヶ月）のうち、紹介があった月の割合。3ヶ月中2ヶ月紹介あり→67%。'],
    ['紹介数・成約数', '『紹介・成約実績』の行数と、そのうち成約日が入っている行数。'],
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
