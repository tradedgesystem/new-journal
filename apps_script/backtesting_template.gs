function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Backtest Template')
    .addItem('Build/Reset Template', 'buildTemplate')
    .addToUi();
}

function buildTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradesSheet = upsertSheet_(ss, 'Trades');
  const summarySheet = upsertSheet_(ss, 'Summary');

  setupTradesSheet_(tradesSheet);
  setupSummarySheet_(summarySheet);
}

function upsertSheet_(spreadsheet, name) {
  const existing = spreadsheet.getSheetByName(name);
  if (existing) {
    existing.clear({ contentsOnly: false });
    existing.clearConditionalFormatRules();
    return existing;
  }
  return spreadsheet.insertSheet(name);
}

function setupTradesSheet_(sheet) {
  const headers = ['Date', 'Ticker', 'Trade #', 'Points', 'Risk', 'Result', 'R Multiple', 'Notes'];
  const maxRows = 500; // adjust if you want more pre-filled rows

  sheet.clearConditionalFormatRules();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#f1f3f4');

  sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 4, maxRows - 1, 2).setNumberFormat('0.00');
  sheet.getRange(2, 7, maxRows - 1, 1).setNumberFormat('0.00');

  const tickerValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['MES', 'M2K', 'MYM', 'MCL'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, maxRows - 1, 1).setDataValidation(tickerValidation);

  const resultValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Green', 'Red', 'Breakeven', 'No Trade'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, maxRows - 1, 1).setDataValidation(resultValidation);

  const tradeNumberValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 2)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(tradeNumberValidation);

  const rFormulaR1C1 = '=IF(OR(RC[-3]="",RC[-2]=""),"",IF(RC[-2]=0,"",RC[-3]/RC[-2]))';
  sheet.getRange(2, 7, maxRows - 1, 1).setFormulaR1C1(rFormulaR1C1);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Green')
      .setBackground('#0f9d58')
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange(2, 6, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Red')
      .setBackground('#d93025')
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange(2, 6, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Breakeven')
      .setBackground('#b0bec5')
      .setFontColor('#000000')
      .setRanges([sheet.getRange(2, 6, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('No Trade')
      .setBackground('#f9ab00')
      .setFontColor('#000000')
      .setRanges([sheet.getRange(2, 6, maxRows - 1, 1)])
      .build(),
  ];
  sheet.setConditionalFormatRules(rules);

  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(8, 220);
}

function setupSummarySheet_(sheet) {
  const rows = [
    ['Metric', 'Value'],
    ['Total Points', '=SUM(Trades!D:D)'],
    ['Total Risk', '=SUM(Trades!E:E)'],
    ['Average Risk', '=IFERROR(AVERAGEIF(Trades!E:E,">0"),"")'],
    ['Overall R/R', '=IF(SUM(Trades!E:E)=0,"",SUM(Trades!D:D)/SUM(Trades!E:E))'],
    ['Average R Multiple', '=IFERROR(AVERAGEIF(Trades!G:G,">0"),"")'],
    ['Wins', '=COUNTIF(Trades!F:F,"Green")'],
    ['Losses', '=COUNTIF(Trades!F:F,"Red")'],
    ['Breakeven', '=COUNTIF(Trades!F:F,"Breakeven")'],
    ['Win/Loss Ratio', '=IF(B8=0,"",B7/B8)'],
    ['Win Rate', '=IF(B12=0,"",B7/B12)'],
    ['Total Trades', '=COUNTIFS(Trades!F:F,"<>",Trades!F:F,"<>No Trade")'],
    ['No Trade Days', '=COUNTIF(Trades!F:F,"No Trade")'],
  ];

  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);

  const tickerHeaderRow = rows.length + 2;
  sheet.getRange(tickerHeaderRow, 1, 1, 3)
    .setValues([['Ticker', 'Trades Taken', 'Total Points']])
    .setFontWeight('bold');

  const tickers = ['MES', 'M2K', 'MYM', 'MCL'];
  tickers.forEach((ticker, index) => {
    const row = tickerHeaderRow + index + 1;
    sheet.getRange(row, 1).setValue(ticker);
    sheet.getRange(row, 2).setFormula(`=COUNTIFS(Trades!B:B,A${row},Trades!F:F,"<>No Trade",Trades!F:F,"<>")`);
    sheet.getRange(row, 3).setFormula(`=SUMIFS(Trades!D:D,Trades!B:B,A${row})`);
  });

  sheet.autoResizeColumns(1, 3);
}
