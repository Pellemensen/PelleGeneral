/**
 * 50/30/20 Budget Tracker — Google Apps Script
 * ─────────────────────────────────────────────
 * HOW TO USE:
 *   1. Open a blank Google Sheet
 *   2. Click Extensions → Apps Script
 *   3. Delete everything in the editor
 *   4. Paste this entire file
 *   5. Click Save (💾), then click Run ▶
 *   6. Accept permissions when prompted
 *   7. Go back to your sheet — it's done!
 */

// ── Colours ──────────────────────────────────────────────────────────────────
var SAGE       = '#5C8C6E';
var LT_SAGE    = '#7AAD8A';
var ROSE       = '#D98C8C';
var SAND       = '#C4A882';
var MINT       = '#E8F0EB';
var ALT_ROW    = '#F5FAF7';
var OFF_WHITE  = '#F9F5F0';
var DK_TEXT    = '#2C3E2D';
var RED_BG     = '#FADBD8';
var RED_TXT    = '#C0392B';
var GRN_BG     = '#D5F5E3';
var GRN_TXT    = '#1E8449';
var YLW_BG     = '#FEF9E7';
var GREY_BDR   = '#BDC3C7';

// ── Categories ───────────────────────────────────────────────────────────────
var NEEDS_CATS   = ['Internet','Electricity','Water','Mobile','Insurance','Gas','City Garbage','Rent','Groceries','Transportation','Household','Pet'];
var WANTS_CATS   = ['Apparel','Dining Out','Entertainment','Travel','Social Life','Beauty','Gift','Netflix','Amazon Prime','Spotify','Subscription Box','Gym Membership'];
var SAVINGS_CATS = ['Emergency Fund','Retirement','Investment','Credit Card','Student Loan','Car Loan','Mortgage Overpayment','Savings Goal'];
var INCOME_CATS  = ['Paycheck','Business','Side Hustle','Dividends','Interest Income','Commission','Income 7','Income 8','Income 9','Income 10','Income 11','Income 12'];
var MONTHS       = ['Jan 2026','Feb 2026','Mar 2026','Apr 2026','May 2026','Jun 2026','Jul 2026','Aug 2026','Sep 2026','Oct 2026','Nov 2026','Dec 2026'];

// ── Sample data ───────────────────────────────────────────────────────────────
var TX = [
  ['2026-01-01','Paycheck',             'Income', 'Needs',          'Paycheck',         1200],
  ['2026-01-01','Business Revenue',     'Income', 'Needs',          'Business',         2350],
  ['2026-01-15','Side Hustle',          'Income', 'Needs',          'Side Hustle',       800],
  ['2026-01-20','Dividend',             'Income', 'Needs',          'Dividends',         175],
  ['2026-01-05','Rent payment',         'Expense','Needs',          'Rent',              725],
  ['2026-01-06','Weekly groceries',     'Expense','Needs',          'Groceries',         204],
  ['2026-01-08','Internet bill',        'Expense','Needs',          'Internet',           40],
  ['2026-01-09','Electricity',          'Expense','Needs',          'Electricity',       200],
  ['2026-01-10','Mobile plan',          'Expense','Needs',          'Mobile',             35],
  ['2026-01-12','Car insurance',        'Expense','Needs',          'Insurance',         200],
  ['2026-01-14','Gas',                  'Expense','Needs',          'Gas',               100],
  ['2026-01-18','Uber to work',         'Expense','Needs',          'Transportation',    140],
  ['2026-01-22','Household supplies',   'Expense','Needs',          'Household',         225],
  ['2026-01-25','Pet food & vet',       'Expense','Needs',          'Pet',               150],
  ['2026-01-03','New jacket',           'Expense','Wants',          'Apparel',            95],
  ['2026-01-07','Restaurant dinner',    'Expense','Wants',          'Dining Out',        100],
  ['2026-01-11','Cinema tickets',       'Expense','Wants',          'Entertainment',     100],
  ['2026-01-15','Flight booking',       'Expense','Wants',          'Travel',            400],
  ['2026-01-17','Bar with friends',     'Expense','Wants',          'Social Life',        90],
  ['2026-01-19','Haircut',              'Expense','Wants',          'Beauty',            135],
  ['2026-01-23','Birthday gift',        'Expense','Wants',          'Gift',               45],
  ['2026-01-25','Netflix',              'Expense','Wants',          'Netflix',            35],
  ['2026-01-26','Amazon Prime',         'Expense','Wants',          'Amazon Prime',       60],
  ['2026-01-28','Spotify',              'Expense','Wants',          'Spotify',            15],
  ['2026-01-29','Mystery Box',          'Expense','Wants',          'Subscription Box',  120],
  ['2026-01-30','Gym',                  'Expense','Wants',          'Gym Membership',    150],
  ['2026-01-02','Emergency fund top-up','Expense','Savings & Debt', 'Emergency Fund',    200],
  ['2026-01-02','Retirement contribution','Expense','Savings & Debt','Retirement',       300],
  ['2026-01-03','Credit card payment',  'Expense','Savings & Debt', 'Credit Card',       360],
  ['2026-02-01','Paycheck',             'Income', 'Needs',          'Paycheck',         1200],
  ['2026-02-01','Business Revenue',     'Income', 'Needs',          'Business',         2100],
  ['2026-02-05','Rent payment',         'Expense','Needs',          'Rent',              725],
  ['2026-02-07','Weekly groceries',     'Expense','Needs',          'Groceries',         185],
  ['2026-02-10','Restaurant dinner',    'Expense','Wants',          'Dining Out',         80],
  ['2026-02-14','Valentine dinner',     'Expense','Wants',          'Dining Out',        120],
  ['2026-02-02','Retirement contribution','Expense','Savings & Debt','Retirement',       300],
  ['2026-03-01','Paycheck',             'Income', 'Needs',          'Paycheck',         1200],
  ['2026-03-05','Rent payment',         'Expense','Needs',          'Rent',              725],
  ['2026-03-08','Cinema tickets',       'Expense','Wants',          'Entertainment',      65],
];

// =============================================================================
// MAIN — run this
// =============================================================================

function createBudgetTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('50/30/20 Budget Tracker');

  // ── Create / rename sheets ──
  var existing = ss.getSheets();
  existing[0].setName('Dashboard');

  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }
  var dash = ss.getSheetByName('Dashboard');
  var log  = getOrCreate('Transactions Log');
  var set  = getOrCreate('Settings');
  var ref  = getOrCreate('Reference');

  // Remove any unexpected extra sheets
  ss.getSheets().forEach(function(s) {
    var keep = ['Dashboard','Transactions Log','Settings','Reference'];
    if (keep.indexOf(s.getName()) === -1) ss.deleteSheet(s);
  });

  // ── Build each tab ──
  buildReference(ref);   SpreadsheetApp.flush();
  buildSettings(set);    SpreadsheetApp.flush();
  buildLog(log);         SpreadsheetApp.flush();
  buildDashboard(dash);  SpreadsheetApp.flush();

  ref.hideSheet();
  ss.setActiveSheet(dash);

  SpreadsheetApp.getUi().alert(
    '✓  Budget Tracker is ready!\n\n' +
    'Select a month in cell B5 on the Dashboard to get started.'
  );
}

// =============================================================================
// REFERENCE TAB
// =============================================================================

function buildReference(sh) {
  sh.clearContents();
  sh.getRange(1,1,12,1).setValues(
    [['January'],['February'],['March'],['April'],['May'],['June'],
     ['July'],['August'],['September'],['October'],['November'],['December']]
  );
  sh.getRange(1,2,3,1).setValues([['Needs'],['Wants'],['Savings & Debt']]);
  sh.getRange(1,3,2,1).setValues([['On Track'],['Over Budget']]);
}

// =============================================================================
// SETTINGS TAB
// =============================================================================

function buildSettings(sh) {
  sh.clearContents();
  sh.clearFormats();

  // ── Ratio section ──
  sh.getRange('A1').setValue('BUDGET RATIOS');
  sh.getRange('A2:B4').setValues([['Needs %',50],['Wants %',30],['Savings & Debt %',20]]);
  sh.getRange('A5').setValue('Ratio check');
  sh.getRange('B5').setFormula('=IF(B2+B3+B4=100,"✓ Ratios OK","⚠ Must sum to 100%")');

  // ── Category mapping header ──
  sh.getRange('A7:B7').setValues([['Category Name','Bucket']]);

  // ── Category rows (income first so they appear in the dropdown) ──
  var mapping = [];
  INCOME_CATS.forEach(function(c)  { mapping.push([c,'Income']); });
  NEEDS_CATS.forEach(function(c)   { mapping.push([c,'Needs']); });
  WANTS_CATS.forEach(function(c)   { mapping.push([c,'Wants']); });
  SAVINGS_CATS.forEach(function(c) { mapping.push([c,'Savings & Debt']); });
  sh.getRange(8, 1, mapping.length, 2).setValues(mapping);

  // ── Formatting ──
  // Title
  sh.getRange('A1:B1').merge().setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  // Mapping header
  sh.getRange('A7:B7').setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold');
  // Alt rows
  for (var i = 0; i < mapping.length; i++) {
    sh.getRange(8+i, 1, 1, 2).setBackground(i % 2 === 0 ? '#FFFFFF' : ALT_ROW);
  }
  // Outer border around mapping table
  sh.getRange(7, 1, mapping.length+1, 2)
    .setBorder(true,true,true,true,true,true, SAGE, SpreadsheetApp.BorderStyle.SOLID);

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 130);
  sh.setFrozenRows(1);

  // Ratio validation CF
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✓')
    .setBackground(GRN_BG).setFontColor(GRN_TXT)
    .setRanges([sh.getRange('B5')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('B5')]).build());
  sh.setConditionalFormatRules(rules);
}

// =============================================================================
// TRANSACTIONS LOG TAB
// =============================================================================

function buildLog(sh) {
  sh.clearContents();
  sh.clearFormats();

  // ── Headers ──
  sh.getRange('A1:I1').setValues([['Date','Description','Type','Bucket','Category','Amount','Month','Year','Notes']]);

  // ── Sample data ──
  var rows = TX.map(function(tx, i) {
    var r = i + 2;
    return [
      tx[0], tx[1], tx[2],
      '=IFERROR(VLOOKUP(E'+r+',Settings!$A$8:$B$55,2,FALSE),"")',
      tx[4], tx[5],
      '=IF(A'+r+'="","",MONTH(A'+r+'))',
      '=IF(A'+r+'="","",YEAR(A'+r+'))',
      ''
    ];
  });
  sh.getRange(2, 1, rows.length, 9).setValues(rows);

  // ── Formats ──
  sh.getRange('A1:I1')
    .setBackground(SAGE).setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange(2, 1, rows.length, 1).setNumberFormat('DD MMM YYYY');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('$#,##0.00');

  // Alt rows
  for (var i = 0; i < rows.length; i++) {
    sh.getRange(i+2, 1, 1, 9).setBackground(i % 2 === 0 ? '#FFFFFF' : ALT_ROW);
  }

  // Column widths
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,200); sh.setColumnWidth(3,80);
  sh.setColumnWidth(4,110); sh.setColumnWidth(5,130); sh.setColumnWidth(6,90);
  sh.setColumnWidth(7,60);  sh.setColumnWidth(8,60);  sh.setColumnWidth(9,200);
  sh.setFrozenRows(1);

  // ── Data validation ──
  // Type dropdown
  sh.getRange(2,3,199,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Income','Expense'], true)
      .setAllowInvalid(false).build()
  );
  // Category dropdown from Settings (warning mode so script can set income categories)
  sh.getRange(2,5,199,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getRange('A8:A55'),
        true
      ).setAllowInvalid(true).build()
  );

  // ── Conditional formatting ──
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint(SAGE)
    .setGradientMinpoint(GRN_BG)
    .setRanges([sh.getRange('F2:F200')]).build());
  sh.setConditionalFormatRules(rules);
}

// =============================================================================
// DASHBOARD TAB
// =============================================================================

function buildDashboard(sh) {
  sh.clearContents();
  sh.clearFormats();

  // ── Background ──
  sh.getRange(1, 1, 80, 25).setBackground(OFF_WHITE);

  // ── Row 1: Title ──
  sh.getRange('A1:J1').merge()
    .setValue('50/30/20 BUDGET DASHBOARD')
    .setBackground(SAGE).setFontColor('#FFFFFF')
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 55);

  // ── Rows 3–5: Controls ──
  sh.getRange(3,1,3,10).setBackground(MINT);
  sh.getRange('A3').setValue('BUDGET PERIOD');
  sh.getRange('B3').setValue('01 Jan 2026');
  sh.getRange('C3').setValue('to');
  sh.getRange('D3').setValue('31 Jan 2026');
  sh.getRange('A4').setValue('SET CURRENCY');
  sh.getRange('B4').setValue('$');
  sh.getRange('A5').setValue('SELECTED MONTH');
  sh.getRange('B5').setValue('Jan 2026');

  sh.getRange('A3:A5').setFontWeight('bold').setFontColor(DK_TEXT);

  // Month selector validation
  sh.getRange('B5').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(MONTHS, true)
      .setAllowInvalid(false).build()
  );

  // ── Row 7: KPI ──
  sh.setRowHeight(7, 65);
  sh.getRange(7,2,1,9).setBackground(MINT);
  var kpiLabels = [['NEEDS','','','WANTS','','','SAVINGS & DEBT','','']];
  sh.getRange('B7:J7').setValues(kpiLabels);
  sh.getRange('B7').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('F7').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('I7').setFontWeight('bold').setFontColor(DK_TEXT);

  // Helper cells: T3 = month number, U3 = year number (locale-safe, no DATEVALUE)
  sh.getRange('T3').setFormula('=MATCH(LEFT(B5,3),{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},0)');
  sh.getRange('U3').setFormula('=VALUE(RIGHT(B5,4))');
  var mf = '$T$3';
  var yf = '$U$3';
  var base = "=IFERROR(SUMIFS('Transactions Log'!F:F,'Transactions Log'!C:C,\"Expense\",'Transactions Log'!D:D,\"{B}\",'Transactions Log'!G:G,"+mf+",'Transactions Log'!H:H,"+yf+"),0)";

  sh.getRange('C7').setFormula(base.replace('{B}','Needs'));
  sh.getRange('F7').setFormula(base.replace('{B}','Wants'));
  sh.getRange('I7').setFormula(base.replace('{B}','Savings & Debt'));
  sh.getRange('C7:C7').setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('F7:F7').setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('I7:I7').setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange(7,2,1,9).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Row 9: Section headers ──
  sh.getRange('A9:D9').merge().setValue('INCOME SUMMARY')
    .setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sh.getRange('F9:K9').merge().setValue('NEEDS SUMMARY')
    .setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sh.getRange('M9:R9').merge().setValue('WANTS SUMMARY')
    .setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sh.getRange('T9:W9').merge().setValue('AMOUNT LEFT TO SPEND')
    .setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);

  // ── Row 10: Column headers ──
  sh.getRange('A10:D10').setValues([['Category','Expected ($)','Actual ($)','Variance']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F10:K10').setValues([['Category','Due Date','Expected ($)','Actual ($)','Progress','Bar']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('M10:R10').setValues([['Category','Due Date','Expected ($)','Actual ($)','Progress','Bar']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  // ── Rows 11–22: Income data ──
  var incomeData = INCOME_CATS.map(function(cat, i) {
    var r = i + 11;
    var act = "=IFERROR(SUMIFS('Transactions Log'!F:F,'Transactions Log'!E:E,A"+r+",'Transactions Log'!C:C,\"Income\",'Transactions Log'!G:G,"+mf+",'Transactions Log'!H:H,"+yf+"),0)";
    return [cat, 0, act, '=C'+r+'-B'+r];
  });
  sh.getRange(11, 1, incomeData.length, 4).setValues(incomeData);

  // ── Row 23: TOTAL INCOME ──
  sh.getRange('A23:D23').setValues([['TOTAL INCOME','=SUM(B11:B22)','=SUM(C11:C22)','=SUM(D11:D22)']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold');

  // ── Row 25: ROLLOVER ──
  sh.getRange('A25').setValue('ROLLOVER');
  sh.getRange('B25').setValue(0);
  sh.getRange(25,1,1,4).setBackground(MINT).setFontWeight('bold').setFontColor(DK_TEXT);

  // ── Row 26: TOTAL BUDGET ──
  sh.getRange('A26:C26').setValues([['TOTAL BUDGET','=C23+B25','']])
    .setBackground(SAGE).setFontColor('#FFFFFF').setFontWeight('bold');

  // Format income money columns
  sh.getRange('B11:D23').setNumberFormat('$#,##0.00');
  sh.getRange('B26').setNumberFormat('$#,##0.00');

  // Alt rows income
  for (var i = 0; i < 12; i++) {
    sh.getRange(11+i, 1, 1, 4).setBackground(i % 2 === 0 ? '#FFFFFF' : ALT_ROW);
  }

  // ── Rows 11–22: Needs table (cols F–K = 6–11) ──
  var needsData = NEEDS_CATS.map(function(cat, i) {
    var r = i + 11;
    var act = "=IFERROR(SUMIFS('Transactions Log'!F:F,'Transactions Log'!E:E,F"+r+",'Transactions Log'!G:G,"+mf+",'Transactions Log'!H:H,"+yf+"),0)";
    var pct = '=IFERROR(I'+r+'/H'+r+',0)';
    var bar = '=IF(H'+r+'=0,"",SPARKLINE(I'+r+'/H'+r+',{"charttype","bar";"max",1;"color1","#5C8C6E";"color2","#F2DDD8"}))';
    return [cat, '', 0, act, pct, bar];
  });
  sh.getRange(11, 6, needsData.length, 6).setValues(needsData);
  sh.getRange(23, 6, 1, 6).setValues([['TOTAL NEEDS','','=SUM(H11:H22)','=SUM(I11:I22)','=IFERROR(I23/H23,0)','']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold');

  sh.getRange('H11:I23').setNumberFormat('$#,##0.00');
  sh.getRange('J11:J23').setNumberFormat('0%');
  for (var i = 0; i < 12; i++) {
    sh.getRange(11+i, 6, 1, 6).setBackground(i % 2 === 0 ? '#FFFFFF' : ALT_ROW);
  }

  // ── Rows 11–22: Wants table (cols M–R = 13–18) ──
  var wantsData = WANTS_CATS.map(function(cat, i) {
    var r = i + 11;
    var act = "=IFERROR(SUMIFS('Transactions Log'!F:F,'Transactions Log'!E:E,M"+r+",'Transactions Log'!G:G,"+mf+",'Transactions Log'!H:H,"+yf+"),0)";
    var pct = '=IFERROR(P'+r+'/O'+r+',0)';
    var bar = '=IF(O'+r+'=0,"",SPARKLINE(P'+r+'/O'+r+',{"charttype","bar";"max",1;"color1","#D98C8C";"color2","#F2DDD8"}))';
    return [cat, '', 0, act, pct, bar];
  });
  sh.getRange(11, 13, wantsData.length, 6).setValues(wantsData);
  sh.getRange(23, 13, 1, 6).setValues([['TOTAL WANTS','','=SUM(O11:O22)','=SUM(P11:P22)','=IFERROR(P23/O23,0)','']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold');

  sh.getRange('O11:P23').setNumberFormat('$#,##0.00');
  sh.getRange('Q11:Q23').setNumberFormat('0%');
  for (var i = 0; i < 12; i++) {
    sh.getRange(11+i, 13, 1, 6).setBackground(i % 2 === 0 ? '#FFFFFF' : ALT_ROW);
  }

  // ── Amount Left panel (cols T–W = 20–23) ──
  sh.getRange('T11').setFormula('=B26-(SUM(I11:I23)+SUM(P11:P23))')
    .setFontSize(26).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('T13').setValue('PERCENTAGE LEFT').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('T14').setFormula('=IFERROR(T11/B26,0)')
    .setFontSize(14).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('0.00%');
  sh.getRange('T15').setFormula('=IF(T14>0.1,"✓ Doing great! You are right on track.","⚠ Watch your spending!")')
    .setFontWeight('bold').setFontColor(DK_TEXT).setFontSize(10);
  sh.getRange(10,20,1,4).setValues([['AMOUNT LEFT','','','']])
    .setBackground(LT_SAGE).setFontColor('#FFFFFF').setFontWeight('bold');

  // Helper cells for KPI targets (row 2, hidden)
  sh.getRange('T2').setFormula('=Settings!B2/100*B26');  // Needs target
  sh.getRange('U2').setFormula('=Settings!B3/100*B26');  // Wants target
  sh.getRange('V2').setFormula('=Settings!B4/100*B26');  // S&D target

  // ── Column widths ──
  sh.setColumnWidth(1,160);  // A
  sh.setColumnWidth(2,90);   // B
  sh.setColumnWidth(3,90);   // C
  sh.setColumnWidth(4,80);   // D
  sh.setColumnWidth(5,16);   // E spacer
  sh.setColumnWidth(6,150);  // F
  sh.setColumnWidth(7,70);   // G
  sh.setColumnWidth(8,90);   // H
  sh.setColumnWidth(9,90);   // I
  sh.setColumnWidth(10,60);  // J
  sh.setColumnWidth(11,100); // K
  sh.setColumnWidth(12,16);  // L spacer
  sh.setColumnWidth(13,150); // M
  sh.setColumnWidth(14,70);  // N
  sh.setColumnWidth(15,90);  // O
  sh.setColumnWidth(16,90);  // P
  sh.setColumnWidth(17,60);  // Q
  sh.setColumnWidth(18,100); // R
  sh.setColumnWidth(19,16);  // S spacer
  sh.setColumnWidth(20,160); // T
  sh.setColumnWidth(21,120); // U
  sh.setColumnWidth(22,120); // V
  sh.setColumnWidth(23,120); // W

  sh.setFrozenRows(1);

  // ── Borders ──
  sh.getRange('A9:D23').setBorder(true,true,true,true,true,true, SAGE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('F9:K23').setBorder(true,true,true,true,true,true, SAGE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('M9:R23').setBorder(true,true,true,true,true,true, SAGE, SpreadsheetApp.BorderStyle.SOLID);

  // ── Conditional formatting ──
  var rules = [];

  // Needs Actual (col I) over budget
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(H11<>0,I11>H11)')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('I11:I22')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(H11<>0,I11<=H11)')
    .setBackground(GRN_BG).setFontColor(GRN_TXT)
    .setRanges([sh.getRange('I11:I22')]).build());

  // Wants Actual (col P) over budget
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(O11<>0,P11>O11)')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('P11:P22')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(O11<>0,P11<=O11)')
    .setBackground(GRN_BG).setFontColor(GRN_TXT)
    .setRanges([sh.getRange('P11:P22')]).build());

  // Progress % — Needs col J
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1)
    .setBackground(RED_BG)
    .setRanges([sh.getRange('J11:J22')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(J11>=0.75,J11<=1)')
    .setBackground(YLW_BG)
    .setRanges([sh.getRange('J11:J22')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.75)
    .setBackground(GRN_BG)
    .setRanges([sh.getRange('J11:J22')]).build());

  // KPI vs targets
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=C7>$T$2')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('C7')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=F7>$U$2')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('F7')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=I7>$V$2')
    .setBackground(RED_BG).setFontColor(RED_TXT)
    .setRanges([sh.getRange('I7')]).build());

  sh.setConditionalFormatRules(rules);

  // ── Charts ──

  // Chart 1 — Cash Flow Bar (Income expected vs actual)
  var chart1 = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('A10:C22'))
    .setPosition(28, 1, 0, 0)
    .setOption('title', 'CASH FLOW SUMMARY')
    .setOption('legend', {position: 'top'})
    .setOption('series', {
      0: {color: '#A8C5B0', labelInLegend: 'Expected'},
      1: {color: SAGE,      labelInLegend: 'Actual'}
    })
    .setOption('width', 380)
    .setOption('height', 260)
    .build();
  sh.insertChart(chart1);

  // Chart 2 — Donut Allocation
  var chart2 = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('B7:I7'))
    .setPosition(28, 6, 0, 0)
    .setOption('title', 'ACTUAL ALLOCATION')
    .setOption('pieHole', 0.5)
    .setOption('legend', {position: 'right'})
    .setOption('slices', {
      0: {color: SAGE},
      1: {color: ROSE},
      2: {color: SAND}
    })
    .setOption('width', 300)
    .setOption('height', 260)
    .build();
  sh.insertChart(chart2);
}
