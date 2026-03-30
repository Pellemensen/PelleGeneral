/**
 * 50/30/20 Budget Tracker — Google Apps Script
 * ─────────────────────────────────────────────
 * HOW TO USE:
 *   1. Open a BRAND NEW blank Google Sheet
 *   2. Click Extensions → Apps Script
 *   3. Delete everything in the editor
 *   4. Paste this entire file
 *   5. Click Save (💾), then click Run ▶  (select createBudgetTracker)
 *   6. Accept permissions when prompted
 *   7. Go back to your sheet — it's done!
 */

// ── Colours ───────────────────────────────────────────────────────────────────
var SAGE      = '#5C8C6E';
var LT_SAGE   = '#7AAD8A';
var ROSE      = '#D98C8C';
var SAND      = '#C4A882';
var MINT      = '#E8F0EB';
var ALT_ROW   = '#F5FAF7';
var OFF_WHITE = '#F9F5F0';
var DK_TEXT   = '#2C3E2D';
var RED_BG    = '#FADBD8';
var RED_TXT   = '#C0392B';
var GRN_BG    = '#D5F5E3';
var GRN_TXT   = '#1E8449';
var YLW_BG    = '#FEF9E7';
var W         = '#FFFFFF';

// ── Categories ────────────────────────────────────────────────────────────────
var INCOME_CATS  = ['Paycheck','Business','Side Hustle','Dividends','Interest Income','Commission','Income 7','Income 8','Income 9','Income 10','Income 11','Income 12'];
var NEEDS_CATS   = ['Internet','Electricity','Water','Mobile','Insurance','Gas','City Garbage','Rent','Groceries','Transportation','Household','Pet'];
var WANTS_CATS   = ['Apparel','Dining Out','Entertainment','Travel','Social Life','Beauty','Gift','Netflix','Amazon Prime','Spotify','Subscription Box','Gym Membership'];
var SAVINGS_CATS = ['Emergency Fund','Retirement','Investment','Credit Card','Student Loan','Car Loan','Mortgage Overpayment','Savings Goal'];
var MONTHS       = ['Jan 2026','Feb 2026','Mar 2026','Apr 2026','May 2026','Jun 2026','Jul 2026','Aug 2026','Sep 2026','Oct 2026','Nov 2026','Dec 2026'];

// ── Sample data ───────────────────────────────────────────────────────────────
var TX = [
  ['2026-01-01','Paycheck',              'Income', 'Needs',         'Paycheck',         1200],
  ['2026-01-01','Business Revenue',      'Income', 'Needs',         'Business',         2350],
  ['2026-01-15','Side Hustle',           'Income', 'Needs',         'Side Hustle',       800],
  ['2026-01-20','Dividend',              'Income', 'Needs',         'Dividends',         175],
  ['2026-01-05','Rent payment',          'Expense','Needs',         'Rent',              725],
  ['2026-01-06','Weekly groceries',      'Expense','Needs',         'Groceries',         204],
  ['2026-01-08','Internet bill',         'Expense','Needs',         'Internet',           40],
  ['2026-01-09','Electricity',           'Expense','Needs',         'Electricity',       200],
  ['2026-01-10','Mobile plan',           'Expense','Needs',         'Mobile',             35],
  ['2026-01-12','Car insurance',         'Expense','Needs',         'Insurance',         200],
  ['2026-01-14','Gas',                   'Expense','Needs',         'Gas',               100],
  ['2026-01-18','Uber to work',          'Expense','Needs',         'Transportation',    140],
  ['2026-01-22','Household supplies',    'Expense','Needs',         'Household',         225],
  ['2026-01-25','Pet food & vet',        'Expense','Needs',         'Pet',               150],
  ['2026-01-03','New jacket',            'Expense','Wants',         'Apparel',            95],
  ['2026-01-07','Restaurant dinner',     'Expense','Wants',         'Dining Out',        100],
  ['2026-01-11','Cinema tickets',        'Expense','Wants',         'Entertainment',     100],
  ['2026-01-15','Flight booking',        'Expense','Wants',         'Travel',            400],
  ['2026-01-17','Bar with friends',      'Expense','Wants',         'Social Life',        90],
  ['2026-01-19','Haircut',               'Expense','Wants',         'Beauty',            135],
  ['2026-01-23','Birthday gift',         'Expense','Wants',         'Gift',               45],
  ['2026-01-25','Netflix',               'Expense','Wants',         'Netflix',            35],
  ['2026-01-26','Amazon Prime',          'Expense','Wants',         'Amazon Prime',       60],
  ['2026-01-28','Spotify',               'Expense','Wants',         'Spotify',            15],
  ['2026-01-29','Mystery Box',           'Expense','Wants',         'Subscription Box',  120],
  ['2026-01-30','Gym',                   'Expense','Wants',         'Gym Membership',    150],
  ['2026-01-02','Emergency fund top-up', 'Expense','Savings & Debt','Emergency Fund',    200],
  ['2026-01-02','Retirement contribution','Expense','Savings & Debt','Retirement',       300],
  ['2026-01-03','Credit card payment',   'Expense','Savings & Debt','Credit Card',       360],
  ['2026-02-01','Paycheck',              'Income', 'Needs',         'Paycheck',         1200],
  ['2026-02-01','Business Revenue',      'Income', 'Needs',         'Business',         2100],
  ['2026-02-05','Rent payment',          'Expense','Needs',         'Rent',              725],
  ['2026-02-07','Weekly groceries',      'Expense','Needs',         'Groceries',         185],
  ['2026-02-10','Restaurant dinner',     'Expense','Wants',         'Dining Out',         80],
  ['2026-02-14','Valentine dinner',      'Expense','Wants',         'Dining Out',        120],
  ['2026-02-02','Retirement contribution','Expense','Savings & Debt','Retirement',       300],
  ['2026-03-01','Paycheck',              'Income', 'Needs',         'Paycheck',         1200],
  ['2026-03-05','Rent payment',          'Expense','Needs',         'Rent',              725],
  ['2026-03-08','Cinema tickets',        'Expense','Wants',         'Entertainment',      65],
];

// ── Helper: repeat value n times ──────────────────────────────────────────────
function rep(val, n) {
  var a = [];
  for (var i = 0; i < n; i++) a.push(val);
  return a;
}

// ── Helper: build alt-row background 2D array ─────────────────────────────────
function altBg(rows, cols) {
  var arr = [];
  for (var i = 0; i < rows; i++) {
    arr.push(rep(i % 2 === 0 ? W : ALT_ROW, cols));
  }
  return arr;
}

// =============================================================================
// MAIN
// =============================================================================

function createBudgetTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('50/30/20 Budget Tracker');

  // Rename first sheet, create others
  var sheets = ss.getSheets();
  sheets[0].setName('Dashboard');

  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }
  var dash = ss.getSheetByName('Dashboard');
  var log  = getOrCreate('Transactions Log');
  var set  = getOrCreate('Settings');
  var ref  = getOrCreate('Reference');

  // Remove any leftover unnamed sheets
  ss.getSheets().forEach(function(s) {
    if (['Dashboard','Transactions Log','Settings','Reference'].indexOf(s.getName()) === -1) {
      ss.deleteSheet(s);
    }
  });

  buildReference(ref);
  buildSettings(set);
  buildLog(log, set);
  buildDashboard(dash);

  ref.hideSheet();
  ss.setActiveSheet(dash);

  SpreadsheetApp.getUi().alert(
    '✓  Budget Tracker is klaar!\n\n' +
    'Selecteer een maand in cel B5 op het Dashboard om te beginnen.'
  );
}

// =============================================================================
// REFERENCE
// =============================================================================

function buildReference(sh) {
  sh.clearContents();
  sh.getRange(1,1,12,1).setValues([['January'],['February'],['March'],['April'],['May'],['June'],['July'],['August'],['September'],['October'],['November'],['December']]);
  sh.getRange(1,2,3,1).setValues([['Needs'],['Wants'],['Savings & Debt']]);
  sh.getRange(1,3,2,1).setValues([['On Track'],['Over Budget']]);
  // Column D: uppercase 3-letter month abbreviations used by MATCH formulas
  sh.getRange(1,4,12,1).setValues([['JAN'],['FEB'],['MAR'],['APR'],['MAY'],['JUN'],['JUL'],['AUG'],['SEP'],['OCT'],['NOV'],['DEC']]);
}

// =============================================================================
// SETTINGS
// =============================================================================

function buildSettings(sh) {
  sh.clearContents();
  sh.clearFormats();
  sh.getRange('A1:B5').clearDataValidations();
  sh.getRange('A7:B60').clearDataValidations();

  // Ratio values
  sh.getRange('A1').setValue('BUDGET RATIOS');
  sh.getRange('A2:B4').setValues([['Needs %',50],['Wants %',30],['Savings & Debt %',20]]);
  sh.getRange('A5').setValue('Ratio check');
  sh.getRange('B5').setFormula('=IF(B2+B3+B4=100,"✓ Ratios OK","⚠ Must sum to 100%")');

  // Category mapping
  sh.getRange('A7:B7').setValues([['Category Name','Bucket']]);
  var mapping = [];
  INCOME_CATS.forEach(function(c)  { mapping.push([c,'Income']); });
  NEEDS_CATS.forEach(function(c)   { mapping.push([c,'Needs']); });
  WANTS_CATS.forEach(function(c)   { mapping.push([c,'Wants']); });
  SAVINGS_CATS.forEach(function(c) { mapping.push([c,'Savings & Debt']); });
  var dataRange = sh.getRange(8, 1, mapping.length, 2);
  dataRange.setValues(mapping);

  // ── Format in bulk ──
  sh.getRange('A1:B1').merge().setBackground(SAGE).setFontColor(W).setFontWeight('bold').setFontSize(11);
  sh.getRange('A7:B7').setBackground(SAGE).setFontColor(W).setFontWeight('bold');

  // Alt rows — single batch call
  sh.getRange(8, 1, mapping.length, 2).setBackgrounds(altBg(mapping.length, 2));

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 130);
  sh.setFrozenRows(1);

  // Ratio CF
  var cfRules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('✓').setBackground(GRN_BG).setFontColor(GRN_TXT).setRanges([sh.getRange('B5')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('⚠').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('B5')]).build(),
  ];
  sh.setConditionalFormatRules(cfRules);
}

// =============================================================================
// TRANSACTIONS LOG
// =============================================================================

function buildLog(sh, settingsSheet) {
  sh.clearContents();
  sh.clearFormats();
  sh.getRange('C2:C200').clearDataValidations();
  sh.getRange('E2:E200').clearDataValidations();

  // Headers
  sh.getRange('A1:I1').setValues([['Date','Description','Type','Bucket','Category','Amount','Month','Year','Notes']]);

  // Static values (no formulas yet)
  var staticVals = TX.map(function(tx) {
    return [tx[0], tx[1], tx[2], '', tx[4], tx[5], '', '', ''];
  });
  sh.getRange(2, 1, staticVals.length, 9).setValues(staticVals);

  // Formulas — set separately so locale doesn't affect them
  var bucketFs = [], monthFs = [], yearFs = [];
  TX.forEach(function(_, i) {
    var r = i + 2;
    bucketFs.push(['=IFERROR(VLOOKUP(E'+r+',Settings!$A$8:$B$55,2,FALSE),"")']);
    monthFs.push(['=IF(A'+r+'="","",MONTH(A'+r+'))']);
    yearFs.push(['=IF(A'+r+'="","",YEAR(A'+r+'))']);
  });
  sh.getRange(2, 4, bucketFs.length, 1).setFormulas(bucketFs);
  sh.getRange(2, 7, monthFs.length,  1).setFormulas(monthFs);
  sh.getRange(2, 8, yearFs.length,   1).setFormulas(yearFs);

  // ── Format in bulk ──
  sh.getRange('A1:I1').setBackground(SAGE).setFontColor(W).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, TX.length, 9).setBackgrounds(altBg(TX.length, 9));  // alt rows in 1 call
  sh.getRange(2, 1, TX.length, 1).setNumberFormat('DD MMM YYYY');
  sh.getRange(2, 6, TX.length, 1).setNumberFormat('$#,##0.00');

  // Column widths
  [100,200,80,110,130,90,60,60,200].forEach(function(w,i){ sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(1);

  // Validation
  sh.getRange('C2:C200').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Income','Expense'],true).setAllowInvalid(false).build()
  );
  sh.getRange('E2:E200').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(settingsSheet.getRange('A8:A55'), true)
      .setAllowInvalid(true).build()
  );

  // Gradient CF on Amount
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().setGradientMaxpoint(SAGE).setGradientMinpoint(GRN_BG).setRanges([sh.getRange('F2:F200')]).build()
  ]);
}

// =============================================================================
// DASHBOARD
// =============================================================================

function buildDashboard(sh) {
  sh.clearContents();
  sh.clearFormats();
  sh.getRange('B5:B5').clearDataValidations();

  // Background
  sh.getRange(1, 1, 60, 24).setBackground(OFF_WHITE);

  // ── Row 1: Title ──
  sh.getRange('A1:J1').merge()
    .setValue('50/30/20 BUDGET DASHBOARD')
    .setBackground(SAGE).setFontColor(W).setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 55);

  // ── Rows 3–5: Controls ──
  sh.getRange(3,1,3,10).setBackground(MINT);
  sh.getRange('A3:D3').setValues([['BUDGET PERIOD','01 Jan 2026','to','31 Jan 2026']]);
  sh.getRange('A4:B4').setValues([['SET CURRENCY','$']]);
  sh.getRange('A5:B5').setValues([['SELECTED MONTH','Jan 2026']]);
  sh.getRange('A3:A5').setFontWeight('bold').setFontColor(DK_TEXT);

  sh.getRange('B5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MONTHS,true).setAllowInvalid(false).build()
  );

  // ── Row 7: KPI ──
  sh.setRowHeight(7, 65);
  sh.getRange(7,2,1,9).setBackground(MINT).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange('B7').setValue('NEEDS').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('E7').setValue('WANTS').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('H7').setValue('SAVINGS & DEBT').setFontWeight('bold').setFontColor(DK_TEXT);

  // KPI formulas — using setFormula to avoid locale issues
  var mf = 'MATCH(UPPER(LEFT($B$5,3)),Reference!$D$1:$D$12,0)';
  var yf = 'VALUE(RIGHT($B$5,4))';
  function kpif(bucket) {
    return '=IFERROR(SUMIFS(\'Transactions Log\'!F:F,\'Transactions Log\'!C:C,"Expense",\'Transactions Log\'!D:D,"'+bucket+'",\'Transactions Log\'!G:G,'+mf+',\'Transactions Log\'!H:H,'+yf+'),0)';
  }
  sh.getRange('C7').setFormula(kpif('Needs')).setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('F7').setFormula(kpif('Wants')).setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('I7').setFormula(kpif('Savings & Debt')).setFontSize(18).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');

  // Helper targets (hidden, row 2)
  sh.getRange('T2').setFormula('=Settings!B2/100*B26');
  sh.getRange('U2').setFormula('=Settings!B3/100*B26');
  sh.getRange('V2').setFormula('=Settings!B4/100*B26');

  // ── Row 9: Section headers ──
  var secHdrFmt = {bg: SAGE, fg: W, bold: true, size: 10};
  sh.getRange('A9:D9').merge().setValue('INCOME SUMMARY').setBackground(SAGE).setFontColor(W).setFontWeight('bold').setFontSize(10);
  sh.getRange('F9:K9').merge().setValue('NEEDS SUMMARY').setBackground(SAGE).setFontColor(W).setFontWeight('bold').setFontSize(10);
  sh.getRange('M9:R9').merge().setValue('WANTS SUMMARY').setBackground(SAGE).setFontColor(W).setFontWeight('bold').setFontSize(10);
  sh.getRange('T9:W9').merge().setValue('AMOUNT LEFT TO SPEND').setBackground(SAGE).setFontColor(W).setFontWeight('bold').setFontSize(10);

  // ── Row 10: Column headers ──
  sh.getRange('A10:D10').setValues([['Category','Expected ($)','Actual ($)','Variance']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F10:K10').setValues([['Category','Due Date','Expected ($)','Actual ($)','Progress','Bar']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('M10:R10').setValues([['Category','Due Date','Expected ($)','Actual ($)','Progress','Bar']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold').setHorizontalAlignment('center');

  // ── Income static values ──
  var incomeStatic = INCOME_CATS.map(function(cat){ return [cat, 0]; });
  sh.getRange(11, 1, 12, 2).setValues(incomeStatic);

  // Income formulas (Actual + Variance) — setFormulas avoids locale issues
  var incActFs = [], incVarFs = [];
  INCOME_CATS.forEach(function(_, i) {
    var r = i + 11;
    incActFs.push(['=IFERROR(SUMIFS(\'Transactions Log\'!F:F,\'Transactions Log\'!E:E,A'+r+',\'Transactions Log\'!C:C,"Income",\'Transactions Log\'!G:G,'+mf+',\'Transactions Log\'!H:H,'+yf+'),0)']);
    incVarFs.push(['=C'+r+'-B'+r]);
  });
  sh.getRange(11, 3, 12, 1).setFormulas(incActFs);
  sh.getRange(11, 4, 12, 1).setFormulas(incVarFs);

  // Income formatting
  sh.getRange(11, 1, 12, 4).setBackgrounds(altBg(12, 4));
  sh.getRange('B11:D23').setNumberFormat('$#,##0.00');

  // Total Income row
  sh.getRange('A23:D23').setValues([['TOTAL INCOME','=SUM(B11:B22)','=SUM(C11:C22)','=SUM(D11:D22)']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold');
  // Rollover
  sh.getRange('A25:B25').setValues([['ROLLOVER', 0]]).setBackground(MINT).setFontWeight('bold').setFontColor(DK_TEXT);
  // Total Budget
  sh.getRange('A26:B26').setValues([['TOTAL BUDGET','=C23+B25']]).setBackground(SAGE).setFontColor(W).setFontWeight('bold');
  sh.getRange('B26').setNumberFormat('$#,##0.00');

  // ── Needs static values ──
  var needsStatic = NEEDS_CATS.map(function(cat){ return [cat,'',0]; });
  sh.getRange(11, 6, 12, 3).setValues(needsStatic);

  // Needs formulas
  var nActFs = [], nPctFs = [], nBarFs = [];
  NEEDS_CATS.forEach(function(_, i) {
    var r = i + 11;
    nActFs.push(['=IFERROR(SUMIFS(\'Transactions Log\'!F:F,\'Transactions Log\'!E:E,F'+r+',\'Transactions Log\'!G:G,'+mf+',\'Transactions Log\'!H:H,'+yf+'),0)']);
    nPctFs.push(['=IFERROR(I'+r+'/H'+r+',0)']);
    nBarFs.push(['=IF(H'+r+'=0,"",SPARKLINE(I'+r+'/H'+r+',{"charttype","bar";"max",1;"color1","#5C8C6E";"color2","#F2DDD8"}))']);
  });
  sh.getRange(11, 9,  12, 1).setFormulas(nActFs);
  sh.getRange(11, 10, 12, 1).setFormulas(nPctFs);
  sh.getRange(11, 11, 12, 1).setFormulas(nBarFs);

  // Needs formatting
  sh.getRange(11, 6, 12, 6).setBackgrounds(altBg(12, 6));
  sh.getRange('H11:I23').setNumberFormat('$#,##0.00');
  sh.getRange('J11:J23').setNumberFormat('0%');

  // Needs total row
  sh.getRange(23, 6, 1, 6).setValues([['TOTAL NEEDS','','=SUM(H11:H22)','=SUM(I11:I22)','=IFERROR(I23/H23,0)','']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold');

  // ── Wants static values ──
  var wantsStatic = WANTS_CATS.map(function(cat){ return [cat,'',0]; });
  sh.getRange(11, 13, 12, 3).setValues(wantsStatic);

  // Wants formulas
  var wActFs = [], wPctFs = [], wBarFs = [];
  WANTS_CATS.forEach(function(_, i) {
    var r = i + 11;
    wActFs.push(['=IFERROR(SUMIFS(\'Transactions Log\'!F:F,\'Transactions Log\'!E:E,M'+r+',\'Transactions Log\'!G:G,'+mf+',\'Transactions Log\'!H:H,'+yf+'),0)']);
    wPctFs.push(['=IFERROR(P'+r+'/O'+r+',0)']);
    wBarFs.push(['=IF(O'+r+'=0,"",SPARKLINE(P'+r+'/O'+r+',{"charttype","bar";"max",1;"color1","#D98C8C";"color2","#F2DDD8"}))']);
  });
  sh.getRange(11, 16, 12, 1).setFormulas(wActFs);
  sh.getRange(11, 17, 12, 1).setFormulas(wPctFs);
  sh.getRange(11, 18, 12, 1).setFormulas(wBarFs);

  // Wants formatting
  sh.getRange(11, 13, 12, 6).setBackgrounds(altBg(12, 6));
  sh.getRange('O11:P23').setNumberFormat('$#,##0.00');
  sh.getRange('Q11:Q23').setNumberFormat('0%');

  // Wants total row
  sh.getRange(23, 13, 1, 6).setValues([['TOTAL WANTS','','=SUM(O11:O22)','=SUM(P11:P22)','=IFERROR(P23/O23,0)','']]).setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold');

  // ── Amount Left panel ──
  sh.getRange('T10:W10').merge().setValue('AMOUNT LEFT').setBackground(LT_SAGE).setFontColor(W).setFontWeight('bold');
  sh.getRange('T11').setFormula('=B26-(SUM(I11:I22)+SUM(P11:P22))').setFontSize(24).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('$#,##0.00');
  sh.getRange('T13').setValue('PERCENTAGE LEFT').setFontWeight('bold').setFontColor(DK_TEXT);
  sh.getRange('T14').setFormula('=IFERROR(T11/B26,0)').setFontSize(14).setFontWeight('bold').setFontColor(DK_TEXT).setNumberFormat('0.00%');
  sh.getRange('T15').setFormula('=IF(T14>0.1,"✓ Doing great! You are right on track.","⚠ Watch your spending!")').setFontWeight('bold').setFontColor(DK_TEXT);

  // ── Column widths ──
  var widths = [160,90,90,80,16,150,70,90,90,60,100,16,150,70,90,90,60,100,16,160,120,120,120];
  widths.forEach(function(w,i){ sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(1);

  // ── Borders ──
  sh.getRange('A9:D23').setBorder(true,true,true,true,true,true,SAGE,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('F9:K23').setBorder(true,true,true,true,true,true,SAGE,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('M9:R23').setBorder(true,true,true,true,true,true,SAGE,SpreadsheetApp.BorderStyle.SOLID);

  // ── Conditional formatting ──
  var rules = [
    // Needs actual: over budget
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(H11<>0,I11>H11)').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('I11:I22')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(H11<>0,I11<=H11)').setBackground(GRN_BG).setFontColor(GRN_TXT).setRanges([sh.getRange('I11:I22')]).build(),
    // Wants actual: over budget
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(O11<>0,P11>O11)').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('P11:P22')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(O11<>0,P11<=O11)').setBackground(GRN_BG).setFontColor(GRN_TXT).setRanges([sh.getRange('P11:P22')]).build(),
    // Progress %
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1).setBackground(RED_BG).setRanges([sh.getRange('J11:J22')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(J11>=0.75,J11<=1)').setBackground(YLW_BG).setRanges([sh.getRange('J11:J22')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.75).setBackground(GRN_BG).setRanges([sh.getRange('J11:J22')]).build(),
    // KPI vs targets
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=C7>$T$2').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('C7')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=F7>$U$2').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('F7')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=I7>$V$2').setBackground(RED_BG).setFontColor(RED_TXT).setRanges([sh.getRange('I7')]).build(),
  ];
  sh.setConditionalFormatRules(rules);

  // ── Charts ──
  var chart1 = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('A10:C22'))
    .setPosition(28, 1, 0, 0)
    .setOption('title','CASH FLOW SUMMARY')
    .setOption('legend',{position:'top'})
    .setOption('series',{0:{color:'#A8C5B0',labelInLegend:'Expected'},1:{color:SAGE,labelInLegend:'Actual'}})
    .setOption('width',380).setOption('height',260)
    .build();
  sh.insertChart(chart1);

  var chart2 = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('B7:I7'))
    .setPosition(28, 6, 0, 0)
    .setOption('title','ACTUAL ALLOCATION')
    .setOption('pieHole',0.5)
    .setOption('legend',{position:'right'})
    .setOption('slices',{0:{color:SAGE},1:{color:ROSE},2:{color:SAND}})
    .setOption('width',300).setOption('height',260)
    .build();
  sh.insertChart(chart2);
}
