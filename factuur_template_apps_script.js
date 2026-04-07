/**
 * Freelancer Factuur & Offerte Template — Google Apps Script
 * Nederlandse versie — Dutch Google Sheets (semicolons, Dutch function names)
 *
 * HOE TE GEBRUIKEN:
 *   1. Open een nieuw leeg Google Sheet
 *   2. Klik op Uitbreidingen → Apps Script
 *   3. Verwijder alles en plak dit script
 *   4. Klik op Opslaan en dan op Uitvoeren → createFactuurTemplate
 *   5. Accepteer de machtigingen
 */

// ── Kleuren ───────────────────────────────────────────────────────────────────
var C_NAVY     = '#1B3A6B';
var C_BLUE     = '#2E5499';
var C_LOGO     = '#F4F6F9';
var C_ALT      = '#EBF0FA';
var C_W        = '#FFFFFF';
var C_TIP      = '#FEF9E7';
var C_BETAALD  = '#D5F5E3';
var C_BT_TXT   = '#1E8449';
var C_ONBET    = '#FEF9E7';
var C_ON_TXT   = '#B7770D';
var C_VERLOPEN = '#FADBD8';
var C_VL_TXT   = '#C0392B';
var C_BODY     = '#1A1A2E';
var C_GREY     = '#666666';
var C_PALE     = '#E6EDF8';

// ── Voorbeeldklanten ──────────────────────────────────────────────────────────
var KLANTEN = [
  ['Bakkerij De Mol','Bakkerij De Mol BV','Hoofdstraat 12','1234 AB Amsterdam','Nederland','NL123456789B01','info@demol.nl','020-1234567','Jan de Mol',''],
  ['Studio Roos','Studio Roos','Keizersgracht 45','2500 GH Den Haag','Nederland','NL987654321B01','hello@studioroos.nl','070-9876543','Rosa Bloem',''],
  ['Tech Startup X','Tech Startup X BV','High Tech Campus','5600 AA Eindhoven','Nederland','','info@techx.nl','','Mark Jansen',''],
  ['Freelancer Y','Y Consulting','Bergweg 8','3000 BC Rotterdam','Nederland','','y@consulting.nl','','Yara Smit',''],
  ['Klant Z','Z Productions','Nieuwendijk 100','1012 MR Amsterdam','Nederland','NL555444333B01','z@productions.nl','06-12345678','Zoe Peters','']
];

// ── Voorbeelddiensten ─────────────────────────────────────────────────────────
var DIENSTEN = [
  ['Consultancy (uurtarief)',   95,   0.21, 'uur',  'Strategisch advies per uur'],
  ['Consultancy (dagtarief)',  700,   0.21, 'dag',  'Volledige werkdag'],
  ['Webdesign',              1500,   0.21, 'vast', 'Website ontwerp project'],
  ['Webdevelopment',         2500,   0.21, 'vast', 'Maatwerk webontwikkeling'],
  ['Logo ontwerp',            650,   0.21, 'vast', 'Merklogo ontwerp'],
  ['Copywriting (per pagina)',150,   0.21, 'stuk', 'Geschreven content per pagina'],
  ['Fotografie (halve dag)',  600,   0.21, 'vast', 'Fotosessie halve dag'],
  ['Fotografie (hele dag)',  1100,   0.21, 'vast', 'Fotosessie hele dag'],
  ['Social media beheer',    450,   0.21, 'vast', 'Maandelijks social media beheer'],
  ['SEO Audit',              350,   0.21, 'vast', 'Website SEO auditrapport'],
  ['Training / Workshop',    800,   0.21, 'vast', 'Trainingsessie of workshop'],
  ['Reiskosten',            0.23,     0,  'km',   'Reiskostenvergoeding per km']
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function rep(val, n) {
  var a = [];
  for (var i = 0; i < n; i++) a.push(val);
  return a;
}

// ── Hoofdfunctie ──────────────────────────────────────────────────────────────
function createFactuurTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('Factuur & Offerte Template');

  var tabNamen = ['Factuur','Offerte','Klantenlijst','Diensten & Producten','Factuurlog','Instellingen','Referentie'];
  var sheets = {};

  tabNamen.forEach(function(n) {
    var s = ss.getSheetByName(n);
    if (!s) s = ss.insertSheet(n);
    sheets[n] = s;
  });

  ['Sheet1','Blad1','Blad2','Blad3'].forEach(function(n) {
    var s = ss.getSheetByName(n);
    if (s) try { ss.deleteSheet(s); } catch(e) {}
  });

  sheets['Instellingen'].hideSheet();
  sheets['Referentie'].hideSheet();

  buildInstellingen(sheets['Instellingen']);
  buildReferentie(sheets['Referentie']);
  buildKlantenlijst(sheets['Klantenlijst']);
  buildDiensten(sheets['Diensten & Producten']);
  buildFactuurlog(sheets['Factuurlog']);
  buildDocument(sheets['Factuur'], false);
  buildDocument(sheets['Offerte'], true);

  ss.setActiveSheet(sheets['Factuur']);
  SpreadsheetApp.getUi().alert('✓ Template klaar!\n\nStap 1: Open "Instellingen" en vul je bedrijfsgegevens in.\nStap 2: Open "Klantenlijst" en voeg je klanten toe.\nStap 3: Open "Factuur" en begin met factureren!');
}

// =============================================================================
// INSTELLINGEN (verborgen)
// =============================================================================
function buildInstellingen(sh) {
  sh.getRange(1,1,20,3).breakApart().clearContents().clearFormats();
  sh.getRange(1,1,20,3).setBackground(C_PALE).setFontColor(C_BODY);

  sh.getRange('A1:C1').merge()
    .setValue('INSTELLINGEN — Vul hier je bedrijfsgegevens in')
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);

  var velden = [
    ['Bedrijfsnaam',             'Jouw Bedrijfsnaam'],
    ['Adres',                    'Straat & Huisnummer'],
    ['Stad & Postcode',          '1234 AB Stad'],
    ['BTW Nummer',               'NL000000000B01'],
    ['KVK Nummer',               '12345678'],
    ['E-mail',                   'hallo@jouwbedrijf.nl'],
    ['IBAN',                     'NL00 BANK 0000 0000 00'],
    ['Telefoon',                 '06-00000000'],
    ['Volgend Factuurnummer',    'FACT-2026-001'],
    ['Betalingstermijn (dagen)', 14],
    ['Volgend Offertenummer',    'OFF-2026-001'],
    ['Offerte geldig (dagen)',   30],
    ['Standaard BTW Tarief',     '21%'],
    ['Valutasymbool',            '€']
  ];

  velden.forEach(function(v, i) {
    var r = i + 2;
    sh.getRange(r, 1).setValue(v[0]).setBackground(C_PALE).setFontWeight('bold').setFontColor(C_NAVY).setFontSize(10);
    sh.getRange(r, 2).setValue(v[1]).setBackground(C_W).setFontColor(C_BODY).setFontSize(10);
    sh.setRowHeight(r, 26);
  });

  // Helper: volgend factuurnummer suggestie
  sh.getRange('C10').setFormula('="Volgend: FACT-2026-"&TEKST(WAARDE(RECHTS(B10;3))+1;"000")')
    .setFontColor(C_GREY).setFontStyle('italic').setFontSize(9);

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 220);
  sh.setFrozenRows(1);
}

// =============================================================================
// REFERENTIE (verborgen)
// =============================================================================
function buildReferentie(sh) {
  sh.clearContents();
  sh.getRange('A1:B1').setValues([['BTW Tarief','Label']]);
  sh.getRange('A2:B4').setValues([[0,'0%'],[0.09,'9%'],[0.21,'21%']]);
  sh.getRange('D1:D4').setValues([['Status'],['Onbetaald'],['Betaald'],['Verlopen']]);
}

// =============================================================================
// KLANTENLIJST
// =============================================================================
function buildKlantenlijst(sh) {
  sh.getRange(1,1,200,10).breakApart().clearContents().clearFormats();
  sh.getRange(1,1,200,10).setBackground(C_W).setFontColor(C_BODY);

  sh.getRange('A1:J1').merge()
    .setValue('Voeg hier je klanten toe — deze verschijnen automatisch in je facturen')
    .setBackground(C_PALE).setFontColor(C_BLUE)
    .setFontStyle('italic').setFontSize(10)
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 30);

  var koppen = ['Klantnaam','Bedrijfsnaam','Adres','Stad & Postcode','Land','BTW Nummer','E-mail','Telefoon','Contactpersoon','Notities'];
  sh.getRange(2,1,1,10).setValues([koppen])
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.setRowHeight(2, 30);

  sh.getRange(3,1,KLANTEN.length,10).setValues(KLANTEN);

  for (var i = 0; i < KLANTEN.length; i++) {
    sh.getRange(3+i, 1, 1, 10).setBackground(i % 2 === 0 ? C_W : C_ALT);
    sh.setRowHeight(3+i, 24);
  }

  [140,160,160,150,80,140,180,110,130,120].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(2);
}

// =============================================================================
// DIENSTEN & PRODUCTEN
// =============================================================================
function buildDiensten(sh) {
  sh.getRange(1,1,50,5).breakApart().clearContents().clearFormats();
  sh.getRange(1,1,50,5).setBackground(C_W).setFontColor(C_BODY);

  sh.getRange('A1:E1').merge()
    .setValue('Voeg hier je diensten en prijzen toe — deze verschijnen automatisch in je facturen')
    .setBackground(C_PALE).setFontColor(C_BLUE)
    .setFontStyle('italic').setFontSize(10)
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 30);

  var koppen = ['Dienst / Product','Standaardprijs (excl. BTW)','BTW Tarief','Eenheid','Omschrijving'];
  sh.getRange(2,1,1,5).setValues([koppen])
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.setRowHeight(2, 30);

  sh.getRange(3,1,DIENSTEN.length,5).setValues(DIENSTEN);
  sh.getRange(3,2,DIENSTEN.length,1).setNumberFormat('"€"#.##0,00');
  sh.getRange(3,3,DIENSTEN.length,1).setNumberFormat('0%');

  for (var j = 0; j < DIENSTEN.length; j++) {
    sh.getRange(3+j,1,1,5).setBackground(j % 2 === 0 ? C_W : C_ALT);
    sh.setRowHeight(3+j, 24);
  }

  [200,150,80,80,250].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(2);
}

// =============================================================================
// FACTUURLOG
// =============================================================================
function buildFactuurlog(sh) {
  sh.getRange(1,1,210,8).breakApart().clearContents().clearFormats();
  sh.setConditionalFormatRules([]);
  sh.getRange(1,1,210,8).setBackground(C_W).setFontColor(C_BODY);

  // Titel
  sh.getRange('A1:H1').merge()
    .setValue('FACTUURLOG')
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 45);

  // Samenvattingsrijen
  sh.getRange('A2:H2').merge()
    .setFormula('="📊  Totaal gefactureerd (dit jaar): "&TEKST(SOMMEN.ALS(F8:F200;B8:B200;">="&DATUM(JAAR(VANDAAG());1;1));"€#.##0,00")')
    .setBackground(C_PALE).setFontColor(C_NAVY).setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle');
  sh.setRowHeight(2, 30);

  sh.getRange('A3:H3').merge()
    .setFormula('="✅  Totaal betaald: "&TEKST(SOMMEN.ALS(F8:F200;G8:G200;"Betaald");"€#.##0,00")')
    .setBackground(C_BETAALD).setFontColor(C_BT_TXT).setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle');
  sh.setRowHeight(3, 30);

  sh.getRange('A4:H4').merge()
    .setFormula('="⏳  Openstaand: "&TEKST(SOMMEN.ALS(F8:F200;G8:G200;"Onbetaald")+SOMMEN.ALS(F8:F200;G8:G200;"Verlopen");"€#.##0,00")')
    .setBackground(C_ONBET).setFontColor(C_ON_TXT).setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle');
  sh.setRowHeight(4, 30);

  // Kolomkoppen rij 5
  var koppen = ['Factuurnummer','Datum','Klant','Subtotaal excl. BTW','BTW Totaal','Totaal incl. BTW','Status','Vervaldatum'];
  sh.getRange(5,1,1,8).setValues([koppen])
    .setBackground(C_BLUE).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.setRowHeight(5, 30);

  // Lege spacers rij 6-7
  sh.setRowHeight(6, 8);
  sh.setRowHeight(7, 8);

  // Formattering datakolommen
  sh.getRange('B8:B200').setNumberFormat('DD-MM-JJJJ');
  sh.getRange('H8:H200').setNumberFormat('DD-MM-JJJJ');
  sh.getRange('D8:F200').setNumberFormat('"€"#.##0,00');

  // Status dropdown
  sh.getRange('G8:G200').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Onbetaald','Betaald','Verlopen','Geannuleerd'], true)
      .setAllowInvalid(false).build()
  );

  // Conditionele opmaak op basis van status
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G8="Betaald"')
      .setBackground(C_BETAALD).setFontColor(C_BT_TXT)
      .setRanges([sh.getRange('A8:H200')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G8="Verlopen"')
      .setBackground(C_VERLOPEN).setFontColor(C_VL_TXT)
      .setRanges([sh.getRange('A8:H200')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G8="Onbetaald"')
      .setBackground(C_ONBET).setFontColor(C_ON_TXT)
      .setRanges([sh.getRange('A8:H200')]).build()
  ];
  sh.setConditionalFormatRules(rules);

  [130,100,160,150,110,150,100,120].forEach(function(w,i) { sh.setColumnWidth(i+1,w); });
  sh.setFrozenRows(5);
}

// =============================================================================
// FACTUUR / OFFERTE DOCUMENT
// =============================================================================
function buildDocument(sh, isOfferte) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sh.getRange(1,1,65,8).breakApart().clearContents().clearFormats();
  sh.setConditionalFormatRules([]);
  sh.getRange(1,1,65,8).setBackground(C_W).setFontColor(C_BODY);

  // Kolombreedtes
  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 280);
  sh.setColumnWidth(3, 60);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 80);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 110);

  // ── HEADER (rijen 1-10) ───────────────────────────────────────────────────
  for (var i = 1; i <= 10; i++) sh.setRowHeight(i, 22);
  sh.setRowHeight(1, 30);

  // Logo placeholder
  sh.getRange('A1:B4').merge()
    .setValue('LOGO\nHIER\nINVOEGEN')
    .setBackground(C_LOGO).setFontColor('#AAAAAA')
    .setFontStyle('italic').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Bedrijfsgegevens rechts
  sh.getRange('D1:H1').merge()
    .setFormula('=Instellingen!B2')
    .setFontSize(16).setFontWeight('bold').setFontColor(C_NAVY).setVerticalAlignment('bottom');
  sh.getRange('D2:H2').merge().setFormula('=Instellingen!B3').setFontColor(C_BODY);
  sh.getRange('D3:H3').merge().setFormula('=Instellingen!B4').setFontColor(C_BODY);
  sh.getRange('D4:H4').merge().setFormula('="BTW: "&Instellingen!B5').setFontColor(C_GREY).setFontSize(9);
  sh.getRange('D5:H5').merge().setFormula('="KVK: "&Instellingen!B6').setFontColor(C_GREY).setFontSize(9);
  sh.getRange('D6:H6').merge().setFormula('=Instellingen!B7').setFontColor(C_GREY).setFontSize(9);
  sh.getRange('D7:H7').merge().setFormula('=Instellingen!B8').setFontColor(C_GREY).setFontSize(9);

  // ── TITEL (rij 11) ────────────────────────────────────────────────────────
  sh.setRowHeight(11, 55);
  sh.getRange('A11:H11').setBackground(C_NAVY);
  sh.getRange('A11:D11').merge()
    .setValue(isOfferte ? 'OFFERTE' : 'FACTUUR')
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontSize(28).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Spacer rij 12
  sh.setRowHeight(12, 10);

  // ── META (rijen 13-17) ────────────────────────────────────────────────────
  for (var m = 13; m <= 17; m++) sh.setRowHeight(m, 24);

  var numLabel   = isOfferte ? 'Offertenummer:' : 'Factuurnummer:';
  var numFormula = isOfferte ? '=Instellingen!B12' : '=Instellingen!B10';
  var dateLabel  = isOfferte ? 'Offertedatum:' : 'Factuurdatum:';
  var dueLabel   = isOfferte ? 'Geldig tot:' : 'Vervaldatum:';
  var dueFormula = isOfferte
    ? '=ALS(B14="";"";B14+Instellingen!B13)'
    : '=ALS(B14="";"";B14+Instellingen!B11)';

  sh.getRange('A13').setValue(numLabel).setFontWeight('bold').setFontColor(C_NAVY).setFontSize(9);
  sh.getRange('B13').setFormula(numFormula).setFontWeight('bold').setFontColor(C_NAVY);
  sh.getRange('A14').setValue(dateLabel).setFontWeight('bold').setFontColor(C_NAVY).setFontSize(9);
  sh.getRange('B14').setNumberFormat('DD-MM-JJJJ').setBackground('#FFFDE7');
  sh.getRange('A15').setValue(dueLabel).setFontWeight('bold').setFontColor(C_NAVY).setFontSize(9);
  sh.getRange('B15').setFormula(dueFormula).setNumberFormat('DD-MM-JJJJ');

  if (!isOfferte) {
    sh.getRange('A16').setValue('Status:').setFontWeight('bold').setFontColor(C_NAVY).setFontSize(9);
    sh.getRange('B16').setValue('Onbetaald')
      .setBackground(C_ONBET).setFontColor(C_ON_TXT);
    sh.getRange('B16').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Onbetaald','Betaald','Verlopen'], true)
        .setAllowInvalid(false).build()
    );
  }

  // ── FACTUURADRES (rij 18) ─────────────────────────────────────────────────
  sh.setRowHeight(18, 28);
  sh.getRange('A18:H18').merge()
    .setValue('FACTUURADRES')
    .setBackground(C_NAVY).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // ── KLANTGEGEVENS (rijen 19-24) ───────────────────────────────────────────
  for (var k = 19; k <= 25; k++) sh.setRowHeight(k, 22);

  // Merge cellen VOOR het instellen van formules
  ['A19:D19','A20:D20','A21:D21','A22:D22','A23:D23','A24:D24'].forEach(function(r) {
    sh.getRange(r).merge();
  });

  // Klant dropdown (A19)
  var klantenSheet = ss.getSheetByName('Klantenlijst');
  sh.getRange('A19').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(klantenSheet.getRange('A3:A100'), true)
      .setAllowInvalid(true).build()
  ).setFontWeight('bold').setFontColor(C_NAVY).setFontSize(11)
   .setBackground(C_PALE);

  // VERT.ZOEKEN klantgegevens
  sh.getRange('A20').setFormula('=ALS.FOUT(VERT.ZOEKEN(A19;Klantenlijst!A:B;2;ONWAAR);"")').setFontColor(C_BODY);
  sh.getRange('A21').setFormula('=ALS.FOUT(VERT.ZOEKEN(A19;Klantenlijst!A:C;3;ONWAAR);"")').setFontColor(C_BODY);
  sh.getRange('A22').setFormula('=ALS.FOUT(VERT.ZOEKEN(A19;Klantenlijst!A:D;4;ONWAAR);"")').setFontColor(C_BODY);
  sh.getRange('A23').setFormula('=ALS.FOUT(VERT.ZOEKEN(A19;Klantenlijst!A:E;5;ONWAAR);"")').setFontColor(C_BODY);
  sh.getRange('A24').setFormula('=ALS.FOUT("BTW: "&VERT.ZOEKEN(A19;Klantenlijst!A:F;6;ONWAAR);"")').setFontColor(C_GREY).setFontSize(9);

  sh.setRowHeight(25, 15);
  sh.setRowHeight(26, 5);

  // ── REGELPOSTEN HEADER (rij 27) ───────────────────────────────────────────
  var kolHeaders = ['#','Omschrijving','Aantal','Stukprijs','BTW %','BTW Bedrag','Totaal excl.','Totaal incl.'];
  sh.getRange(27,1,1,8).setValues([kolHeaders])
    .setBackground(C_BLUE).setFontColor(C_W)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.setRowHeight(27, 30);

  // ── REGELPOSTEN (rijen 28-42) ─────────────────────────────────────────────
  var dienstenSheet = ss.getSheetByName('Diensten & Producten');

  for (var li = 0; li < 15; li++) {
    var r = 28 + li;
    var bg = li % 2 === 0 ? C_W : C_ALT;
    sh.getRange(r,1,1,8).setBackground(bg).setFontSize(10);
    sh.setRowHeight(r, 24);

    // Rijnummer
    sh.getRange(r,1).setValue(li+1).setHorizontalAlignment('center')
      .setFontColor(C_GREY).setFontSize(9).setBackground(bg);

    // Omschrijving dropdown (B)
    sh.getRange(r,2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(dienstenSheet.getRange('A3:A50'), true)
        .setAllowInvalid(true).build()
    ).setBackground(bg);

    // Aantal (C) - handmatige invoer
    sh.getRange(r,3).setHorizontalAlignment('center').setBackground(bg);

    // Stukprijs auto-fill via VERT.ZOEKEN (D)
    sh.getRange(r,4).setFormula(
      '=ALS.FOUT(VERT.ZOEKEN(B'+r+';\'Diensten & Producten\'!A:B;2;ONWAAR);"")'
    ).setNumberFormat('"€"#.##0,00').setHorizontalAlignment('right').setBackground(bg);

    // BTW % auto-fill + handmatige override (E)
    sh.getRange(r,5).setFormula(
      '=ALS.FOUT(VERT.ZOEKEN(B'+r+';\'Diensten & Producten\'!A:C;3;ONWAAR);0,21)'
    ).setNumberFormat('0%').setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(r,5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([0, 0.09, 0.21], true)
        .setAllowInvalid(true).build()
    );

    // BTW Bedrag (F)
    sh.getRange(r,6).setFormula(
      '=ALS.FOUT(ALS(C'+r+'="";"";C'+r+'*D'+r+'*E'+r+');"")'
    ).setNumberFormat('"€"#.##0,00').setHorizontalAlignment('right').setBackground(bg);

    // Totaal excl. BTW (G)
    sh.getRange(r,7).setFormula(
      '=ALS.FOUT(ALS(C'+r+'="";"";C'+r+'*D'+r+');"")'
    ).setNumberFormat('"€"#.##0,00').setHorizontalAlignment('right').setBackground(bg);

    // Totaal incl. BTW (H)
    sh.getRange(r,8).setFormula(
      '=ALS.FOUT(ALS(C'+r+'="";"";G'+r+'+F'+r+');"")'
    ).setNumberFormat('"€"#.##0,00').setHorizontalAlignment('right').setBackground(bg);
  }

  // Spacer boven totalen
  sh.setRowHeight(43, 8);

  // ── TOTALEN (rijen 44-50) ─────────────────────────────────────────────────
  var totaalRijen = [
    [44, 'Subtotaal (excl. BTW)', '=SOM(G28:G42)',                     false],
    [45, 'BTW 0%',                '=SOMMEN.ALS(F28:F42;E28:E42;0)',    false],
    [46, 'BTW 9%',                '=SOMMEN.ALS(F28:F42;E28:E42;0,09)', false],
    [47, 'BTW 21%',               '=SOMMEN.ALS(F28:F42;E28:E42;0,21)', false],
    [48, 'Totaal BTW',            '=SOM(F28:F42)',                      false],
    [50, 'TOTAAL TE BETALEN',     '=SOM(H28:H42)',                      true]
  ];

  totaalRijen.forEach(function(t) {
    var row = t[0], label = t[1], formula = t[2], isTotal = t[3];
    var bg  = isTotal ? C_NAVY : C_W;
    var fg  = isTotal ? C_W : C_BODY;
    var sz  = isTotal ? 13 : 10;
    var bld = isTotal ? 'bold' : 'normal';

    sh.setRowHeight(row, isTotal ? 35 : 24);
    sh.getRange(row,1,1,8).setBackground(bg);
    sh.getRange(row,6,1,2).merge()
      .setValue(label)
      .setBackground(bg).setFontColor(fg)
      .setFontSize(sz).setFontWeight(bld)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange(row,8)
      .setFormula(formula)
      .setNumberFormat('"€"#.##0,00')
      .setBackground(bg).setFontColor(fg)
      .setFontSize(sz).setFontWeight(bld)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // Spacer rij 49
  sh.setRowHeight(49, 6);
  sh.getRange('A49:H49').setBackground(C_W);

  // ── VOETTEKST (rijen 52-56) ───────────────────────────────────────────────
  sh.setRowHeight(51, 12);

  sh.getRange('A52').setValue('Betaalgegevens').setFontWeight('bold').setFontColor(C_NAVY).setFontSize(10);
  sh.getRange('A53').setFormula('="IBAN: "&Instellingen!B8').setFontColor(C_BODY).setFontSize(10);

  if (isOfferte) {
    sh.getRange('A54').setFormula('=ALS.FOUT("Deze offerte is geldig tot "&TEKST(B15;"DD-MM-JJJJ")&". Neem contact op om te bevestigen.";"")').setFontColor(C_GREY).setFontSize(9);
  } else {
    sh.getRange('A54').setFormula('="Vermeld bij betaling factuurnummer: "&Instellingen!B10').setFontColor(C_BODY).setFontSize(9);
  }

  sh.getRange('A55').setValue('[Notities / aanvullende informatie]').setFontColor(C_GREY).setFontStyle('italic').setFontSize(9);
  sh.getRange('A56').setFormula('="Bedankt voor uw opdracht! — "&Instellingen!B2').setFontColor(C_GREY).setFontStyle('italic').setFontSize(9);

  for (var fr = 52; fr <= 56; fr++) sh.setRowHeight(fr, 22);

  // Spacer rij 57
  sh.setRowHeight(57, 10);

  // ── TIP BOX (rij 58) ──────────────────────────────────────────────────────
  sh.getRange('A58:H58').merge()
    .setValue('💡  PDF opslaan: Bestand → Downloaden → PDF')
    .setBackground(C_TIP).setFontColor(C_ON_TXT)
    .setFontStyle('italic').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(58, 28);

  // ── RANDEN ────────────────────────────────────────────────────────────────
  sh.getRange('A1:H58').setBorder(
    true, true, true, true, false, false,
    C_NAVY, SpreadsheetApp.BorderStyle.MEDIUM
  );
  sh.getRange('A27:H42').setBorder(
    null, null, null, null, true, true,
    '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID
  );
  sh.getRange('A44:H44').setBorder(
    true, null, null, null, null, null,
    C_NAVY, SpreadsheetApp.BorderStyle.MEDIUM
  );

  // ── CONDITIONELE OPMAAK (alleen Factuur) ─────────────────────────────────
  if (!isOfferte) {
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B$16="Betaald"')
        .setBackground(C_BETAALD).setFontColor(C_BT_TXT)
        .setRanges([sh.getRange('B16')]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B$16="Verlopen"')
        .setBackground(C_VERLOPEN).setFontColor(C_VL_TXT)
        .setRanges([sh.getRange('B16')]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B$16="Onbetaald"')
        .setBackground(C_ONBET).setFontColor(C_ON_TXT)
        .setRanges([sh.getRange('B16')]).build()
    ]);
  }
}
