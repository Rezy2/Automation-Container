function setupSchema() {
  const ss = getSpreadsheet_();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  const data = ensureSheet_(ss, SHEET.DATA);
  assertCompatible_(data, DATA_HEADERS);
  TX_SHEETS.forEach(function (name) {
    assertCompatible_(ensureSheet_(ss, name), TX_HEADERS);
  });

  setupDataSheet_(data);
  ss.setActiveSheet(data);

  TX_SHEETS.forEach(function (name) {
    setupTransactionSheet_(ensureSheet_(ss, name));
  });

  setupInternalSheets_(ss);
  cleanupLegacyOpenWA_();
  applyBalanceFormulas_(data, 2, data.getLastRow());
  syncCodeSequences_();
  ensureMenuTrigger_(ss);
  SpreadsheetApp.flush();

  const message = 'Setup selesai. Jalankan seedDummy() untuk mengisi data contoh.';
  Logger.log(message);
  return message;
}

function repairFormulas() {
  const ss = ensureReady_();
  const data = ss.getSheetByName(SHEET.DATA);
  applyBalanceFormulas_(data, 2, data.getLastRow());
  syncCodeSequences_();
  SpreadsheetApp.flush();
  const message = 'Formula saldo dan sequence kode diperbaiki.';
  Logger.log(message);
  return message;
}

function buildMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory LCI')
    .addItem('Setup Schema', 'setupSchema')
    .addItem('Seed Data Dummy', 'seedDummy')
    .addItem('Hapus Data Dummy', 'clearDummy')
    .addSeparator()
    .addItem('Perbaiki Formula Saldo', 'repairFormulas')
    .addItem('Buka Config', 'showConfig')
    .addSeparator()
    .addItem('Fonnte Status', 'fonnteStatus')
    .addItem('Test Kirim Fonnte', 'testFonnteSend')
    .addItem('Proses Outbox', 'processOutbox')
    .addItem('Pasang Trigger Outbox', 'installFonnteOutboxTrigger')
    .addToUi();
}

function showConfig() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET.CONFIG);
  if (!sheet) throw new Error('Tab Config belum ada. Jalankan setupSchema().');
  if (sheet.isSheetHidden()) sheet.showSheet();
  ss.setActiveSheet(sheet);
}

function assertCompatible_(sheet, headers) {
  if (sheet.getLastRow() < 2) return;
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const same = headers.every(function (header, index) {
    return String(current[index]) === String(header);
  });
  if (!same) {
    throw new Error(
      'Tab ' +
        sheet.getName() +
        ' sudah berisi data dengan header lama. Buat salinan spreadsheet atau kosongkan data sebelum setup.'
    );
  }
}

function setupDataSheet_(sheet) {
  assertCompatible_(sheet, DATA_HEADERS);
  writeHeaders_(sheet, DATA_HEADERS, '#1f4e78');
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.setTabColor('#1f4e78');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 38);
  sheet.hideColumns(DATA_COL.IS_DUMMY, 1);

  setWidths_(sheet, {
    1: 50,
    2: 120,
    3: 120,
    4: 230,
    5: 80,
    6: 180,
    7: 110,
    8: 80,
    9: 85,
    10: 100,
    11: 105,
    12: 95,
    13: 100,
    14: 105,
    15: 95,
    16: 95,
    17: 95,
    18: 90,
    19: 80
  });

  const lastRow = Math.max(sheet.getMaxRows(), 2);
  sheet
    .getRange(2, 1, lastRow - 1, DATA_HEADERS.length)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.getRange(2, DATA_COL.MIN, lastRow - 1, 7).setNumberFormat('#,##0.###');

  sheet
    .getRange('E2:E')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['F', 'M', 'C', 'T'], true)
        .setAllowInvalid(false)
        .build()
    );

  sheet
    .getRange('R2:R')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['ACTIVE', 'INACTIVE'], true)
        .setAllowInvalid(false)
        .build()
    );

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setRanges([sheet.getRange('L2:L'), sheet.getRange('O2:O')])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('YA')
      .setBackground('#fff2cc')
      .setFontColor('#7f6000')
      .setRanges([sheet.getRange('P2:P'), sheet.getRange('Q2:Q')])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function setupTransactionSheet_(sheet) {
  assertCompatible_(sheet, TX_HEADERS);
  writeHeaders_(sheet, TX_HEADERS, '#1f4e78');
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.setTabColor(sheet.getName().indexOf('RWT') === 0 ? '#b45f06' : '#2e7d32');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 38);
  sheet.hideColumns(TX_COL.TX_ID, TX_HEADERS.length - TX_COL.TX_ID + 1);

  setWidths_(sheet, {
    1: 50,
    2: 140,
    3: 120,
    4: 230,
    5: 70,
    6: 70,
    7: 120,
    8: 240,
    9: 120,
    10: 140,
    11: 160,
    12: 140,
    13: 80,
    14: 100,
    15: 140,
    16: 80
  });

  const lastRow = Math.max(sheet.getMaxRows(), 2);
  sheet
    .getRange(2, 1, lastRow - 1, TX_HEADERS.length)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.getRange(2, TX_COL.DATE, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(2, TX_COL.QTY, lastRow - 1, 1).setNumberFormat('#,##0.###');
  sheet.getRange(2, TX_COL.BALANCE_AFTER, lastRow - 1, 1).setNumberFormat('#,##0.###');
  sheet.getRange(2, TX_COL.CREATED_AT, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  sheet
    .getRange('M2:M')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['POSTED', 'VOID'], true)
        .setAllowInvalid(false)
        .build()
    );

  sheet
    .getRange('P2:P')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['WA', 'DUMMY', 'MANUAL'], true)
        .setAllowInvalid(false)
        .build()
    );
}

function setupInternalSheets_(ss) {
  const config = ensureSheet_(ss, SHEET.CONFIG);
  writeHeaders_(config, INTERNAL_HEADERS[SHEET.CONFIG], '#444444');

  Object.keys(INTERNAL_HEADERS).forEach(function (name) {
    if (name === SHEET.CONFIG) return;
    const sheet = ensureSheet_(ss, name);
    writeHeaders_(sheet, INTERNAL_HEADERS[name], '#444444');
  });

  ensureConfigDefaults_();
  ensureCodeSequenceRows_();

  ss.setActiveSheet(ss.getSheetByName(SHEET.DATA));
  Object.keys(INTERNAL_HEADERS).forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (sheet && !sheet.isSheetHidden()) sheet.hideSheet();
  });
}

function applyBalanceFormulas_(sheet, startRow, endRow) {
  const first = Math.max(startRow, 2);
  if (!sheet || endRow < first) return;

  const formulas = [];
  for (let row = first; row <= endRow; row++) {
    formulas.push([
      `=IF($C${row}="","",SUMIF('WWT In'!$C$2:$C,$C${row},'WWT In'!$E$2:$E))`,
      `=IF($C${row}="","",SUMIF('WWT Out'!$C$2:$C,$C${row},'WWT Out'!$E$2:$E))`,
      `=IF($C${row}="","",J${row}-K${row})`,
      `=IF($C${row}="","",SUMIF('RWT In'!$C$2:$C,$C${row},'RWT In'!$E$2:$E))`,
      `=IF($C${row}="","",SUMIF('RWT Out'!$C$2:$C,$C${row},'RWT Out'!$E$2:$E))`,
      `=IF($C${row}="","",M${row}-N${row})`,
      `=IF(OR($C${row}="",$I${row}=""),"",IF(L${row}<$I${row},"YA",""))`,
      `=IF(OR($C${row}="",$I${row}=""),"",IF(O${row}<$I${row},"YA",""))`
    ]);
  }

  sheet.getRange(first, DATA_COL.IN_WWT, formulas.length, 8).setFormulas(formulas);
}

function syncCodeSequences_() {
  const ss = getSpreadsheet_();
  const data = ss.getSheetByName(SHEET.DATA);
  const codes = ss.getSheetByName(SHEET.CODES);
  if (!data || !codes) return;

  const maxFromData = {
    F: 0,
    M: 0,
    C: 0,
    T: 0
  };

  const lastRow = data.getLastRow();
  if (lastRow >= 2) {
    const values = data.getRange(2, DATA_COL.CODE, lastRow - 1, 1).getValues();
    values.forEach(function (row) {
      const match = String(row[0]).trim().match(/^([FMCT])-(\d+)$/i);
      if (!match) return;
      const prefix = match[1].toUpperCase();
      const number = parseInt(match[2], 10);
      if (number > maxFromData[prefix]) maxFromData[prefix] = number;
    });
  }

  const existing = {
    F: 0,
    M: 0,
    C: 0,
    T: 0
  };

  const lastCodeRow = codes.getLastRow();
  if (lastCodeRow >= 2) {
    const values = codes.getRange(2, 1, lastCodeRow - 1, 2).getValues();
    values.forEach(function (row) {
      const prefix = String(row[0]).trim().toUpperCase();
      if (prefix in existing) {
        existing[prefix] = Math.max(existing[prefix], Number(row[1]) || 0);
      }
    });
  }

  const rows = CODE_PREFIXES.map(function (prefix) {
    return [prefix, Math.max(existing[prefix], maxFromData[prefix]), new Date()];
  });
  codes.getRange(2, 1, rows.length, 3).setValues(rows);
}

function ensureMenuTrigger_(ss) {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'buildMenu';
  });
  if (!exists) {
    ScriptApp.newTrigger('buildMenu').forSpreadsheet(ss).onOpen().create();
  }
}
