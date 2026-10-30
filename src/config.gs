const SPREADSHEET_ID = '19KlQdT11UzJlUOjZ_CMMUMEK-7GgRt2bJEA4rfIhGPU';
const TIMEZONE = 'Asia/Jakarta';
const CODE_PREFIXES = ['F', 'M', 'C', 'T'];

const SHEET = {
  DATA: 'Data Material',
  WWT_IN: 'WWT In',
  WWT_OUT: 'WWT Out',
  RWT_IN: 'RWT In',
  RWT_OUT: 'RWT Out',
  CONFIG: 'Config',
  DRAFTS: 'Drafts',
  INBOX: 'Inbox',
  OUTBOX: 'Outbox',
  AUDIT: 'Audit Log',
  CODES: 'Code Sequences',
  ALIASES: 'Material Alias'
};

const TX_SHEETS = [SHEET.WWT_IN, SHEET.WWT_OUT, SHEET.RWT_IN, SHEET.RWT_OUT];

const DATA_HEADERS = [
  'No.',
  'Picture',
  'Material Code',
  'Material Name',
  'Category',
  'Specification',
  'Bahan',
  'Unit',
  'Min Stock',
  'Stock In WWT',
  'Stock Out WWT',
  'Current WWT',
  'Stock In RWT',
  'Stock Out RWT',
  'Current RWT',
  'Reorder WWT',
  'Reorder RWT',
  'Status',
  'Is Dummy'
];

const TX_HEADERS = [
  'No.',
  'Date',
  'Material Code',
  'Material Name',
  'Qty',
  'Unit',
  'User',
  'Remarks',
  'Picture',
  'Tx ID',
  'WA Message ID',
  'Draft ID',
  'Status',
  'Balance After',
  'Created At',
  'Source'
];

const INTERNAL_HEADERS = {
  [SHEET.CONFIG]: ['Key', 'Value', 'Description', 'Updated At'],
  [SHEET.DRAFTS]: [
    'Draft ID',
    'Chat ID',
    'Sender',
    'WA Message ID',
    'Raw Message',
    'Direction',
    'Material Code',
    'Material Name',
    'Qty',
    'Unit',
    'Container',
    'User',
    'Note',
    'Status',
    'Candidates JSON',
    'Created At',
    'Updated At'
  ],
  [SHEET.INBOX]: ['WA Message ID', 'Received At', 'Sender', 'Raw Payload', 'Status', 'Error'],
  [SHEET.OUTBOX]: [
    'ID',
    'Chat ID',
    'Payload',
    'Status',
    'Attempts',
    'Next Attempt At',
    'Last Error',
    'Created At',
    'Sent At'
  ],
  [SHEET.AUDIT]: [
    'ID',
    'Timestamp',
    'Actor',
    'Action',
    'Entity Type',
    'Entity ID',
    'Before JSON',
    'After JSON',
    'WA Message ID',
    'Notes'
  ],
  [SHEET.CODES]: ['Prefix', 'Last Number', 'Updated At'],
  [SHEET.ALIASES]: ['Alias', 'Normalized Alias', 'Material Code', 'Created At']
};

const DATA_COL = {
  NO: 1,
  PICTURE: 2,
  CODE: 3,
  NAME: 4,
  CATEGORY: 5,
  SPEC: 6,
  BAHAN: 7,
  UNIT: 8,
  MIN: 9,
  IN_WWT: 10,
  OUT_WWT: 11,
  CUR_WWT: 12,
  IN_RWT: 13,
  OUT_RWT: 14,
  CUR_RWT: 15,
  REORDER_WWT: 16,
  REORDER_RWT: 17,
  STATUS: 18,
  IS_DUMMY: 19
};

const TX_COL = {
  NO: 1,
  DATE: 2,
  CODE: 3,
  NAME: 4,
  QTY: 5,
  UNIT: 6,
  USER: 7,
  REMARKS: 8,
  PICTURE: 9,
  TX_ID: 10,
  WA_MESSAGE_ID: 11,
  DRAFT_ID: 12,
  STATUS: 13,
  BALANCE_AFTER: 14,
  CREATED_AT: 15,
  SOURCE: 16
};

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureReady_() {
  const ss = getSpreadsheet_();
  const required = [
    SHEET.DATA,
    SHEET.WWT_IN,
    SHEET.WWT_OUT,
    SHEET.RWT_IN,
    SHEET.RWT_OUT,
    SHEET.CONFIG,
    SHEET.CODES
  ];
  const missing = required.some(function (name) {
    return !ss.getSheetByName(name);
  });
  if (missing || !isSchemaReady_(ss)) setupSchema();
  return getSpreadsheet_();
}

function isSchemaReady_(ss) {
  const data = ss.getSheetByName(SHEET.DATA);
  const transaction = ss.getSheetByName(SHEET.WWT_IN);
  if (!data || !transaction) return false;
  const dataHeader = data.getRange(1, 1, 1, DATA_HEADERS.length).getValues()[0].join('|');
  const transactionHeader = transaction
    .getRange(1, 1, 1, TX_HEADERS.length)
    .getValues()[0]
    .join('|');
  return dataHeader === DATA_HEADERS.join('|') && transactionHeader === TX_HEADERS.join('|');
}

function writeHeaders_(sheet, headers, color) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range
    .setBackground(color || '#1f4e78')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 38);
}

function setWidths_(sheet, widths) {
  Object.keys(widths).forEach(function (column) {
    sheet.setColumnWidth(Number(column), widths[column]);
  });
}

function getConfig_(key) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET.CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0]) === key) return String(values[index][1]);
  }
  return '';
}

function setConfig_(key, value, description) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET.CONFIG);
  if (!sheet) throw new Error('Tab Config belum ada. Jalankan setupSchema().');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for (let index = 0; index < values.length; index++) {
      if (String(values[index][0]) === key) {
        sheet
          .getRange(index + 2, 2, 1, 3)
          .setValues([[value, description || values[index][2] || '', new Date()]]);
        return;
      }
    }
  }
  sheet.appendRow([key, value, description || '', new Date()]);
}

function ensureConfigDefaults_() {
  const defaults = [
    ['TIMEZONE', TIMEZONE, 'Timezone timestamp dan laporan'],
    ['FUZZY_THRESHOLD', '0.75', 'Ambang minimal fuzzy matching'],
    ['FUZZY_MARGIN', '0.10', 'Selisih minimal kandidat teratas dan kedua'],
    ['AI_ENABLED', 'FALSE', 'Aktifkan AI vision fase 2'],
    ['DUMMY_SEEDED', 'FALSE', 'Penanda data dummy aktif'],
    ['FONNTE_ENABLED', 'FALSE', 'Aktifkan webhook Fonnte'],
    ['FONNTE_ALLOWED_NUMBERS', '', 'Nomor WhatsApp yang diizinkan, pisahkan dengan koma'],
    ['FONNTE_QUOTE_REPLY', 'TRUE', 'Balas dengan quote inboxid Fonnte']
  ];
  defaults.forEach(function (row) {
    if (getConfig_(row[0]) === '') setConfig_(row[0], row[1], row[2]);
  });
}

function ensureCodeSequenceRows_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET.CODES);
  if (!sheet) throw new Error('Tab Code Sequences belum ada.');
  const lastRow = sheet.getLastRow();
  const existing = {};
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach(function (row) {
      if (row[0]) existing[String(row[0])] = Number(row[1]) || 0;
    });
  }
  CODE_PREFIXES.forEach(function (prefix) {
    if (!(prefix in existing)) existing[prefix] = 0;
  });
  const rows = CODE_PREFIXES.map(function (prefix) {
    return [prefix, existing[prefix], new Date()];
  });
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

function nextNo_(sheet, column) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  let max = 0;
  values.forEach(function (row) {
    const value = Number(row[0]);
    if (value > max) max = value;
  });
  return max + 1;
}

function renumberNo_(sheet, column) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  const rows = values.map(function (_, index) {
    return [index + 1];
  });
  sheet.getRange(2, column, rows.length, 1).setValues(rows);
}
