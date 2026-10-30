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


const PARSER_UNIT_ALIASES = {
  pc: 'pcs',
  pcs: 'pcs',
  buah: 'pcs',
  roll: 'roll',
  rol: 'roll',
  pasang: 'pasang',
  kg: 'kg',
  kilogram: 'kg',
  unit: 'unit',
  box: 'box',
  pack: 'pack',
  set: 'set',
  batang: 'batang',
  lembar: 'lembar',
  meter: 'meter',
  m: 'meter',
  liter: 'liter',
  l: 'liter',
  tube: 'tube',
  kaleng: 'kaleng'
};

const PARSER_UNIT_PRIORITY = {
  pcs: 100,
  roll: 95,
  pasang: 90,
  unit: 80,
  box: 70,
  pack: 70,
  set: 70,
  batang: 60,
  lembar: 60,
  tube: 60,
  kaleng: 60,
  kg: 40,
  liter: 35,
  meter: 30
};

const PARSER_USER_NAMES = ['Fariz', 'Ismat', 'Dimas', 'Budi', 'Admin'];

const PARSER_USER_STOPWORDS = {
  container: true,
  kontainer: true,
  wwt: true,
  rwt: true,
  gudang: true,
  stock: true,
  barang: true,
  kerja: true,
  maintenance: true,
  perbaikan: true,
  line: true,
  hose: true,
  panel: true,
  valve: true,
  kabel: true,
  listrik: true,
  area: true,
  support: true,
  fitting: true,
  flange: true,
  cleaning: true,
  packing: true,
  material: true,
  disimpan: true,
  simpan: true,
  masuk: true,
  di: true,
  ke: true,
  untuk: true,
  buat: true,
  dari: true,
  oleh: true
};

const PARSER_CATEGORY_RULES = [
  {
    category: 'T',
    words: [
      'kunci',
      'tang',
      'obeng',
      'palu',
      'cutter',
      'gunting',
      'test pen',
      'multimeter',
      'wrench',
      'sealant gun',
      'sapu',
      'sikat'
    ]
  },
  {
    category: 'F',
    words: [
      'elbow',
      'tee',
      'connector',
      'hose',
      'selang',
      'clamp',
      'flange',
      'valve',
      'fitting',
      'pipa',
      'pvc',
      'reducer',
      'union',
      'nipple',
      'bushing',
      'adapter',
      'plug',
      'cap',
      'coupling',
      'gasket'
    ]
  },
  {
    category: 'C',
    words: [
      'seal tape',
      'teflon tape',
      'tape',
      'lakban',
      'isolasi',
      'cable tie',
      'sarung tangan',
      'masker',
      'safety goggles',
      'goggles',
      'elektroda',
      'majun',
      'lap',
      'kawat las',
      'mata gerinda',
      'batu gerinda',
      'oli',
      'grease',
      'cleaner',
      'solvent',
      'tissue',
      'kertas'
    ]
  }
];

const PARSER_MASTER = [
  {
    code: 'F-001',
    name: 'Elbow 90° PVC 1"',
    unit: 'pcs',
    aliases: ['Elbow 90° PVC 1"', 'Elbow 90 PVC 1"', 'Elbow PVC 1"', 'Elbow 90', 'Elbow']
  },
  {
    code: 'F-002',
    name: 'Tee PVC 1"',
    unit: 'pcs',
    aliases: ['Tee PVC 1"', 'Tee PVC', 'Tee 1"']
  },
  {
    code: 'M-001',
    name: 'Bearing 6204',
    unit: 'pcs',
    aliases: ['Bearing 6204', 'Bearing']
  },
  {
    code: 'M-002',
    name: 'Selang Fleksibel 1/2"',
    unit: 'meter',
    aliases: ['Selang Fleksibel 1/2"', 'Selang Fleksibel', 'Selang 1/2"']
  },
  {
    code: 'C-001',
    name: 'Seal Tape',
    unit: 'roll',
    aliases: ['Seal Tape', 'Teflon Tape', 'Tape Teflon']
  },
  {
    code: 'C-002',
    name: 'Sarung Tangan Nitrile',
    unit: 'pasang',
    aliases: ['Sarung Tangan Nitrile', 'Sarung Tangan Kerja Nitrile']
  },
  {
    code: 'T-001',
    name: 'Kunci Pas 10 mm',
    unit: 'pcs',
    aliases: ['Kunci Pas 10 mm', 'Kunci Pas 10']
  },
  {
    code: 'T-002',
    name: 'Tang Kombinasi 8"',
    unit: 'pcs',
    aliases: ['Tang Kombinasi 8"', 'Tang Kombinasi']
  }
];

function parseMessage(raw) {
  const original = String(raw == null ? '' : raw);
  const text = stripParserMeta_(original);
  const container = findParserContainer_(text);
  const direction = findParserDirection_(text, container);
  const user = findParserUser_(text);
  const qty = findParserQty_(text);
  const spans = [];

  if (container.start >= 0) spans.push({ start: container.start, end: container.end });
  if (user.start >= 0) spans.push({ start: user.start, end: user.end });
  if (qty.start >= 0) spans.push({ start: qty.start, end: qty.end });

  const withoutCore = removeParserSpans_(text, spans);
  const noteInfo = findParserNote_(withoutCore);
  const materialSource = noteInfo
    ? removeParserSpans_(withoutCore, [{ start: noteInfo.start, end: noteInfo.end }])
    : withoutCore;
  const materialText = cleanParserMaterial_(materialSource);
  const match = matchParserMaterial_(materialText);
  const missing = [];

  if (!direction) missing.push('direction');
  if (!materialText) missing.push('material');
  if (qty.qty == null) missing.push('qty');
  if (!qty.unit) missing.push('unit');
  if (!container.value) missing.push('container');
  if (!user.value) missing.push('user');

  const unitMismatch = !!(match && qty.unit && match.unit && qty.unit !== match.unit);
  const needsMaster = !!(materialText && !match);
  const classification = match
    ? { category: match.code.charAt(0), keyword: '' }
    : classifyMaterial_(materialText);

  return {
    raw: original,
    text: text,
    direction: direction,
    materialText: materialText,
    materialCode: match ? match.code : '',
    materialName: match ? match.name : materialText,
    qty: qty.qty,
    unit: qty.unit,
    container: container.value,
    user: user.value,
    note: noteInfo ? noteInfo.note : '',
    missing: missing,
    missingFields: missing,
    needsMaster: needsMaster,
    autoCreate: needsMaster,
    category: classification.category,
    categoryKeyword: classification.keyword,
    unitMismatch: unitMismatch,
    needsClarification: missing.length > 0 || unitMismatch,
    candidates: {
      container: container.candidates || [],
      material: match ? [match.code] : [],
      qty: qty.ambiguous ? qty.candidates : []
    }
  };
}

function classifyMaterial_(text) {
  const lower = String(text || '').toLowerCase().replace(/\s+/g, ' ');
  const haystack = ' ' + lower + ' ';

  for (let ruleIndex = 0; ruleIndex < PARSER_CATEGORY_RULES.length; ruleIndex++) {
    const rule = PARSER_CATEGORY_RULES[ruleIndex];
    for (let wordIndex = 0; wordIndex < rule.words.length; wordIndex++) {
      const word = rule.words[wordIndex];
      if (haystack.indexOf(' ' + word + ' ') !== -1) {
        return { category: rule.category, keyword: word };
      }
    }
  }

  return { category: 'M', keyword: '' };
}

function stripParserMeta_(text) {
  return String(text || '')
    .replace(/\((?:container|kontainer)\s+tidak\s+disebut\)/gi, ' ')
    .replace(/\(\d+\s*\/\s*\d+\)\s*$/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findParserContainer_(text) {
  const patterns = [
    /\b(?:ke|di|dari|untuk|masuk|keluar)\s+(?:container|kontainer)\s+(WWT|RWT)\b/i,
    /\b(?:container|kontainer)\s+(WWT|RWT)\b/i,
    /\b(?:ke|di|dari|untuk|masuk|keluar)\s+(WWT|RWT)\b/i,
    /\b(WWT|RWT)\b/i
  ];
  const all = [];
  String(text).replace(/\b(WWT|RWT)\b/gi, function (token) {
    const upper = token.toUpperCase();
    if (all.indexOf(upper) === -1) all.push(upper);
    return token;
  });

  if (all.length > 1) {
    return { value: '', start: -1, end: -1, candidates: all };
  }

  for (let index = 0; index < patterns.length; index++) {
    const match = patterns[index].exec(text);
    if (match) {
      return {
        value: match[1].toUpperCase(),
        start: match.index,
        end: match.index + match[0].length,
        candidates: all
      };
    }
  }

  return { value: '', start: -1, end: -1, candidates: [] };
}

function findParserDirection_(text, container) {
  const lower = String(text || '').toLowerCase();

  if (/\b(?:keluar(?:kan)?|out)\b/.test(lower)) return 'OUT';

  const hasOutVerb = /\b(?:ambil|mengambil|minta|diminta)\b/.test(lower);
  const hasIn = /\b(?:masuk(?:kan)?|memasukkan|in|terima|diterima|menerima|setor|menyetor|menyerahkan|diserahkan|menyimpan|simpan|membawa|kirim|mengirim|stock[-\s]?in)\b/.test(
    lower
  );

  if (hasOutVerb && hasIn) {
    if (/\b(?:dari|keluar)\s+(?:container|kontainer)?\s*(WWT|RWT)\b/i.test(text)) return 'OUT';
    if (/\b(?:ke|di|masuk|untuk)\s+(?:container|kontainer)?\s*(WWT|RWT)\b/i.test(text)) return 'IN';
    return 'OUT';
  }

  if (hasOutVerb) return 'OUT';
  if (hasIn) return 'IN';

  if (container.value) {
    if (/\b(?:dari|keluar)\s+(?:container|kontainer)?\s*(WWT|RWT)\b/i.test(text)) return 'OUT';
    if (/\b(?:ke|di|masuk|untuk)\s+(?:container|kontainer)?\s*(WWT|RWT)\b/i.test(text)) return 'IN';
  }

  return '';
}

function findParserUser_(text) {
  const patterns = [
    {
      regex: /\b(?:peminta|pemberi|supplier|diserahkan(?:\s+oleh)?|oleh|dari)\s+([A-Za-z][A-Za-z.'-]*)\b/gi,
      mode: 'cue'
    },
    {
      regex: /\b([A-Za-z][A-Za-z.'-]*)\s+(?:ambil|mengambil|minta|memasukkan|menyerahkan|membawa|menyetor|setor|kirim|mengirim)\b/gi,
      mode: 'actor'
    },
    {
      regex: /\b(?:container|kontainer)\s+(?:WWT|RWT)\s+([A-Za-z][A-Za-z.'-]*)\b/gi,
      mode: 'container'
    },
    {
      regex: /,\s*([A-Za-z][A-Za-z.'-]*)\.?\s*$/i,
      mode: 'comma'
    }
  ];

  for (let p = 0; p < patterns.length; p++) {
    const pattern = patterns[p];
    let match;
    while ((match = pattern.regex.exec(text)) !== null) {
      const name = cleanParserName_(match[1]);
      if (!name || isParserStopword_(name)) continue;

      let start = match.index;
      let end = match.index + match[0].length;

      if (pattern.mode === 'container') {
        const offset = match[0].lastIndexOf(match[1]);
        start = match.index + offset;
        end = start + match[1].length;
      }

      return { value: name, start: start, end: end };
    }
  }

  return { value: '', start: -1, end: -1 };
}

function cleanParserName_(name) {
  const value = String(name || '')
    .replace(/^[^A-Za-z]+/, '')
    .replace(/[^A-Za-z.'-]+$/, '')
    .replace(/[.'-]+$/, '')
    .trim();
  if (!value) return '';

  const canonical = PARSER_USER_NAMES.filter(function (candidate) {
    return candidate.toLowerCase() === value.toLowerCase();
  })[0];

  if (canonical) return canonical;
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
}

function isParserStopword_(name) {
  return !!PARSER_USER_STOPWORDS[String(name || '').toLowerCase()];
}

function findParserQty_(text) {
  const regex = /\b(\d+(?:[.,]\d+)?)\s*(pcs|pc|buah|roll|rol|pasang|kg|kilogram|unit|box|pack|set|batang|lembar|meter|m|liter|l|tube|kaleng)\b/gi;
  const candidates = [];
  let match;

  while ((match = regex.exec(text)) !== null) {
    const unit = PARSER_UNIT_ALIASES[match[2].toLowerCase()] || match[2].toLowerCase();
    candidates.push({
      qty: Number(String(match[1]).replace(',', '.')),
      unit: unit,
      start: match.index,
      end: match.index + match[0].length,
      priority: PARSER_UNIT_PRIORITY[unit] || 0
    });
  }

  if (!candidates.length) {
    return { qty: null, unit: '', start: -1, end: -1, candidates: [], ambiguous: false };
  }

  const sorted = candidates.slice().sort(function (a, b) {
    return b.priority - a.priority || a.start - b.start;
  });
  const best = sorted[0];
  const ambiguous =
    candidates.length > 1 &&
    candidates.filter(function (candidate) {
      return candidate.priority === best.priority;
    }).length > 1;

  return {
    qty: best.qty,
    unit: best.unit,
    start: best.start,
    end: best.end,
    candidates: candidates,
    ambiguous: ambiguous
  };
}

function findParserNote_(text) {
  const match = /\b(?:untuk|buat)\b\s*(.*)$/i.exec(text);
  if (!match) return null;

  const raw = String(match[1] || '').trim();
  return {
    note: cleanParserNote_(raw),
    raw: raw,
    start: match.index,
    end: match.index + match[0].length
  };
}

function cleanParserNote_(note) {
  let value = String(note || '')
    .replace(/\b(?:oleh|dari)\s+[A-Za-z][A-Za-z.'-]*\s*$/i, '')
    .replace(/[\s,;.]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (!value) return '';

  const storage = {
    untuk: true,
    buat: true,
    masuk: true,
    memasukkan: true,
    disimpan: true,
    simpan: true,
    ditaruh: true,
    taruh: true,
    di: true,
    ke: true,
    dari: true,
    container: true,
    kontainer: true,
    wwt: true,
    rwt: true
  };
  const words = value.toLowerCase().split(/\s+/).filter(Boolean);

  if (
    words.every(function (word) {
      return !!storage[word];
    })
  ) {
    return '';
  }

  return value;
}

function cleanParserMaterial_(text) {
  let value = String(text || '');
  value = value.replace(
    /\b(?:mohon|tolong|barang|stock[-\s]?in|stock\s+masuk|memasukkan|masuk(?:kan)?|in|keluar(?:kan)?|out|ambil|mengambil|minta|diminta|terima|diterima|menerima|setor|menyetor|menyerahkan|diserahkan|menyimpan|simpan|membawa|kirim|mengirim|supplier|pemberi|peminta|oleh|dari|ke|di|untuk|buat|sebanyak|dengan|berupa|yaitu|adalah)\b/gi,
    ' '
  );
  value = value.replace(/\b(?:container|kontainer|wwt|rwt)\b/gi, ' ');
  value = value.replace(/[;,:.!?]+/g, ' ');
  value = value.replace(/[()]/g, ' ');
  value = value.replace(/\s+/g, ' ').trim();
  value = value.replace(/^[-–—\s]+|[-–—\s]+$/g, '');
  value = value.replace(/^(?:yang|adalah)\s+/i, '');
  return value.trim();
}

function matchParserMaterial_(materialText) {
  const key = normalizeParserMaterialKey_(materialText);
  if (!key) return null;

  let found = null;
  PARSER_MASTER.forEach(function (material) {
    material.aliases.forEach(function (alias) {
      if (normalizeParserMaterialKey_(alias) === key) found = material;
    });
  });
  return found;
}

function normalizeParserMaterialKey_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[“”]/g, '"')
    .replace(/[°º]/g, ' ')
    .replace(/\b(?:inch|inchi)\b/g, '"')
    .replace(/["']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function removeParserSpans_(text, spans) {
  const valid = spans
    .filter(function (span) {
      return span && span.start >= 0 && span.end > span.start;
    })
    .sort(function (a, b) {
      return b.start - a.start;
    });

  let output = String(text || '');
  valid.forEach(function (span) {
    output = output.slice(0, span.start) + ' ' + output.slice(span.end);
  });
  return output;
}


function resolveParsedMaterial_(parsed) {
  if (!parsed || !parsed.materialText || !parsed.unit) return parsed;

  const material = ensureMaterialByText_(parsed.materialText, parsed.unit);
  const unitMismatch = !!(parsed.unit && material.unit && parsed.unit !== material.unit);
  const missing = (parsed.missingFields || []).slice();

  return Object.assign({}, parsed, {
    materialCode: material.code,
    materialName: material.name,
    category: material.category,
    masterUnit: material.unit,
    needsMaster: false,
    autoCreate: false,
    autoCreated: !!material.created,
    unitMismatch: unitMismatch,
    needsClarification: missing.length > 0 || unitMismatch
  });
}

function ensureMaterialByText_(materialText, unit) {
  const text = String(materialText || '').trim();
  const parsedUnit = String(unit || '').trim();
  if (!text) throw new Error('Nama material kosong.');
  if (!parsedUnit) {
    throw new Error('Satuan material wajib diisi sebelum membuat material baru.');
  }

  const ss = ensureReady_();
  let found = findMaterialByText_(ss, text);
  if (found) return found;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    found = findMaterialByText_(ss, text);
    if (found) return found;

    const classified = classifyMaterial_(text);
    const code = nextMaterialCode_(ss, classified.category);
    const name = displayMaterialName_(text);
    const data = ss.getSheetByName(SHEET.DATA);
    const row = Math.max(data.getLastRow() + 1, 2);
    const number = nextNo_(data, DATA_COL.NO);

    data.getRange(row, 1, 1, 9).setValues([
      [number, '', code, name, classified.category, '', '', parsedUnit, 0]
    ]);
    data.getRange(row, DATA_COL.STATUS, 1, 2).setValues([['ACTIVE', false]]);
    applyBalanceFormulas_(data, row, row);
    appendMaterialAlias_(ss, text, code);
    SpreadsheetApp.flush();

    return {
      code: code,
      name: name,
      unit: parsedUnit,
      category: classified.category,
      row: row,
      created: true
    };
  } finally {
    lock.releaseLock();
  }
}

function findMaterialByText_(ss, text) {
  const key = normalizeMaterialText_(text);
  if (!key) return null;

  const data = ss.getSheetByName(SHEET.DATA);
  if (!data || data.getLastRow() < 2) return null;

  const lastRow = data.getLastRow();
  const values = data.getRange(2, DATA_COL.CODE, lastRow - 1, 6).getValues();
  const byCode = {};

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    const code = String(row[0]).trim();
    const name = String(row[1]).trim();
    const category = String(row[2]).trim();
    const unit = String(row[5]).trim();
    const record = {
      code: code,
      name: name,
      unit: unit,
      category: category,
      row: index + 2,
      created: false
    };

    if (code) byCode[code.toUpperCase()] = record;
    if (normalizeMaterialText_(name) === key) return record;
  }

  const aliases = ss.getSheetByName(SHEET.ALIASES);
  if (!aliases || aliases.getLastRow() < 2) return null;

  const aliasValues = aliases.getRange(2, 1, aliases.getLastRow() - 1, 3).getValues();
  for (let index = 0; index < aliasValues.length; index++) {
    const row = aliasValues[index];
    const normalized = String(row[1]).trim() || normalizeMaterialText_(row[0]);
    if (normalized !== key) continue;

    const code = String(row[2]).trim().toUpperCase();
    if (byCode[code]) return byCode[code];
  }

  return null;
}

function nextMaterialCode_(ss, prefix) {
  const codes = ss.getSheetByName(SHEET.CODES);
  if (!codes) throw new Error('Tab Code Sequences belum ada.');

  if (codes.getLastRow() < 2) ensureCodeSequenceRows_();

  const lastRow = codes.getLastRow();
  const values = codes.getRange(2, 1, lastRow - 1, 3).getValues();
  let targetRow = -1;
  let lastNumber = 0;

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    if (String(row[0]).trim().toUpperCase() === prefix) {
      targetRow = index + 2;
      lastNumber = Number(row[1]) || 0;
      break;
    }
  }

  const next = Math.max(lastNumber, maxMaterialNumber_(ss, prefix)) + 1;

  if (targetRow < 0) {
    codes.appendRow([prefix, next, new Date()]);
  } else {
    codes.getRange(targetRow, 2, 1, 2).setValues([[next, new Date()]]);
  }

  return prefix + '-' + String(next).padStart(3, '0');
}

function maxMaterialNumber_(ss, prefix) {
  const data = ss.getSheetByName(SHEET.DATA);
  if (!data || data.getLastRow() < 2) return 0;

  const values = data.getRange(2, DATA_COL.CODE, data.getLastRow() - 1, 1).getValues();
  const pattern = new RegExp('^' + prefix + '-(\\d+)$', 'i');
  let max = 0;

  values.forEach(function (row) {
    const match = String(row[0]).trim().match(pattern);
    if (match) max = Math.max(max, parseInt(match[1], 10));
  });

  return max;
}

function appendMaterialAlias_(ss, alias, code) {
  const sheet = ss.getSheetByName(SHEET.ALIASES);
  if (!sheet) return;

  const normalized = normalizeMaterialText_(alias);
  if (!normalized) return;

  if (sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let index = 0; index < values.length; index++) {
      const row = values[index];
      const existing = String(row[1]).trim() || normalizeMaterialText_(row[0]);
      if (existing === normalized) return;
    }
  }

  sheet.appendRow([alias, normalized, code, new Date()]);
}

function displayMaterialName_(text) {
  const value = String(text || '').trim().replace(/\s+/g, ' ');
  if (!value) return '';
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function normalizeMaterialText_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[“”]/g, '"')
    .replace(/[°º]/g, ' ')
    .replace(/\b(?:inch|inchi)\b/g, '"')
    .replace(/["']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


const DRAFT_STATUS = {
  PENDING: 'PENDING',
  NEEDS_CLARIFICATION: 'NEEDS_CLARIFICATION',
  REVISI: 'REVISI',
  POSTED: 'POSTED',
  CANCELLED: 'CANCELLED'
};

const DRAFT_FIELD_LABELS = {
  direction: 'arah transaksi (IN/OUT)',
  material: 'nama material',
  qty: 'qty',
  unit: 'satuan',
  container: 'container (WWT/RWT)',
  user: 'nama user'
};

function handleIncomingMessage_(rawMessage, meta) {
  const info = meta || {};
  const text = String(rawMessage == null ? '' : rawMessage).trim();
  const reply = normalizeDraftReply_(text);
  const draftId = String(info.draftId || '').trim();

  if (draftId) return handleDraftReply_(draftId, text, info);

  if (reply.command === 'OK' || reply.command === 'REVISI') {
    const latest = findLatestPendingDraft_(info.chatId || '', info.sender || '');
    if (!latest) {
      return {
        status: 'ERROR',
        message: 'Tidak ada draft yang menunggu konfirmasi. Kirim pesan transaksi baru.'
      };
    }
    return handleDraftReply_(latest.draftId, text, info);
  }

  return createDraft_(text, info);
}

function createDraft_(rawMessage, meta) {
  const info = meta || {};

  if (info.waMessageId) {
    const existing = findDraftByWaMessageId_(info.waMessageId);
    if (existing) return existingDraftResult_(existing);
  }

  const parsed = parseMessage(rawMessage);

  if (!parsed.text) {
    return clarificationResult_(parsed, 'Pesan kosong. Kirim pesan transaksi lengkap.');
  }

  if (parsed.missingFields.length > 0 || (parsed.candidates.container || []).length > 1) {
    return clarificationResult_(parsed, buildClarificationMessage_(parsed));
  }

  const resolved = resolveParsedMaterial_(parsed);

  if (resolved.unitMismatch) {
    return clarificationResult_(resolved, buildClarificationMessage_(resolved));
  }

  if (!resolved.materialCode) {
    return clarificationResult_(resolved, buildClarificationMessage_(resolved));
  }

  const ss = ensureReady_();
  const sheet = ss.getSheetByName(SHEET.DRAFTS);
  if (!sheet) {
    throw new Error('Tab Drafts belum ada. Jalankan setupSchema() terlebih dahulu.');
  }

  const draftId = makeDraftId_();
  const now = new Date();
  const row = [
    draftId,
    info.chatId || '',
    info.sender || '',
    info.waMessageId || '',
    String(rawMessage || ''),
    resolved.direction,
    resolved.materialCode,
    resolved.materialName,
    resolved.qty,
    resolved.unit,
    resolved.container,
    resolved.user,
    resolved.note || '',
    DRAFT_STATUS.PENDING,
    JSON.stringify(resolved.candidates || {}),
    now,
    now
  ];

  sheet.appendRow(row);
  SpreadsheetApp.flush();

  return {
    status: DRAFT_STATUS.PENDING,
    draftId: draftId,
    parsed: resolved,
    message: renderDraftPreview_(draftId, resolved)
  };
}

function handleDraftReply_(draftId, replyText, meta) {
  const info = meta || {};
  const draft = findDraftById_(draftId);
  if (!draft) {
    return {
      status: 'ERROR',
      draftId: draftId,
      message: 'Draft ' + draftId + ' tidak ditemukan.'
    };
  }

  if (draft.status === DRAFT_STATUS.POSTED) {
    return {
      status: DRAFT_STATUS.POSTED,
      draftId: draft.draftId,
      message: 'Draft ' + draft.draftId + ' sudah tersimpan. Tidak ada transaksi baru.'
    };
  }

  if (draft.status === DRAFT_STATUS.REVISI || draft.status === DRAFT_STATUS.CANCELLED) {
    return {
      status: draft.status,
      draftId: draft.draftId,
      message: 'Draft ' + draft.draftId + ' sudah dibatalkan. Kirim pesan transaksi baru.'
    };
  }

  const reply = normalizeDraftReply_(replyText);

  if (reply.command === 'OK') {
    return commitDraft_(draft.draftId);
  }

  if (reply.command === 'REVISI') {
    updateDraftStatus_(draft.draftId, DRAFT_STATUS.REVISI);

    if (!reply.payload) {
      return {
        status: DRAFT_STATUS.REVISI,
        draftId: draft.draftId,
        message: 'Draft ' + draft.draftId + ' dibatalkan. Kirim ulang pesan lengkap.'
      };
    }

    return createDraft_(reply.payload, {
      chatId: draft.chatId,
      sender: draft.sender,
      waMessageId: info.waMessageId || '',
      draftId: ''
    });
  }

  return {
    status: 'WAITING_CONFIRMATION',
    draftId: draft.draftId,
    message: renderDraftPreview_(draft.draftId, draft)
  };
}

function commitDraft_(draftId) {
  const ss = ensureReady_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const draft = findDraftById_(draftId, ss);
    if (!draft) throw new Error('Draft ' + draftId + ' tidak ditemukan.');
    if (draft.status === DRAFT_STATUS.POSTED) {
      return {
        status: DRAFT_STATUS.POSTED,
        draftId: draft.draftId,
        message: 'Draft ' + draft.draftId + ' sudah tersimpan. Tidak ada transaksi baru.'
      };
    }
    if (draft.status !== DRAFT_STATUS.PENDING) {
      throw new Error('Draft ' + draft.draftId + ' tidak bisa disimpan karena status ' + draft.status + '.');
    }

    validateDraftForCommit_(draft);

    const material = findMaterialByCode_(ss, draft.materialCode);
    if (!material) {
      throw new Error(
        'Material ' + draft.materialCode + ' tidak ditemukan di Data Material. Jalankan repairFormulas() atau periksa master.'
      );
    }
    if (String(material.unit).trim() !== String(draft.unit).trim()) {
      throw new Error(
        'Satuan draft ' + draft.unit + ' tidak sesuai master ' + material.unit + ' untuk ' + draft.materialCode + '.'
      );
    }

    const duplicate = findPostedTransactionByWaMessageId_(ss, draft.waMessageId);
    if (duplicate) {
      updateDraftStatus_(draft.draftId, DRAFT_STATUS.POSTED, ss);
      return {
        status: DRAFT_STATUS.POSTED,
        draftId: draft.draftId,
        txId: duplicate.txId,
        balanceAfter: duplicate.balanceAfter,
        message: 'Transaksi untuk pesan ini sudah tersimpan sebelumnya. Tx ID: ' + duplicate.txId
      };
    }

    const currentStock = getStock_(ss, draft.container, draft.materialCode);
    if (draft.direction === 'OUT' && Number(draft.qty) > currentStock) {
      throw new Error(
        'Stok ' +
          draft.container +
          ' untuk ' +
          draft.materialCode +
          ' hanya ' +
          currentStock +
          ' ' +
          draft.unit +
          '. Tidak cukup untuk OUT ' +
          draft.qty +
          ' ' +
          draft.unit +
          '.'
      );
    }

    const balanceAfter = currentStock + (draft.direction === 'IN' ? Number(draft.qty) : -Number(draft.qty));
    const sheetName = draftTxSheetName_(draft.direction, draft.container);
    const txSheet = ss.getSheetByName(sheetName);
    if (!txSheet) throw new Error('Tab ' + sheetName + ' belum ada.');

    const now = new Date();
    const txId = 'TX-' + draft.draftId.replace(/^DRF-/, '');
    const row = Math.max(txSheet.getLastRow() + 1, 2);
    const number = nextNo_(txSheet, TX_COL.NO);

    txSheet.getRange(row, 1, 1, TX_HEADERS.length).setValues([
      [
        number,
        now,
        draft.materialCode,
        draft.materialName,
        Number(draft.qty),
        draft.unit,
        draft.user,
        draft.note || '',
        '',
        txId,
        draft.waMessageId || '',
        draft.draftId,
        'POSTED',
        balanceAfter,
        now,
        'WA'
      ]
    ]);
    txSheet.getRange(row, TX_COL.DATE, 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    txSheet.getRange(row, TX_COL.CREATED_AT, 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

    updateDraftStatus_(draft.draftId, DRAFT_STATUS.POSTED, ss);
    appendAudit_(
      ss,
      'POST_TRANSACTION',
      'Draft',
      draft.draftId,
      JSON.stringify(draft),
      JSON.stringify({
        txId: txId,
        sheet: sheetName,
        materialCode: draft.materialCode,
        qty: Number(draft.qty),
        balanceAfter: balanceAfter
      }),
      draft.waMessageId || '',
      ''
    );
    SpreadsheetApp.flush();

    return {
      status: DRAFT_STATUS.POSTED,
      draftId: draft.draftId,
      txId: txId,
      balanceAfter: balanceAfter,
      message: renderCommitMessage_(draft, balanceAfter, txId)
    };
  } finally {
    lock.releaseLock();
  }
}

function validateDraftForCommit_(draft) {
  const missing = [];
  if (draft.direction !== 'IN' && draft.direction !== 'OUT') missing.push('direction');
  if (draft.container !== 'WWT' && draft.container !== 'RWT') missing.push('container');
  if (!draft.materialCode) missing.push('materialCode');
  if (!draft.materialName) missing.push('materialName');
  if (!draft.unit) missing.push('unit');
  if (!draft.user) missing.push('user');
  if (!(Number(draft.qty) > 0)) missing.push('qty');

  if (missing.length > 0) {
    throw new Error('Draft ' + draft.draftId + ' belum lengkap: ' + missing.join(', ') + '.');
  }
}

function draftTxSheetName_(direction, container) {
  if (container !== 'WWT' && container !== 'RWT') {
    throw new Error('Container draft tidak valid: ' + container + '.');
  }
  if (direction !== 'IN' && direction !== 'OUT') {
    throw new Error('Arah draft tidak valid: ' + direction + '.');
  }
  return container + ' ' + (direction === 'IN' ? 'In' : 'Out');
}

function findDraftByWaMessageId_(waMessageId, ss) {
  const target = String(waMessageId || '').trim();
  if (!target) return null;

  const sheet = (ss || getSpreadsheet_()).getSheetByName(SHEET.DRAFTS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][3]).trim() === target) return draftFromRow_(values[index], index + 2);
  }
  return null;
}

function findDraftById_(draftId, ss) {
  const target = String(draftId || '').trim();
  if (!target) return null;

  const sheet = (ss || getSpreadsheet_()).getSheetByName(SHEET.DRAFTS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0]).trim() === target) return draftFromRow_(values[index], index + 2);
  }
  return null;
}

function findLatestPendingDraft_(chatId, sender) {
  const ss = ensureReady_();
  const sheet = ss.getSheetByName(SHEET.DRAFTS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  for (let index = values.length - 1; index >= 0; index--) {
    const row = values[index];
    if (String(row[13]).trim() !== DRAFT_STATUS.PENDING) continue;
    if (chatId && String(row[1]).trim() !== String(chatId).trim()) continue;
    if (sender && String(row[2]).trim() !== String(sender).trim()) continue;
    return draftFromRow_(row, index + 2);
  }
  return null;
}

function existingDraftResult_(draft) {
  if (draft.status === DRAFT_STATUS.POSTED) {
    return {
      status: DRAFT_STATUS.POSTED,
      draftId: draft.draftId,
      message: 'Draft ' + draft.draftId + ' sudah tersimpan. Tidak ada transaksi baru.'
    };
  }

  if (draft.status === DRAFT_STATUS.PENDING) {
    return {
      status: DRAFT_STATUS.PENDING,
      draftId: draft.draftId,
      message: renderDraftPreview_(draft.draftId, draft)
    };
  }

  return {
    status: draft.status,
    draftId: draft.draftId,
    message: 'Draft ' + draft.draftId + ' berstatus ' + draft.status + '. Kirim pesan transaksi baru.'
  };
}

function draftFromRow_(row, rowNumber) {
  return {
    draftId: String(row[0]).trim(),
    chatId: String(row[1]),
    sender: String(row[2]),
    waMessageId: String(row[3]),
    rawMessage: String(row[4]),
    direction: String(row[5]).trim().toUpperCase(),
    materialCode: String(row[6]).trim(),
    materialName: String(row[7]).trim(),
    qty: Number(row[8]),
    unit: String(row[9]).trim(),
    container: String(row[10]).trim().toUpperCase(),
    user: String(row[11]).trim(),
    note: String(row[12]).trim(),
    status: String(row[13]).trim().toUpperCase(),
    candidatesJson: String(row[14]),
    createdAt: row[15],
    updatedAt: row[16],
    row: rowNumber
  };
}

function updateDraftStatus_(draftId, status, ss) {
  const sheet = (ss || getSpreadsheet_()).getSheetByName(SHEET.DRAFTS);
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0]).trim() !== String(draftId).trim()) continue;
    const row = index + 2;
    sheet.getRange(row, 14, 1, 1).setValue(status);
    sheet.getRange(row, 17, 1, 1).setValue(new Date());
    return;
  }
}

function findMaterialByCode_(ss, code) {
  const data = ss.getSheetByName(SHEET.DATA);
  if (!data || data.getLastRow() < 2) return null;

  const target = String(code || '').trim().toUpperCase();
  if (!target) return null;

  const values = data.getRange(2, DATA_COL.CODE, data.getLastRow() - 1, 6).getValues();
  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    if (String(row[0]).trim().toUpperCase() !== target) continue;
    return {
      code: String(row[0]).trim(),
      name: String(row[1]).trim(),
      category: String(row[2]).trim(),
      unit: String(row[5]).trim(),
      row: index + 2
    };
  }
  return null;
}

function getStock_(ss, container, code) {
  const inSheet = ss.getSheetByName(container + ' In');
  const outSheet = ss.getSheetByName(container + ' Out');
  return sumTxQty_(inSheet, code) - sumTxQty_(outSheet, code);
}

function sumTxQty_(sheet, code) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const target = String(code || '').trim().toUpperCase();
  const values = sheet.getRange(2, TX_COL.CODE, sheet.getLastRow() - 1, 11).getValues();
  let total = 0;

  values.forEach(function (row) {
    if (String(row[0]).trim().toUpperCase() !== target) return;
    const status = String(row[10]).trim().toUpperCase();
    if (status && status !== 'POSTED') return;
    total += Number(row[2]) || 0;
  });

  return total;
}

function findPostedTransactionByWaMessageId_(ss, waMessageId) {
  const target = String(waMessageId || '').trim();
  if (!target) return null;

  for (let index = 0; index < TX_SHEETS.length; index++) {
    const name = TX_SHEETS[index];
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) continue;

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, TX_HEADERS.length).getValues();
    for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      if (String(row[TX_COL.STATUS - 1]).trim().toUpperCase() !== 'POSTED') continue;
      if (String(row[TX_COL.WA_MESSAGE_ID - 1]).trim() !== target) continue;
      return {
        sheet: name,
        row: rowIndex + 2,
        txId: String(row[TX_COL.TX_ID - 1]).trim(),
        balanceAfter: Number(row[TX_COL.BALANCE_AFTER - 1])
      };
    }
  }
  return null;
}

function appendAudit_(ss, action, entityType, entityId, beforeJson, afterJson, waMessageId, notes) {
  const sheet = ss.getSheetByName(SHEET.AUDIT);
  if (!sheet) return;

  sheet.appendRow([
    'AUD-' + makeIdSuffix_(),
    new Date(),
    'SYSTEM',
    action,
    entityType,
    entityId,
    beforeJson || '',
    afterJson || '',
    waMessageId || '',
    notes || ''
  ]);
}

function normalizeDraftReply_(text) {
  const raw = String(text == null ? '' : text).trim();
  const commandText = raw.replace(/[\s.!?,]+$/g, '').trim();
  const upper = commandText.toUpperCase();

  if (/^(OK|OKE|OKAY|SIMPAN|POST|LANJUT|BENAR|YA)$/.test(upper)) {
    return { command: 'OK', payload: '' };
  }

  const revision = /^REVISI\b\s*([\s\S]*)$/i.exec(commandText);
  if (revision) {
    return { command: 'REVISI', payload: String(revision[1] || '').trim() };
  }

  return { command: '', payload: raw };
}

function buildClarificationMessage_(parsed) {
  const lines = ['Perlu klarifikasi:'];
  const containerCandidates = (parsed.candidates && parsed.candidates.container) || [];

  if (containerCandidates.length > 1) {
    lines.push('- Container ambigu: ' + containerCandidates.join(' / ') + '. Sebutkan satu.');
  }

  (parsed.missingFields || []).forEach(function (field) {
    if (field === 'container' && containerCandidates.length > 1) return;
    lines.push('- ' + (DRAFT_FIELD_LABELS[field] || field) + ' belum jelas.');
  });

  if (parsed.unitMismatch) {
    lines.push(
      '- Satuan "' +
        parsed.unit +
        '" tidak sesuai master' +
        (parsed.masterUnit ? ' (master: ' + parsed.masterUnit + ')' : '') +
        '. Mohon pakai satuan master.'
    );
  }

  const qtyCandidates = (parsed.candidates && parsed.candidates.qty) || [];
  if (qtyCandidates.length > 1) {
    lines.push(
      '- Qty ambigu: ' +
        qtyCandidates
          .map(function (candidate) {
            return candidate.qty + ' ' + candidate.unit;
          })
          .join(' / ') +
        '. Sebutkan satu.'
    );
  }

  if (lines.length === 1) lines.push('- Data transaksi belum lengkap.');
  lines.push('');
  lines.push('Kirim ulang pesan lengkap, contoh:');
  lines.push('out seal tape 1 roll dari WWT oleh Fariz untuk perbaikan pompa');
  return lines.join('\n');
}

function clarificationResult_(parsed, message) {
  return {
    status: DRAFT_STATUS.NEEDS_CLARIFICATION,
    draftId: '',
    parsed: parsed,
    message: message
  };
}

function renderDraftPreview_(draftId, draft) {
  return [
    'Material: ' + draft.materialName,
    'Material Code: ' + draft.materialCode,
    'Qty: ' + formatDraftQty_(draft.qty, draft.unit),
    'Container: ' + draft.container,
    'User: ' + draft.user,
    'Catatan: ' + (draft.note || '-'),
    '',
    'Draft: ' + draftId,
    'Balas OK untuk simpan, atau REVISI <perbaikan>.'
  ].join('\n');
}

function renderCommitMessage_(draft, balanceAfter, txId) {
  return [
    'Tersimpan.',
    'Material: ' + draft.materialName + ' (' + draft.materialCode + ')',
    'Qty: ' + formatDraftQty_(draft.qty, draft.unit),
    'Container: ' + draft.container,
    'User: ' + draft.user,
    'Saldo ' + draft.container + ': ' + formatDraftQty_(balanceAfter, draft.unit),
    'Tx ID: ' + txId
  ].join('\n');
}

function formatDraftQty_(qty, unit) {
  const value = Number(qty);
  return (isFinite(value) ? String(value) : String(qty || '')) + ' ' + String(unit || '');
}

function makeDraftId_() {
  return 'DRF-' + makeIdSuffix_();
}

function makeIdSuffix_() {
  const now = new Date();
  const pad = function (number) {
    return String(number).padStart(2, '0');
  };
  const stamp =
    now.getFullYear() +
    pad(now.getMonth() + 1) +
    pad(now.getDate()) +
    '-' +
    pad(now.getHours()) +
    pad(now.getMinutes()) +
    pad(now.getSeconds());
  return stamp + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
}


const FONNTE_CONFIG_KEYS = {
  ENABLED: 'FONNTE_ENABLED',
  ALLOWED_NUMBERS: 'FONNTE_ALLOWED_NUMBERS',
  QUOTE_REPLY: 'FONNTE_QUOTE_REPLY'
};

const FONNTE_PROPERTY_KEYS = {
  TOKEN: 'FONNTE_TOKEN',
  WEBHOOK_TOKEN: 'FONNTE_WEBHOOK_TOKEN'
};

const FONNTE_INBOX_STATUS = {
  PROCESSING: 'PROCESSING',
  PROCESSED: 'PROCESSED',
  FAILED: 'FAILED'
};

const FONNTE_OUTBOX_STATUS = {
  PENDING: 'PENDING',
  SENT: 'SENT',
  FAILED: 'FAILED'
};

const FONNTE_API_URL = 'https://api.fonnte.com/send';

function doPost(e) {
  try {
    handleFonnteWebhook_(e);
  } catch (error) {
    Logger.log('Fonnte doPost error: ' + String(error && error.message ? error.message : error));
  }
  return HtmlService.createHtmlOutput('OK');
}

function doGet() {
  return HtmlService.createHtmlOutput('Fonnte webhook aktif.');
}

function handleFonnteWebhook_(e) {
  const config = getFonnteConfig_();
  if (!config.enabled) return { ok: true, ignored: true, reason: 'disabled' };
  if (!config.webhookToken) return { ok: false, error: 'webhook_token_missing' };

  const token = e && e.parameter ? String(e.parameter.token || '') : '';
  if (!fonnteSafeEqual_(token, config.webhookToken)) {
    Logger.log('Fonnte webhook unauthorized.');
    return { ok: false, error: 'unauthorized' };
  }

  const raw = e && e.postData ? String(e.postData.contents || '') : '';
  if (!raw) return { ok: false, error: 'empty_body' };

  let payload;
  try {
    payload = JSON.parse(raw);
  } catch (error) {
    return { ok: false, error: 'invalid_json' };
  }

  return processFonnteWebhook_(payload, raw, config);
}

function processFonnteWebhook_(payload, raw, config) {
  const data = payload || {};
  const senderRaw = String(data.sender || '').trim();
  const member = String(data.member || '').trim();
  const hasGroupMember = member !== '' && member !== '0' && member.toLowerCase() !== 'null';

  if (hasGroupMember || senderRaw.toLowerCase().indexOf('@g.us') !== -1) {
    return { ok: true, ignored: true, reason: 'not_dm' };
  }

  const sender = normalizeFonnteNumber_(senderRaw);
  if (!sender) return { ok: true, ignored: true, reason: 'missing_sender' };
  if (!fonnteNumberAllowed_(sender, config.allowedNumbers)) {
    return { ok: true, ignored: true, reason: 'sender_not_allowed' };
  }

  const text = String(data.message || data.text || '').trim();
  const hasMedia =
    String(data.url || '').trim() !== '' ||
    String(data.filename || '').trim() !== '' ||
    String(data.extension || '').trim() !== '' ||
    String(data.pollname || '').trim() !== '' ||
    String(data.choices || '').trim() !== '';
  if (!text || hasMedia) return { ok: true, ignored: true, reason: 'not_text' };

  const inboxid = fonnteInboxId_(data.inboxid);
  const fallback = [data.device || '', sender, data.timestamp || '', text].join('|');
  const waMessageId = inboxid ? 'fn_' + inboxid : 'fn_' + fonnteHash_(fallback);
  const chatId = sender + '@c.us';

  const ss = ensureReady_();
  const claim = claimFonnteInbox_(ss, {
    waMessageId: waMessageId,
    sender: sender,
    rawPayload: raw
  });

  if (claim.duplicate) {
    return { ok: true, duplicate: true, waMessageId: waMessageId, status: claim.status || '' };
  }

  try {
    const result = handleIncomingMessage_(text, {
      chatId: chatId,
      sender: sender,
      waMessageId: waMessageId
    });

    let outboxId = '';
    if (result && result.message) {
      outboxId = enqueueFonnteReply_(ss, {
        chatId: chatId,
        target: sender,
        text: result.message,
        inboxid: inboxid,
        replyTo: waMessageId
      });
    }

    markFonnteInbox_(ss, claim.row, FONNTE_INBOX_STATUS.PROCESSED, '');
    if (outboxId) flushFonnteOutboxById_(ss, outboxId);

    return {
      ok: true,
      waMessageId: waMessageId,
      status: result && result.status ? result.status : '',
      draftId: result && result.draftId ? result.draftId : ''
    };
  } catch (error) {
    const message = 'Gagal memproses pesan: ' + String(error && error.message ? error.message : error);
    try {
      const outboxId = enqueueFonnteReply_(ss, {
        chatId: chatId,
        target: sender,
        text: message,
        inboxid: inboxid,
        replyTo: waMessageId
      });
      markFonnteInbox_(ss, claim.row, FONNTE_INBOX_STATUS.FAILED, message);
      if (outboxId) flushFonnteOutboxById_(ss, outboxId);
    } catch (secondary) {
      Logger.log('Fonnte error reply failed: ' + String(secondary && secondary.message ? secondary.message : secondary));
    }
    return { ok: true, failed: true, waMessageId: waMessageId, error: message };
  }
}

function claimFonnteInbox_(ss, info) {
  const sheet = ss.getSheetByName(SHEET.INBOX);
  if (!sheet) throw new Error('Tab Inbox belum ada. Jalankan setupSchema().');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let index = 0; index < values.length; index++) {
        const row = values[index];
        if (String(row[0]).trim() !== info.waMessageId) continue;

        const status = String(row[4]).trim().toUpperCase();
        const age = Date.now() - new Date(row[1]).getTime();
        if (status === FONNTE_INBOX_STATUS.PROCESSING && isFinite(age) && age < 5 * 60 * 1000) {
          return { row: index + 2, duplicate: true, status: status };
        }
        if (status === FONNTE_INBOX_STATUS.PROCESSING && !isFinite(age)) {
          return { row: index + 2, duplicate: true, status: status };
        }
        if (status === FONNTE_INBOX_STATUS.PROCESSING || status === '') {
          sheet
            .getRange(index + 2, 1, 1, 6)
            .setValues([
              [
                info.waMessageId,
                new Date(),
                info.sender,
                fonnteTruncate_(info.rawPayload, 45000),
                FONNTE_INBOX_STATUS.PROCESSING,
                ''
              ]
            ]);
          return { row: index + 2, duplicate: false, status: status };
        }

        return { row: index + 2, duplicate: true, status: status };
      }
    }

    const row = sheet.getLastRow() + 1;
    sheet
      .getRange(row, 1, 1, 6)
      .setValues([
        [
          info.waMessageId,
          new Date(),
          info.sender,
          fonnteTruncate_(info.rawPayload, 45000),
          FONNTE_INBOX_STATUS.PROCESSING,
          ''
        ]
      ]);
    return { row: row, duplicate: false, status: '' };
  } finally {
    lock.releaseLock();
  }
}

function markFonnteInbox_(ss, row, status, error) {
  const sheet = ss.getSheetByName(SHEET.INBOX);
  if (!sheet || !row) return;
  sheet.getRange(row, 5, 1, 2).setValues([[status, fonnteTruncate_(error, 1000)]]);
}

function enqueueFonnteReply_(ss, info) {
  const sheet = ss.getSheetByName(SHEET.OUTBOX);
  if (!sheet) throw new Error('Tab Outbox belum ada. Jalankan setupSchema().');

  const target = normalizeFonnteNumber_(info.target || info.chatId);
  const text = String(info.text || '').trim();
  if (!target || !text) return '';

  const chatId = String(info.chatId || '').trim() || target + '@c.us';
  const inboxid = fonnteInboxId_(info.inboxid);
  const replyTo = String(info.replyTo || '').trim();
  const idempotencyKey = 'reply_' + fonnteHash_(String(inboxid || replyTo || target) + '|' + text);
  const id = 'OUT-' + fonnteHash_(idempotencyKey);

  if (findFonnteOutboxRow_(sheet, id)) return id;

  const payload = {
    target: target,
    chatId: chatId,
    text: text,
    inboxid: inboxid,
    replyTo: replyTo,
    idempotencyKey: idempotencyKey
  };

  sheet.appendRow([id, chatId, JSON.stringify(payload), FONNTE_OUTBOX_STATUS.PENDING, 0, new Date(), '', new Date(), '']);
  return id;
}

function findFonnteOutboxRow_(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0]).trim() === String(id).trim()) return index + 2;
  }
  return 0;
}

function flushFonnteOutboxById_(ss, id) {
  const sheet = ss.getSheetByName(SHEET.OUTBOX);
  if (!sheet) return 0;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return 0;

  try {
    const rowNumber = findFonnteOutboxRow_(sheet, id);
    if (!rowNumber) return 0;

    const config = getFonnteConfig_();
    if (!config.enabled || !fonnteSendConfigReady_(config)) return 0;

    const row = sheet.getRange(rowNumber, 1, 1, 9).getValues()[0];
    if (String(row[3]).trim().toUpperCase() !== FONNTE_OUTBOX_STATUS.PENDING) return 0;

    const nextAttempt = row[5] ? new Date(row[5]).getTime() : 0;
    if (isFinite(nextAttempt) && nextAttempt > Date.now()) return 0;

    sendFonnteOutboxRow_(sheet, rowNumber, row, config);
    return 1;
  } finally {
    lock.releaseLock();
  }
}

function flushFonnteOutbox_(ss, limit) {
  const sheet = ss.getSheetByName(SHEET.OUTBOX);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const config = getFonnteConfig_();
  if (!config.enabled || !fonnteSendConfigReady_(config)) return 0;

  const max = Number(limit) > 0 ? Number(limit) : 10;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  let processed = 0;

  for (let index = 0; index < values.length && processed < max; index++) {
    const row = values[index];
    if (String(row[3]).trim().toUpperCase() !== FONNTE_OUTBOX_STATUS.PENDING) continue;

    const nextAttempt = row[5] ? new Date(row[5]).getTime() : 0;
    if (isFinite(nextAttempt) && nextAttempt > Date.now()) continue;

    sendFonnteOutboxRow_(sheet, index + 2, row, config);
    processed++;
  }

  return processed;
}

function sendFonnteOutboxRow_(sheet, rowNumber, row, config) {
  let payload;
  try {
    payload = JSON.parse(String(row[2]));
  } catch (error) {
    sheet
      .getRange(rowNumber, 4, 1, 6)
      .setValues([[FONNTE_OUTBOX_STATUS.FAILED, Number(row[4]) || 0, '', 'Payload JSON tidak valid', row[7], '']]);
    return;
  }

  const result = sendFonnte_(payload, config);
  const attempts = (Number(row[4]) || 0) + 1;

  if (result.ok) {
    sheet
      .getRange(rowNumber, 4, 1, 6)
      .setValues([[FONNTE_OUTBOX_STATUS.SENT, attempts, '', '', row[7] || new Date(), new Date()]]);
    return;
  }

  const retry = result.retryable === true && attempts < 5;
  const nextAttempt = retry ? new Date(Date.now() + fonnteBackoffMs_(attempts)) : '';
  sheet.getRange(rowNumber, 4, 1, 6).setValues([
    [
      retry ? FONNTE_OUTBOX_STATUS.PENDING : FONNTE_OUTBOX_STATUS.FAILED,
      attempts,
      nextAttempt,
      fonnteTruncate_(result.error, 1000),
      row[7] || new Date(),
      ''
    ]
  ]);
}

function sendFonnte_(message, config) {
  const settings = config || getFonnteConfig_();
  if (!fonnteSendConfigReady_(settings)) {
    return { ok: false, statusCode: 0, retryable: false, error: 'Konfigurasi Fonnte belum lengkap.' };
  }

  const target = normalizeFonnteNumber_(message.target || message.chatId);
  const text = fonnteTruncate_(String(message.text || '').trim(), 10000);
  if (!target || !text) {
    return { ok: false, statusCode: 400, retryable: false, error: 'target atau text kosong.' };
  }

  const inboxid = fonnteInboxId_(message.inboxid);
  const fields = [
    ['target', target],
    ['message', text],
    ['preview', 'false'],
    ['connectOnly', 'false']
  ];
  if (settings.quoteReply !== false && inboxid) fields.push(['inboxid', String(inboxid)]);

  const body = fields
    .map(function (pair) {
      return encodeURIComponent(pair[0]) + '=' + encodeURIComponent(pair[1]);
    })
    .join('&');

  try {
    const response = UrlFetchApp.fetch(FONNTE_API_URL, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      headers: { Authorization: settings.token },
      payload: body,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const responseBody = response.getContentText();
    let parsed = {};
    try {
      parsed = JSON.parse(responseBody);
    } catch (error) {
      parsed = {};
    }

    if (status >= 200 && status < 300 && fonnteResponseOk_(parsed)) {
      return {
        ok: true,
        statusCode: status,
        messageId: Array.isArray(parsed.id) ? String(parsed.id[0] || '') : String(parsed.id || ''),
        requestId: parsed.requestid || '',
        detail: String(parsed.detail || '')
      };
    }

    const reason = String(parsed.reason || parsed.detail || responseBody || 'HTTP ' + status);
    return {
      ok: false,
      statusCode: status,
      retryable: fonnteIsRetryable_(status, reason),
      error: reason
    };
  } catch (error) {
    return { ok: false, statusCode: 0, retryable: true, error: String(error && error.message ? error.message : error) };
  }
}

function fonnteResponseOk_(parsed) {
  if (!parsed || typeof parsed !== 'object') return false;
  const status = parsed.status;
  const capitalStatus = parsed.Status;
  return (
    status === true ||
    capitalStatus === true ||
    String(status).toLowerCase() === 'true' ||
    String(capitalStatus).toLowerCase() === 'true'
  );
}

function fonnteIsRetryable_(statusCode, reason) {
  const status = Number(statusCode) || 0;
  if ([0, 408, 409, 425, 429, 500, 502, 503, 504].indexOf(status) !== -1) return true;

  const text = String(reason || '').toLowerCase();
  if (!text) return false;
  return text.indexOf('device disconnected') !== -1 || text.indexOf('server issue') !== -1;
}

function fonnteBackoffMs_(attempts) {
  const value = Math.max(1, Number(attempts) || 1);
  return Math.min(15 * 60 * 1000, 30 * 1000 * Math.pow(2, value - 1));
}

function processOutbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Fonnte outbox sibuk, dilewati.');
    return 'busy';
  }

  try {
    const ss = ensureReady_();
    const processed = flushFonnteOutbox_(ss, 20);
    Logger.log('Fonnte outbox diproses: ' + processed);
    return processed;
  } finally {
    lock.releaseLock();
  }
}

function installFonnteOutboxTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'processOutbox';
  });
  if (exists) return 'Trigger processOutbox sudah ada.';
  ScriptApp.newTrigger('processOutbox').timeBased().everyMinutes(1).create();
  return 'Trigger processOutbox dibuat setiap 1 menit.';
}

function fonnteStatus() {
  const config = getFonnteConfig_();
  const ss = getSpreadsheet_();
  const outbox = ss.getSheetByName(SHEET.OUTBOX);
  let pending = 0;
  if (outbox && outbox.getLastRow() >= 2) {
    const values = outbox.getRange(2, 4, outbox.getLastRow() - 1, 1).getValues();
    values.forEach(function (row) {
      if (String(row[0]).trim().toUpperCase() === FONNTE_OUTBOX_STATUS.PENDING) pending++;
    });
  }

  const lines = [
    'Fonnte enabled: ' + (config.enabled ? 'TRUE' : 'FALSE'),
    'Allowed numbers: ' + (config.allowedNumbers.length ? config.allowedNumbers.join(', ') : '(kosong)'),
    'Quote reply: ' + (config.quoteReply ? 'TRUE' : 'FALSE'),
    'Token: ' + (config.token ? 'terisi' : '(kosong)'),
    'Webhook token: ' + (config.webhookToken ? 'terisi' : '(kosong)'),
    'Outbox pending: ' + pending
  ];
  Logger.log(lines.join('\n'));
  return lines.join('\n');
}

function testFonnteSend() {
  const config = getFonnteConfig_();
  if (!config.allowedNumbers.length) {
    throw new Error('Isi ' + FONNTE_CONFIG_KEYS.ALLOWED_NUMBERS + ' di tab Config dengan nomor tujuan.');
  }

  const result = sendFonnte_(
    {
      target: config.allowedNumbers[0],
      text: 'Tes koneksi Fonnte dari Inventory LCI.'
    },
    config
  );
  Logger.log(JSON.stringify(result));
  return result;
}

function getFonnteConfig_() {
  const properties = PropertiesService.getScriptProperties();
  return {
    enabled: String(getConfig_(FONNTE_CONFIG_KEYS.ENABLED)).trim().toUpperCase() === 'TRUE',
    allowedNumbers: parseFonnteAllowedNumbers_(getConfig_(FONNTE_CONFIG_KEYS.ALLOWED_NUMBERS)),
    quoteReply: String(getConfig_(FONNTE_CONFIG_KEYS.QUOTE_REPLY)).trim().toUpperCase() !== 'FALSE',
    token: String(properties.getProperty(FONNTE_PROPERTY_KEYS.TOKEN) || '').trim(),
    webhookToken: String(properties.getProperty(FONNTE_PROPERTY_KEYS.WEBHOOK_TOKEN) || '').trim()
  };
}

function parseFonnteAllowedNumbers_(value) {
  return String(value || '')
    .split(/[\s,;]+/)
    .map(normalizeFonnteNumber_)
    .filter(function (item) {
      return item !== '';
    });
}

function normalizeFonnteNumber_(value) {
  let digits = String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/@.*$/, '')
    .replace(/\D/g, '');
  if (!digits) return '';
  if (digits.charAt(0) === '0') return '62' + digits.slice(1);
  if (digits.charAt(0) === '8') return '62' + digits;
  return digits;
}

function fonnteNumberAllowed_(number, allowedNumbers) {
  const target = normalizeFonnteNumber_(number);
  if (!target || !allowedNumbers || !allowedNumbers.length) return false;
  return allowedNumbers.indexOf(target) !== -1;
}

function fonnteSendConfigReady_(config) {
  return Boolean(config && config.token);
}

function fonnteInboxId_(value) {
  const number = Number(value);
  return isFinite(number) && number > 0 ? Math.floor(number) : 0;
}

function fonnteSafeEqual_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  if (a.length !== b.length) return false;
  let mismatch = 0;
  for (let index = 0; index < a.length; index++) {
    mismatch |= a.charCodeAt(index) ^ b.charCodeAt(index);
  }
  return mismatch === 0;
}

function fonnteTruncate_(value, limit) {
  const text = String(value == null ? '' : value);
  return text.length > limit ? text.slice(0, limit) : text;
}

function fonnteHash_(value) {
  const text = String(value == null ? '' : value);
  let first = 2166136261;
  let second = 2246822519;
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index);
    first = Math.imul(first ^ code, 16777619);
    second = Math.imul(second ^ code, 2654435761);
  }
  return (first >>> 0).toString(16).padStart(8, '0') + (second >>> 0).toString(16).padStart(8, '0');
}

function cleanupLegacyOpenWA_() {
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEET.CONFIG);
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let index = values.length - 1; index >= 0; index--) {
        if (String(values[index][0]).indexOf('OPENWA_') === 0) sheet.deleteRow(index + 2);
      }
    }
  } catch (error) {
    Logger.log('Gagal membersihkan Config OpenWA: ' + String(error && error.message ? error.message : error));
  }

  try {
    const properties = PropertiesService.getScriptProperties();
    ['OPENWA_API_KEY', 'OPENWA_RELAY_TOKEN'].forEach(function (key) {
      properties.deleteProperty(key);
    });
  } catch (error) {
    Logger.log('Gagal membersihkan properti OpenWA: ' + String(error && error.message ? error.message : error));
  }
}


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


function seedDummy() {
  const ss = ensureReady_();
  const data = ss.getSheetByName(SHEET.DATA);

  if (String(getConfig_('DUMMY_SEEDED')).toUpperCase() === 'TRUE') {
    throw new Error('Data dummy sudah ada. Jalankan clearDummy() terlebih dahulu.');
  }

  const materials = getDummyMaterials_();
  assertNoCodeConflicts_(data, materials);

  let nextNo = nextNo_(data, DATA_COL.NO);
  const materialRows = materials.map(function (material) {
    return [
      nextNo++,
      '',
      material.code,
      material.name,
      material.category,
      material.spec,
      material.bahan,
      material.unit,
      material.min
    ];
  });

  const startRow = Math.max(data.getLastRow() + 1, 2);
  data.getRange(startRow, 1, materialRows.length, 9).setValues(materialRows);
  data.getRange(startRow, DATA_COL.STATUS, materialRows.length, 2).setValues(
    materials.map(function () {
      return ['ACTIVE', true];
    })
  );

  const endRow = startRow + materialRows.length - 1;
  applyBalanceFormulas_(data, startRow, endRow);

  appendDummyTransactions_(ss, SHEET.WWT_IN, getDummyTxWwtIn_());
  appendDummyTransactions_(ss, SHEET.WWT_OUT, getDummyTxWwtOut_());
  appendDummyTransactions_(ss, SHEET.RWT_IN, getDummyTxRwtIn_());
  appendDummyTransactions_(ss, SHEET.RWT_OUT, getDummyTxRwtOut_());

  syncCodeSequences_();
  setConfig_('DUMMY_SEEDED', 'TRUE', 'Penanda data dummy aktif');
  SpreadsheetApp.flush();

  const message = 'Data dummy ditambahkan: ' + materials.length + ' material, 10 transaksi.';
  Logger.log(message);
  return message;
}

function clearDummy() {
  const ss = ensureReady_();
  const data = ss.getSheetByName(SHEET.DATA);

  let deletedMaterials = 0;
  for (let row = data.getLastRow(); row >= 2; row--) {
    const value = data.getRange(row, DATA_COL.IS_DUMMY).getValue();
    if (value === true || String(value).toUpperCase() === 'TRUE') {
      data.deleteRow(row);
      deletedMaterials++;
    }
  }

  let deletedTransactions = 0;
  TX_SHEETS.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    for (let row = sheet.getLastRow(); row >= 2; row--) {
      const value = sheet.getRange(row, TX_COL.SOURCE).getValue();
      if (String(value).toUpperCase() === 'DUMMY') {
        sheet.deleteRow(row);
        deletedTransactions++;
      }
    }
    renumberNo_(sheet, TX_COL.NO);
  });

  renumberNo_(data, DATA_COL.NO);
  applyBalanceFormulas_(data, 2, data.getLastRow());
  rebuildCodeSequencesFromMaterials_();
  setConfig_('DUMMY_SEEDED', 'FALSE', 'Penanda data dummy aktif');
  SpreadsheetApp.flush();

  const message =
    'Data dummy dihapus: ' + deletedMaterials + ' material, ' + deletedTransactions + ' transaksi.';
  Logger.log(message);
  return message;
}

function appendDummyTransactions_(ss, sheetName, transactions) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab ' + sheetName + ' belum ada.');
  if (!transactions.length) return;

  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  const now = new Date();
  let nextNo = nextNo_(sheet, TX_COL.NO);

  const rows = transactions.map(function (transaction, index) {
    const transactionDate = new Date(now.getTime() - (transactions.length - index) * 60000);
    return [
      nextNo++,
      transactionDate,
      transaction.code,
      transaction.name,
      transaction.qty,
      transaction.unit,
      transaction.user,
      transaction.remarks || '',
      '',
      'DUMMY-' +
        sheetName.replace(/\s+/g, '-').toUpperCase() +
        '-' +
        String(index + 1).padStart(3, '0'),
      '',
      '',
      'POSTED',
      '',
      transactionDate,
      'DUMMY'
    ];
  });

  sheet.getRange(startRow, 1, rows.length, TX_HEADERS.length).setValues(rows);
  sheet.getRange(startRow, TX_COL.DATE, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(startRow, TX_COL.CREATED_AT, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

function assertNoCodeConflicts_(data, materials) {
  const lastRow = data.getLastRow();
  if (lastRow < 2) return;
  const existing = {};
  const values = data.getRange(2, DATA_COL.CODE, lastRow - 1, 1).getValues();
  values.forEach(function (row) {
    const code = String(row[0]).trim().toUpperCase();
    if (code) existing[code] = true;
  });
  materials.forEach(function (material) {
    if (existing[material.code.toUpperCase()]) {
      throw new Error(
        'Kode ' + material.code + ' sudah dipakai. Hapus atau ganti material tersebut dulu.'
      );
    }
  });
}

function rebuildCodeSequencesFromMaterials_() {
  const ss = getSpreadsheet_();
  const data = ss.getSheetByName(SHEET.DATA);
  const codes = ss.getSheetByName(SHEET.CODES);
  if (!data || !codes) return;

  const max = {
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
      if (number > max[prefix]) max[prefix] = number;
    });
  }

  const rows = CODE_PREFIXES.map(function (prefix) {
    return [prefix, max[prefix], new Date()];
  });
  codes.getRange(2, 1, rows.length, 3).setValues(rows);
}

function getDummyMaterials_() {
  return [
    {
      code: 'F-001',
      category: 'F',
      name: 'Elbow 90° PVC 1"',
      spec: '1 inch, 90 derajat',
      bahan: 'PVC',
      unit: 'pcs',
      min: 5
    },
    {
      code: 'F-002',
      category: 'F',
      name: 'Tee PVC 1"',
      spec: '1 inch',
      bahan: 'PVC',
      unit: 'pcs',
      min: 5
    },
    {
      code: 'M-001',
      category: 'M',
      name: 'Bearing 6204',
      spec: '6204',
      bahan: 'Steel',
      unit: 'pcs',
      min: 2
    },
    {
      code: 'M-002',
      category: 'M',
      name: 'Selang Fleksibel 1/2"',
      spec: '1/2 inch',
      bahan: 'PVC',
      unit: 'meter',
      min: 10
    },
    {
      code: 'C-001',
      category: 'C',
      name: 'Seal Tape',
      spec: '0.1 mm x 12 mm',
      bahan: 'PTFE',
      unit: 'roll',
      min: 10
    },
    {
      code: 'C-002',
      category: 'C',
      name: 'Sarung Tangan Nitrile',
      spec: 'Size L',
      bahan: 'Nitrile',
      unit: 'pasang',
      min: 12
    },
    {
      code: 'T-001',
      category: 'T',
      name: 'Kunci Pas 10 mm',
      spec: '10 mm',
      bahan: 'Steel',
      unit: 'pcs',
      min: 1
    },
    {
      code: 'T-002',
      category: 'T',
      name: 'Tang Kombinasi 8"',
      spec: '8 inch',
      bahan: 'Steel',
      unit: 'pcs',
      min: 1
    }
  ];
}

function getDummyTxWwtIn_() {
  return [
    {
      code: 'C-001',
      name: 'Seal Tape',
      qty: 10,
      unit: 'roll',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    },
    {
      code: 'F-001',
      name: 'Elbow 90° PVC 1"',
      qty: 10,
      unit: 'pcs',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    },
    {
      code: 'T-001',
      name: 'Kunci Pas 10 mm',
      qty: 2,
      unit: 'pcs',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    },
    {
      code: 'T-002',
      name: 'Tang Kombinasi 8"',
      qty: 1,
      unit: 'pcs',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    }
  ];
}

function getDummyTxWwtOut_() {
  return [
    {
      code: 'C-001',
      name: 'Seal Tape',
      qty: 2,
      unit: 'roll',
      user: 'Fariz',
      remarks: 'modifikasi line naocl'
    },
    {
      code: 'F-001',
      name: 'Elbow 90° PVC 1"',
      qty: 1,
      unit: 'pcs',
      user: 'Budi',
      remarks: 'perbaikan pompa'
    }
  ];
}

function getDummyTxRwtIn_() {
  return [
    {
      code: 'M-001',
      name: 'Bearing 6204',
      qty: 4,
      unit: 'pcs',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    },
    {
      code: 'C-002',
      name: 'Sarung Tangan Nitrile',
      qty: 20,
      unit: 'pasang',
      user: 'Admin',
      remarks: 'Stok awal dummy'
    }
  ];
}

function getDummyTxRwtOut_() {
  return [
    {
      code: 'M-001',
      name: 'Bearing 6204',
      qty: 1,
      unit: 'pcs',
      user: 'Fariz',
      remarks: 'bearing aus'
    },
    {
      code: 'C-002',
      name: 'Sarung Tangan Nitrile',
      qty: 5,
      unit: 'pasang',
      user: 'Budi',
      remarks: 'shifting'
    }
  ];
}
