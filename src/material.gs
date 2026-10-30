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
