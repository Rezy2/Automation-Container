const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
    this.sheet.ensure(row + numRows - 1, col + numCols - 1);
  }

  getValues() {
    const rows = [];
    for (let r = 0; r < this.numRows; r++) {
      rows.push(this.sheet.cells[this.row - 1 + r].slice(this.col - 1, this.col - 1 + this.numCols));
    }
    return rows;
  }

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        this.sheet.cells[this.row - 1 + r][this.col - 1 + c] = values[r][c];
      }
    }
    return this;
  }

  getValue() {
    return this.sheet.cells[this.row - 1][this.col - 1];
  }

  setValue(value) {
    this.sheet.cells[this.row - 1][this.col - 1] = value;
    return this;
  }

  setFormula(value) {
    return this.setValue(value);
  }

  setFormulas(values) {
    return this.setValues(values);
  }

  setNumberFormat() {
    return this;
  }
}

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.maxRows = 100;
    this.maxCols = 30;
    this.cells = Array.from({ length: this.maxRows }, () => Array(this.maxCols).fill(''));
    this.hidden = false;
  }

  ensure(row, col) {
    while (this.cells.length < row) this.cells.push(Array(this.maxCols).fill(''));
    if (row > this.maxRows) this.maxRows = row;
    if (col > this.maxCols) {
      this.cells.forEach((line) => {
        while (line.length < col) line.push('');
      });
      this.maxCols = col;
    }
  }

  getRange(row, col, numRows, numCols) {
    return new FakeRange(this, row, col, numRows || 1, numCols || 1);
  }

  getLastRow() {
    for (let row = this.cells.length - 1; row >= 0; row--) {
      if (this.cells[row].some((value) => value !== '' && value !== null && value !== undefined)) {
        return row + 1;
      }
    }
    return 0;
  }

  appendRow(values) {
    const row = this.getLastRow() + 1;
    this.getRange(row, 1, 1, values.length).setValues([values]);
  }

  getName() {
    return this.name;
  }

  isSheetHidden() {
    return this.hidden;
  }

  showSheet() {
    this.hidden = false;
    return this;
  }

  hideSheet() {
    this.hidden = true;
    return this;
  }
}

class FakeSpreadsheet {
  constructor() {
    this.sheets = [];
  }

  insertSheet(name) {
    const sheet = new FakeSheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  getSheetByName(name) {
    return this.sheets.find((sheet) => sheet.name === name) || null;
  }
}

function sheetWithHeaders(ss, name, headers) {
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function materialRow(number, code, name, category, unit, min, spec, bahan) {
  const row = Array(19).fill('');
  row[0] = number;
  row[2] = code;
  row[3] = name;
  row[4] = category;
  row[5] = spec || '';
  row[6] = bahan || '';
  row[7] = unit;
  row[8] = min;
  row[17] = 'ACTIVE';
  row[18] = false;
  return row;
}

const code = [
  fs.readFileSync(path.join(__dirname, '..', 'src', 'config.gs'), 'utf8'),
  fs.readFileSync(path.join(__dirname, '..', 'src', 'parser.gs'), 'utf8'),
  fs.readFileSync(path.join(__dirname, '..', 'src', 'setup.gs'), 'utf8'),
  fs.readFileSync(path.join(__dirname, '..', 'src', 'material.gs'), 'utf8')
].join('\n\n');

const ss = new FakeSpreadsheet();
const data = sheetWithHeaders(ss, 'Data Material', [
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
]);
data.getRange(2, 1, 4, 19).setValues([
  materialRow(1, 'F-001', 'Elbow 90° PVC 1"', 'F', 'pcs', 5, '1 inch', 'PVC'),
  materialRow(2, 'C-001', 'Seal Tape', 'C', 'roll', 10, '0.1 mm', 'PTFE'),
  materialRow(3, 'M-001', 'Bearing 6204', 'M', 'pcs', 2, '6204', 'Steel'),
  materialRow(4, 'T-001', 'Kunci Pas 10 mm', 'T', 'pcs', 1, '10 mm', 'Steel')
]);

const txHeaders = [
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
['WWT In', 'WWT Out', 'RWT In', 'RWT Out'].forEach((name) => {
  sheetWithHeaders(ss, name, txHeaders);
});

sheetWithHeaders(ss, 'Config', ['Key', 'Value', 'Description', 'Updated At']);
const codes = sheetWithHeaders(ss, 'Code Sequences', ['Prefix', 'Last Number', 'Updated At']);
codes.getRange(2, 1, 4, 3).setValues([
  ['F', 1, new Date()],
  ['M', 1, new Date()],
  ['C', 1, new Date()],
  ['T', 1, new Date()]
]);
sheetWithHeaders(ss, 'Material Alias', ['Alias', 'Normalized Alias', 'Material Code', 'Created At']);

const sandbox = {
  console,
  SpreadsheetApp: {
    openById(id) {
      assert.strictEqual(id, '19KlQdT11UzJlUOjZ_CMMUMEK-7GgRt2bJEA4rfIhGPU');
      return ss;
    },
    flush() {}
  },
  LockService: {
    getScriptLock() {
      return {
        waitLock() {},
        releaseLock() {}
      };
    }
  },
  Logger: {
    log() {}
  }
};

const context = vm.createContext(sandbox);
vm.runInContext(code, context, { filename: 'material.gs' });

function parse(text) {
  return vm.runInContext('parseMessage(' + JSON.stringify(text) + ')', context);
}

function resolve(parsed) {
  context.parsedInput = parsed;
  return vm.runInContext('resolveParsedMaterial_(parsedInput)', context);
}

function ensure(text, unit) {
  return vm.runInContext(
    'ensureMaterialByText_(' + JSON.stringify(text) + ', ' + JSON.stringify(unit) + ')',
    context
  );
}

const pressure = parse('out pressure gauge 1 pcs dari RWT oleh Fariz');
assert.strictEqual(pressure.materialText, 'pressure gauge');
assert.strictEqual(pressure.category, 'M');
assert.strictEqual(pressure.autoCreate, true);
assert.strictEqual(pressure.needsClarification, false);

const created = resolve(pressure);
assert.strictEqual(created.materialCode, 'M-002');
assert.strictEqual(created.materialName, 'Pressure gauge');
assert.strictEqual(created.autoCreated, true);
assert.strictEqual(created.autoCreate, false);
assert.strictEqual(created.needsMaster, false);
assert.strictEqual(created.needsClarification, false);

assert.strictEqual(data.getLastRow(), 6);
assert.strictEqual(data.cells[5][0], 5);
assert.strictEqual(data.cells[5][2], 'M-002');
assert.strictEqual(data.cells[5][3], 'Pressure gauge');
assert.strictEqual(data.cells[5][4], 'M');
assert.strictEqual(data.cells[5][7], 'pcs');
assert.strictEqual(data.cells[5][8], 0);
assert.strictEqual(data.cells[5][17], 'ACTIVE');
assert.strictEqual(data.cells[5][18], false);
assert(String(data.cells[5][9]).includes("SUMIF('WWT In'!$C$2:$C"));
assert.strictEqual(codes.cells[2][1], 2);
assert.strictEqual(ss.getSheetByName('Material Alias').cells[1][0], 'pressure gauge');
assert.strictEqual(ss.getSheetByName('Material Alias').cells[1][1], 'pressure gauge');
assert.strictEqual(ss.getSheetByName('Material Alias').cells[1][2], 'M-002');

const again = resolve(parse('out pressure gauge 1 pcs dari RWT oleh Fariz'));
assert.strictEqual(again.materialCode, 'M-002');
assert.strictEqual(again.autoCreated, false);
assert.strictEqual(data.getLastRow(), 6);

const pipe = ensure('kunci pipa 8"', 'pcs');
assert.strictEqual(pipe.code, 'T-002');
assert.strictEqual(pipe.category, 'T');
assert.strictEqual(pipe.created, true);

codes.cells[1][1] = 0;
const connector = ensure('connector 3/4"', 'pcs');
assert.strictEqual(connector.code, 'F-002');
assert.strictEqual(connector.category, 'F');
assert.strictEqual(connector.created, true);
assert.strictEqual(codes.cells[1][1], 2);

const duplicate = ensure('Kunci Pipa 8"', 'pcs');
assert.strictEqual(duplicate.code, 'T-002');
assert.strictEqual(duplicate.created, false);
assert.strictEqual(data.getLastRow(), 8);

const missingUnit = parse('out laser pointer dari RWT oleh Fariz');
assert.strictEqual(missingUnit.unit, '');
assert.strictEqual(missingUnit.autoCreate, true);
const unresolved = resolve(missingUnit);
assert.strictEqual(unresolved.materialCode, '');
assert.strictEqual(unresolved.autoCreate, true);
assert.strictEqual(data.getLastRow(), 8);

const mismatch = resolve(parse('out elbow 90 pvc 1" 2 roll dari WWT oleh Fariz'));
assert.strictEqual(mismatch.materialCode, 'F-001');
assert.strictEqual(mismatch.unitMismatch, true);
assert.strictEqual(mismatch.needsClarification, true);

assert.throws(function () {
  ensure('material tanpa satuan', '');
}, /Satuan material wajib/);

assert.strictEqual(vm.runInContext('classifyMaterial_("sikat kawat").category', context), 'T');
assert.strictEqual(vm.runInContext('classifyMaterial_("seal tape").category', context), 'C');
assert.strictEqual(vm.runInContext('classifyMaterial_("connector 1/2").category', context), 'F');
assert.strictEqual(vm.runInContext('classifyMaterial_("laser pointer").category', context), 'M');

console.log('material auto-create tests passed');
