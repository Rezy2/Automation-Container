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
    this.cells = Array.from({ length: 200 }, () => Array(30).fill(''));
    this.hidden = false;
  }

  ensure(row, col) {
    while (this.cells.length < row) this.cells.push(Array(30).fill(''));
    this.cells.forEach((line) => {
      while (line.length < col) line.push('');
    });
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
    this.getRange(this.getLastRow() + 1, 1, 1, values.length).setValues([values]);
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

function materialRow(number, code, name, category, unit, min) {
  const row = Array(19).fill('');
  row[0] = number;
  row[2] = code;
  row[3] = name;
  row[4] = category;
  row[7] = unit;
  row[8] = min;
  row[17] = 'ACTIVE';
  row[18] = false;
  return row;
}

const code = fs.readFileSync(path.join(__dirname, '..', 'setup.gs'), 'utf8');

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
data.getRange(2, 1, 3, 19).setValues([
  materialRow(1, 'F-001', 'Elbow 90° PVC 1"', 'F', 'pcs', 5),
  materialRow(2, 'C-001', 'Seal Tape', 'C', 'roll', 10),
  materialRow(3, 'T-001', 'Kunci Pas 10 mm', 'T', 'pcs', 1)
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
  ['M', 0, new Date()],
  ['C', 1, new Date()],
  ['T', 1, new Date()]
]);
sheetWithHeaders(ss, 'Material Alias', ['Alias', 'Normalized Alias', 'Material Code', 'Created At']);
sheetWithHeaders(ss, 'Drafts', [
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
]);
sheetWithHeaders(ss, 'Audit Log', [
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
]);

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
vm.runInContext(code, context, { filename: 'app.gs' });

function incoming(text, meta) {
  context.rawMessage = text;
  context.messageMeta = meta || {};
  return vm.runInContext('handleIncomingMessage_(rawMessage, messageMeta)', context);
}

function draftCount() {
  const sheet = ss.getSheetByName('Drafts');
  return Math.max(0, sheet.getLastRow() - 1);
}

function lastDraftStatus() {
  const sheet = ss.getSheetByName('Drafts');
  return sheet.cells[sheet.getLastRow() - 1][13];
}

function draftStatusById(draftId) {
  const sheet = ss.getSheetByName('Drafts');
  for (let row = 1; row < sheet.getLastRow(); row++) {
    if (sheet.cells[row][0] === draftId) return sheet.cells[row][13];
  }
  return null;
}

const draft1 = incoming('in pressure gauge 5 pcs ke RWT oleh Fariz', {
  chatId: 'C1',
  sender: 'Fariz',
  waMessageId: 'WA-1'
});
assert.strictEqual(draft1.status, 'PENDING');
assert.match(draft1.draftId, /^DRF-/);
assert.strictEqual(draft1.message.split('\n').slice(0, 6).join('\n'), [
  'Material: Pressure gauge',
  'Material Code: M-001',
  'Qty: 5 pcs',
  'Container: RWT',
  'User: Fariz',
  'Catatan: -'
].join('\n'));
assert.strictEqual(data.getLastRow(), 5);
assert.strictEqual(data.cells[4][2], 'M-001');
assert.strictEqual(draftCount(), 1);

const commit1 = incoming('OK', {
  chatId: 'C1',
  sender: 'Fariz',
  waMessageId: 'WA-2',
  draftId: draft1.draftId
});
assert.strictEqual(commit1.status, 'POSTED');
assert.match(commit1.message, /^Tersimpan\./);
assert.ok(commit1.message.includes('Saldo RWT: 5 pcs'));
assert.strictEqual(ss.getSheetByName('RWT In').getLastRow(), 2);
assert.strictEqual(ss.getSheetByName('RWT In').cells[1][4], 5);
assert.strictEqual(lastDraftStatus(), 'POSTED');

const repeat = incoming('OK', {
  chatId: 'C1',
  sender: 'Fariz',
  draftId: draft1.draftId
});
assert.strictEqual(repeat.status, 'POSTED');
assert.ok(repeat.message.includes('sudah tersimpan'));
assert.strictEqual(ss.getSheetByName('RWT In').getLastRow(), 2);

const draft2 = incoming('out pressure gauge 3 pcs dari RWT oleh Budi', {
  chatId: 'C2',
  sender: 'Budi',
  waMessageId: 'WA-3'
});
assert.strictEqual(draft2.status, 'PENDING');
assert.ok(draft2.message.includes('Material Code: M-001'));

const commit2 = incoming('OK', {
  chatId: 'C2',
  sender: 'Budi',
  draftId: draft2.draftId
});
assert.strictEqual(commit2.status, 'POSTED');
assert.strictEqual(commit2.balanceAfter, 2);
assert.strictEqual(ss.getSheetByName('RWT Out').getLastRow(), 2);

const draft3 = incoming('out pressure gauge 3 pcs dari RWT oleh Budi', {
  chatId: 'C3',
  sender: 'Budi',
  waMessageId: 'WA-5'
});
assert.strictEqual(draft3.status, 'PENDING');

let stockError = null;
try {
  incoming('OK', {
    chatId: 'C3',
    sender: 'Budi',
    draftId: draft3.draftId
  });
} catch (error) {
  stockError = error;
}
assert.ok(stockError);
assert.match(stockError.message, /Stok RWT/);
assert.strictEqual(ss.getSheetByName('RWT Out').getLastRow(), 2);
assert.strictEqual(lastDraftStatus(), 'PENDING');

const draftsBeforeMismatch = draftCount();
const mismatch = incoming('out seal tape 1 pcs dari RWT oleh Fariz', {
  chatId: 'C4',
  sender: 'Fariz',
  waMessageId: 'WA-6'
});
assert.strictEqual(mismatch.status, 'NEEDS_CLARIFICATION');
assert.ok(mismatch.message.includes('master: roll'));
assert.strictEqual(draftCount(), draftsBeforeMismatch);

const dataRowsBeforeMissing = data.getLastRow();
const missing = incoming('out connector 3/4" 1 pcs untuk hose fariz', {
  chatId: 'C5',
  sender: 'Fariz',
  waMessageId: 'WA-7'
});
assert.strictEqual(missing.status, 'NEEDS_CLARIFICATION');
assert.ok(missing.message.includes('container (WWT/RWT)'));
assert.ok(missing.message.includes('nama user'));
assert.strictEqual(data.getLastRow(), dataRowsBeforeMissing);

const ambiguous = incoming('out seal tape 1 roll dari WWT ke RWT oleh Fariz', {
  chatId: 'C6',
  sender: 'Fariz',
  waMessageId: 'WA-8'
});
assert.strictEqual(ambiguous.status, 'NEEDS_CLARIFICATION');
assert.ok(ambiguous.message.includes('Container ambigu: WWT / RWT'));

const draft4 = incoming('in pressure gauge 5 pcs ke RWT oleh Fariz', {
  chatId: 'C7',
  sender: 'Fariz',
  waMessageId: 'WA-9'
});
assert.strictEqual(draft4.status, 'PENDING');

const revision = incoming('REVISI in pressure gauge 6 pcs ke RWT oleh Fariz', {
  chatId: 'C7',
  sender: 'Fariz',
  waMessageId: 'WA-10',
  draftId: draft4.draftId
});
assert.strictEqual(revision.status, 'PENDING');
assert.notStrictEqual(revision.draftId, draft4.draftId);
assert.ok(revision.message.includes('Qty: 6 pcs'));
assert.strictEqual(draftStatusById(draft4.draftId), 'REVISI');

const revisionCommit = incoming('OK', {
  chatId: 'C7',
  sender: 'Fariz',
  draftId: revision.draftId
});
assert.strictEqual(revisionCommit.status, 'POSTED');
assert.strictEqual(revisionCommit.balanceAfter, 8);

const draft5 = incoming('in pressure gauge 2 pcs ke RWT oleh Fariz', {
  chatId: 'C8',
  sender: 'Fariz',
  waMessageId: 'WA-11'
});
const cancelled = incoming('REVISI', {
  chatId: 'C8',
  sender: 'Fariz',
  draftId: draft5.draftId
});
assert.strictEqual(cancelled.status, 'REVISI');
const cancelledOk = incoming('OK', {
  chatId: 'C8',
  sender: 'Fariz',
  draftId: draft5.draftId
});
assert.strictEqual(cancelledOk.status, 'REVISI');
assert.ok(cancelledOk.message.includes('sudah dibatalkan'));

const draft6 = incoming('in pressure gauge 2 pcs ke RWT oleh Fariz', {
  chatId: 'C9',
  sender: 'Fariz',
  waMessageId: 'WA-12'
});
assert.strictEqual(draft6.status, 'PENDING');
const latestCommit = incoming('OK', {
  chatId: 'C9',
  sender: 'Fariz'
});
assert.strictEqual(latestCommit.status, 'POSTED');
assert.strictEqual(latestCommit.balanceAfter, 10);

const firstDuplicate = incoming('in pressure gauge 1 pcs ke RWT oleh Fariz', {
  chatId: 'C10',
  sender: 'Fariz',
  waMessageId: 'WA-SAME'
});
incoming('OK', {
  chatId: 'C10',
  sender: 'Fariz',
  draftId: firstDuplicate.draftId
});
const beforeDuplicateTx = ss.getSheetByName('RWT In').getLastRow();
const secondDuplicate = incoming('in pressure gauge 1 pcs ke RWT oleh Fariz', {
  chatId: 'C10',
  sender: 'Fariz',
  waMessageId: 'WA-SAME'
});
const duplicate = incoming('OK', {
  chatId: 'C10',
  sender: 'Fariz',
  draftId: secondDuplicate.draftId
});
assert.strictEqual(secondDuplicate.draftId, firstDuplicate.draftId);
assert.strictEqual(duplicate.status, 'POSTED');
assert.ok(duplicate.message.includes('sudah tersimpan'));
assert.strictEqual(ss.getSheetByName('RWT In').getLastRow(), beforeDuplicateTx);

const draft7 = incoming('in pressure gauge 1 pcs ke RWT oleh Fariz', {
  chatId: 'C11',
  sender: 'Fariz',
  waMessageId: 'WA-13'
});
const invalidReply = incoming('halo', {
  chatId: 'C11',
  sender: 'Fariz',
  draftId: draft7.draftId
});
assert.strictEqual(invalidReply.status, 'WAITING_CONFIRMATION');
assert.ok(invalidReply.message.includes('Balas OK'));

assert.strictEqual(ss.getSheetByName('Audit Log').getLastRow() > 1, true);
console.log('draft preview and commit tests passed');
