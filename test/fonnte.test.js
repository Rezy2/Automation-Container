'use strict';

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
    this.cells = Array.from({ length: 300 }, () => Array(30).fill(''));
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

  deleteRow(row) {
    this.cells.splice(row - 1, 1);
    this.cells.push(Array(30).fill(''));
  }

  getName() {
    return this.name;
  }

  isSheetHidden() {
    return this.hidden;
  }

  hideSheet() {
    this.hidden = true;
    return this;
  }

  showSheet() {
    this.hidden = false;
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

const dataHeaders = [
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

const ss = new FakeSpreadsheet();
sheetWithHeaders(ss, 'Data Material', dataHeaders);
['WWT In', 'WWT Out', 'RWT In', 'RWT Out'].forEach((name) => {
  sheetWithHeaders(ss, name, txHeaders);
});
const config = sheetWithHeaders(ss, 'Config', ['Key', 'Value', 'Description', 'Updated At']);
config.getRange(2, 1, 3, 4).setValues([
  ['FONNTE_ENABLED', 'TRUE', '', new Date()],
  ['FONNTE_ALLOWED_NUMBERS', '628111', '', new Date()],
  ['FONNTE_QUOTE_REPLY', 'TRUE', '', new Date()]
]);
const codes = sheetWithHeaders(ss, 'Code Sequences', ['Prefix', 'Last Number', 'Updated At']);
codes.getRange(2, 1, 4, 3).setValues([
  ['F', 0, new Date()],
  ['M', 0, new Date()],
  ['C', 0, new Date()],
  ['T', 0, new Date()]
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
sheetWithHeaders(ss, 'Inbox', ['WA Message ID', 'Received At', 'Sender', 'Raw Payload', 'Status', 'Error']);
sheetWithHeaders(ss, 'Outbox', [
  'ID',
  'Chat ID',
  'Payload',
  'Status',
  'Attempts',
  'Next Attempt At',
  'Last Error',
  'Created At',
  'Sent At'
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

const properties = new Map([
  ['FONNTE_TOKEN', 'fonnte-token-1'],
  ['FONNTE_WEBHOOK_TOKEN', 'webhook-token-1']
]);
const fetchCalls = [];
let fetchMode = 'ok';

function fakeResponse(code, body) {
  return {
    getResponseCode() {
      return code;
    },
    getContentText() {
      return body;
    }
  };
}

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
        tryLock() {
          return true;
        },
        releaseLock() {}
      };
    }
  },
  PropertiesService: {
    getScriptProperties() {
      return {
        getProperty(key) {
          return properties.has(key) ? properties.get(key) : null;
        },
        deleteProperty(key) {
          properties.delete(key);
        }
      };
    }
  },
  UrlFetchApp: {
    fetch(url, options) {
      fetchCalls.push({ url: url, options: options });
      if (fetchMode === '503') {
        return fakeResponse(503, JSON.stringify({ status: false, reason: 'server busy' }));
      }
      if (fetchMode === 'invalid') {
        return fakeResponse(200, JSON.stringify({ status: false, reason: 'invalid token', requestid: 12 }));
      }
      return fakeResponse(
        200,
        JSON.stringify({
          status: true,
          detail: 'success! message in queue',
          id: ['80367170'],
          process: 'pending',
          requestid: 2937124,
          target: ['628111']
        })
      );
    }
  },
  HtmlService: {
    createHtmlOutput(text) {
      return {
        text: text,
        getContent() {
          return this.text;
        }
      };
    }
  },
  Logger: { log() {} }
};

const code = fs.readFileSync(path.join(__dirname, '..', 'setup.gs'), 'utf8');
const context = vm.createContext(sandbox);
vm.runInContext(code, context, { filename: 'setup.gs' });

function webhookEvent(payload, token) {
  return {
    parameter: token === undefined ? {} : { token: token },
    postData: { contents: JSON.stringify(payload) }
  };
}

function dispatch(payload, token) {
  context.webhookEvent = webhookEvent(payload, token);
  return vm.runInContext('handleFonnteWebhook_(webhookEvent)', context);
}

function post(payload, token) {
  context.webhookEvent = webhookEvent(payload, token);
  return vm.runInContext('doPost(webhookEvent)', context).getContent();
}

function payload(inboxid, message, overrides) {
  return Object.assign(
    {
      device: '628999',
      sender: '628111',
      message: message,
      name: 'Fariz',
      timestamp: 1730000000 + Number(inboxid || 0),
      inboxid: inboxid
    },
    overrides || {}
  );
}

function parseForm(body) {
  const result = {};
  String(body || '')
    .split('&')
    .forEach(function (pair) {
      const parts = pair.split('=');
      result[decodeURIComponent(parts[0])] = decodeURIComponent(parts.slice(1).join('=') || '');
    });
  return result;
}

function outboxRow(index) {
  const sheet = ss.getSheetByName('Outbox');
  return sheet.cells[index][0] ? sheet.cells[index] : null;
}

const first = dispatch(payload(1001, 'in pressure gauge 5 pcs ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(first.ok, true);
assert.strictEqual(first.status, 'PENDING');
assert.strictEqual(ss.getSheetByName('Drafts').getLastRow(), 2);
assert.strictEqual(ss.getSheetByName('Drafts').cells[1][13], 'PENDING');
assert.strictEqual(ss.getSheetByName('Inbox').cells[1][4], 'PROCESSED');
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), 2);
assert.strictEqual(outboxRow(1)[3], 'SENT');
assert.strictEqual(outboxRow(1)[4], 1);

assert.strictEqual(fetchCalls.length, 1);
assert.strictEqual(fetchCalls[0].url, 'https://api.fonnte.com/send');
assert.strictEqual(fetchCalls[0].options.headers.Authorization, 'fonnte-token-1');
const sent = parseForm(fetchCalls[0].options.payload);
assert.strictEqual(sent.target, '628111');
assert.match(sent.message, /^Material: Pressure gauge/);
assert.strictEqual(sent.inboxid, '1001');
assert.strictEqual(sent.preview, 'false');
assert.strictEqual(sent.connectOnly, 'false');

const duplicate = dispatch(payload(1001, 'in pressure gauge 5 pcs ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(duplicate.ok, true);
assert.strictEqual(duplicate.duplicate, true);
assert.strictEqual(ss.getSheetByName('Drafts').getLastRow(), 2);
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), 2);
assert.strictEqual(fetchCalls.length, 1);

const commit = dispatch(payload(1002, 'OK'), 'webhook-token-1');
assert.strictEqual(commit.ok, true);
assert.strictEqual(commit.status, 'POSTED');
assert.strictEqual(ss.getSheetByName('RWT In').getLastRow(), 2);
assert.strictEqual(ss.getSheetByName('RWT In').cells[1][4], 5);
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), 3);
assert.strictEqual(outboxRow(2)[3], 'SENT');
const commitForm = parseForm(fetchCalls[1].options.payload);
assert.strictEqual(commitForm.inboxid, '1002');
assert.match(commitForm.message, /^Tersimpan\./);

const dupCommit = dispatch(payload(1002, 'OK'), 'webhook-token-1');
assert.strictEqual(dupCommit.ok, true);
assert.strictEqual(dupCommit.duplicate, true);
assert.strictEqual(ss.getSheetByName('RWT In').getLastRow(), 2);

const directResponse = post(payload(1003, 'in seal tape 1 roll ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(directResponse, 'OK');
assert.strictEqual(ss.getSheetByName('Drafts').getLastRow(), 3);

const unauthorizedResponse = post(payload(1004, 'OK'), 'wrong-token');
assert.strictEqual(unauthorizedResponse, 'OK');
const outboxBeforeUnauthorized = ss.getSheetByName('Outbox').getLastRow();
const unauthorized = dispatch(payload(1004, 'OK'), 'wrong-token');
assert.strictEqual(unauthorized.ok, false);
assert.strictEqual(unauthorized.error, 'unauthorized');
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), outboxBeforeUnauthorized);

const group = dispatch(payload(1005, 'in seal tape 2 roll ke RWT oleh Fariz', { member: '628111' }), 'webhook-token-1');
assert.strictEqual(group.ignored, true);
assert.strictEqual(group.reason, 'not_dm');

const groupBySender = dispatch(
  payload(1006, 'in seal tape 2 roll ke RWT oleh Fariz', { sender: '123@g.us', member: '' }),
  'webhook-token-1'
);
assert.strictEqual(groupBySender.ignored, true);
assert.strictEqual(groupBySender.reason, 'not_dm');

const media = dispatch(
  payload(1007, 'in seal tape 2 roll ke RWT oleh Fariz', { url: 'https://example.com/photo.jpg' }),
  'webhook-token-1'
);
assert.strictEqual(media.ignored, true);
assert.strictEqual(media.reason, 'not_text');

const otherNumber = dispatch(
  payload(1008, 'in seal tape 2 roll ke RWT oleh Fariz', { sender: '628222' }),
  'webhook-token-1'
);
assert.strictEqual(otherNumber.ignored, true);
assert.strictEqual(otherNumber.reason, 'sender_not_allowed');

const outboxBeforeNoInbox = ss.getSheetByName('Outbox').getLastRow();
const noInbox = dispatch(payload(0, 'in seal tape 2 roll ke RWT oleh Fariz', { inboxid: 0 }), 'webhook-token-1');
assert.strictEqual(noInbox.ok, true);
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), outboxBeforeNoInbox + 1);
const noInboxForm = parseForm(fetchCalls[fetchCalls.length - 1].options.payload);
assert.strictEqual(noInboxForm.inboxid, undefined);
const noInboxDuplicate = dispatch(payload(0, 'in seal tape 2 roll ke RWT oleh Fariz', { inboxid: 0 }), 'webhook-token-1');
assert.strictEqual(noInboxDuplicate.duplicate, true);
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), outboxBeforeNoInbox + 1);

config.cells[3][1] = 'FALSE';
const noQuote = dispatch(payload(1009, 'in seal tape 3 roll ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(noQuote.ok, true);
const noQuoteForm = parseForm(fetchCalls[fetchCalls.length - 1].options.payload);
assert.strictEqual(noQuoteForm.inboxid, undefined);
config.cells[3][1] = 'TRUE';

fetchMode = '503';
const retryPending = dispatch(payload(1010, 'in seal tape 4 roll ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(retryPending.ok, true);
const outboxRetryIndex = ss.getSheetByName('Outbox').getLastRow() - 1;
assert.strictEqual(outboxRow(outboxRetryIndex)[3], 'PENDING');
assert.strictEqual(outboxRow(outboxRetryIndex)[4], 1);
assert.ok(String(outboxRow(outboxRetryIndex)[6]).includes('server busy'));

fetchMode = 'ok';
ss.getSheetByName('Outbox').cells[outboxRetryIndex][5] = new Date(Date.now() - 1000);
vm.runInContext('processOutbox()', context);
assert.strictEqual(outboxRow(outboxRetryIndex)[3], 'SENT');
assert.strictEqual(outboxRow(outboxRetryIndex)[4], 2);

fetchMode = 'invalid';
const invalidToken = dispatch(payload(1011, 'in seal tape 5 roll ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(invalidToken.ok, true);
const outboxInvalidIndex = ss.getSheetByName('Outbox').getLastRow() - 1;
assert.strictEqual(outboxRow(outboxInvalidIndex)[3], 'FAILED');
assert.strictEqual(outboxRow(outboxInvalidIndex)[4], 1);
assert.ok(String(outboxRow(outboxInvalidIndex)[6]).includes('invalid token'));

fetchMode = 'ok';
const outboxBeforeDisabled = ss.getSheetByName('Outbox').getLastRow();
config.cells[1][1] = 'FALSE';
const disabled = dispatch(payload(1012, 'in seal tape 6 roll ke RWT oleh Fariz'), 'webhook-token-1');
assert.strictEqual(disabled.ok, true);
assert.strictEqual(disabled.ignored, true);
assert.strictEqual(disabled.reason, 'disabled');
assert.strictEqual(ss.getSheetByName('Outbox').getLastRow(), outboxBeforeDisabled);
config.cells[1][1] = 'TRUE';

const normalize = vm.runInContext('normalizeFonnteNumber_', context);
assert.strictEqual(normalize('+62 811-2222'), '628112222');
assert.strictEqual(normalize('0811-2222'), '628112222');
assert.strictEqual(normalize('8112222'), '628112222');

const configLastRow = config.getLastRow();
config.appendRow(['OPENWA_ENABLED', 'TRUE', '', new Date()]);
properties.set('OPENWA_API_KEY', 'legacy-token');
vm.runInContext('cleanupLegacyOpenWA_()', context);
assert.strictEqual(config.getLastRow(), configLastRow);
assert.strictEqual(properties.has('OPENWA_API_KEY'), false);

console.log('fonnte tests passed');
