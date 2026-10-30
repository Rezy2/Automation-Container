const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const parserPath = path.join(__dirname, '..', 'src', 'parser.gs');
const outputPath = path.join(__dirname, 'parser-corpus.csv');
const code = fs.readFileSync(parserPath, 'utf8');
const context = vm.createContext({ console });
vm.runInContext(code, context, { filename: 'parser.gs' });
const parse = vm.runInContext('parseMessage', context);

const cases = [
  {
    group: 'OUT',
    raw: 'kunci pipa 8" oleh Fariz dari kontainer wwt 1 pcs untuk kerja',
    direction: 'OUT',
    material: 'kunci pipa 8"',
    qty: 1,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'kerja'
  },
  {
    group: 'OUT',
    raw: 'ambil seal tape 1 pcs untuk modif line kontainer RWT oleh ismat',
    direction: 'OUT',
    material: 'seal tape',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'modif line',
    code: 'C-001',
    mismatch: true
  },
  {
    group: 'OUT',
    raw: 'out connector 3/4" 1 pcs untuk hose fariz (container tidak disebut)',
    direction: 'OUT',
    material: 'connector 3/4"',
    qty: 1,
    unit: 'pcs',
    container: '',
    user: '',
    note: 'hose fariz',
    missing: ['container', 'user']
  },
  {
    group: 'OUT',
    raw: 'Ambil tang kombinasi 1 pcs oleh Fariz dari container WWT untuk perbaikan panel.',
    direction: 'OUT',
    material: 'tang kombinasi',
    qty: 1,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'perbaikan panel',
    code: 'T-002'
  },
  {
    group: 'OUT',
    raw: 'Keluar kunci inggris 12 inch, 1 pcs, RWT, untuk kerja oleh Ismat.',
    direction: 'OUT',
    material: 'kunci inggris 12 inch',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'kerja'
  },
  {
    group: 'OUT',
    raw: 'Dimas ambil hose clamp 3/4" sebanyak 2 pcs dari WWT untuk perbaikan hose.',
    direction: 'OUT',
    material: 'hose clamp 3/4"',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    note: 'perbaikan hose'
  },
  {
    group: 'OUT',
    raw: 'Fariz mengambil 1 pcs gasket pump dari container WWT untuk maintenance.',
    direction: 'OUT',
    material: 'gasket pump',
    qty: 1,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'maintenance'
  },
  {
    group: 'OUT',
    raw: 'Out 2 pcs lampu portable dari RWT untuk inspeksi area, oleh Ismat.',
    direction: 'OUT',
    material: 'lampu portable',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'inspeksi area'
  },
  {
    group: 'OUT',
    raw: 'Ambil 1 roll teflon tape dari container RWT untuk modifikasi line oleh Fariz.',
    direction: 'OUT',
    material: 'teflon tape',
    qty: 1,
    unit: 'roll',
    container: 'RWT',
    user: 'Fariz',
    note: 'modifikasi line',
    code: 'C-001'
  },
  {
    group: 'OUT',
    raw: 'Keluar 3 pcs obeng plus dari WWT, peminta Dimas, untuk maintenance.',
    direction: 'OUT',
    material: 'obeng plus',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    note: 'maintenance'
  },
  {
    group: 'OUT',
    raw: 'Ismat ambil 1 pcs adjustable wrench dari kontainer RWT untuk perbaikan valve.',
    direction: 'OUT',
    material: 'adjustable wrench',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'perbaikan valve'
  },
  {
    group: 'OUT',
    raw: 'Mohon keluarkan 2 pcs safety goggles dari WWT untuk pekerjaan hari ini, Fariz.',
    direction: 'OUT',
    material: 'safety goggles',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'pekerjaan hari ini'
  },
  {
    group: 'OUT',
    raw: 'Out 4 pcs cable tie ukuran besar dari RWT oleh Dimas untuk rapikan kabel.',
    direction: 'OUT',
    material: 'cable tie ukuran besar',
    qty: 4,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas',
    note: 'rapikan kabel'
  },
  {
    group: 'OUT',
    raw: 'Fariz minta 1 pcs kunci L 6 mm dari container WWT buat kerja.',
    direction: 'OUT',
    material: 'kunci L 6 mm',
    qty: 1,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'kerja'
  },
  {
    group: 'OUT',
    raw: 'Ambil 2 pcs sarung tangan kerja dari RWT oleh Ismat.',
    direction: 'OUT',
    material: 'sarung tangan kerja',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat'
  },
  {
    group: 'OUT',
    raw: 'Keluarkan 1 pcs cutter dari WWT untuk potong gasket, peminta Dimas.',
    direction: 'OUT',
    material: 'cutter',
    qty: 1,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    note: 'potong gasket'
  },
  {
    group: 'OUT',
    raw: 'Dimas mengambil 1 roll isolasi listrik dari kontainer RWT untuk perbaikan kabel.',
    direction: 'OUT',
    material: 'isolasi listrik',
    qty: 1,
    unit: 'roll',
    container: 'RWT',
    user: 'Dimas',
    note: 'perbaikan kabel'
  },
  {
    group: 'OUT',
    raw: 'Out 3 pcs baut stainless M10 dari WWT untuk perbaikan bracket oleh Fariz.',
    direction: 'OUT',
    material: 'baut stainless M10',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'perbaikan bracket'
  },
  {
    group: 'OUT',
    raw: 'Ambil 1 pcs tang potong dari RWT untuk pekerjaan instrument, Ismat.',
    direction: 'OUT',
    material: 'tang potong',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'pekerjaan instrument'
  },
  {
    group: 'OUT',
    raw: 'Keluar 2 pcs kunci pas 14 mm dari container WWT, untuk maintenance oleh Dimas.',
    direction: 'OUT',
    material: 'kunci pas 14 mm',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    note: 'maintenance'
  },
  {
    group: 'OUT',
    raw: 'Fariz ambil 1 pcs pressure gauge dari RWT untuk pengecekan line.',
    direction: 'OUT',
    material: 'pressure gauge',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Fariz',
    note: 'pengecekan line'
  },
  {
    group: 'OUT',
    raw: 'Mohon ambil 5 pcs masker dari WWT untuk pekerjaan, oleh Ismat.',
    direction: 'OUT',
    material: 'masker',
    qty: 5,
    unit: 'pcs',
    container: 'WWT',
    user: 'Ismat',
    note: 'pekerjaan'
  },
  {
    group: 'OUT',
    raw: 'Out 1 pcs sealant gun dari container RWT untuk sealing line, Dimas.',
    direction: 'OUT',
    material: 'sealant gun',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas',
    note: 'sealing line'
  },
  {
    group: 'OUT',
    raw: 'Keluar 2 roll lakban dari WWT buat packing material oleh Fariz.',
    direction: 'OUT',
    material: 'lakban',
    qty: 2,
    unit: 'roll',
    container: 'WWT',
    user: 'Fariz',
    note: 'packing material'
  },
  {
    group: 'OUT',
    raw: 'Ismat minta 1 pcs palu 1 kg dari kontainer RWT untuk kerja.',
    direction: 'OUT',
    material: 'palu 1 kg',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'kerja'
  },
  {
    group: 'OUT',
    raw: 'Ambil 6 pcs elektroda las dari WWT untuk repair support, Dimas.',
    direction: 'OUT',
    material: 'elektroda las',
    qty: 6,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    note: 'repair support'
  },
  {
    group: 'OUT',
    raw: 'Out 1 pcs kunci pipa 10" dari RWT oleh Fariz untuk bongkar fitting.',
    direction: 'OUT',
    material: 'kunci pipa 10"',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Fariz',
    note: 'bongkar fitting'
  },
  {
    group: 'OUT',
    raw: 'Keluar 2 pcs sikat kawat dari container WWT untuk bersihkan flange, Ismat.',
    direction: 'OUT',
    material: 'sikat kawat',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Ismat',
    note: 'bersihkan flange'
  },
  {
    group: 'OUT',
    raw: 'Dimas ambil 1 pcs test pen dari RWT untuk pengecekan listrik.',
    direction: 'OUT',
    material: 'test pen',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas',
    note: 'pengecekan listrik'
  },
  {
    group: 'OUT',
    raw: 'Keluarkan 3 pcs lap majun dari WWT untuk cleaning oleh Fariz.',
    direction: 'OUT',
    material: 'lap majun',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    note: 'cleaning'
  },
  {
    group: 'OUT',
    raw: 'Ambil 1 pcs connector 1/2 inch dari container RWT untuk hose, Ismat.',
    direction: 'OUT',
    material: 'connector 1/2 inch',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat',
    note: 'hose'
  },
  {
    group: 'IN',
    raw: 'sapu masuk container wwt 2 pcs dari Dimas',
    direction: 'IN',
    material: 'sapu',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'In sikat 1 pcs container RWT Dimas',
    direction: 'IN',
    material: 'sikat',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Dimas memasukkan 3 pcs tang kombinasi ke container WWT.',
    direction: 'IN',
    material: 'tang kombinasi',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas',
    code: 'T-002'
  },
  {
    group: 'IN',
    raw: 'Barang masuk 2 pcs kunci inggris 12" ke container RWT dari Ismat.',
    direction: 'IN',
    material: 'kunci inggris 12"',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Terima hose clamp 3/4" sebanyak 5 pcs untuk container RWT dari Dimas.',
    direction: 'IN',
    material: 'hose clamp 3/4"',
    qty: 5,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Masuk 4 pcs gasket pump ke WWT, diserahkan Fariz.',
    direction: 'IN',
    material: 'gasket pump',
    qty: 4,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Dimas menyerahkan 2 pcs lampu portable untuk masuk container RWT.',
    direction: 'IN',
    material: 'lampu portable',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: '3 roll teflon tape masuk ke WWT dari Ismat.',
    direction: 'IN',
    material: 'teflon tape',
    qty: 3,
    unit: 'roll',
    container: 'WWT',
    user: 'Ismat',
    code: 'C-001'
  },
  {
    group: 'IN',
    raw: 'Stock-in 10 pcs obeng plus container WWT, supplier Dimas.',
    direction: 'IN',
    material: 'obeng plus',
    qty: 10,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Fariz mengirim 1 pcs adjustable wrench ke container RWT.',
    direction: 'IN',
    material: 'adjustable wrench',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Barang diterima: safety goggles 2 pcs, masuk WWT dari Ismat.',
    direction: 'IN',
    material: 'safety goggles',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'In cable tie ukuran besar 20 pcs ke container RWT dari Dimas.',
    direction: 'IN',
    material: 'cable tie ukuran besar',
    qty: 20,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Masukkan 3 pcs kunci L 6 mm ke WWT, dari Fariz.',
    direction: 'IN',
    material: 'kunci L 6 mm',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Ismat setor 5 pasang sarung tangan kerja ke container RWT.',
    direction: 'IN',
    material: 'sarung tangan kerja',
    qty: 5,
    unit: 'pasang',
    container: 'RWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Masuk cutter 2 pcs ke container WWT, diserahkan oleh Dimas.',
    direction: 'IN',
    material: 'cutter',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Diterima 4 roll isolasi listrik untuk RWT dari Fariz.',
    direction: 'IN',
    material: 'isolasi listrik',
    qty: 4,
    unit: 'roll',
    container: 'RWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Stock masuk baut stainless M10 sebanyak 12 pcs ke WWT dari Ismat.',
    direction: 'IN',
    material: 'baut stainless M10',
    qty: 12,
    unit: 'pcs',
    container: 'WWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Dimas membawa 2 pcs tang potong, simpan di container RWT.',
    direction: 'IN',
    material: 'tang potong',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Barang masuk: kunci pas 14 mm 3 pcs ke WWT dari Fariz.',
    direction: 'IN',
    material: 'kunci pas 14 mm',
    qty: 3,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'In 1 pcs pressure gauge container RWT dari Ismat.',
    direction: 'IN',
    material: 'pressure gauge',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Masuk masker 20 pcs ke container WWT, pemberi Dimas.',
    direction: 'IN',
    material: 'masker',
    qty: 20,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Fariz menyerahkan 2 pcs sealant gun untuk disimpan di RWT.',
    direction: 'IN',
    material: 'sealant gun',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Diterima 6 roll lakban, masuk ke WWT dari Ismat.',
    direction: 'IN',
    material: 'lakban',
    qty: 6,
    unit: 'roll',
    container: 'WWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Masuk 1 pcs palu 1 kg container RWT dari Dimas. (1/2)',
    direction: 'IN',
    material: 'palu 1 kg',
    qty: 1,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Stock-in elektroda las 3 kg ke WWT, dari Fariz.',
    direction: 'IN',
    material: 'elektroda las',
    qty: 3,
    unit: 'kg',
    container: 'WWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Ismat kirim 2 pcs kunci pipa 10" ke container RWT.',
    direction: 'IN',
    material: 'kunci pipa 10"',
    qty: 2,
    unit: 'pcs',
    container: 'RWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Barang diterima 4 pcs sikat kawat untuk WWT dari Dimas.',
    direction: 'IN',
    material: 'sikat kawat',
    qty: 4,
    unit: 'pcs',
    container: 'WWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'In test pen 3 pcs ke RWT, diserahkan Fariz.',
    direction: 'IN',
    material: 'test pen',
    qty: 3,
    unit: 'pcs',
    container: 'RWT',
    user: 'Fariz'
  },
  {
    group: 'IN',
    raw: 'Masuk lap majun 10 pcs container WWT dari Ismat.',
    direction: 'IN',
    material: 'lap majun',
    qty: 10,
    unit: 'pcs',
    container: 'WWT',
    user: 'Ismat'
  },
  {
    group: 'IN',
    raw: 'Dimas menyerahkan connector 1/2" sebanyak 5 pcs untuk container RWT.',
    direction: 'IN',
    material: 'connector 1/2"',
    qty: 5,
    unit: 'pcs',
    container: 'RWT',
    user: 'Dimas'
  },
  {
    group: 'IN',
    raw: 'Terima 2 pcs seal tape untuk WWT dari Fariz.',
    direction: 'IN',
    material: 'seal tape',
    qty: 2,
    unit: 'pcs',
    container: 'WWT',
    user: 'Fariz',
    code: 'C-001',
    mismatch: true
  }
];

const categoryByMaterial = {
  'kunci pipa 8"': 'T',
  'seal tape': 'C',
  'connector 3/4"': 'F',
  'tang kombinasi': 'T',
  'kunci inggris 12 inch': 'T',
  'hose clamp 3/4"': 'F',
  'gasket pump': 'F',
  'lampu portable': 'M',
  'teflon tape': 'C',
  'obeng plus': 'T',
  'adjustable wrench': 'T',
  'safety goggles': 'C',
  'cable tie ukuran besar': 'C',
  'kunci L 6 mm': 'T',
  'sarung tangan kerja': 'C',
  'cutter': 'T',
  'isolasi listrik': 'C',
  'baut stainless M10': 'M',
  'tang potong': 'T',
  'kunci pas 14 mm': 'T',
  'pressure gauge': 'M',
  'masker': 'C',
  'sealant gun': 'T',
  'lakban': 'C',
  'palu 1 kg': 'T',
  'elektroda las': 'C',
  'kunci pipa 10"': 'T',
  'sikat kawat': 'T',
  'test pen': 'T',
  'lap majun': 'C',
  'connector 1/2 inch': 'F',
  'sapu': 'T',
  'sikat': 'T',
  'kunci inggris 12"': 'T',
  'connector 1/2"': 'F'
};

function csvCell(value) {
  return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
}

const rows = [];
cases.forEach(function (test, index) {
  const result = parse(test.raw);
  const missing = test.missing || [];
  const code = test.code || '';
  const mismatch = !!test.mismatch;
  const label = index + 1 + '. ' + test.group + ': ' + test.raw;

  assert.strictEqual(result.direction, test.direction, label + ' direction');
  assert.strictEqual(result.materialText, test.material, label + ' material');
  assert.strictEqual(result.qty, test.qty, label + ' qty');
  assert.strictEqual(result.unit, test.unit, label + ' unit');
  assert.strictEqual(result.container, test.container, label + ' container');
  assert.strictEqual(result.user, test.user, label + ' user');
  assert.strictEqual(result.note, test.note || '', label + ' note');
  assert.deepStrictEqual(Array.from(result.missingFields), missing, label + ' missing');
  assert.strictEqual(result.materialCode, code, label + ' code');
  assert.strictEqual(result.unitMismatch, mismatch, label + ' unit mismatch');
  assert.strictEqual(result.needsMaster, !code, label + ' needs master');
  assert.strictEqual(result.autoCreate, !code, label + ' auto create');
  assert.strictEqual(
    result.category,
    categoryByMaterial[test.material],
    label + ' category'
  );
  assert.strictEqual(
    result.needsClarification,
    missing.length > 0 || mismatch,
    label + ' needs clarification'
  );

  rows.push([
    index + 1,
    test.group,
    test.raw,
    result.direction,
    result.materialText,
    result.materialCode,
    result.category,
    result.qty,
    result.unit,
    result.container,
    result.user,
    result.note,
    result.missingFields.join(';'),
    result.unitMismatch ? 'YA' : '',
    result.autoCreate ? 'YA' : '',
    result.needsMaster ? 'YA' : '',
    result.needsClarification ? 'YA' : ''
  ]);
});

const ambiguousContainer = parse('out seal tape 1 pcs dari WWT ke RWT');
assert.strictEqual(ambiguousContainer.container, '');
assert.deepStrictEqual(Array.from(ambiguousContainer.candidates.container), ['WWT', 'RWT']);
assert(Array.from(ambiguousContainer.missingFields).includes('container'));

const missingQty = parse('out seal tape oleh Fariz dari WWT');
assert.strictEqual(missingQty.qty, null);
assert.strictEqual(missingQty.unit, '');
assert.deepStrictEqual(Array.from(missingQty.missingFields), ['qty', 'unit']);

const missingDirection = parse('seal tape 1 pcs');
assert.strictEqual(missingDirection.direction, '');
assert.deepStrictEqual(
  Array.from(missingDirection.missingFields),
  ['direction', 'container', 'user']
);

const header = [
  'No',
  'Group',
  'Raw',
  'Direction',
  'Material',
  'Material Code',
  'Category',
  'Qty',
  'Unit',
  'Container',
  'User',
  'Note',
  'Missing',
  'Unit Mismatch',
  'Auto Create',
  'Unknown Material',
  'Needs Clarification'
];

const csv = [header].concat(rows).map(function (row) {
  return row.map(csvCell).join(',');
}).join('\n') + '\n';

fs.writeFileSync(outputPath, csv);
console.log('parser tests passed: ' + cases.length + ' messages');
console.log('parser edge cases passed');
console.log('results written: test/parser-corpus.csv');
