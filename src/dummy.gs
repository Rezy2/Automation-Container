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
