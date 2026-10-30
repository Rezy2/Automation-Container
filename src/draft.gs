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
