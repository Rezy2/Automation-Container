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
