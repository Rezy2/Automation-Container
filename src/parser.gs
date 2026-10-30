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
