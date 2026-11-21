/**
 * Reimbursement Request → Expense Claim Form (ECF) pipeline.
 *
 * Fills the ORIGINAL ECF .xlsx by editing the XML inside it, so the
 * checkboxes, dropdowns, formatting and formulas are left untouched.
 * One form response = one ECF.
 */

// Settings (CONFIG) live in Config.gs; see Config.example.gs.

// Response sheet headers for the claimant fields.
const HEADERS = {
  status: 'Reimbursment Status',
  name: 'Full Name',
  email: 'Email Address',
  studentNumber: 'Student Number',
  signature: 'Signature',              // Google Drive link to the uploaded signature image
  firstCategory: 'Expense Category',   // start of purchase block 1
};

// Each purchase is a block of 13 columns; offsets from the block's first column.
// Headers inside the blocks are inconsistent, so blocks are read by position.
const BLOCK_WIDTH = 13;
const BLOCK = {
  category: 0,
  event: 1,
  date: 2,
  vendor: 3,
  amount: 4,
  rate: 5,
  missingReceipt: 6,
  description: 7,
  morePurchases: 12,
};
const MAX_PURCHASES = 10;

// ECF cell positions (Claim Form sheet, 2026-27 template).
const ECF = {
  SHEET_PATH: 'xl/worksheets/sheet1.xml',
  name: 'E4',
  date: 'K4',
  studentNumber: 'E5',
  team: 'K5',
  email: 'E7',
  worktag: 'E8',
  claimantPrintName: 'E42',
  claimantSignDate: 'K42',
  approverPrintName: 'E45',
  approverSignDate: 'K45',
  firstRow: 14,               // purchase 1 → row 14 ... purchase 10 → row 23 (rows 24-28 stay blank)
  // Column H (Project ID) isn't on the form, so it's left blank.
  cols: { date: 'B', vendor: 'D', amount: 'E', rate: 'F', cad: 'G', category: 'I', event: 'J', description: 'L' },
  missingReceiptCol: 10,      // 0-based column K, where the checkboxes sit
  total: 'E29',

  // Signature boxes: claimant G42:I42, team's designated approver G45:I45.
  // Positions are 0-based; sizes in EMU (914400 per inch).
  DRAWING_PATH: 'xl/drawings/drawing1.xml',
  DRAWING_RELS_PATH: 'xl/drawings/_rels/drawing1.xml.rels',
  signatures: {
    claimant: { shapeId: 2000, col: 6, colOff: 190500, row: 41, rowOff: 25400, maxWidth: 2286000, maxHeight: 431800 },
    approver: { shapeId: 2001, col: 6, colOff: 190500, row: 44, rowOff: 25400, maxWidth: 2286000, maxHeight: 457200 },
  },
};

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ECF')
    .addItem('Generate ECF for selected row(s)', 'generateForSelection')
    .addItem(`Generate ECFs for all "${CONFIG.STATUS_TO_PROCESS}" rows`, 'generateForPending')
    .addToUi();
}

function generateForSelection() {
  const sheet = getResponsesSheet_();
  const ranges = sheet.getActiveRangeList();
  if (!ranges || SpreadsheetApp.getActiveSheet().getName() !== sheet.getName()) {
    SpreadsheetApp.getUi().alert(`Select one or more rows on the "${CONFIG.RESPONSES_SHEET}" tab first.`);
    return;
  }
  const rows = new Set();
  ranges.getRanges().forEach(r => {
    for (let i = r.getRow(); i <= r.getLastRow(); i++) if (i > 1) rows.add(i);
  });
  run_([...rows].sort((a, b) => a - b), true);
}

function generateForPending() {
  run_(null, true);
}

/** Use this one for a time-driven trigger (no UI available there). */
function generateForPendingScheduled() {
  run_(null, false);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
function run_(rowNumbers, showUi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = getResponsesSheet_();
  const values = sheet.getDataRange().getValues();
  const col = findColumns_(values[0]);

  if (!rowNumbers) {
    rowNumbers = [];
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][col.status]).trim() === CONFIG.STATUS_TO_PROCESS) rowNumbers.push(r + 1);
    }
  }

  const template = DriveApp.getFileById(CONFIG.TEMPLATE_FILE_ID).getBlob();
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const log = getLogSheet_(ss);
  const today = new Date();
  const summary = [];

  let approverSignature = null;
  if (CONFIG.APPROVER_SIGNATURE_FILE_ID) {
    const errors = [];
    approverSignature = loadSignature_(CONFIG.APPROVER_SIGNATURE_FILE_ID, errors);
    if (!approverSignature) throw new Error(`Approver ${errors[0]} Check APPROVER_SIGNATURE_FILE_ID in CONFIG.`);
  }

  rowNumbers.forEach(rowNumber => {
    const row = values[rowNumber - 1];
    if (!row || row.every(v => v === '' || v === null)) return;

    const claim = parseClaim_(row, col, tz);
    claim.signature = loadSignature_(row[col.signature], claim.errors);
    claim.approverSignature = approverSignature;
    if (claim.errors.length) {
      setStatus_(sheet, rowNumber, col.status, CONFIG.STATUS_FAILED);
      log.appendRow([today, rowNumber, claim.name, 'Needs Info', claim.errors.join('\n')]);
      summary.push(`Row ${rowNumber} (${claim.name || 'no name'}): ${claim.errors.length} problem(s), see ECF Log`);
      return;
    }

    claim.ecfDate = Utilities.formatDate(today, tz, 'yyyy-MM-dd');   // text: the ECF's DATE cells aren't date-formatted
    // Time included so a regenerated ECF doesn't share a name with the earlier one.
    const fileName = `ECF - ${claim.name} - ${Utilities.formatDate(today, tz, 'yyyy-MM-dd HHmm')}.xlsx`;
    const file = folder.createFile(buildEcf_(template, claim, fileName));

    setStatus_(sheet, rowNumber, col.status, CONFIG.STATUS_DONE);
    log.appendRow([today, rowNumber, claim.name, `ECF Generated (${claim.purchases.length} purchase(s))`, file.getUrl()]);
    summary.push(`Row ${rowNumber} (${claim.name}): ${claim.purchases.length} purchase(s) → ${fileName}`);
  });

  if (showUi) {
    SpreadsheetApp.getUi().alert(summary.length ? summary.join('\n') : 'Nothing to process.');
  }
}

// ---------------------------------------------------------------------------
// Reading + checking a response
// ---------------------------------------------------------------------------
function findColumns_(headerRow) {
  const clean = h => String(h).replace(/\s+/g, ' ').trim();
  const headers = headerRow.map(clean);
  const col = {};
  Object.keys(HEADERS).forEach(key => {
    const i = headers.indexOf(HEADERS[key]);
    if (i === -1) throw new Error(`Column "${HEADERS[key]}" not found on "${CONFIG.RESPONSES_SHEET}".`);
    col[key] = i;
  });
  return col;
}

function parseClaim_(row, col, tz) {
  const errors = [];
  const claim = {
    name: String(row[col.name]).trim(),
    email: String(row[col.email]).trim(),
    studentNumber: String(row[col.studentNumber]).trim().replace(/\.0+$/, ''),
    purchases: [],
    errors,
  };

  if (!claim.name) errors.push('Full Name is empty.');
  if (!/^\d+$/.test(claim.studentNumber)) errors.push(`Student Number "${row[col.studentNumber]}" is not a number.`);

  for (let n = 1; n <= MAX_PURCHASES; n++) {
    const start = col.firstCategory + BLOCK_WIDTH * (n - 1);
    const cell = key => row[start + BLOCK[key]];
    const text = key => String(cell(key) === null || cell(key) === undefined ? '' : cell(key)).trim();

    if (!text('category') && !text('vendor') && !text('amount')) break;   // empty block

    const p = `Purchase ${n}`;
    const purchase = {
      category: CONFIG.CATEGORY_ALIASES[text('category')] || text('category'),
      event: text('event'),
      vendor: text('vendor'),
      description: text('description'),
      missingReceipt: /^yes$/i.test(text('missingReceipt')),
    };

    purchase.amount = toNumber_(cell('amount'));
    if (purchase.amount === null || purchase.amount <= 0) errors.push(`${p}: Receipt Amount "${text('amount')}" needs exactly one number above 0.`);

    purchase.rate = toNumber_(cell('rate'));
    if (purchase.rate === null || purchase.rate <= 0) errors.push(`${p}: Exchange Rate "${text('rate')}" needs exactly one number above 0.`);

    purchase.dateSerial = toDateSerial_(cell('date'), tz);
    if (purchase.dateSerial === null) errors.push(`${p}: Purchase Date "${text('date')}" is not a date.`);

    if (CONFIG.VALID_CATEGORIES.indexOf(purchase.category) === -1) {
      errors.push(`${p}: Expense Category "${purchase.category}" is not one of: ${CONFIG.VALID_CATEGORIES.join(', ')}.`);
    }
    if (!purchase.vendor) errors.push(`${p}: Vendor is empty.`);

    claim.purchases.push(purchase);
    if (!/^yes$/i.test(text('morePurchases'))) break;
  }

  if (!claim.purchases.length) errors.push('No purchases found in this response.');
  return claim;
}

/** Loads the signature image from its Drive link; returns null and records an error on failure. */
function loadSignature_(link, errors) {
  const id = (String(link || '').match(/[-\w]{25,}/) || [])[0];
  if (!id) {
    errors.push('Signature: no Google Drive link found.');
    return null;
  }

  let blob;
  try {
    blob = DriveApp.getFileById(id).getBlob();
  } catch (e) {
    errors.push(`Signature: can't open the Drive file (${e.message}).`);
    return null;
  }

  let type = blob.getContentType();
  if (type !== 'image/png' && type !== 'image/jpeg') {
    try {
      blob = blob.getAs('image/png');   // e.g. GIF/BMP; iPhone HEIC photos can't be converted
      type = 'image/png';
    } catch (e) {
      errors.push(`Signature: "${type}" files aren't supported. Upload a PNG or JPG.`);
      return null;
    }
  }

  const bytes = blob.getBytes();
  const size = getImageSize_(bytes);
  if (!size) {
    errors.push('Signature: the image file could not be read.');
    return null;
  }
  return { bytes, ext: type === 'image/png' ? 'png' : 'jpeg', width: size.width, height: size.height };
}

/** Reads pixel width/height from PNG or JPEG bytes. */
function getImageSize_(bytes) {
  const b = i => bytes[i] & 0xff;
  if (b(0) === 0x89 && b(1) === 0x50) {   // PNG: IHDR
    return {
      width: (b(16) << 24 | b(17) << 16 | b(18) << 8 | b(19)) >>> 0,
      height: (b(20) << 24 | b(21) << 16 | b(22) << 8 | b(23)) >>> 0,
    };
  }
  if (b(0) === 0xff && b(1) === 0xd8) {   // JPEG: find the start-of-frame marker
    let i = 2;
    while (i + 9 < bytes.length) {
      if (b(i) !== 0xff) { i++; continue; }
      const marker = b(i + 1);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: b(i + 5) << 8 | b(i + 6), width: b(i + 7) << 8 | b(i + 8) };
      }
      i += 2 + (b(i + 2) << 8 | b(i + 3));
    }
  }
  return null;
}

/**
 * Pulls the number out of text like "100 CAD", "CAD 100", "$1,234.50" or "C$ 12.5".
 * Returns null if there's no number, or more than one (e.g. "1 USD = 1.38 CAD"),
 * since then it isn't clear which one is meant.
 */
function toNumber_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;
  const s = String(value === null || value === undefined ? '' : value);
  const numbers = s.match(/\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+/g) || [];
  if (numbers.length !== 1) return null;
  return Number(numbers[0].replace(/,/g, ''));
}

/** Converts a sheet date to an Excel date serial (days since 1899-12-30). */
function toDateSerial_(value, tz) {
  let y, m, d;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    [y, m, d] = Utilities.formatDate(value, tz, 'yyyy-M-d').split('-').map(Number);
  } else {
    const match = String(value === null || value === undefined ? '' : value).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) return null;
    [m, d, y] = [Number(match[1]), Number(match[2]), Number(match[3])];
  }
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

// ---------------------------------------------------------------------------
// Writing the ECF .xlsx
// ---------------------------------------------------------------------------
function buildEcf_(templateBlob, claim, fileName) {
  const zipBlob = templateBlob.copyBlob().setContentType('application/zip');
  const parts = Utilities.unzip(zipBlob);

  const text = {};
  parts.forEach(b => {
    if (/\.(xml|vml|rels)$/.test(b.getName())) text[b.getName()] = b.getDataAsString('UTF-8');
  });

  const changed = fillEcfXml_(text, claim);

  const out = parts.map(b =>
    changed[b.getName()] !== undefined
      ? Utilities.newBlob(changed[b.getName()], 'application/xml', b.getName())
      : b
  );
  signaturesOf_(claim).forEach(([slot, sig]) => {
    out.push(Utilities.newBlob(sig.bytes, 'image/' + sig.ext, signatureMediaPath_(slot, sig)));
  });
  return Utilities.zip(out, fileName)
    .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * Pure XML step: takes {path: xmlText} for the template and returns
 * {path: newXmlText} for every part that changed.
 */
function fillEcfXml_(files, claim) {
  let sheet = files[ECF.SHEET_PATH];
  const out = {};

  sheet = setCell_(sheet, ECF.name, claim.name);
  const ecfDate = claim.ecfDate;   // already text, e.g. '2026-11-21'
  sheet = setCell_(sheet, ECF.date, ecfDate);
  sheet = setCell_(sheet, ECF.studentNumber, Number(claim.studentNumber));
  sheet = setCell_(sheet, ECF.email, claim.email);
  if (CONFIG.TEAM_NAME) sheet = setCell_(sheet, ECF.team, CONFIG.TEAM_NAME);
  if (CONFIG.WORKTAG) sheet = setCell_(sheet, ECF.worktag, CONFIG.WORKTAG);
  sheet = setCell_(sheet, ECF.claimantPrintName, claim.name);
  if (CONFIG.APPROVER_NAME) sheet = setCell_(sheet, ECF.approverPrintName, CONFIG.APPROVER_NAME);

  // Date each signature with the day the ECF is generated.
  if (claim.signature) sheet = setCell_(sheet, ECF.claimantSignDate, ecfDate);
  if (claim.approverSignature) sheet = setCell_(sheet, ECF.approverSignDate, ecfDate);

  let total = 0;
  const missingRows = [];
  claim.purchases.forEach((p, i) => {
    const r = ECF.firstRow + i;
    const c = ECF.cols;
    sheet = setCell_(sheet, c.date + r, p.dateSerial);
    sheet = setCell_(sheet, c.vendor + r, p.vendor);
    sheet = setCell_(sheet, c.amount + r, p.amount);
    sheet = setCell_(sheet, c.rate + r, p.rate);
    sheet = setCell_(sheet, c.category + r, p.category);
    sheet = setCell_(sheet, c.event + r, p.event);
    sheet = setCell_(sheet, c.description + r, p.description);

    const cad = Math.round(p.amount * p.rate * 100) / 100;
    sheet = setCachedFormulaValue_(sheet, c.cad + r, cad);
    total += cad;
    if (p.missingReceipt) missingRows.push(r);
  });
  sheet = setCachedFormulaValue_(sheet, ECF.total, Math.round(total * 100) / 100);

  // Open at the top of the form.
  sheet = sheet.replace(/topLeftCell="[A-Z]+\d+"/, 'topLeftCell="A1"');
  out[ECF.SHEET_PATH] = sheet;

  if (missingRows.length) Object.assign(out, tickCheckboxes_(files, sheet, missingRows));
  signaturesOf_(claim).forEach(([slot, sig]) => {
    // Each image builds on the drawing/rels the previous one changed.
    Object.assign(out, addSignatureImage_(Object.assign({}, files, out), sig, slot));
  });

  // Make Excel recalculate the CAD amounts and TOTAL when the file is opened.
  const workbook = files['xl/workbook.xml'];
  out['xl/workbook.xml'] = /fullCalcOnLoad=/.test(workbook)
    ? workbook
    : workbook.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');

  return out;
}

/** Replaces an existing cell in the sheet XML, keeping its style. */
function setCell_(sheetXml, ref, value) {
  const re = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`);
  const match = sheetXml.match(re);
  if (!match) throw new Error(`ECF template has no cell ${ref}.`);
  const style = (match[1].match(/\ss="\d+"/) || [''])[0];

  let cell;
  if (typeof value === 'number') {
    cell = `<c r="${ref}"${style}><v>${value}</v></c>`;
  } else {
    cell = `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXml_(value)}</t></is></c>`;
  }
  return sheetXml.replace(re, cell);
}

/** Updates the stored result of a formula cell so previews show the right number. */
function setCachedFormulaValue_(sheetXml, ref, value) {
  const re = new RegExp(`(<c r="${ref}"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>))<v>[^<]*</v>`);
  return sheetXml.replace(re, `$1<v>${value}</v>`);
}

/** Ticks the Missing Receipt checkbox on each given sheet row. */
function tickCheckboxes_(files, sheetXml, rows) {
  const out = {};
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const rels = files[relsPath];

  // Map control → its ctrlProp file and the cell it sits in.
  const controls = [];
  const controlRe = /<control shapeId="(\d+)" r:id="(rId\d+)"[^>]*>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/g;
  let m;
  while ((m = controlRe.exec(sheetXml))) {
    const target = (rels.match(new RegExp(`Id="${m[2]}"[^>]*Target="([^"]+)"`)) ||
                    rels.match(new RegExp(`Target="([^"]+)"[^>]*Id="${m[2]}"`)) || [])[1];
    controls.push({ shapeId: m[1], col: Number(m[3]), row: Number(m[4]) + 1, target });
  }

  const vmlPath = Object.keys(files).find(p => /^xl\/drawings\/vmlDrawing\d*\.vml$/.test(p));
  let vml = files[vmlPath];

  rows.forEach(row => {
    const ctl = controls.find(c => c.col === ECF.missingReceiptCol && c.row === row);
    if (!ctl || !ctl.target) throw new Error(`No Missing Receipt checkbox found on ECF row ${row}.`);

    const propPath = 'xl/' + ctl.target.replace(/^\.\.\//, '');
    const prop = files[propPath];
    if (!/checked=/.test(prop)) out[propPath] = prop.replace('<formControlPr ', '<formControlPr checked="Checked" ');

    const shapeRe = new RegExp(`(<v:shape [^>]*id="_x0000_s${ctl.shapeId}"[\\s\\S]*?<x:ClientData ObjectType="Checkbox">)([\\s\\S]*?)(</x:ClientData>)`);
    vml = vml.replace(shapeRe, (all, open, body, close) =>
      /<x:Checked>/.test(body) ? all : `${open}${body}<x:Checked>1</x:Checked>${close}`);
  });

  out[vmlPath] = vml;
  return out;
}

/** [slot, image] pairs for the signatures this ECF gets. */
function signaturesOf_(claim) {
  return [['claimant', claim.signature], ['approver', claim.approverSignature]].filter(([, sig]) => sig);
}

function signatureMediaPath_(slot, sig) {
  return `xl/media/${slot}-signature.${sig.ext}`;
}

/** Places a signature image in its box ('claimant' or 'approver'), scaled to fit. */
function addSignatureImage_(files, sig, slot) {
  const out = {};
  const box = ECF.signatures[slot];
  const scale = Math.min(box.maxWidth / sig.width, box.maxHeight / sig.height);
  const cx = Math.round(sig.width * scale);
  const cy = Math.round(sig.height * scale);

  // Link the image file to the drawing.
  const rels = files[ECF.DRAWING_RELS_PATH];
  const relId = `rId_${slot}Signature`;
  out[ECF.DRAWING_RELS_PATH] = rels.replace('</Relationships>',
    `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
    `Target="../media/${signatureMediaPath_(slot, sig).split('/').pop()}"/></Relationships>`);

  // Add the picture to the drawing.
  const anchor =
    '<xdr:oneCellAnchor>' +
      `<xdr:from><xdr:col>${box.col}</xdr:col><xdr:colOff>${box.colOff}</xdr:colOff>` +
      `<xdr:row>${box.row}</xdr:row><xdr:rowOff>${box.rowOff}</xdr:rowOff></xdr:from>` +
      `<xdr:ext cx="${cx}" cy="${cy}"/>` +
      '<xdr:pic>' +
        `<xdr:nvPicPr><xdr:cNvPr id="${box.shapeId}" name="${slot} signature"/>` +
        '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
        `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${relId}"/>` +
        '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
        `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
      '</xdr:pic>' +
      '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>';
  out[ECF.DRAWING_PATH] = files[ECF.DRAWING_PATH].replace('</xdr:wsDr>', anchor + '</xdr:wsDr>');

  // Register the image type (the template only declares jpeg).
  const typesPath = '[Content_Types].xml';
  const types = files[typesPath];
  if (!new RegExp(`<Default Extension="${sig.ext}"`).test(types)) {
    out[typesPath] = types.replace('<Default ', `<Default Extension="${sig.ext}" ContentType="image/${sig.ext}"/><Default `);
  }
  return out;
}

function escapeXml_(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------
function getResponsesSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.RESPONSES_SHEET);
  if (!sheet) throw new Error(`Sheet "${CONFIG.RESPONSES_SHEET}" not found.`);
  return sheet;
}

function getLogSheet_(ss) {
  let log = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET);
    log.appendRow(['Run at', 'Response row', 'Name', 'Result', 'ECF link / problems']);
    log.setFrozenRows(1);
  }
  return log;
}

/** The status column is a dropdown; if the value isn't one of its options, log instead of crashing. */
function setStatus_(sheet, rowNumber, statusCol, status) {
  try {
    sheet.getRange(rowNumber, statusCol + 1).setValue(status);
  } catch (e) {
    console.warn(`Could not set status "${status}" on row ${rowNumber}: ${e.message}`);
  }
}
