/**
 * ============================================================================
 * Karjat Protection Settings Portal — Google Drive/Sheets backend
 * ============================================================================
 * This Apps Script backs TWO things for the portal:
 *
 * 1) "Relay/Bay/Substation Overview" — one tab per Overview section, styled
 *    like a Quick Reference sheet (Section header rows, then Parameter /
 *    Value / Unit / Notes rows), with the Parameter column restricted to a
 *    dropdown of real setting titles. (actions: get / ensure / save)
 *
 * 2) The full settings database — every bay, relay, and setting, stored as
 *    proper rows in dedicated sheets (_Bays / _Relays / _Settings / _Meta)
 *    inside a "Karjat Protection Portal Data" Drive folder, plus the SLD
 *    reference image as a real Drive file. The portal fetches this at
 *    startup instead of embedding ~4MB of JSON in the HTML file, so the
 *    HTML stays small. (actions: bootstrap / getFullData / getSld)
 *
 * The portal calls this script's Web App URL directly from the browser
 * (no login prompt for the person using the portal).
 *
 * ----------------------------------------------------------------------
 * DEPLOY / UPDATE THIS AS A WEB APP
 * ----------------------------------------------------------------------
 * 1. Go to https://script.google.com, open your existing project (or
 *    create a new one), and replace Code.gs with this whole file.
 * 2. Click Deploy → Manage deployments → Edit (pencil) → New version →
 *    Deploy. This keeps your existing /exec URL working — no need to
 *    change anything in the portal.
 *    (First time only: Deploy → New deployment → Web app → Execute as
 *    "Me" → Who has access "Anyone" → Deploy → authorise → copy the URL.)
 * 3. In the portal, open Admin (as Admin) → "Data source" card → click
 *    "Sync full data to Google Drive" once. That reads the data already
 *    loaded in your browser and writes it into the sheets described
 *    above. After that succeeds, every future page load fetches from
 *    Drive instead of the built-in copy.
 * ============================================================================
 */

// PERFORMANCE: paste your spreadsheet ID here after the first sync. Without it, every
// single request has to search Drive by folder/file name first, which is slow — this is
// the single biggest speed win available. After running "Sync full data to Google Drive"
// in the portal once, the Admin page shows the spreadsheet ID to copy in here; then
// Deploy → Manage deployments → Edit → New version to apply it.
const SS_ID = '';
const ROOT_FOLDER_NAME = 'Karjat Protection Portal Data';
const IMAGES_FOLDER_NAME = 'Images';
const SS_NAME = 'Karjat Relay Overviews';

const LOOKUP_SHEET = '_lookup'; // reserved name, kept out of entity tab listings

const BAYS_SHEET = '_Bays';
const RELAYS_SHEET = '_Relays';
const SETTINGS_SHEET = '_Settings';
const META_SHEET = '_Meta';
const BAYS_HEADER = ['id','num','name','voltage','scheme','dia','pos','isBusZone'];
const RELAYS_HEADER = ['id','bayId','name','file','family','model','kind','n_settings','vitalsJson'];
const SETTINGS_HEADER = ['relayId','group','t','v','u','s','r','desc','vm'];

/* ---------------------------------------------------------------------- */
/* Drive folder helpers                                                   */
/* ---------------------------------------------------------------------- */
function getRootFolder_() {
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}
function getImagesFolder_() {
  const root = getRootFolder_();
  const it = root.getFoldersByName(IMAGES_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return root.createFolder(IMAGES_FOLDER_NAME);
}

/* ---------------------------------------------------------------------- */
/* Spreadsheet / tab helpers                                              */
/* ---------------------------------------------------------------------- */

function getSS_() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  const root = getRootFolder_();
  const files = root.getFilesByName(SS_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  // create loose then move into the folder (Sheets API creates in "My Drive" root by default)
  const ss = SpreadsheetApp.create(SS_NAME);
  const file = DriveApp.getFileById(ss.getId());
  root.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  Logger.log('Created new spreadsheet, id=' + ss.getId() +
    ' — paste this into SS_ID for stability.');
  return ss;
}

// Sheet tab names can't contain [ ] * / \ ? : — so we sanitize the entityId, then
// join it to the blockId with a delimiter ("~~") that also survives sanitization.
// Both tabName_ and entityPrefix_ sanitize the entityId the SAME way, so a tab always
// starts with its entity's prefix and listBlocks_ can find it reliably.
const BLOCK_DELIM = '~~';
function sanitizeId_(id) {
  return String(id).replace(/[\[\]\*\/\\\?:]/g, '_');
}
function tabName_(entityId, blockId) {
  const raw = blockId ? (sanitizeId_(entityId) + BLOCK_DELIM + sanitizeId_(blockId)) : sanitizeId_(entityId);
  return raw.slice(0, 95);
}
function entityPrefix_(entityId) {
  return sanitizeId_(entityId) + BLOCK_DELIM;
}
const SYSTEM_SHEETS = [BAYS_SHEET, RELAYS_SHEET, SETTINGS_SHEET, META_SHEET, LOOKUP_SHEET];

// Creates a blank tab (no imposed structure at all) the first time it's opened.
// If it already exists, returns it untouched — whatever is in it stays exactly as is.
function ensureTab_(ss, entityId, blockId, title) {
  const name = tabName_(entityId, blockId);
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  if (title) sh.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  return sh;
}

// Lists every tab that belongs to this entity (relay/bay/substation), so the portal
// can build a tab-strip. blockId is whatever follows "entityId::" in the tab name.
function listBlocks_(ss, entityId) {
  const prefix = entityPrefix_(entityId);
  return ss.getSheets()
    .filter(function (sh) { return sh.getName().indexOf(prefix) === 0; })
    .map(function (sh) { return { blockId: sh.getName().slice(prefix.length), gid: sh.getSheetId() }; });
}

/* ---------------------------------------------------------------------- */
/* Read / write RAW grid content — no imposed columns or rows at all.     */
/* Whatever is typed or pasted into the sheet is exactly what the portal  */
/* mirrors back, including each column's width so it looks the same.     */
/* ---------------------------------------------------------------------- */

function readTabRaw_(sh) {
  const range = sh.getDataRange();
  const numRows = range.getNumRows(), numCols = range.getNumColumns();
  const rows = (numRows > 0 && numCols > 0) ? range.getValues() : [];
  const colWidths = [];
  for (let c = 1; c <= Math.max(numCols, 1); c++) colWidths.push(sh.getColumnWidth(c));
  return { rows: rows, colWidths: colWidths };
}

function writeTabRaw_(sh, rows) {
  sh.clear();
  if (!rows || !rows.length) return;
  const numCols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const padded = rows.map(function (r) {
    const rr = r.slice();
    while (rr.length < numCols) rr.push('');
    return rr;
  });
  sh.getRange(1, 1, padded.length, numCols).setValues(padded);
}

/* ---------------------------------------------------------------------- */
/* Web app entry points                                                   */
/* ---------------------------------------------------------------------- */

/* ---------------------------------------------------------------------- */
/* Full settings database (bootstrap / getFullData / getSld)              */
/* ---------------------------------------------------------------------- */

function writeSheetBulk_(ss, name, header, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#2A4A6B').setFontColor('#FBEFC9');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  return sh;
}

function bootstrapDatabase_(payload) {
  const ss = getSS_();
  const sub = payload.substation || {};

  // _Bays: one row per real bay, plus one per bus-zone (isBusZone=true)
  const bayRows = [];
  (payload.bays || []).forEach(function (b) {
    bayRows.push([b.id, b.num || '', b.name, b.voltage, b.scheme || '', b.dia || '', b.pos || '', false]);
  });
  (payload.busZones || []).forEach(function (z) {
    bayRows.push([z.id, '', z.name, z.voltage, 'busbar', '', '', true]);
  });
  writeSheetBulk_(ss, BAYS_SHEET, BAYS_HEADER, bayRows);

  // _Relays and _Settings: walk every bay's relays (+ bus-zone relays)
  const relayRows = [];
  const settingRows = [];
  function walkRelays(bayId, relays) {
    (relays || []).forEach(function (r) {
      relayRows.push([r.id, bayId, r.name, r.file || '', r.family || '', r.model || '', r.kind || '',
        r.n_settings || 0, JSON.stringify(r.vitals || [])]);
      (r.groups || []).forEach(function (g) {
        (g.settings || []).forEach(function (s) {
          settingRows.push([r.id, g.group, s.t, s.v, s.u || '', s.s || '', s.r || '', s.desc || '', s.vm || '']);
        });
      });
    });
  }
  (payload.bays || []).forEach(function (b) { walkRelays(b.id, b.relays); });
  (payload.busZones || []).forEach(function (z) { walkRelays(z.id, z.relays); });
  writeSheetBulk_(ss, RELAYS_SHEET, RELAYS_HEADER, relayRows);
  writeSheetBulk_(ss, SETTINGS_SHEET, SETTINGS_HEADER, settingRows);

  // SLD reference image -> a real Drive file, not a giant cell
  let sldFileId = '';
  if (payload.sldBase64) {
    const imgFolder = getImagesFolder_();
    const existing = imgFolder.getFilesByName('SLD_reference.jpg');
    if (existing.hasNext()) existing.next().setTrashed(true);
    const bytes = Utilities.base64Decode(payload.sldBase64);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', 'SLD_reference.jpg');
    sldFileId = imgFolder.createFile(blob).getId();
  }

  // _Meta: small key/value sheet
  let metaSh = ss.getSheetByName(META_SHEET);
  if (!metaSh) metaSh = ss.insertSheet(META_SHEET);
  metaSh.clear();
  const metaRows = [
    ['key', 'value'],
    ['name', sub.name || ''],
    ['ssCode', sub.ssCode || ''],
    ['docDate', sub.docDate || ''],
    ['statsBays', (sub.stats && sub.stats.bays) || bayRows.length],
    ['statsRelays', (sub.stats && sub.stats.relays) || relayRows.length],
    ['statsSettings', (sub.stats && sub.stats.settings) || settingRows.length],
    ['sldFileId', sldFileId],
    ['syncedAt', new Date().toISOString()]
  ];
  metaSh.getRange(1, 1, metaRows.length, 2).setValues(metaRows);
  metaSh.getRange(1, 1, 1, 2).setFontWeight('bold');

  return { bays: bayRows.length, relays: relayRows.length, settings: settingRows.length, sldFileId: sldFileId, spreadsheetId: ss.getId() };
}

function readMeta_(ss) {
  const sh = ss.getSheetByName(META_SHEET);
  const meta = {};
  if (!sh) return meta;
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) meta[values[i][0]] = values[i][1];
  return meta;
}

function readFullDatabase_() {
  const ss = getSS_();
  const meta = readMeta_(ss);

  const bSh = ss.getSheetByName(BAYS_SHEET);
  const rSh = ss.getSheetByName(RELAYS_SHEET);
  const sSh = ss.getSheetByName(SETTINGS_SHEET);
  if (!bSh || !rSh || !sSh) return { ok: false, error: 'Database not synced yet — run "Sync full data to Google Drive" from Admin first.' };

  const bayRows = bSh.getLastRow() > 1 ? bSh.getRange(2, 1, bSh.getLastRow() - 1, BAYS_HEADER.length).getValues() : [];
  const relayRows = rSh.getLastRow() > 1 ? rSh.getRange(2, 1, rSh.getLastRow() - 1, RELAYS_HEADER.length).getValues() : [];
  const settingRows = sSh.getLastRow() > 1 ? sSh.getRange(2, 1, sSh.getLastRow() - 1, SETTINGS_HEADER.length).getValues() : [];

  // group settings by relayId -> group -> [settings]
  const settingsByRelay = {};
  settingRows.forEach(function (row) {
    const relayId = row[0], group = row[1];
    if (!settingsByRelay[relayId]) settingsByRelay[relayId] = {};
    if (!settingsByRelay[relayId][group]) settingsByRelay[relayId][group] = [];
    const s = { t: row[2], v: row[3], u: row[4], s: row[5], r: row[6] };
    if (row[7]) s.desc = row[7];
    if (row[8]) s.vm = row[8];
    settingsByRelay[relayId][group].push(s);
  });

  // build relays keyed by id, with groups[] in first-seen order
  const relaysById = {};
  relayRows.forEach(function (row) {
    const id = row[0];
    const groupsMap = settingsByRelay[id] || {};
    const groups = Object.keys(groupsMap).map(function (g) { return { group: g, settings: groupsMap[g] }; });
    let vitals = [];
    try { vitals = JSON.parse(row[8] || '[]'); } catch (e) { vitals = []; }
    relaysById[id] = {
      id: id, name: row[2], file: row[3], family: row[4], model: row[5], kind: row[6],
      n_settings: row[7], vitals: vitals, groups: groups, bayId: row[1]
    };
  });

  // assemble bays / busZones, attaching their relays (indexed by bayId — avoids re-scanning
  // the full relay list once per bay)
  const relaysByBay = {};
  relayRows.forEach(function (row) {
    const bayId = row[1];
    if (!relaysByBay[bayId]) relaysByBay[bayId] = [];
    relaysByBay[bayId].push(relaysById[row[0]]);
  });
  const bays = [], busZones = [];
  bayRows.forEach(function (row) {
    const id = row[0], isBusZone = row[7] === true || row[7] === 'TRUE';
    const relays = relaysByBay[id] || [];
    if (isBusZone) {
      busZones.push({ id: id, name: row[2], voltage: row[3], relays: relays });
    } else {
      bays.push({ id: id, num: row[1], name: row[2], voltage: row[3], scheme: row[4], dia: row[5] || null, pos: row[6] || null, relays: relays });
    }
  });

  return {
    ok: true,
    name: meta.name || '', ssCode: meta.ssCode || '', docDate: meta.docDate || '',
    stats: { bays: Number(meta.statsBays) || bays.length, relays: Number(meta.statsRelays) || relayRows.length, settings: Number(meta.statsSettings) || settingRows.length },
    bays: bays, busZones: busZones, syncedAt: meta.syncedAt || ''
  };
}

function readSld_() {
  const ss = getSS_();
  const meta = readMeta_(ss);
  if (!meta.sldFileId) return { ok: false, error: 'No SLD image synced yet.' };
  const file = DriveApp.getFileById(meta.sldFileId);
  const bytes = file.getBlob().getBytes();
  return { ok: true, base64: Utilities.base64Encode(bytes) };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetUrlFor_(ss, sh) {
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

// GET ?action=get&relayId=...   — read-only fetch, used by the portal's "Refresh" button.
function doGet(e) {
  try {
    const action = e.parameter.action || 'get';

    if (action === 'getFullData') return jsonOut_(readFullDatabase_());
    if (action === 'getSld') return jsonOut_(readSld_());

    const entityId = e.parameter.entityId;
    if (action === 'listBlocks') {
      if (!entityId) return jsonOut_({ ok: false, error: 'entityId required' });
      return jsonOut_({ ok: true, blocks: listBlocks_(getSS_(), entityId) });
    }
    if (action === 'get') {
      const blockId = e.parameter.blockId;
      if (!entityId || !blockId) return jsonOut_({ ok: false, error: 'entityId and blockId required' });
      const ss = getSS_();
      const sh = ss.getSheetByName(tabName_(entityId, blockId));
      if (!sh) return jsonOut_({ ok: true, exists: false, rows: [], colWidths: [] });
      const data = readTabRaw_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }
    return jsonOut_({ ok: false, error: 'unknown action for GET: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// POST body (JSON): {action:'listBlocks'|'ensure'|'get'|'save'|'deleteBlock', entityId, blockId, title, rows}
// or {action:'bootstrap', substation:{...}, bays:[...], busZones:[...], sldBase64}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'bootstrap') {
      const result = bootstrapDatabase_(body);
      return jsonOut_(Object.assign({ ok: true }, result));
    }
    if (action === 'getFullData') return jsonOut_(readFullDatabase_());
    if (action === 'getSld') return jsonOut_(readSld_());

    const entityId = body.entityId;
    if (!entityId) return jsonOut_({ ok: false, error: 'entityId required' });
    const ss = getSS_();

    if (action === 'listBlocks') {
      return jsonOut_({ ok: true, blocks: listBlocks_(ss, entityId) });
    }

    const blockId = body.blockId;
    if (!blockId) return jsonOut_({ ok: false, error: 'blockId required' });

    if (action === 'get') {
      const sh = ss.getSheetByName(tabName_(entityId, blockId));
      if (!sh) return jsonOut_({ ok: true, exists: false, rows: [], colWidths: [] });
      const data = readTabRaw_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }

    // Creates a completely blank tab the first time — no imposed structure.
    // If it already exists, its content is left exactly as the user last had it.
    if (action === 'ensure') {
      const sh = ensureTab_(ss, entityId, blockId, body.title);
      const data = readTabRaw_(sh);
      return jsonOut_(Object.assign({ ok: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }

    // Raw bulk overwrite — used by the "Upload Excel" feature to push parsed
    // spreadsheet rows straight into the tab, no column mapping imposed.
    if (action === 'save') {
      let sh = ss.getSheetByName(tabName_(entityId, blockId));
      if (!sh) sh = ensureTab_(ss, entityId, blockId, body.title);
      writeTabRaw_(sh, body.rows || []);
      return jsonOut_({ ok: true, sheetUrl: sheetUrlFor_(ss, sh) });
    }

    if (action === 'deleteBlock') {
      const sh = ss.getSheetByName(tabName_(entityId, blockId));
      if (sh) ss.deleteSheet(sh);
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: 'unknown action for POST: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
