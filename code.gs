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
// ---------------------------------------------------------------------
// BACKEND VERSION STAMP
// Bump this whenever code.gs changes. The portal compares it against the
// version it expects and tells you plainly if the deployed script is out of
// date — instead of surfacing a cryptic error like "blockId required" that
// actually just means "you forgot to redeploy".
// REMINDER: saving in the editor is NOT enough. You must do
//   Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy
// ---------------------------------------------------------------------
const BACKEND_VERSION = 7;

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

// ======================================================================
// OVERVIEW = one SEPARATE Google Spreadsheet per entity (relay / bay / substation).
// Each such spreadsheet can hold as many tabs as the user wants, fully editable in
// Google Sheets with real formatting. The portal mirrors each tab faithfully —
// values, cell background colours, font colours/weights/styles/sizes, alignments,
// merged cells and column widths — blended with the site theme for unstyled cells.
// A registry sheet in the master spreadsheet maps entityId -> that entity's file id.
// ======================================================================
const OVERVIEW_REGISTRY = '_OverviewRegistry';
const OVERVIEW_FOLDER_NAME = 'Overviews';

function sanitizeId_(id) {
  return String(id).replace(/[\[\]\*\/\\\?:]/g, '_');
}
function getOverviewFolder_() {
  const root = getRootFolder_();
  const it = root.getFoldersByName(OVERVIEW_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return root.createFolder(OVERVIEW_FOLDER_NAME);
}
function overviewRegistrySheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(OVERVIEW_REGISTRY);
  if (!sh) {
    sh = ss.insertSheet(OVERVIEW_REGISTRY);
    sh.getRange(1, 1, 1, 3).setValues([['entityId', 'spreadsheetId', 'title']]);
    sh.hideSheet();
  }
  return sh;
}
function overviewLookup_(entityId) {
  const sh = overviewRegistrySheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(entityId)) return { spreadsheetId: rows[i][1], title: rows[i][2], row: i + 2 };
  return null;
}
// Find-or-create the entity's OWN spreadsheet (with its own tabs).
function ensureOverviewSpreadsheet_(entityId, title) {
  const existing = overviewLookup_(entityId);
  if (existing && existing.spreadsheetId) {
    try { return SpreadsheetApp.openById(existing.spreadsheetId); }
    catch (e) { /* file was deleted — fall through and recreate */ }
  }
  const name = 'Overview — ' + (title || entityId);
  const newSs = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(newSs.getId());
  getOverviewFolder_().addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  // give the first tab a friendly name
  newSs.getSheets()[0].setName('Overview');
  const reg = overviewRegistrySheet_();
  if (existing && existing.row) {
    reg.getRange(existing.row, 2, 1, 2).setValues([[newSs.getId(), title || entityId]]);
  } else {
    reg.appendRow([entityId, newSs.getId(), title || entityId]);
  }
  return newSs;
}

// ---------------------------------------------------------------------
// "Database (auto)" tab — the first tab of every overview spreadsheet.
// It holds a filtered view of the master database for that entity (relay /
// bay / substation), written by this script. Because it lives INSIDE the
// entity's own spreadsheet, the user's other tabs can reference it with
// plain FILTER / VLOOKUP / QUERY — no IMPORTRANGE and no authorisation
// prompts. It is rewritten whenever the entity's overview is created or
// the portal asks for a database sync.
// ---------------------------------------------------------------------
const DB_TAB_NAME = 'Database (auto)';

function entityDbRows_(entityId) {
  const ss = getSS_();
  const rSh = ss.getSheetByName(RELAYS_SHEET);
  const sSh = ss.getSheetByName(SETTINGS_SHEET);
  const bSh = ss.getSheetByName(BAYS_SHEET);
  if (!rSh || !sSh || !bSh) return { header: [], rows: [], error: 'Database not synced yet.' };

  const relayRows = rSh.getLastRow() > 1 ? rSh.getRange(2, 1, rSh.getLastRow() - 1, RELAYS_HEADER.length).getValues() : [];
  const settingRows = sSh.getLastRow() > 1 ? sSh.getRange(2, 1, sSh.getLastRow() - 1, SETTINGS_HEADER.length).getValues() : [];
  const bayRows = bSh.getLastRow() > 1 ? bSh.getRange(2, 1, bSh.getLastRow() - 1, BAYS_HEADER.length).getValues() : [];

  const bayNameById = {}; bayRows.forEach(function (r) { bayNameById[r[0]] = r[2]; });
  const relayById = {}; relayRows.forEach(function (r) { relayById[r[0]] = { name: r[2], bayId: r[1], kind: r[6] }; });

  let scope;
  if (String(entityId) === 'SUBSTATION') scope = 'sub';
  else if (relayById[entityId]) scope = 'relay';
  else scope = 'bay';

  let header, rows = [];
  if (scope === 'relay') {
    header = ['Group', 'Setting', 'Value', 'Unit', 'Step', 'Range', 'Setting Meaning', 'Value Meaning'];
    settingRows.forEach(function (s) {
      if (String(s[0]) !== String(entityId)) return;
      rows.push([s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8]]);
    });
  } else if (scope === 'bay') {
    header = ['Relay', 'Relay Type', 'Group', 'Setting', 'Value', 'Unit', 'Step', 'Range', 'Setting Meaning', 'Value Meaning'];
    settingRows.forEach(function (s) {
      const r = relayById[s[0]];
      if (!r || String(r.bayId) !== String(entityId)) return;
      rows.push([r.name, r.kind, s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8]]);
    });
  } else {
    header = ['Bay', 'Relay', 'Relay Type', 'Group', 'Setting', 'Value', 'Unit', 'Step', 'Range', 'Setting Meaning', 'Value Meaning'];
    settingRows.forEach(function (s) {
      const r = relayById[s[0]];
      if (!r) return;
      rows.push([bayNameById[r.bayId] || r.bayId, r.name, r.kind, s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8]]);
    });
  }
  return { header: header, rows: rows, scope: scope };
}

function syncDbTab_(ss, entityId) {
  const data = entityDbRows_(entityId);
  if (!data.header.length) return { rows: 0, error: data.error || 'no data' };

  let sh = ss.getSheetByName(DB_TAB_NAME);
  if (!sh) {
    sh = ss.insertSheet(DB_TAB_NAME, 0);
  } else {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(1); // keep it as the first tab
  }
  sh.clear();

  const nCols = data.header.length;
  const need = data.rows.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());

  sh.getRange(1, 1, 1, nCols).setValues([data.header])
    .setFontWeight('bold').setBackground('#2E4A66').setFontColor('#EFE8D2').setHorizontalAlignment('center');
  if (data.rows.length) sh.getRange(2, 1, data.rows.length, nCols).setValues(data.rows);
  sh.setFrozenRows(1);
  return { rows: data.rows.length, scope: data.scope };
}


// ---------------------------------------------------------------------
// SHARED USER DATA — the fix for "my desktop changes aren't on my phone".
// Everything the user creates (revisions, cards, meanings, corrections,
// interruptions, users) is stored as a single JSON file in Drive rather than
// in one browser's localStorage. Any device reads the same file.
// A Drive FILE is used instead of a sheet cell because cells cap at ~50k
// characters and this payload grows past that quickly.
// Concurrency: last-write-wins, but each write carries the updatedAt it was
// based on; if Storage has moved on since, we report a conflict instead of
// silently clobbering the other device's work.
// ---------------------------------------------------------------------
const USER_DATA_FILE = 'user_data.json';

function userDataFile_() {
  const root = getRootFolder_();
  const it = root.getFilesByName(USER_DATA_FILE);
  if (it.hasNext()) return it.next();
  return root.createFile(USER_DATA_FILE, JSON.stringify({ data: {}, updatedAt: '', updatedBy: '' }), 'application/json');
}

function readUserData_() {
  try {
    const f = userDataFile_();
    let obj;
    try { obj = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { obj = {}; }
    return { ok: true, data: (obj && obj.data) || {}, updatedAt: (obj && obj.updatedAt) || '', updatedBy: (obj && obj.updatedBy) || '' };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// Timestamp only — used by the portal's poll loop so checking "has anything changed?"
// doesn't drag the whole user-data payload across the wire every 20 seconds.
function readUserDataMeta_() {
  try {
    const f = userDataFile_();
    let obj;
    try { obj = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { obj = {}; }
    return { ok: true, updatedAt: (obj && obj.updatedAt) || '', updatedBy: (obj && obj.updatedBy) || '' };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function writeUserData_(data, updatedBy, baseUpdatedAt, force) {
  const f = userDataFile_();
  let current = {};
  try { current = JSON.parse(f.getBlob().getDataAsString()) || {}; } catch (e) { current = {}; }
  const remoteAt = current.updatedAt || '';
  // Another device saved after the copy this one started from.
  if (!force && baseUpdatedAt !== undefined && baseUpdatedAt !== null && remoteAt && remoteAt !== baseUpdatedAt) {
    return { ok: false, conflict: true, remoteUpdatedAt: remoteAt, remoteUpdatedBy: current.updatedBy || '' };
  }
  const payload = { data: data, updatedAt: new Date().toISOString(), updatedBy: updatedBy || '' };
  f.setContent(JSON.stringify(payload));
  return { ok: true, updatedAt: payload.updatedAt };
}

function listOverviewTabs_(ss) {
  return ss.getSheets().filter(function (sh) { return !sh.isSheetHidden(); }).map(function (sh) {
    return { blockId: String(sh.getSheetId()), name: sh.getName(), gid: sh.getSheetId() };
  });
}
function tabByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (String(sheets[i].getSheetId()) === String(gid)) return sheets[i];
  return null;
}

// Read one tab with full formatting.
function readOverviewTab_(sh) {
  const rng = sh.getDataRange();
  const nRows = rng.getNumRows(), nCols = rng.getNumColumns();
  if (nRows === 0 || nCols === 0) return { rows: [], colWidths: [sh.getColumnWidth(1)], fmt: {}, rowHeights: [] };

  const values = rng.getDisplayValues();
  const bg = rng.getBackgrounds();
  const fc = rng.getFontColors();
  const fw = rng.getFontWeights();
  const fs = rng.getFontStyles();
  const fsz = rng.getFontSizes();
  const ha = rng.getHorizontalAlignments();
  const va = rng.getVerticalAlignments();
  const wrap = rng.getWraps();

  const colWidths = [];
  for (var c = 1; c <= nCols; c++) colWidths.push(sh.getColumnWidth(c));
  const rowHeights = [];
  for (var r = 1; r <= nRows; r++) rowHeights.push(sh.getRowHeight(r));

  // merged cells -> list of {row,col,rows,cols} (0-indexed within the data range)
  const merges = sh.getRange(1, 1, nRows, nCols).getMergedRanges().map(function (m) {
    return { r: m.getRow() - 1, c: m.getColumn() - 1, rs: m.getNumRows(), cs: m.getNumColumns() };
  });

  return {
    rows: values,
    colWidths: colWidths,
    rowHeights: rowHeights,
    merges: merges,
    fmt: { bg: bg, fc: fc, fw: fw, fs: fs, fsz: fsz, ha: ha, va: va, wrap: wrap }
  };
}

// Overwrite a tab's values from a raw grid (used by bulk Excel upload). Formatting the
// user later applies in Google Sheets is preserved on subsequent reads.
function writeOverviewTab_(sh, rows) {
  sh.clearContents();
  if (!rows || !rows.length) return;
  const numCols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const padded = rows.map(function (r) { const rr = r.slice(); while (rr.length < numCols) rr.push(''); return rr; });
  if (sh.getMaxRows() < padded.length) sh.insertRowsAfter(sh.getMaxRows(), padded.length - sh.getMaxRows());
  if (sh.getMaxColumns() < numCols) sh.insertColumnsAfter(sh.getMaxColumns(), numCols - sh.getMaxColumns());
  sh.getRange(1, 1, padded.length, numCols).setValues(padded);
}

// Persist a single column's width, set by dragging on the website.
function setOverviewColWidth_(sh, colIndex1, width) {
  if (colIndex1 >= 1 && width > 20) sh.setColumnWidth(colIndex1, Math.round(width));
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
  // A new sheet defaults to 1000 rows / 26 cols. Bulk-writing a range larger than the
  // current grid throws "range exceeds grid limits", so grow the grid FIRST. This is the
  // key fix for large sheets like _Settings (45k+ rows).
  const needRows = rows.length + 1;      // +1 for the header row
  const needCols = header.length;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());

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
    bays: bays, busZones: busZones, syncedAt: meta.syncedAt || '',
    backendVersion: BACKEND_VERSION
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

function sheetUrlFor_(ss, gid) {
  return ss.getUrl() + (gid != null ? ('#gid=' + gid) : '');
}

// GET actions: getFullData, getSld, listBlocks (tabs of an entity's sheet), get (one tab w/ formatting)
function doGet(e) {
  try {
    const action = e.parameter.action || 'get';

    if (action === 'version') return jsonOut_({ ok: true, backendVersion: BACKEND_VERSION });
    if (action === 'getFullData') return jsonOut_(readFullDatabase_());
    if (action === 'getSld') return jsonOut_(readSld_());
    if (action === 'getUserData') return jsonOut_(readUserData_());
    if (action === 'getUserDataMeta') return jsonOut_(readUserDataMeta_());

    const entityId = e.parameter.entityId;
    if (!entityId) return jsonOut_({ ok: false, error: 'entityId required' });

    if (action === 'listBlocks') {
      const lk = overviewLookup_(entityId);
      if (!lk || !lk.spreadsheetId) return jsonOut_({ ok: true, exists: false, blocks: [] });
      const ss = SpreadsheetApp.openById(lk.spreadsheetId);
      return jsonOut_({ ok: true, exists: true, spreadsheetUrl: ss.getUrl(), blocks: listOverviewTabs_(ss) });
    }
    if (action === 'get') {
      const gid = e.parameter.blockId;
      const lk = overviewLookup_(entityId);
      if (!lk || !lk.spreadsheetId) return jsonOut_({ ok: true, exists: false, rows: [] });
      const ss = SpreadsheetApp.openById(lk.spreadsheetId);
      const sh = gid ? tabByGid_(ss, gid) : ss.getSheets()[0];
      if (!sh) return jsonOut_({ ok: true, exists: false, rows: [] });
      const data = readOverviewTab_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh.getSheetId()), spreadsheetUrl: ss.getUrl() }, data));
    }
    return jsonOut_({ ok: false, error: 'unknown action for GET: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// POST actions: bootstrap, getFullData, getSld, ensure (entity spreadsheet + optional new tab),
//               listBlocks, get, save (bulk rows into a tab), addTab, deleteTab, setColWidth
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'bootstrap') {
      const result = bootstrapDatabase_(body);
      return jsonOut_(Object.assign({ ok: true }, result));
    }
    if (action === 'version') return jsonOut_({ ok: true, backendVersion: BACKEND_VERSION });
    if (action === 'getFullData') return jsonOut_(readFullDatabase_());
    if (action === 'getSld') return jsonOut_(readSld_());
    if (action === 'getUserData') return jsonOut_(readUserData_());
    if (action === 'getUserDataMeta') return jsonOut_(readUserDataMeta_());
    if (action === 'saveUserData') {
      return jsonOut_(writeUserData_(body.data || {}, body.updatedBy, body.baseUpdatedAt, body.force));
    }

    const entityId = body.entityId;
    if (!entityId) return jsonOut_({ ok: false, error: 'entityId required' });

    // Ensure this entity's own spreadsheet exists; returns its tabs + first tab content.
    if (action === 'ensure') {
      const ss = ensureOverviewSpreadsheet_(entityId, body.title);
      // First tab is always the auto-filtered database view for this entity.
      const dbInfo = syncDbTab_(ss, entityId);
      const first = ss.getSheets()[0];
      const data = readOverviewTab_(first);
      return jsonOut_(Object.assign({ ok: true, spreadsheetUrl: ss.getUrl(), sheetUrl: sheetUrlFor_(ss, first.getSheetId()),
        blocks: listOverviewTabs_(ss), dbRows: dbInfo.rows }, data));
    }

    // Rebuild just the "Database (auto)" tab from the master database.
    if (action === 'syncDb') {
      const ss2 = overviewLookup_(entityId) ? SpreadsheetApp.openById(overviewLookup_(entityId).spreadsheetId)
                                            : ensureOverviewSpreadsheet_(entityId, body.title);
      const info = syncDbTab_(ss2, entityId);
      if (info.error) return jsonOut_({ ok: false, error: info.error });
      return jsonOut_({ ok: true, dbRows: info.rows, blocks: listOverviewTabs_(ss2), spreadsheetUrl: ss2.getUrl() });
    }

    if (action === 'listBlocks') {
      const lk = overviewLookup_(entityId);
      if (!lk) return jsonOut_({ ok: true, exists: false, blocks: [] });
      const ss = SpreadsheetApp.openById(lk.spreadsheetId);
      return jsonOut_({ ok: true, exists: true, spreadsheetUrl: ss.getUrl(), blocks: listOverviewTabs_(ss) });
    }

    const lk = overviewLookup_(entityId);
    const ss = lk ? SpreadsheetApp.openById(lk.spreadsheetId) : ensureOverviewSpreadsheet_(entityId, body.title);

    if (action === 'get') {
      const sh = body.blockId ? tabByGid_(ss, body.blockId) : ss.getSheets()[0];
      if (!sh) return jsonOut_({ ok: true, exists: false, rows: [] });
      const data = readOverviewTab_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh.getSheetId()), spreadsheetUrl: ss.getUrl() }, data));
    }

    // Add a new tab (new named sheet) to this entity's spreadsheet.
    if (action === 'addTab') {
      const name = (body.name || 'New tab').slice(0, 90);
      const sh = ss.insertSheet(name);
      return jsonOut_({ ok: true, blockId: String(sh.getSheetId()), name: sh.getName(), blocks: listOverviewTabs_(ss) });
    }

    // Bulk overwrite a tab from parsed Excel rows (creates the tab if a name is given & missing).
    if (action === 'save') {
      let sh = body.blockId ? tabByGid_(ss, body.blockId) : ss.getSheets()[0];
      if (!sh && body.name) sh = ss.insertSheet(String(body.name).slice(0, 90));
      if (!sh) return jsonOut_({ ok: false, error: 'tab not found' });
      writeOverviewTab_(sh, body.rows || []);
      return jsonOut_({ ok: true, blockId: String(sh.getSheetId()), sheetUrl: sheetUrlFor_(ss, sh.getSheetId()) });
    }

    // Persist a column width dragged on the website.
    if (action === 'setColWidth') {
      const sh = body.blockId ? tabByGid_(ss, body.blockId) : ss.getSheets()[0];
      if (sh) setOverviewColWidth_(sh, Number(body.col) + 1, Number(body.width));
      return jsonOut_({ ok: true });
    }

    if (action === 'deleteTab') {
      const sh = tabByGid_(ss, body.blockId);
      if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
      return jsonOut_({ ok: true, blocks: listOverviewTabs_(ss) });
    }

    return jsonOut_({ ok: false, error: 'unknown action for POST: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
