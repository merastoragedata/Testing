/**
 * ============================================================================
 * Karjat Protection Settings Portal — "Relay Settings Overview" backend
 * ============================================================================
 * This Apps Script turns a Google Sheet into the backing store for the
 * portal's "Relay Settings Overview" feature: one tab per relay, styled like
 * a Quick Reference sheet (Section header rows, then Parameter / Value /
 * Unit / Notes rows), with the Parameter column restricted to a dropdown of
 * that relay's real setting titles so entries stay linked to actual data.
 *
 * The portal calls this script's Web App URL directly from the browser
 * (no login prompt for the person using the portal) and reads back JSON.
 * Full editing — adding rows, picking from the dropdown, writing notes — is
 * done by opening the returned sheet URL in a normal Google Sheets tab;
 * the portal itself only shows a read-only mirror plus "Refresh".
 *
 * ----------------------------------------------------------------------
 * DEPLOY THIS AS A WEB APP
 * ----------------------------------------------------------------------
 * 1. Go to https://script.google.com, create a new project, and paste this
 *    whole file in as Code.gs (replace the default content).
 * 2. (Optional but recommended) Run once manually: select the function
 *    `getSS_` in the toolbar dropdown and click Run, to authorise the
 *    script and create the master spreadsheet. Check the Execution Log
 *    for a line like "Created new spreadsheet, id=...— paste this into
 *    SS_ID for stability." Copy that ID into the SS_ID constant below so
 *    the script always reuses the same spreadsheet (otherwise it searches
 *    Drive by filename each time, which works but is slightly slower).
 * 3. Click Deploy → New deployment.
 *      Type: "Web app"
 *      Execute as: "Me"
 *      Who has access: "Anyone"
 *    (This lets the offline HTML portal call it from any browser without
 *    the person needing a Google login — the script itself runs with your
 *    permissions to read/write the sheet.)
 * 4. Click Deploy, authorise the requested permissions, and copy the
 *    resulting URL (ends in /exec).
 * 5. Send me that URL — I'll hardcode it into the portal's
 *    GOOGLE_SHEET_API_URL constant and the "Relay Settings Overview"
 *    section will go live with no other changes needed.
 *
 * If you ever need to redeploy after editing this script, use
 * Deploy → Manage deployments → Edit (pencil icon) → New version, so the
 * /exec URL already hardcoded into the portal keeps working.
 * ============================================================================
 */

// Paste your spreadsheet ID here after the first run (see step 2 above).
// Leave blank and the script will find-or-create one called
// "Karjat Relay Overviews" in your Drive automatically.
const SS_ID = '';

const LOOKUP_SHEET = '_lookup';
const HEADER = ['Section', 'Parameter', 'Value', 'Unit', 'Notes / Meaning'];
const DATA_ROWS_HEADROOM = 500; // how many rows of dropdown validation to pre-apply per tab

/* ---------------------------------------------------------------------- */
/* Spreadsheet / tab helpers                                              */
/* ---------------------------------------------------------------------- */

function getSS_() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  const name = 'Karjat Relay Overviews';
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  const ss = SpreadsheetApp.create(name);
  Logger.log('Created new spreadsheet, id=' + ss.getId() +
    ' — paste this into SS_ID for stability.');
  return ss;
}

function tabName_(relayId) {
  // Sheet tab names: max 100 chars, and can't contain [ ] * / \ ? :
  return String(relayId).replace(/[\[\]\*\/\\\?:]/g, '_').slice(0, 90);
}

function ensureLookup_(ss, relayId, settingsList) {
  let sh = ss.getSheetByName(LOOKUP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOOKUP_SHEET);
    sh.hideSheet();
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let col = header.indexOf(relayId) + 1;
  if (col === 0) {
    col = lastCol + (header[lastCol - 1] ? 1 : 0);
    if (col === 0) col = 1;
    sh.getRange(1, col).setValue(relayId);
  }
  const labels = (settingsList || []).map(function (s) {
    return s.group + ' — ' + s.title;
  });
  // clear old values below the header in this column first, then write fresh
  const maxRows = Math.max(sh.getMaxRows() - 1, labels.length);
  if (maxRows > 0) sh.getRange(2, col, maxRows, 1).clearContent();
  if (labels.length) sh.getRange(2, col, labels.length, 1).setValues(labels.map(function (l) { return [l]; }));
  return { col: col, count: labels.length };
}

function ensureTab_(ss, relayId, relayName, bayName, settingsList, initialSections) {
  const name = tabName_(relayId);
  let sh = ss.getSheetByName(name);
  if (sh) return sh; // already exists — leave existing content untouched

  sh = ss.insertSheet(name);
  sh.getRange(1, 1).setValue(relayName || relayId).setFontWeight('bold').setFontSize(12);
  sh.getRange(2, 1).setValue(bayName || '').setFontColor('#6B7A99');
  sh.getRange(3, 1, 1, HEADER.length).setValues([HEADER]);
  sh.getRange(3, 1, 1, HEADER.length)
    .setFontWeight('bold').setBackground('#2A4A6B').setFontColor('#FBEFC9');
  sh.setFrozenRows(3);
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 380);

  const lookup = ensureLookup_(ss, relayId, settingsList);
  if (lookup.count > 0) {
    const lookupSheet = ss.getSheetByName(LOOKUP_SHEET);
    const colA1 = columnToLetter_(lookup.col);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lookupSheet.getRange(colA1 + '2:' + colA1 + (1 + lookup.count)), true)
      .setAllowInvalid(true) // engineers can still type a value not in the list if needed
      .setHelpText('Pick a setting from this relay, or type your own label.')
      .build();
    sh.getRange(4, 2, DATA_ROWS_HEADROOM, 1).setDataValidation(rule); // Parameter column
  }

  let row = 4;
  (initialSections || []).forEach(function (sec) {
    sh.getRange(row, 1).setValue(sec.name).setFontWeight('bold').setBackground('#FAF3DC');
    row++;
    (sec.entries || []).forEach(function (en) {
      sh.getRange(row, 2, 1, 4).setValues([[en.parameter || '', en.value || '', en.unit || '', en.notes || '']]);
      row++;
    });
    row++; // blank spacer row between sections
  });

  return sh;
}

function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - rem) / 26);
  }
  return letter;
}

/* ---------------------------------------------------------------------- */
/* Read / write the Section / Parameter / Value / Unit / Notes rows       */
/* ---------------------------------------------------------------------- */

function readTab_(sh) {
  const lastRow = sh.getLastRow();
  const relayName = sh.getRange(1, 1).getValue();
  const bayName = sh.getRange(2, 1).getValue();
  if (lastRow < 4) return { relayName: relayName, bayName: bayName, sections: [] };

  const values = sh.getRange(4, 1, lastRow - 3, 5).getValues();
  const sections = [];
  let cur = null;
  values.forEach(function (row) {
    const section = row[0], parameter = row[1], value = row[2], unit = row[3], notes = row[4];
    const isSectionHeader = section && !parameter && !value && !notes;
    if (isSectionHeader) {
      cur = { name: String(section), entries: [] };
      sections.push(cur);
    } else if (parameter || value || notes) {
      if (!cur) { cur = { name: '(General)', entries: [] }; sections.push(cur); }
      cur.entries.push({
        parameter: String(parameter || ''),
        value: String(value || ''),
        unit: String(unit || ''),
        notes: String(notes || '')
      });
    }
  });
  return { relayName: relayName, bayName: bayName, sections: sections };
}

function writeTab_(sh, sections) {
  const lastRow = sh.getLastRow();
  if (lastRow >= 4) sh.getRange(4, 1, lastRow - 3, 5).clearContent();
  let row = 4;
  (sections || []).forEach(function (sec) {
    sh.getRange(row, 1).setValue(sec.name).setFontWeight('bold').setBackground('#FAF3DC');
    row++;
    (sec.entries || []).forEach(function (en) {
      sh.getRange(row, 2, 1, 4).setValues([[en.parameter || '', en.value || '', en.unit || '', en.notes || '']]);
      row++;
    });
    row++;
  });
}

/* ---------------------------------------------------------------------- */
/* Web app entry points                                                   */
/* ---------------------------------------------------------------------- */

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
    const relayId = e.parameter.relayId;
    if (!relayId) return jsonOut_({ ok: false, error: 'relayId required' });

    const ss = getSS_();
    const sh = ss.getSheetByName(tabName_(relayId));
    if (action === 'get') {
      if (!sh) return jsonOut_({ ok: true, exists: false, sections: [] });
      const data = readTab_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }
    return jsonOut_({ ok: false, error: 'unknown action for GET: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// POST body (JSON): {action:'ensure'|'save'|'get', relayId, relayName, bayName,
//                     settingsList:[{group,title}], sections|initialSections}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const relayId = body.relayId;
    if (!relayId) return jsonOut_({ ok: false, error: 'relayId required' });

    const ss = getSS_();

    if (action === 'get') {
      const sh = ss.getSheetByName(tabName_(relayId));
      if (!sh) return jsonOut_({ ok: true, exists: false, sections: [] });
      const data = readTab_(sh);
      return jsonOut_(Object.assign({ ok: true, exists: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }

    // Creates the tab (with dropdown-linked Parameter column) if it doesn't exist yet,
    // seeding it with initialSections the first time only. Always returns current data
    // and the sheet URL, so the portal can open straight to this relay's tab.
    if (action === 'ensure') {
      const sh = ensureTab_(ss, relayId, body.relayName, body.bayName, body.settingsList, body.initialSections);
      const data = readTab_(sh);
      return jsonOut_(Object.assign({ ok: true, sheetUrl: sheetUrlFor_(ss, sh) }, data));
    }

    // Overwrites the tab's data rows with exactly the given sections — used when the
    // portal wants to push local-only edits (made before the sheet was connected, or
    // added quickly in-app) back up into the sheet.
    if (action === 'save') {
      let sh = ss.getSheetByName(tabName_(relayId));
      if (!sh) sh = ensureTab_(ss, relayId, body.relayName, body.bayName, body.settingsList, []);
      writeTab_(sh, body.sections || []);
      return jsonOut_({ ok: true, sheetUrl: sheetUrlFor_(ss, sh) });
    }

    return jsonOut_({ ok: false, error: 'unknown action for POST: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
