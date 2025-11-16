/**
 * Exports "Clipboard" A1:L100 as PDFs for each dropdown value in the external sheet's B2.
 * - MODIFIED: Preserves the file ID/URL of existing PDFs by updating their content.
 * - Values-only copy into a local "Print" sheet (no formulas), then export that sheet.
 * - Waits for IMPORTRANGE to refresh after setting each name (using hash and witness cell).
 * - Uses OUTPUT_FOLDER_ID if valid; otherwise falls back to "Clipboard Exports (auto)" in My Drive.
 */

// -----------------------------------------------------------------------------
// == CONFIGURATION CONSTANTS ==
// -----------------------------------------------------------------------------

// --- File & Sheet Settings ---
const OUTPUT_FOLDER_ID = '1LkqlKN80bug1MrR0vrzygjJ4QfBL6Zxg'; // <-- paste your folder ID or leave blank
const REPORT_SHEET_NAME = 'Clipboard'; // Printable tab in THIS file (Source of data)
const PRINT_SHEET_NAME = 'Print'; // Temp output tab (auto-managed, where values are copied)
const SRC_RANGE_A1 = 'A1:L100'; // Area to print/copy
const LOCAL_Q1_RANGE_A1 = 'Clipboard!Q1'; // Cell containing full URL of the external file
const EXTERNAL_B2_A1 = 'B2'; // Dropdown cell in the external file (source of names)
const CSV_FILENAME = 'Clipboard Export Log.csv'; // name of the log file

// --- Timing & Backoff ---
const WAIT_MS_AFTER_SET = 2500; // Original requested wait time (now handled by waitFor functions)
const BACKOFF_MAX_RETRIES = 5;
const BACKOFF_BASE_MS = 600; // Start wait time for network retries
const BETWEEN_JOBS_MS = 250; // Small pause after each PDF export

// --- PDF Export Options ---
const EXPORT_OPTS = {
  format: 'pdf',
  size: 'letter',
  portrait: 'true', // 'true' portrait, 'false' landscape
  scale: '2', // 1=Normal, 2=Fit to width, 3=Fit to page (Fit to width)
  top_margin: '0.5',
  bottom_margin: '0.5',
  left_margin: '0.5',
  right_margin: '0.5',
  sheetnames: 'false',
  printtitle: 'false',
  pagenumbers: 'false',
  gridlines: 'true',
  fzr: 'false'
};

// --- Refresh/Wait Settings ---
const WITNESS_CELL_A1 = 'B1'; // Cell on REPORT_SHEET that reflects the selected name
const WITNESS_TIMEOUT_MS = 10000; // 10s max wait for witness cell
const WITNESS_POLL_MS = 300; // poll every 0.3s
const REFRESH_TIMEOUT_MS = 10000; // 10s max wait for data change
const POLL_INTERVAL_MS = 300; // 0.3s poll interval

// --- Multi-pass Settings ---
const MAX_NAMES_PER_RUN = 15;                  // how many names to process per execution
const STATE_KEY = 'ClipboardExportState';      // script property key for progress

// -----------------------------------------------------------------------------
// == MAIN FUNCTION (NOW MULTI-PASS) ==
// -----------------------------------------------------------------------------

function BatchExportClipboardForAllNames() {
  const ss = SpreadsheetApp.getActive();
  const report = mustGetSheet_(ss, REPORT_SHEET_NAME);
  const folder = getOutputFolder_();
  const scriptProps = PropertiesService.getScriptProperties();

  // 1) Get External Sheet and Name List
  const sourceUrl = ss.getRange(LOCAL_Q1_RANGE_A1).getValue().toString().trim();
  const sourceId = extractIdFromUrl_(sourceUrl);
  if (!sourceId) throw new Error(`Could not extract spreadsheet ID from: ${sourceUrl}`);

  const ext = SpreadsheetApp.openById(sourceId);
  const extSheet = findSheetWithB2Dropdown_(ext) || ext.getSheets()[0];
  const b2 = extSheet.getRange(EXTERNAL_B2_A1);
  const names = getDropdownValues_(b2);
  if (!names.length) throw new Error('No names found from dropdown in external B2.');

  // 1b) Determine where to start (multi-pass state)
  let startIndex = 0;
  const stateStr = scriptProps.getProperty(STATE_KEY);
  if (stateStr) {
    try {
      const st = JSON.parse(stateStr);
      if (st && typeof st.nextIndex === 'number') {
        startIndex = st.nextIndex;
      }
    } catch (e) {
      // Bad state; start from 0
      startIndex = 0;
    }
  }

  // If we already finished previously, reset and start from 0 again
  if (startIndex >= names.length) {
    scriptProps.deleteProperty(STATE_KEY);
    ss.toast('All names were already exported. Starting over from the first name.');
    startIndex = 0;
  }

  const endIndex = Math.min(names.length, startIndex + MAX_NAMES_PER_RUN);

  // 2) Prepare Local Print Sheet
  const srcRange = report.getRange(SRC_RANGE_A1);
  const rows = srcRange.getNumRows(),
        cols = srcRange.getNumColumns();
  const printSheet = ensurePrintSheetSized_(ss, PRINT_SHEET_NAME, rows, cols);
  copyFormatsOnce_(report, printSheet, SRC_RANGE_A1);

  // 3) Build Export URL and Headers
  const ssId = ss.getId();
  const gid = printSheet.getSheetId();
  const query = Object.keys(EXPORT_OPTS).map(k => `${k}=${encodeURIComponent(EXPORT_OPTS[k])}`).join('&');
  const baseExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?${query}&gid=${gid}`;
  const headers = {
    Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
  };
  const dstRange = printSheet.getRange(1, 1, rows, cols);

  // 4) Loop through each name IN THIS PASS ONLY
  const witnessCell = report.getRange(WITNESS_CELL_A1);
  let prevHash = hashValues_(srcRange.getValues()); // Baseline hash before first change in this pass

  for (let i = startIndex; i < endIndex; i++) {
    const name = names[i];

    // A) Set the external selector & flush
    b2.setValue(name);
    SpreadsheetApp.flush();

    // B) Wait for witness cell (B1) to show the new name
    waitForWitnessCell_(witnessCell, name);

    // C) Wait for the main data range (A1:L100) to change/refresh
    const { vals, hash } = waitForRangeChange_(srcRange, prevHash);

    // D) Paste values into Print sheet & flush
    dstRange.clearContent();
    dstRange.setValues(vals);
    SpreadsheetApp.flush();
    Utilities.sleep(100); // Small settle time

    // E) Export to PDF (with backoff) and UPDATE/CREATE the file
    const resp = fetchWithBackoff_(baseExportUrl, headers);
    const blob = resp.getBlob();
    const fname = `Clipboard for - ${sanitize_(name)}.pdf`;

    const pdfFile = createOrUpdateFile_(folder, blob, fname);

    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // F) Log the new file link to CSV
    upsertCsvRow_(folder, name, pdfFile.getUrl());

    // G) Update baseline and pause
    prevHash = hash;
    Utilities.sleep(BETWEEN_JOBS_MS);
  }

  // 5) Save progress for next pass
  if (endIndex >= names.length) {
    scriptProps.deleteProperty(STATE_KEY);
    ss.toast(`Finished all ${names.length} exports in this pass.`);
  } else {
    scriptProps.setProperty(
      STATE_KEY,
      JSON.stringify({ nextIndex: endIndex, total: names.length })
    );
    ss.toast(
      `Exported ${endIndex} of ${names.length} names so far. ` +
      `Run "Batch Export → Export PDFs for all names" again to continue.`
    );
  }
}

// -----------------------------------------------------------------------------
// == RESET FUNCTION FOR MULTI-PASS STATE ==
// -----------------------------------------------------------------------------

function ResetClipboardExportProgress() {
  PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
  SpreadsheetApp.getActive().toast('Clipboard export progress has been reset.');
}

// -----------------------------------------------------------------------------
// == MODIFIED HELPER FUNCTION - DRIVE I/O ==
// -----------------------------------------------------------------------------

/**
 * Creates a new file or updates the content of an existing file with the same name.
 * Uses Advanced Drive API so the original File ID and URL are preserved.
 * NOTE: Requires Advanced Drive service enabled as "Drive" (v2).
 */
function createOrUpdateFile_(folder, blob, filename) {
  // Look for existing file by name in this folder
  const it = folder.getFilesByName(filename);

  if (it.hasNext()) {
    const existingFile = it.next();
    const fileId = existingFile.getId();

    // Safety: trash any additional duplicates with the same name
    while (it.hasNext()) {
      it.next().setTrashed(true);
    }

    // Resource body for Drive.Files.update (Drive v2)
    const resource = {
      title: filename,
      mimeType: blob.getContentType(),
      parents: [{ id: folder.getId() }]
    };

    // Overwrite file content while preserving the file ID
    const updated = Drive.Files.update(resource, fileId, blob);

    // Return a DriveApp.File so the rest of your code works unchanged
    return DriveApp.getFileById(updated.id);
  }

  // No existing file with that name: create a new one
  return folder.createFile(blob.setName(filename));
}

// -----------------------------------------------------------------------------
// == OTHER HELPER FUNCTIONS (UNCHANGED) ==
// -----------------------------------------------------------------------------

// Retrieves the output folder, falling back to an auto-created one if the ID fails.
function getOutputFolder_() {
  try {
    if (OUTPUT_FOLDER_ID && !OUTPUT_FOLDER_ID.includes('PUT_YOUR_FOLDER_ID_HERE')) {
      return DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    }
  } catch (e) {
    SpreadsheetApp.getActive().toast('Could not open folder by ID; using "Clipboard Exports (auto)" in My Drive.');
  }
  return ensureFolderByName_('Clipboard Exports (auto)');
}

// Gets a folder by name, creating it if necessary.
function ensureFolderByName_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// Finds the sheet in the external spreadsheet that contains the required dropdown in B2.
function findSheetWithB2Dropdown_(extSpreadsheet) {
  const sheets = extSpreadsheet.getSheets();
  for (const sh of sheets) {
    const rule = sh.getRange(EXTERNAL_B2_A1).getDataValidation();
    if (!rule) continue;
    const t = rule.getCriteriaType();
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
        t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return sh;
    }
  }
  return null;
}

// Extracts the values used in the dropdown menu for the given range.
function getDropdownValues_(range) {
  const rule = range.getDataValidation();
  if (!rule) return [];
  const type = rule.getCriteriaType();
  const vals = rule.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return (vals[0] || []).map(v => (v || '').toString().trim()).filter(Boolean);
  }
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const src = vals[0];
    if (!src) return [];
    return src.getValues().flat().map(v => (v || '').toString().trim()).filter(Boolean);
  }
  const v = (range.getValue() || '').toString().trim();
  return v ? [v] : [];
}

// Ensures the temporary print sheet exists and is sized to fit the data.
function ensurePrintSheetSized_(ss, name, rows, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getMaxRows() < rows) sh.insertRows(1, rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumns(1, cols - sh.getMaxColumns());
  return sh;
}

// Copies only the formats and column widths from the report sheet to the print sheet.
function copyFormatsOnce_(reportSheet, printSheet, srcA1) {
  const src = reportSheet.getRange(srcA1);
  const rows = src.getNumRows(),
        cols = src.getNumColumns();
  const dst = printSheet.getRange(1, 1, rows, cols);
  dst.clear({ contentsOnly: true });
  src.copyTo(dst, { formatOnly: true });
  for (let c = 1; c <= cols; c++) {
    printSheet.setColumnWidth(c, reportSheet.getColumnWidth(c));
  }
}

// Retrieves a sheet by name or throws an error if not found.
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" not found.`);
  return sh;
}

// Extracts the spreadsheet ID from a Google Sheet URL.
function extractIdFromUrl_(url) {
  const m = (url || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : '';
}

// Sanitizes a string for use as a file name.
function sanitize_(s) {
  return s.replace(/[\\/:*?"<>|]/g, '_').trim();
}

// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Batch Export')
    .addItem('Export PDFs for all names', 'BatchExportClipboardForAllNames')
    .addItem('Reset export progress', 'ResetClipboardExportProgress')
    .addToUi();
}

// Waits for a specific cell's display value to match an expected value (e.g., IMPORTRANGE result).
function waitForWitnessCell_(cellRange, expected, timeoutMs = WITNESS_TIMEOUT_MS) {
  const want = String(expected).trim();
  const t0 = Date.now();
  while (true) {
    SpreadsheetApp.flush();
    const got = String(cellRange.getDisplayValue()).trim();
    if (got === want) return;
    if (Date.now() - t0 > timeoutMs) {
      throw new Error(`Timed out waiting for ${cellRange.getA1Notation()} to equal "${want}" (got "${got}")`);
    }
    Utilities.sleep(WITNESS_POLL_MS);
  }
}

// Generates a deterministic hash for a 2D array of values.
function hashValues_(vals) {
  return vals.map(r => r.join('\u0001')).join('\u0002');
}

// Waits for the data range's hash to change, signaling a refresh is complete.
function waitForRangeChange_(range, previousHash, timeoutMs = REFRESH_TIMEOUT_MS) {
  const t0 = Date.now();
  while (true) {
    SpreadsheetApp.flush();
    const vals = range.getValues();
    const h = hashValues_(vals);
    if (h !== previousHash) return { vals, hash: h };
    if (Date.now() - t0 > timeoutMs) {
      throw new Error('Timed out waiting for A1:L100 to refresh.');
    }
    Utilities.sleep(POLL_INTERVAL_MS);
  }
}

// Fetches a URL with exponential backoff on retryable (429, 5xx) errors.
function fetchWithBackoff_(url, headers) {
  let attempt = 0, lastErr = null;
  while (attempt <= BACKOFF_MAX_RETRIES) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        headers,
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();

      if (code >= 200 && code < 300) return resp;

      if (code === 429 || (code >= 500 && code < 600)) {
        lastErr = new Error(`HTTP ${code}: ${resp.getContentText().slice(0,200)}`);
      } else {
        throw new Error(`HTTP ${code}: ${resp.getContentText().slice(0,200)}`);
      }
    } catch (e) {
      lastErr = e;
    }
    const wait = BACKOFF_BASE_MS * Math.pow(2, attempt) + Math.floor(Math.random() * 300);
    Utilities.sleep(wait);
    attempt++;
  }
  throw lastErr || new Error('fetchWithBackoff_ failed.');
}

// Gets the CSV log file, creating it with a header if it doesn't exist.
function getOrCreateCsv_(folder) {
  const it = folder.getFilesByName(CSV_FILENAME);
  if (it.hasNext()) return it.next();
  const header = 'Name,URL\n';
  return folder.createFile(CSV_FILENAME, header, MimeType.CSV);
}

// Updates the CSV log: ensures exactly one row per name (latest link wins).
function upsertCsvRow_(folder, name, url) {
  const csvFile = getOrCreateCsv_(folder);
  const map = readCsvMap_(csvFile);
  map.set(name, url);
  writeCsvFromMap_(csvFile, map);
  return csvFile;
}

// Reads CSV into a Map (Name -> URL). Handles quoted fields with "" escapes.
function readCsvMap_(csvFile) {
  const text = csvFile.getBlob().getDataAsString() || '';
  const lines = text.split(/\r?\n/).filter(Boolean);
  const map = new Map();
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (i === 0 && /^name\s*,\s*url$/i.test(line.trim())) continue;
    const rec = parseCsv2Fields_(line);
    if (!rec) continue;
    const [name, url] = rec;
    if (name) map.set(name, url || '');
  }
  return map;
}

// Writes the Map back to the CSV file (sorted by name).
function writeCsvFromMap_(csvFile, map) {
  const rows = Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  let out = 'Name,URL\n';
  for (const [name, url] of rows) {
    out += csvQ_(name) + ',' + csvQ_(url) + '\n';
  }
  csvFile.setContent(out);
}

// Parses a CSV line assuming exactly two fields, handling quoting rules.
function parseCsv2Fields_(line) {
  const res = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQ = false;
        }
      } else {
        cur += ch;
      }
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') {
        res.push(cur);
        cur = '';
      } else cur += ch;
    }
  }
  res.push(cur);
  if (res.length !== 2) return null;
  return res;
}

// Quotes a string and escapes internal double quotes for CSV format.
function csvQ_(s) {
  s = s == null ? '' : String(s);
  return '"' + s.replace(/"/g, '""') + '"';
}

