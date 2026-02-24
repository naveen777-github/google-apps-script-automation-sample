const TAB_CONFIG = "config";
const TAB_DATA = "data";
const TAB_SUMMARY = "summary";
const TAB_LOGS = "logs";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Automation Sample")
    .addItem("Run Import (Config)", "runImportFromConfig")
    .addSeparator()
    .addItem("Clear Data", "clearData")
    .addToUi();
}

function runImportFromConfig() {
  const start = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const configSheet = getOrCreateSheet_(ss, TAB_CONFIG);
  const dataSheet = getOrCreateSheet_(ss, TAB_DATA);
  const summarySheet = getOrCreateSheet_(ss, TAB_SUMMARY);
  const logsSheet = getOrCreateSheet_(ss, TAB_LOGS);

  try {
    ensureHeaders_(dataSheet, ["timestamp", "id", "name", "type", "dimension"]);
    ensureHeaders_(summarySheet, ["metric", "value"]);
    ensureHeaders_(logsSheet, ["timestamp", "level", "message", "context"]);

    const cfg = readConfig_(configSheet);
    const baseUrl = cfg.api_url;
    const maxPages = parseInt(cfg.max_pages || "1", 10);
    const mode = (cfg.mode || "upsert").toLowerCase();

    log_(logsSheet, "INFO", "Starting import", { baseUrl, maxPages, mode });

    // 1) Fetch pages
    const items = fetchLocationsPages_(baseUrl, maxPages, logsSheet);

    // 2) Upsert into sheet
    const result = upsertRows_(dataSheet, items, mode);

    // 3) Write summary insights
    writeSummary_(summarySheet, dataSheet, result);

    const durationMs = Date.now() - start;
    log_(logsSheet, "INFO", "Import complete", { ...result, durationMs });

    SpreadsheetApp.getUi().alert(
      `Done. Imported: ${result.inserted}, Updated: ${result.updated}, Skipped: ${result.skipped}. (${Math.round(durationMs/1000)}s)`
    );
  } catch (err) {
    log_(logsSheet, "ERROR", "Import failed", { error: String(err) });

    // Optional email alert (uncomment + set your email if you want)
    // GmailApp.sendEmail("your.email@example.com", "Automation Sample: Import failed", String(err));

    SpreadsheetApp.getUi().alert(`Import failed. Check "${TAB_LOGS}" tab.`);
  }
}

/* ------------------ Fetching ------------------ */

function fetchLocationsPages_(baseUrl, maxPages, logsSheet) {
  const all = [];
  for (let page = 1; page <= maxPages; page++) {
    const url = `${baseUrl}?page=${page}`;
    log_(logsSheet, "INFO", "Fetching page", { page, url });

    const res = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: { Accept: "application/json" }
    });

    const code = res.getResponseCode();
    if (code !== 200) {
      log_(logsSheet, "ERROR", "Non-200 response", { page, code, body: res.getContentText().slice(0, 200) });
      throw new Error(`API returned status ${code} on page ${page}`);
    }

    const json = JSON.parse(res.getContentText());
    if (!json || !Array.isArray(json.results)) {
      log_(logsSheet, "ERROR", "Unexpected JSON shape", { page, keys: Object.keys(json || {}) });
      throw new Error("Unexpected API response structure");
    }

    // Basic validation + normalization
    json.results.forEach(loc => {
      if (!loc || !loc.id || !loc.name) return;
      all.push({
        id: String(loc.id).trim(),
        name: String(loc.name).trim(),
        type: String(loc.type || "").trim(),
        dimension: String(loc.dimension || "").trim()
      });
    });
  }
  return all;
}

/* ------------------ Upsert ------------------ */

function upsertRows_(dataSheet, items, mode) {
  // mode: "append" or "upsert"
  const now = new Date();

  // Read existing ids -> row index
  const lastRow = dataSheet.getLastRow();
  const existingMap = new Map();

  if (lastRow >= 2) {
    const idValues = dataSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B = id
    idValues.forEach((row, i) => {
      const id = String(row[0] || "").trim();
      if (id) existingMap.set(id, i + 2); // actual sheet row
    });
  }

  let inserted = 0, updated = 0, skipped = 0;

  const appendRows = [];
  const updateOps = []; // {rowIndex, values}

  items.forEach(item => {
    if (!item.id) { skipped++; return; }

    const values = [now, item.id, item.name, item.type, item.dimension];

    if (mode === "append") {
      appendRows.push(values);
      inserted++;
      return;
    }

    // upsert
    const rowIndex = existingMap.get(item.id);
    if (rowIndex) {
      updateOps.push({ rowIndex, values });
      updated++;
    } else {
      appendRows.push(values);
      inserted++;
    }
  });

  // Batch append
  if (appendRows.length > 0) {
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
  }

  // Batch update (still efficient enough for small sample)
  updateOps.forEach(op => {
    dataSheet.getRange(op.rowIndex, 1, 1, op.values.length).setValues([op.values]);
  });

  return { inserted, updated, skipped, totalFetched: items.length };
}

/* ------------------ Summary ------------------ */

function writeSummary_(summarySheet, dataSheet, importResult) {
  // Compute simple insights: total rows, distinct types, top 5 types
  const lastRow = dataSheet.getLastRow();
  const rowsCount = Math.max(0, lastRow - 1); // excluding header

  // If no data, just write basic summary
  summarySheet.clearContents();
  ensureHeaders_(summarySheet, ["metric", "value"]);

  const metrics = [];
  metrics.push(["Total rows in sheet", rowsCount]);
  metrics.push(["Imported (new)", importResult.inserted]);
  metrics.push(["Updated", importResult.updated]);
  metrics.push(["Skipped", importResult.skipped]);

  if (rowsCount === 0) {
    summarySheet.getRange(2, 1, metrics.length, 2).setValues(metrics);
    return;
  }

  // Read "type" column (D)
  const typeVals = dataSheet.getRange(2, 4, rowsCount, 1).getValues().flat().map(v => String(v || "").trim());
  const counts = new Map();
  typeVals.forEach(t => {
    const key = t || "(blank)";
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  metrics.push(["Distinct types", counts.size]);

  // Top 5 types
  const top = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).slice(0, 5);
  top.forEach(([k, v], i) => metrics.push([`Top type #${i + 1}`, `${k}: ${v}`]));

  summarySheet.getRange(2, 1, metrics.length, 2).setValues(metrics);
}

/* ------------------ Config + Helpers ------------------ */

function readConfig_(configSheet) {
  // expects key/value in A/B
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    // If empty, set defaults
    configSheet.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
    configSheet.getRange(2, 1, 3, 2).setValues([
      ["api_url", "https://rickandmortyapi.com/api/location"],
      ["max_pages", "3"],
      ["mode", "upsert"]
    ]);
  }

  const range = configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getValues();
  const cfg = {};
  for (let i = 1; i < range.length; i++) {
    const key = String(range[i][0] || "").trim();
    const val = String(range[i][1] || "").trim();
    if (key) cfg[key] = val;
  }
  return cfg;
}

function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, TAB_DATA);
  sheet.clearContents();
  ensureHeaders_(sheet, ["timestamp", "id", "name", "type", "dimension"]);
  SpreadsheetApp.getUi().alert(`Cleared "${TAB_DATA}".`);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const ok = headers.every((h, i) => String(firstRow[i] || "").toLowerCase() === h.toLowerCase());
  if (!ok) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function log_(logsSheet, level, message, context) {
  ensureHeaders_(logsSheet, ["timestamp", "level", "message", "context"]);
  logsSheet.appendRow([new Date(), level, message, JSON.stringify(context || {})]);
}