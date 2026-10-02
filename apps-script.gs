/** @OnlyCurrentDoc */
// Vitamin VVV landing page -> Google Sheet "Vitamin VVV 落地页回答", tab "Responses".
// First install: Extensions > Apps Script, paste, Deploy > New deployment > Web app,
// Execute as: Me, Who has access: Anyone. Put the /exec URL into SUBMIT_URL in the page.
// To update the code later: Deploy > Manage deployments > pencil > Version: New version > Deploy.
// That keeps the same /exec URL ("New deployment" would make a new one).

const VERSION = "2";
const TAB = "Responses";
const COLS = ["Time", "Response id", "Name", "Wears", "Rank 1", "Rank 2", "Rank 3", "Rank 4", "Rank 5", "Rank 6",
  "Shown order", "Round 2 style", "Left side", "Fashion places", "Practical places", "Prefers", "Own places",
  "Email", "Address", "Device", "Test", "Seq"];
const ID_RE = /^[A-Za-z0-9-]{8,40}$/;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  const raw = (e && e.postData && e.postData.contents) || "";
  let d;
  try { d = JSON.parse(raw); } catch (err) { return out_("bad request"); }
  if (raw.length > 10000 || !d || (d.type !== "response" && d.type !== "email") || typeof d.id !== "string" || !ID_RE.test(d.id)) return out_("bad request");
  if (d.type === "email" && !EMAIL_RE.test(String(d.email || "").trim())) return out_("bad request");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_(keepUnsaved_(raw, "busy") ? "ok" : "error");
  try {
    save_(d);
    SpreadsheetApp.flush();
    return out_("ok");
  } catch (err) {
    console.error(err, raw.slice(0, 2000));
    return out_(keepUnsaved_(raw, String(err)) ? "ok" : "error");
  } finally {
    lock.releaseLock();
  }
}

// Opening the /exec URL in a browser shows this, so it's easy to check which code is live.
function doGet() {
  return out_("Vitamin VVV form endpoint is live (v" + VERSION + ").");
}

function save_(d) {
  const sh = sheet_();
  const col = columns_(sh);
  const id = String(d.id);
  const seq = Number(d.seq) || 0;
  const found = findRow_(sh, col, id);
  const answers = answersFor_(d, seq);
  if (!found.row) {
    const values = Object.assign({ "Time": new Date(), "Response id": id }, answers);
    if (d.type === "email") Object.assign(values, contact_(d));
    write_(sh, col, found.next, values);
    return;
  }
  // A later answer from the same visitor (after Undo or a change in round 2) replaces the earlier one.
  const stored = Number(sh.getRange(found.row, col["Seq"]).getValue()) || 0;
  const values = {};
  if (d.type === "response" && seq > stored) Object.assign(values, answers);
  if (d.type === "email") Object.assign(values, contact_(d));
  write_(sh, col, found.row, values);
}

function answersFor_(d, seq) {
  const rank = (Array.isArray(d.ranking) ? d.ranking : []).slice(0, 6);
  while (rank.length < 6) rank.push("");
  const r2 = d.round2 || {};
  const scenes = r2.scenes || {};
  const list = a => clean_((Array.isArray(a) ? a : []).filter(Boolean).join(", "));
  const v = {
    "Name": clean_(d.name), "Wears": clean_(d.gender),
    "Shown order": list(d.shownOrder), "Round 2 style": clean_(r2.style),
    "Left side": clean_((r2.sides || [])[0]),
    "Fashion places": list(scenes.fashion), "Practical places": list(scenes.practical),
    "Prefers": clean_(r2.prefer), "Own places": list(r2.others),
    "Device": clean_(d.device), "Test": d.test ? "yes" : "", "Seq": seq
  };
  rank.forEach((r, i) => { v["Rank " + (i + 1)] = clean_(r); });
  return v;
}

function contact_(d) {
  return { "Email": clean_(String(d.email || "").trim()), "Address": clean_(d.address) };
}

// The "Responses" tab. On first run the existing first tab is renamed if it holds no answers yet.
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    const first = ss.getSheets()[0];
    sh = first.getLastRow() <= 1 ? first.setName(TAB) : ss.insertSheet(TAB, 0);
  }
  return sh;
}

// Column number for each name in COLS, read from the header. Never overwrites a filled row 1:
// missing columns are added at the end, and if row 1 isn't a header a new header row is inserted above it.
function columns_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.setFrozenRows(1);
  } else {
    let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (head.indexOf("Response id") === -1) {
      sh.insertRowBefore(1);
      sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
      sh.setFrozenRows(1);
      head = COLS.slice();
    }
    if (sh.getFrozenRows() === 0) sh.setFrozenRows(1);
    let width = head.length;
    while (width > 0 && head[width - 1] === "") width--;
    const missing = COLS.filter(c => head.indexOf(c) === -1);
    if (missing.length) {
      if (sh.getMaxColumns() < width + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), width + missing.length - sh.getMaxColumns());
      sh.getRange(1, width + 1, 1, missing.length).setValues([missing]);
      head = head.slice(0, width).concat(missing);
    }
    const col = {};
    COLS.forEach(c => { col[c] = head.indexOf(c) + 1; });
    return col;
  }
  const col = {};
  COLS.forEach((c, i) => { col[c] = i + 1; });
  return col;
}

// Finds the visitor's row by id, and the first row after the last id (so helper columns
// such as gift checkboxes further down don't push new rows to the bottom of the sheet).
function findRow_(sh, col, id) {
  const last = sh.getLastRow();
  if (last < 2) return { row: 0, next: 2 };
  const ids = sh.getRange(2, col["Response id"], last - 1, 1).getValues();
  let next = 2;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === "") continue;
    if (next === 2) next = i + 3;
    if (ids[i][0] === id) return { row: i + 2, next: next };
  }
  return { row: 0, next: next };
}

// Writes only our own cells, one range per run of adjacent columns, so other columns stay as they are.
function write_(sh, col, row, values) {
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows() + 100);
  const cells = Object.keys(values).map(k => [col[k], values[k]]).sort((a, b) => a[0] - b[0]);
  for (let i = 0; i < cells.length;) {
    let j = i + 1;
    while (j < cells.length && cells[j][0] === cells[j - 1][0] + 1) j++;
    sh.getRange(row, cells[i][0], 1, j - i).setValues([cells.slice(i, j).map(c => c[1])]);
    i = j;
  }
}

// Anything that couldn't be saved goes to an "Unsaved" tab as raw text, so no answer is lost.
function keepUnsaved_(raw, why) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Unsaved") || ss.insertSheet("Unsaved");
    sh.appendRow([new Date(), why.slice(0, 300), "'" + raw.slice(0, 40000)]);
    return true;
  } catch (err) {
    console.error("keepUnsaved_", err, raw.slice(0, 2000));
    return false;
  }
}

function out_(text) {
  return ContentService.createTextOutput(text);
}

// Plain text only, capped. A value that could run as a formula keeps a visible ' in front,
// so it stays text in Sheets and in a CSV export.
function clean_(v) {
  const s = v == null ? "" : String(v).slice(0, 300);
  return /^[=+\-@\t\r]/.test(s) ? "''" + s : s;
}
