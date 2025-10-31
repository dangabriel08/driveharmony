/**************************************
 * Group List ( Automated ) — main list + per-group folder tabs
 * BATCHED "ALL groups" to avoid timeouts.
 *
 * Main sheet columns (order):
 *   Group Name, Email, Description, Member count, Members, Group ID
 *
 * Menu:
 *   - Refresh Group List
 *   - Build folder tab for selected row
 *   - Build folder tabs for ALL groups (queued)
 *   - Cancel running build
 *   - Recreate Daily Trigger (3 AM)
 *
 * Per-group tabs:
 *   - One sheet per group (named by group name, sanitized)
 *   - Re-run clears & updates existing tab (no duplicates)
 *   - Columns: Path, Folder name, Link, Folder ID, Drive
 *
 * Advanced Services (enable in Apps Script):
 *   - AdminDirectory (Admin SDK)
 *   - Drive (Drive v3)
 *
 * Suggested manifest scopes (appsscript.json):
 *   - https://www.googleapis.com/auth/spreadsheets
 *   - https://www.googleapis.com/auth/admin.directory.group.readonly
 *   - https://www.googleapis.com/auth/admin.directory.group.member.readonly
 *   - https://www.googleapis.com/auth/drive.metadata.readonly
 *   - (Optional; used by trigger helpers) https://www.googleapis.com/auth/script.scriptapp
 **************************************/

const SHEET_NAME = 'Group List ( Automated )';
const DAILY_TRIGGER_HOUR = 3; // 3 AM

// === HEADERS ===
const headers = [
  'Group Name',
  'Email',
  'Description',
  'Member count',
  'Members',
  'Group ID'
];

// === Column indices (1-based) ===
const COL = {
  NAME: 1,
  EMAIL: 2,
  DESCRIPTION: 3,
  MEMBER_COUNT: 4,
  MEMBERS: 5,
  GROUP_ID: 6,
};

// === Queueing constants (batched “ALL groups”) ===
const BUILD_BATCH_SIZE = 3; // process N groups per invocation
const BUILD_QUEUE_KEY   = 'FOLDER_TABS_BUILD_QUEUE';  // JSON array of items
const BUILD_INDEX_KEY   = 'FOLDER_TABS_BUILD_INDEX';  // current index (0-based)
const BUILD_STATE_KEY   = 'FOLDER_TABS_BUILD_STATE';  // "RUNNING" | "IDLE"
const BUILD_TRIG_TAG    = 'continueBuildAllGroupTabs_'; // handler name

/* =========================
   MENU
   ========================= */
function onOpen() {
  ensureSheetStructure_();
  SpreadsheetApp.getUi()
    .createMenu('Group Tools')
    .addItem('Refresh Group List', 'refreshGroups')
    .addSeparator()
    .addItem('Build folder tab for selected row', 'buildFolderTabForSelectedRow')
    .addItem('Build folder tabs for ALL groups (queued)', 'startBuildFolderTabsForAllGroups')
    .addItem('Cancel running build', 'cancelBuildAllGroupTabs')
    .addSeparator()
    .addItem('Recreate Daily Trigger (3 AM)', 'createDailyTrigger_')
    .addToUi();
}

/* =========================
   MAIN LIST
   ========================= */
function refreshGroups() {
  writeGroupsToSheet_();
}

function createDailyTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshGroups')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshGroups')
    .timeBased().everyDays(1).atHour(DAILY_TRIGGER_HOUR).create();
}

function ensureSheetStructure_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }
  const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const mismatch = headers.some((h, i) => existing[i] !== h);
  if (mismatch) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function writeGroupsToSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  // CLEAR then headers
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  const rows = [];
  let pageToken = null;

  do {
    const resp = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 200,
      pageToken,
      fields: 'nextPageToken,groups(id,email,name,description,directMembersCount)'
    });
    const groups = resp.groups || [];
    for (const g of groups) {
      const email = g.email || '';
      const name = g.name || '';
      const desc = g.description || '';
      const { directCount, memberList } = getGroupMembers_(email);

      rows.push([
        name,              // Group Name
        email,             // Email
        desc,              // Description
        directCount,       // Member count
        memberList,        // Members (newline-separated)
        g.id || ''         // Group ID
      ]);
    }
    pageToken = resp.nextPageToken || null;
  } while (pageToken);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(2, COL.DESCRIPTION, rows.length, 1).setWrap(true);
    sh.getRange(2, COL.MEMBERS, rows.length, 1).setWrap(true);
    sh.autoResizeColumns(1, headers.length);
  }

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sh.getRange(1, 1).setNote(`Last sync: ${stamp} (${tz})`);
}

/* Members helper */
function getGroupMembers_(groupEmail) {
  const members = [];
  let pageToken = null, count = 0;
  try {
    do {
      const resp = AdminDirectory.Members.list(groupEmail, {
        pageToken,
        maxResults: 200,
        fields: 'nextPageToken,members(email,role)'
      });
      const arr = resp.members || [];
      for (const m of arr) {
        count++;
        const role = m.role ? ` (${capitalize_(m.role)})` : '';
        members.push(`${m.email}${role}`);
      }
      pageToken = resp.nextPageToken || null;
    } while (pageToken);
  } catch (err) {
    members.push(`(Unable to list members: ${err.message})`);
  }
  return { directCount: count, memberList: members.join('\n') };
}

function capitalize_(s) {
  return (s || '').toString().charAt(0).toUpperCase() + (s || '').toString().slice(1);
}

/* =========================
   PER-GROUP TABS
   ========================= */
function buildFolderTabForSelectedRow() {
  const ss = SpreadsheetApp.getActive();
  const main = ss.getSheetByName(SHEET_NAME);
  const r = main.getActiveCell().getRow();
  if (r < 2) {
    SpreadsheetApp.getUi().alert('Select a data row first.');
    return;
  }
  const name = String(main.getRange(r, COL.NAME).getValue() || '');
  const email = String(main.getRange(r, COL.EMAIL).getValue() || '');
  const id = String(main.getRange(r, COL.GROUP_ID).getValue() || '');
  const key = email || id || name;
  if (!key) {
    SpreadsheetApp.getUi().alert('Row has no Email or Group ID.');
    return;
  }
  // Prefer not to burn Admin API calls if email is available
  const groupDisplayName = name || email || key;
  const groupEmail = email || resolveGroupIdentity_(key).email;

  buildOneGroupFolderTab_(groupDisplayName, groupEmail);
  SpreadsheetApp.getUi().alert(`Built tab for: ${groupDisplayName}`);
}

/**
 * Start queued (batched) build for ALL groups.
 * Creates a queue in Script Properties and schedules the worker.
 */
function startBuildFolderTabsForAllGroups() {
  const ss = SpreadsheetApp.getActive();
  const main = ss.getSheetByName(SHEET_NAME);
  const last = main.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('No groups to process. Refresh the main list first.');
    return;
  }

  // Build a queue from the main list (prefer email if present)
  const values = main.getRange(2, 1, last - 1, headers.length).getValues();
  const items = values.map(row => ({
    name: String(row[COL.NAME - 1] || ''),
    email: String(row[COL.EMAIL - 1] || ''),
    id: String(row[COL.GROUP_ID - 1] || '')
  }));

  const props = PropertiesService.getScriptProperties();
  props.setProperty(BUILD_QUEUE_KEY, JSON.stringify(items));
  props.setProperty(BUILD_INDEX_KEY, '0');
  props.setProperty(BUILD_STATE_KEY, 'RUNNING');

  // Ensure only one worker trigger exists
  deleteTriggersFor_(BUILD_TRIG_TAG);
  ScriptApp.newTrigger(BUILD_TRIG_TAG).timeBased().after(1000).create(); // kick off in ~1s

  SpreadsheetApp.getUi().alert(`Queued ${items.length} group(s). Building in batches of ${BUILD_BATCH_SIZE}…`);
}

/** Worker: process next batch and reschedule if more remain */
function continueBuildAllGroupTabs_() {
  const props = PropertiesService.getScriptProperties();
  const state = props.getProperty(BUILD_STATE_KEY);
  if (state !== 'RUNNING') {
    // No-op if canceled or not started
    deleteTriggersFor_(BUILD_TRIG_TAG);
    return;
  }

  const raw = props.getProperty(BUILD_QUEUE_KEY);
  const idxStr = props.getProperty(BUILD_INDEX_KEY);
  if (!raw || idxStr === null) {
    // Nothing to do
    props.deleteProperty(BUILD_STATE_KEY);
    deleteTriggersFor_(BUILD_TRIG_TAG);
    return;
  }

  const queue = JSON.parse(raw);
  let i = Number(idxStr) || 0;
  const end = Math.min(queue.length, i + BUILD_BATCH_SIZE);

  for (; i < end; i++) {
    const item = queue[i];
    try {
      const display = item.name || item.email || item.id;
      const email = item.email || resolveGroupIdentity_(item.id || item.name).email;
      buildOneGroupFolderTab_(display, email);
      // yield a tiny bit (helps UI/quotas)
      Utilities.sleep(200);
    } catch (e) {
      Logger.log(`Error on ${item && (item.email || item.name || item.id)}: ${e && e.message ? e.message : e}`);
    }
  }

  props.setProperty(BUILD_INDEX_KEY, String(i));

  if (i < queue.length) {
    // More to do — schedule another one-shot run
    deleteTriggersFor_(BUILD_TRIG_TAG);
    ScriptApp.newTrigger(BUILD_TRIG_TAG).timeBased().after(1500).create();
  } else {
    // Done
    props.deleteProperty(BUILD_QUEUE_KEY);
    props.deleteProperty(BUILD_INDEX_KEY);
    props.setProperty(BUILD_STATE_KEY, 'IDLE');
    deleteTriggersFor_(BUILD_TRIG_TAG);
  }
}

/** Cancel current queued build */
function cancelBuildAllGroupTabs() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(BUILD_STATE_KEY, 'IDLE');
  deleteTriggersFor_(BUILD_TRIG_TAG);
  SpreadsheetApp.getUi().alert('Canceled the running build.');
}

/* Trigger helper */
function deleteTriggersFor_(handlerName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handlerName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * Create or reuse a sheet named after the group and list folders in a tree-like format.
 * If the tab already exists, it is CLEARED and updated (no duplicates).
 * Columns: Path, Folder name, Link, Folder ID, Drive
 */
function buildOneGroupFolderTab_(groupName, groupEmail) {
  const ss = SpreadsheetApp.getActive();
  const tabName = sanitizeSheetName_(groupName).slice(0, 80); // safe length

  // Reuse existing tab if present; otherwise create it.
  let sh = ss.getSheetByName(tabName);
  if (!sh) {
    sh = ss.insertSheet(tabName);
  } else {
    sh.clear(); // clear & update
  }

  const HEAD = ['Path', 'Folder name', 'Link', 'Folder ID', 'Drive'];
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);

  // Get all directly shared folders (with parents info)
  const folders = listFoldersDirectlySharedToGroup_(groupEmail);

  // Cache: id -> {id,name,parents,driveId}
  const cache = new Map();
  folders.forEach(f => cache.set(f.id, { id: f.id, name: f.name, parents: f.parents || [], driveId: f.driveId || null }));

  const rows = folders.map(f => {
    const { pathParts, depth, driveName } = computeFullPath_(f, cache);
    const path = pathParts.join(' / ');
    const indentPrefix = ' '.repeat(depth * 2) + (depth > 0 ? '• ' : '');
    const nameCell = indentPrefix + f.name;
    const linkFormula = `=HYPERLINK("${f.link}","Open")`;
    return [path, nameCell, linkFormula, f.id, driveName || ''];
  });

  // Sort by Path then Name
  rows.sort((a, b) => (a[0] || '').localeCompare(b[0] || '') || (a[1] || '').localeCompare(b[1] || ''));

  if (rows.length) {
    sh.getRange(2, 1, rows.length, HEAD.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 2).setWrap(true);
    sh.autoResizeColumns(1, HEAD.length);
  }

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sh.getRange(1, 1).setNote(`Group: ${groupName} <${groupEmail}> — generated ${stamp}`);
}

/** Build full path by walking parents (caches Drive lookups) */
function computeFullPath_(file, cache) {
  const pathIds = [];
  let cursor = file;

  while (cursor) {
    pathIds.push(cursor.id);
    const parents = cursor.parents || [];
    if (!parents.length) break;
    const parentId = parents[0];
    let parent = cache.get(parentId);
    if (!parent) {
      try {
        const p = Drive.Files.get(parentId, { fields: 'id,name,parents,driveId' });
        parent = { id: p.id, name: p.name, parents: p.parents || [], driveId: p.driveId || null };
        cache.set(parent.id, parent);
      } catch (_) {
        parent = null; // stop if inaccessible
      }
    }
    cursor = parent;
  }

  const parts = pathIds.reverse().map(id => (cache.get(id) ? cache.get(id).name : id));
  const depth = Math.max(0, parts.length - 1);

  let driveName = '';
  const withDrive = cache.get(file.id);
  const driveId = withDrive && withDrive.driveId ? withDrive.driveId : null;
  if (driveId) {
    try {
      const d = Drive.Drives.get(driveId, { fields: 'id,name' });
      driveName = d && d.name ? d.name : '';
    } catch (_) {}
  }

  return { pathParts: parts, depth, driveName };
}

/* ===== Group/Drive helpers ===== */
function resolveGroupIdentity_(groupKey) {
  try {
    const g = AdminDirectory.Groups.get(groupKey);
    return { email: g.email, id: g.id, name: g.name || g.email };
  } catch (err) {
    if (/@/.test(groupKey)) return { email: groupKey, id: '', name: groupKey };
    throw new Error(`Unable to resolve group: ${groupKey} (${err.message})`);
  }
}

function listFoldersDirectlySharedToGroup_(groupEmail) {
  const results = [];
  let pageToken = null;

  const q = [
    "mimeType = 'application/vnd.google-apps.folder'",
    "trashed = false",
    `('${escapeForQuery_(groupEmail)}' in readers or '${escapeForQuery_(groupEmail)}' in writers or '${escapeForQuery_(groupEmail)}' in owners)`
  ].join(' and ');

  do {
    const resp = Drive.Files.list({
      q,
      pageSize: 200,
      pageToken,
      fields: "nextPageToken, files(id, name, webViewLink, driveId, parents)",
      corpora: 'allDrives',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    (resp.files || []).forEach(f => results.push({
      id: f.id,
      name: f.name,
      link: f.webViewLink,
      driveId: f.driveId || null,
      parents: f.parents || []
    }));
    pageToken = resp.nextPageToken || null;
  } while (pageToken);

  return results;
}

function escapeForQuery_(s) { return String(s).replace(/'/g, "\\'"); }
