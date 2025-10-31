/***** CONFIG *****/
const WATCH_SHEET = 'Watched Folders';
const DEFAULT_SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/xxx/yyy/zzz'; // fallback if sheet cell is blank
const ALERT_CHANNEL_LABEL = '#sec-drive-alerts'; // display label only
const CONFIDENTIAL_SHEET = 'Confidential Folders';
const CONFIDENTIAL_HEADERS = ['Folder Name', 'Link', 'Folder ID'];
const WATCH_HEADERS = ['Enabled', 'Folder Name', 'Folder ID', 'Slack Webhook URL'];
const SHARED_DRIVES_SHEET = 'Shared Drives';
const SHARED_DRIVES_HEADERS = ['Drive Name', 'Drive ID', 'Created', 'Restrictions', 'Can Manage', 'Can Share', 'Can Add Members'];
const ACCESS_CHANGES_SHEET = 'Recent Access Changes';
const ACCESS_HEADERS = ['Watched Folder', 'Item', 'Link', 'Change Type', 'Granted/Removed For', 'Permission Role', 'Actor', 'When'];
const ACCESS_LOOKBACK_MINUTES = 60 * 24; // 24 hours

/***** ENTRYPOINT *****/
function checkUserSharesToSlackFromSheet() {
  const watchRows = getWatchList_(); // [{enabled, name, folderId, webhook}]
  if (!watchRows.length) return;

  try {
    requireDriveService_();
  } catch (err) {
    console.error(err && err.message ? err.message : err);
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const nowIso = new Date().toISOString();

  const alerts = [];

  for (const row of watchRows) {
    if (!row.enabled) continue;

    const sinceKey = `since:${row.folderId}`;
    const since = props.getProperty(sinceKey) || new Date(Date.now() - 15 * 60 * 1000).toISOString();

    // Query Drive Activity for this folder subtree
    const req = {
      ancestorName: `items/${row.folderId}`,
      filter: `detail:PERMISSION_CHANGE AND time >= "${since}"`
    };

    let res;
    try {
      res = DriveActivity.Activity.query(req) || {};
    } catch (e) {
      console.error(`DriveActivity query failed for ${row.folderId}`, e);
      continue;
    }

    const activities = res.activities || [];
    const folderAlerts = extractUserShareAlerts_(activities);

    // Enrich each alert with fileId, link, user emails via Drive v3
    for (const a of folderAlerts) {
      const fileId = resourceNameToId_(a.targetResName);
      const itemUrl = fileId ? `https://drive.google.com/open?id=${fileId}` : '';
      const grantedUserEmails = fileId ? listUserEmails_(fileId) : [];

      alerts.push({
        folderName: row.name,
        folderId: row.folderId,
        webhook: row.webhook || DEFAULT_SLACK_WEBHOOK_URL,
        targetName: a.targetTitle || '(item)',
        itemUrl,
        when: a.when,
        actor: a.actor,
        grantedUserEmails: dedupe_(grantedUserEmails)
      });
    }

    // advance this folder’s cursor
    props.setProperty(sinceKey, nowIso);
  }

  // Simple dedupe within this run (file + emails + minute bucket)
  const seen = new Set();
  for (const a of alerts) {
    const key = `${a.itemUrl}|${a.grantedUserEmails.join(',')}|${a.when?.slice(0,16)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    postToSlack_(a);
  }
}

/***** MENU & REPORTS *****/
function addDriveAlertsMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Drive Alerts')
    .addItem('List folders with "Confidential"', 'listConfidentialFolders')
    .addItem('Write recent permission changes', 'listRecentAccessChanges')
    .addItem('List Shared Drives', 'listAllSharedDrives')
    .addToUi();
}

function listConfidentialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert('This function must be run from a spreadsheet-bound script.');
    return;
  }

  try {
    requireDriveService_();
  } catch (err) {
    SpreadsheetApp.getUi().alert(err && err.message ? err.message : String(err));
    return;
  }

  const existing = ss.getSheetByName(CONFIDENTIAL_SHEET);
  const createNew = !existing;
  let sh = existing;
  if (!sh) {
    sh = ss.insertSheet(CONFIDENTIAL_SHEET);
  } else {
    sh.clear();
  }

  sh.getRange(1, 1, 1, CONFIDENTIAL_HEADERS.length).setValues([CONFIDENTIAL_HEADERS]);
  sh.setFrozenRows(1);

  const rows = [];
  const seen = new Set();
  const q = [
    "mimeType = 'application/vnd.google-apps.folder'",
    "trashed = false",
    "name contains 'Confidential'"
  ].join(' and ');

  let pageToken = null;
  do {
    const resp = Drive.Files.list({
      q,
      pageToken,
      pageSize: 200,
      fields: "nextPageToken, files(id,name,webViewLink)",
      corpora: 'allDrives',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true
    });
    const files = resp.files || [];
    files.forEach(f => {
      const name = f.name || '';
      if (!/confidential/i.test(name)) return;
      if (seen.has(f.id)) return;
      seen.add(f.id);
      const linkFormula = f.webViewLink ? `=HYPERLINK("${f.webViewLink}","Open")` : '';
      rows.push([name, linkFormula, f.id || '']);
    });
    pageToken = resp.nextPageToken || null;
  } while (pageToken);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, CONFIDENTIAL_HEADERS.length).setValues(rows);
    sh.autoResizeColumns(1, CONFIDENTIAL_HEADERS.length);
  } else if (!createNew) {
    const extraRows = sh.getMaxRows() - 1;
    if (extraRows > 0) sh.deleteRows(2, extraRows);
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert(rows.length ? `Found ${rows.length} folder(s) with "Confidential" in the name.` : 'No folders with "Confidential" found.');
}

function listAllSharedDrives() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert('This function must be run from a spreadsheet-bound script.');
    return;
  }

  try {
    requireDriveService_();
  } catch (err) {
    SpreadsheetApp.getUi().alert(err && err.message ? err.message : String(err));
    return;
  }

  const existing = ss.getSheetByName(SHARED_DRIVES_SHEET);
  const createNew = !existing;
  let sh = existing;
  if (!sh) {
    sh = ss.insertSheet(SHARED_DRIVES_SHEET);
  } else {
    sh.clear();
  }

  sh.getRange(1, 1, 1, SHARED_DRIVES_HEADERS.length).setValues([SHARED_DRIVES_HEADERS]);
  sh.setFrozenRows(1);

  const rows = [];
  let pageToken = null;
  do {
    const resp = Drive.Drives.list({
      pageSize: 100,
      pageToken,
      useDomainAdminAccess: true,
      fields: 'nextPageToken, drives(id,name,createdTime,restrictions(adminManagedRestrictions,copyRequiresWriterPermission,domainUsersOnly,sharingFoldersRequiresOrganizerPermission),capabilities(canManageMembers,canShare,canAddChildren))'
    });
    const drives = resp.drives || [];
    drives.forEach(d => {
      const created = formatTimestampLocal_(d.createdTime);
      const restrictions = summarizeRestrictions_(d.restrictions || {});
      const canManage = extractCapability_(d.capabilities, 'canManageMembers');
      const canShare = extractCapability_(d.capabilities, 'canShare');
      const canAdd = extractCapability_(d.capabilities, 'canAddChildren');
      rows.push([
        d.name || '(unnamed)',
        d.id || '',
        created,
        restrictions,
        canManage,
        canShare,
        canAdd
      ]);
    });
    pageToken = resp.nextPageToken || null;
  } while (pageToken);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, SHARED_DRIVES_HEADERS.length).setValues(rows);
    sh.autoResizeColumns(1, SHARED_DRIVES_HEADERS.length);
  } else if (!createNew) {
    const extraRows = sh.getMaxRows() - 1;
    if (extraRows > 0) sh.deleteRows(2, extraRows);
  }

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sh.getRange(1, 1).setNote(`Generated ${stamp}`);

  SpreadsheetApp.getUi().alert(rows.length ? `Listed ${rows.length} shared drive(s).` : 'No shared drives found or accessible.');
}

function listRecentAccessChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert('This function must be run from a spreadsheet-bound script.');
    return;
  }

  const watchRows = getWatchList_().filter(row => row.enabled);
  if (!watchRows.length) {
    SpreadsheetApp.getUi().alert('No enabled folders in the watch list. Add at least one folder to "Watched Folders".');
    return;
  }

  const existing = ss.getSheetByName(ACCESS_CHANGES_SHEET);
  const createNew = !existing;
  let sh = existing;
  if (!sh) {
    sh = ss.insertSheet(ACCESS_CHANGES_SHEET);
  } else {
    sh.clear();
  }

  sh.getRange(1, 1, 1, ACCESS_HEADERS.length).setValues([ACCESS_HEADERS]);
  sh.setFrozenRows(1);

  const sinceIso = new Date(Date.now() - ACCESS_LOOKBACK_MINUTES * 60 * 1000).toISOString();
  const events = [];

  for (const row of watchRows) {
    const req = {
      ancestorName: `items/${row.folderId}`,
      filter: `detail:PERMISSION_CHANGE AND time >= "${sinceIso}"`
    };
    let res;
    try {
      res = DriveActivity.Activity.query(req) || {};
    } catch (err) {
      console.error(`DriveActivity query failed for ${row.folderId}`, err);
      continue;
    }
    const activities = res.activities || [];
    const perms = extractPermissionChangeEvents_(activities);
    perms.forEach(evt => {
      const fileId = resourceNameToId_(evt.targetResName);
      events.push({
        watchName: row.name || '(watch)',
        itemName: evt.targetTitle || '(item)',
        itemUrl: fileId ? `https://drive.google.com/open?id=${fileId}` : '',
        changeType: evt.changeType,
        entity: evt.entity,
        role: evt.role || '',
        actor: evt.actor,
        when: evt.when
      });
    });
  }

  events.sort((a, b) => new Date(b.when).getTime() - new Date(a.when).getTime());

  if (events.length) {
    const rows = events.map(evt => {
      const link = evt.itemUrl ? `=HYPERLINK("${evt.itemUrl}","Open")` : '';
      const whenText = formatTimestampLocal_(evt.when);
      return [
        evt.watchName,
        evt.itemName,
        link,
        evt.changeType,
        evt.entity,
        evt.role,
        evt.actor,
        whenText
      ];
    });
    sh.getRange(2, 1, rows.length, ACCESS_HEADERS.length).setValues(rows);
    sh.autoResizeColumns(1, ACCESS_HEADERS.length);
  } else if (!createNew) {
    const extraRows = sh.getMaxRows() - 1;
    if (extraRows > 0) sh.deleteRows(2, extraRows);
  }

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sh.getRange(1, 1).setNote(`Generated ${stamp}`);

  SpreadsheetApp.getUi().alert(events.length ? `Wrote ${events.length} recent permission change(s).` : 'No permission changes found in the last 24 hours.');
}

/***** CORE HELPERS *****/
// Parse sheet
function getWatchList_() {
  ensureWatchSheet_();
  const sh = SpreadsheetApp.getActive().getSheetByName(WATCH_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${WATCH_SHEET}`);

  const [hdr, ...rows] = sh.getDataRange().getValues();
  const idx = headerIndex_(hdr, ['Enabled','Folder Name','Folder ID','Slack Webhook URL']);

  const out = [];
  for (const r of rows) {
    const enabledVal = String((r[idx.Enabled] ?? '')).trim().toUpperCase();
    const enabled = enabledVal === 'Y' || enabledVal === 'YES' || enabledVal === 'TRUE';
    const name = String(r[idx['Folder Name']] ?? '').trim();
    const folderId = String(r[idx['Folder ID']] ?? '').trim();
    const webhook = idx['Slack Webhook URL'] != null ? String(r[idx['Slack Webhook URL']] ?? '').trim() : '';

    if (!folderId) continue;
    out.push({ enabled, name, folderId, webhook });
  }
  return out;
}

function ensureWatchSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  let sh = ss.getSheetByName(WATCH_SHEET);
  if (!sh) {
    sh = ss.insertSheet(WATCH_SHEET);
    sh.getRange(1, 1, 1, WATCH_HEADERS.length).setValues([WATCH_HEADERS]);
    sh.setFrozenRows(1);
    return;
  }

  const neededLength = WATCH_HEADERS.length;
  const currentCols = Math.max(sh.getLastColumn(), neededLength);
  const header = sh.getRange(1, 1, 1, currentCols).getValues()[0];
  let mismatch = false;
  WATCH_HEADERS.forEach((expected, idx) => {
    if (header[idx] !== expected) mismatch = true;
  });
  if (mismatch) {
    sh.getRange(1, 1, 1, neededLength).setValues([WATCH_HEADERS]);
  }
  sh.setFrozenRows(1);
}

function headerIndex_(hdr, needed) {
  const map = {};
  needed.forEach(h => map[h] = hdr.indexOf(h));
  // tolerate missing optional columns
  if (map['Enabled'] === -1 || map['Folder Name'] === -1 || map['Folder ID'] === -1) {
    throw new Error(`Header row must include: Enabled, Folder Name, Folder ID (Slack Webhook URL optional)`);
  }
  return map;
}

// Extract user-share events from Drive Activity response
function extractUserShareAlerts_(activities) {
  const out = [];
  (activities || []).forEach(a => {
    const targets = a.targets || [];
    const target = (targets.find(t => t.driveItem) || {}).driveItem || {};
    const targetTitle = target.title || '(item)';
    const targetResName = target.name; // items/FILE_ID

    (a.actions || []).forEach(ac => {
      const permChange = ac.detail && ac.detail.permissionChange;
      if (!permChange) return;

      const when = ac.timestamp || a.timestamp || new Date().toISOString();
      const actor = resolveActor_(ac.actor);

      (permChange.addedPermissions || []).forEach(p => {
        const entityType = detectPermissionEntity_(p);
        if (entityType === 'USER') {
          out.push({ targetTitle, targetResName, when, actor });
        }
      });
    });
  });
  return out;
}

function extractPermissionChangeEvents_(activities) {
  const out = [];
  (activities || []).forEach(a => {
    const targets = a.targets || [];
    const target = (targets.find(t => t.driveItem) || {}).driveItem || {};
    const targetTitle = target.title || '(item)';
    const targetResName = target.name;

    (a.actions || []).forEach(ac => {
      const permChange = ac.detail && ac.detail.permissionChange;
      if (!permChange) return;

      const when = ac.timestamp || a.timestamp || new Date().toISOString();
      const actor = resolveActor_(ac.actor);

      (permChange.addedPermissions || []).forEach(p => {
        out.push({
          targetTitle,
          targetResName,
          when,
          actor,
          changeType: 'Added',
          entity: describePermissionEntity_(p),
          role: (p && p.role) || ''
        });
      });

      (permChange.removedPermissions || []).forEach(p => {
        out.push({
          targetTitle,
          targetResName,
          when,
          actor,
          changeType: 'Removed',
          entity: describePermissionEntity_(p),
          role: (p && p.role) || ''
        });
      });
    });
  });
  return out;
}

// Drive v3 permissions → all user emails (post-change)
function listUserEmails_(fileId) {
  requireDriveService_();
  try {
    const permList = Drive.Permissions.list(fileId, { supportsAllDrives: true }) || {};
    return (permList.permissions || [])
      .filter(p => p.type === 'user')
      .map(p => p.emailAddress)
      .filter(Boolean);
  } catch (e) {
    console.warn('Permissions.list failed for', fileId, e.message);
    return [];
  }
}

function resourceNameToId_(resourceName) {
  if (!resourceName) return '';
  const parts = resourceName.split('/');
  return parts.length === 2 ? parts[1] : '';
}

function detectPermissionEntity_(perm) {
  if (perm && perm.user && perm.user.knownUser) return 'USER';
  if (perm && perm.group) return 'GROUP';
  if (perm && perm.domain) return 'DOMAIN';
  if (perm && perm.anyone) return 'ANYONE';
  return 'UNKNOWN';
}

function resolveActor_(actor) {
  try {
    if (actor && actor.user && actor.user.knownUser && actor.user.knownUser.personName) {
      return `actor: ${actor.user.knownUser.personName}`; // people/xxx (not email)
    }
    if (actor && actor.user && actor.user.deletedUser) return '(deleted user)';
  } catch (e) {}
  return '(actor unavailable)';
}

function dedupe_(arr) {
  return Array.from(new Set(arr || [])).sort();
}

function describePermissionEntity_(perm) {
  if (!perm) return 'Unknown entity';
  if (perm.user && perm.user.knownUser) {
    const person = perm.user.knownUser.personName || perm.user.knownUser.userAccount || '';
    return person ? `User: ${person}` : 'User';
  }
  if (perm.group && perm.group.email) {
    return `Group: ${perm.group.email}`;
  }
  if (perm.domain && perm.domain.domain) {
    return `Domain: ${perm.domain.domain}`;
  }
  if (perm.anyone) {
    return 'Anyone with link';
  }
  if (perm.link) {
    return 'Shared via link';
  }
  if (perm.permissionId) {
    return `Permission ${perm.permissionId}`;
  }
  return 'Unknown entity';
}

function formatTimestampLocal_(isoString) {
  if (!isoString) return '';
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return isoString;
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm:ss');
}

function requireDriveService_() {
  if (typeof Drive === 'undefined' || !Drive || !Drive.Files) {
    throw new Error(
      'Drive advanced service is disabled. In Apps Script, open Services (✚ icon), enable Drive API, and authorize the script.'
    );
  }
}

function summarizeRestrictions_(restrictions) {
  const bits = [];
  if (!restrictions) return 'None';
  if (restrictions.adminManagedRestrictions) bits.push('Admin managed');
  if (restrictions.copyRequiresWriterPermission) bits.push('Copy needs writer');
  if (restrictions.domainUsersOnly) bits.push('Domain only');
  if (restrictions.sharingFoldersRequiresOrganizerPermission) bits.push('Sharing needs organizer');
  return bits.length ? bits.join(', ') : 'None';
}

function extractCapability_(caps, key) {
  if (!caps || typeof caps[key] === 'undefined') return 'Unknown';
  return caps[key] ? 'Yes' : 'No';
}

// Slack poster (supports per-row webhook)
function postToSlack_(evt) {
  const title = `User-share detected (not a group)`;
  const linkLine = evt.itemUrl ? `<${evt.itemUrl}|Open in Drive>` : '(no link)';
  const emails = evt.grantedUserEmails.length ? evt.grantedUserEmails.join(', ') : '(email not exposed yet)';
  const ts = new Date(evt.when).toLocaleString();

  const payload = {
    text: `${title} in ${ALERT_CHANNEL_LABEL}`,
    blocks: [
      { type: 'header', text: { type: 'plain_text', text: title, emoji: true } },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text:
`*Folder:* ${evt.folderName || '(watch)'}
*Item:* ${evt.targetName}
*Link:* ${linkLine}
*Added to (user emails):* ${emails}
*By:* ${evt.actor}
*When:* ${ts}`
        }
      }
    ]
  };

  const params = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = evt.webhook || DEFAULT_SLACK_WEBHOOK_URL;
  try {
    const res = UrlFetchApp.fetch(url, params);
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
      console.error('Slack post failed', res.getResponseCode(), res.getContentText());
    }
  } catch (e) {
    console.error('Slack post error', e);
  }
}

function onOpen() {
  if (typeof ensureSheetStructure_ === 'function') {
    try {
      ensureSheetStructure_();
    } catch (err) {
      console.warn('ensureSheetStructure_ failed:', err && err.message ? err.message : err);
    }
  }
  try {
    ensureWatchSheet_();
  } catch (err) {
    console.warn('ensureWatchSheet_ failed:', err && err.message ? err.message : err);
  }
  if (typeof addGroupToolsMenu_ === 'function') {
    try {
      addGroupToolsMenu_();
    } catch (err) {
      console.warn('addGroupToolsMenu_ failed:', err && err.message ? err.message : err);
    }
  }
  addDriveAlertsMenu_();
}
