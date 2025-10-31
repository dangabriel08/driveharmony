/***** CONFIG *****/
const WATCH_SHEET = 'Watched Folders';
const DEFAULT_SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/xxx/yyy/zzz'; // fallback if sheet cell is blank
const ALERT_CHANNEL_LABEL = '#sec-drive-alerts'; // display label only

/***** ENTRYPOINT *****/
function checkUserSharesToSlackFromSheet() {
  const watchRows = getWatchList_(); // [{enabled, name, folderId, webhook}]
  if (!watchRows.length) return;

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

/***** CORE HELPERS *****/
// Parse sheet
function getWatchList_() {
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

// Drive v3 permissions → all user emails (post-change)
function listUserEmails_(fileId) {
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
