# Developement

## Enable github pages to serve index.html

From github CLI Create a new repo and enable pages
```
gh repo create bjacques/40k-rpg --public --source=. --remote=origin --push
gh api -X POST repos/bjacques/40k-rpg/pages -f "source[branch]=main" -f "source[path]=/" --jq '.html_url'
```


## Google Sheets save (preferred) with localStorage and CSV

The page always auto-saves to the browser (localStorage). Saving to the GM's Google Sheet is the
preferred way to keep a character: once a player pastes a **group passphrase** into the Save bar at the
bottom of the page (press **More**), every save is also copied to the Sheet, and the **Save** / **Restore from sheet**
buttons are there for manual use. Download/Load CSV (under **More**) is the secondary fallback file backup.

### GM setup (one time, free)
1. Create a Google Sheet.
2. Extensions > Apps Script, replace the contents with the script below.
3. Project Settings > Script Properties > add `TABLE_CODE` = a shared passphrase (no `:` needed, any text).
4. Deploy > New deployment > type **Web app**, Execute as **Me**, Who has access **Anyone**. Authorise when asked (the "app hasn't been verified" warning is expected for your own script: Advanced > Go to project > Allow).
5. Copy the deployment ID (the part of the `/exec` URL between `/s/` and `/exec`) and set it as `SHEET_ID` in `index.html`. It is not secret; the passphrase is what protects the Sheet. Changing deployments gives a new ID.
6. Send each player the passphrase (`TABLE_CODE`) privately. Do not commit it.
7. After editing the script later: Deploy > Manage deployments > Edit > New version.

### Code.gs
```js
/** @OnlyCurrentDoc */
const SHEET_NAME = 'Characters';
const HEADERS = ['Key', 'Player', 'Character', 'Chapter', 'Speciality', 'Rank', 'UpdatedAt', 'CSV'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  return sh;
}
function ok_(code) {
  const t = PropertiesService.getScriptProperties().getProperty('TABLE_CODE');
  return !!t && code === t;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false }); }
  if (!ok_(d.code) || !d.player || typeof d.csv !== 'string' || d.csv.length > 45000) return json_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const key = (d.player + '|' + (d.character || '')).toLowerCase();
    const row = [key, d.player, d.character || '', d.chapter || '', d.spec || '', d.rank || '', new Date().toISOString(), d.csv];
    const n = sh.getLastRow() - 1;
    const keys = n > 0 ? sh.getRange(2, 1, n, 1).getValues().flat() : [];
    const i = keys.indexOf(key);
    const r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
    const range = sh.getRange(r, 1, 1, 8);
    range.setNumberFormat('@');
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function doGet(e) {
  const p = e.parameter || {};
  if (!ok_(p.code) || !p.player) return json_({ ok: false });
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, rows: [] });
  const want = String(p.player).trim().toLowerCase();
  const rows = sh.getRange(2, 1, n, 8).getValues()
    .filter(r => String(r[1]).trim().toLowerCase() === want)
    .map(r => ({ character: r[2], updatedAt: r[6], csv: r[7] }));
  return json_({ ok: true, rows });
}
```

## Instructions to send to each player

Copy, fill in the passphrase, and send privately (e.g. WhatsApp DM).

```
Deathwatch Character Generator
Link: https://bjacques.github.io/40k-rpg/

1. Open the link in Safari (iPhone) or Chrome (Android).
   Do NOT use the WhatsApp in-app browser, it keeps its own separate storage.
2. Optional, iPhone: Share > Add to Home Screen. This keeps your character the longest.
3. Fill in the form. It saves itself on your device as you go.
   Use the same browser every time you come back.

Save to the GM's Google Sheet (preferred, lets you restore on another phone):
4. The Save bar is always at the bottom of the page. Press "More".
5. Type this passphrase into "Group passphrase (Google Sheets)":
   PASSPHRASE
6. Make sure "Players Name" is filled in (names are matched in lower case, so
   "Brad" and "brad" are the same). Edit anything and wait a few seconds.
   Or press "Save". The bar should say "Saved to sheet HH:MM".

Moving to a new phone or browser:
7. Open the link, type the same passphrase, type your Players Name, and press
   "Restore from sheet". Pick your character if asked.

Secondary fallback (file backup):
- Under "More", "Download CSV" saves a file. "Load CSV" brings it back. Do this now and then
  as a safety net, especially before clearing browser data.

Notes:
- Don't use private/incognito tabs, or clear site data, if you want local saves to last.
- Keep the passphrase private. Anyone with it can read and change the shared sheet.
```
