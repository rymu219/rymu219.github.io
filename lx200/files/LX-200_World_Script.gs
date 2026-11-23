/**
 * LX-200 World: reset, deliver Message 67, and a read-only feed for the demo hub.
 * Paste into Extensions > Apps Script in the LX-200 World Google Sheet.
 * Fictional practice data only.
 */
const LIVE_TABS = ['Inbox','Team chat','Gates','Schedule','Budget & Hours','Payments','Documents','RAID','Contacts'];
const MESSAGE_67 = [67, "Nov 23, 2026", "Jason Wu <jason.wu@hengtaimold.example>", "Ryan Murphy <ryan.murphy@cypressbend.example>", "", "URGENT: LX-200 Rev C impact on tool and T1 date", "", "CAUTION: This email originated from outside Cypress Bend Molding. Do not click links or open attachments unless you recognize the sender.\n\nDear Ryan,\n\nToday we received drawing Rev C from your sales team. The rib change from 1.5 mm to 2.0 mm affects the slider design. We need 8 more working days for tool design, and the design must be approved again.\n\nNew T1 date: Feb 5, 2027.\n\nPlease also note our factory closes for Chinese New Year from Jan 30 and reopens Feb 22. If T1 cannot be completed before Jan 30, it will be after Feb 22.\n\nWe are still waiting for the deposit of USD 28,000 to order steel.\n\nPlease advise how you want to proceed.\n\nBest regards,\nJason Wu\nProject Manager, Hengtai Precision Mold"];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('LX-200')
    .addItem('Reset world', 'resetWorld')
    .addItem('Deliver Message 67', 'deliver67')
    .addSeparator()
    .addItem('Set hub key', 'setHubKey')
    .addToUi();
}

function resetWorld() {
  const ss = SpreadsheetApp.getActive();
  LIVE_TABS.forEach(name => {
    const live = ss.getSheetByName(name), base = ss.getSheetByName('_base_' + name);
    if (!live || !base) throw new Error('Missing tab: ' + name);
    live.clearContents();
    base.getDataRange().copyTo(live.getRange(1, 1));
  });
  const f = ss.getSheetByName('Findings');
  if (f.getLastRow() > 1) f.getRange(2, 1, f.getLastRow() - 1, Math.max(f.getLastColumn(), 7)).clearContent();
  setControl_('message67_delivered', 'FALSE');
  setControl_('last_reset', new Date().toISOString());
  setControl_('run_label', 'Opening');
  return 'reset';
}

function deliver67() {
  if (getControl_('message67_delivered') === 'TRUE') return 'already delivered';
  const inbox = SpreadsheetApp.getActive().getSheetByName('Inbox');
  inbox.appendRow(MESSAGE_67);
  setControl_('message67_delivered', 'TRUE');
  setControl_('run_label', 'Message 67');
  return 'delivered';
}

function setHubKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Hub key', 'Pick a short private word. The hub asks for it before Reset or Deliver.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('HUB_KEY', r.getResponseText().trim());
    ui.alert('Saved.');
  }
}

/** Web app: GET returns the world as JSON. ?action=reset|deliver67&key=... changes it. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let result = 'ok';
  if (p.action) {
    const key = PropertiesService.getScriptProperties().getProperty('HUB_KEY') || '';
    if (!key || p.key !== key) return json_({ error: 'bad key' });
    if (p.action === 'reset') result = resetWorld();
    else if (p.action === 'deliver67') result = deliver67();
    else return json_({ error: 'unknown action' });
  }
  const ss = SpreadsheetApp.getActive();
  return json_({
    result: result,
    updated: new Date().toISOString(),
    control: readControl_(),
    inbox: rows_(ss.getSheetByName('Inbox')),
    chat: rows_(ss.getSheetByName('Team chat')),
    findings: rows_(ss.getSheetByName('Findings'))
  });
}

function rows_(sh) {
  const v = sh.getDataRange().getDisplayValues();
  const head = v.shift();
  return v.filter(r => r.some(x => String(x).trim() !== '')).map(r => {
    const o = {}; head.forEach((h, i) => o[h] = r[i]); return o;
  });
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function readControl_() {
  const v = SpreadsheetApp.getActive().getSheetByName('Control').getDataRange().getDisplayValues();
  const o = {}; v.slice(1).forEach(r => o[r[0]] = r[1]); return o;
}
function getControl_(k) { return readControl_()[k]; }
function setControl_(k, val) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Control');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val]);
}
