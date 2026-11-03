/**
 * App Inventor Academy prototype — receives anonymous test sessions from the
 * GitHub Pages site and keeps one row per session in the "Sessions" sheet.
 * A "Summary" sheet computes the H1–H6 indicators against their targets.
 *
 * Deploy: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 */
const HEADERS = [
  'id', 'startedAt', 'updatedAt', 'lang', 'device', 'furthest', 'path',
  'registered', 'registerSeconds', 'dashboardActionSeconds', 'country', 'level', 'area',
  'stepsDone', 'evidenceType', 'evidenceStarted', 'evidenceStep', 'evidenceCompleted',
  'certificateViewed', 'shareOpened', 'shareAsWritten', 'downloadClicked',
  'upgradeViewed', 'educatorClicked', 'institutionalClicked',
  'useInClass', 'evidenceEase', 'wouldPay', 'comment',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data || typeof data.id !== 'string' || !/^s[a-z0-9]{6,24}$/.test(data.id)) return reply_({ ok: false });
    const sheet = sessionsSheet_();
    const row = HEADERS.map(h => clean_(data[h]));
    const last = sheet.getLastRow();
    const ids = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]) : [];
    const index = ids.indexOf(data.id);
    if (index >= 0) sheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply_({ ok: true, service: 'App Inventor Academy prototype sessions' });
}

// Run once from the editor to create both sheets before the first test.
function setup() {
  sessionsSheet_();
  summarySheet_();
}

function sessionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    sheet = ss.insertSheet('Sessions');
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    summarySheet_();
  }
  return sheet;
}

function summarySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Summary')) return;
  const sheet = ss.insertSheet('Summary', 0);
  const s = 'Sessions!';
  const rows = [
    ['Hypothesis', 'Indicator', 'Observed', 'Target', 'n', 'Status'],
    ['H1', 'Registered in under 3 minutes',
      `=IFERROR(COUNTIFS(${s}H2:H,TRUE,${s}I2:I,"<180")/COUNTIF(${s}H2:H,TRUE),"")`, 0.8, `=COUNTIF(${s}H2:H,TRUE)`],
    ['H2', 'Tapped the next step within 60 s (proxy, confirm aloud)',
      `=IFERROR(COUNTIFS(${s}H2:H,TRUE,${s}J2:J,"<=60")/COUNTIF(${s}H2:H,TRUE),"")`, '', `=COUNTIF(${s}H2:H,TRUE)`],
    ['H3', 'Would use Module 1 in their next class',
      `=IFERROR(COUNTIF(${s}Z2:Z,"yes")/COUNTA(${s}Z2:Z),"")`, 0.7, `=COUNTA(${s}Z2:Z)`],
    ['H4', 'Completed the evidence upload after starting it',
      `=IFERROR(COUNTIFS(${s}P2:P,TRUE,${s}R2:R,TRUE)/COUNTIF(${s}P2:P,TRUE),"")`, 0.75, `=COUNTIF(${s}P2:P,TRUE)`],
    ['H5', 'Would share the certificate as written',
      `=IFERROR(COUNTIF(${s}U2:U,"yes")/COUNTA(${s}U2:U),"")`, 0.6, `=COUNTA(${s}U2:U)`],
    ['H6', 'Opened a paid plan after seeing Module 2 locked',
      `=IFERROR((COUNTIFS(${s}W2:W,TRUE,${s}X2:X,TRUE)+COUNTIFS(${s}W2:W,TRUE,${s}X2:X,FALSE,${s}Y2:Y,TRUE))/COUNTIF(${s}W2:W,TRUE),"")`, 0.4, `=COUNTIF(${s}W2:W,TRUE)`],
  ];
  rows.forEach((r, i) => {
    const line = i + 1;
    sheet.getRange(line, 1, 1, 2).setValues([[r[0], r[1]]]);
    if (i === 0) { sheet.getRange(1, 1, 1, 6).setValues([r]).setFontWeight('bold'); return; }
    sheet.getRange(line, 3).setFormula(r[2]);
    sheet.getRange(line, 4).setValue(r[3]);
    sheet.getRange(line, 5).setFormula(r[4]);
    sheet.getRange(line, 6).setFormula(
      `=IF(OR(E${line}<3,C${line}=""),"Not enough data",IF(D${line}="","Confirm in session",IF(C${line}>D${line},"Meets target","Below target")))`);
  });
  sheet.getRange('C2:D7').setNumberFormat('0%');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
}

// Keep text short and stop spreadsheet formulas from being injected through the comment box.
function clean_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean' || typeof v === 'number') return v;
  return String(v).slice(0, 600).replace(/^[=+\-@]/, "'$&");
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
