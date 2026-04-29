var REFRESH_CONTROL_SHEET_NAME = 'Refresh Control';
var REFRESH_CONTROL_HEADERS = ['field', 'value'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SLA Report')
    .addItem('Refresh Report', 'requestRefresh')
    .addItem('Initialize Refresh Control', 'initializeRefreshControl')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function initializeRefreshControl() {
  writeRefreshControl_({
    lastRefreshStatus: 'Never Run',
    requestedAt: '',
    requestedBy: '',
    backupReference: '',
    message: 'Click "Refresh Report" to mark a request, then run ./scripts/refresh_report_local.sh from this repo.',
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Refresh Control initialized.',
    'SLA Report',
    5
  );
}

function requestRefresh() {
  var currentBackupReference = readRefreshControlField_('backup_reference') || '';
  var requestedAt = formatTimestamp_(new Date());
  var requestedBy = getRequesterEmail_();

  writeRefreshControl_({
    lastRefreshStatus: 'Requested',
    requestedAt: requestedAt,
    requestedBy: requestedBy,
    backupReference: currentBackupReference,
    message: 'Refresh requested. Run ./scripts/refresh_report_local.sh from this repo to rebuild the report.',
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Refresh requested. Run ./scripts/refresh_report_local.sh from this repo.',
    'SLA Report',
    8
  );
}

function readRefreshControlField_(fieldName) {
  var sheet = getOrCreateRefreshControlSheet_();
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === fieldName) {
      return values[i][1];
    }
  }

  return '';
}

function writeRefreshControl_(state) {
  var sheet = getOrCreateRefreshControlSheet_();
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([REFRESH_CONTROL_HEADERS]);
  sheet.getRange(2, 1, 5, 2).setValues([
    ['last_refresh_status', state.lastRefreshStatus],
    ['requested_at', state.requestedAt],
    ['requested_by', state.requestedBy],
    ['backup_reference', state.backupReference],
    ['message', state.message],
  ]);
  sheet.getRange('A1:B6').setVerticalAlignment('middle');
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 560);
}

function getOrCreateRefreshControlSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REFRESH_CONTROL_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(REFRESH_CONTROL_SHEET_NAME);
  }

  return sheet;
}

function getRequesterEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      return email;
    }
  } catch (error) {
    // Ignore and fall back to the next available identity source.
  }

  try {
    var fallbackEmail = Session.getEffectiveUser().getEmail();
    if (fallbackEmail) {
      return fallbackEmail;
    }
  } catch (error) {
    // Ignore and fall back to a human-readable placeholder.
  }

  return 'Not available';
}

function formatTimestamp_(date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd HH:mm:ss z');
}
