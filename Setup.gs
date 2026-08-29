/**
 * ============================================
 * SETUP SCRIPT — Bitrix24 Task Management System
 * ============================================
 * 
 * Run setupAllSheets() once to initialize everything.
 * Safe to re-run — skips tabs that already exist.
 * 
 * Architecture:
 *   - Sheets referenced by GID (immutable), never by name
 *   - Columns referenced by header name, never by index
 *   - Config sheet (GID 0) is the central control panel
 */

// ============================================
// CORE HELPERS — Used by all scripts
// ============================================

/**
 * Finds a sheet by its GID (sheet ID).
 * @param {number} gid
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheetByGid(gid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === parseInt(gid)) return sheets[i];
  }
  return null;
}

/**
 * Gets a config value from the Config sheet (GID 0).
 * @param {string} settingName
 * @returns {string}
 */
function getConfig(settingName) {
  var sheet = getSheetByGid(0);
  if (!sheet) throw new Error('Config sheet (GID 0) not found');
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === settingName) return String(data[i][1]).trim();
  }
  throw new Error('Config setting "' + settingName + '" not found');
}

/**
 * Sets a config value. Updates existing or appends new.
 * @param {string} settingName
 * @param {*} value
 */
function setConfig(settingName, value) {
  var sheet = getSheetByGid(0);
  if (!sheet) throw new Error('Config sheet (GID 0) not found');
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([settingName, value]);
}

/**
 * Gets a sheet using its GID stored in Config.
 * @param {string} configKey — e.g. 'MASTER_DATA_GID'
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetFromConfig(configKey) {
  var gid = parseInt(getConfig(configKey));
  var sheet = getSheetByGid(gid);
  if (!sheet) throw new Error('Sheet for config key "' + configKey + '" (GID ' + gid + ') not found');
  return sheet;
}

/**
 * Gets column index (1-based) by header name.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} headerName
 * @returns {number} 1-based column index
 */
function getColumnByHeader(sheet, headerName) {
  var headerRow = 1;
  try { headerRow = parseInt(getConfig('HEADER_ROW')) || 1; } catch (e) { /* default 1 */ }
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet "' + sheet.getName() + '" has no columns');
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) return i + 1;
  }
  throw new Error('Header "' + headerName + '" not found in sheet "' + sheet.getName() + '"');
}

/**
 * Gets multiple column indices as a map.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headerNames
 * @returns {Object} { headerName: 1-based index }
 */
function getColumnMap(sheet, headerNames) {
  var headerRow = 1;
  try { headerRow = parseInt(getConfig('HEADER_ROW')) || 1; } catch (e) { /* default 1 */ }
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet "' + sheet.getName() + '" has no columns');
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var h = 0; h < headerNames.length; h++) {
    var found = false;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === headerNames[h]) {
        map[headerNames[h]] = i + 1;
        found = true;
        break;
      }
    }
    if (!found) throw new Error('Header "' + headerNames[h] + '" not found in sheet "' + sheet.getName() + '"');
  }
  return map;
}

/**
 * Converts 1-based column index to letter (A, B, ... Z, AA, AB ...).
 * @param {number} col — 1-based
 * @returns {string}
 */
function colToLetter(col) {
  var letter = '';
  while (col > 0) {
    col--;
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col = Math.floor(col / 26);
  }
  return letter;
}

/**
 * Gets column letter by header name (convenience).
 */
function getColumnLetter(sheet, headerName) {
  return colToLetter(getColumnByHeader(sheet, headerName));
}


// ============================================
// SHEET DEFINITIONS
// ============================================

/**
 * Returns all sheet definitions with their headers.
 * Single source of truth for the entire schema.
 */
function getSheetDefinitions_() {
  return {
    CONFIG: {
      configKey: null, // GID 0, always exists
      tabName: 'Config',
      headers: ['Setting', 'Value'],
      initialData: [
        ['BITRIX_WEBHOOK_URL', 'https://eazyerp.bitrix24.in/rest/9797/ywdtfi50dffg2afp/'],
        ['BITRIX_USER_ID', '9797'],
        ['HEADER_ROW', '1'],
        ['DATA_START_ROW', '2'],
        ['MASTER_DATA_GID', ''],
        ['NEW_TASK_REVIEW_GID', ''],
        ['CLOSED_REVIEW_GID', ''],
        ['ARCHIVE_GID', ''],
        ['CUSTOMER_MASTER_GID', ''],
        ['TEAMMATES_GID', ''],
        ['BITRIX_STATUS_INCLUDE', '0,1,2,3,4,7,8,9'],
        ['BITRIX_GROUP_IDS', '0,164,182,241,343,345,347,419'],
        ['ARCHIVE_MAX_ROWS', '1000']
      ]
    },

    MASTER_DATA: {
      configKey: 'MASTER_DATA_GID',
      tabName: 'Master Data',
      headers: [
        'Bitrix ID',
        'Client Name',
        'Task Title',
        'Description',
        "Goutham's Remarks",
        'Tech Remarks',
        'Next Follow Up Date',
        'Contact Person',
        'Deadline Given',
        'Platform',
        'Task Owner',
        'Task Type',
        'Priority',
        'Current Status',
        'Created Date',
        'Project',
        'Sprint',
        'Backlog',
        'Parent ID',
        'Chat ID',
        'Team',
        'Total Points',
        'My Priority',
        'Chain',
        'Tags',
        'Is Exists?',
        'Focus'
      ]
    },

    NEW_TASK_REVIEW: {
      configKey: 'NEW_TASK_REVIEW_GID',
      tabName: 'New Task Review',
      headers: [
        'Bitrix ID',
        'Task Title',
        'Description',
        'Deadline',
        'Task Owner',
        'Assignee',
        'Stage',
        'Project',
        'Created Date',
        'Chat ID',
        'Client Name',
        'Contact Person',
        'Platform',
        'Task Type',
        'Priority',
        'Tags',
        'Pulled Date',
        'Review Status'
      ]
    },

    CLOSED_REVIEW: {
      configKey: 'CLOSED_REVIEW_GID',
      tabName: 'Closed Review',
      headers: [
        'Bitrix ID',
        'Task Title',
        'Client Name',
        'Last Known Status',
        'Last Known Stage',
        'Deadline',
        'Contact Person',
        'Tags',
        'Closed By',
        'Closed Date',
        'Verification Note',
        'Review Status'
      ]
    },

    ARCHIVE: {
      configKey: 'ARCHIVE_GID',
      tabName: 'Archive',
      headers: [
        'Bitrix ID',
        'Client Name',
        'Task Title',
        "Goutham's Remarks",
        'Tech Remarks',
        'Next Follow Up Date',
        'Contact Person',
        'Deadline Given',
        'Platform',
        'Task Owner',
        'Task Type',
        'Priority',
        'Current Status',
        'Created Date',
        'Team',
        'Total Points',
        'My Priority',
        'Chain',
        'Tags',
        'Focus',
        'Archive Status',
        'Archived Date'
      ]
    },

    CUSTOMER_MASTER: {
      configKey: 'CUSTOMER_MASTER_GID',
      tabName: 'Customer Master',
      headers: [
        'Client Name',
        'Weight',
        'Active'
      ]
    },

    TEAMMATES: {
      configKey: 'TEAMMATES_GID',
      tabName: 'Teammates',
      headers: [
        'Name',
        'Bitrix User ID',
        'Role',
        'Department',
        'Team',
        'Telegram Chat ID',
        'Active'
      ]
    }
  };
}


// ============================================
// SEED DATA
// ============================================

/**
 * Returns seed data for the Teammates sheet.
 * Source: Knowledge base §8 Contact Person routing.
 */
function getTeammatesSeedData_() {
  return [
    // Mobile Dev Team
    ['Suraj Singh',         9135,  'Mobile Dev Lead',    'Development',  'Tech', '', 'Yes'],
    ['Kanhaiya Lal',        8438,  'Mobile Developer',   'Development',  'Tech', '', 'Yes'],
    ['Anil Kumar Saini',    10033, 'Mobile Developer',   'Development',  'Tech', '', 'Yes'],
    ['Ajit Kumar',          9475,  'Mobile Developer',   'Development',  'Tech', '', 'Yes'],
    ['Kartik Panchal',      9887,  'Mobile Developer',   'Development',  'Tech', '', 'Yes'],

    // Web Dev Team
    ['Lokesh Kumar',        9102,  'Web Dev Lead',       'Development',  'Tech', '', 'Yes'],
    ['Arvind Kumar',        9473,  'Web Dev / Integration', 'Development', 'Tech', '', 'Yes'],
    ['Nilendra Pulipati',   9106,  'Web Developer',      'Development',  'Tech', '', 'Yes'],
    ['Wahid Zamal Siddiqui', 9112, 'Web Developer',      'Development',  'Tech', '', 'Yes'],
    ['Akshay Kumar',        9717,  'Web Developer',      'Development',  'Tech', '', 'Yes'],
    ['Radheshyam Gupta',    9805,  'Web Developer',      'Development',  'Tech', '', 'Yes'],
    ['Nirmal',              9743,  'Web Developer',      'Development',  'Tech', '', 'Yes'],

    // Testing Team
    ['Palvinder Kaur',      10251, 'Testing Manager',    'Testing',      'Tech', '', 'Yes'],
    ['Bhawna Bisht',        19,    'Test Head',          'Testing',      'Tech', '', 'Yes'],
    ['Bindu AN',            9153,  'Tester',             'Testing',      'Tech', '', 'Yes'],
    ['Shambhavi Kumari',    9145,  'Tester',             'Testing',      'Tech', '', 'Yes'],
    ['Shivani Malhotra',    10115, 'Tester',             'Testing',      'Tech', '', 'Yes'],
    ['Srihari B Sondur',    9755,  'Tester',             'Testing',      'Tech', '', 'Yes'],

    // Support / Coordination
    ['Goutham K',           9797,  'Technical Success Coordinator', 'Support', 'Support', '', 'Yes']
  ];
}


// ============================================
// SETUP: Main Entry Point
// ============================================

/**
 * Run this function once to set up all sheets.
 * Safe to re-run — existing tabs are skipped.
 */
function setupAllSheets() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '🔧 Setup Confirmation',
    'This will create all required tabs and configure the sheet.\n\n' +
    '• Config (GID 0) — settings & GID registry\n' +
    '• Master Data — main task tracking\n' +
    '• New Task Review — Bitrix pull staging\n' +
    '• Closed Review — disappeared task verification\n' +
    '• Archive — confirmed closed tasks\n' +
    '• Customer Master — client lookup\n' +
    '• Teammates — contact person lookup\n\n' +
    'Existing tabs will NOT be overwritten.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Setup cancelled.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = getSheetDefinitions_();

  try {
    // Step 1: Setup Config sheet (GID 0)
    Logger.log('Step 1: Setting up Config sheet...');
    setupConfigSheet_(ss, defs.CONFIG);

    // Step 2: Create all other tabs
    Logger.log('Step 2: Creating tabs...');
    var createdSheets = {};
    var tabKeys = ['MASTER_DATA', 'NEW_TASK_REVIEW', 'CLOSED_REVIEW', 'ARCHIVE', 'CUSTOMER_MASTER', 'TEAMMATES'];
    for (var t = 0; t < tabKeys.length; t++) {
      var key = tabKeys[t];
      var def = defs[key];
      var sheet = createTabIfNotExists_(ss, def.tabName, def.headers);
      createdSheets[key] = sheet;
      Logger.log('  Created/found: ' + def.tabName + ' (GID ' + sheet.getSheetId() + ')');
    }

    // Step 3: Record GIDs to Config
    Logger.log('Step 3: Recording GIDs to Config...');
    for (var t = 0; t < tabKeys.length; t++) {
      var key = tabKeys[t];
      var def = defs[key];
      if (def.configKey) {
        setConfig(def.configKey, createdSheets[key].getSheetId());
      }
    }

    // Step 4: Seed Teammates data
    Logger.log('Step 4: Seeding Teammates data...');
    seedTeammatesData_(createdSheets.TEAMMATES);

    // Step 5: Format all sheets
    Logger.log('Step 5: Formatting sheets...');
    formatSheet_(getSheetByGid(0), defs.CONFIG.headers.length);
    for (var t = 0; t < tabKeys.length; t++) {
      var key = tabKeys[t];
      var def = defs[key];
      formatSheet_(createdSheets[key], def.headers.length);
    }

    // Step 6: Set Archive data validation for Archive Status
    Logger.log('Step 6: Setting up data validations...');
    setupDataValidations_(createdSheets);

    SpreadsheetApp.flush();

    ui.alert(
      '✅ Setup Complete!',
      'All sheets created and configured successfully.\n\n' +
      'Next steps:\n' +
      '1. Add your clients to the "Customer Master" tab\n' +
      '2. Add any additional teammates to the "Teammates" tab\n' +
      '3. Check the Config tab — all GIDs are recorded there\n\n' +
      '⚠️ Do NOT rename the Config tab or change GID 0.',
      ui.ButtonSet.OK
    );

    Logger.log('✅ Setup complete!');

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log(e.stack);
    ui.alert('❌ Setup Error', e.message + '\n\nCheck Execution Log for details.', ui.ButtonSet.OK);
  }
}


// ============================================
// SETUP: Internal Functions
// ============================================

/**
 * Sets up the Config sheet (GID 0) with headers and initial settings.
 */
function setupConfigSheet_(ss, configDef) {
  var sheet = getSheetByGid(0);
  if (!sheet) throw new Error('Cannot find the default sheet (GID 0). It should already exist.');

  // Rename to 'Config' if not already
  if (sheet.getName() !== configDef.tabName) {
    sheet.setName(configDef.tabName);
  }

  // Clear and write headers + data
  sheet.clear();
  sheet.getRange(1, 1, 1, configDef.headers.length).setValues([configDef.headers]);

  // Write initial settings
  var data = configDef.initialData;
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * Creates a tab if it doesn't already exist. Writes headers to Row 1.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function createTabIfNotExists_(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  } else {
    // Tab exists — check if headers match, update if needed
    var existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var headersMatch = true;
    for (var i = 0; i < headers.length; i++) {
      if (String(existingHeaders[i]).trim() !== headers[i]) {
        headersMatch = false;
        break;
      }
    }
    if (headersMatch) return sheet; // Already set up correctly
  }

  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

/**
 * Seeds the Teammates sheet with known team member data.
 * Only adds if the sheet has no data below the header row.
 */
function seedTeammatesData_(sheet) {
  var dataStartRow = 2; // Teammates doesn't use Formula Row
  var lastRow = sheet.getLastRow();

  // Skip if data already exists
  if (lastRow >= dataStartRow) {
    Logger.log('  Teammates already has data, skipping seed.');
    return;
  }

  var seedData = getTeammatesSeedData_();
  if (seedData.length > 0) {
    sheet.getRange(dataStartRow, 1, seedData.length, seedData[0].length).setValues(seedData);
    Logger.log('  Seeded ' + seedData.length + ' teammates.');
  }
}



/**
 * Formats a sheet: freeze header row, bold headers, auto-resize, background color.
 */
function formatSheet_(sheet, numCols) {
  if (!sheet || numCols <= 0) return;

  // Freeze header row
  sheet.setFrozenRows(1);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setBackground('#e8edf2');
  headerRange.setFontColor('#334155');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(false);

  // Set row height for header
  sheet.setRowHeight(1, 32);

  // Set minimum column width
  for (var c = 1; c <= numCols; c++) {
    sheet.setColumnWidth(c, 140);
  }

  // Add border below header
  headerRange.setBorder(null, null, true, null, null, null, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Sets up data validations for dropdown columns.
 */
function setupDataValidations_(sheets) {
  // Archive Status validation: "Completed" or "On-Hold"
  var archiveSheet = sheets.ARCHIVE;
  if (archiveSheet) {
    try {
      var statusCol = getColumnByHeader(archiveSheet, 'Archive Status');
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Completed', 'On-Hold'], true)
        .setAllowInvalid(false)
        .setHelpText('Select: Completed or On-Hold')
        .build();
      // Apply to column from row 2 down (1000 rows max)
      archiveSheet.getRange(2, statusCol, 1000, 1).setDataValidation(statusRule);
      Logger.log('  ✓ Archive Status validation set');
    } catch (e) {
      Logger.log('  ⚠ Could not set Archive Status validation: ' + e.message);
    }
  }

  // New Task Review — Review Status validation
  var nrSheet = sheets.NEW_TASK_REVIEW;
  if (nrSheet) {
    try {
      var rsCol = getColumnByHeader(nrSheet, 'Review Status');
      var rsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pending', 'Confirmed', 'Skipped'], true)
        .setAllowInvalid(false)
        .setHelpText('Select: Pending, Confirmed, or Skipped')
        .build();
      nrSheet.getRange(2, rsCol, 1000, 1).setDataValidation(rsRule);
      Logger.log('  ✓ New Task Review Status validation set');
    } catch (e) {
      Logger.log('  ⚠ Could not set NTR Review Status validation: ' + e.message);
    }
  }

  // Closed Review — Review Status validation
  var crSheet = sheets.CLOSED_REVIEW;
  if (crSheet) {
    try {
      var crRsCol = getColumnByHeader(crSheet, 'Review Status');
      var crRsRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pending', 'Archived', 'Kept in Master'], true)
        .setAllowInvalid(false)
        .setHelpText('Select: Pending, Archived, or Kept in Master')
        .build();
      crSheet.getRange(2, crRsCol, 1000, 1).setDataValidation(crRsRule);
      Logger.log('  ✓ Closed Review Status validation set');
    } catch (e) {
      Logger.log('  ⚠ Could not set CR Review Status validation: ' + e.message);
    }
  }

  // Customer Master — Active validation
  var cmSheet = sheets.CUSTOMER_MASTER;
  if (cmSheet) {
    try {
      var cmActiveCol = getColumnByHeader(cmSheet, 'Active');
      var cmActiveRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true)
        .setAllowInvalid(false)
        .build();
      cmSheet.getRange(2, cmActiveCol, 1000, 1).setDataValidation(cmActiveRule);
      Logger.log('  ✓ Customer Master Active validation set');
    } catch (e) {
      Logger.log('  ⚠ Could not set CM Active validation: ' + e.message);
    }
  }

  // Teammates — Active validation
  var tmSheet = sheets.TEAMMATES;
  if (tmSheet) {
    try {
      var tmActiveCol = getColumnByHeader(tmSheet, 'Active');
      var tmActiveRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Yes', 'No'], true)
        .setAllowInvalid(false)
        .build();
      tmSheet.getRange(2, tmActiveCol, 1000, 1).setDataValidation(tmActiveRule);
      Logger.log('  ✓ Teammates Active validation set');
    } catch (e) {
      Logger.log('  ⚠ Could not set Teammates Active validation: ' + e.message);
    }
  }
}



// ============================================
// CUSTOM MENU
// ============================================

/**
 * Creates a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Bitrix Tools')
    .addItem('📋 Setup All Sheets', 'setupAllSheets')
    .addSeparator()
    .addItem('ℹ️ Show Config', 'showConfig')
    .addToUi();
}

/**
 * Shows current config in a dialog.
 */
function showConfig() {
  var sheet = getSheetByGid(0);
  var data = sheet.getDataRange().getValues();
  var lines = [];
  for (var i = 1; i < data.length; i++) { // skip header
    var val = String(data[i][1]);
    // Mask webhook URL for safety
    if (String(data[i][0]).indexOf('WEBHOOK') >= 0 && val.length > 20) {
      val = val.substring(0, 40) + '...***';
    }
    lines.push(data[i][0] + ':  ' + val);
  }
  SpreadsheetApp.getUi().alert('⚙️ Current Configuration', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
