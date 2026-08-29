/**
 * ============================================
 * BITRIX24 TASK REVIEW — WEB APP BACKEND
 * ============================================
 *
 * Google Apps Script web app for reviewing new Bitrix tasks
 * and managing them into the Master Data sheet.
 *
 * Architecture:
 *   - Sheets referenced by GID (from Config sheet, GID 0)
 *   - Columns referenced by header name, never by index
 *   - Config sheet (GID 0) is the central control panel
 */


// ============================================
// CONFIGURATION & HELPERS
// ============================================

/**
 * Finds a sheet by its GID (sheet ID).
 * @param {number} gid
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheetByGid(gid) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === parseInt(gid)) return sheets[i];
  }
  return null;
}

function getConfigSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  return ss.getSheetByName('Config') || ss.getSheetByName('config') || getSheetByGid(0) || ss.getSheets()[0];
}

/**
 * Gets a config value from the Config sheet.
 * Includes built-in fallbacks so missing optional keys never crash the app.
 */
function getConfig(settingName, defaultValue) {
  var defaults = {
    'SUPER_ADMIN_EMAIL': 'goutham@eazyerp.com',
    'MANDATORY_FIELDS': 'Client Name,Contact Person,Platform,Task Type,Priority',
    'HEADER_ROW': '1',
    'DATA_START_ROW': '2',
    'ARCHIVE_MAX_ROWS': '1000',
    'TELEGRAM_MINI_APP_URL': 'https://gouthamk-eazy.github.io/EazyERP-Task-Hub/'
  };

  var fallback = defaultValue !== undefined ? defaultValue : defaults[settingName];

  var sheet = getConfigSheet_();
  if (!sheet) {
    if (fallback !== undefined) return fallback;
    throw new Error('Config sheet not found');
  }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === settingName && data[i][1] !== undefined && data[i][1] !== null && String(data[i][1]).trim() !== '') {
        return String(data[i][1]).trim();
      }
    }
  } catch(e) {}

  if (fallback !== undefined) return fallback;
  throw new Error('Config setting "' + settingName + '" not found');
}

/**
 * Sets a config value. Updates existing or appends new.
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
 */
function getSheetFromConfig(configKey) {
  var gid = parseInt(getConfig(configKey));
  var sheet = getSheetByGid(gid);
  if (!sheet) throw new Error('Sheet for "' + configKey + '" (GID ' + gid + ') not found');
  return sheet;
}

/**
 * Gets column index (1-based) by header name, using normalized matching.
 */
function getColumnByHeader(sheet, headerName) {
  var headerRow = 1;
  try { headerRow = parseInt(getConfig('HEADER_ROW', '1')) || 1; } catch (e) {}
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet "' + sheet.getName() + '" has no columns');
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var targetLower = String(headerName).trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  for (var i = 0; i < headers.length; i++) {
    var hLower = String(headers[i]).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (hLower === targetLower) return i + 1;
  }

  // Fallback: check substring match for formula-generated headers
  for (var j = 0; j < headers.length; j++) {
    var hj = String(headers[j]).trim().toLowerCase();
    if (hj.indexOf(String(headerName).toLowerCase()) >= 0) return j + 1;
  }

  throw new Error('Header "' + headerName + '" not found in sheet "' + sheet.getName() + '"');
}

/**
 * Gets all headers from a sheet as an array.
 */
function getHeaders(sheet) {
  var headerRow = 1;
  try { headerRow = parseInt(getConfig('HEADER_ROW', '1')) || 1; } catch (e) {}
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
}

/**
 * Gets the actual last data row in a sheet, ignoring ARRAYFORMULA blank results.
 * Scans backwards from sheet.getLastRow() in the key column (or first column) for the last non-empty value.
 * [Project Level] — prevents ARRAYFORMULA from causing getLastRow() to return max sheet rows (e.g. 1000).
 */
function getLastDataRow(sheet, keyColumnHeader) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return lastRow;

  var headerRow = 1;
  try { headerRow = parseInt(getConfig('HEADER_ROW', '1')) || 1; } catch (e) {}
  if (lastRow <= headerRow) return headerRow;

  if (keyColumnHeader) {
    try {
      var col = getColumnByHeader(sheet, keyColumnHeader);
      var values = sheet.getRange(1, col, lastRow, 1).getValues();
      for (var r = values.length - 1; r >= headerRow; r--) {
        var val = String(values[r][0]).trim();
        if (val !== '' && val !== 'null' && val !== 'undefined') {
          return r + 1; // 1-based row index
        }
      }
      return headerRow;
    } catch (e) {}
  }

  // Fallback: check column 1
  try {
    var values1 = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var r1 = values1.length - 1; r1 >= headerRow; r1--) {
      var val1 = String(values1[r1][0]).trim();
      if (val1 !== '' && val1 !== 'null' && val1 !== 'undefined') {
        return r1 + 1;
      }
    }
  } catch (e) {}

  return headerRow;
}

/**
 * Reads all data rows from a sheet as an array of objects (keyed by header name).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object[]}
 */
function getSheetDataAsObjects(sheet) {
  var headerRow = 1;
  var dataStartRow = 2;
  try { headerRow = parseInt(getConfig('HEADER_ROW', '1')) || 1; } catch (e) {}
  try { dataStartRow = parseInt(getConfig('DATA_START_ROW', '2')) || 2; } catch (e) {}

  var lastRow = getLastDataRow(sheet, 'Bitrix ID');
  var lastCol = sheet.getLastColumn();

  if (lastRow < dataStartRow || lastCol === 0) return [];

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
  var result = [];

  for (var r = 0; r < data.length; r++) {
    // Skip fully empty rows
    var isEmpty = data[r].every(function(cell) { return cell === '' || cell === null || cell === undefined; });
    if (isEmpty) continue;

    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c]).trim();
      if (key) {
        var val = data[r][c];
        // Convert dates to ISO strings for JSON transport
        if (val instanceof Date) {
          obj[key] = val.toISOString();
        } else {
          obj[key] = val;
        }
      }
    }
    result.push(obj);
  }
  return result;
}

/**
 * Helper to get a value from a data object for any header, supporting aliases,
 * case-insensitivity, and trimmed comparisons.
 */
function getValueForHeader_(header, dataObj) {
  if (!header || !dataObj) return '';
  var hClean = String(header).trim();
  var hLower = hClean.toLowerCase().replace(/[^a-z0-9]/g, '');

  // 1. Exact match
  if (dataObj.hasOwnProperty(hClean) && dataObj[hClean] !== undefined && dataObj[hClean] !== null) {
    return dataObj[hClean];
  }

  // 2. Case-insensitive / normalized key match
  for (var k in dataObj) {
    if (dataObj.hasOwnProperty(k)) {
      var kLower = String(k).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      if (kLower === hLower && dataObj[k] !== undefined && dataObj[k] !== null) {
        return dataObj[k];
      }
    }
  }

  // 3. Known Aliases Mapping
  var ALIASES = {
    'bitrixid': ['bitrixid', 'id', 'taskid'],
    'clientname': ['clientname', 'client', 'customer', 'customername', 'shortcode'],
    'tasktitle': ['tasktitle', 'title', 'name'],
    'currentstatus': ['currentstatus', 'status', 'stage', 'currentstage', 'lastknownstatus', 'lastknownstage'],
    'stage': ['stage', 'currentstatus', 'status', 'lastknownstage'],
    'deadlinegiven': ['deadlinegiven', 'deadline', 'duedate'],
    'deadline': ['deadline', 'deadlinegiven', 'duedate'],
    'createddate': ['createddate', 'created', 'datecreated'],
    'nextfollowupdate': ['nextfollowupdate', 'nextfollowup', 'followupdate', 'followup'],
    'project': ['project', 'group', 'projectgroup', 'projectscrum', 'projectscrumid'],
    'sprint': ['sprint', 'sprintname', 'sprintid'],
    'parentid': ['parentid', 'parenttaskid', 'parent'],
    'chatid': ['chatid', 'chat'],
    'backlog': ['backlog', 'isbacklog', 'backlogid'],
    'gouthamsremarks': ['gouthamsremarks', 'gouthamremarks', 'remarks', 'myremarks'],
    'techremarks': ['techremarks', 'technicalremarks'],
    'platform': ['platform', 'platforms'],
    'tasktype': ['tasktype', 'type'],
    'priority': ['priority'],
    'mypriority': ['mypriority'],
    'team': ['team'],
    'totalpoints': ['totalpoints', 'points'],
    'chain': ['chain'],
    'isexists': ['isexists', 'isexist', 'exists'],
    'description': ['description', 'desc'],
    'reviewstatus': ['reviewstatus', 'status'],
    'contactperson': ['contactperson', 'assignee', 'contact'],
    'taskowner': ['taskowner', 'owner', 'creator'],
    'closedby': ['closedby'],
    'closeddate': ['closeddate']
  };

  if (ALIASES[hLower]) {
    var candidateAliases = ALIASES[hLower];
    for (var c = 0; c < candidateAliases.length; c++) {
      var targetAlias = candidateAliases[c];
      for (var dk in dataObj) {
        if (dataObj.hasOwnProperty(dk)) {
          var dkLower = String(dk).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
          if (dkLower === targetAlias && dataObj[dk] !== undefined && dataObj[dk] !== null && String(dataObj[dk]).trim() !== '') {
            return dataObj[dk];
          }
        }
      }
    }
  }

  return '';
}

var FORMULA_HEADERS_SET = {
  'mypriority': true,
  'contactperson': true,
  'taskowner': true,
  'team': true,
  'totalpoints': true
};

/**
 * Safely writes values to a row in a sheet, skipping columns managed by array formulas
 * (My Priority [F], Contact Person [I], Task Owner [L], Team [S], Total Points [W])
 * so array formulas in Row 1 are never blocked or broken.
 */
function writeRowSafe_(sheet, rowIndex, dataObj) {
  var headers = getHeaders(sheet);
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    var hClean = String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (FORMULA_HEADERS_SET[hClean]) {
      continue; // Skip formula columns
    }
    var val = getValueForHeader_(h, dataObj);
    sheet.getRange(rowIndex, c + 1).setValue(val);
  }
}

/**
 * Writes an object (keyed by header name) as a new row in a sheet.
 * Uses getLastDataRow to write to the exact next non-empty row, bypassing ARRAYFORMULA issue.
 */
function appendRowAsObject(sheet, dataObj) {
  var nextRow = getLastDataRow(sheet, 'Bitrix ID') + 1;
  writeRowSafe_(sheet, nextRow, dataObj);
}

/**
 * Updates a specific row in a sheet using an object mapped by header names.
 */
function updateRowAsObject(sheet, rowIndex, dataObj) {
  writeRowSafe_(sheet, rowIndex, dataObj);
}

/**
 * Strips Bitrix BBCode tags from text, returning clean plain text.
 * Handles: [b], [i], [u], [s], [url], [list], [*], [disk], [code], [quote], etc.
 * [Project Level] — used across the entire app.
 */
function stripBBCode(text) {
  if (!text) return '';
  return String(text)
    .replace(/\[disk[^\]]*\]/gi, '')                         // [disk file id=... width=... height=...]
    .replace(/\[url=([^\]]*)\]([^\[]*)\[\/url\]/gi, '$2')    // [url=...]text[/url] → text
    .replace(/\[img\][^\[]*\[\/img\]/gi, '')                 // [img]...[/img] → remove
    .replace(/\[\/?(?:b|i|u|s|code|quote|color|size|font|align|center|right|left|indent|table|tr|td|th|hr|p|br)\b[^\]]*\]/gi, '')  // standard tags
    .replace(/\[\/?list\]/gi, '')                            // [list] [/list]
    .replace(/\[\*\]/gi, '• ')                               // [*] → bullet
    .replace(/\[\/?[a-z]+[^\]]*\]/gi, '')                    // catch-all for remaining tags
    .replace(/\r\n/g, '\n')                                  // normalize line endings
    .replace(/\n{3,}/g, '\n\n')                              // collapse excessive newlines
    .trim();
}

/**
 * Truncates text to a max length, appending "..." if truncated.
 */
function truncateText(text, maxLength) {
  if (!text) return '';
  var str = String(text);
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength) + '...';
}


/**
 * Converts column index (1-based) to letter.
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
 * Default Project/Scrum ID -> Name Map
 * [Project Level] — ensures Project column displays readable names instead of numeric IDs.
 */
var DEFAULT_GROUP_NAMES = {
  '0': 'General',
  '164': 'DMS Dev Team',
  '182': 'Dummy Android App',
  '241': 'Recibo Production Issue',
  '343': 'Recibo 2.0 (iOS / Hybrid App) - Customization',
  '345': 'Recibo 1.0 - Android app - Customization',
  '347': 'Recibo Web Customization',
  '419': 'Recibo SFA Scrum'
};

/**
 * Resolves Group/Project ID to a human-readable Group Name.
 * [Project Level]
 */
function getGroupName(groupId, groupNameMap) {
  var idStr = String(groupId || '0').trim();
  if (groupNameMap && groupNameMap[idStr]) return groupNameMap[idStr];
  if (DEFAULT_GROUP_NAMES[idStr]) return DEFAULT_GROUP_NAMES[idStr];
  if (idStr === '0' || idStr === '') return 'General';

  // Fallback: try fetching from Bitrix API sonet_group.get
  try {
    var res = callBitrixApi('sonet_group.get', { GROUP_ID: parseInt(idStr) });
    if (res && res.result && res.result[0] && res.result[0].NAME) {
      return res.result[0].NAME;
    }
  } catch (e) {
    Logger.log('sonet_group.get failed for group ID ' + idStr + ': ' + e.message);
  }

  return 'Project ' + idStr;
}

/**
 * Fetches Stage, Epic & Sprint lookups from Bitrix API.
 * Returns { stageMap, sprintMap } mapping numeric IDs to human-readable names.
 * [Project Level] — resolves numeric stageId, epicId, and sprintId to names.
 */
function getScrumLookupMaps(groupIds) {
  var stageMap = {};
  var sprintMap = {};

  function addStage(id, name) {
    if (id && name) {
      stageMap[String(id)] = String(name).trim();
    }
  }

  function addSprint(id, name) {
    if (id && name) {
      sprintMap[String(id)] = String(name).trim();
    }
  }

  function parseStages(res) {
    if (res && res.result) {
      var stages = res.result;
      if (Array.isArray(stages)) {
        stages.forEach(function(st) {
          addStage(st.id || st.ID, st.name || st.NAME || st.title || st.TITLE);
        });
      } else if (typeof stages === 'object') {
        for (var sid in stages) {
          var stObj = stages[sid];
          if (stObj) {
            addStage(stObj.id || stObj.ID || sid, stObj.name || stObj.NAME || stObj.title || stObj.TITLE);
          }
        }
      }
    }
  }

  // 1. Fetch default / global stages (entityId = 0)
  try {
    parseStages(callBitrixApi('task.stages.get', { entityId: 0 }));
  } catch (e) {
    Logger.log('task.stages.get entityId=0 warning: ' + e.message);
  }

  if (!groupIds || groupIds.length === 0) return { stageMap: stageMap, sprintMap: sprintMap };

  var unique = [];
  for (var i = 0; i < groupIds.length; i++) {
    var id = parseInt(groupIds[i]);
    if (id > 0 && unique.indexOf(id) === -1) unique.push(id);
  }

  for (var j = 0; j < unique.length; j++) {
    var gid = unique[j];

    // 2. Standard Group Kanban stages
    try {
      parseStages(callBitrixApi('task.stages.get', { entityId: gid }));
      parseStages(callBitrixApi('task.stages.get', { entityId: gid, entityType: 'G' }));
    } catch (e) {
      Logger.log('task.stages.get failed for group ' + gid + ': ' + e.message);
    }

    // 3. Scrum Epics (tasks.api.scrum.epic.list)
    try {
      var epicRes = callBitrixApi('tasks.api.scrum.epic.list', {
        filter: { GROUP_ID: gid },
        select: ['ID', 'NAME']
      });
      if (epicRes && epicRes.result && Array.isArray(epicRes.result)) {
        epicRes.result.forEach(function(ep) {
          addStage(ep.id || ep.ID, ep.name || ep.NAME);
        });
      }
    } catch (eEpic) {
      Logger.log('tasks.api.scrum.epic.list warning for group ' + gid + ': ' + eEpic.message);
    }

    // 4. Scrum Sprints & Scrum Kanban Stages (tasks.api.scrum.sprint.list & kanban.getStages)
    try {
      var sprintRes = callBitrixApi('tasks.api.scrum.sprint.list', {
        filter: { GROUP_ID: gid },
        select: ['ID', 'NAME']
      });
      if (sprintRes && sprintRes.result && Array.isArray(sprintRes.result)) {
        sprintRes.result.forEach(function(sp) {
          var spId = sp.id || sp.ID;
          var spName = sp.name || sp.NAME;
          if (spId) {
            addSprint(spId, spName || ('Sprint ' + spId));
            try {
              var kStages = callBitrixApi('tasks.api.scrum.kanban.getStages', { sprintId: parseInt(spId) });
              parseStages(kStages);
            } catch (eK) {
              Logger.log('tasks.api.scrum.kanban.getStages warning for sprint ' + spId + ': ' + eK.message);
            }
          }
        });
      }
    } catch (eSprint) {
      Logger.log('tasks.api.scrum.sprint.list warning for group ' + gid + ': ' + eSprint.message);
    }
  }

  return { stageMap: stageMap, sprintMap: sprintMap };
}

/** Legacy alias */
function getStageMapForGroups(groupIds) {
  return getScrumLookupMaps(groupIds).stageMap;
}

/**
 * Formats a Date object or date string into standard 'dd MMM yyyy' for writing to Google Sheets.
 * [Project Level] — keeps date only (no time component) across all sheets and UI.
 */
function formatDateForSheet(dateVal) {
  if (!dateVal) return '';
  var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);

  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var day = ('0' + d.getDate()).slice(-2);
  var month = months[d.getMonth()];
  var year = d.getFullYear();

  return day + ' ' + month + ' ' + year;
}

/**
 * Resolves Next Follow Up Date logic:
 * Next Follow Up Date = N_Date
 * If N_Date == "" OR N_Date < Today then Next Follow Up Date = Today
 * If N_Date >= Today then N_Date remains same
 * [Project Level]
 */
function resolveNextFollowUpDate(nDateVal) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  if (!nDateVal) {
    return formatDateForSheet(today);
  }

  var nDate = (nDateVal instanceof Date) ? nDateVal : new Date(nDateVal);
  if (isNaN(nDate.getTime())) {
    var parts = String(nDateVal).trim().split(' ');
    if (parts.length >= 3) {
      var months = { Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5, Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11 };
      var dayNum = parseInt(parts[0], 10);
      var monthNum = months[parts[1]];
      var yearNum = parseInt(parts[2], 10);
      if (!isNaN(dayNum) && monthNum !== undefined && !isNaN(yearNum)) {
        nDate = new Date(yearNum, monthNum, dayNum);
      }
    }
  }

  if (!nDate || isNaN(nDate.getTime())) {
    return formatDateForSheet(today);
  }

  var nDateZero = new Date(nDate.getFullYear(), nDate.getMonth(), nDate.getDate());
  if (nDateZero.getTime() < today.getTime()) {
    return formatDateForSheet(today);
  }

  return formatDateForSheet(nDateZero);
}

/**
 * Scans a sheet (e.g. MASTER_DATA_GID or CLOSED_REVIEW_GID) and updates any
 * 'Next Follow Up Date' that is empty or < Today to Today's date.
 * [Project Level]
 */
function refreshNextFollowUpDatesInSheet(sheetConfigKey) {
  try {
    var sheet = getSheetFromConfig(sheetConfigKey);
    var colNextFollowUp = getColumnByHeader(sheet, 'Next Follow Up Date');
    var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;
    var lastRow = getLastDataRow(sheet, 'Bitrix ID');

    if (lastRow < dataStartRow) return;

    var range = sheet.getRange(dataStartRow, colNextFollowUp, lastRow - dataStartRow + 1, 1);
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
      var currentVal = values[i][0];
      var resolvedVal = resolveNextFollowUpDate(currentVal);
      if (String(currentVal).trim() !== String(resolvedVal).trim()) {
        sheet.getRange(dataStartRow + i, colNextFollowUp).setValue(resolvedVal);
      }
    }
  } catch (e) {
    Logger.log('refreshNextFollowUpDatesInSheet error (' + sheetConfigKey + '): ' + e.message);
  }
}


// ============================================
// WEB APP & API ENTRY POINTS
// ============================================

/**
 * Serves the web app HTML or JSON API responses for GET requests.
 */
function doGet(e) {
  // Support API calls via GET (e.g. ?api=1&action=getInitialData)
  if (e && e.parameter && (e.parameter.api === '1' || e.parameter.action)) {
    var action = e.parameter.action || 'getInitialData';
    var authOverride = {
      telegramInitData: e.parameter.telegramInitData || ''
    };
    var responseData = {};
    if (action === 'getInitialData' || action === 'reloadSheet') {
      responseData = getInitialData(authOverride);
    } else if (action === 'getCurrentUserAuth') {
      responseData = { success: true, data: getCurrentUserAuth('', '', authOverride.telegramInitData) };
    } else {
      responseData = { success: false, error: 'Unknown GET action: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('Bitrix24 Task Review')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Unified JSON API router for Telegram Mini App and external web deployments.
 * Supports all CRUD operations.
 */
function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    // Direct Telegram Bot Webhook update (inline_query or message)
    if (payload.inline_query || payload.message) {
      var tgRes = handleTelegramWebhook_(payload);
      return ContentService.createTextOutput(JSON.stringify(tgRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var action = String(payload.action || 'getInitialData').trim();
    var authOverride = {
      telegramInitData: payload.telegramInitData || ''
    };

    var responseData = {};

    switch (action) {
      case 'getInitialData':
      case 'reloadSheet':
        responseData = getInitialData(authOverride);
        break;

      case 'confirmTask':
        responseData = confirmTask(payload.reviewData, authOverride);
        break;

      case 'skipTask':
        responseData = skipTask(payload.bitrixId, authOverride);
        break;

      case 'updateTask':
        responseData = updateTaskInSheet(payload.sheetGidKey, payload.taskData, authOverride);
        break;

      case 'moveTask':
        responseData = moveTaskBetweenSheets(payload.fromGidKey, payload.toGidKey, payload.bitrixId, payload.additionalUpdates, authOverride);
        break;

      case 'saveConfig':
        responseData = saveMasterConfig(payload.configData, authOverride);
        break;

      case 'syncBitrix':
        responseData = pullAndStageNewTasks(authOverride);
        break;

      case 'getCurrentUserAuth':
        responseData = { success: true, data: getCurrentUserAuth('', '', authOverride.telegramInitData) };
        break;

      default:
        responseData = { success: false, error: 'Unknown API action: ' + action };
        break;
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message || 'Internal server error in doPost.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Includes an HTML file (for templating).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Verifies Telegram Mini App initData and returns the trusted Telegram user.
 * Requires TELEGRAM_BOT_TOKEN in Apps Script Script Properties.
 */
function verifyTelegramInitData_(initData) {
  initData = String(initData || '').trim();
  if (!initData) return null;

  var botToken = String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  if (!botToken) throw new Error('Telegram bot token is not configured. Set TELEGRAM_BOT_TOKEN in Script Properties.');

  var pairs = initData.split('&');
  var values = {};
  var checkPairs = [];
  var receivedHash = '';

  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i];
    var eq = pair.indexOf('=');
    if (eq === -1) continue;

    var key = decodeURIComponent(pair.slice(0, eq));
    var value = decodeURIComponent(pair.slice(eq + 1).replace(/\+/g, '%20'));
    values[key] = value;

    if (key === 'hash') {
      receivedHash = value;
    } else {
      checkPairs.push(key + '=' + value);
    }
  }

  if (!receivedHash) throw new Error('Telegram init data is missing hash.');
  checkPairs.sort();

  var dataCheckString = checkPairs.join('\n');
  var secretKey = Utilities.computeHmacSha256Signature(botToken, 'WebAppData');
  var calculatedHash = bytesToHex_(Utilities.computeHmacSha256Signature(dataCheckString, secretKey));

  if (!constantTimeEqual_(calculatedHash, receivedHash.toLowerCase())) {
    throw new Error('Telegram login verification failed.');
  }

  var authDate = parseInt(values.auth_date || '0', 10);
  var maxAgeSeconds = 24 * 60 * 60;
  if (!authDate || Math.floor(Date.now() / 1000) - authDate > maxAgeSeconds) {
    throw new Error('Telegram login session expired. Please reopen the Mini App from Telegram.');
  }

  if (!values.user) throw new Error('Telegram init data is missing user.');
  return JSON.parse(values.user);
}

function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function constantTimeEqual_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;

  var mismatch = 0;
  for (var i = 0; i < a.length; i++) {
    mismatch |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return mismatch === 0;
}

/**
 * Handles Telegram Bot Webhook updates (messages, /start commands, inline queries).
 */
function handleTelegramWebhook_(update) {
  if (update.inline_query) {
    return handleTelegramInlineQuery_(update.inline_query);
  } else if (update.message) {
    return handleTelegramMessage_(update.message);
  }
  return { ok: true };
}

function handleTelegramMessage_(message) {
  var botToken = String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  if (!botToken) return { ok: false };

  var chatId = message.chat.id;
  var text = String(message.text || '').trim();
  var appUrl = String(getConfig('TELEGRAM_MINI_APP_URL', 'https://goutham.github.io/bitrix-tasks/')).trim();

  if (text.indexOf('/start') === 0 || text.indexOf('/tasks') === 0) {
    var welcomeText = '👋 *Welcome to Bitrix Task Manager!*\n\nTap the button below to review new tasks, view your client portfolio, and update technical remarks in real-time.';
    var payload = {
      chat_id: chatId,
      text: welcomeText,
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [
            {
              text: '🚀 Open Task Dashboard',
              web_app: { url: appUrl }
            }
          ]
        ]
      }
    };
    UrlFetchApp.fetch('https://api.telegram.org/bot' + botToken + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }
  return { ok: true };
}

function handleTelegramInlineQuery_(inlineQuery) {
  var botToken = String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  if (!botToken) return { ok: false };

  var queryId = inlineQuery.id;
  var q = String(inlineQuery.query || '').toLowerCase().trim();
  var appUrl = String(getConfig('TELEGRAM_MINI_APP_URL', 'https://goutham.github.io/bitrix-tasks/')).trim();

  var masterRes = getMasterSheetData();
  var newRes = getNewTaskReviewData();
  var master = masterRes.data || [];
  var newTasks = newRes.data || [];
  var all = master.concat(newTasks);

  var matches = all.filter(function(t) {
    if (!q) return true;
    var idStr = String(t['Bitrix ID'] || '').toLowerCase();
    var title = String(t['Task Title'] || '').toLowerCase();
    var client = String(t['Client Name'] || '').toLowerCase();
    var owner = String(t['Task Owner'] || '').toLowerCase();
    var contact = String(t['Contact Person'] || '').toLowerCase();
    return idStr.indexOf(q) >= 0 || title.indexOf(q) >= 0 || client.indexOf(q) >= 0 || owner.indexOf(q) >= 0 || contact.indexOf(q) >= 0;
  }).slice(0, 15);

  var results = matches.map(function(t, idx) {
    var idStr = String(t['Bitrix ID'] || '').trim();
    var title = String(t['Task Title'] || 'Untitled Task').trim();
    var client = String(t['Client Name'] || '-').trim();
    var owner = String(t['Task Owner'] || '-').trim();
    var contact = String(t['Contact Person'] || '-').trim();
    var deadline = String(t['Deadline Given'] || t['Deadline'] || '-').trim();
    var techRemarks = String(t['Tech Remarks'] || '-').trim();
    var followUp = String(t['Next Follow Up Date'] || '-').trim();
    var status = String(t['Current Status'] || t['Stage'] || t['Review Status'] || 'Active').trim();

    var msgText = 
      '📋 *Task #' + idStr + ': ' + title + '*\n' +
      '🏢 *Client:* ' + client + '\n' +
      '👤 *Owner:* ' + owner + (contact !== '-' ? ' | *Contact:* ' + contact : '') + '\n' +
      '⏰ *Deadline:* ' + deadline + ' | *Status:* ' + status + '\n' +
      (techRemarks !== '-' ? '🛠️ *Tech Remarks:* ' + techRemarks + '\n' : '') +
      (followUp !== '-' ? '📅 *Next Follow Up:* ' + followUp + '\n' : '');

    return {
      type: 'article',
      id: String(idStr || idx),
      title: '#' + idStr + ' - ' + title.substring(0, 45),
      description: client + ' | ' + owner + ' | ' + deadline,
      input_message_content: {
        message_text: msgText,
        parse_mode: 'Markdown'
      },
      reply_markup: {
        inline_keyboard: [
          [
            {
              text: '📱 Open Task in App',
              web_app: { url: appUrl }
            }
          ]
        ]
      }
    };
  });

  var payload = {
    inline_query_id: queryId,
    results: results,
    cache_time: 5,
    is_personal: true
  };

  UrlFetchApp.fetch('https://api.telegram.org/bot' + botToken + '/answerInlineQuery', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return { ok: true };
}

/**
 * Registers this Google Apps Script Web App URL as the Telegram Bot Webhook in one click.
 */
function setTelegramWebhook() {
  var botToken = String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  if (!botToken) {
    throw new Error('TELEGRAM_BOT_TOKEN is not set in Script Properties.');
  }
  var webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    throw new Error('Web App is not deployed yet. Deploy as Web App first.');
  }

  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + botToken + '/setWebhook?url=' + encodeURIComponent(webAppUrl), {
    muteHttpExceptions: true
  });
  Logger.log('setWebhook result: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}


// ============================================
// BITRIX24 API LAYER
// ============================================

/**
 * Calls a Bitrix24 REST API method via the inbound webhook.
 * @param {string} method — e.g. 'tasks.task.list'
 * @param {Object} params — request parameters
 * @returns {Object} parsed JSON response
 */
function callBitrixApi(method, params) {
  var webhookUrl = getConfig('BITRIX_WEBHOOK_URL');
  // Ensure trailing slash
  if (webhookUrl.charAt(webhookUrl.length - 1) !== '/') webhookUrl += '/';

  var url = webhookUrl + method + '.json';

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(params || {}),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code === 503) {
    throw new Error('Bitrix rate limit exceeded (503). Wait and retry.');
  }
  if (code === 429) {
    throw new Error('Bitrix operation time limit (429). Wait 10 minutes.');
  }
  if (code !== 200) {
    throw new Error('Bitrix API error: HTTP ' + code + ' — ' + response.getContentText().substring(0, 200));
  }

  return JSON.parse(response.getContentText());
}

/**
 * Fetches tasks from Bitrix with pagination.
 * Respects rate limits with 0.5s delay between calls.
 * @param {Object} filter — Bitrix filter object
 * @param {string[]} [select] — fields to return (defaults to all fields)
 * @returns {Object[]} all matching tasks
 */
function fetchBitrixTasks(filter, select) {
  var allTasks = [];
  var start = 0;
  var hasMore = true;
  var pageCount = 0;
  var MAX_PAGES = 50; // safety cap

  while (hasMore && pageCount < MAX_PAGES) {
    var params = {
      filter: filter || {},
      select: select || ['*', 'UF_*'],
      start: start
    };

    var result = callBitrixApi('tasks.task.list', params);

    if (result.result && result.result.tasks) {
      allTasks = allTasks.concat(result.result.tasks);
    }

    // Check for more pages
    if (result.next && result.next > start) {
      start = result.next;
      pageCount++;
      // Rate limit: 0.5s delay between paginated calls
      Utilities.sleep(500);
    } else {
      hasMore = false;
    }
  }

  return allTasks;
}

/**
 * Resolves Bitrix user IDs to display names.
 * @param {number[]} userIds — array of user IDs
 * @returns {Object} { userId: 'Full Name' }
 */
function resolveUserNames(userIds) {
  if (!userIds || userIds.length === 0) return {};

  // Deduplicate
  var unique = [];
  var seen = {};
  for (var i = 0; i < userIds.length; i++) {
    var id = String(userIds[i]);
    if (!seen[id]) {
      seen[id] = true;
      unique.push(id);
    }
  }

  var nameMap = {};

  // Bitrix user.get supports batch — but we'll do chunks of 50
  for (var start = 0; start < unique.length; start += 50) {
    var chunk = unique.slice(start, start + 50);
    var idFilter = {};
    for (var j = 0; j < chunk.length; j++) {
      idFilter[j] = chunk[j];
    }

    var result = callBitrixApi('user.get', { ID: chunk });

    if (result.result) {
      for (var k = 0; k < result.result.length; k++) {
        var user = result.result[k];
        nameMap[user.ID] = (user.NAME || '') + ' ' + (user.LAST_NAME || '');
        nameMap[user.ID] = nameMap[user.ID].trim();
      }
    }

    if (start + 50 < unique.length) Utilities.sleep(500);
  }

  return nameMap;
}


// ============================================
// DATA: READ OPERATIONS (called by frontend)
// ============================================

/**
 * Returns pending tasks from the New Task Review sheet.
 * Called by frontend via google.script.run
 */
function getNewTaskReviewData() {
  try {
    var sheet = getSheetFromConfig('NEW_TASK_REVIEW_GID');
    var allData = getSheetDataAsObjects(sheet);

    // Filter: only Pending tasks (or tasks not marked Confirmed/Skipped)
    var pending = allData.filter(function(row) {
      var rStatus = String(row['Review Status'] || '').trim();
      return rStatus === 'Pending' || rStatus === '' || rStatus === '-';
    });

    return { success: true, data: pending };
  } catch (e) {
    Logger.log('getNewTaskReviewData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Returns all active tasks from Master Data sheet.
 */
function getMasterSheetData() {
  try {
    var sheet = getSheetFromConfig('MASTER_DATA_GID');
    var allData = getSheetDataAsObjects(sheet);
    var valid = allData.filter(function(row) {
      return String(row['Bitrix ID'] || '').trim() !== '';
    });
    return { success: true, data: valid };
  } catch (e) {
    Logger.log('getMasterSheetData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Returns all completed tasks from Closed Review sheet.
 */
function getClosedSheetData() {
  try {
    var sheet = getSheetFromConfig('CLOSED_REVIEW_GID');
    var allData = getSheetDataAsObjects(sheet);
    var valid = allData.filter(function(row) {
      return String(row['Bitrix ID'] || '').trim() !== '';
    });
    return { success: true, data: valid };
  } catch (e) {
    Logger.log('getClosedSheetData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Returns all archived tasks from Archive sheet.
 */
function getArchiveSheetData() {
  try {
    var sheet = getSheetFromConfig('ARCHIVE_GID');
    var allData = getSheetDataAsObjects(sheet);
    var valid = allData.filter(function(row) {
      return String(row['Bitrix ID'] || '').trim() !== '';
    });
    return { success: true, data: valid };
  } catch (e) {
    Logger.log('getArchiveSheetData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Returns all Bitrix IDs currently in Master Data (for duplicate checking).
 */
function getMasterBitrixIds() {
  try {
    var sheet = getSheetFromConfig('MASTER_DATA_GID');
    var col = getColumnByHeader(sheet, 'Bitrix ID');
    var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;
    var lastRow = getLastDataRow(sheet, 'Bitrix ID');

    if (lastRow < dataStartRow) return [];

    var values = sheet.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1).getValues();
    var ids = [];
    for (var i = 0; i < values.length; i++) {
      var val = String(values[i][0]).trim();
      if (val) ids.push(val);
    }
    return ids;
  } catch (e) {
    Logger.log('getMasterBitrixIds error: ' + e.message);
    return [];
  }
}

/**
 * Returns active customers for the Client Name dropdown.
 * Uses "Short Code" as the display value and prepends "ALL COMPANY".
 * [Project Level]
 */
function getActiveCustomers() {
  try {
    var sheet = getSheetFromConfig('CUSTOMER_MASTER_GID');
    var allData = getSheetDataAsObjects(sheet);

    var active = allData.filter(function(row) {
      var st = String(row['Status'] || 'Active').trim().toLowerCase();
      return st === 'active' || st === '';
    }).map(function(row) {
      return {
        shortCode: String(row['Short Code'] || row['Client Name'] || '').trim(),
        clientName: String(row['Client Name'] || '').trim(),
        taskOwner: String(row['Task Owner'] || '').trim(),
        apm: String(row['APM'] || '').trim(),
        cse: String(row['CSE'] || '').trim(),
        weight: row['Weight'] || 1
      };
    }).filter(function(c) {
      return c.shortCode !== '' && c.shortCode !== 'ALL COMPANY';
    });

    // Prepend "ALL COMPANY" option (Task Owner: Goutham K)
    active.unshift({
      shortCode: 'ALL COMPANY',
      clientName: 'ALL COMPANY',
      taskOwner: 'Goutham K',
      apm: 'Goutham K',
      cse: 'Goutham K',
      weight: 0
    });

    return { success: true, data: active };
  } catch (e) {
    Logger.log('getActiveCustomers error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Detects the active user from Google Session and resolves their role from the Teammates sheet
 * and Customer Master APM/CSE hierarchy.
 */
function getCurrentUserAuth(overrideEmail, overrideTelegramId, telegramInitData) {
  try {
    var email = String(overrideEmail || '').trim().toLowerCase();
    var tgId = String(overrideTelegramId || '').trim();

    if (telegramInitData) {
      try {
        var telegramUser = verifyTelegramInitData_(telegramInitData);
        if (telegramUser && telegramUser.id) {
          tgId = String(telegramUser.id).trim();
        }
      } catch(tgErr) {
        Logger.log('Telegram verification warning: ' + tgErr.message);
      }
    }

    if (!email && !tgId) {
      try {
        email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
      } catch(e) {}
    }

    var superAdminEmail = String(getConfig('SUPER_ADMIN_EMAIL', 'goutham@eazyerp.com')).trim().toLowerCase();

    // Direct check for Goutham / Default Super Admin
    if ((!email && !tgId) || (email && (email.indexOf('goutham') >= 0 || (superAdminEmail && email === superAdminEmail)))) {
      return {
        email: email || 'goutham@eazyerp.com',
        name: 'Goutham K',
        role: 'Super Admin',
        team: 'Management',
        status: 'Active',
        isSuperAdmin: true,
        isAdmin: true,
        isManager: true,
        isBlocked: false
      };
    }

    var tmSheet = getSheetFromConfig('TEAMMATES_GID');
    var tmData = getSheetDataAsObjects(tmSheet);

    var custSheet = getSheetFromConfig('CUSTOMER_MASTER_GID');
    var custData = getSheetDataAsObjects(custSheet);

    // Find teammate by Telegram ID or Email or Name match
    var matched = null;
    for (var i = 0; i < tmData.length; i++) {
      var row = tmData[i];
      var rowTg = String(row['Telegram Chat ID'] || row['Telegram ID'] || row['Chat ID'] || '').trim();
      var rowEmail = String(row['Email'] || row['email'] || row['Mail'] || row['Gmail'] || '').trim().toLowerCase();
      var rowName = String(row['Name'] || '').trim().toLowerCase();

      if (tgId && rowTg && rowTg === tgId) {
        matched = row;
        break;
      }
      if (email && ((rowEmail && rowEmail === email) || (email.indexOf(rowName) >= 0 && rowName.length > 3))) {
        matched = row;
        break;
      }
    }

    if (!matched) {
      if (email && email.indexOf('goutham') >= 0) {
        return {
          email: email,
          name: 'Goutham K',
          role: 'Super Admin',
          team: 'Management',
          status: 'Active',
          isSuperAdmin: true,
          isAdmin: true,
          isManager: true,
          isBlocked: false
        };
      }

      return {
        email: email,
        name: email || (tgId ? 'Telegram ID: ' + tgId : 'Unlisted User'),
        role: 'Blocked',
        team: '',
        status: 'Unlisted',
        isSuperAdmin: false,
        isAdmin: false,
        isManager: false,
        isBlocked: true,
        message: 'Your account (' + (email || ('Telegram ID ' + tgId)) + ') is not registered in the Teammates sheet. Please contact Goutham to request access.'
      };
    }

    var status = String(matched['Status'] || 'Active').trim();
    var rawRole = String(matched['Role'] || '').trim();
    var name = String(matched['Name'] || email).trim();
    var team = String(matched['Team'] || '').trim();

    if (status.toLowerCase() === 'blocked' || status.toLowerCase() === 'inactive') {
      return {
        email: email,
        name: name,
        role: 'Blocked',
        team: team,
        status: status,
        isSuperAdmin: false,
        isAdmin: false,
        isManager: false,
        isBlocked: true,
        message: 'Your account access is currently ' + status + '. Please contact your administrator.'
      };
    }

    // Role resolution with APM / CSE Hierarchy
    var roleNorm = 'User';
    var isSuper = false;
    var isAdmin = false;
    var isManager = false;

    var nameLower = name.toLowerCase();
    var roleLower = rawRole.toLowerCase();

    if (nameLower.indexOf('goutham') >= 0 || roleLower.indexOf('super') >= 0 || roleLower.indexOf('owner') >= 0) {
      roleNorm = 'Super Admin';
      isSuper = true;
      isAdmin = true;
      isManager = true;
    } else if (roleLower.indexOf('admin') >= 0 || roleLower.indexOf('lead') >= 0) {
      roleNorm = 'Admin';
      isAdmin = true;
      isManager = true;
    } else {
      // Check if user is an APM / Task Owner for any client
      var isApm = custData.some(function(c) {
        var apm = String(c['APM'] || '').trim();
        var own = String(c['Task Owner'] || '').trim();
        return apm === name || own === name;
      });

      if (isApm || roleLower.indexOf('manager') >= 0 || roleLower.indexOf('apm') >= 0) {
        roleNorm = 'Manager (APM)';
        isManager = true;
        isAdmin = true;
      } else {
        roleNorm = 'User (CSE)';
      }
    }

    return {
      email: email,
      name: name,
      telegramId: tgId,
      bitrixUserId: matched['Bitrix User ID'] || '',
      role: roleNorm,
      rawRole: rawRole,
      team: team,
      department: String(matched['Department'] || '').trim(),
      status: status,
      isSuperAdmin: isSuper,
      isAdmin: isAdmin,
      isManager: isManager,
      isBlocked: false
    };
  } catch(e) {
    Logger.log('getCurrentUserAuth error: ' + e.message);
    return {
      email: '',
      name: 'Blocked User',
      role: 'Blocked',
      team: '',
      status: 'Auth Error',
      isSuperAdmin: false,
      isAdmin: false,
      isManager: false,
      isBlocked: true,
      message: e.message || 'Unable to verify your account.'
    };
  }
}

/**
 * Returns active teammates for the Contact Person dropdown.
 */
function getActiveTeammates() {
  try {
    var sheet = getSheetFromConfig('TEAMMATES_GID');
    var allData = getSheetDataAsObjects(sheet);

    var active = allData.filter(function(row) {
      var st = String(row['Status'] || '').trim().toLowerCase();
      return st === 'active' || st === '';
    }).map(function(row) {
      return {
        name: String(row['Name'] || '').trim(),
        email: String(row['Email'] || row['email'] || '').trim().toLowerCase(),
        bitrixUserId: row['Bitrix User ID'] || '',
        role: String(row['Role'] || 'User').trim(),
        department: String(row['Department'] || '').trim(),
        team: String(row['Team'] || '').trim(),
        status: String(row['Status'] || 'Active').trim()
      };
    }).filter(function(t) {
      return t.name !== '';
    });

    return { success: true, data: active };
  } catch (e) {
    Logger.log('getActiveTeammates error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Returns dropdown options for Platform, Task Type, Priority.
 * Reads from Config if available, otherwise returns defaults.
 */
function getDropdownOptions() {
  return {
    platforms: ['1.0', '2.0', 'Android 2.0', 'iPhone 2.0', 'Web Portal', 'Integration'],
    taskTypes: ['Issue', 'Customization', 'Understanding Gap'],
    priorities: [
      { value: 'Critical', color: '#dc2626' },
      { value: 'High', color: '#ea580c' },
      { value: 'Medium', color: '#d97706' },
      { value: 'Low', color: '#16a34a' }
    ]
  };
}

/**
 * Returns all data the frontend needs on initial load for all 4 views.
 * Single call instead of multiple round-trips.
 */
function getInitialData(authOverride) {
  try {
    var userAuth = null;
    if (authOverride && (authOverride.telegramInitData || authOverride.email || authOverride.telegramId)) {
      userAuth = getCurrentUserAuth(authOverride.email, authOverride.telegramId, authOverride.telegramInitData);
    } else {
      userAuth = getCurrentUserAuth();
    }
    if (userAuth.isBlocked) {
      return {
        success: true,
        userAuth: userAuth,
        newTasks: [],
        masterTasks: [],
        closedTasks: [],
        archiveTasks: [],
        customers: [],
        teammates: [],
        dropdowns: getDropdownOptions(),
        config: {}
      };
    }

    var newTasks = getNewTaskReviewData();
    var masterTasks = getMasterSheetData();
    var closedTasks = getClosedSheetData();
    var archiveTasks = getArchiveSheetData();
    var customers = getActiveCustomers();
    var teammates = getActiveTeammates();
    var dropdowns = getDropdownOptions();
    var masterConfig = getMasterConfig();

    return {
      success: true,
      userAuth: userAuth,
      newTasks: newTasks.success ? newTasks.data : [],
      masterTasks: masterTasks.success ? masterTasks.data : [],
      closedTasks: closedTasks.success ? closedTasks.data : [],
      archiveTasks: archiveTasks.success ? archiveTasks.data : [],
      customers: customers.success ? customers.data : [],
      teammates: teammates.success ? teammates.data : [],
      dropdowns: dropdowns,
      config: masterConfig.success ? masterConfig.config : {},
      errors: [
        newTasks.success ? null : 'New Tasks: ' + newTasks.error,
        masterTasks.success ? null : 'Master Tasks: ' + masterTasks.error,
        closedTasks.success ? null : 'Closed Tasks: ' + closedTasks.error,
        archiveTasks.success ? null : 'Archive Tasks: ' + archiveTasks.error,
        customers.success ? null : 'Customers: ' + customers.error,
        teammates.success ? null : 'Teammates: ' + teammates.error
      ].filter(Boolean)
    };
  } catch (e) {
    Logger.log('getInitialData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Retrieves master configuration for web app.
 */
function getMasterConfig() {
  try {
    var raw = getConfig('MANDATORY_FIELDS', '');
    if (raw) {
      return { success: true, config: { mandatoryFields: JSON.parse(raw) } };
    }
  } catch (e) {}
  return {
    success: true,
    config: {
      mandatoryFields: ['client', 'platform', 'type', 'priority']
    }
  };
}

/**
 * Saves master configuration to Config sheet.
 * Super Admin only.
 */
function saveMasterConfig(configObj, authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (!userAuth.isSuperAdmin) {
      throw new Error('Access denied: Only Super Admin can modify Master Configuration');
    }

    if (configObj && configObj.mandatoryFields) {
      setConfig('MANDATORY_FIELDS', JSON.stringify(configObj.mandatoryFields));
    }
    return { success: true, message: 'Configuration saved successfully' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ============================================
// DATA: WRITE OPERATIONS (called by frontend)
// ============================================

/**
 * Confirms a task from New Task Review → Master Data.
 * Super Admin & Admin only.
 * @param {Object} reviewData — task data with manual fields filled
 * @returns {Object} { success: boolean, message: string }
 */
function confirmTask(reviewData, authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (userAuth.isBlocked) throw new Error('Access denied: Your account is blocked or unregistered');
    if (!userAuth.isAdmin && !userAuth.isSuperAdmin) {
      throw new Error('Access denied: Only Admins can confirm tasks into Master Sheet');
    }

    var ntrSheet = getSheetFromConfig('NEW_TASK_REVIEW_GID');
    var masterSheet = getSheetFromConfig('MASTER_DATA_GID');

    var bitrixId = String(reviewData.bitrixId || reviewData['Bitrix ID'] || '').trim();
    if (!bitrixId) throw new Error('Bitrix ID is required');

    // Retrieve existing row from New Task Review for complete Bitrix properties
    var ntrData = getSheetDataAsObjects(ntrSheet);
    var existingNtrRow = {};
    for (var i = 0; i < ntrData.length; i++) {
      if (String(ntrData[i]['Bitrix ID']).trim() === bitrixId) {
        existingNtrRow = ntrData[i];
        break;
      }
    }

    // Validate required manual fields according to dynamic configuration
    var mandatoryKeys = ['Client Name', 'Platform', 'Task Type', 'Priority'];
    try {
      var rawMandatory = getConfig('MANDATORY_FIELDS', '');
      if (rawMandatory) {
        var parsed = JSON.parse(rawMandatory);
        var fieldKeyMap = {
          'client': 'Client Name',
          'platform': 'Platform',
          'type': 'Task Type',
          'priority': 'Priority',
          'gouthamRemarks': "Goutham's Remarks",
          'techRemarks': 'Tech Remarks',
          'nextFollowUp': 'Next Follow Up Date'
        };
        mandatoryKeys = parsed.map(function(k) { return fieldKeyMap[k] || k; });
      }
    } catch(e) {}

    for (var j = 0; j < mandatoryKeys.length; j++) {
      var field = mandatoryKeys[j];
      var val = reviewData[field] || existingNtrRow[field];
      if (!val || String(val).trim() === '' || String(val).trim() === '-') {
        throw new Error('"' + field + '" is required');
      }
    }

    function cleanVal(v1, v2) {
      if (v1 !== undefined && v1 !== null && String(v1).trim() !== '' && String(v1).trim() !== '-') {
        return String(v1).trim();
      }
      if (v2 !== undefined && v2 !== null && String(v2).trim() !== '' && String(v2).trim() !== '-') {
        return String(v2).trim();
      }
      return '';
    }

    var selectedClient = cleanVal(reviewData['Client Name'], existingNtrRow['Client Name']);
    var titleVal = cleanVal(reviewData['Task Title'], existingNtrRow['Task Title']);
    var statusVal = cleanVal(reviewData['Current Status'] || reviewData['Stage'], existingNtrRow['Current Status'] || existingNtrRow['Stage']);
    var priorityVal = cleanVal(reviewData['Priority'], existingNtrRow['Priority']);
    var taskTypeVal = cleanVal(reviewData['Task Type'], existingNtrRow['Task Type']);
    var platformVal = cleanVal(reviewData['Platform'], existingNtrRow['Platform']);
    var deadlineRaw = cleanVal(reviewData['Deadline Given'] || reviewData['Deadline'], existingNtrRow['Deadline Given'] || existingNtrRow['Deadline']);
    var deadlineVal = formatDateForSheet(deadlineRaw);
    var createdRaw = cleanVal(reviewData['Created Date'], existingNtrRow['Created Date']);
    var createdVal = formatDateForSheet(createdRaw);
    var rawNextFollow = cleanVal(reviewData['Next Follow Up Date'], existingNtrRow['Next Follow Up Date']);
    var nextFollowVal = resolveNextFollowUpDate(rawNextFollow);
    var projectVal = cleanVal(reviewData['Project'], existingNtrRow['Project']);
    var sprintVal = cleanVal(reviewData['Sprint'], existingNtrRow['Sprint']);
    var backlogVal = cleanVal(reviewData['Backlog'], existingNtrRow['Backlog']);
    var parentIdVal = cleanVal(reviewData['Parent ID'] || reviewData['parentId'], existingNtrRow['Parent ID'] || existingNtrRow['Parent Task ID']);
    var chatIdVal = cleanVal(reviewData['Chat ID'], existingNtrRow['Chat ID']);
    var descVal = cleanVal(reviewData['Description'], existingNtrRow['Description']);
    var remarksVal = cleanVal(reviewData["Goutham's Remarks"], existingNtrRow["Goutham's Remarks"]);
    var techRemarksVal = cleanVal(reviewData['Tech Remarks'], existingNtrRow['Tech Remarks']);

    // Build Master Data row object with formatted values for all possible column names
    var masterRow = {
      'Bitrix ID': bitrixId,
      'Client Name': selectedClient,
      'Task Title': titleVal,
      'Current Status': statusVal,
      'Stage': statusVal,
      'Priority': priorityVal,
      'My Priority': cleanVal(reviewData['My Priority'], existingNtrRow['My Priority']),
      "Goutham's Remarks": remarksVal,
      'Tech Remarks': techRemarksVal,
      'Contact Person': '',
      'Next Follow Up Date': nextFollowVal,
      'Deadline Given': deadlineVal,
      'Deadline': deadlineVal,
      'Task Owner': '',
      'Created Date': createdVal,
      'Project': projectVal,
      'Sprint': sprintVal,
      'Task Type': taskTypeVal,
      'Platform': platformVal,
      'Team': cleanVal(reviewData['Team'], existingNtrRow['Team']),
      'Parent ID': parentIdVal,
      'Parent Task ID': parentIdVal,
      'Chat ID': chatIdVal,
      'Backlog': backlogVal,
      'Description': descVal,
      'Total Points': cleanVal(reviewData['Total Points'], existingNtrRow['Total Points']),
      'Chain': cleanVal(reviewData['Chain'], existingNtrRow['Chain']),
      'Is Exists?': 'Yes',
      'Closed By': '',
      'Closed Date': '',
      'Review Status': 'Confirmed'
    };

    // Check if task already exists in Master Data — update in-place or append new
    var masterData = getSheetDataAsObjects(masterSheet);
    var existingMasterRowIndex = -1;
    var dataStartRow = parseInt(getConfig('DATA_START_ROW', '2')) || 2;
    for (var m = 0; m < masterData.length; m++) {
      if (String(masterData[m]['Bitrix ID']).trim() === bitrixId) {
        existingMasterRowIndex = dataStartRow + m;
        break;
      }
    }

    if (existingMasterRowIndex > 0) {
      updateRowAsObject(masterSheet, existingMasterRowIndex, masterRow);
    } else {
      appendRowAsObject(masterSheet, masterRow);
    }

    // Update Review Status in NTR to "Confirmed"
    updateReviewStatus_(ntrSheet, bitrixId, 'Confirmed');

    return { success: true, message: 'Task ' + bitrixId + ' added to Master Data' };
  } catch (e) {
    Logger.log('confirmTask error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Updates a task's manual fields in-place in any of the 4 sheets.
 * Super Admin & Admin can edit all fields.
 * Regular Users can only update Tech Remarks & Next Follow Up Date.
 * @param {string} sheetConfigKey - e.g. 'MASTER_DATA_GID', 'CLOSED_REVIEW_GID', 'ARCHIVE_GID'
 * @param {Object} updateData - updated task fields
 */
function updateTaskInSheet(sheetConfigKey, updateData, authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (userAuth.isBlocked) throw new Error('Access denied: Your account is blocked or unregistered');

    var sheet = getSheetFromConfig(sheetConfigKey);
    var bitrixId = String(updateData.bitrixId || updateData['Bitrix ID'] || '').trim();
    if (!bitrixId) throw new Error('Bitrix ID is required');

    var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;
    var colBid = getColumnByHeader(sheet, 'Bitrix ID');
    var lastRow = getLastDataRow(sheet, 'Bitrix ID');

    if (lastRow < dataStartRow) throw new Error('No data rows in sheet');

    var bidValues = sheet.getRange(dataStartRow, colBid, lastRow - dataStartRow + 1, 1).getValues();
    var targetRowIndex = -1;

    for (var i = 0; i < bidValues.length; i++) {
      if (String(bidValues[i][0]).trim() === bitrixId) {
        targetRowIndex = dataStartRow + i;
        break;
      }
    }

    if (targetRowIndex === -1) throw new Error('Task ' + bitrixId + ' not found in sheet');

    // Build row payload with role-based field restrictions
    var rowPayload = {};

    if (userAuth.isAdmin || userAuth.isSuperAdmin) {
      if (updateData['Client Name'] !== undefined) rowPayload['Client Name'] = updateData['Client Name'];
      if (updateData['Platform'] !== undefined) rowPayload['Platform'] = updateData['Platform'];
      if (updateData['Task Type'] !== undefined) rowPayload['Task Type'] = updateData['Task Type'];
      if (updateData['Priority'] !== undefined) rowPayload['Priority'] = updateData['Priority'];
      if (updateData["Goutham's Remarks"] !== undefined) rowPayload["Goutham's Remarks"] = updateData["Goutham's Remarks"];
    }

    // Both Admins and Users can update Tech Remarks & Follow Up Date
    if (updateData['Tech Remarks'] !== undefined) rowPayload['Tech Remarks'] = updateData['Tech Remarks'];
    if (updateData['Next Follow Up Date'] !== undefined) {
      rowPayload['Next Follow Up Date'] = resolveNextFollowUpDate(updateData['Next Follow Up Date']);
    }

    writeRowSafe_(sheet, targetRowIndex, rowPayload);

    return { success: true, message: 'Task updated successfully' };
  } catch (e) {
    Logger.log('updateTaskInSheet error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Moves a task between sheets (e.g. Master -> Closed, Master -> Archive, Closed -> Master, etc.)
 * Super Admin & Admin only.
 */
function moveTaskBetweenSheets(fromSheetKey, toSheetKey, bitrixId, updatedData, authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (userAuth.isBlocked) throw new Error('Access denied: Your account is blocked or unregistered');
    if (!userAuth.isAdmin && !userAuth.isSuperAdmin) {
      throw new Error('Access denied: Only Admins can move tasks between status sheets');
    }

    var fromSheet = getSheetFromConfig(fromSheetKey);
    var toSheet = getSheetFromConfig(toSheetKey);

    var bId = String(bitrixId || '').trim();
    if (!bId) throw new Error('Bitrix ID is required');

    var fromData = getSheetDataAsObjects(fromSheet);
    var existingRow = null;
    var fromRowIndex = -1;
    var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;

    for (var i = 0; i < fromData.length; i++) {
      if (String(fromData[i]['Bitrix ID']).trim() === bId) {
        existingRow = fromData[i];
        fromRowIndex = dataStartRow + i;
        break;
      }
    }

    if (!existingRow) throw new Error('Task ' + bId + ' not found in source sheet');

    // Merge updated manual fields into existing row data
    var merged = {};
    for (var k in existingRow) {
      if (existingRow.hasOwnProperty(k)) merged[k] = existingRow[k];
    }
    if (updatedData) {
      for (var uk in updatedData) {
        if (updatedData.hasOwnProperty(uk)) merged[uk] = updatedData[uk];
      }
    }

    // Format dates
    if (merged['Next Follow Up Date']) merged['Next Follow Up Date'] = resolveNextFollowUpDate(merged['Next Follow Up Date']);
    if (merged['Deadline Given']) merged['Deadline Given'] = formatDateForSheet(merged['Deadline Given']);
    if (merged['Created Date']) merged['Created Date'] = formatDateForSheet(merged['Created Date']);

    // Append to destination sheet
    appendRowAsObject(toSheet, merged);

    // Delete row from source sheet
    fromSheet.deleteRow(fromRowIndex);

    return { success: true, message: 'Task moved successfully' };
  } catch (e) {
    Logger.log('moveTaskBetweenSheets error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Stages tasks from Master Data that are completed / closed in Bitrix into the Closed Review sheet.
 * [Project Level]
 */
function stageClosedTasksIntoClosedReview() {
  try {
    var masterSheet = getSheetFromConfig('MASTER_DATA_GID');
    var closedSheet = getSheetFromConfig('CLOSED_REVIEW_GID');
    var masterData = getSheetDataAsObjects(masterSheet);
    var closedData = getSheetDataAsObjects(closedSheet);

    var closedExistingMap = {};
    closedData.forEach(function(row) {
      var bid = String(row['Bitrix ID']).trim();
      if (bid) closedExistingMap[bid] = true;
    });

    var masterBitrixIds = masterData.map(function(r) { return String(r['Bitrix ID']).trim(); }).filter(Boolean);
    if (masterBitrixIds.length === 0) return { success: true, stagedCount: 0 };

    var closedCount = 0;
    for (var i = 0; i < masterBitrixIds.length; i += 50) {
      var chunkIds = masterBitrixIds.slice(i, i + 50);
      var fetched = fetchBitrixTasks({ ID: chunkIds });
      var fetchedMap = {};
      fetched.forEach(function(t) { fetchedMap[String(t.id)] = t; });

      for (var j = 0; j < chunkIds.length; j++) {
        var idStr = chunkIds[j];
        if (closedExistingMap[idStr]) continue;

        var bTask = fetchedMap[idStr];
        var isClosedInBitrix = false;
        var closedByName = '';
        var closedDateVal = '';

        if (bTask) {
          var statusNum = parseInt(bTask.status, 10);
          if (statusNum === 5) {
            isClosedInBitrix = true;
            closedByName = bTask.closedBy || '';
            closedDateVal = formatDateForSheet(bTask.closedDate || new Date());
          }
        }

        if (isClosedInBitrix) {
          var mRow = masterData.find(function(r) { return String(r['Bitrix ID']).trim() === idStr; }) || {};
          var closedRow = {
            'Bitrix ID': idStr,
            'Task Title': mRow['Task Title'] || (bTask ? bTask.title : ''),
            'Client Name': mRow['Client Name'] || '',
            'Last Known Status': mRow['Current Status'] || mRow['Stage'] || 'Completed',
            'Last Known Stage': mRow['Current Status'] || mRow['Stage'] || 'Completed',
            'Deadline': mRow['Deadline Given'] || mRow['Deadline'] || (bTask ? formatDateForSheet(bTask.deadline) : ''),
            'Contact Person': mRow['Contact Person'] || '',
            'Closed By': closedByName,
            'Closed Date': closedDateVal,
            'Verification Note': 'Closed in Bitrix',
            'Review Status': 'Pending'
          };
          appendRowAsObject(closedSheet, closedRow);
          closedExistingMap[idStr] = true;
          closedCount++;
        }
      }
      if (i + 50 < masterBitrixIds.length) Utilities.sleep(500);
    }

    return { success: true, stagedCount: closedCount };
  } catch (e) {
    Logger.log('stageClosedTasksIntoClosedReview error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Skips a task in New Task Review (marks as Skipped).
 * @param {string} bitrixId
 */
function skipTask(bitrixId, authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (userAuth.isBlocked) throw new Error('Access denied: Your account is blocked or unregistered');
    if (!userAuth.isAdmin && !userAuth.isSuperAdmin) {
      throw new Error('Access denied: Only Admins can skip tasks in New Task Review');
    }

    var ntrSheet = getSheetFromConfig('NEW_TASK_REVIEW_GID');
    updateReviewStatus_(ntrSheet, String(bitrixId).trim(), 'Skipped');
    return { success: true, message: 'Task ' + bitrixId + ' skipped' };
  } catch (e) {
    Logger.log('skipTask error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Updates the Review Status column for a given Bitrix ID in a sheet.
 */
function updateReviewStatus_(sheet, bitrixId, newStatus) {
  var bidCol = getColumnByHeader(sheet, 'Bitrix ID');
  var statusCol = getColumnByHeader(sheet, 'Review Status');
  var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;
  var lastRow = getLastDataRow(sheet, 'Bitrix ID');

  if (lastRow < dataStartRow) return;

  var bidValues = sheet.getRange(dataStartRow, bidCol, lastRow - dataStartRow + 1, 1).getValues();

  for (var i = 0; i < bidValues.length; i++) {
    if (String(bidValues[i][0]).trim() === bitrixId) {
      sheet.getRange(dataStartRow + i, statusCol).setValue(newStatus);
      return;
    }
  }
}


// ============================================
// BITRIX SYNC: Pull & Stage New Tasks
// ============================================

/**
 * Pulls tasks from Bitrix24 and stages new ones into the New Task Review sheet.
 * Compares against Master Data to identify truly new tasks.
 * 
 * Uses the PROVEN 3-pass approach from the old working code:
 *   Pass 1: CREATED_BY = user (server-side filter — works)
 *   Pass 2: RESPONSIBLE_ID = user (server-side filter — works)
 *   Pass 3: Scan by GROUP_IDs, check accomplices/auditors client-side
 *           (because Bitrix silently IGNORES these filters server-side)
 *   Then: filter by included statuses CLIENT-SIDE
 */
function pullAndStageNewTasks(authOverride) {
  try {
    var userAuth = authOverride ? getCurrentUserAuth('', '', authOverride.telegramInitData) : getCurrentUserAuth();
    if (userAuth.isBlocked) throw new Error('Access denied: Your account is blocked or unregistered');
    if (!userAuth.isAdmin && !userAuth.isSuperAdmin) {
      throw new Error('Access denied: Only Admins can trigger Bitrix Sync');
    }

    var userId = getConfig('BITRIX_USER_ID');
    var statusInclude = getConfig('BITRIX_STATUS_INCLUDE').split(',').map(function(s) { return parseInt(s.trim()); });
    var groupIds = getConfig('BITRIX_GROUP_IDS').split(',').map(function(s) { return parseInt(s.trim()); });

    Logger.log('=== pullAndStageNewTasks started ===');
    Logger.log('User ID: ' + userId);
    Logger.log('Status include: [' + statusInclude.join(',') + ']');
    Logger.log('Group IDs: [' + groupIds.join(',') + ']');

    // Get existing Bitrix IDs from Master Data and New Task Review
    var masterIds = getMasterBitrixIds();
    var ntrSheet = getSheetFromConfig('NEW_TASK_REVIEW_GID');
    var ntrData = getSheetDataAsObjects(ntrSheet);
    var ntrIds = ntrData.map(function(row) { return String(row['Bitrix ID']).trim(); });

    var existingIds = {};
    masterIds.forEach(function(id) { existingIds[id] = true; });
    ntrIds.forEach(function(id) { existingIds[id] = true; });
    Logger.log('Existing IDs (Master + NTR): ' + (masterIds.length + ntrIds.length));

    // Dedupe map
    var taskMap = {};

    // --- Pass 1: CREATED_BY (server-side filter — works) ---
    Logger.log('--- Pass 1: CREATED_BY = ' + userId + ' ---');
    var createdTasks = fetchBitrixTasks({ CREATED_BY: parseInt(userId) });
    Logger.log('Pass 1 returned: ' + createdTasks.length + ' tasks');
    createdTasks.forEach(function(t) { taskMap[t.id] = t; });

    // --- Pass 2: RESPONSIBLE_ID (server-side filter — works) ---
    Logger.log('--- Pass 2: RESPONSIBLE_ID = ' + userId + ' ---');
    var responsibleTasks = fetchBitrixTasks({ RESPONSIBLE_ID: parseInt(userId) });
    Logger.log('Pass 2 returned: ' + responsibleTasks.length + ' tasks');
    responsibleTasks.forEach(function(t) { taskMap[t.id] = t; });

    // --- Pass 3: Group scan for Participant/Observer (client-side check) ---
    Logger.log('--- Pass 3: Scanning groups for Participant/Observer ---');
    for (var g = 0; g < groupIds.length; g++) {
      var gid = groupIds[g];
      if (gid === 0) continue; // skip invalid group ID

      var groupTasks = fetchBitrixTasks({ GROUP_ID: gid });
      Logger.log('Group ' + gid + ': ' + groupTasks.length + ' total tasks scanned');

      var matchedCount = 0;
      for (var t = 0; t < groupTasks.length; t++) {
        var task = groupTasks[t];

        // Already have this task from Pass 1 or 2
        if (taskMap[task.id]) continue;

        // Check accomplices (participants) and auditors (observers)
        var accomplices = Array.isArray(task.accomplices) ? task.accomplices.map(String) : [];
        var auditors = Array.isArray(task.auditors) ? task.auditors.map(String) : [];
        var isParticipant = accomplices.indexOf(String(userId)) >= 0;
        var isObserver = auditors.indexOf(String(userId)) >= 0;

        if (isParticipant || isObserver) {
          taskMap[task.id] = task;
          matchedCount++;
        }
      }
      Logger.log('Group ' + gid + ': ' + matchedCount + ' matched Participant/Observer');

      if (g < groupIds.length - 1) Utilities.sleep(500);
    }

    // Merge all unique tasks
    var merged = [];
    for (var id in taskMap) {
      if (taskMap.hasOwnProperty(id)) merged.push(taskMap[id]);
    }
    Logger.log('Merged unique tasks (all roles): ' + merged.length);

    // --- Client-side status filter ---
    // Exclude only Completed (5) and Deferred (6).
    // Meta-statuses like -1 (overdue), -2 (not viewed), -3 (almost overdue)
    // are still active tasks and should be included.
    var EXCLUDED_STATUSES = [5, 6];
    var statusFiltered = merged.filter(function(t) {
      var status = parseInt(t.status, 10);
      return EXCLUDED_STATUSES.indexOf(status) === -1;
    });
    Logger.log('After status filter (excluding 5,6): ' + statusFiltered.length + ' tasks');

    // Get Master Data IDs
    var masterIds = getMasterBitrixIds();
    var masterIdsMap = {};
    masterIds.forEach(function(id) { masterIdsMap[id] = true; });

    // Map existing rows in New Task Review sheet
    var ntrSheet = getSheetFromConfig('NEW_TASK_REVIEW_GID');
    var ntrData = getSheetDataAsObjects(ntrSheet);
    var dataStartRow = parseInt(getConfig('DATA_START_ROW')) || 2;

    var pendingNtrMap = {};   // bitrixId -> 1-based sheet row index
    var completedNtrMap = {}; // bitrixId -> true

    for (var r = 0; r < ntrData.length; r++) {
      var bId = String(ntrData[r]['Bitrix ID']).trim();
      var rStatus = String(ntrData[r]['Review Status'] || '').trim();
      if (bId) {
        if (rStatus === 'Pending' || rStatus === '') {
          pendingNtrMap[bId] = dataStartRow + r;
        } else {
          completedNtrMap[bId] = true;
        }
      }
    }

    // Filter tasks: exclude tasks in Master Data or Confirmed/Skipped in NTR
    var tasksToProcess = statusFiltered.filter(function(task) {
      var tid = String(task.id);
      return !masterIdsMap[tid] && !completedNtrMap[tid];
    });

    Logger.log('Tasks to process (new + pending updates): ' + tasksToProcess.length);

    if (tasksToProcess.length === 0) {
      Logger.log('No new or updated tasks found.');
      return { success: true, newCount: 0, updatedCount: 0, message: 'All tasks up to date.' };
    }

    // Resolve user names for creators and assignees
    var userIdsToResolve = [];
    tasksToProcess.forEach(function(task) {
      if (task.createdBy) userIdsToResolve.push(task.createdBy);
      if (task.responsibleId) userIdsToResolve.push(task.responsibleId);
    });
    var nameMap = resolveUserNames(userIdsToResolve);

    // Resolve group IDs to project names (if available in config)
    var groupNameMap = {};
    try {
      var groupNamesConfig = getConfig('BITRIX_GROUP_NAMES');
      if (groupNamesConfig) {
        var pairs = groupNamesConfig.split(';');
        pairs.forEach(function(pair) {
          var parts = pair.split(':');
          if (parts.length === 2) groupNameMap[parts[0].trim()] = parts[1].trim();
        });
      }
    } catch (e) { /* BITRIX_GROUP_NAMES not configured, skip */ }

    // Fetch stage & sprint titles for involved groups
    var scrumMaps = getScrumLookupMaps(groupIds);
    var stageMap = scrumMaps.stageMap;
    var sprintMap = scrumMaps.sprintMap;

    var now = new Date();
       // Safe lookup for column indices in NTR sheet
    var colTitle = null, colDeadlineGiven = null, colOwner = null, colContact = null, colCurrentStatus = null, colProject = null, colSprint = null, colBacklog = null, colParentId = null, colCreated = null, colChatId = null, colNextFollowUp = null;
    try { colTitle = getColumnByHeader(ntrSheet, 'Task Title'); } catch(e) {}
    try { colDeadlineGiven = getColumnByHeader(ntrSheet, 'Deadline Given'); } catch(e) {
      try { colDeadlineGiven = getColumnByHeader(ntrSheet, 'Deadline'); } catch(e2) {}
    }
    try { colOwner = getColumnByHeader(ntrSheet, 'Task Owner'); } catch(e) {}
    try { colContact = getColumnByHeader(ntrSheet, 'Contact Person'); } catch(e) {
      try { colContact = getColumnByHeader(ntrSheet, 'Assignee'); } catch(e2) {}
    }
    try { colCurrentStatus = getColumnByHeader(ntrSheet, 'Current Status'); } catch(e) {
      try { colCurrentStatus = getColumnByHeader(ntrSheet, 'Stage'); } catch(e2) {}
    }
    try { colProject = getColumnByHeader(ntrSheet, 'Project'); } catch(e) {}
    try { colSprint = getColumnByHeader(ntrSheet, 'Sprint'); } catch(e) {}
    try { colBacklog = getColumnByHeader(ntrSheet, 'Backlog'); } catch(e) {}
    try { colParentId = getColumnByHeader(ntrSheet, 'Parent ID'); } catch(e) {}
    try { colCreated = getColumnByHeader(ntrSheet, 'Created Date'); } catch(e) {}
    try { colChatId = getColumnByHeader(ntrSheet, 'Chat ID'); } catch(e) {}
    try { colNextFollowUp = getColumnByHeader(ntrSheet, 'Next Follow Up Date'); } catch(e) {}

    var newCount = 0;
    var updatedCount = 0;

    for (var n = 0; n < tasksToProcess.length; n++) {
      var nt = tasksToProcess[n];
      var tid = String(nt.id);
      var groupName = getGroupName(nt.groupId, groupNameMap);
      var stageName = stageMap[String(nt.stageId)] || stageMap[String(nt.epicId)] || stageMap[String(nt.epic)] || stageMap[String(nt.stage)] || '';

      var sprintId = nt.sprintId || nt.sprint;
      var sprintName = (sprintId && sprintMap[String(sprintId)]) ? sprintMap[String(sprintId)] : (sprintId ? ('Sprint ' + sprintId) : '-');

      var isBacklog = 'No';
      if (nt.backlogId || (nt.groupId > 0 && (!sprintId || parseInt(sprintId) === 0))) {
        isBacklog = 'Yes';
      } else if (!nt.groupId || parseInt(nt.groupId) === 0) {
        isBacklog = '-';
      }

      var parentIdRaw = nt.parentId || nt.parent || nt.PARENT_ID;
      var parentIdVal = (parentIdRaw && String(parentIdRaw) !== '0') ? String(parentIdRaw) : '';

      var titleVal = nt.title || '';
      var deadlineVal = formatDateForSheet(nt.deadline);
      var ownerVal = nameMap[nt.createdBy] || nt.createdBy || '';
      var assigneeVal = nameMap[nt.responsibleId] || nt.responsibleId || '';
      var createdVal = formatDateForSheet(nt.createdDate);
      var chatIdVal = nt.chatId || '';

      if (pendingNtrMap[tid]) {
        // UPDATE existing Pending row with fresh Bitrix details (Contact Person & Task Owner left untouched for formulas)
        var rowIndex = pendingNtrMap[tid];
        if (colTitle) ntrSheet.getRange(rowIndex, colTitle).setValue(titleVal);
        if (colDeadlineGiven) ntrSheet.getRange(rowIndex, colDeadlineGiven).setValue(deadlineVal);
        if (colCurrentStatus) ntrSheet.getRange(rowIndex, colCurrentStatus).setValue(stageName);
        if (colProject) ntrSheet.getRange(rowIndex, colProject).setValue(groupName);
        if (colSprint) ntrSheet.getRange(rowIndex, colSprint).setValue(sprintName);
        if (colBacklog) ntrSheet.getRange(rowIndex, colBacklog).setValue(isBacklog);
        if (colParentId) ntrSheet.getRange(rowIndex, colParentId).setValue(parentIdVal);
        if (colCreated) ntrSheet.getRange(rowIndex, colCreated).setValue(createdVal);
        if (colChatId) ntrSheet.getRange(rowIndex, colChatId).setValue(chatIdVal);

        // Evaluate Next Follow Up Date logic on every run
        if (colNextFollowUp) {
          var existingNextFollow = ntrSheet.getRange(rowIndex, colNextFollowUp).getValue();
          var updatedNextFollow = resolveNextFollowUpDate(existingNextFollow);
          ntrSheet.getRange(rowIndex, colNextFollowUp).setValue(updatedNextFollow);
        }
        updatedCount++;
      } else {
        // APPEND brand new row (26-column unified schema)
        // Contact Person and Task Owner left empty for Google Sheet formulas
        var ntrRow = {
          'Bitrix ID': tid,
          'Client Name': '',
          'Task Title': titleVal,
          'Current Status': stageName,
          'Priority': '',
          'My Priority': '',
          "Goutham's Remarks": '',
          'Tech Remarks': '',
          'Contact Person': '',
          'Next Follow Up Date': resolveNextFollowUpDate(''),
          'Deadline Given': deadlineVal,
          'Task Owner': '',
          'Created Date': createdVal,
          'Project': groupName,
          'Sprint': sprintName,
          'Task Type': '',
          'Platform': '',
          'Team': '',
          'Parent ID': parentIdVal,
          'Chat ID': chatIdVal,
          'Backlog': isBacklog,
          'Total Points': '',
          'Closed By': '',
          'Closed Date': '',
          'Review Status': 'Pending'
        };
        appendRowAsObject(ntrSheet, ntrRow);
        newCount++;
      }
    }

    // Map active task IDs returned from Bitrix
    var activeBitrixIdsMap = {};
    statusFiltered.forEach(function(t) { activeBitrixIdsMap[String(t.id)] = true; });

    // Auto-close tasks in New Task Review if they are no longer active in Bitrix
    var autoClosedCount = 0;
    var colReviewStatus = getColumnByHeader(ntrSheet, 'Review Status');
    for (var pId in pendingNtrMap) {
      if (!activeBitrixIdsMap[pId] && !masterIdsMap[pId]) {
        var closedRowIdx = pendingNtrMap[pId];
        ntrSheet.getRange(closedRowIdx, colReviewStatus).setValue('Closed in Bitrix');
        autoClosedCount++;
      }
    }

    var msg = [];
    if (newCount > 0) msg.push(newCount + ' new task(s) added');
    if (updatedCount > 0) msg.push(updatedCount + ' pending task(s) updated');
    if (autoClosedCount > 0) msg.push(autoClosedCount + ' task(s) marked "Closed in Bitrix"');
    var messageText = msg.length > 0 ? msg.join(', ') + '.' : 'All tasks up to date.';

    Logger.log('Done: ' + messageText);
    // Stage any closed tasks from Master Data into Closed Review
    try {
      var closedRes = stageClosedTasksIntoClosedReview();
      if (closedRes && closedRes.stagedCount > 0) {
        Logger.log('Staged ' + closedRes.stagedCount + ' closed task(s) into Closed Review');
      }
    } catch (eClosed) {
      Logger.log('stageClosedTasks warning: ' + eClosed.message);
    }

    // Refresh Next Follow Up Date in Master Data and Closed Review sheets on every sync run
    refreshNextFollowUpDatesInSheet('MASTER_DATA_GID');
    refreshNextFollowUpDatesInSheet('CLOSED_REVIEW_GID');

    return {
      success: true,
      newCount: newCount,
      updatedCount: updatedCount,
      autoClosedCount: autoClosedCount,
      message: 'Pulled ' + newCount + ' new tasks, updated ' + updatedCount + ' existing pending tasks, auto-closed ' + autoClosedCount + ' tasks in Review, and refreshed Next Follow Up dates across Master & Closed sheets.'
    };
  } catch (e) {
    Logger.log('pullAndStageNewTasks error: ' + e.message);
    Logger.log(e.stack);
    return { success: false, error: e.message };
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
    .addItem('🔄 Pull New Tasks from Bitrix', 'pullAndStageNewTasksMenu')
    .addSeparator()
    .addItem('🌐 Open Web App', 'openWebApp')
    .addItem('ℹ️ Show Config', 'showConfig')
    .addToUi();
}

/**
 * Menu wrapper for pullAndStageNewTasks (shows UI feedback).
 */
function pullAndStageNewTasksMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('🔄 Pulling tasks...', 'Fetching new tasks from Bitrix24.\nThis may take a moment due to API rate limits.', ui.ButtonSet.OK);

  var result = pullAndStageNewTasks();

  if (result.success) {
    ui.alert('✅ Pull Complete', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Pull Failed', result.error, ui.ButtonSet.OK);
  }
}

/**
 * Opens the deployed web app URL.
 */
function openWebApp() {
  var url = ScriptApp.getService().getUrl();
  if (url) {
    var html = HtmlService.createHtmlOutput(
      '<script>window.open("' + url + '", "_blank");google.script.host.close();</script>'
    ).setWidth(200).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening Web App...');
  } else {
    SpreadsheetApp.getUi().alert('Web app not deployed yet.\n\nGo to Deploy → New deployment → Web app.');
  }
}

/**
 * Shows current config in a dialog (masks sensitive values).
 */
function showConfig() {
  var sheet = getSheetByGid(0);
  var data = sheet.getDataRange().getValues();
  var lines = [];
  for (var i = 1; i < data.length; i++) {
    var val = String(data[i][1]);
    if (String(data[i][0]).indexOf('WEBHOOK') >= 0 && val.length > 20) {
      val = val.substring(0, 40) + '...***';
    }
    lines.push(data[i][0] + ':  ' + val);
  }
  SpreadsheetApp.getUi().alert('⚙️ Current Configuration', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
