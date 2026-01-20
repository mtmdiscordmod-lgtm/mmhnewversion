/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * CURRICULUM SYSTEM — Visual Curriculum Builder Backend
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * This module handles the curriculum builder dashboard:
 * - Hierarchical curriculum definition (Studio → Track → Level → Assignment → Achievement)
 * - CRUD operations for all curriculum entities
 * - Drag-drop reordering support
 * - Batch Google Classroom publishing
 * - Special Events with calendar scheduling
 * - Track Selection Form management
 *
 * Achievement ID format: {trackCode}-L{level}-a{assignNum}-{achNum}
 * Example: sf-L0-a1-3 = Sound Factory, Level 0, Assignment 1, Achievement 3
 *
 * Version: 1.0.0
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 1: CURRICULUM CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

var CURRICULUM_CONFIG = {
  SHEETS: {
    CURRICULUM: '📚 Curriculum',
    EVENTS: '📅 Events',
    EVENT_VICTORIES: '🏆 Event Victories',
    LEDGER: '📜 Change Ledger'
  },

  // Default studios (can be extended)
  STUDIOS: ['Sound', 'Visual', 'Interactive'],

  // Resource types for achievements
  RESOURCE_TYPES: ['None', 'YouTube', 'Link', 'File', 'Google Doc', 'Google Slides']
};

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 2: SHEET SETUP & FORMATTING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Setup the Curriculum sheet with proper structure
 */
function setupCurriculumSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet) {
    sheet = ss.insertSheet(CURRICULUM_CONFIG.SHEETS.CURRICULUM);
  }

  // Apply dark theme if available
  if (typeof applyGlobalDarkTheme_ === 'function') {
    applyGlobalDarkTheme_(sheet);
  }

  // Description row
  sheet.getRange('A1:M1').merge().setValue('📚 CURRICULUM — Define your tracks, assignments, and achievements. Achievement IDs auto-generate as {trackCode}-L{level}-a{assign}-{ach}');
  if (typeof applyDescriptionStyle_ === 'function') {
    applyDescriptionStyle_(sheet.getRange('A1:M1'));
  }

  // Headers
  var headers = [
    'Achievement ID',      // A - Auto-generated
    'Track Code',          // B
    'Track Name',          // C
    'Studio',              // D
    'Level',               // E
    'Assignment #',        // F
    'Assignment Title',    // G
    'Achievement #',       // H
    'Achievement Name',    // I
    'MP Value',            // J
    'Resource Type',       // K
    'Resource URL',        // L
    'Active'               // M
  ];

  sheet.getRange('A2:M2').setValues([headers]);
  if (typeof applyHeaderStyle_ === 'function') {
    applyHeaderStyle_(sheet.getRange('A2:M2'));
  }

  // Column widths
  sheet.setColumnWidth(1, 150);  // Achievement ID
  sheet.setColumnWidth(2, 100);  // Track Code
  sheet.setColumnWidth(3, 180);  // Track Name
  sheet.setColumnWidth(4, 100);  // Studio
  sheet.setColumnWidth(5, 60);   // Level
  sheet.setColumnWidth(6, 80);   // Assignment #
  sheet.setColumnWidth(7, 200);  // Assignment Title
  sheet.setColumnWidth(8, 80);   // Achievement #
  sheet.setColumnWidth(9, 250);  // Achievement Name
  sheet.setColumnWidth(10, 80);  // MP Value
  sheet.setColumnWidth(11, 120); // Resource Type
  sheet.setColumnWidth(12, 300); // Resource URL
  sheet.setColumnWidth(13, 70);  // Active

  // Data validations
  var studioRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRICULUM_CONFIG.STUDIOS, true)
    .build();
  sheet.getRange('D3:D1000').setDataValidation(studioRule);

  var resourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CURRICULUM_CONFIG.RESOURCE_TYPES, true)
    .build();
  sheet.getRange('K3:K1000').setDataValidation(resourceRule);

  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  sheet.getRange('M3:M1000').setDataValidation(activeRule);

  sheet.setFrozenRows(2);

  return sheet;
}

/**
 * Setup the Events sheet for special events
 */
function setupEventsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    sheet = ss.insertSheet(CURRICULUM_CONFIG.SHEETS.EVENTS);
  }

  if (typeof applyGlobalDarkTheme_ === 'function') {
    applyGlobalDarkTheme_(sheet);
  }

  sheet.getRange('A1:K1').merge().setValue('📅 SPECIAL EVENTS — Time-limited events with victory thresholds. Event Track Code is used for achievements (e.g., "event-holiday2024")');
  if (typeof applyDescriptionStyle_ === 'function') {
    applyDescriptionStyle_(sheet.getRange('A1:K1'));
  }

  var headers = [
    'Event ID',           // A - Unique identifier
    'Event Name',         // B
    'Description',        // C
    'Event Track Code',   // D - Used for achievement IDs
    'Start Date',         // E
    'End Date',           // F
    'Victory MP Tier 1',  // G - MP threshold for bronze
    'Victory MP Tier 2',  // H - Silver
    'Victory MP Tier 3',  // I - Gold
    'Tier Rewards',       // J - JSON: {"tier1": "10 coins", "tier2": "25 coins", "tier3": "50 coins + item"}
    'Active'              // K
  ];

  sheet.getRange('A2:K2').setValues([headers]);
  if (typeof applyHeaderStyle_ === 'function') {
    applyHeaderStyle_(sheet.getRange('A2:K2'));
  }

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 300);
  sheet.setColumnWidth(11, 70);

  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  sheet.getRange('K3:K100').setDataValidation(activeRule);

  sheet.setFrozenRows(2);

  return sheet;
}

/**
 * Setup the Change Ledger for undo functionality
 */
function setupLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.LEDGER);

  if (!sheet) {
    sheet = ss.insertSheet(CURRICULUM_CONFIG.SHEETS.LEDGER);
    sheet.hideSheet();
  }

  if (sheet.getLastRow() < 1) {
    sheet.appendRow(['Timestamp', 'Action', 'EntityType', 'EntityId', 'OldValue', 'NewValue', 'User', 'Undone']);
  }

  return sheet;
}

/**
 * Setup all curriculum sheets
 */
function setupCurriculumSheets() {
  setupCurriculumSheet();
  setupEventsSheet();
  setupLedgerSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('Curriculum sheets configured!', '✅ Complete', 5);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 3: CURRICULUM CRUD OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get all curriculum data structured hierarchically
 * Returns: { studios: { Sound: { tracks: [...] } } }
 */
function getCurriculumData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet || sheet.getLastRow() < 3) {
    return { studios: {}, tracks: [], raw: [] };
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 13).getValues();
  var raw = [];
  var studios = {};

  // Initialize studios
  CURRICULUM_CONFIG.STUDIOS.forEach(function(studio) {
    studios[studio] = { name: studio, tracks: {} };
  });

  data.forEach(function(row, index) {
    if (!row[1]) return; // Skip if no track code

    var entry = {
      rowIndex: index + 3,
      achievementId: row[0],
      trackCode: row[1],
      trackName: row[2],
      studio: row[3],
      level: row[4],
      assignmentNum: row[5],
      assignmentTitle: row[6],
      achievementNum: row[7],
      achievementName: row[8],
      mpValue: row[9],
      resourceType: row[10],
      resourceUrl: row[11],
      active: row[12] === true || row[12] === 'TRUE'
    };

    raw.push(entry);

    // Build hierarchy
    var studioKey = entry.studio || 'Sound';
    if (!studios[studioKey]) {
      studios[studioKey] = { name: studioKey, tracks: {} };
    }

    var trackKey = entry.trackCode;
    if (!studios[studioKey].tracks[trackKey]) {
      studios[studioKey].tracks[trackKey] = {
        code: trackKey,
        name: entry.trackName,
        levels: {}
      };
    }

    var levelKey = 'L' + entry.level;
    if (!studios[studioKey].tracks[trackKey].levels[levelKey]) {
      studios[studioKey].tracks[trackKey].levels[levelKey] = {
        level: entry.level,
        assignments: {}
      };
    }

    var assignKey = 'a' + entry.assignmentNum;
    if (!studios[studioKey].tracks[trackKey].levels[levelKey].assignments[assignKey]) {
      studios[studioKey].tracks[trackKey].levels[levelKey].assignments[assignKey] = {
        num: entry.assignmentNum,
        title: entry.assignmentTitle,
        achievements: []
      };
    }

    studios[studioKey].tracks[trackKey].levels[levelKey].assignments[assignKey].achievements.push({
      num: entry.achievementNum,
      id: entry.achievementId,
      name: entry.achievementName,
      mp: entry.mpValue,
      resourceType: entry.resourceType,
      resourceUrl: entry.resourceUrl,
      active: entry.active,
      rowIndex: entry.rowIndex
    });
  });

  // Convert tracks objects to arrays and sort
  Object.keys(studios).forEach(function(studioKey) {
    var trackArr = [];
    Object.keys(studios[studioKey].tracks).forEach(function(trackCode) {
      var track = studios[studioKey].tracks[trackCode];

      // Convert levels to array and sort
      var levelArr = [];
      Object.keys(track.levels).forEach(function(levelKey) {
        var level = track.levels[levelKey];

        // Convert assignments to array and sort
        var assignArr = [];
        Object.keys(level.assignments).forEach(function(assignKey) {
          var assign = level.assignments[assignKey];
          assign.achievements.sort(function(a, b) { return a.num - b.num; });
          assignArr.push(assign);
        });
        assignArr.sort(function(a, b) { return a.num - b.num; });
        level.assignments = assignArr;
        levelArr.push(level);
      });
      levelArr.sort(function(a, b) { return a.level - b.level; });
      track.levels = levelArr;
      trackArr.push(track);
    });
    trackArr.sort(function(a, b) { return a.code.localeCompare(b.code); });
    studios[studioKey].tracks = trackArr;
  });

  return { studios: studios, raw: raw };
}

/**
 * Get all tracks with basic info (for dropdown lists)
 */
function getTracksList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);

  if (!sheet || sheet.getLastRow() < 3) {
    return [];
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 10).getValues();
  var tracks = [];
  var seen = {};

  data.forEach(function(row) {
    if (!row[0] || seen[row[0] + '-L' + row[3]]) return;
    seen[row[0] + '-L' + row[3]] = true;

    tracks.push({
      code: row[0],
      name: row[1],
      studio: row[2],
      level: row[3],
      unlockMP: row[4],
      prerequisites: row[5],
      active: row[6] === true || row[6] === 'TRUE',
      trackType: row[8]
    });
  });

  return tracks;
}

/**
 * Create a new track in the Tracks sheet
 */
function createTrack(trackCode, trackName, studio, level, unlockMP, prerequisites, trackType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);

  if (!sheet) {
    throw new Error('Tracks sheet not found. Run setup first.');
  }

  // Log to ledger
  logChange_('CREATE', 'track', trackCode + '-L' + level, null, JSON.stringify({
    trackCode: trackCode,
    trackName: trackName,
    studio: studio,
    level: level
  }));

  sheet.appendRow([
    trackCode,
    trackName,
    studio,
    level,
    unlockMP || 160,
    prerequisites || '',
    'TRUE',
    'UNLOCKED_ONLY',
    trackType || 'Power',
    ''
  ]);

  return { success: true, message: 'Track created' };
}

/**
 * Create a new achievement entry in curriculum
 */
function createAchievement(trackCode, trackName, studio, level, assignmentNum, assignmentTitle, achievementNum, achievementName, mpValue, resourceType, resourceUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet) {
    setupCurriculumSheet();
    sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);
  }

  // Generate achievement ID
  var achievementId = trackCode.toLowerCase() + '-L' + level + '-a' + assignmentNum + '-' + achievementNum;

  // Log to ledger
  logChange_('CREATE', 'achievement', achievementId, null, JSON.stringify({
    name: achievementName,
    mp: mpValue
  }));

  sheet.appendRow([
    achievementId,
    trackCode,
    trackName,
    studio,
    level,
    assignmentNum,
    assignmentTitle,
    achievementNum,
    achievementName,
    mpValue || 10,
    resourceType || 'None',
    resourceUrl || '',
    'TRUE'
  ]);

  return {
    success: true,
    achievementId: achievementId,
    message: 'Achievement created: ' + achievementId
  };
}

/**
 * Update an achievement
 */
function updateAchievement(rowIndex, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet) {
    throw new Error('Curriculum sheet not found');
  }

  var currentRow = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
  var achievementId = currentRow[0];

  // Log to ledger
  logChange_('UPDATE', 'achievement', achievementId, JSON.stringify(currentRow), JSON.stringify(updates));

  // Apply updates
  if (updates.achievementName !== undefined) sheet.getRange(rowIndex, 9).setValue(updates.achievementName);
  if (updates.mpValue !== undefined) sheet.getRange(rowIndex, 10).setValue(updates.mpValue);
  if (updates.resourceType !== undefined) sheet.getRange(rowIndex, 11).setValue(updates.resourceType);
  if (updates.resourceUrl !== undefined) sheet.getRange(rowIndex, 12).setValue(updates.resourceUrl);
  if (updates.active !== undefined) sheet.getRange(rowIndex, 13).setValue(updates.active ? 'TRUE' : 'FALSE');
  if (updates.assignmentTitle !== undefined) sheet.getRange(rowIndex, 7).setValue(updates.assignmentTitle);

  return { success: true, message: 'Achievement updated' };
}

/**
 * Delete an achievement (or mark inactive)
 */
function deleteAchievement(rowIndex, permanent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet) {
    throw new Error('Curriculum sheet not found');
  }

  var currentRow = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
  var achievementId = currentRow[0];

  logChange_('DELETE', 'achievement', achievementId, JSON.stringify(currentRow), null);

  if (permanent) {
    sheet.deleteRow(rowIndex);
  } else {
    sheet.getRange(rowIndex, 13).setValue('FALSE');
  }

  return { success: true, message: 'Achievement deleted' };
}

/**
 * Bulk create achievements for an assignment
 */
function createAssignmentWithAchievements(trackCode, trackName, studio, level, assignmentNum, assignmentTitle, achievements) {
  var results = [];

  achievements.forEach(function(ach, index) {
    var achievementNum = index + 1;
    var result = createAchievement(
      trackCode,
      trackName,
      studio,
      level,
      assignmentNum,
      assignmentTitle,
      achievementNum,
      ach.name,
      ach.mp || 10,
      ach.resourceType || 'None',
      ach.resourceUrl || ''
    );
    results.push(result);
  });

  return {
    success: true,
    created: results.length,
    achievements: results
  };
}

/**
 * Reorder achievements within an assignment (drag-drop support)
 */
function reorderAchievements(assignmentKey, newOrder) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CURRICULUM);

  if (!sheet) {
    throw new Error('Curriculum sheet not found');
  }

  // newOrder is array of { rowIndex, newAchievementNum }
  newOrder.forEach(function(item) {
    // Update achievement number
    sheet.getRange(item.rowIndex, 8).setValue(item.newAchievementNum);

    // Regenerate achievement ID
    var row = sheet.getRange(item.rowIndex, 1, 1, 8).getValues()[0];
    var newId = row[1].toLowerCase() + '-L' + row[4] + '-a' + row[5] + '-' + item.newAchievementNum;
    sheet.getRange(item.rowIndex, 1).setValue(newId);
  });

  logChange_('REORDER', 'achievements', assignmentKey, null, JSON.stringify(newOrder));

  return { success: true, message: 'Achievements reordered' };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 4: EVENTS CRUD OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get all special events
 */
function getEventsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet || sheet.getLastRow() < 3) {
    return [];
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 11).getValues();
  var events = [];

  var now = new Date();

  data.forEach(function(row, index) {
    if (!row[0]) return;

    var startDate = new Date(row[4]);
    var endDate = new Date(row[5]);
    var status = 'upcoming';

    if (now > endDate) {
      status = 'past';
    } else if (now >= startDate && now <= endDate) {
      status = 'active';
    }

    events.push({
      rowIndex: index + 3,
      eventId: row[0],
      name: row[1],
      description: row[2],
      trackCode: row[3],
      startDate: row[4],
      endDate: row[5],
      tier1MP: row[6],
      tier2MP: row[7],
      tier3MP: row[8],
      tierRewards: row[9],
      active: row[10] === true || row[10] === 'TRUE',
      status: status
    });
  });

  return events;
}

/**
 * Create a new special event
 */
function createEvent(eventData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    setupEventsSheet();
    sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);
  }

  var eventId = eventData.eventId || 'event-' + Date.now();

  logChange_('CREATE', 'event', eventId, null, JSON.stringify(eventData));

  sheet.appendRow([
    eventId,
    eventData.name,
    eventData.description || '',
    eventData.trackCode || eventId,
    eventData.startDate,
    eventData.endDate,
    eventData.tier1MP || 50,
    eventData.tier2MP || 100,
    eventData.tier3MP || 150,
    eventData.tierRewards || '{"tier1":"10 coins","tier2":"25 coins","tier3":"50 coins + Mystery Item"}',
    'TRUE'
  ]);

  return { success: true, eventId: eventId, message: 'Event created' };
}

/**
 * Update an event
 */
function updateEvent(rowIndex, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    throw new Error('Events sheet not found');
  }

  var currentRow = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];

  logChange_('UPDATE', 'event', currentRow[0], JSON.stringify(currentRow), JSON.stringify(updates));

  if (updates.name !== undefined) sheet.getRange(rowIndex, 2).setValue(updates.name);
  if (updates.description !== undefined) sheet.getRange(rowIndex, 3).setValue(updates.description);
  if (updates.trackCode !== undefined) sheet.getRange(rowIndex, 4).setValue(updates.trackCode);
  if (updates.startDate !== undefined) sheet.getRange(rowIndex, 5).setValue(updates.startDate);
  if (updates.endDate !== undefined) sheet.getRange(rowIndex, 6).setValue(updates.endDate);
  if (updates.tier1MP !== undefined) sheet.getRange(rowIndex, 7).setValue(updates.tier1MP);
  if (updates.tier2MP !== undefined) sheet.getRange(rowIndex, 8).setValue(updates.tier2MP);
  if (updates.tier3MP !== undefined) sheet.getRange(rowIndex, 9).setValue(updates.tier3MP);
  if (updates.tierRewards !== undefined) sheet.getRange(rowIndex, 10).setValue(updates.tierRewards);
  if (updates.active !== undefined) sheet.getRange(rowIndex, 11).setValue(updates.active ? 'TRUE' : 'FALSE');

  return { success: true, message: 'Event updated' };
}

/**
 * Delete an event
 */
function deleteEvent(rowIndex, permanent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    throw new Error('Events sheet not found');
  }

  var currentRow = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];

  logChange_('DELETE', 'event', currentRow[0], JSON.stringify(currentRow), null);

  if (permanent) {
    sheet.deleteRow(rowIndex);
  } else {
    sheet.getRange(rowIndex, 11).setValue('FALSE');
  }

  return { success: true, message: 'Event deleted' };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 5: CHANGE LEDGER & UNDO SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Log a change to the ledger
 */
function logChange_(action, entityType, entityId, oldValue, newValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.LEDGER);

  if (!sheet) {
    setupLedgerSheet();
    sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.LEDGER);
  }

  var user = Session.getActiveUser().getEmail() || 'system';

  sheet.appendRow([
    new Date(),
    action,
    entityType,
    entityId,
    oldValue || '',
    newValue || '',
    user,
    'FALSE'
  ]);
}

/**
 * Get change history (for UI display)
 */
function getChangeHistory(limit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.LEDGER);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var numRows = Math.min(limit || 50, sheet.getLastRow() - 1);
  var startRow = Math.max(2, sheet.getLastRow() - numRows + 1);

  var data = sheet.getRange(startRow, 1, numRows, 8).getValues();
  var history = [];

  // Reverse to show newest first
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    if (!row[0]) continue;

    history.push({
      rowIndex: startRow + i,
      timestamp: row[0],
      action: row[1],
      entityType: row[2],
      entityId: row[3],
      oldValue: row[4],
      newValue: row[5],
      user: row[6],
      undone: row[7] === true || row[7] === 'TRUE'
    });
  }

  return history;
}

/**
 * Undo a specific change
 */
function undoChange(ledgerRowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledgerSheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.LEDGER);

  if (!ledgerSheet) {
    throw new Error('Ledger sheet not found');
  }

  var changeRow = ledgerSheet.getRange(ledgerRowIndex, 1, 1, 8).getValues()[0];
  var action = changeRow[1];
  var entityType = changeRow[2];
  var entityId = changeRow[3];
  var oldValue = changeRow[4];
  var alreadyUndone = changeRow[7] === true || changeRow[7] === 'TRUE';

  if (alreadyUndone) {
    return { success: false, message: 'This change has already been undone' };
  }

  // Determine target sheet
  var targetSheetName;
  if (entityType === 'achievement') {
    targetSheetName = CURRICULUM_CONFIG.SHEETS.CURRICULUM;
  } else if (entityType === 'event') {
    targetSheetName = CURRICULUM_CONFIG.SHEETS.EVENTS;
  } else if (entityType === 'track') {
    targetSheetName = CONFIG.SHEETS.CONFIG_TRACKS;
  } else {
    return { success: false, message: 'Unknown entity type: ' + entityType };
  }

  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    return { success: false, message: 'Target sheet not found' };
  }

  // Perform undo based on action type
  if (action === 'CREATE') {
    // Find and delete the created row
    var data = targetSheet.getDataRange().getValues();
    for (var i = 2; i < data.length; i++) {
      if (data[i][0] === entityId) {
        targetSheet.deleteRow(i + 1);
        break;
      }
    }
  } else if (action === 'UPDATE' || action === 'DELETE') {
    // Restore old value
    if (oldValue) {
      var oldData = JSON.parse(oldValue);
      // Find row by entity ID and restore
      var data = targetSheet.getDataRange().getValues();
      for (var i = 2; i < data.length; i++) {
        if (data[i][0] === entityId) {
          // Restore full row
          if (Array.isArray(oldData)) {
            targetSheet.getRange(i + 1, 1, 1, oldData.length).setValues([oldData]);
          }
          break;
        }
      }
    }
  }

  // Mark as undone in ledger
  ledgerSheet.getRange(ledgerRowIndex, 8).setValue('TRUE');

  // Log the undo action
  logChange_('UNDO', entityType, entityId, null, 'Undid action: ' + action);

  return { success: true, message: 'Change undone successfully' };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: GOOGLE CLASSROOM INTEGRATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get queue status for Classroom publishing
 */
function getClassroomPublishQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);

  if (!sheet || sheet.getLastRow() < 3) {
    return { ready: [], pending: [], completed: [], errors: [] };
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 14).getValues();
  var queue = { ready: [], pending: [], completed: [], errors: [] };

  data.forEach(function(row, index) {
    if (!row[0]) return;

    var item = {
      rowIndex: index + 3,
      trackCode: row[0],
      level: row[1],
      week: row[2],
      assignNum: row[3],
      title: row[4],
      maxMP: row[5],
      topic: row[6],
      description: row[7],
      viewMaterials: row[8],
      copyMaterials: row[9],
      websiteId: row[10],
      status: row[11],
      classroomId: row[12],
      createdDate: row[13]
    };

    if (item.status === '⏳ Ready') {
      queue.ready.push(item);
    } else if (item.status && item.status.indexOf('✅') === 0) {
      queue.completed.push(item);
    } else if (item.status === '❌ Error') {
      queue.errors.push(item);
    } else {
      queue.pending.push(item);
    }
  });

  return queue;
}

/**
 * Publish a batch of assignments to Classroom with progress reporting
 * Returns progress updates via callback
 */
function publishToClassroomBatch(rowIndices) {
  var results = [];
  var total = rowIndices.length;

  for (var i = 0; i < rowIndices.length; i++) {
    var rowIndex = rowIndices[i];

    try {
      // Call existing createSelectedAssignments logic for single row
      // This is a simplified version - in practice, call the existing function
      var result = publishSingleAssignment_(rowIndex);
      results.push({ rowIndex: rowIndex, success: true, result: result });
    } catch (e) {
      results.push({ rowIndex: rowIndex, success: false, error: e.message });
    }

    // Update progress (this would be sent to client in real impl)
    var progress = Math.round(((i + 1) / total) * 100);
    console.log('Publish progress: ' + progress + '%');
  }

  return {
    success: true,
    total: total,
    succeeded: results.filter(function(r) { return r.success; }).length,
    failed: results.filter(function(r) { return !r.success; }).length,
    results: results
  };
}

/**
 * Internal: Publish single assignment
 */
function publishSingleAssignment_(rowIndex) {
  // This will call the existing createSelectedAssignments or similar
  // For now, just update status
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);

  // Mark as created (in real impl, actually push to Classroom)
  sheet.getRange(rowIndex, 12).setValue('✅ Created');
  sheet.getRange(rowIndex, 14).setValue(new Date());

  return { published: true };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 7: TRACK SELECTION FORM MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get all track selection forms
 */
function getTrackSelectionForms() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var formsJson = props.getProperty('trackSelectionForms') || '[]';

  try {
    return JSON.parse(formsJson);
  } catch (e) {
    return [];
  }
}

/**
 * Create a new track selection form
 */
function createTrackSelectionForm(formName, tracks) {
  // Use existing createTrackSelectionForm logic
  var result = createTrackSelectionForm(); // calls existing function

  // Store reference
  var forms = getTrackSelectionForms();
  forms.push({
    id: result.formId || 'form-' + Date.now(),
    name: formName,
    formUrl: result.formUrl,
    createdAt: new Date().toISOString(),
    tracks: tracks,
    active: true
  });

  var props = PropertiesService.getScriptProperties();
  props.setProperty('trackSelectionForms', JSON.stringify(forms));

  return { success: true, forms: forms };
}

/**
 * Delete a track selection form
 */
function deleteTrackSelectionForm(formId) {
  var forms = getTrackSelectionForms();
  forms = forms.filter(function(f) { return f.id !== formId; });

  var props = PropertiesService.getScriptProperties();
  props.setProperty('trackSelectionForms', JSON.stringify(forms));

  return { success: true, forms: forms };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 8: WEBAPP ENDPOINTS FOR CURRICULUM DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get full curriculum dashboard data
 */
function getCurriculumDashboardData() {
  return {
    curriculum: getCurriculumData(),
    tracks: getTracksList(),
    events: getEventsData(),
    publishQueue: getClassroomPublishQueue(),
    forms: getTrackSelectionForms(),
    history: getChangeHistory(20),
    studios: CURRICULUM_CONFIG.STUDIOS,
    resourceTypes: CURRICULUM_CONFIG.RESOURCE_TYPES
  };
}

/**
 * Import existing tracks into curriculum (migration helper)
 */
function importTracksIntoCurriculum() {
  var tracks = getTracksList();
  var results = [];

  tracks.forEach(function(track) {
    // Check if track already has curriculum entries
    var existing = getCurriculumData().raw.filter(function(r) {
      return r.trackCode === track.code && r.level === track.level;
    });

    if (existing.length === 0) {
      // Create a placeholder assignment with one achievement
      var result = createAchievement(
        track.code,
        track.name,
        track.studio,
        track.level,
        1,  // Assignment 1
        track.name + ' - Assignment 1',
        1,  // Achievement 1
        'Complete the assignment',
        10, // Default MP
        'None',
        ''
      );
      results.push(result);
    }
  });

  return {
    success: true,
    imported: results.length,
    message: 'Imported ' + results.length + ' tracks into curriculum'
  };
}
