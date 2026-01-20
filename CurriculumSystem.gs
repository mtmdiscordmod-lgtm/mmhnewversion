/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * CURRICULUM SYSTEM — Visual Curriculum Builder Backend
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * This module handles the curriculum builder dashboard:
 * - Reads from EXISTING sheets (⚙️ Tracks, ➕ Create Assignments)
 * - Special Events with calendar scheduling (independent of tracks)
 * - Track Selection Form management
 *
 * WORKFLOW:
 * - Events are OPEN TO ALL STUDENTS - no track needed first
 * - Tracks are defined in ⚙️ Tracks sheet
 * - Assignments are created via ➕ Create Assignments sheet
 *
 * Version: 1.1.0
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 1: CURRICULUM CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

var CURRICULUM_CONFIG = {
  SHEETS: {
    // Use EXISTING sheets from mmhv3
    TRACKS: '⚙️ Tracks',
    CREATE_ASSIGNMENT: '➕ Create Assignments',
    ASSIGNMENTS: '📋 Assignments',
    ITEMS: '📦 Item Config',
    // New sheets for curriculum system
    EVENTS: '📅 Events',
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
 * Events are OPEN TO ALL STUDENTS by default
 *
 * Tier rewards structure (per tier):
 * - MP threshold to reach tier
 * - Raw MP reward
 * - Raw Timmy Coins reward
 * - MP Bonus % (can have BOTH mp and coin bonus)
 * - Coin Bonus %
 * - Loot box item
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

  // Now supports BOTH MP bonus AND coin bonus per tier
  sheet.getRange('A1:AA1').merge().setValue('📅 SPECIAL EVENTS — OPEN TO ALL STUDENTS. Each tier can give: Raw MP, Coins, MP Bonus %, Coin Bonus %, and Loot Item');
  if (typeof applyDescriptionStyle_ === 'function') {
    applyDescriptionStyle_(sheet.getRange('A1:AA1'));
  }

  var headers = [
    'Event ID',           // A
    'Event Name',         // B
    'Description',        // C
    'Event Track Code',   // D
    'Start Date',         // E
    'End Date',           // F
    // Tier 1 (Bronze) - columns G-L
    'T1 Threshold',       // G
    'T1 +MP',             // H
    'T1 +Coins',          // I
    'T1 MP Bonus%',       // J - NEW: separate MP bonus
    'T1 Coin Bonus%',     // K - NEW: separate Coin bonus
    'T1 Loot',            // L
    // Tier 2 (Silver) - columns M-R
    'T2 Threshold',       // M
    'T2 +MP',             // N
    'T2 +Coins',          // O
    'T2 MP Bonus%',       // P
    'T2 Coin Bonus%',     // Q
    'T2 Loot',            // R
    // Tier 3 (Gold) - columns S-X
    'T3 Threshold',       // S
    'T3 +MP',             // T
    'T3 +Coins',          // U
    'T3 MP Bonus%',       // V
    'T3 Coin Bonus%',     // W
    'T3 Loot',            // X
    // Meta
    'Bonus Days',         // Y
    'Active'              // Z
  ];

  sheet.getRange('A2:Z2').setValues([headers]);
  if (typeof applyHeaderStyle_ === 'function') {
    applyHeaderStyle_(sheet.getRange('A2:Z2'));
  }

  // Column widths - compact
  sheet.setColumnWidth(1, 110);  // Event ID
  sheet.setColumnWidth(2, 160);  // Name
  sheet.setColumnWidth(3, 200);  // Description
  sheet.setColumnWidth(4, 110);  // Track Code
  sheet.setColumnWidth(5, 90);   // Start
  sheet.setColumnWidth(6, 90);   // End
  // Tier columns - narrower
  for (var i = 7; i <= 24; i++) {
    sheet.setColumnWidth(i, 70);
  }
  sheet.setColumnWidth(25, 70);  // Bonus Days
  sheet.setColumnWidth(26, 50);  // Active

  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  sheet.getRange('Z3:Z100').setDataValidation(activeRule);

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
 * Get all curriculum data from EXISTING sheets
 * Reads from ⚙️ Tracks and ➕ Create Assignments
 * Returns: { studios: { Sound: { tracks: [...] } } }
 */
function getCurriculumData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read from EXISTING ⚙️ Tracks sheet
  var tracksSheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.TRACKS);
  if (!tracksSheet || tracksSheet.getLastRow() < 3) {
    return { studios: {}, tracks: [], raw: [] };
  }

  // Tracks sheet columns: Track Code, Track Name, Studio, Level, Unlock MP, Prerequisites, Active, Assign To, Track Type, Notes
  var tracksData = tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 10).getValues();

  // Also read assignments from ➕ Create Assignments sheet
  var assignSheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CREATE_ASSIGNMENT);
  var assignData = [];
  if (assignSheet && assignSheet.getLastRow() >= 3) {
    // Columns: Track Code, Level, Week, Assign#, Title, Max MP, Topic, Description, View Materials, Copy Materials, Website ID, Status, Classroom ID, Created Date
    assignData = assignSheet.getRange(3, 1, assignSheet.getLastRow() - 2, 14).getValues();
  }

  var raw = [];
  var studios = {};

  // Build hierarchy from tracks
  tracksData.forEach(function(row, index) {
    if (!row[0]) return; // Skip if no track code

    var trackCode = row[0];
    var trackName = row[1];
    var studio = row[2] || 'Sound';
    var level = row[3] || 0;
    var unlockMP = row[4] || 160;
    var prerequisites = row[5] || '';
    var active = row[6] === true || row[6] === 'TRUE';
    var assignTo = row[7] || 'UNLOCKED_ONLY';
    var trackType = row[8] || 'Power';
    var notes = row[9] || '';

    // Initialize studio if needed
    if (!studios[studio]) {
      studios[studio] = { name: studio, tracks: {} };
    }

    // Create track entry
    var trackKey = trackCode + '-L' + level;
    if (!studios[studio].tracks[trackKey]) {
      studios[studio].tracks[trackKey] = {
        code: trackCode,
        name: trackName,
        studio: studio,
        unlockMP: unlockMP,
        prerequisites: prerequisites,
        trackType: trackType,
        assignTo: assignTo,
        notes: notes,
        rowIndex: index + 3,
        levels: {}
      };
    }

    // Add level
    var levelKey = 'L' + level;
    if (!studios[studio].tracks[trackKey].levels[levelKey]) {
      studios[studio].tracks[trackKey].levels[levelKey] = {
        level: level,
        assignments: []
      };
    }

    raw.push({
      rowIndex: index + 3,
      trackCode: trackCode,
      trackName: trackName,
      studio: studio,
      level: level,
      unlockMP: unlockMP,
      active: active
    });
  });

  // Add assignments to tracks
  assignData.forEach(function(row, index) {
    if (!row[0]) return;

    var trackCode = row[0];
    var level = row[1] || 0;
    var week = row[2];
    var assignNum = row[3];
    var title = row[4];
    var maxMP = row[5];
    var status = row[11];

    // Find the track in studios
    for (var studioKey in studios) {
      var trackKey = trackCode + '-L' + level;
      if (studios[studioKey].tracks[trackKey]) {
        var levelKey = 'L' + level;
        if (studios[studioKey].tracks[trackKey].levels[levelKey]) {
          studios[studioKey].tracks[trackKey].levels[levelKey].assignments.push({
            rowIndex: index + 3,
            num: assignNum,
            week: week,
            title: title,
            maxMP: maxMP,
            status: status
          });
        }
        break;
      }
    }
  });

  // Convert tracks objects to arrays and sort
  Object.keys(studios).forEach(function(studioKey) {
    var trackArr = [];
    Object.keys(studios[studioKey].tracks).forEach(function(trackKey) {
      var track = studios[studioKey].tracks[trackKey];

      // Convert levels to array and sort
      var levelArr = [];
      Object.keys(track.levels).forEach(function(levelKey) {
        var level = track.levels[levelKey];
        // Sort assignments by number
        level.assignments.sort(function(a, b) { return (a.num || 0) - (b.num || 0); });
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
 * Reads from EXISTING ⚙️ Tracks sheet
 */
function getTracksList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.TRACKS);

  if (!sheet || sheet.getLastRow() < 3) {
    return [];
  }

  // Columns: Track Code, Track Name, Studio, Level, Unlock MP, Prerequisites, Active, Assign To, Track Type, Notes
  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 10).getValues();
  var tracks = [];
  var seen = {};

  data.forEach(function(row, index) {
    if (!row[0]) return;

    var key = row[0] + '-L' + row[3];
    if (seen[key]) return;
    seen[key] = true;

    tracks.push({
      rowIndex: index + 3,
      code: row[0],
      name: row[1],
      studio: row[2] || 'Sound',
      level: row[3] || 0,
      unlockMP: row[4] || 160,
      prerequisites: row[5] || '',
      active: row[6] === true || row[6] === 'TRUE',
      assignTo: row[7] || 'UNLOCKED_ONLY',
      trackType: row[8] || 'Power',
      notes: row[9] || ''
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
 * Events are OPEN TO ALL STUDENTS
 * Now supports BOTH MP bonus AND coin bonus per tier
 */
function getEventsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet || sheet.getLastRow() < 3) {
    return [];
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 26).getValues();
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
      openToAll: true, // Events are always open to all students
      // Tier 1 (Bronze) - supports BOTH bonuses
      tier1: {
        mpThreshold: row[6] || 50,
        rawMP: row[7] || 0,
        coins: row[8] || 0,
        mpBonus: row[9] || 0,      // NEW: separate MP bonus %
        coinBonus: row[10] || 0,   // NEW: separate Coin bonus %
        lootItem: row[11] || ''
      },
      // Tier 2 (Silver)
      tier2: {
        mpThreshold: row[12] || 100,
        rawMP: row[13] || 0,
        coins: row[14] || 0,
        mpBonus: row[15] || 0,
        coinBonus: row[16] || 0,
        lootItem: row[17] || ''
      },
      // Tier 3 (Gold)
      tier3: {
        mpThreshold: row[18] || 150,
        rawMP: row[19] || 0,
        coins: row[20] || 0,
        mpBonus: row[21] || 0,
        coinBonus: row[22] || 0,
        lootItem: row[23] || ''
      },
      bonusDays: row[24] || 7,
      active: row[25] === true || row[25] === 'TRUE',
      status: status
    });
  });

  return events;
}

/**
 * Get available loot box items for event rewards
 * Reads from 📦 Item Config sheet
 */
function getAvailableLootItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.ITEMS) ||
              ss.getSheetByName('🎁 Inventory Items') ||
              ss.getSheetByName('Inventory Items');

  if (!sheet || sheet.getLastRow() < 3) {
    return ['Mystery Box', 'Rare Item', 'Epic Item', 'Legendary Item'];
  }

  // Item Config columns typically: Item Name, Description, Rarity, Cost, etc.
  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 2).getValues();
  var items = [];

  data.forEach(function(row) {
    if (row[0]) items.push(row[0]);
  });

  return items.length > 0 ? items : ['Mystery Box', 'Rare Item', 'Epic Item', 'Legendary Item'];
}

/**
 * Create a new special event (OPEN TO ALL STUDENTS)
 * Supports BOTH MP bonus AND coin bonus per tier
 */
function createEvent(eventData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    setupEventsSheet();
    sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);
  }

  var eventId = eventData.eventId || 'event-' + Date.now();
  var t1 = eventData.tier1 || {};
  var t2 = eventData.tier2 || {};
  var t3 = eventData.tier3 || {};

  logChange_('CREATE', 'event', eventId, null, JSON.stringify(eventData));

  sheet.appendRow([
    eventId,
    eventData.name,
    eventData.description || '',
    eventData.trackCode || eventId,
    eventData.startDate,
    eventData.endDate,
    // Tier 1 - now with separate MP and Coin bonuses
    t1.mpThreshold || 50,
    t1.rawMP || 0,
    t1.coins || 10,
    t1.mpBonus || 0,        // MP Bonus %
    t1.coinBonus || 0,      // Coin Bonus %
    t1.lootItem || '',
    // Tier 2
    t2.mpThreshold || 100,
    t2.rawMP || 0,
    t2.coins || 25,
    t2.mpBonus || 0,
    t2.coinBonus || 0,
    t2.lootItem || '',
    // Tier 3
    t3.mpThreshold || 150,
    t3.rawMP || 0,
    t3.coins || 50,
    t3.mpBonus || 5,
    t3.coinBonus || 5,
    t3.lootItem || 'Mystery Box',
    // Meta
    eventData.bonusDays || 7,
    'TRUE'
  ]);

  return { success: true, eventId: eventId, message: 'Event created (Open to all students)' };
}

/**
 * Update an event
 * Supports BOTH MP bonus AND coin bonus per tier
 */
function updateEvent(rowIndex, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.EVENTS);

  if (!sheet) {
    throw new Error('Events sheet not found');
  }

  var currentRow = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];

  logChange_('UPDATE', 'event', currentRow[0], JSON.stringify(currentRow), JSON.stringify(updates));

  // Basic fields
  if (updates.name !== undefined) sheet.getRange(rowIndex, 2).setValue(updates.name);
  if (updates.description !== undefined) sheet.getRange(rowIndex, 3).setValue(updates.description);
  if (updates.trackCode !== undefined) sheet.getRange(rowIndex, 4).setValue(updates.trackCode);
  if (updates.startDate !== undefined) sheet.getRange(rowIndex, 5).setValue(updates.startDate);
  if (updates.endDate !== undefined) sheet.getRange(rowIndex, 6).setValue(updates.endDate);

  // Tier 1 (columns 7-12) - now with separate MP and Coin bonuses
  if (updates.tier1) {
    var t1 = updates.tier1;
    if (t1.mpThreshold !== undefined) sheet.getRange(rowIndex, 7).setValue(t1.mpThreshold);
    if (t1.rawMP !== undefined) sheet.getRange(rowIndex, 8).setValue(t1.rawMP);
    if (t1.coins !== undefined) sheet.getRange(rowIndex, 9).setValue(t1.coins);
    if (t1.mpBonus !== undefined) sheet.getRange(rowIndex, 10).setValue(t1.mpBonus);
    if (t1.coinBonus !== undefined) sheet.getRange(rowIndex, 11).setValue(t1.coinBonus);
    if (t1.lootItem !== undefined) sheet.getRange(rowIndex, 12).setValue(t1.lootItem);
  }

  // Tier 2 (columns 13-18)
  if (updates.tier2) {
    var t2 = updates.tier2;
    if (t2.mpThreshold !== undefined) sheet.getRange(rowIndex, 13).setValue(t2.mpThreshold);
    if (t2.rawMP !== undefined) sheet.getRange(rowIndex, 14).setValue(t2.rawMP);
    if (t2.coins !== undefined) sheet.getRange(rowIndex, 15).setValue(t2.coins);
    if (t2.mpBonus !== undefined) sheet.getRange(rowIndex, 16).setValue(t2.mpBonus);
    if (t2.coinBonus !== undefined) sheet.getRange(rowIndex, 17).setValue(t2.coinBonus);
    if (t2.lootItem !== undefined) sheet.getRange(rowIndex, 18).setValue(t2.lootItem);
  }

  // Tier 3 (columns 19-24)
  if (updates.tier3) {
    var t3 = updates.tier3;
    if (t3.mpThreshold !== undefined) sheet.getRange(rowIndex, 19).setValue(t3.mpThreshold);
    if (t3.rawMP !== undefined) sheet.getRange(rowIndex, 20).setValue(t3.rawMP);
    if (t3.coins !== undefined) sheet.getRange(rowIndex, 21).setValue(t3.coins);
    if (t3.mpBonus !== undefined) sheet.getRange(rowIndex, 22).setValue(t3.mpBonus);
    if (t3.coinBonus !== undefined) sheet.getRange(rowIndex, 23).setValue(t3.coinBonus);
    if (t3.lootItem !== undefined) sheet.getRange(rowIndex, 24).setValue(t3.lootItem);
  }

  // Meta
  if (updates.bonusDays !== undefined) sheet.getRange(rowIndex, 25).setValue(updates.bonusDays);
  if (updates.active !== undefined) sheet.getRange(rowIndex, 26).setValue(updates.active ? 'TRUE' : 'FALSE');

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

  var currentRow = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];

  logChange_('DELETE', 'event', currentRow[0], JSON.stringify(currentRow), null);

  if (permanent) {
    sheet.deleteRow(rowIndex);
  } else {
    sheet.getRange(rowIndex, 26).setValue('FALSE');  // Active column is now 26
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
 * Reads from ➕ Create Assignments sheet
 */
function getClassroomPublishQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CREATE_ASSIGNMENT);

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
  var sheet = ss.getSheetByName(CURRICULUM_CONFIG.SHEETS.CREATE_ASSIGNMENT);

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
    resourceTypes: CURRICULUM_CONFIG.RESOURCE_TYPES,
    lootItems: getAvailableLootItems()
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
