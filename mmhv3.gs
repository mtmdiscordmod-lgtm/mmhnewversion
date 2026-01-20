/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * MULTIMEDIA HEROES — Gamified Classroom Management System
 * ═══════════════════════════════════════════════════════════════════════════════
 * * Version: 1.5.0 
 * Author: Mr. B (Millennial Tech Middle School)
 * * v1.5.0 Changes:
 * - VISUAL: Aligned color palette with Web App (Slate/Cyan/Amber)
 * - VISUAL: Added Logo integration to Dashboard
 * - VISUAL: Dashboard now uses "Dark Mode" HUD aesthetic
 * * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 1: CONSTANTS & CONFIGURATION (UPDATED FOR DARK MODE)
// ═══════════════════════════════════════════════════════════════════════════════

var CONFIG = {
  SHEETS: {
    DASHBOARD: '📊 Dashboard',
    LEADERBOARD: '🏆 Leaderboard',
    CONFIG_TRACKS: '⚙️ Tracks',
    CONFIG_CLASSROOMS: '🏫 Classrooms',
    CREATE_ASSIGNMENT: '➕ Create Assignments',
    ASSIGNMENTS: '📋 Assignments',
    PROGRESS: '📈 Progress',
    MANUAL_UNLOCKS: '🔓 Manual Unlocks',
    UNLOCK_LOG: '📜 Unlock Log',
    TRACK_SELECTIONS: '📝 Track Selections',
    HELP: '❓ Help',
    AUCTIONS: '💰 Auctions',
    TREASURY: '💰 Treasury',
    RATIONS: '📦 Rations',
    ITEMS: '📦 Item Config',
    GRADING: '📊 Grading Export',
    LEDGER_DB: 'hidden_ledger_db',
    INVENTORY_DB: 'hidden_inventory_db',
    BONUS_DB: 'hidden_bonus_mp_db',
    SONG_REQUESTS: '🎵 Song Requests',
    BONUS_ROUNDS: '🎯 Bonus Rounds'
  },
  
  //DARK PALETTE (Tailwind-based)
  COLORS: {
    // Brand Colors
    PRIMARY: '#22d3ee',       // Cyan 400 (Main Text/Borders)
    SECONDARY: '#fbbf24',     // Amber 400 (Accents/Power)
    MUTED: '#94a3b8',         // Slate 400 (Secondary Text)
    
    // UI Backgrounds
    BG_DEEP: '#020617',       // Slate 950 (Main Background)
    BG_HEADER: '#0f172a',     // Slate 900 (Headers)
    BG_ROW_1: '#0f172a',      // Slate 900 (Row A)
    BG_ROW_2: '#1e293b',      // Slate 800 (Row B)
    ROW_NORMAL: '#0f172a',    // Alias for code compatibility
ROW_ALT: '#1e293b',       // Alias for code compatibility
    BG_INPUT: '#334155',      // Slate 700 (Input Fields - Lighter for visibility)
    BORDER: '#334155',        // Slate 700
    
    // Status Colors (Dark Mode Optimized)
    STATUS_READY_BG: '#451a03',   // Dark Amber
    STATUS_READY_FG: '#fbbf24',   // Bright Amber
    
    STATUS_SUCCESS_BG: '#064e3b', // Dark Emerald
    STATUS_SUCCESS_FG: '#34d399', // Bright Emerald
    
    STATUS_ERROR_BG: '#450a0a',   // Dark Red
    STATUS_ERROR_FG: '#f87171',   // Bright Red
    
    STATUS_INFO_BG: '#0c4a6e',    // Dark Sky
    STATUS_INFO_FG: '#38bdf8'     // Bright Sky
  },
  
  ASSETS: {
    // Using the PNG version for Sheets compatibility
    LOGO_URL: 'https://multimediaheroes.com/multimediaheroesassets/mmhlogo.png'
  },
  
  DEFAULTS: {
    UNLOCK_THRESHOLD: 160,
    SYNC_INTERVAL_MINUTES: 5,
    WEEKLY_TARGET_MP: 100,
    WEEKLY_MAX_MP: 130
  },
  
  PROPS: {
    SETUP_COMPLETE: 'SETUP_COMPLETE',
    LAST_SYNC: 'LAST_SYNC',
    SELECTED_CLASSROOM: 'SELECTED_CLASSROOM',
    TRACK_FORM_ID: 'TRACK_FORM_ID',
    TRACK_FORM_ASSIGNMENT_IDS: 'TRACK_FORM_ASSIGNMENT_IDS'
  }
};


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 2: MENU SETUP
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var isSetupComplete = props.getProperty(CONFIG.PROPS.SETUP_COMPLETE) === 'true';

  var menu = ui.createMenu('🎮 Multimedia Heroes');

  if (!isSetupComplete) {
    menu.addItem('⚙️ Setup Wizard (Start Here!)', 'showSetupWizard');
  } else {
    // Main actions
    menu.addItem('🔄 Sync & Check Unlocks', 'syncAndCheckUnlocks')
        .addSeparator()
        // Assignments submenu
        .addSubMenu(ui.createMenu('📝 Assignments')
            .addItem('➕ Create Selected Assignments', 'createSelectedAssignments')
            .addItem('🗑️ Delete Assignment...', 'deleteAssignmentDialog'))
        // Students submenu
        .addSubMenu(ui.createMenu('👥 Students')
            .addItem('📝 Create Track Selection Form', 'createTrackSelectionForm')
            .addItem('🔓 Process Manual Unlocks', 'processManualUnlocks')
            .addItem('💰 Process Treasury', 'processTreasuryBatch'))
        .addSeparator()
        // MMH Overlord submenu (Plugin/Skey management)
        .addSubMenu(ui.createMenu('🎮 MMH Overlord')
            .addItem('⚙️ Setup Overlord Sheets', 'setupOverlordSheets')
            .addItem('🔑 Deploy Skeys for Period...', 'showDeploySkeyDialog')
            .addItem('📊 View Deployment Status', 'showDeploymentStatus')
            .addSeparator()
            .addItem('🔍 Toggle SafeSearch', 'toggleSafeSearchUI'))
        .addSeparator()
        // Settings submenu
        .addSubMenu(ui.createMenu('⚙️ Settings')
            .addItem('📥 Import Students from Classroom', 'importStudents')
            .addItem('🔄 Sync Student Registry', 'syncStudentRegistry')
            .addItem('⏰ Toggle Auto-Sync', 'toggleAutoSync')
            .addItem('🔃 Reorder Tabs', 'reorderSheets_')
            .addItem('🔄 Re-run Setup Wizard', 'showSetupWizard'));
  }

  menu.addToUi();
}

/**
 * Check if auto-sync trigger is currently enabled
 */
function isAutoSyncEnabled_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSyncTrigger') {
      return true;
    }
  }
  return false;
}

/**
 * Toggle auto-sync on/off
 */
/**
 * Toggle auto-sync on/off
 */
function toggleAutoSync() {
  var ui = SpreadsheetApp.getUi();
  if (isAutoSyncEnabled_()) {
    // Disable
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'autoSyncTrigger') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    ui.alert('⏹️ Auto-Sync Disabled', 
      'Auto-sync has been turned OFF.',
      ui.ButtonSet.OK);
  } else {
    // Enable
    ScriptApp.newTrigger('autoSyncTrigger')
      .timeBased()
      .everyHours(1) // <--- CHANGED THIS from everyMinutes(5)
      .create();
    ui.alert('✅ Auto-Sync Enabled', 
      'Auto-sync is now ON.\n\n' +
      'The system will check for grades and unlocks every HOUR.',
      ui.ButtonSet.OK);
  }
}


/**
 * Auto-sync trigger function
 */
function autoSyncTrigger() {
  processNewTrackSelections(); 
  syncGrades();
  runUnlockCheck();
  refreshDashboard();
  refreshLeaderboard();
}

/**
 * Combined sync + unlock + refresh (main user action)
 */
function syncAndCheckUnlocks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // --- NEW STEP ADDED HERE ---
    ss.toast('Processing new track selections...', '🚀 Step 0/5', 30);
    processNewTrackSelections(); 
    // ---------------------------

    ss.toast('Syncing grades...', '🔄 Step 1/5', 30);
    syncGrades();
    
    ss.toast('Updating assignment list...', '📋 Step 2/5', 30);
    updateAssignmentsReferenceSheet_();
    
    ss.toast('Checking unlocks...', '🔓 Step 3/5', 30);
    runUnlockCheck();
    
    ss.toast('Refreshing dashboard...', '📊 Step 4/5', 10);
    refreshDashboard();
    
    ss.toast('Refreshing leaderboard...', '🏆 Step 5/5', 10);
    refreshLeaderboard();
    
    ss.toast('All done!', '✅ Complete', 5);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error('syncAndCheckUnlocks error:', error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 3: SETUP WIZARD & FORMATTING
// ═══════════════════════════════════════════════════════════════════════════════

function showSetupWizard() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var isSetupComplete = props.getProperty(CONFIG.PROPS.SETUP_COMPLETE) === 'true';
  
  var confirmReset = true;
  if (isSetupComplete) {
    var response = ui.alert('⚠️ Re-run Setup?', 'This will re-apply the Multimedia Heroes Dark Mode theme to all sheets.\nData will NOT be deleted.\nContinue?', ui.ButtonSet.YES_NO);
    confirmReset = (response === ui.Button.YES);
  }
  
  if (!confirmReset) return;
  
  var startSetup = ui.alert(
    '🎮 Welcome to Multimedia Heroes', 
    'Initializing the universal classroom gamification system.\n' +
    '(Designed for Multimedia, adaptable for any subject)\n\n' +
    '1. Create/Update sheet tabs\n' +
    '2. Apply Dark Mode Visual Theme\n' +
    '3. Connect to Google Classrooms\n' +
    '4. Import students\n' +
    '5. Generate Starter Data (if empty)\n\n' +
    'Ready to initialize?',
    ui.ButtonSet.YES_NO
  );
  if (startSetup !== ui.Button.YES) return;
  
  try {
    showProgress_('Initializing sheets...', 10);
    setupAllSheets_();
    
    showProgress_('Applying Dark Mode theme...', 30);
    formatAllSheets_();
    
    showProgress_('Loading classrooms...', 50);
    var classrooms = loadClassrooms_();
    
    if (classrooms.length === 0) {
      ui.alert('⚠️ No Classrooms', 'No Google Classrooms found.', ui.ButtonSet.OK);
      return;
    }
    
    showProgress_('Importing students...', 70);
    importStudentsToSheet_(classrooms);
    
    // --- UPDATED: Add Samples for Tracks AND Assignments ---
    showProgress_('Checking for starter data...', 85);
    addSampleTracks_();      // Checks if empty first
    addSampleAssignments_(); // Checks if empty first
    addSampleItems_();       // Checks if empty first
    addSampleBonusQuestions_(); // Checks if empty first
    // -----------------------------------------------------

    showProgress_('Organizing tabs...', 95);
    reorderSheets_();
    
    showProgress_('System Online', 100);
    props.setProperty(CONFIG.PROPS.SETUP_COMPLETE, 'true');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
    if (dashboard) ss.setActiveSheet(dashboard);
    
    onOpen();
    ui.alert('🎉 System Online', 'Multimedia Heroes is ready to gamify your classroom.', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
    console.error('Setup error:', error);
  }
}

function showProgress_(message, percent) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '⚙️ System Setup', 3);
}

function setupAllSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetNames = [
    CONFIG.SHEETS.HELP,
    CONFIG.SHEETS.UNLOCK_LOG,
    CONFIG.SHEETS.TRACK_SELECTIONS,
    CONFIG.SHEETS.MANUAL_UNLOCKS,
    CONFIG.SHEETS.PROGRESS,
    CONFIG.SHEETS.ASSIGNMENTS,
    CONFIG.SHEETS.CREATE_ASSIGNMENT,
    CONFIG.SHEETS.CONFIG_CLASSROOMS,
    CONFIG.SHEETS.CONFIG_TRACKS,
    CONFIG.SHEETS.LEADERBOARD,
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.AUCTIONS,
    CONFIG.SHEETS.ITEMS,
    CONFIG.SHEETS.TREASURY,
    CONFIG.SHEETS.RATIONS,
    CONFIG.SHEETS.GRADING,
    CONFIG.SHEETS.SONG_REQUESTS,
    CONFIG.SHEETS.BONUS_ROUNDS
  ];
  
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name, 0);
    }
  });
  
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0) {
    ss.deleteSheet(sheet1);
  }
}

function formatAllSheets_() {
  formatDashboard_();
  formatLeaderboardSheet_();
  formatTracksSheet_();
  formatClassroomsSheet_();
  formatCreateAssignmentSheet_();
  formatAssignmentsSheet_();
  formatProgressSheet_();
  formatManualUnlocksSheet_();
  formatUnlockLogSheet_();
  formatTrackSelectionsSheet_();
  formatHelpSheet_();
  formatAuctionsSheet_();
  formatItemsSheet_();
  formatTreasurySheet_();
  formatRationsSheet_();
  formatGradingSheet_();
  formatSongRequestsSheet_();
  formatBonusRoundsSheet_();
}

/**
 * ----------------------------------------------------------------------------
 * DARK MODE HELPER FUNCTIONS
 * ----------------------------------------------------------------------------
 */

// Applies base Dark Mode to the entire sheet
function applyGlobalDarkTheme_(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setBackground(CONFIG.COLORS.BG_DEEP)
    .setFontColor(CONFIG.COLORS.MUTED)
    .setFontFamily('Verdana')
    .setFontSize(10)
    .setVerticalAlignment('middle');
}

// Styles the Header Row (Row 2 usually)
function applyHeaderStyle_(range) {
  range.setBackground(CONFIG.COLORS.BG_HEADER)
       .setFontColor(CONFIG.COLORS.PRIMARY)
       .setFontFamily('Consolas')
       .setFontWeight('bold')
       .setFontSize(11)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setBorder(false, false, true, false, false, false, CONFIG.COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Styles the Description Row (Row 1 usually)
function applyDescriptionStyle_(range) {
  range.setBackground(CONFIG.COLORS.BG_ROW_2)
       .setFontColor(CONFIG.COLORS.SECONDARY)
       .setFontFamily('Consolas')
       .setFontSize(10)
       .setFontStyle('italic')
       .setWrap(true)
       .setVerticalAlignment('middle')
       .setBorder(false, false, true, false, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.DASHED);
}

// Styles data rows with alternating dark colors
function applyTableBanding_(sheet, startRow, numCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols);
  // Clear existing banding if any
  range.applyRowBanding(SpreadsheetApp.BandingTheme.GREY, false, false); 
  var banding = range.getBandings()[0];
  if (banding) banding.remove();
  
  // Apply manual colors for total control
  for (var i = 0; i < lastRow - startRow + 1; i++) {
    var rowNum = startRow + i;
    var color = (i % 2 === 0) ? CONFIG.COLORS.BG_ROW_1 : CONFIG.COLORS.BG_ROW_2;
    sheet.getRange(rowNum, 1, 1, numCols).setBackground(color);
  }
}

/**
 * ----------------------------------------------------------------------------
 * SHEET-SPECIFIC FORMATTERS
 * ----------------------------------------------------------------------------
 */

function formatDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) return;
  sheet.clear();
  applyGlobalDarkTheme_(sheet);
  
  // 1. Logo Insertion
  var images = sheet.getImages();
  images.forEach(function(img) { img.remove(); });
  try {
    sheet.insertImage(CONFIG.ASSETS.LOGO_URL, 2, 1, 0, 10).setHeight(80).setWidth(80);
  } catch (e) { console.log("Image error: " + e.message); }

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 50);

  // 2. Main Title
  sheet.getRange('C2:F2').merge().setValue('MULTIMEDIA HEROES')
    .setFontFamily('Consolas').setFontSize(24).setFontWeight('bold').setFontColor(CONFIG.COLORS.PRIMARY).setVerticalAlignment('middle');
  
  // 3. Sub-Title & Signature (Split Cell)
  sheet.getRange('C3:E3').merge().setValue('// UNIVERSAL GAMIFICATION SYSTEM //')
    .setFontFamily('Consolas').setFontSize(10).setFontColor(CONFIG.COLORS.MUTED).setVerticalAlignment('top');

  sheet.getRange('F3').setValue('Made by Mr. B')
    .setFontFamily('Consolas').setFontSize(10).setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.SECONDARY) // Gold/Amber
    .setHorizontalAlignment('right')
    .setVerticalAlignment('top');

  sheet.getRange('E4:F4').merge().setValue('LAST SYNC: WAITING...')
    .setFontFamily('Consolas').setFontSize(9).setFontColor(CONFIG.COLORS.SECONDARY).setHorizontalAlignment('right');

  var sectionStyle = SpreadsheetApp.newTextStyle().setFontSize(12).setFontFamily('Consolas').setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY).build();

  // Stats Box
  sheet.getRange('B6').setValue('>> MISSION STATUS').setTextStyle(sectionStyle);
  var statsRange = sheet.getRange('B7:C10');
  statsRange.setBackground(CONFIG.COLORS.BG_HEADER).setFontColor('#fff').setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B7').setValue('Active Agents');
  sheet.getRange('B8').setValue('On Track (>100 MP)');
  sheet.getRange('B9').setValue('Needs Support');
  sheet.getRange('B10').setValue('Unlocks This Week');
  sheet.getRange('C7:C10').setValue('—').setHorizontalAlignment('center').setFontColor(CONFIG.COLORS.PRIMARY);

  // Unlocks Box
  sheet.getRange('E6').setValue('>> RECENT UNLOCKS').setTextStyle(sectionStyle);
  sheet.getRange('E7:F11').setBackground(CONFIG.COLORS.BG_HEADER).setFontColor('#fff').setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('E7').setValue('System initialized. Run sync.').setFontColor(CONFIG.COLORS.MUTED).setFontStyle('italic');

  // Attention Box
  sheet.getRange('B13').setValue('>> PROXIMITY ALERT').setTextStyle(sectionStyle);
  sheet.getRange('B14:C23').setBackground(CONFIG.COLORS.BG_HEADER).setFontColor(CONFIG.COLORS.SECONDARY).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B14').setValue('No alerts.').setFontColor(CONFIG.COLORS.MUTED).setFontStyle('italic');

  // Critical Box
  sheet.getRange('E13').setValue('>> CRITICAL ATTENTION').setTextStyle(sectionStyle.copy().setForegroundColor(CONFIG.COLORS.STATUS_ERROR_FG).build());
  var filterRule = SpreadsheetApp.newDataValidation().requireValueInList(['▶ Click to show', '📊 All Periods', 'Period 1', 'Period 2', 'Period 3', 'Period 4', 'Period 5', 'Period 6', 'Period 7'], true).build();
  sheet.getRange('F13').setDataValidation(filterRule).setValue('▶ Click to show').setBackground('#1e293b').setFontColor(CONFIG.COLORS.MUTED).setHorizontalAlignment('center');
  sheet.getRange('E14:F23').setBackground(CONFIG.COLORS.BG_HEADER).setFontColor(CONFIG.COLORS.STATUS_ERROR_FG).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('E14').setValue('Select dropdown to reveal.').setFontColor(CONFIG.COLORS.MUTED).setFontStyle('italic');

  // 4. Footer Description
  sheet.getRange('B25:F27').merge().setValue('💡 COMMAND CONSOLE: A gamification system for Multimedia Arts (or any subject). Use the menu above to Sync Grades & Manage Content.')
    .setBackground(CONFIG.COLORS.BG_ROW_2).setFontColor(CONFIG.COLORS.PRIMARY).setFontFamily('Consolas').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setBorder(true, true, true, true, false, false, CONFIG.COLORS.PRIMARY, SpreadsheetApp.BorderStyle.DASHED);

  sheet.setFrozenRows(0);
}

function formatLeaderboardSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LEADERBOARD);
  if (!sheet) return;
  sheet.clear();
  applyGlobalDarkTheme_(sheet);

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 50);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 100);

  sheet.getRange('B1:H1').merge().setValue('🏆 LEADERBOARD').setFontFamily('Consolas').setFontSize(24).setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.PRIMARY).setHorizontalAlignment('center').setBackground(CONFIG.COLORS.BG_DEEP);

  sheet.getRange('B3').setValue('Filter:').setFontWeight('bold').setHorizontalAlignment('right').setFontColor(CONFIG.COLORS.SECONDARY);
  var filterRule = SpreadsheetApp.newDataValidation().requireValueInList(['📊 All Periods', 'Period 1', 'Period 2', 'Period 3', 'Period 4', 'Period 5', 'Period 6', 'Period 7'], true).build();
  sheet.getRange('C3').setDataValidation(filterRule).setValue('📊 All Periods').setBackground(CONFIG.COLORS.BG_INPUT).setFontColor('#fff');

  sheet.getRange('B5:D5').merge().setValue('⚡ THIS WEEK').setFontColor(CONFIG.COLORS.STATUS_INFO_FG).setFontWeight('bold').setHorizontalAlignment('center').setBackground(CONFIG.COLORS.BG_HEADER);
  sheet.getRange('F5:H5').merge().setValue('👑 ALL TIME').setFontColor(CONFIG.COLORS.SECONDARY).setFontWeight('bold').setHorizontalAlignment('center').setBackground(CONFIG.COLORS.BG_HEADER);

  var headers = sheet.getRange('B6:D6'); headers.setValues([['Rank', 'Agent Name', 'MP']]); applyHeaderStyle_(headers);
  var headers2 = sheet.getRange('F6:H6'); headers2.setValues([['Rank', 'Agent Name', 'MP']]); applyHeaderStyle_(headers2);

  sheet.setFrozenRows(6);
}

function formatTracksSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  var headers = ['Track Code', 'Track Name', 'Studio', 'Level', 'Unlock MP', 'Prerequisites', 'Active', 'Assign To', 'Track Type', 'Notes'];
  sheet.getRange('A1:J1').merge().setValue('📋 TRACK CONFIGURATION — Define skill paths. Level 0 tracks are gated by the Track Selection Form.');
  applyDescriptionStyle_(sheet.getRange('A1:J1'));
  
  sheet.getRange('A2:J2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:J2'));
  
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(10, 200);

  // [FIX] Clear data validations for Column C so you can type ANY studio name (Math, Science, etc.)
  sheet.getRange('C3:C100').clearDataValidations();

  // Keep validation for other columns
  var activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE'], true).build();
  sheet.getRange('G3:G100').setDataValidation(activeRule);
  
  var assignRule = SpreadsheetApp.newDataValidation().requireValueInList(['ALL_STUDENTS', 'UNLOCKED_ONLY'], true).build();
  sheet.getRange('H3:H100').setDataValidation(assignRule);

  applyTableBanding_(sheet, 3, 10);
  sheet.setFrozenRows(2);
}

function formatClassroomsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:F1').merge().setValue('🏫 CLASSROOM CONFIGURATION — Set Active to TRUE for classrooms you want to sync.');
  applyDescriptionStyle_(sheet.getRange('A1:F1'));
  
  var headers = ['Course ID', 'Classroom Name', 'Section', 'Period', 'Active', 'Student Count'];
  sheet.getRange('A2:F2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:F2'));
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  var activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE'], true).build();
  sheet.getRange('E3:E50').setDataValidation(activeRule);
  
  applyTableBanding_(sheet, 3, 6);
  sheet.setFrozenRows(2);
}

function formatItemsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  if (!sheet) return;

  applyGlobalDarkTheme_(sheet);

  sheet.getRange('A1:D1').merge().setValue('📦 ITEM CONFIG — Define items for the Auction Web App.');
  applyDescriptionStyle_(sheet.getRange('A1:D1'));

  var headers = ['Item Name', 'Image URL (Square/GIF)', 'Min Bid', 'Description'];
  sheet.getRange('A2:D2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:D2'));

  sheet.setColumnWidth(1, 200); // Name
  sheet.setColumnWidth(2, 400); // Image
  sheet.setColumnWidth(3, 100); // Cost
  sheet.setColumnWidth(4, 300); // Desc

  applyTableBanding_(sheet, 3, 4);
  sheet.setFrozenRows(2);
}

function addSampleItems_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  if (!sheet) return;

  if (sheet.getLastRow() > 2) return;

  var samples = [
    ['Song Request', 'https://media.giphy.com/media/l41lI4bYmcsPJX9Go/giphy.gif', 50, 'DJ Mr. B plays your jam.'],
    ['Dining Pass', 'https://media.giphy.com/media/3o7TKr3nzbh5WgCFxe/giphy.gif', 100, 'Skip the lunch line.'],
    ['Mystery Box', 'https://media.giphy.com/media/ge91zAgmwUqOY/giphy.gif', 250, 'What is inside? Who knows.']
  ];

  sheet.getRange(3, 1, samples.length, 4).setValues(samples);
  applyTableBanding_(sheet, 3, 4);
}

function formatCreateAssignmentSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:N1').merge().setValue('➕ CREATE ASSIGNMENTS — Inputs A-K. Set Status (Col L) to "⏳ Ready".');
  applyDescriptionStyle_(sheet.getRange('A1:N1'));
  
  // UPDATED COLUMN ORDER: Added "Assign #" at Column D
  var headers = [
    'Track Code',      // A
    'Level',           // B
    'Week',            // C
    'Assign #',        // D (NEW)
    'Title',           // E
    'Max MP',          // F
    'Topic',           // G
    'Description',     // H
    'View Materials',  // I
    'Copy Materials',  // J
    'Website ID',      // K
    'Status',          // L
    'Classroom ID',    // M (Hidden)
    'Created Date'     // N (Hidden)
  ];
  
  sheet.getRange('A2:N2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:N2'));
  
  // Adjust Column Widths
  sheet.setColumnWidth(4, 70);  // Assign #
  sheet.setColumnWidth(5, 200); // Title
  sheet.setColumnWidth(8, 250); // Description
  sheet.setColumnWidth(11, 80); // Website ID
  sheet.setColumnWidth(12, 100);
  sheet.hideColumns(13, 2); // Hide Generated Columns (M, N)
  
  // Highlight Input Area (A-K)
  sheet.getRange('A3:K500').setBackground(CONFIG.COLORS.BG_INPUT).setFontColor('#ffffff').setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Validations
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  if (tracksSheet) {
    var trackCodeRule = SpreadsheetApp.newDataValidation().requireValueInRange(tracksSheet.getRange('A3:A500'), true).build();
    sheet.getRange('A3:A500').setDataValidation(trackCodeRule);
  }
  var weekRule = SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D'], true).build();
  sheet.getRange('C3:C500').setDataValidation(weekRule);
  
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['⏳ Ready', '✅ Created', '✅ Updated', '❌ Error'], true).build();
  sheet.getRange('L3:L500').setDataValidation(statusRule);
  
  // Conditional Formatting
  var readyRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('⏳ Ready')
    .setBackground(CONFIG.COLORS.STATUS_READY_BG).setFontColor(CONFIG.COLORS.STATUS_READY_FG).setRanges([sheet.getRange('L3:L500')]).build();
  var createdRule = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✅')
    .setBackground(CONFIG.COLORS.STATUS_SUCCESS_BG).setFontColor(CONFIG.COLORS.STATUS_SUCCESS_FG).setRanges([sheet.getRange('L3:L500')]).build();
  var errorRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('❌ Error')
    .setBackground(CONFIG.COLORS.STATUS_ERROR_BG).setFontColor(CONFIG.COLORS.STATUS_ERROR_FG).setRanges([sheet.getRange('L3:L500')]).build();
    
  sheet.setConditionalFormatRules([readyRule, createdRule, errorRule]);
  sheet.setFrozenRows(2);
}




function formatAssignmentsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ASSIGNMENTS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:K1').merge().setValue('📋 ASSIGNMENTS — System generated. Read only.');
  applyDescriptionStyle_(sheet.getRange('A1:K1'));
  
  var headers = ['Period', 'Tags', 'Classroom ID', 'Assignment ID', 'Title', 'Track Code', 'Level', 'Week', 'Max Points', 'State', 'Last Updated'];
  sheet.getRange('A2:K2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:K2'));
  
  sheet.setColumnWidth(5, 250);
  if (sheet.getFilter() === null) sheet.getRange('A2:K' + sheet.getMaxRows()).createFilter();
  applyTableBanding_(sheet, 3, 11);
  sheet.setFrozenRows(2);
}

function formatProgressSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:L1').merge().setValue('📈 PROGRESS — Auto-calculated from grades. Week A/B breakdown determines unlocks.');
  applyDescriptionStyle_(sheet.getRange('A1:L1'));
  
  var headers = ['Student Email', 'Student Name', 'Period', 'Track Code', 'Track Name', 'Current Level', 'Track MP Total', 'Current Week MP', 'Week A MP', 'Week B MP', 'Ready to Unlock?', 'Last Activity'];
  sheet.getRange('A2:L2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:L2'));
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(5, 150);
  
  var readyUnlockRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✅ YES')
    .setBackground(CONFIG.COLORS.STATUS_SUCCESS_BG)
    .setFontColor(CONFIG.COLORS.STATUS_SUCCESS_FG)
    .setBold(true)
    .setRanges([sheet.getRange('K3:K1000')])
    .build();
  sheet.setConditionalFormatRules([readyUnlockRule]);
  
  applyTableBanding_(sheet, 3, 12);
  sheet.setFrozenRows(2);
}

function formatManualUnlocksSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.MANUAL_UNLOCKS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:F1').merge().setValue('🔓 MANUAL UNLOCKS — Grant access regardless of MP.');
  applyDescriptionStyle_(sheet.getRange('A1:F1'));
  
  var headers = ['Student Email', 'Track Code', 'Level', 'Reason', 'Status', 'Processed Date'];
  sheet.getRange('A2:F2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:F2'));
  
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(4, 250);
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['⏳ Pending', '✅ Processed', '❌ Error'], true).build();
  sheet.getRange('E3:E500').setDataValidation(statusRule);
  
  applyTableBanding_(sheet, 3, 6);
  sheet.setFrozenRows(2);
}

function formatUnlockLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.UNLOCK_LOG);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:F1').merge().setValue('📜 UNLOCK LOG — System history.');
  applyDescriptionStyle_(sheet.getRange('A1:F1'));
  
  var headers = ['Timestamp', 'Student Email', 'Student Name', 'Unlocked', 'Type', 'Notes'];
  sheet.getRange('A2:F2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:F2'));
  
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(6, 250);
  sheet.setFrozenRows(2);
}

function formatTrackSelectionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TRACK_SELECTIONS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:E1').merge().setValue('📝 TRACK SELECTIONS — Incoming form data.');
  applyDescriptionStyle_(sheet.getRange('A1:E1'));
  
  var headers = ['Timestamp', 'Student Email', 'Selected Track', 'Processed', 'Processed Date'];
  sheet.getRange('A2:E2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:E2'));
  
  sheet.setColumnWidth(2, 220);
  sheet.setFrozenRows(2);
}

function formatHelpSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.HELP);
  if (!sheet) return;
  
  sheet.clear();
  applyGlobalDarkTheme_(sheet);
  
  // Set wide column for readability (document style)
  sheet.setColumnWidth(1, 1000);
  
  var helpContent = [
    ['🎮 MULTIMEDIA HEROES // OPERATOR MANUAL'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['1. MISSION BRIEFING (THE PHILOSOPHY)'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Multimedia Heroes turns your classroom into a role-playing game.'],
    ['Students choose a specialization ("Track"), earn XP ("Media Points"), and unlock new content automatically.'],
    [''],
    ['• AGENCY: Students pick their own path (e.g., Sound, Visual, Code, Math, Science).'],
    ['• MASTERY: They cannot advance until they prove competence (Level Up).'],
    ['• AUTOMATION: You grade in Google Classroom; this system handles the unlocking logic.'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['2. CRITICAL CONCEPTS'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['[ TRACKS ]'],
    ['  Specific skill trees (e.g., "PNO" for Piano, "ALG" for Algebra).'],
    ['  Defined in the "⚙️ Tracks" tab. You can invent any Track for any subject.'],
    [''],
    ['[ LEVELS ]'],
    ['  Level 0 (Foundation): The starter content. Unlocked immediately via the "Track Selection Form".'],
    ['  Level 1+ (Apprentice): Locked content. Must be earned by completing the previous level.'],
    [''],
    ['[ MP (MEDIA POINTS) & THE "WEEK A/B" SYSTEM ]'],
    ['  To ensure balanced learning, leveling up requires two types of work:'],
    ['  • Week A (Training): Practice, worksheets, tutorials. (Target: ~80 MP)'],
    ['  • Week B (Power Test): Projects, boss battles, creative application. (Target: ~80 MP)'],
    ['  >> A student MUST earn enough points in BOTH categories to unlock the next level.'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['3. SETUP CHECKLIST (START HERE)'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Step 1: Configure Classrooms'],
    ['  Go to "🏫 Classrooms". Set "Active" to TRUE for the classes you want to manage.'],
    [''],
    ['Step 2: Define Your Tracks'],
    ['  Go to "⚙️ Tracks". This is the blueprint of your curriculum.'],
    ['  Ensure you have "Level 0" tracks for beginners to start with.'],
    [''],
    ['Step 3: The Sorting Hat'],
    ['  Menu > Students > "Create Track Selection Form".'],
    ['  This posts a form to Google Classroom. When students submit it, they automatically get their Level 0 work.'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['4. CREATING ASSIGNMENTS (THE RULES)'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['Do NOT create assignments manually in Google Classroom if you want them connected to the game.'],
    ['Use the "➕ Create Assignments" tab.'],
    [''],
    ['1. Fill out the row (Track, Level, Week, Title, etc).'],
    ['2. IMPORTANT: "Week" determines the category.'],
    ['   • Select "A" for Training/Practice assignments.'],
    ['   • Select "B" for Projects/Tests.'],
    ['3. "Website ID" (Optional): Links the assignment to the visual Skill Tree web app.'],
    ['   • Leave this blank if you are just using Google Classroom without the visual map.'],
    ['4. Set Status to "⏳ Ready".'],
    ['5. Menu > Assignments > "Create Selected Assignments".'],
    [''],
    ['The system will post them as DRAFT (if locked) or PUBLISHED (if unlocked).'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['5. THE DAILY ROUTINE'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['1. TEACH: Students do work.'],
    ['2. GRADE: You return grades in Google Classroom.'],
    ['3. SYNC: Click "🔄 Sync & Check Unlocks" in the custom menu.'],
    [''],
    ['What happens during Sync?'],
    ['  • Grades are pulled from Classroom.'],
    ['  • The system calculates if a student has enough Week A AND Week B points.'],
    ['  • If they qualify, the NEXT level assignments are automatically assigned to them.'],
    [''],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['6. TROUBLESHOOTING'],
    ['════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['"A student has enough points but didn\'t level up!"'],
    ['  • Check the breakdown on the "📈 Progress" tab.'],
    ['  • Do they have enough "Week A" points? Do they have enough "Week B" points?'],
    ['  • Remember: 160 total MP usually means 80 in A and 80 in B.'],
    [''],
    ['"I need to unlock someone manually."'],
    ['  • Go to "🔓 Manual Unlocks".'],
    ['  • Enter their email, the track code, and the level you want to give them.'],
    ['  • Set status to "⏳ Pending" and run Menu > Students > "Process Manual Unlocks".'],
    [''],
    ['"I added a new studio but the form didn\'t update."'],
    ['  • If you add new Track types, create a NEW form or update the existing one via the menu.']
  ];
  
  // Write content
  sheet.getRange(1, 1, helpContent.length, 1).setValues(helpContent);
  
  // --- VISUAL STYLING (THE SYNDICATE THEME) ---
  
  var headerStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily('Consolas')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY) // Cyan
    .build();
    
  var sectionHeaderStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily('Consolas')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.SECONDARY) // Amber
    .build();
    
  var subHeaderStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily('Consolas')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY) // Cyan
    .build();
    
  var bodyStyle = SpreadsheetApp.newTextStyle()
    .setFontFamily('Verdana')
    .setFontSize(10)
    .setBold(false)
    .setForegroundColor(CONFIG.COLORS.MUTED) // Slate
    .build();

  // Apply Styles Row by Row
  
  // Title Row
  sheet.getRange('A1').setTextStyle(headerStyle).setBackground(CONFIG.COLORS.BG_HEADER).setBorder(false, false, true, false, false, false, CONFIG.COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Content Rows
  for (var i = 1; i < helpContent.length; i++) {
    var row = i + 1;
    var text = helpContent[i][0];
    var range = sheet.getRange(row, 1);
    
    if (text.startsWith('1.') || text.startsWith('2.') || text.startsWith('3.') || text.startsWith('4.') || text.startsWith('5.') || text.startsWith('6.')) {
      // Section Headers (e.g., 1. MISSION BRIEFING)
      range.setTextStyle(sectionHeaderStyle).setBackground(CONFIG.COLORS.BG_ROW_1);
    } else if (text.startsWith('═')) {
      // Separators
      range.setFontColor(CONFIG.COLORS.BORDER);
    } else if (text.startsWith('[') || text.startsWith('Step')) {
      // Sub-headers (e.g., [ TRACKS ] or Step 1:)
      range.setTextStyle(subHeaderStyle);
    } else {
      // Body Text
      range.setTextStyle(bodyStyle);
    }
  }
  
  sheet.setFrozenRows(1);
}

function formatAuctionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.AUCTIONS);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:F1').merge().setValue('💰 AUCTIONS — Spend Timmy Coins. Transactions are final.');
  applyDescriptionStyle_(sheet.getRange('A1:F1'));
  
  var headers = ['Item', 'Student ID (Number Only)', 'Bid/Cost', 'Check Bid', 'Status', 'Confirm Transaction'];
  sheet.getRange('A2:F2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:F2'));
  
  sheet.setColumnWidth(1, 250); // Item
  sheet.setColumnWidth(2, 150); // Student ID
  sheet.setColumnWidth(3, 100); // Cost
  sheet.setColumnWidth(4, 150); // Check Button
  sheet.setColumnWidth(5, 300); // Status
  sheet.setColumnWidth(6, 150); // Confirm Button
  
  // Item Dropdown
  var items = ['Song Request', 'Dining Pass', 'Video Game Lunch', 'Mr. B Game Challenge', 'Showcase Skip', '33MP Ration'];
  var itemRule = SpreadsheetApp.newDataValidation().requireValueInList(items, true).build();
  sheet.getRange('A3:A500').setDataValidation(itemRule);
  
  // Checkboxes for Buttons (simulated buttons)
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('D3:D500').setDataValidation(checkRule);
  sheet.getRange('F3:F500').setDataValidation(checkRule); 
  
  applyTableBanding_(sheet, 3, 6);
  sheet.setFrozenRows(2);
}

function formatTreasurySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TREASURY);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:F1').merge().setValue('💰 TREASURY — Manually Add or Subtract Timmy Coins.');
  applyDescriptionStyle_(sheet.getRange('A1:F1'));
  
  var headers = ['Student Email', 'Amount', 'Type', 'Reason', 'Status', 'Processed Date'];
  sheet.getRange('A2:F2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:F2'));
  
  sheet.setColumnWidth(1, 250);
  
  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['ADD', 'SUBTRACT'], true).build();
  sheet.getRange('C3:C500').setDataValidation(typeRule);
  
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['⏳ Pending', '✅ Processed', '❌ Error'], true).build();
  sheet.getRange('E3:E500').setDataValidation(statusRule);
  
  applyTableBanding_(sheet, 3, 6);
  sheet.setFrozenRows(2);
}

function formatRationsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RATIONS);
  if (!sheet) return;
  
  sheet.clear(); 
  applyGlobalDarkTheme_(sheet);
  
  sheet.getRange('A1:E1').merge().setValue('📦 RATIONS — Students with >= 1 Ration. Use to add 33 Bonus MP (Week Only).');
  applyDescriptionStyle_(sheet.getRange('A1:E1'));
  
  // Updated headers to match refreshRationsList logic: [Email, Name, Count, Consume, Status]
  var headers = ['Student Email', 'Student Name', 'Rations Owned', 'Consume (33 MP)', 'Status'];
  sheet.getRange('A2:E2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:E2'));
  
  sheet.setColumnWidth(1, 200); // Email
  sheet.setColumnWidth(2, 200); // Name
  sheet.setColumnWidth(3, 120); // Count
  sheet.setColumnWidth(4, 120); // Consume Button
  sheet.setColumnWidth(5, 250); // Status
  
  // Checkbox for "Consume" (Col D/4)
  var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('D3:D500').setDataValidation(checkRule);
  
  sheet.setFrozenRows(2);
}

function formatGradingSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.GRADING);
  if (!sheet) return;
  
  applyGlobalDarkTheme_(sheet);
  
  // Config Section
  sheet.getRange('A1:B1').merge().setValue('⚙️ CONFIGURATION');
  applyHeaderStyle_(sheet.getRange('A1:B1'));
  
  var configs = [
    ['Assignment Name:', 'Weekly MP Check'],
    ['Due Date (YYYY-MM-DD):', new Date().toISOString().slice(0, 10)],
    ['Points Possible:', '100'],
    ['Score Type:', 'POINTS'],
    ['Weekly Target (MP):', '100'], // Added
    ['Weekly Max (MP):', '130']     // Added
  ];
  sheet.getRange('A2:B7').setValues(configs);
  sheet.getRange('A2:A7').setFontWeight('bold').setHorizontalAlignment('right').setFontColor(CONFIG.COLORS.SECONDARY);
  sheet.getRange('B2:B7').setBackground(CONFIG.COLORS.BG_INPUT).setFontColor('#ffffff').setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.getRange('D1:E1').merge().setValue('🏫 PERIOD CONFIGURATION');
  applyHeaderStyle_(sheet.getRange('D1:E1'));
  sheet.getRange('D2:E2').setValues([['Period', 'Teacher Name']]);
  
  var periods = [];
  for(var i=1; i<=7; i++) periods.push(['Period ' + i, '']);
  sheet.getRange('D3:E9').setValues(periods);
  sheet.getRange('E3:E9').setBackground(CONFIG.COLORS.BG_INPUT).setFontColor('#ffffff').setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Generate Button Area
  sheet.getRange('A9:B9').merge().setValue('👇 ACTION');
  applyHeaderStyle_(sheet.getRange('A9:B9'));
  
  var generateRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('A10').setDataValidation(generateRule).setValue(false);
  sheet.getRange('B10').setValue('Check box to Generate CSV');
  
  // Output Area
  sheet.getRange('A12:E12').merge().setValue('📄 CSV OUTPUT (Copy below)');
  applyHeaderStyle_(sheet.getRange('A12:E12'));
  sheet.getRange('A13:E30').merge().setValue('waiting for generation...').setVerticalAlignment('top').setWrap(true);
}

function formatSongRequestsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);
  if (!sheet) return;

  applyGlobalDarkTheme_(sheet);

  sheet.getRange('A1:H1').merge().setValue('🎵 SONG REQUESTS — Students pay "Song Request" items to add to this list.');
  applyDescriptionStyle_(sheet.getRange('A1:H1'));

  var headers = ['Timestamp', 'Student Email', 'Student Name', 'Period', 'Artist', 'Song', 'Link', 'Status'];
  sheet.getRange('A2:H2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:H2'));

  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 200); // Email
  sheet.setColumnWidth(3, 150); // Name
  sheet.setColumnWidth(4, 50);  // Period
  sheet.setColumnWidth(5, 150); // Artist
  sheet.setColumnWidth(6, 150); // Song
  sheet.setColumnWidth(7, 200); // Link
  sheet.setColumnWidth(8, 100); // Status

  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDING', 'STAGED', 'PLAYED'], true).build();
  sheet.getRange('H3:H500').setDataValidation(statusRule);

  applyTableBanding_(sheet, 3, 8);
  sheet.setFrozenRows(2);
}

/**
 * Helper to get the Weekly Target MP from the Grading sheet.
 * Defaults to CONFIG.DEFAULTS.WEEKLY_TARGET_MP if missing.
 */
function getWeeklyTargetMP_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.GRADING);
  if (!sheet) return CONFIG.DEFAULTS.WEEKLY_TARGET_MP;
  
  var val = sheet.getRange('B6').getValue(); // Row 6 is Target MP
  var num = parseInt(val);
  return isNaN(num) ? CONFIG.DEFAULTS.WEEKLY_TARGET_MP : num;
}

/**
 * Helper to get the Weekly Max MP from the Grading sheet.
 * Defaults to CONFIG.DEFAULTS.WEEKLY_MAX_MP if missing.
 */
function getWeeklyMaxMP_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.GRADING);
  if (!sheet) return CONFIG.DEFAULTS.WEEKLY_MAX_MP;
  
  var val = sheet.getRange('B7').getValue(); // Row 7 is Max MP
  var num = parseInt(val);
  return isNaN(num) ? CONFIG.DEFAULTS.WEEKLY_MAX_MP : num;
}

/**
 * Reorders sheets to a logical flow.
 */
function reorderSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Desired Order
  var order = [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.LEADERBOARD,
    CONFIG.SHEETS.CREATE_ASSIGNMENT,
    CONFIG.SHEETS.GRADING,
    CONFIG.SHEETS.TREASURY,
    CONFIG.SHEETS.AUCTIONS,
    CONFIG.SHEETS.MANUAL_UNLOCKS,
    CONFIG.SHEETS.PROGRESS,
    CONFIG.SHEETS.ASSIGNMENTS,
    CONFIG.SHEETS.RATIONS,
    CONFIG.SHEETS.UNLOCK_LOG,
    CONFIG.SHEETS.TRACK_SELECTIONS,
    CONFIG.SHEETS.CONFIG_TRACKS,
    CONFIG.SHEETS.CONFIG_CLASSROOMS,
    CONFIG.SHEETS.HELP
  ];
  
  order.forEach(function(name, index) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });
  
  // Return to Dashboard
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (dashboard) ss.setActiveSheet(dashboard);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 4: CLASSROOM API FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function loadClassrooms_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  
  // 1. Read existing config to preserve Period/Active/Count
  var existingConfig = {};
  if (sheet && sheet.getLastRow() >= 3) {
    var currentData = sheet.getRange(3, 1, sheet.getLastRow() - 2, 6).getValues();
    currentData.forEach(function(row) {
      if (row[0]) {
        existingConfig[row[0]] = {
          period: row[3],
          active: row[4],
          count: row[5]
        };
      }
    });
  }

  // 2. Fetch API
  var courses = [];
  var pageToken = null;
  
  do {
    var response = Classroom.Courses.list({
      pageSize: 100,
      courseStates: ['ACTIVE'],
      pageToken: pageToken
    });
    
    if (response.courses) {
      courses = courses.concat(response.courses);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  // 3. Merge
  if (courses.length > 0 && sheet) {
    var data = courses.map(function(course) {
      var prev = existingConfig[course.id];
      return [
        course.id,
        course.name,
        course.section || '',
        prev ? prev.period : '',      // Preserve or default
        prev ? prev.active : 'TRUE',  // Preserve or default
        prev ? prev.count : ''        // Preserve or default
      ];
    });
    
    sheet.getRange(3, 1, data.length, 6).setValues(data);
  }
  
  return courses;
}

function importStudentsToSheet_(classrooms) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  
  classrooms.forEach(function(course, index) {
    try {
      var students = [];
      var pageToken = null;
      
      do {
        var response = Classroom.Courses.Students.list(course.id, {
          pageSize: 100,
          pageToken: pageToken
        });
        
        if (response.students) {
          students = students.concat(response.students);
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
      
      if (configSheet) {
        configSheet.getRange(3 + index, 6).setValue(students.length);
      }
      
    } catch (error) {
      console.error('Error importing students from ' + course.name + ':', error);
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 14: GRADING DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

function generatePowerSchoolCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.GRADING);
  if (!sheet) return;
  
  // 1. Read Config
  var assignmentName = sheet.getRange('B2').getValue();
  var dueDate = sheet.getRange('B3').getValue();
  var pointsPossible = sheet.getRange('B4').getValue();
  var extraPoints = 0.0;
  var scoreType = sheet.getRange('B5').getValue();
  
  var periodMap = {};
  var periodData = sheet.getRange('D3:E9').getValues();
  periodData.forEach(function(row) {
    if (row[1]) periodMap[row[0].replace('Period ', '')] = row[1];
  });
  
  if (!assignmentName) {
    SpreadsheetApp.getUi().alert('❌ Missing Config', 'Enter Assignment Name.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 2. Calculate Date Range (Current Week: Mon-Sun)
  var now = new Date();
  var day = now.getDay();
  var diff = now.getDate() - day + (day == 0 ? -6 : 1); // Monday
  var monday = new Date(now.setDate(diff));
  monday.setHours(0,0,0,0);
  
  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23,59,59,999);
  
  // 3. Get Student Data (Progress + Hidden Grades + Hidden Bonus)
  // We need to iterate ALL students who have data.
  // Best source: Hidden Grades DB + Hidden Bonus DB unique emails.
  var studentScores = {}; // { email: { mp: 0, bonus: 0, name: '', period: '' } }
  
  // A. Grades DB (Assignments)
  var gradesSheet = ss.getSheetByName("hidden_grades_db");
  if (gradesSheet && gradesSheet.getLastRow() >= 2) {
    var gData = gradesSheet.getRange(2, 1, gradesSheet.getLastRow() - 1, 9).getValues();
    gData.forEach(function(row) {
      // row[8] = Updated/GradeTime. Check if in range.
      var date = new Date(row[8]);
      if (date >= monday && date <= sunday) {
        var email = row[0];
        if (!studentScores[email]) studentScores[email] = { mp: 0, bonus: 0, name: row[2], period: row[3] };
        studentScores[email].mp += (row[7] || 0);
      }
    });
  }
  
  // B. Bonus DB (Rations)
  var bonusSheet = ss.getSheetByName(CONFIG.SHEETS.BONUS_DB);
  if (bonusSheet && bonusSheet.getLastRow() >= 2) {
    var bData = bonusSheet.getRange(2, 1, bonusSheet.getLastRow() - 1, 3).getValues();
    bData.forEach(function(row) {
      // row[0] = Timestamp
      var date = new Date(row[0]);
      if (date >= monday && date <= sunday) {
        var email = row[1];
        // If student exists from grades, update. If not, try to find name/period later (or skip if unknown).
        if (studentScores[email]) {
          studentScores[email].bonus += (row[2] || 0);
        } else {
          // Edge case: Student has ONLY bonus points this week. Need to look up info.
          // Try finding info from Progress sheet cache
          var info = lookupStudentInfo_(email);
          if (info) {
            studentScores[email] = { mp: 0, bonus: (row[2] || 0), name: info.name, period: info.period };
          }
        }
      }
    });
  }
  
  // 4. Generate CSV (One Block Per Period in Column A)
  var periodGroups = {};
  Object.keys(studentScores).forEach(function(email) {
    var s = studentScores[email];
    var p = s.period ? s.period.toString() : 'Unknown';
    if (!periodGroups[p]) periodGroups[p] = [];
    periodGroups[p].push(s);
  });
  
  // Clear previous output
  var lastRow = sheet.getLastRow();
  if (lastRow > 11) sheet.getRange(12, 1, lastRow - 11, 1).clearContent();
  
  var currentRow = 12;
  var periods = Object.keys(periodGroups).sort();
  
  periods.forEach(function(period) {
    var teacherName = periodMap[period] || 'Unknown Teacher';
    var students = periodGroups[period];
    var csvBlock = '';
    
    // CSV Header
    csvBlock += 'Teacher Name:,' + teacherName + ',\n';
    csvBlock += 'Class:,' + 'Period ' + period + ',\n'; 
    csvBlock += 'Assignment Name:,' + assignmentName + ',\n';
    var dStr = (dueDate instanceof Date) ? dueDate.toISOString().slice(0, 10) : dueDate;
    csvBlock += 'Due Date:,' + dStr + ',\n';
    csvBlock += 'Points Possible:,' + pointsPossible + ',\n';
    csvBlock += 'Extra Points:,' + extraPoints + ',\n';
    csvBlock += 'Score Type:,' + scoreType + ',\n';
    csvBlock += 'Student Num,Student Name,Score\n';
    
    // Student Rows
    students.forEach(function(s) {
      var idMatch = s.email ? s.email.match(/\d+/) : null;
      var id = idMatch ? idMatch[0] : '00000';
      var score = s.mp + s.bonus;
      var safeName = '"' + s.name.replace(/"/g, '""') + '"';
      csvBlock += id + ',' + safeName + ',' + score + '\n';
    });
    
    // Write Block Title
    sheet.getRange(currentRow, 1).setValue('📂 Period ' + period + ' CSV (Copy content below):')
         .setFontWeight('bold').setFontColor(CONFIG.COLORS.SECONDARY);
    currentRow++;
    
    // Write CSV Content
    sheet.getRange(currentRow, 1).setValue(csvBlock)
         .setFontFamily('Consolas').setVerticalAlignment('top').setWrap(true)
         .setBackground(CONFIG.COLORS.BG_ROW_1).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    
    // Add extra spacing rows for visual separation if content is long, but cell auto-resizes.
    // We just move to next row.
    currentRow += 2; 
  });
  
  SpreadsheetApp.getUi().alert('✅ Generated', 'CSV data for ' + periods.length + ' periods generated below.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function lookupStudentInfo_(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (pSheet && pSheet.getLastRow() >= 3) {
    var data = pSheet.getRange(3, 1, pSheet.getLastRow() - 2, 3).getValues(); // Email, Name, Period
    for(var i=0; i<data.length; i++) {
      if (data[i][0] === email) {
        return { name: data[i][1], period: data[i][2], email: email };
      }
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 13: RATIONS & BONUS SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

function addToBonusDB_(email, amount, week) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.BONUS_DB);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.BONUS_DB);
    sheet.hideSheet();
    sheet.appendRow(['Timestamp', 'Email', 'Amount', 'Week']);
  }
  
  sheet.appendRow([new Date(), email, amount, week]);
}

function refreshRationsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RATIONS);
  if (!sheet) return;
  
  // Clear data but keep headers area
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) sheet.getRange(3, 1, lastRow - 2, 5).clearContent().removeCheckboxes();
  
  // 1. Tally Inventory
  var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_DB);
  var rationCounts = {};
  
  if (invSheet && invSheet.getLastRow() >= 2) {
    var invData = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 3).getValues(); // [Time, Email, Item]
    invData.forEach(function(row) {
      var email = row[1];
      var item = row[2];
      if (item === '33MP Ration') {
        if (!rationCounts[email]) rationCounts[email] = 0;
        rationCounts[email]++;
      } else if (item === 'Consumed 33MP Ration') {
        if (!rationCounts[email]) rationCounts[email] = 0;
        rationCounts[email]--;
      }
    });
  }
  
  // 2. Get Names
  var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  var nameMap = {};
  if (progressSheet && progressSheet.getLastRow() >= 3) {
    var pData = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 2).getValues();
    pData.forEach(function(r) { nameMap[r[0]] = r[1]; });
  }
  
  // 3. Build Rows
  var output = [];
  var emails = Object.keys(rationCounts);
  
  emails.forEach(function(email) {
    if (rationCounts[email] > 0) {
      output.push([
        email,                    // Col A: Email
        nameMap[email] || email,  // Col B: Name
        rationCounts[email],      // Col C: Count
        false,                    // Col D: Checkbox
        ''                        // Col E: Last Consumed
      ]);
    }
  });
  
  // Sort by Name (Col B)
  output.sort(function(a, b) {
    var nameA = a[1].toString().toLowerCase();
    var nameB = b[1].toString().toLowerCase();
    return nameA.localeCompare(nameB);
  });
  
  // Update Headers to match structure
  sheet.getRange('A2:E2').setValues([['Student Email', 'Student Name', 'Rations Owned', 'Consume (33 MP)', 'Status']]);
  applyHeaderStyle_(sheet.getRange('A2:E2'));
  
  if (output.length > 0) {
    sheet.getRange(3, 1, output.length, 5).setValues(output);
    sheet.getRange(3, 4, output.length, 1).insertCheckboxes();
  } else {
    sheet.getRange('A3').setValue('No students have rations.');
  }
}

function consumeRations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RATIONS);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  
  var data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  
  data.forEach(function(row, index) {
    var isChecked = row[3]; // Col D (Index 3)
    if (isChecked === true) {
      var email = row[0];
      var currentCount = row[2];
      var rowIndex = index + 3;
      
      if (currentCount > 0) {
        // Calculate Week (Monday of current week)
        var d = new Date();
        var day = d.getDay();
        var diff = d.getDate() - day + (day == 0 ? -6 : 1);
        var monday = new Date(d.setDate(diff));
        var weekStr = monday.toISOString().slice(0, 10);
        
        try {
          // Consume and Grant
          addToInventoryDB_(email, 'Consumed 33MP Ration');
          addToBonusDB_(email, 33, weekStr);
          
          sheet.getRange(rowIndex, 5).setValue('✅ Consumed for week ' + weekStr);
          sheet.getRange(rowIndex, 3).setValue(currentCount - 1);
          
        } catch (e) {
          console.error(e);
          sheet.getRange(rowIndex, 5).setValue('❌ Error');
        }
      } else {
        sheet.getRange(rowIndex, 5).setValue('❌ No rations left');
      }
      
      sheet.getRange(rowIndex, 4).uncheck();
    }
  });
}

function importStudents() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var classrooms = loadClassrooms_();
    importStudentsToSheet_(classrooms);
    
    ui.alert('✅ Import Complete', 
      'Updated student counts for ' + classrooms.length + ' classroom(s).',
      ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ Import Error', error.message, ui.ButtonSet.OK);
  }
}

function addSampleTracks_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  if (!sheet) return;
  
  // SAFEGUARD: If rows exist beyond the header (row 2), do NOT add samples.
  if (sheet.getLastRow() > 2) {
    console.log('Tracks sheet already has data. Skipping sample generation.');
    return;
  }
  
  // Sample Data: 3 Distinct Branches for Multimedia/Tech
  var sampleTracks = [
    // Sound Studio Branch
    ['SND', 'Audio Engineering', 'Sound', 0, 160, '', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Starter path for Audio'],
    ['SND', 'Sound Design I', 'Sound', 1, 160, 'SND-L0', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Requires Level 0'],
    
    // Visual Studio Branch
    ['VIS', 'Graphic Design', 'Visual', 0, 160, '', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Starter path for Visuals'],
    ['VIS', 'Digital Art I', 'Visual', 1, 160, 'VIS-L0', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Requires Level 0'],
    
    // Interactive Studio Branch
    ['DEV', 'Game Development', 'Interactive', 0, 160, '', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Starter path for Coding'],
    ['DEV', 'Game Logic I', 'Interactive', 1, 160, 'DEV-L0', 'TRUE', 'UNLOCKED_ONLY', 'Power', 'Requires Level 0']
  ];
  
  sheet.getRange(3, 1, sampleTracks.length, 10).setValues(sampleTracks);
  
  // Apply visual banding (Dark Mode style)
  applyTableBanding_(sheet, 3, 10);
}
function addSampleAssignments_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);
  if (!sheet) return;

  if (sheet.getLastRow() > 2) {
    console.log('Assignment sheet already has data. Skipping sample generation.');
    return;
  }

  // Sample Assignments: Matched to the new column order
  // [Track, Level, Week, Title, MaxMP, Topic, Desc, ViewMat, CopyMat, WebID, Status]
  var sampleAssignments = [
    // Sound Examples (With Website ID)
    ['SND', 0, 'A', 'Intro to Soundwaves', 100, 'Sound Studio', 'Watch the video.', 'https://www.youtube.com/watch?v=example', '', 'snd-01', '⏳ Ready'],
    ['SND', 0, 'B', 'Record Your Intro', 100, 'Sound Studio', 'Record a 30-second intro.', '', '', 'snd-02', '⏳ Ready'],
    
    // Visual Example (Blank Website ID -> No Link)
    ['VIS', 0, 'A', 'Elements of Art', 100, 'Visual Studio', 'Identify line and shape.', 'https://drive.google.com/file/d/example/view', '', '', '⏳ Ready'],
    
    // Interactive Example (Blank Website ID -> No Link)
    ['DEV', 0, 'A', 'Hello World', 100, 'Interactive Studio', 'Write your first code.', '', '', '', '⏳ Ready']
  ];

  // Write to columns A-K (Indices 1-11)
  sheet.getRange(3, 1, sampleAssignments.length, 11).setValues(sampleAssignments);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 5: GRADE SYNC & PROGRESS CALCULATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * REWRITTEN SYNC (SELF-HEALING): 
 * 1. Clears old triggers.
 * 2. Fetches grades until time runs out.
 * 3. If timed out: SAVES progress, SETS trigger for 30s later, and EXITS.
 * 4. If finished: SAVES progress, CALCULATES final stats, and STOPS.
 */
function syncGrades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();
  var MAX_EXECUTION_TIME = 1000 * 60 * 4.5; // Stop after 4.5 minutes

  // 1. CLEANUP: Delete any old "restart" triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncGrades' && 
        triggers[i].getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  try {
    // 2. PREP: Get active classrooms
    var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    var activeClassrooms = [];
    if (classroomsSheet.getLastRow() >= 3) {
      activeClassrooms = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 6).getValues()
        .filter(function(row) { return row[4] === 'TRUE' || row[4] === true; });
    }
    if (activeClassrooms.length === 0) return;

    // 3. QUEUE: Build the Master Assignment List
    var queue = [];
    var syncDb = getSyncDatabase_(); // Load "Hidden Sync Timestamp DB"

    activeClassrooms.forEach(function(classroom) {
      var courseId = classroom[0];
      var period = classroom[3] || '?';
      try {
        var coursework = getAllCoursework_(courseId); 
        coursework.forEach(function(work) {
          if (work.state !== 'PUBLISHED') return;
          var parsed = parseAssignmentName_(work.title);
          if (!parsed) return;
          
          var lastSync = syncDb[work.id] || 0; 
          queue.push({
            id: work.id,
            courseId: courseId,
            title: work.title,
            period: period,
            parsed: parsed,
            lastSync: lastSync 
          });
        });
      } catch (e) { console.error('Error listing course ' + courseId, e); }
    });

    // SORT: Oldest ("Never Synced") First
    queue.sort(function(a, b) { return a.lastSync - b.lastSync; });
    console.log('Sync Queue Loaded: ' + queue.length + ' assignments.');

    // 4. PROCESS: The Work Loop
    var updatesForDB = []; // Rows to save to Grade DB
    var processedIDs = []; // IDs to update in Sync DB
    var studentCache = {};
    var isTimedOut = false;

    for (var i = 0; i < queue.length; i++) {
      // ⏰ TIMEOUT CHECK
      if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
        console.warn('⚠️ Time limit reached. Saving batch and auto-scheduling.');
        ss.toast('⏳ Saving progress & auto-resuming in 30s...', 'Sync Paused', 10);
        
        // A. Save what we have so far
        if (processedIDs.length > 0) {
          if (updatesForDB.length > 0) saveToGradeDatabase_(updatesForDB);
          updateSyncDatabase_(processedIDs);
        }

        // B. Create Trigger to restart
        ScriptApp.newTrigger('syncGrades')
                 .timeBased()
                 .after(1000 * 30) // 30 seconds
                 .create();
        
        isTimedOut = true;
        return; // EXIT FUNCTION HERE
      }

      var task = queue[i];
      
      // Cache student lists
      if (!studentCache[task.courseId]) {
        studentCache[task.courseId] = getAllStudents_(task.courseId);
      }
      var students = studentCache[task.courseId];

      try {
        // Fetch Submissions
        var submissions = fetchWithBackoff_(function() {
          return getSubmissions_(task.courseId, task.id);
        });

        submissions.forEach(function(sub) {
          var grade = sub.assignedGrade;
          if (grade === null || grade === undefined) return;

          var student = students.find(function(s) { return s.userId === sub.userId; });
          if (!student) return;

          // SAFETY CHECK: Prevent collisions by skipping students with no email
          if (!student.profile.emailAddress) {
            console.warn('⚠️ SKIPPING: Student ' + (student.profile.name ? student.profile.name.fullName : sub.userId) + ' has no email address. Check permissions.');
            return; // Skip this iteration to prevent overwriting other students
          }

          updatesForDB.push([
            student.profile.emailAddress,   // Col 1: Email
            task.id,                        // Col 2: AssignID
            student.profile.name.fullName,  // Col 3: Name
            task.period,                    // Col 4: Period
            task.parsed.trackCode,          // Col 5: Track
            task.parsed.level,              // Col 6: Level
            task.parsed.week,               // Col 7: Week
            grade,                          // Col 8: MP
            new Date()                      // Col 9: Updated
          ]);
        });

        processedIDs.push({ id: task.id, courseId: task.courseId });
      } catch (e) {
        console.error('Failed to sync assignment ' + task.id, e);
      }
    }

    // 5. FINALIZE (Only reached if NO timeout)
    if (processedIDs.length > 0) {
      if (updatesForDB.length > 0) saveToGradeDatabase_(updatesForDB);
      updateSyncDatabase_(processedIDs);
    }

    // 6. COMPILE PROGRESS (Only runs when queue is fully cleared to save time)
    ss.toast('Compiling final progress...', '📈 Calculating', 10);
    compileProgressFromDB_(); 

    // 7. Update Dashboard Timestamp
    var now = new Date();
    PropertiesService.getScriptProperties().setProperty(CONFIG.PROPS.LAST_SYNC, now.toISOString());
    var dashSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
    if (dashSheet) {
      dashSheet.getRange('E4:F4').merge().setValue('LAST SYNC: ' + now.toLocaleTimeString());
    }
    
    ss.toast('Sync Complete!', '✅ Success', 5);

  } catch (error) {
    console.error('Sync error:', error);
    ss.toast('Error during sync. Check logs.', '❌ Error', 20);
  }
}

/**
 * HELPER: Upserts grades into the hidden database
 */


/**
 * HELPER: Reads full history and updates Progress tab
 */


/**
 * Get all coursework from a classroom
 * Includes both PUBLISHED and DRAFT assignments
 */
function getAllCoursework_(courseId) {
  var coursework = [];
  var pageToken = null;
  
  do {
    var response = Classroom.Courses.CourseWork.list(courseId, {
      pageSize: 100,
      pageToken: pageToken,
      courseWorkStates: ['PUBLISHED', 'DRAFT']
    });
    
    if (response.courseWork) {
      coursework = coursework.concat(response.courseWork);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return coursework;
}
// -------------------------------------------------------------------------
// GLOBAL CACHE (Optimizes speed by storing student lists temporarily)
// -------------------------------------------------------------------------
var STUDENT_CACHE = {}; 

function getCachedStudents_(courseId) {
  // If we haven't fetched this class yet, fetch it now
  if (!STUDENT_CACHE[courseId]) {
    STUDENT_CACHE[courseId] = getAllStudents_(courseId);
  }
  // Return the list we already have in memory
  return STUDENT_CACHE[courseId];
}
function getAllStudents_(courseId) {
  // Use enriched version that fills in missing emails from student registry
  return getAllStudentsEnriched_(courseId);
}

function getSubmissions_(courseId, courseworkId) {
  var submissions = [];
  var pageToken = null;
  
  do {
    var response = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, courseworkId, {
      pageSize: 100,
      pageToken: pageToken
    });
    
    if (response.studentSubmissions) {
      submissions = submissions.concat(response.studentSubmissions);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return submissions;
}

function parseAssignmentName_(title) {
  // NEW FLEXIBLE REGEX:
  // 1. (\d+)([A-D])(\d+)  -> Looks for Level + Week + Num (e.g., 0A1) anywhere.
  // 2. .* -> Any characters in between
  // 3. \[([A-Z]{2,4})\]   -> Looks for [TRACK] tag (e.g., [SND])
  // 4. 'i' flag           -> Case insensitive (allows 0a1 or 0A1)
  
  var regex = /(\d+)([A-D])(\d+).*\[([A-Z]{2,4})\]/i;
  var match = title.match(regex);
  
  if (!match) return null;
  
  // Note indices change slightly because we removed the start anchor (^)
  // match[1] = Level
  // match[2] = Week
  // match[3] = Number
  // match[4] = Track Code
  
  return {
    trackCode: match[4].toUpperCase(), // Ensure capitalized
    level: parseInt(match[1]),
    week: match[2].toUpperCase(),      // Ensure capitalized
    num: parseInt(match[3])
  };
}

function parsePrerequisites_(prereqString) {
  // New Structure: Returns an array of 'groups'.
  // Each group is an array of 'requirements'.
  // Logic: (Group1_Req1 AND Group1_Req2) OR (Group2_Req1)
  
  if (!prereqString) return [];
  
  // 1. Split by OR (|) to get groups
  var groups = prereqString.split('|');
  
  return groups.map(function(groupStr) {
    // 2. Split each group by AND (,) to get individual requirements
    return groupStr.split(',').map(function(s) { 
      return s.trim(); 
    }).filter(function(s) { 
      return s.length > 0; 
    });
  });
}

function buildTrackMap_(tracksData) {
  var map = {};
  
  tracksData.forEach(function(row) {
    var code = row[0];
    var level = row[3];
    var key = code + '-L' + level;
    var prereqString = row[5] ? row[5].toString().trim() : '';
    
    // UPDATED: Use the new parser which returns groups of requirements
    var parsedPrereqs = parsePrerequisites_(prereqString);
    
    // Parse Track Type tags from column 9 (index 8)
    var trackTypeString = row[8] ? row[8].toString().trim() : '';
    var trackTags = [];
    if (trackTypeString) {
      trackTags = trackTypeString.split(',').map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });
    }
    
    map[key] = {
      code: code,
      name: row[1],
      studio: row[2],
      level: level,
      unlockMP: row[4] || CONFIG.DEFAULTS.UNLOCK_THRESHOLD,
      
      // NEW LOGIC: We just store the parsed groups here.
      // The checking function will handle interpreting "MIN_MP", "STUDIO_COUNT", etc.
      prerequisites: parsedPrereqs,      
      
      active: row[6] === 'TRUE' || row[6] === true,
      assignTo: row[7],
      trackTags: trackTags
    };
  });
  
  return map;
}

/**
 * Aggregate progress data with Week A/B MP breakdown
 * UPDATED: Includes trackTags to distinguish Power/Hero vs Theory/Events
 */
function aggregateProgress_(progressData, trackMap) {
  var aggregated = {};
  var now = new Date();
  var dayOfWeek = now.getDay();
  var mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  var weekStart = new Date(now);
  weekStart.setDate(now.getDate() + mondayOffset);
  weekStart.setHours(0, 0, 0, 0);

  progressData.forEach(function(entry) {
    var key = entry.email + '|' + entry.trackCode + '|' + entry.level;
    
    if (!aggregated[key]) {
      var trackKey = entry.trackCode + '-L' + entry.level;
      var trackInfo = trackMap[trackKey] || {};
      
      var unlockThreshold = trackInfo.unlockMP || CONFIG.DEFAULTS.UNLOCK_THRESHOLD;
      
      aggregated[key] = {
        email: entry.email,
        name: entry.name,
        period: entry.period,
        trackCode: entry.trackCode,
        trackName: trackInfo.name || entry.trackCode,
        trackTags: trackInfo.trackTags || [], // Store tags for logic check
        level: entry.level,
        totalMP: 0,
        weekMP: 0,
        weekAMP: 0,
        weekBMP: 0,
        lastActivity: null,
        unlockMP: unlockThreshold
      };
    }
    
    aggregated[key].totalMP += entry.mp;
    
    // Track Week A vs Week B+ MP
    if (entry.week === 'A') {
      aggregated[key].weekAMP += entry.mp;
    } else {
      // B, C, D all count as "Week B" (Power Test / project work / Hero Trials)
      aggregated[key].weekBMP += entry.mp;
    }
    
    var gradeDate = new Date(entry.gradeTime);
    if (gradeDate >= weekStart) {
      aggregated[key].weekMP += entry.mp;
    }
    
    var activityDate = new Date(entry.gradeTime);
    if (!aggregated[key].lastActivity || activityDate > aggregated[key].lastActivity) {
      aggregated[key].lastActivity = activityDate;
    }
  });
  
  return Object.values(aggregated);
}

/**
 * Write progress data to sheet
 * UPDATED: 
 * - Power/Hero Tracks: Require (Unlock MP / 2) in Week A AND Week B to unlock.
 * - Theory/Events: Never show "Ready to Unlock" based on MP.
 */
/**
 * Write progress data to sheet
 * UPDATED: Uses GLOBAL COUNT Logic.
 * Checks total # of Power Wins vs Total # of Theory Wins.
 */
function writeProgressSheet_(progressData, trackMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, 12).clear();
  }
  
  if (progressData.length === 0) return;
  
  // Sort for readability
  progressData.sort(function(a, b) {
    if (a.email !== b.email) return a.email.localeCompare(b.email);
    return a.level - b.level;
  });

  // 1. CALCULATE GLOBAL COUNTS PER STUDENT
  var studentStats = {};

  progressData.forEach(function(p) {
    if (!studentStats[p.email]) {
      studentStats[p.email] = { powerWins: 0, theoryWins: 0, theoryLevels: {} };
    }
    
    // Check if this row is "Finished" (80/80)
    var isFinished = (p.weekAMP >= 80 && p.weekBMP >= 80);
    
    if (isFinished) {
      var isTheory = p.trackTags.some(function(t) { return t === 'Theory'; });
      
      if (isTheory) {
        // For theory, we group by level. (Theory A L1 + Theory B L1 = 1 Win)
        // We track which levels have passing grades. 
        // Note: Strict logic would check if BOTH A & B are done. 
        // Here we assume if the row is present and passed, it counts towards that level's bucket.
        if (!studentStats[p.email].theoryLevels[p.level]) {
          studentStats[p.email].theoryLevels[p.level] = [];
        }
        studentStats[p.email].theoryLevels[p.level].push(true);
      } else {
        // For Power, every track is a unique win
        studentStats[p.email].powerWins++;
      }
    }
  });

  // Consolidate Theory Wins (Must have at least 2 passed tracks per level to count as a "Win")
  // Adjust '2' to '1' if you only have 1 theory track per level.
  Object.keys(studentStats).forEach(function(email) {
    var levels = studentStats[email].theoryLevels;
    Object.keys(levels).forEach(function(lvl) {
      if (levels[lvl].length >= 2) { // Assuming Theory A + Theory B
        studentStats[email].theoryWins++;
      }
    });
  });

  var rows = progressData.map(function(p) {
    var statusMsg = '';
    var isTheory = p.trackTags.some(function(t) { return t === 'Theory'; });
    var isFinished = (p.weekAMP >= 80 && p.weekBMP >= 80);
    
    if (isTheory) {
      statusMsg = isFinished ? '✅ COMPLETED' : '🛠️ IN PROGRESS';
    } else {
      // POWER TRACK LOGIC
      if (isFinished) {
        // It's done. Can we unlock the NEXT one?
        // Check Global Balance.
        var stats = studentStats[p.email];
        
        // If I have 5 Power Wins and 4 Theory Wins... 
        // My debt is high. I need 5 Theory Wins to match my 5 Power Wins.
        if (stats.theoryWins >= stats.powerWins) {
           statusMsg = '🚀 AUTO-LEVELING';
        } else {
           // Calculate which theory level they actually need
           // If PowerWins = 5, they need Theory Win #5, which is likely Level 4 (if starting at 0)
           statusMsg = '🔒 NEEDS THEORY #' + (stats.powerWins);
        }
      } else {
        statusMsg = '🛠️ WORKING';
      }
    }
    
    var lastActivity = p.lastActivity ? p.lastActivity.toLocaleDateString() : '';
    
    return [
      p.email, p.name, p.period, p.trackCode, p.trackName,
      'Level ' + p.level, p.totalMP, p.weekMP, p.weekAMP, p.weekBMP,
      statusMsg, lastActivity
    ];
  });

  sheet.getRange(3, 1, rows.length, 12).setValues(rows);
  
  // Conditional Formatting (Same as before)
  var range = sheet.getRange(3, 11, rows.length, 1);
  sheet.clearConditionalFormatRules();
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('🚀 AUTO-LEVELING').setBackground('#064e3b').setFontColor('#34d399').setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('🔒').setBackground('#451a03').setFontColor('#fbbf24').setBold(true).setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
  
  // Banding
  for (var i = 0; i < rows.length; i++) {
    var rowNum = 3 + i;
    var bgColor = (i % 2 === 0) ? CONFIG.COLORS.ROW_NORMAL : CONFIG.COLORS.ROW_ALT;
    sheet.getRange(rowNum, 1, 1, 12).setBackground(bgColor);
  }
}
// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: UNLOCK LOGIC (STRICT LOCKSTEP SYSTEM)
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: UNLOCK LOGIC (UPDATED: 2 THEORIES & LOCKSTEP)
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: UNLOCK LOGIC (GLOBAL COUNT SYSTEM)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * GLOBAL COUNT UNLOCKER
 * 1. Counts Total Power Nodes Completed (80/80).
 * 2. Counts Total Theory Levels Completed (80/80 in A & B).
 * 3. Unlocks the FIRST valid Power Node where Theory Count >= Power Count.
 */
function runUnlockCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // 1. LOAD ALL DATA
    var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
    var progressData = progressSheet.getLastRow() >= 3 ?
      progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 12).getValues() : [];
    
    var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
    var tracksData = tracksSheet.getLastRow() >= 3 ? 
      tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 10).getValues() : [];
    var trackMap = buildTrackMap_(tracksData);

    var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    var activeClassrooms = [];
    var classroomsLastRow = classroomsSheet.getLastRow();
    if (classroomsSheet.getLastRow() >= 3) {
      activeClassrooms = classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues()
        .filter(function(row) { return row[4] === 'TRUE' || row[4] === true; });
    }

    // 2. CACHE COURSEWORK
    var courseWorkCache = {};
    activeClassrooms.forEach(function(classroom) {
      try {
        courseWorkCache[classroom[0]] = getAllCoursework_(classroom[0]);
      } catch (e) { console.error('Error caching course ' + classroom[0], e); }
    });

    // 3. MAP STUDENT DATA & CALCULATE COUNTS
    var studentMap = {};

    progressData.forEach(function(row) {
      var email = row[0];
      if (!studentMap[email]) {
        studentMap[email] = {
          rows: [],
          name: row[1],
          powerWins: 0,
          theoryLevels: {} // { 0: 2, 1: 2 } -> Level: Count of passed tracks
        };
      }
      studentMap[email].rows.push(row);
      
      // Calculate Stats Inline
      var trackCode = row[3];
      var level = parseInt(row[5].replace('Level ', ''));
      var trackKey = trackCode + '-L' + level;
      var config = trackMap[trackKey];
      
      if (config) {
        var isFinished = (row[8] >= 80 && row[9] >= 80);
        var isTheory = config.trackTags.indexOf('Theory') !== -1;
        
        if (isFinished) {
          if (isTheory) {
            if (!studentMap[email].theoryLevels[level]) studentMap[email].theoryLevels[level] = 0;
            studentMap[email].theoryLevels[level]++;
          } else {
            studentMap[email].powerWins++;
          }
        }
      }
    });

    // 4. MAIN CHECK LOOP
    var unlocks = [];
    var processedStudents = []; 

    Object.keys(studentMap).forEach(function(email) {
      if (processedStudents.indexOf(email) !== -1) return; // Only 1 unlock per student per sync

      var student = studentMap[email];
      
      // Calculate Total Theory Wins (Levels where they passed 2 tracks, e.g. A & B)
      var theoryWins = 0;
      Object.keys(student.theoryLevels).forEach(function(lvl) {
        if (student.theoryLevels[lvl] >= 2) theoryWins++; // Assumes 2 theory tracks per level
      });

      // Find a Power Track that is finished (80/80) but can progress
      for (var i = 0; i < student.rows.length; i++) {
        var row = student.rows[i];
        var trackCode = row[3];
        var level = parseInt(row[5].replace('Level ', ''));
        var weekA = row[8];
        var weekB = row[9];
        
        var trackKey = trackCode + '-L' + level;
        var config = trackMap[trackKey];
        
        // Filter candidates: Must be Power Track, Must be Finished
        if (config && config.trackTags.indexOf('Theory') === -1 && weekA >= 80 && weekB >= 80) {
           
           // THE GOLDEN RULE:
           // Theory Wins must be >= Power Wins to unlock the NEXT thing.
           // Note: 'student.powerWins' includes the track we are currently looking at.
           
           if (theoryWins >= student.powerWins) {
             
             // Check if next level exists
             var nextLevel = level + 1;
             var nextKey = trackCode + '-L' + nextLevel;
             var nextTrack = trackMap[nextKey];
             
             if (nextTrack && nextTrack.active) {
               unlocks.push({
                 email: email,
                 name: student.name,
                 trackCode: trackCode,
                 toLevel: nextLevel,
                 type: 'Power',
                 reason: 'Global Count: ' + theoryWins + ' Theory vs ' + student.powerWins + ' Power'
               });
               processedStudents.push(email);
               break; // Stop scanning for this student
             }
           }
        }
      }
    });

    if (unlocks.length === 0) return;

    // 5. EXECUTE UNLOCKS
    var logEntries = [];
    var uniqueUnlocks = []; // Dedupe
    var seenKeys = {};
    
    unlocks.forEach(function(u) {
      var key = u.email + '|' + u.trackCode + '|' + u.toLevel;
      if (!seenKeys[key]) { seenKeys[key] = true; uniqueUnlocks.push(u); }
    });

    uniqueUnlocks.forEach(function(unlock) {
      try {
        var assignments = findAssignmentsForLevelCached_(activeClassrooms, courseWorkCache, unlock.trackCode, unlock.toLevel);
        var anySuccess = false;

        assignments.forEach(function(assignment) {
           var success = addStudentToAssignment_(assignment.courseId, assignment.assignmentId, unlock.email, assignment.state);
           if (success) anySuccess = true;
        });

        if (anySuccess) {
          logEntries.push([
            new Date(), unlock.email, unlock.name, 
            unlock.trackCode + '-L' + unlock.toLevel, 
            unlock.type + ' Auto', 
            unlock.reason
          ]);
        }
      } catch (e) { console.error(e); }
    });

    if (logEntries.length > 0) {
      var logSheet = ss.getSheetByName(CONFIG.SHEETS.UNLOCK_LOG);
      var logLastRow = Math.max(2, logSheet.getLastRow());
      logSheet.getRange(logLastRow + 1, 1, logEntries.length, 6).setValues(logEntries);
    }
    
  } catch (error) {
    console.error('Unlock error:', error);
  }
}

// HELPER: Find assignments using the cache (avoids API calls)
function findAssignmentsForLevelCached_(classrooms, cache, trackCode, level) {
  var assignments = [];
  classrooms.forEach(function(classroom) {
    var courseId = classroom[0];
    var coursework = cache[courseId] || []; // Use cache!
    
    coursework.forEach(function(work) {
      var parsed = parseAssignmentName_(work.title);
      if (!parsed) return;
      if (parsed.trackCode === trackCode && parsed.level === level) {
        assignments.push({
          courseId: courseId,
          assignmentId: work.id,
          title: work.title,
          state: work.state
        });
      }
    });
  });
  return assignments;
}

/**
 * HELPER: Calculates the Highest COMPLETED Level for Power and Theory per studio.
 * Returns: { "Sound": { power: 0, theory: 0 }, "Visual": { power: -1, theory: 0 } }
 * (-1 means they haven't finished Level 0 yet)
 */
function calculateStudioLevels_(email, progressData, trackMap) {
  var levels = {};
  
  var studentWork = progressData.filter(function(row) { return row[0] === email; });
  
  studentWork.forEach(function(row) {
    var trackCode = row[3];
    var currentLevel = parseInt(row[5].replace('Level ', ''));
    var readyToUnlock = row[10] === '✅ YES'; // Finished?
    var trackKey = trackCode + '-L' + currentLevel;
    var trackInfo = trackMap[trackKey];
    
    if (!trackInfo) return;
    
    var studio = trackInfo.studio;
    if (!levels[studio]) levels[studio] = { power: -1, theory: -1 };
    
    var isTheory = trackInfo.trackTags.indexOf('Theory') !== -1;
    
    // If this level is marked "Ready to Unlock", it counts as COMPLETED.
    // We update the max level found.
    if (readyToUnlock) {
      if (isTheory) {
        if (currentLevel > levels[studio].theory) levels[studio].theory = currentLevel;
      } else {
        if (currentLevel > levels[studio].power) levels[studio].power = currentLevel;
      }
    } else {
      // Even if not finished with current, they finished previous.
      // (e.g. working on L1 means L0 is done)
      var completedPrev = currentLevel - 1;
      if (isTheory) {
        if (completedPrev > levels[studio].theory) levels[studio].theory = completedPrev;
      } else {
        if (completedPrev > levels[studio].power) levels[studio].power = completedPrev;
      }
    }
  });
  
  return levels;
}

function checkPrerequisites_(email, trackConfig, progressData, trackMap) {
  // If no prereqs defined (and parsed as empty array), return true
  var prereqGroups = trackConfig.prerequisites || []; 
  // Note: We need to update buildTrackMap_ to use the new parser output structure
  // mapped to 'prerequisites' property. See below.
  
  if (prereqGroups.length === 0) return true;

  // OR LOGIC: Loop through groups. If ANY group returns true, the student passes.
  return prereqGroups.some(function(group) {
    
    // AND LOGIC: Inside a group, ALL requirements must be met.
    return group.every(function(reqString) {
      
      // -- STUDIO COUNT (e.g., STUDIO_COUNT:Interactive:2) --
      if (reqString.indexOf('STUDIO_COUNT:') === 0) {
        var parts = reqString.split(':');
        if (parts.length === 3) {
          var studio = parts[1].trim();
          var count = parseInt(parts[2]);
          return countCompletedLevelsPerStudio_(email, studio, progressData, trackMap) >= count;
        }
        return false;
      }
      
      // -- MIN MP (e.g., MIN_MP:500) --
      if (reqString.indexOf('MIN_MP:') === 0) {
        var minParts = reqString.split(':');
        return getTotalStudentMP_(email, progressData) >= parseInt(minParts[1]);
      }

      // -- TAG COUNT (e.g., TAG_COUNT:Assessment:3) --
      if (reqString.indexOf('TAG_COUNT:') === 0) {
        var tagParts = reqString.split(':');
        return countCompletedTracksByTag_(email, tagParts[1].trim(), progressData, trackMap) >= parseInt(tagParts[2]);
      }

      // -- STANDARD TRACK LEVEL (e.g., SND-L1) --
      // This handles the "SND-L1" part
      return checkSingleTrackLevelPrereq_(email, reqString, progressData, trackMap);
    });
  });
}

function checkSingleTrackLevelPrereq_(email, prereqString, progressData, trackMap) {
  var match = prereqString.match(/^([A-Z]+)-L(\d+)$/);
  if (!match) return false;
  
  var reqTrack = match[1];
  var reqLevel = parseInt(match[2]);
  var reqKey = reqTrack + '-L' + reqLevel;
  var reqConfig = trackMap[reqKey];
  
  if (!reqConfig) return false;
  
  var studentProgress = progressData.find(function(row) {
    return row[0] === email && 
      row[3] === reqTrack && 
      parseInt(row[5].replace('Level ', '')) === reqLevel;
  });
  
  if (!studentProgress) return false;
  
  var mp = studentProgress[6];
  var threshold = reqConfig.unlockMP || CONFIG.DEFAULTS.UNLOCK_THRESHOLD;
  
  if (threshold === 0) return true;
  return mp >= threshold;
}

function countCompletedLevelsPerStudio_(email, studio, progressData, trackMap) {
  var count = 0;
  
  var studioTracks = Object.values(trackMap).filter(function(t) {
    // Avoid circular logic: Don't count theory tracks as completions for theory logic
    return t.studio === studio && (!t.studioCountPrereqs || t.studioCountPrereqs.length === 0);
  });
  
  studioTracks.forEach(function(track) {
    var studentProgress = progressData.find(function(row) {
      return row[0] === email &&
        row[3] === track.code &&
        parseInt(row[5].replace('Level ', '')) === track.level;
    });
    
    if (studentProgress) {
      var mp = studentProgress[6] || 0;
      var threshold = track.unlockMP || CONFIG.DEFAULTS.UNLOCK_THRESHOLD;
      
      if ((!threshold && mp > 0) || (threshold && mp >= threshold)) {
        count++;
      }
    }
  });
  
  return count;
}

/**
 * Get total MP across all tracks for a student
 */
function getTotalStudentMP_(email, progressData) {
  var total = 0;
  progressData.forEach(function(row) {
    if (row[0] === email) {
      total += (row[6] || 0);
    }
  });
  return total;
}

/**
 * Count completed track-levels with a specific tag for a student
 */
function countCompletedTracksByTag_(email, tag, progressData, trackMap) {
  var count = 0;
  
  var taggedTracks = Object.values(trackMap).filter(function(t) {
    return t.trackTags && t.trackTags.indexOf(tag) !== -1;
  });
  
  taggedTracks.forEach(function(track) {
    var studentProgress = progressData.find(function(row) {
      return row[0] === email &&
        row[3] === track.code &&
        parseInt(row[5].replace('Level ', '')) === track.level;
    });
    
    if (studentProgress) {
      var mp = studentProgress[6] || 0;
      var threshold = track.unlockMP || CONFIG.DEFAULTS.UNLOCK_THRESHOLD;
      
      if (threshold === 0 || mp >= threshold) {
        count++;
      }
    }
  });
  
  return count;
}

function getTheoryTracks_(trackMap) {
  return Object.values(trackMap).filter(function(t) {
    return (t.studioCountPrereqs && t.studioCountPrereqs.length > 0) ||
           (t.tagCountPrereqs && t.tagCountPrereqs.length > 0) ||
           t.minMP;
  });
}

function findAssignmentsForLevel_(classrooms, trackCode, level) {
  var assignments = [];
  
  classrooms.forEach(function(classroom) {
    var courseId = classroom[0];
    
    try {
      var coursework = getAllCoursework_(courseId);
      
      coursework.forEach(function(work) {
        var parsed = parseAssignmentName_(work.title);
        if (!parsed) return;
        
        if (parsed.trackCode === trackCode && parsed.level === level) {
          assignments.push({
            courseId: courseId,
            assignmentId: work.id,
            title: work.title,
            state: work.state
          });
        }
      });
      
    } catch (error) {
      console.error('Error finding assignments in ' + courseId + ':', error);
    }
  });
  
  return assignments;
}

function addStudentToAssignment_(courseId, assignmentId, studentEmail, assignmentState) {
  // Use cached list instead of calling API every time
  var students = getCachedStudents_(courseId);
  var student = students.find(function(s) {
    return s.profile.emailAddress === studentEmail;
  });

  if (!student) 
    return false;
  
  
  var studentId = student.userId;

  try {
    // WRAP IN BACKOFF TO PREVENT CRASHES
    return fetchWithBackoff_(function() {
        if (assignmentState === 'DRAFT') {
          // FIRST UNLOCKER: Change assignee mode, add student, then publish
          Classroom.Courses.CourseWork.modifyAssignees({
            assigneeMode: 'INDIVIDUAL_STUDENTS',
            modifyIndividualStudentsOptions: {
              addStudentIds: [studentId]
            }
          }, courseId, assignmentId);

          Classroom.Courses.CourseWork.patch(
            { state: 'PUBLISHED' },
            courseId,
            assignmentId,
            { updateMask: 'state' }
          );
          console.log('Published DRAFT assignment ' + assignmentId + ' with first student ' + studentEmail);
        } else {
          // SUBSEQUENT UNLOCKERS
          Classroom.Courses.CourseWork.modifyAssignees({
            assigneeMode: 'INDIVIDUAL_STUDENTS', 
            modifyIndividualStudentsOptions: {
              addStudentIds: [studentId]
            }
          }, courseId, assignmentId);
          console.log('Added ' + studentEmail + ' to published assignment ' + assignmentId);
        }
        return true;
    });
  } catch (e) {
    // Check if error is "Student already assigned" (Google throws 400 or 409 often)
    if (e.message.indexOf('already assigned') !== -1 || e.message.indexOf('409') !== -1) {
       return true; // Treat as success
    }
    console.error('Error adding ' + studentEmail + ' to ' + assignmentId + ': ' + e.message);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 7: TRACK SELECTION FORM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create the Track Selection Form and post to all Classrooms
 */
/**
 * UPDATED: Create the Track Selection Form (Verified Email Collection)
 */
function createTrackSelectionForm() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  
  // Check if form already exists
  var existingFormId = props.getProperty(CONFIG.PROPS.TRACK_FORM_ID);
  if (existingFormId) {
    var response = ui.alert(
      '⚠️ Form Already Exists',
      'A Track Selection Form has already been created.\n\n' +
      'Do you want to create a NEW form? (The old one will remain but won\'t be linked.)',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }
  
  // Get Level 0 tracks
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksLastRow = tracksSheet.getLastRow();
  if (tracksLastRow < 3) {
    ui.alert('⚠️ No Tracks', 'Add tracks to the Tracks sheet first.', ui.ButtonSet.OK);
    return;
  }
  
// Ensure we read 10 columns (to include Track Type at index 8)
  var tracksData = tracksSheet.getRange(3, 1, tracksLastRow - 2, 10).getValues();

  var level0Tracks = tracksData.filter(function(row) {
    var level = row[3];
    var active = (row[6] === 'TRUE' || row[6] === true);
    var typeTags = row[8] ? row[8].toString() : ''; // Column I is Index 8

    // RULE: Must be Level 0, Active, AND have the 'Power' tag
    // (This automatically excludes 'Theory' tracks unless you tagged them as both)
    return level === 0 && active && typeTags.indexOf('Power') !== -1;
  });

  if (level0Tracks.length === 0) {
    ui.alert('⚠️ No Level 0 Power Tracks', 'No active Level 0 tracks with the "Power" tag found.', ui.ButtonSet.OK);
    return;
  }
  
  var trackNames = level0Tracks.map(function(row) {
    return row[1]; // Track Name
  });
  
  try {
    ss.toast('Creating form...', '📝 Step 1/3', 10);
    
    // Create the form
    var form = FormApp.create('🎮 Multimedia Heroes — Choose Your Starting Track');
    
    // 1. UPDATED DESCRIPTION: Mentions automatic collection
    form.setDescription(
      'Welcome to Multimedia Heroes!\n\n' +
      'Pick ONE track to start your journey. This will unlock all starter content in that studio.\n\n' +
      '• Ensure you are signed in with your school account.\n' +
      '• Your email will be recorded automatically.'
    );
    
    // 2. EMAIL SETTINGS
    // setCollectEmail(true) enables collection.
    // setLimitOneResponsePerUser(true) forces login, which is required for "Verified" collection.
    form.setCollectEmail(true);
    form.setLimitOneResponsePerUser(true); 
    form.setAllowResponseEdits(false);
    
    // Add the track selection question
    var trackQuestion = form.addListItem();
    trackQuestion.setTitle('Which track do you want to start with?');
    trackQuestion.setChoiceValues(trackNames);
    trackQuestion.setRequired(true);
    
   // ... inside createTrackSelectionForm ...

    // Store form ID
    props.setProperty(CONFIG.PROPS.TRACK_FORM_ID, form.getId());
    
    // -------------------------------------------------------------------------
    // NEW: Link form and AUTO-HIDE the clutter tab
    // -------------------------------------------------------------------------
    ss.toast('Linking and hiding response sheet...', '📝 Step 2/3', 10);

    // 1. Get list of existing sheet IDs
    var currentSheets = ss.getSheets().map(function(s) { return s.getSheetId(); });

    // 2. Link the form (This creates the new "Form Responses" tab)
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // 3. Force the spreadsheet to update immediately
    SpreadsheetApp.flush();

    // 4. Find the new sheet that wasn't there before and hide it
    var allSheets = ss.getSheets();
    var newSheet = allSheets.find(function(s) {
      return currentSheets.indexOf(s.getSheetId()) === -1;
    });

    if (newSheet) {
      newSheet.hideSheet();
      // Optional: Rename it so you know what it is in the hidden list
      newSheet.setName("Hidden_Raw_Responses_" + form.getId().substring(0, 5));
    }
    // -------------------------------------------------------------------------

    ss.toast('Setting up trigger...', '📝 Step 2/3', 10);
    
    // Create form submit trigger
    setupFormSubmitTrigger_(form.getId());

// ... rest of the function ...
    
    ss.toast('Posting to Classrooms...', '📝 Step 3/3', 10);
    
    // Get active classrooms
    var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    var classroomsLastRow = classroomsSheet.getLastRow();
    var classroomData = classroomsLastRow >= 3 ? classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues() : [];
    var activeClassrooms = classroomData.filter(function(row) {
      return row[4] === 'TRUE' || row[4] === true;
    });
    
    var formUrl = form.getPublishedUrl();
    var editUrl = form.getEditUrl();
    var assignmentIds = {};
    
    // Post to each classroom
    activeClassrooms.forEach(function(classroom) {
      var courseId = classroom[0];
      
      try {
        var topicId = getOrCreateTopic_(courseId, '🚀 Start Here');
        
        var assignment = Classroom.Courses.CourseWork.create({
          title: '🚀 Choose Your Starting Track',
          description: 'Fill out this form to pick your starting track and unlock your first assignments!\n\n' +
                       'This is required before you can see any assignments.',
          workType: 'ASSIGNMENT',
          maxPoints: 0,
          state: 'PUBLISHED',
          topicId: topicId,
          assigneeMode: 'ALL_STUDENTS',
          materials: [{
            link: {
              url: formUrl,
              title: 'Track Selection Form'
            }
          }]
        }, courseId);
        
        assignmentIds[courseId] = assignment.id;
        
      } catch (error) {
        console.error('Error posting to classroom ' + courseId + ':', error);
      }
    });
    
    props.setProperty(CONFIG.PROPS.TRACK_FORM_ASSIGNMENT_IDS, JSON.stringify(assignmentIds));
    
    // UPDATED ALERT: Reminds you to check the "Verified" setting
    ui.alert(
      '✅ Form Created!',
      'Track Selection Form has been created and posted to ' + Object.keys(assignmentIds).length + ' classroom(s).\n\n' +
      '⚠️ IMPORTANT: VERIFY SETTINGS\n' +
      'Google often defaults email collection to "Manual Input" even when scripted.\n' +
      '1. Click the Edit Link below.\n' +
      '2. Go to Settings → Responses.\n' +
      '3. Ensure "Collect email addresses" is set to "Verified" (so students don\'t have to type it).\n\n' +
      'Edit Link:\n' + editUrl,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Error', 'Failed to create form:\n\n' + error.message, ui.ButtonSet.OK);
    console.error('Form creation error:', error);
  }
}

/**
 * Set up the form submit trigger
 */
function setupFormSubmitTrigger_(formId) {
  // Remove any existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onTrackSelectionSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  var form = FormApp.openById(formId);
  ScriptApp.newTrigger('onTrackSelectionSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}
/**
 * BATCH PROCESSOR: Checks raw form responses against processed log.
 * OPTIMIZED: Caches coursework to prevent API quotas from exploding.
 */
function processNewTrackSelections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();
  var MAX_EXECUTION_TIME = 1000 * 60 * 4.5; // Stop after 4.5 minutes
  
  console.log('Starting Track Selection Batch Process...');

  // 1. CLEANUP: Delete any previous "restart" triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processNewTrackSelections' && 
        triggers[i].getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 2. Locate Data Sources
  var rawSheet = findRawResponseSheet_();
  if (!rawSheet) {
    console.error('Could not find Raw Form Response sheet.');
    return;
  }
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.TRACK_SELECTIONS);
  
  // 3. Get All Raw Data
  var rawLastRow = rawSheet.getLastRow();
  if (rawLastRow < 2) return; 
  var rawData = rawSheet.getRange(2, 1, rawLastRow - 1, 3).getValues(); // [Timestamp, Email, TrackChoice]
  
  // 4. Get All Processed Keys (Email + Track Name)
  var logLastRow = logSheet.getLastRow();
  var processedKeys = []; // CHANGED: We now track "Email-Track" pairs
  
  if (logLastRow >= 3) {
    var logData = logSheet.getRange(3, 1, logLastRow - 2, 3).getValues(); // Read Col A, B, C
    // Create a unique signature for every finished entry: "student@school.edu-SoundTrack"
    processedKeys = logData.map(function(row) { 
      return row[1] + '-' + row[2]; 
    }); 
  }
  
  // 5. Filter: Who hasn't been processed for THIS SPECIFIC TRACK yet?
  var toProcess = rawData.filter(function(row) {
    var email = row[1];
    var trackChoice = row[2];
    var key = email + '-' + trackChoice;
    
    // CHANGED: Check if the Key (Email + Track) exists, not just the email
    return email && processedKeys.indexOf(key) === -1;
  });

  if (toProcess.length === 0) {
    console.log('No new track selections found.');
    ss.toast('All clear! No new selections.', '✅ Complete', 5);
    return;
  }

  console.log('Found ' + toProcess.length + ' new selections to process.');
  ss.toast('Processing ' + toProcess.length + ' students...', '🚀 Working', 30);

  // 6. Pre-load Classroom & Track Data
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksData = tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 10).getValues();
  
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var classroomsLastRow = classroomsSheet.getLastRow();
  var activeClassrooms = [];
  if (classroomsLastRow >= 3) {
     activeClassrooms = classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues()
    .filter(function(row) { return row[4] === 'TRUE' || row[4] === true; });
  }

  // --- CRITICAL FIX: CACHE ALL COURSEWORK HERE ---
  // We fetch all assignments ONCE per classroom, not once per student.
  var courseWorkCache = {};
  activeClassrooms.forEach(function(classroom) {
    try {
      // Use backoff just in case
      var work = fetchWithBackoff_(function() {
         return getAllCoursework_(classroom[0]);
      });
      courseWorkCache[classroom[0]] = work;
    } catch (e) { 
      console.error('Error caching course ' + classroom[0], e); 
      courseWorkCache[classroom[0]] = [];
    }
  });
  // ------------------------------------------------

  // 7. Processing Loop
  var newLogEntries = [];
  var isTimedOut = false;

  for (var i = 0; i < toProcess.length; i++) {
    // ⏰ TIMEOUT CHECK
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      console.warn('⚠️ Time limit reached. Scheduling auto-restart.');
      ss.toast('⏳ Time limit hit. Auto-restarting in 30s...', '♻️ Queued', 10);
      
      ScriptApp.newTrigger('processNewTrackSelections')
               .timeBased()
               .after(1000 * 30) // Resume in 30s
               .create();
      
      isTimedOut = true;
      break; 
    }

    var studentRow = toProcess[i];
    var email = studentRow[1];
    var selectedTrackName = studentRow[2];

    try {
      var trackInfo = tracksData.find(function(row) {
        return row[1] === selectedTrackName && row[3] === 0;
      });

      if (trackInfo) {
        var studio = trackInfo[2];
        
        // Find all L0 tracks in this studio
        var studioL0Tracks = tracksData.filter(function(t) {
          return t[2] === studio && t[3] === 0 && (t[6] === 'TRUE' || t[6] === true);
        });
        
        // Unlock them
        var unlockedNames = [];
        studioL0Tracks.forEach(function(t) {
          // USE THE CACHED FUNCTION HERE
          var assignments = findAssignmentsForLevelCached_(activeClassrooms, courseWorkCache, t[0], 0);
          
          assignments.forEach(function(assign) {
            // Updated to use backoff inside
            var success = addStudentToAssignment_(assign.courseId, assign.assignmentId, email, assign.state);
            if (success && unlockedNames.indexOf(t[1]) === -1) unlockedNames.push(t[1]);
          });
        });

        // Log Success
        newLogEntries.push([
          new Date(),
          email,
          selectedTrackName,
          '✅',
          new Date()
        ]);

        // Clean up the Form Assignment


      } else {
        console.error('Track not found for ' + email + ': ' + selectedTrackName);
        // Log error so we don't retry forever? 
        // Optional: you might want to skip logging this so it retries if it was a typo fix later.
      }
      
    } catch (e) {
      console.error('Error processing ' + email, e);
    }
  }

  // 8. Bulk Write Logs
  if (newLogEntries.length > 0) {
    var startRow = logSheet.getLastRow() + 1;
    if (startRow < 3) startRow = 3;
    logSheet.getRange(startRow, 1, newLogEntries.length, 5).setValues(newLogEntries);
    console.log('Batch saved. Processed ' + newLogEntries.length + ' students.');
  }
  
if (!isTimedOut) {
     ss.toast('Success! All students processed.', '✅ Done', 5);
  }
} // <--- This closing brace was missing!
/**
 * Unassign the form assignment from a student after they submit
 */
function unassignFormFromStudent_(studentEmail) {
  var props = PropertiesService.getScriptProperties();
  var assignmentIdsJson = props.getProperty(CONFIG.PROPS.TRACK_FORM_ASSIGNMENT_IDS);
  
  if (!assignmentIdsJson) return;
  
  var assignmentIds = JSON.parse(assignmentIdsJson);
  
  Object.keys(assignmentIds).forEach(function(courseId) {
    var assignmentId = assignmentIds[courseId];
    
    try {
      var students = getAllStudents_(courseId);
      var student = students.find(function(s) {
        return s.profile.emailAddress === studentEmail;
      });
      
      if (!student) return;
      
      // Get current assignment to check assignee mode
      var assignment = Classroom.Courses.CourseWork.get(courseId, assignmentId);
      
      if (assignment.assigneeMode === 'ALL_STUDENTS') {
        // First student to submit: change to individual mode, remove them
        Classroom.Courses.CourseWork.modifyAssignees({
          assigneeMode: 'INDIVIDUAL_STUDENTS',
          modifyIndividualStudentsOptions: {
            removeStudentIds: [student.userId]
          }
        }, courseId, assignmentId);
        
        // Now we need to add back everyone EXCEPT this student
        var allStudentIds = students.map(function(s) { return s.userId; });
        var remainingIds = allStudentIds.filter(function(id) { return id !== student.userId; });
        
        if (remainingIds.length > 0) {
          Classroom.Courses.CourseWork.modifyAssignees({
            modifyIndividualStudentsOptions: {
              addStudentIds: remainingIds
            }
          }, courseId, assignmentId);
        }
      } else {
        // Already individual mode: just remove this student
        Classroom.Courses.CourseWork.modifyAssignees({
          modifyIndividualStudentsOptions: {
            removeStudentIds: [student.userId]
          }
        }, courseId, assignmentId);
      }
      
      console.log('Unassigned form from ' + studentEmail + ' in course ' + courseId);
      
    } catch (error) {
      console.error('Error unassigning form from ' + studentEmail + ' in ' + courseId + ':', error);
    }
  });
}

/**
 * Get all Level 0 tracks for a given studio
 */
function getLevel0TracksByStudio_(studio) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksLastRow = tracksSheet.getLastRow();
  
  if (tracksLastRow < 3) return [];
  
  // UPDATED: Reading 10 columns
  var tracksData = tracksSheet.getRange(3, 1, tracksLastRow - 2, 10).getValues();
  
  return tracksData.filter(function(row) {
    return row[2] === studio && row[3] === 0 && (row[6] === 'TRUE' || row[6] === true);
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 8: ASSIGNMENT CREATION
// ═══════════════════════════════════════════════════════════════════════════════

function parseUrlToMaterial_(url, shareMode) {
  url = url.trim();
  if (!url) return null;
  
  var youtubeMatch = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([\w-]+)/);
  if (youtubeMatch) {
    return {
      youtubeVideo: {
        id: youtubeMatch[1]
      }
    };
  }
  
  var driveFileMatch = url.match(/drive\.google\.com\/file\/d\/([\w-]+)/);
  if (driveFileMatch) {
    return {
      driveFile: {
        driveFile: { id: driveFileMatch[1] },
        shareMode: shareMode
      }
    };
  }
  
  var docsMatch = url.match(/docs\.google\.com\/(?:document|spreadsheets|presentation|forms)\/d\/([\w-]+)/);
  if (docsMatch) {
    return {
      driveFile: {
        driveFile: { id: docsMatch[1] },
        shareMode: shareMode
      }
    };
  }
  
  var driveIdMatch = url.match(/drive\.google\.com\/.*[?&]id=([\w-]+)/);
  if (driveIdMatch) {
    return {
      driveFile: {
        driveFile: { id: driveIdMatch[1] },
        shareMode: shareMode
      }
    };
  }
  
  var driveOpenMatch = url.match(/drive\.google\.com\/open\?id=([\w-]+)/);
  if (driveOpenMatch) {
    return {
      driveFile: {
        driveFile: { id: driveOpenMatch[1] },
        shareMode: shareMode
      }
    };
  }
  
  return {
    link: {
      url: url
    }
  };
}

function parseUrlsToMaterials_(urlString, shareMode) {
  if (!urlString || urlString.toString().trim() === '') return [];
  
  var urls = urlString.toString().split(',');
  var materials = [];
  
  urls.forEach(function(url) {
    var material = parseUrlToMaterial_(url.trim(), shareMode);
    if (material) {
      materials.push(material);
    }
  });
  
  return materials;
}

/**
 * Refresh track dropdowns (private helper)
 */
function refreshTrackDropdowns_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var createSheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  
  if (!createSheet || !tracksSheet) return;

  // FIXED: Changed A100 to A500 to catch all your Theory tracks
  var trackCodeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(tracksSheet.getRange('A3:A500'), true) 
    .setAllowInvalid(false)
    .build();
    
  createSheet.getRange('A3:A500').setDataValidation(trackCodeRule);
}

/**
 * UPDATED: Validates using the new 12-column structure
 * A=Track, B=Level, C=Week, D=Assign#, E=Title ... L=Status
 */
function validateAssignmentData_(data) {
  var issues = [];
  data.forEach(function(row, index) {
    // Row indices: 0=Track, 1=Level, 2=Week, 3=Assign#, 4=Title, 11=Status
    
    // Skip empty rows (must have Track AND Title)
    if (!row[0] && !row[4]) return; 
    
    // Only check rows marked "Ready"
    if (row[11] !== '⏳ Ready') return; 
    
    var rowNum = index + 3;
    var trackCode = row[0];
    var level = row[1];
    var week = row[2];
    var assignNum = row[3]; 
    var title = row[4];
    
    if (!trackCode) issues.push('Row ' + rowNum + ': Missing Track Code');
    if (trackCode && !/^[A-Z]{2,4}$/.test(trackCode)) issues.push('Row ' + rowNum + ': Track Code should be 2-4 uppercase letters');
    if (level === '' || level === null || level === undefined) issues.push('Row ' + rowNum + ': Missing Level');
    if (!week) issues.push('Row ' + rowNum + ': Missing Week (A, B, C, or D)');
    if (!assignNum) issues.push('Row ' + rowNum + ': Missing Assignment #');
    if (!title) issues.push('Row ' + rowNum + ': Missing Title');
  });
  return issues;
}

/**
 * Create OR Update selected assignments
 * UPDATED: Matches new column layout (Inputs A-J, Status K)
 */
/**
 * Create OR Update selected assignments
 * UPDATED: Uses ID-based tracking (JSON) for robustness against renaming.
 */
function createSelectedAssignments() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  refreshTrackDropdowns_();
  
  var createSheet = ss.getSheetByName(CONFIG.SHEETS.CREATE_ASSIGNMENT);
  var lastRow = createSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('ℹ️ Nothing to Process', 'No assignments found.', ui.ButtonSet.OK);
    return;
  }
  var data = createSheet.getRange(3, 1, lastRow - 2, 14).getValues();
  // --- ADD THIS BLOCK ---
  var issues = validateAssignmentData_(data);
  if (issues.length > 0) {
    ui.alert('⚠️ Validation Errors', 'Please fix the following issues before processing:\n\n' + issues.slice(0, 10).join('\n'), ui.ButtonSet.OK);
    return;
  }
  // ----------------------
  
  // Read columns A to N (14 cols)

  
  var toProcess = data.filter(function(row) {
    var trackCode = row[0];        // Col A
    var title = row[4];            // Col E
    var status = row[11] ? row[11].toString() : ''; // Col L
    var hasContent = (trackCode !== '' && trackCode != null) && (title !== '' && title != null);
    // Process if it's "Ready" OR "Update Needed" (custom statuses allowed)
    var isReady = status === '⏳ Ready' || status === '⏳ Update'; 
    return hasContent && isReady;
  });
  
  if (toProcess.length === 0) {
    ui.alert('ℹ️ Nothing to Process', 'No assignments have Status "⏳ Ready".', ui.ButtonSet.OK);
    return;
  }
  
  var confirm = ui.alert('📝 Process Assignments?', 'Ready to process ' + toProcess.length + ' assignment(s).\n\nExisting assignments will be UPDATED (renames are safe).\nNew assignments will be CREATED.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  
  // 1. Get Active Classrooms
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var classroomsLastRow = classroomsSheet.getLastRow();
  var classroomData = classroomsLastRow >= 3 ?
    classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues() : [];
  var activeClassrooms = classroomData.filter(function(row) { return row[4] === 'TRUE' || row[4] === true; });
  
  if (activeClassrooms.length === 0) {
    ui.alert('⚠️ No Active Classrooms', 'Set at least one classroom to Active = TRUE.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Scanning classrooms...', '🔍 Indexing', 60);
  
  // 2. PRE-FETCH & INDEX: Map ID->Work and Title->Work for every classroom
  // This allows us to find assignments by ID (robust) OR Title (fallback)
  var classroomIndices = {};
  
  activeClassrooms.forEach(function(classroom) {
    var courseId = classroom[0];
    classroomIndices[courseId] = {
      byId: {},
      byTitle: {}
    };
    
    try {
      var work = getAllCoursework_(courseId);
      work.forEach(function(w) { 
        classroomIndices[courseId].byId[w.id] = w;
        classroomIndices[courseId].byTitle[w.title] = w;
      });
    } catch (e) { console.log('Error indexing course ' + courseId); }
  });
  
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksData = tracksSheet.getLastRow() >= 3 ?
    tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 10).getValues() : [];
  var trackMap = buildTrackMap_(tracksData);
  
  var successCount = 0, updateCount = 0, errorCount = 0;
  
  // 3. PROCESS ROWS
  // Reverse loop so we can process bottom-up (optional preference)
  toProcess.reverse().forEach(function(row) {
    var rowNum = data.indexOf(row) + 3;
    
    var trackCode = row[0];
    var level = parseInt(row[1]); if (isNaN(level)) level = 0;
    var week = row[2];
    var assignNum = row[3];
    var title = row[4];
    var maxMP = row[5];
    
    var topicRaw = row[6];
    var description = row[7];
    var viewMaterialsRaw = row[8];
    var copyMaterialsRaw = row[9];
    
    // Parse Stored IDs (Column M/13)
    var storedIdsJson = row[12]; // Index 12 is Col M
    var storedIds = {};
    try {
      if (storedIdsJson && storedIdsJson.toString().trim() !== '') {
        // Try parsing JSON. If it fails (old format), we ignore it.
        if (storedIdsJson.toString().trim().startsWith('{')) {
            storedIds = JSON.parse(storedIdsJson);
        }
      }
    } catch (e) { console.log('Could not parse stored IDs for row ' + rowNum); }

    // Logic setup
    var trackKey = trackCode + '-L' + level;
    var trackConfig = trackMap[trackKey];
    var topicName = topicRaw || (trackConfig ? (trackConfig.studio + ' Studio: ' + trackConfig.name) : (trackCode + ' Level ' + level));

    var displayTitle = title;
    var prefixMatch = title.toString().match(/^[\d\w\s.:]+:\s*(.*)/);
    if (prefixMatch) displayTitle = prefixMatch[1];
    var fullTitle = level + week + assignNum + ': ' + displayTitle + ' [' + trackCode + ']';
    
    var viewMaterials = parseUrlsToMaterials_(viewMaterialsRaw, 'VIEW');
    var copyMaterials = parseUrlsToMaterials_(copyMaterialsRaw, 'STUDENT_COPY');
    var allMaterials = viewMaterials.concat(copyMaterials);
    
    if (trackCode) {
      var generatedId = trackCode.toLowerCase() + '-L' + level + '-' + week.toLowerCase() + assignNum;
      var fullUrl = 'https://www.multimediaheroes.com/?view=' + generatedId;
      allMaterials.push({ link: { url: fullUrl, title: '🚀 View on Skill Tree' } });
    }
    
    var finalIdsMap = {}; // Will store { courseId: assignId } for saving
    var hasError = false;
    
    activeClassrooms.forEach(function(classroom) {
      var courseId = classroom[0];
      var index = classroomIndices[courseId];
      
      try {
        // A. Topic Handling (Using Sticky Topics)
        var topicId = getStickyTopicId_(courseId, topicName);
        
        // B. Find Existing Assignment
        // Strategy: 1. Check Stored ID (Best) -> 2. Check Title Match (Fallback)
        var existingWork = null;
        var storedId = storedIds[courseId]; // Get ID for this specific course
        
        if (storedId && index.byId[storedId]) {
            existingWork = index.byId[storedId]; // Found by ID! Rename safe.
        } else if (index.byTitle[fullTitle]) {
            existingWork = index.byTitle[fullTitle]; // Found by Title (Migration/Fallback)
        }
        
        if (existingWork) {
          // UPDATE
          updateAssignment_(courseId, existingWork.id, {
            title: fullTitle,
            description: description || '',
            maxPoints: maxMP || 0,
            topicId: topicId,
            materials: allMaterials 
          });
          finalIdsMap[courseId] = existingWork.id;
          updateCount++;
        } else {
          // CREATE
          var assignmentBody = {
            title: fullTitle,
            description: description || '',
            workType: 'ASSIGNMENT',
            maxPoints: maxMP || 0,
            state: 'DRAFT',
            topicId: topicId,
            assigneeMode: 'ALL_STUDENTS'
          };
          if (allMaterials.length > 0) assignmentBody.materials = allMaterials;
          
          var created = Classroom.Courses.CourseWork.create(assignmentBody, courseId);
          finalIdsMap[courseId] = created.id;
          successCount++;
          
          // Add to index immediately in case we double-process (unlikely but safe)
          index.byId[created.id] = created; 
        }
      } catch (error) {
        console.error('Error in ' + courseId + ': ' + error.message);
        hasError = true;
      }
    });
    
    // Save Results
    if (hasError) {
      createSheet.getRange(rowNum, 12).setValue('❌ Error');
    } else {
      var statusMsg = (updateCount > 0 && successCount === 0) ? '✅ Updated' : '✅ Created';
      createSheet.getRange(rowNum, 12).setValue(statusMsg);
      // NEW: Save as JSON Map
      createSheet.getRange(rowNum, 13).setValue(JSON.stringify(finalIdsMap)); 
      createSheet.getRange(rowNum, 14).setValue(new Date());
    }
  });
  
  ss.toast('Done! ' + successCount + ' new, ' + updateCount + ' updated.', '✅ Complete', 5);
}

/**
 * Helper function to patch an existing assignment
 */
/**
 * Helper function to patch an existing assignment
 * UPDATED: Removed maxPoints to prevent "Non-supported update mask" errors
 */
/**
 * Helper function to patch an existing assignment
 * UPDATED: Removed maxPoints AND materials (neither can be updated via API)
 */
function updateAssignment_(courseId, assignmentId, updates) {
  // 1. Only include fields that Google allows us to update
  var mask = 'title,description,topicId'; 
  
  var resource = {
    title: updates.title,
    description: updates.description,
    topicId: updates.topicId
  };

  // NOTE: 'materials' is intentionally omitted. 
  // The Google Classroom API does NOT allow updating attachments on existing assignments.
  
  Classroom.Courses.CourseWork.patch(resource, courseId, assignmentId, {
    updateMask: mask
  });
}

function getOrCreateTopic_(courseId, topicName) {
  var topics = [];
  var pageToken = null;
  
  do {
    var response = Classroom.Courses.Topics.list(courseId, {
      pageSize: 100,
      pageToken: pageToken
    });
    
    if (response.topic) {
      topics = topics.concat(response.topic);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  var existing = topics.find(function(t) {
    return t.name === topicName;
  });
  if (existing) return existing.topicId;
  
  var newTopic = Classroom.Courses.Topics.create({
    name: topicName
  }, courseId);
  
  return newTopic.topicId;
}

function deleteAssignmentDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== CONFIG.SHEETS.ASSIGNMENTS) {
    ui.alert('📋 Go to Assignments Tab',
      'To delete assignments:\n\n' +
      '1. Go to the "📋 Assignments" tab\n' +
      '2. Select the row(s) of assignments you want to delete\n' +
      '3. Run this menu option again',
      ui.ButtonSet.OK);
    return;
  }
  
  var selection = sheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  
  if (startRow < 3) {
    ui.alert('⚠️ Select Assignment Rows',
      'Select rows 3 or below containing assignments.',
      ui.ButtonSet.OK);
    return;
  }
  
  // 1. Batch read selection
  var dataRange = sheet.getRange(startRow, 1, numRows, 9).getValues();
  var assignmentsToDelete = [];
  
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row < 3) continue;
    
    var rowData = dataRange[i];
    // Mapping based on your Sheet columns:
    // Col C (Index 2) = Classroom ID
    // Col D (Index 3) = Assignment ID
    // Col E (Index 4) = Title
    var courseId = rowData[2]; 
    var assignmentId = rowData[3];
    var title = rowData[4];
    
    if (courseId && assignmentId) {
      assignmentsToDelete.push({
        row: row,
        courseId: courseId,
        assignmentId: assignmentId,
        title: title
      });
    }
  }
  
  if (assignmentsToDelete.length === 0) {
    ui.alert('⚠️ No Valid Assignments Selected', 'Make sure you selected valid assignment rows.', ui.ButtonSet.OK);
    return;
  }
  
  var confirm = ui.alert(
    '🗑️ Delete ' + assignmentsToDelete.length + ' Assignment(s)?',
    'Are you sure? This will delete the assignments from Google Classroom and remove the rows.\n\n' + 
    '⚠️ Empty Topics will also be deleted to reduce clutter.',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  ss.toast('Scanning classrooms & topics...', '🔍 Preparing', 60);

  // 2. Pre-load Classroom Data & Topic Map
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var classroomsLastRow = classroomsSheet.getLastRow();
  var classroomData = classroomsLastRow >= 3 ? classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues() : [];
  var activeClassrooms = classroomData.filter(function(row) {
    return row[4] === 'TRUE' || row[4] === true;
  });

  // Map to store assignment IDs and their Topic IDs so we know what to check later
  // Structure: { courseId: { 'Title': { id: '...', topicId: '...' } } }
  var classroomCourseWorkMap = {}; 
  
  // Set to track which topics might need deleting: { courseId: [topicId1, topicId2] }
  var topicsToCheck = {}; 

  activeClassrooms.forEach(function(classroom) {
    var cId = classroom[0];
    try {
      var cw = getAllCoursework_(cId);
      var titleMap = {};
      
      cw.forEach(function(w) { 
        titleMap[w.title] = { 
          id: w.id, 
          topicId: w.topicId // Capture the Topic ID!
        }; 
      });
      classroomCourseWorkMap[cId] = titleMap;
    } catch (e) {
      console.log('Error mapping course ' + cId);
    }
  });
  
  ss.toast('Deleting assignments...', '🗑️ Deleting', 30);
  
  var totalDeleted = 0;
  var errors = [];
  
  // Sort reverse to preserve row indices during deletion
  assignmentsToDelete.sort(function(a, b) { return b.row - a.row; });
  
  // 3. Deletion Loop
  for (var idx = 0; idx < assignmentsToDelete.length; idx++) {
    var assignment = assignmentsToDelete[idx];
    try {
      // --- A. Delete from Main Classroom ---
      // We look up the topicId first so we can check it later
      var mainMap = classroomCourseWorkMap[assignment.courseId];
      if (mainMap && mainMap[assignment.title] && mainMap[assignment.title].topicId) {
        var tId = mainMap[assignment.title].topicId;
        if (!topicsToCheck[assignment.courseId]) topicsToCheck[assignment.courseId] = [];
        if (topicsToCheck[assignment.courseId].indexOf(tId) === -1) {
          topicsToCheck[assignment.courseId].push(tId);
        }
      }

      try {
        Classroom.Courses.CourseWork.remove(assignment.courseId, assignment.assignmentId);
        totalDeleted++;
      } catch (e) {
        if (e.message.indexOf('404') === -1) console.log('Error deleting ' + assignment.title);
      }
      
      // --- B. Delete from Other Active Classrooms ---
      activeClassrooms.forEach(function(classroom) {
        var otherCourseId = classroom[0];
        if (otherCourseId === assignment.courseId) return;
        
        var map = classroomCourseWorkMap[otherCourseId];
        if (map && map[assignment.title]) {
           var otherData = map[assignment.title];
           
           // Mark this topic for cleanup check too
           if (otherData.topicId) {
             if (!topicsToCheck[otherCourseId]) topicsToCheck[otherCourseId] = [];
             if (topicsToCheck[otherCourseId].indexOf(otherData.topicId) === -1) {
               topicsToCheck[otherCourseId].push(otherData.topicId);
             }
           }

           try {
             Classroom.Courses.CourseWork.remove(otherCourseId, otherData.id);
             totalDeleted++;
           } catch (e) { console.log('Could not delete copy'); }
        }
      });
      
      // --- C. Delete Spreadsheet Row ---
      sheet.deleteRow(assignment.row);
      SpreadsheetApp.flush(); 
      
    } catch (error) {
      errors.push(assignment.title + ': ' + error.message);
    }
  }

  // 4. Topic Cleanup Phase (The "Clutter Killer")
  ss.toast('Cleaning up empty topics...', '🧹 Tidying', 10);
  var topicsDeleted = 0;

  var affectedCourses = Object.keys(topicsToCheck);
  
  affectedCourses.forEach(function(cId) {
    try {
      // Get FRESH coursework list to see what's left
      var remainingWork = getAllCoursework_(cId);
      
      // Get a list of all Topic IDs currently in use
      var usedTopicIds = remainingWork.map(function(w) { return w.topicId; });
      
      // Check our candidates
      var candidates = topicsToCheck[cId];
      candidates.forEach(function(topicId) {
        // If the candidate is NOT in the used list, it is empty!
        if (usedTopicIds.indexOf(topicId) === -1) {
          try {
            Classroom.Courses.Topics.remove(cId, topicId);
            topicsDeleted++;
            console.log('Deleted empty topic ' + topicId + ' in ' + cId);
          } catch (e) {
            console.log('Could not delete topic: ' + e.message);
          }
        }
      });
    } catch (e) {
      console.log('Error cleaning topics for course ' + cId);
    }
  });
  
  if (errors.length === 0) {
    var msg = 'Deleted ' + totalDeleted + ' assignments.';
    if (topicsDeleted > 0) msg += '\n🧹 Removed ' + topicsDeleted + ' empty topics.';
    ui.alert('✅ Cleanup Complete', msg, ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Partial Success', 'Errors:\n' + errors.join('\n'), ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 9: DASHBOARD & LEADERBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Triggers on cell edits (dropdowns)
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Dashboard dropdown (F12)
  if (sheet.getName() === CONFIG.SHEETS.DASHBOARD && 
      range.getRow() === 12 && 
      range.getColumn() === 6) {
    refreshDashboard();
  }
  
  // Leaderboard dropdown (C3)
  if (sheet.getName() === CONFIG.SHEETS.LEADERBOARD && 
      range.getRow() === 3 && 
      range.getColumn() === 3) {
    refreshLeaderboard();
  }

  // Grading Export Checkbox (Now Row 10)
  if (sheet.getName() === CONFIG.SHEETS.GRADING && 
      range.getRow() === 10 && // Updated from 9 to 10
      range.getColumn() === 1 &&
      range.isChecked()) {
    generatePowerSchoolCSV();
    range.uncheck();
  }
  
  // Auction Logic (Checkbox Triggers)
  if (sheet.getName() === CONFIG.SHEETS.AUCTIONS) {
    var col = range.getColumn();
    // Col 4 = Check Bid, Col 6 = Confirm
    if ((col === 4 || col === 6) && range.isChecked()) {
      if (col === 4) checkAuctionBids();
      if (col === 6) confirmAuctionTransactions();
    }
  }
}

/**
 * Refresh dashboard with current data
 */
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!dashboard) return;
  
  try {
    // Get progress data (now 12 columns)
    var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
    var progressLastRow = progressSheet.getLastRow();
    
    var progressData = progressLastRow >= 3 
      ? progressSheet.getRange(3, 1, progressLastRow - 2, 12).getValues().filter(function(row) { return row[0]; })
      : [];
    
    // Build student totals from progress data
    var studentTotals = {};
    progressData.forEach(function(row) {
      var email = row[0];
      var name = row[1];
      var period = row[2] ? row[2].toString() : '?';
      var trackMP = row[6] || 0;
      var weekMP = row[7] || 0;
      
      if (!studentTotals[email]) {
        studentTotals[email] = {
          email: email,
          name: name,
          period: period,
          totalMP: 0,
          weekMP: 0
        };
      }
      studentTotals[email].totalMP += trackMP;
      studentTotals[email].weekMP += weekMP;
    });
    
    // Get full roster from classrooms (to include 0 MP students)
    var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    var classroomsLastRow = classroomsSheet.getLastRow();
    if (classroomsLastRow >= 3) {
      var classroomData = classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues();
      var activeClassrooms = classroomData.filter(function(row) {
        return row[4] === 'TRUE' || row[4] === true;
      });
      
      activeClassrooms.forEach(function(classroom) {
        var courseId = classroom[0];
        var period = classroom[3] ? classroom[3].toString() : '?';
        
        try {
// SAFE REPLACEMENT: Use spreadsheet data instead of API calls
          // This prevents "Permission" errors on the Dashboard
          if (!studentTotals[courseId]) {
             // We just skip adding empty students here. 
             // If they have 0 MP and aren't in the Progress tab yet, 
             // they will appear after the next full Sync.
          }
        } catch (error) {
          console.log('Could not get students from ' + courseId + ': ' + error.message);
        }
      });
    }
    
    var studentList = Object.values(studentTotals);
    var totalStudents = studentList.length;
    
    var targetMP = getWeeklyTargetMP_();
    var maxMP = getWeeklyMaxMP_();
    
    var onTrack = studentList.filter(function(s) {
      return s.weekMP >= targetMP;
    }).length;
    
    var needsSupport = studentList.filter(function(s) {
      return s.weekMP < 60;
    }).length;
    
    var readyToUnlock = progressData.filter(function(row) {
      return row[10] === '✅ YES'; // Column K (index 10)
    }).length;
    
    dashboard.getRange('C6').setValue(totalStudents);
    dashboard.getRange('C7').setValue(onTrack);
    dashboard.getRange('C8').setValue(needsSupport);
    dashboard.getRange('C9').setValue(readyToUnlock);
    
    // Recent unlocks
    var logSheet = ss.getSheetByName(CONFIG.SHEETS.UNLOCK_LOG);
    var logLastRow = logSheet.getLastRow();
    var logData = logLastRow >= 3 
      ? logSheet.getRange(3, 1, logLastRow - 2, 6).getValues()
          .filter(function(row) { return row[0]; })
          .slice(-5)
          .reverse()
      : [];
    
    dashboard.getRange('E6:F10').clear();
    if (logData.length > 0) {
      var unlockDisplay = logData.map(function(row) {
        return [row[2] || row[1], row[3]]; // Name or email, what was unlocked
      });
      dashboard.getRange(6, 5, Math.min(5, unlockDisplay.length), 2).setValues(unlockDisplay.slice(0, 5));
    } else {
      dashboard.getRange('E6').setValue('No unlocks yet — run a sync!')
        .setFontColor(CONFIG.COLORS.TEXT_MUTED)
        .setFontStyle('italic');
    }
    
    // Close to unlocking
    var closeToUnlock = progressData.filter(function(row) {
      var trackMP = row[6] || 0;
      var threshold = CONFIG.DEFAULTS.UNLOCK_THRESHOLD;
      return trackMP >= threshold - 20 && trackMP < threshold;
    }).slice(0, 5);
    
    dashboard.getRange('B13:C22').clear();
    if (closeToUnlock.length > 0) {
      var closeDisplay = closeToUnlock.map(function(row) {
        return [row[1], row[6] + '/' + CONFIG.DEFAULTS.UNLOCK_THRESHOLD + ' MP in ' + row[3]];
      });
      dashboard.getRange(13, 2, closeDisplay.length, 2).setValues(closeDisplay)
        .setFontColor(CONFIG.COLORS.WARNING);
    } else {
      dashboard.getRange('B13').setValue('No students close to unlocking')
        .setFontColor(CONFIG.COLORS.TEXT_MUTED)
        .setFontStyle('italic');
    }
    
    // Bottom 10 by Total MP (with period filter dropdown)
    var filterValue = dashboard.getRange('F12').getValue();
    
    dashboard.getRange('E13:F22').clear();
    
    if (filterValue === '▶ Click to show' || !filterValue) {
      dashboard.getRange('E13').setValue('Select dropdown to reveal')
        .setFontColor(CONFIG.COLORS.TEXT_MUTED)
        .setFontStyle('italic');
    } else {
      var filteredStudents = studentList;
      
      if (filterValue !== '📊 All Periods') {
        var periodMatch = filterValue.match(/Period (\d+)/);
        if (periodMatch) {
          var selectedPeriod = periodMatch[1];
          filteredStudents = studentList.filter(function(s) {
            return s.period.toString() === selectedPeriod;
          });
        }
      }
      
      filteredStudents.sort(function(a, b) {
        return a.totalMP - b.totalMP;
      });
      
      var bottom10 = filteredStudents.slice(0, 10);
      
      if (bottom10.length > 0) {
        var attentionDisplay = bottom10.map(function(s) {
          return [s.name, s.totalMP + ' total MP'];
        });
        dashboard.getRange(13, 5, attentionDisplay.length, 2).setValues(attentionDisplay)
          .setFontColor(CONFIG.COLORS.DANGER);
      } else {
        dashboard.getRange('E13').setValue('No students in this period')
          .setFontColor(CONFIG.COLORS.TEXT_MUTED)
          .setFontStyle('italic');
      }
    }
    
  } catch (error) {
    console.error('Dashboard refresh error:', error);
  }
}

/**
 * Refresh leaderboard with This Week and All Time columns
 */
function refreshLeaderboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LEADERBOARD);
  if (!sheet) return;
  
  try {
    // Get progress data (now 12 columns)
    var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
    var progressLastRow = progressSheet.getLastRow();
    var progressData = progressLastRow >= 3 
      ? progressSheet.getRange(3, 1, progressLastRow - 2, 12).getValues().filter(function(row) { return row[0]; })
      : [];
    
    // Build student totals
    var studentTotals = {};
    progressData.forEach(function(row) {
      var email = row[0];
      var name = row[1];
      var period = row[2] ? row[2].toString() : '?';
      var trackMP = row[6] || 0;
      var weekMP = row[7] || 0;
      
      if (!studentTotals[email]) {
        studentTotals[email] = {
          email: email,
          name: name,
          period: period,
          totalMP: 0,
          weekMP: 0
        };
      }
      studentTotals[email].totalMP += trackMP;
      studentTotals[email].weekMP += weekMP;
    });
    
    var studentList = Object.values(studentTotals);
    
    // Get filter value
    var filterValue = sheet.getRange('C3').getValue();
    
    // Filter by period if needed
    var filteredStudents = studentList;
    
    if (filterValue !== '📊 All Periods') {
      var periodMatch = filterValue.match(/Period (\d+)/);
      if (periodMatch) {
        var selectedPeriod = periodMatch[1];
        filteredStudents = studentList.filter(function(s) {
          return s.period.toString() === selectedPeriod;
        });
      }
    }
    
    // Clear data areas
    sheet.getRange('B7:D16').clear();
    sheet.getRange('F7:H16').clear();
    
    // THIS WEEK - Top 10 by weekMP
    var thisWeekSorted = filteredStudents.filter(function(s) {
      return s.weekMP > 0;
    }).sort(function(a, b) {
      return b.weekMP - a.weekMP;
    }).slice(0, 10);
    
    if (thisWeekSorted.length > 0) {
      var thisWeekData = thisWeekSorted.map(function(s, index) {
        var rank = index + 1;
        var rankDisplay = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : rank.toString();
        return [rankDisplay, s.name, s.weekMP + ' MP'];
      });
      
      sheet.getRange(7, 2, thisWeekData.length, 3).setValues(thisWeekData);
      
      // Style rows
      for (var i = 0; i < thisWeekData.length; i++) {
        var rowNum = 7 + i;
        var bgColor = (i % 2 === 0) ? CONFIG.COLORS.ROW_NORMAL : CONFIG.COLORS.ROW_ALT;
        if (i === 0) bgColor = '#FEF3C7';
        else if (i === 1) bgColor = '#F3F4F6';
        else if (i === 2) bgColor = '#FED7AA';
        sheet.getRange(rowNum, 2, 1, 3).setBackground(bgColor);
      }
      sheet.getRange(7, 2, thisWeekData.length, 1).setHorizontalAlignment('center');
      sheet.getRange(7, 4, thisWeekData.length, 1).setHorizontalAlignment('center');
    } else {
      sheet.getRange('B7').setValue('No MP earned this week')
        .setFontColor(CONFIG.COLORS.TEXT_MUTED)
        .setFontStyle('italic');
    }
    
    // ALL TIME - Top 10 by totalMP
    var allTimeSorted = filteredStudents.filter(function(s) {
      return s.totalMP > 0;
    }).sort(function(a, b) {
      return b.totalMP - a.totalMP;
    }).slice(0, 10);
    
    if (allTimeSorted.length > 0) {
      var allTimeData = allTimeSorted.map(function(s, index) {
        var rank = index + 1;
        var rankDisplay = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : rank.toString();
        return [rankDisplay, s.name, s.totalMP + ' MP'];
      });
      
      sheet.getRange(7, 6, allTimeData.length, 3).setValues(allTimeData);
      
      // Style rows
      for (var j = 0; j < allTimeData.length; j++) {
        var rowNum2 = 7 + j;
        var bgColor2 = (j % 2 === 0) ? CONFIG.COLORS.ROW_NORMAL : CONFIG.COLORS.ROW_ALT;
        if (j === 0) bgColor2 = '#FEF3C7';
        else if (j === 1) bgColor2 = '#F3F4F6';
        else if (j === 2) bgColor2 = '#FED7AA';
        sheet.getRange(rowNum2, 6, 1, 3).setBackground(bgColor2);
      }
      sheet.getRange(7, 6, allTimeData.length, 1).setHorizontalAlignment('center');
      sheet.getRange(7, 8, allTimeData.length, 1).setHorizontalAlignment('center');
    } else {
      sheet.getRange('F7').setValue('No MP earned yet')
        .setFontColor(CONFIG.COLORS.TEXT_MUTED)
        .setFontStyle('italic');
    }
    
  } catch (error) {
    console.error('Leaderboard refresh error:', error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 10: MANUAL UNLOCKS
// ═══════════════════════════════════════════════════════════════════════════════

function processManualUnlocks() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var unlockSheet = ss.getSheetByName(CONFIG.SHEETS.MANUAL_UNLOCKS);
  var lastRow = unlockSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('ℹ️ No Pending Unlocks', 'Add entries to the Manual Unlocks tab.', ui.ButtonSet.OK);
    return;
  }
  
  var data = unlockSheet.getRange(3, 1, lastRow - 2, 6).getValues();
  
  var pending = data.filter(function(row) {
    return row[4] === '⏳ Pending';
  });
  
  if (pending.length === 0) {
    ui.alert('ℹ️ No Pending Unlocks', 'Set Status to "⏳ Pending" first.', ui.ButtonSet.OK);
    return;
  }
  
  var confirm = ui.alert(
    '🔓 Process Manual Unlocks?',
    'Found ' + pending.length + ' pending unlock(s). Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var classroomsLastRow = classroomsSheet.getLastRow();
  var classroomData = classroomsLastRow >= 3 ? classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues() : [];
  var activeClassrooms = classroomData.filter(function(row) {
    return row[4] === 'TRUE' || row[4] === true;
  });
  
  var successCount = 0;
  var logEntries = [];
  
  pending.forEach(function(row) {
    var dataIndex = data.indexOf(row);
    var rowNum = dataIndex + 3;
    
    var email = row[0];
    var trackCode = row[1];
    var level = row[2];
    var reason = row[3];
    
    try {
      var assignments = findAssignmentsForLevel_(activeClassrooms, trackCode, level);
      
      var anySuccess = false;
      assignments.forEach(function(assignment) {
        var result = addStudentToAssignment_(assignment.courseId, assignment.assignmentId, email, assignment.state);
        if (result) anySuccess = true;
      });
      
      if (anySuccess) {
        unlockSheet.getRange(rowNum, 5).setValue('✅ Processed');
        unlockSheet.getRange(rowNum, 6).setValue(new Date());
        successCount++;
        
        logEntries.push([
          new Date(),
          email,
          '',
          trackCode + '-L' + level,
          'Manual',
          reason || 'Manual unlock'
        ]);
      } else {
        unlockSheet.getRange(rowNum, 5).setValue('❌ Error');
      }
      
    } catch (error) {
      unlockSheet.getRange(rowNum, 5).setValue('❌ Error');
      console.error('Error processing ' + email + ':', error);
    }
  });
  
  if (logEntries.length > 0) {
    var logSheet = ss.getSheetByName(CONFIG.SHEETS.UNLOCK_LOG);
    var logLastRow = Math.max(2, logSheet.getLastRow());
    logSheet.getRange(logLastRow + 1, 1, logEntries.length, 6).setValues(logEntries);
  }
  
  ui.alert('🔓 Done', 'Processed: ' + successCount + ' / ' + pending.length, ui.ButtonSet.OK);
}

/**
 * 🕵️ HIDDEN DB HANDLER
 * Reads the hidden sync log or creates it if missing.
 * Returns: Map { "assignmentId": timestamp }
 */
function getSyncDatabase_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "hidden_sync_db"; // The hidden sheet
  var sheet = ss.getSheetByName(sheetName);
  
  // If missing, create it and hide it
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
    sheet.getRange("A1:C1").setValues([["Assignment ID", "Course ID", "Last Synced"]]);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  
  // Read all data: Col A (ID), Col C (Timestamp)
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var db = {};
  
  data.forEach(function(row) {
    if (row[0]) db[row[0]] = new Date(row[2]).getTime();
  });
  
  return db;
}

/**
 * 💾 SAVE TO DB
 * Updates the hidden sheet with new timestamps.
 */
function updateSyncDatabase_(updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("hidden_sync_db");
  var lastRow = sheet.getLastRow();
  
  // 1. Read existing to map rows
  var existingMap = {}; // { "assignmentId": rowNumber }
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(function(r, i) {
      if(r[0]) existingMap[r[0]] = i + 2; // Row index is +2 (header + 0-index)
    });
  }
  
  // 2. Process updates
  var newRows = [];
  updates.forEach(function(u) {
    if (existingMap[u.id]) {
      // Update existing row
      sheet.getRange(existingMap[u.id], 3).setValue(new Date());
    } else {
      // Add new row
      newRows.push([u.id, u.courseId, new Date()]);
    }
  });
  
  // 3. Batch write new rows
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }
}
/**
 * Saves specific assignment grades to the hidden database.
 * Uses a Map to update existing rows or add new ones.
 */
function saveToGradeDatabase_(newEntries) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "hidden_grades_db";
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
    // Headers: Email, AssignID, Name, Period, Track, Level, Week, MP, Updated
    sheet.appendRow(["Email", "AssignID", "Name", "Period", "Track", "Level", "Week", "MP", "Updated"]);
  }
  
  var lastRow = sheet.getLastRow();
  var currentData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 9).getValues() : [];
  
  // Map Key: Email + AssignmentID
  var dbMap = new Map();
  
  // 1. Load existing DB into Map
  currentData.forEach((row, index) => {
    var key = row[0] + "_" + row[1];
    dbMap.set(key, { row: row, index: index });
  });
  
  // 2. Update Map with New Entries
  newEntries.forEach(entry => {
    var key = entry[0] + "_" + entry[1];
    // entry format matches DB columns
    dbMap.set(key, { row: entry, index: -1 }); // Index -1 means "new"
  });
  
  // 3. Convert Map back to Array
  var output = Array.from(dbMap.values()).map(item => item.row);
  
  // 4. Overwrite Sheet (This is safer/easier than splicing rows in GAS)
  if (output.length > 0) {
    sheet.getRange(2, 1, sheet.getLastRow(), 9).clearContent(); // Clear old
    sheet.getRange(2, 1, output.length, 9).setValues(output);   // Write new
  }
}

/**
 * Reads the ENTIRE grade database to calculate totals.
 * This ensures "Partial Syncs" don't destroy student levels.
 */
function compileProgressFromDB_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName("hidden_grades_db");
  if (!dbSheet || dbSheet.getLastRow() < 2) return;
  
  var allGrades = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 9).getValues();
  
  // Use your existing mapping helper
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksData = tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 10).getValues();
  var trackMap = buildTrackMap_(tracksData);

  // Convert DB format to the format aggregateProgress_ expects
  var progressObjects = allGrades.map(row => {
    return {
      email: row[0],
      name: row[2],
      period: row[3],
      trackCode: row[4],
      level: row[5],
      week: row[6],
      mp: row[7],
      gradeTime: row[8]
    };
  });

  // Now run your existing aggregation logic on the FULL dataset
  var aggregated = aggregateProgress_(progressObjects, trackMap);
  writeProgressSheet_(aggregated);
}
/**
 * NEW: Populates the '📋 Assignments' tab with all formatted coursework found in active classrooms
 */
/**
 * UPDATED: Populates '📋 Assignments' with Period and Tags
 */
function updateAssignmentsReferenceSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ASSIGNMENTS);
  if (!sheet) return;

  // AUTO-FIX: If header is old, re-format immediately
  var currentHeader = sheet.getRange('A2').getValue();
  if (currentHeader !== 'Period') {
    formatAssignmentsSheet_();
  }

  // 1. Get active classrooms and map Course IDs to Periods
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var classroomsLastRow = classroomsSheet.getLastRow();
  var classroomData = classroomsLastRow >= 3 ? classroomsSheet.getRange(3, 1, classroomsLastRow - 2, 6).getValues() : [];
  
  var activeClassrooms = [];
  var periodMap = {}; // Map: courseId -> period
  
  classroomData.forEach(function(row) {
    if (row[4] === 'TRUE' || row[4] === true) {
      activeClassrooms.push(row);
      periodMap[row[0]] = row[3]; // Store period
    }
  });

  if (activeClassrooms.length === 0) return;

  // 2. Get Track Tags using the existing buildTrackMap_ helper
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_TRACKS);
  var tracksLastRow = tracksSheet.getLastRow();
  // Ensure we read 10 columns to get the Tags column
  var tracksData = tracksLastRow >= 3 ? tracksSheet.getRange(3, 1, tracksLastRow - 2, 10).getValues() : [];
  var trackMap = buildTrackMap_(tracksData);

  var rows = [];

  // 3. Scan all coursework
  activeClassrooms.forEach(function(classroom) {
    var courseId = classroom[0];
    try {
      var coursework = getAllCoursework_(courseId);
      
      coursework.forEach(function(work) {
        var parsed = parseAssignmentName_(work.title);
        if (parsed) {
           // Look up Tags for this specific track level
           var trackKey = parsed.trackCode + '-L' + parsed.level;
           var trackInfo = trackMap[trackKey];
           var tags = '';
           if (trackInfo && trackInfo.trackTags && trackInfo.trackTags.length > 0) {
             tags = trackInfo.trackTags.join(', ');
           }

           rows.push([
             periodMap[courseId] || '', // Period (Col A)
             tags,                      // Tags (Col B)
             courseId,
             work.id,
             work.title,
             parsed.trackCode,
             parsed.level,
             parsed.week,
             work.maxPoints || 0,
             work.state,
             new Date(work.updateTime || new Date())
           ]);
        }
      });
    } catch (e) {
      console.error('Error fetching assignments for ' + courseId, e);
    }
  });

  // 4. Sort by Period (Numeric) -> Track -> Level -> Week
  rows.sort(function(a, b) {
     // Robust sort: extract numbers from period (e.g. "Period 1" -> 1)
     var getNum = function(s) {
       var match = s.toString().match(/\d+/);
       return match ? parseInt(match[0]) : 9999;
     };
     var pA = getNum(a[0]);
     var pB = getNum(b[0]);
     
     if (pA !== pB) return pA - pB;
     if (a[5] !== b[5]) return a[5].localeCompare(b[5]); // Track Code
     if (a[6] !== b[6]) return a[6] - b[6]; // Level
     return a[7].localeCompare(b[7]); // Week
  });

  // 5. Clear old data and write new data (11 columns)
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, 11).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, 11).setValues(rows);
    
    // Alternating row colors
    for (var i = 0; i < rows.length; i++) {
      var rowNum = 3 + i;
      var bgColor = (i % 2 === 0) ? CONFIG.COLORS.ROW_NORMAL : CONFIG.COLORS.ROW_ALT;
      sheet.getRange(rowNum, 1, 1, 11).setBackground(bgColor);
    }
  }
}
/**
 * HELPER: Finds the sheet where the raw form data is landing.
 * It looks for the special name we gave it ("Hidden_Raw_Responses...") 
 * or falls back to the default "Form Responses".
 */
function findRawResponseSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  // 1. Priority: Look for the specific hidden sheet we created
  var hiddenSheet = sheets.find(function(s) {
    return s.getName().indexOf("Hidden_Raw_Responses") !== -1;
  });
  if (hiddenSheet) return hiddenSheet;
  
  // 2. Fallback: Look for a standard Google Forms sheet
  var standardSheet = sheets.find(function(s) {
    return s.getName().indexOf("Form Responses") !== -1;
  });
  
  return standardSheet || null;
}
/**
 * ⚡ EXPONENTIAL BACKOFF WRAPPER
 * Tries to run a function. If it hits a rate limit (429 or 503),
 * it waits (1s, 2s, 4s, 8s...) and retries.
 */
function fetchWithBackoff_(func) {
  var maxRetries = 5;
  
  for (var n = 0; n <= maxRetries; n++) {
    try {
      return func();
    } catch (e) {
      // If we are out of retries, throw the error for real
      if (n === maxRetries) throw e;
      
      // Check if it's a rate limit error
      var msg = e.message.toLowerCase();
      if (msg.indexOf('too many requests') !== -1 || 
          msg.indexOf('limit exceeded') !== -1 || 
          msg.indexOf('quota') !== -1 || 
          msg.indexOf('503') !== -1 || 
          msg.indexOf('429') !== -1) {
            
        // 🛑 WAIT: Exponential backoff (1s, 2s, 4s, 8s, 16s) + random jitter
        var waitTime = (Math.pow(2, n) * 1000) + (Math.round(Math.random() * 1000));
        console.warn('⚠️ API Limit hit. Retrying in ' + waitTime + 'ms...');
        Utilities.sleep(waitTime);
        
      } else {
        // If it's a different error (like "Not Found"), throw it immediately
        throw e;
      }
    }
  }
}
/**
 * GET OR CREATE TOPIC (STICKY VERSION)
 * Uses a hidden sheet to remember Topic IDs.
 * This allows you to rename topics in Classroom without creating duplicates.
 */
function getStickyTopicId_(courseId, topicName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheetName = "hidden_topic_cache";
  var dbSheet = ss.getSheetByName(dbSheetName);
  
  // 1. Initialize Cache Sheet if needed
  if (!dbSheet) {
    dbSheet = ss.insertSheet(dbSheetName);
    dbSheet.hideSheet();
    dbSheet.appendRow(["Course ID", "Topic Name", "Topic ID"]);
  }
  
  // 2. Check Cache
  var lastRow = dbSheet.getLastRow();
  var cacheData = lastRow > 1 ? dbSheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  
  var cachedRow = cacheData.find(function(row) {
    return row[0] == courseId && row[1] == topicName;
  });
  
  if (cachedRow) {
    var cachedId = cachedRow[2];
    // Verify it still exists in Classroom
    try {
      Classroom.Courses.Topics.get(courseId, cachedId);
      return cachedId; // It's valid! Return it.
    } catch (e) {
      console.log('Cached topic ' + cachedId + ' not found. Will recreate.');
      // If error (404), it was deleted. We fall through to create new.
    }
  }
  
  // 3. Fallback: Search Classroom by Name (for existing un-cached topics)
  var allTopics = [];
  var pageToken = null;
  do {
    var response = Classroom.Courses.Topics.list(courseId, { pageSize: 100, pageToken: pageToken });
    if (response.topic) allTopics = allTopics.concat(response.topic);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  var existingTopic = allTopics.find(function(t) { return t.name === topicName; });
  
  if (existingTopic) {
    // Found by name! Cache it for next time.
    dbSheet.appendRow([courseId, topicName, existingTopic.topicId]);
    return existingTopic.topicId;
  }
  
  // 4. Create New
  var newTopic = Classroom.Courses.Topics.create({ name: topicName }, courseId);
  dbSheet.appendRow([courseId, topicName, newTopic.topicId]);
  return newTopic.topicId;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 11: TIMMY COINS & ECONOMY SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Calculates current Timmy Coin balance.
 * Balance = (Total MP Earned) + (Ledger Adjustments)
 */
function calculateTimmyCoinBalance_(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sum Total MP from Grades DB
  var gradesSheet = ss.getSheetByName("hidden_grades_db");
  var totalMP = 0;
  if (gradesSheet && gradesSheet.getLastRow() > 1) {
    // Read entire DB (Caching strategy used in other functions would be better for bulk, but this is per-student on demand)
    // For optimization, we can use the Grade DB cache if available, but let's read it for now to be safe.
    var data = gradesSheet.getRange(2, 1, gradesSheet.getLastRow() - 1, 8).getValues(); // Col 1=Email, Col 8=MP
    data.forEach(function(row) {
      if (row[0] === email) {
        totalMP += (row[7] || 0);
      }
    });
  }
  
  // 2. Sum Adjustments from Ledger DB
  var ledgerSheet = ss.getSheetByName(CONFIG.SHEETS.LEDGER_DB);
  var ledgerTotal = 0;
  if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
    // Ledger Format: [Timestamp, Email, Amount, Type, Reason]
    var lData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 3).getValues(); // Col 2=Email, Col 3=Amount
    lData.forEach(function(row) {
      if (row[1] === email) {
        ledgerTotal += (row[2] || 0);
      }
    });
  }
  
  return totalMP + ledgerTotal;
}

/**
 * Adds a transaction to the hidden ledger.
 * Amount can be positive (Grant) or negative (Spend/Fine).
 */
function addToLedgerDB_(email, amount, type, reason) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LEDGER_DB);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LEDGER_DB);
    sheet.hideSheet();
    sheet.appendRow(['Timestamp', 'Email', 'Amount', 'Type', 'Reason']);
  }
  
  sheet.appendRow([new Date(), email, amount, type, reason]);
}

/**
 * Processes manual transactions from the Treasury sheet.
 */
function processTreasuryBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TREASURY);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('ℹ️ No Data', 'No transactions found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  var pending = data.filter(function(row) { return row[4] === '⏳ Pending'; });
  
  if (pending.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Pending Transactions', 'Set Status to "⏳ Pending".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var confirm = SpreadsheetApp.getUi().alert(
    '💰 Process Treasury?',
    'Found ' + pending.length + ' pending transaction(s).\nThis will affect student coin balances immediately.',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;
  
  var successCount = 0;
  
  // Iterate
  pending.forEach(function(row) {
    var rowIndex = data.indexOf(row) + 3;
    var email = row[0];
    var amount = Number(row[1]);
    var type = row[2]; // ADD or SUBTRACT
    var reason = row[3];
    
    if (!email || isNaN(amount)) {
      sheet.getRange(rowIndex, 5).setValue('❌ Error');
      return;
    }
    
    // Calculate final Signed Amount
    var finalAmount = (type === 'SUBTRACT') ? -Math.abs(amount) : Math.abs(amount);
    
    try {
      addToLedgerDB_(email, finalAmount, 'Manual ' + type, reason);
      sheet.getRange(rowIndex, 5).setValue('✅ Processed');
      sheet.getRange(rowIndex, 6).setValue(new Date());
      successCount++;
    } catch (e) {
      console.error(e);
      sheet.getRange(rowIndex, 5).setValue('❌ Error');
    }
  });
  
  SpreadsheetApp.getUi().alert('✅ Complete', 'Processed ' + successCount + ' transaction(s).', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 12: AUCTION SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

function findStudentEmailById_(studentId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idStr = studentId.toString().trim();
  
  // Create a regex to match ONLY if the ID is at the start of the email (e.g. "12345@...")
  // or exactly the numeric part if domain varies.
  // Assumes format: 12345@stu.sandi.net
  var idRegex = new RegExp("^" + idStr + "@");
  
  // 1. Try Progress Sheet (Most active students)
  var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (progressSheet && progressSheet.getLastRow() >= 3) {
    var data = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 1).getValues(); // Col A = Email
    for (var i = 0; i < data.length; i++) {
      var email = data[i][0];
      if (email && idRegex.test(email)) {
        return email;
      }
    }
  }
  
  // 2. Fallback: Check Hidden Grades DB (History)
  var dbSheet = ss.getSheetByName("hidden_grades_db");
  if (dbSheet && dbSheet.getLastRow() >= 2) {
     var dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 1).getValues();
     for (var j = 0; j < dbData.length; j++) {
       var dbEmail = dbData[j][0];
       if (dbEmail && idRegex.test(dbEmail)) return dbEmail;
     }
  }
  
  return null;
}

function addToInventoryDB_(email, item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_DB);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.INVENTORY_DB);
    sheet.hideSheet();
    sheet.appendRow(['Timestamp', 'Email', 'Item']);
  }
  
  sheet.appendRow([new Date(), email, item]);
}

function checkAuctionBids() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.AUCTIONS);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  
  var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  
  data.forEach(function(row, index) {
    var isChecked = row[3]; // Column D (Index 3) = Check Bid
    if (isChecked === true) {
      var item = row[0];
      var studentId = row[1];
      var bid = row[2];
      var rowIndex = index + 3;
      
      if (!studentId || !bid) {
        sheet.getRange(rowIndex, 5).setValue('❌ Missing ID or Bid');
        sheet.getRange(rowIndex, 4).uncheck();
        return;
      }
      
      var email = findStudentEmailById_(studentId);
      if (!email) {
        sheet.getRange(rowIndex, 5).setValue('❌ Student Not Found');
        sheet.getRange(rowIndex, 4).uncheck();
        return;
      }
      
      var balance = calculateTimmyCoinBalance_(email);
      
      if (balance >= bid) {
        sheet.getRange(rowIndex, 5).setValue('✅ Approved. Bal: ' + balance);
      } else {
        var short = bid - balance;
        sheet.getRange(rowIndex, 5).setValue('❌ Short by ' + short + ' Coins');
      }
      
      sheet.getRange(rowIndex, 4).uncheck(); 
    }
  });
}

function confirmAuctionTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.AUCTIONS);
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  
  var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  
  data.forEach(function(row, index) {
    var isChecked = row[5]; // Column F (Index 5) = Confirm
    if (isChecked === true) {
      var item = row[0];
      var studentId = row[1];
      var bid = row[2];
      var rowIndex = index + 3;
      
      var email = findStudentEmailById_(studentId);
      if (!email) {
        sheet.getRange(rowIndex, 5).setValue('❌ Error: Student ID invalid');
        sheet.getRange(rowIndex, 6).uncheck();
        return;
      }
      
      var balance = calculateTimmyCoinBalance_(email);
      if (balance < bid) {
        sheet.getRange(rowIndex, 5).setValue('❌ Funds insufficient');
        sheet.getRange(rowIndex, 6).uncheck();
        return;
      }
      
      try {
        addToLedgerDB_(email, -bid, 'Auction', 'Bought: ' + item);
        addToInventoryDB_(email, item);
        
        sheet.getRange(rowIndex, 5).setValue('🎉 Sold! Remaining: ' + (balance - bid));
      } catch (e) {
        console.error(e);
        sheet.getRange(rowIndex, 5).setValue('❌ System Error');
      }
      
      sheet.getRange(rowIndex, 6).uncheck();
    }
  });
}

function formatBonusRoundsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.BONUS_ROUNDS);
  if (!sheet) return;

  applyGlobalDarkTheme_(sheet);

  sheet.getRange('A1:G1').merge().setValue('🎯 BONUS ROUNDS — Configure trivia questions for the class game.');
  applyDescriptionStyle_(sheet.getRange('A1:G1'));

  var headers = ['RoundID', 'Question Text', 'Media URL', 'Type', 'Answer JSON', 'Damage', 'TimeLimit'];
  sheet.getRange('A2:G2').setValues([headers]);
  applyHeaderStyle_(sheet.getRange('A2:G2'));

  sheet.setColumnWidth(1, 100); // RoundID
  sheet.setColumnWidth(2, 300); // Question
  sheet.setColumnWidth(3, 200); // Media
  sheet.setColumnWidth(4, 100); // Type
  sheet.setColumnWidth(5, 300); // JSON
  sheet.setColumnWidth(6, 80);  // Damage
  sheet.setColumnWidth(7, 80);  // Time

  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRIVIA', 'DECISION'], true).build();
  sheet.getRange('D3:D500').setDataValidation(typeRule);

  applyTableBanding_(sheet, 3, 7);
  sheet.setFrozenRows(2);
}

function addSampleBonusQuestions_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.BONUS_ROUNDS);
  if (!sheet) return;

  if (sheet.getLastRow() > 2) return;

  var samples = [
    [Utilities.getUuid(), 'What is the powerhouse of the cell?', 'https://media.giphy.com/media/13FrdxzhRcBqP6/giphy.gif', 'TRIVIA', JSON.stringify({options: ['Nucleus', 'Mitochondria', 'Ribosome', 'Golgi'], correct: 1}), 10, 30],
    [Utilities.getUuid(), 'Should we open the mystery box?', '', 'DECISION', JSON.stringify({options: ['Yes', 'No'], correct: 0}), 20, 15]
  ];

  sheet.getRange(3, 1, samples.length, 7).setValues(samples);
  applyTableBanding_(sheet, 3, 7);
}
