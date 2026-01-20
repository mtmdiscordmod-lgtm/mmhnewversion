/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * MMH OVERLORD SYSTEM — Plugin Gateway & Achievement Engine
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 * This module handles:
 * - doPost endpoint for Render relay communication
 * - Achievement-based MP system (replaces grades-based)
 * - Token/Skey management for anonymous student identification
 * - YouTube search with SafeSearch for song requests
 * - Inventory management via token
 *
 * Version: 1.0.0
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 1: OVERLORD CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

var OVERLORD_CONFIG = {
  SHEETS: {
    TOKEN_REGISTRY: 'hidden_token_registry',
    ACHIEVEMENTS_DB: 'hidden_achievements_db',
    ACHIEVEMENT_RULES: '🎖️ Achievement Rules'
  },

  // SafeSearch toggle - can be changed by teacher
  SAFE_SEARCH: true,

  // Anonymizer word lists (same style as Ghost in the Machine)
  ADJECTIVES: [
    'Swift', 'Brave', 'Clever', 'Mighty', 'Noble', 'Quick', 'Wise', 'Bold',
    'Calm', 'Eager', 'Fierce', 'Grand', 'Happy', 'Jolly', 'Keen', 'Lucky',
    'Merry', 'Neat', 'Proud', 'Royal', 'Sharp', 'Tough', 'Vivid', 'Warm',
    'Zesty', 'Agile', 'Cosmic', 'Daring', 'Epic', 'Flash', 'Gleam', 'Hyper'
  ],

  NOUNS: [
    'Wolf', 'Eagle', 'Tiger', 'Bear', 'Hawk', 'Lion', 'Fox', 'Deer',
    'Owl', 'Raven', 'Shark', 'Dragon', 'Phoenix', 'Falcon', 'Panther', 'Cobra',
    'Storm', 'Blaze', 'Thunder', 'Frost', 'Shadow', 'Spark', 'Comet', 'Nova',
    'Knight', 'Ninja', 'Pilot', 'Scout', 'Hero', 'Sage', 'Ace', 'Star'
  ],

  // WSS URL for skeys
  WSS_URL: 'wss://mmhnet.onrender.com'
};

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 2: doPost ENDPOINT (Relay Gateway)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Main entry point for Render relay requests
 * All plugin/webapp calls come through here
 */
function doPost(e) {
  try {
    var request = JSON.parse(e.postData.contents);
    var action = request.action;
    var params = request.params || {};

    console.log('[Overlord] doPost action:', action);

    var result;

    switch (action) {
      // ========== ACHIEVEMENT SYSTEM ==========
      case 'rewardMP':
        result = rewardMP(params.studentToken, params.achievementId, params.amount);
        break;

      case 'getBalance':
        result = getBalanceByToken(params.studentToken);
        break;

      case 'getAchievements':
        result = getAchievementsByToken(params.studentToken);
        break;

      // ========== INVENTORY SYSTEM ==========
      case 'getInventory':
        result = getInventoryByToken(params.studentToken);
        break;

      case 'addInventory':
        result = addInventoryByToken(params.studentToken, params.item, params.quantity || 1);
        break;

      case 'removeInventory':
        result = removeInventoryByToken(params.studentToken, params.item, params.quantity || 1);
        break;

      case 'hasItem':
        result = hasItemByToken(params.studentToken, params.item);
        break;

      // ========== SONG REQUEST SYSTEM ==========
      case 'searchYouTube':
        result = searchYouTube(params.query, params.maxResults || 5);
        break;

      case 'submitSongRequest':
        result = submitSongRequestByToken(params.studentToken, params.videoId, params.title, params.artist);
        break;

      case 'getSongRequestInventory':
        result = getSongRequestInventoryByToken(params.studentToken);
        break;

      // ========== TOKEN/AUTH SYSTEM ==========
      case 'validateToken':
        result = validateToken(params.studentToken);
        break;

      case 'getStudentInfo':
        result = getStudentInfoByToken(params.studentToken);
        break;

      // ========== GENERIC BACKEND CALL ==========
      case 'callBackend':
        result = callBackendFunction(params.functionName, params.functionParams);
        break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('[Overlord] doPost error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message || 'Unknown error'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Generic backend function caller (power-user multitool)
 */
function callBackendFunction(functionName, functionParams) {
  try {
    // Security: Only allow whitelisted functions
    var allowedFunctions = [
      'getLeaderboardData',
      'getAuctionData',
      'checkStudentBalance',
      'getDataForDashboard'
    ];

    if (allowedFunctions.indexOf(functionName) === -1) {
      return { success: false, error: 'Function not allowed: ' + functionName };
    }

    var func = this[functionName];
    if (typeof func !== 'function') {
      return { success: false, error: 'Function not found: ' + functionName };
    }

    var result = func.apply(null, functionParams || []);
    return { success: true, data: result };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 3: ACHIEVEMENT SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Reward MP for an achievement
 * Checks rules to determine if achievement can be earned
 *
 * @param {string} studentToken - Anonymous token (e.g., "SwiftWolf42")
 * @param {string} achievementId - Achievement ID (e.g., "sf-L0-a1-1")
 * @param {number} amount - MP to award
 * @returns {Object} Result with success, awarded, mp, newBalance, etc.
 */
function rewardMP(studentToken, achievementId, amount) {
  try {
    if (!studentToken || !achievementId || !amount) {
      return { success: false, error: 'Missing required parameters' };
    }

    var numAmount = Number(amount);
    if (isNaN(numAmount) || numAmount <= 0) {
      return { success: false, error: 'Invalid amount' };
    }

    // 1. Get achievement rules
    var rules = getAchievementRules_(achievementId);

    // 2. Check if already earned
    var existingAchievement = checkAchievementEarned_(studentToken, achievementId);

    if (existingAchievement) {
      // Check if repeatable
      if (!rules.repeatable) {
        return {
          success: true,
          awarded: false,
          mp: 0,
          achievementId: achievementId,
          message: 'Achievement already earned',
          firstTime: false,
          newBalance: calculateAchievementBalance_(studentToken)
        };
      }

      // Check cooldown
      if (rules.cooldownHours > 0) {
        var hoursSinceEarned = (Date.now() - new Date(existingAchievement.earnedAt).getTime()) / (1000 * 60 * 60);
        if (hoursSinceEarned < rules.cooldownHours) {
          var hoursRemaining = Math.ceil(rules.cooldownHours - hoursSinceEarned);
          return {
            success: true,
            awarded: false,
            mp: 0,
            achievementId: achievementId,
            message: 'Try again in ' + hoursRemaining + ' hour(s)',
            cooldownRemaining: hoursRemaining,
            firstTime: false,
            newBalance: calculateAchievementBalance_(studentToken)
          };
        }
      }

      // Check max per day
      if (rules.maxPerDay > 0) {
        var todayCount = countAchievementsToday_(studentToken, achievementId);
        if (todayCount >= rules.maxPerDay) {
          return {
            success: true,
            awarded: false,
            mp: 0,
            achievementId: achievementId,
            message: 'Daily limit reached (' + rules.maxPerDay + '/day)',
            firstTime: false,
            newBalance: calculateAchievementBalance_(studentToken)
          };
        }
      }
    }

    // 3. Award the achievement
    logAchievement_(studentToken, achievementId, numAmount, rules.repeatable);

    // 4. Calculate new balance
    var newBalance = calculateAchievementBalance_(studentToken);

    return {
      success: true,
      awarded: true,
      mp: numAmount,
      achievementId: achievementId,
      message: existingAchievement ? 'Achievement earned again!' : 'Achievement unlocked!',
      firstTime: !existingAchievement,
      newBalance: newBalance
    };

  } catch (error) {
    console.error('[rewardMP] Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get achievement rules for an achievement ID
 * Matches against patterns in the rules sheet
 */
function getAchievementRules_(achievementId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENT_RULES);

  // Default rules
  var defaultRules = {
    repeatable: false,
    cooldownHours: 0,
    maxPerDay: 0
  };

  if (!sheet || sheet.getLastRow() < 3) {
    return defaultRules;
  }

  // Columns: Pattern | Repeatable | CooldownHours | MaxPerDay | Notes
  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 4).getValues();

  for (var i = 0; i < data.length; i++) {
    var pattern = String(data[i][0]).trim();
    if (!pattern) continue;

    // Convert pattern to regex
    // * matches any characters
    var regex = new RegExp('^' + pattern.replace(/\*/g, '.*') + '$', 'i');

    if (regex.test(achievementId)) {
      return {
        repeatable: data[i][1] === true || String(data[i][1]).toUpperCase() === 'TRUE',
        cooldownHours: Number(data[i][2]) || 0,
        maxPerDay: Number(data[i][3]) || 0
      };
    }
  }

  return defaultRules;
}

/**
 * Check if achievement was already earned
 */
function checkAchievementEarned_(studentToken, achievementId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  // Columns: Timestamp | StudentToken | AchievementId | MP | Repeatable
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  // Find the most recent matching achievement
  var latest = null;
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][1] === studentToken && data[i][2] === achievementId) {
      latest = {
        earnedAt: data[i][0],
        mp: data[i][3],
        repeatable: data[i][4]
      };
      break;
    }
  }

  return latest;
}

/**
 * Count how many times achievement was earned today
 */
function countAchievementsToday_(studentToken, achievementId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var count = 0;

  for (var i = 0; i < data.length; i++) {
    var earnedDate = new Date(data[i][0]);
    earnedDate.setHours(0, 0, 0, 0);

    if (data[i][1] === studentToken &&
        data[i][2] === achievementId &&
        earnedDate.getTime() === today.getTime()) {
      count++;
    }
  }

  return count;
}

/**
 * Log an achievement to the database
 */
function logAchievement_(studentToken, achievementId, mp, repeatable) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);

  if (!sheet) {
    sheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);
    sheet.hideSheet();
    sheet.appendRow(['Timestamp', 'StudentToken', 'AchievementId', 'MP', 'Repeatable']);
  }

  sheet.appendRow([new Date(), studentToken, achievementId, mp, repeatable || false]);
}

/**
 * Calculate total balance from achievements + ledger
 * REPLACES the old grades-based calculation
 */
function calculateAchievementBalance_(studentToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get email from token (for ledger lookups)
  var email = getEmailFromToken_(studentToken);

  // 2. Sum MP from achievements
  var achievementSheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);
  var achievementTotal = 0;

  if (achievementSheet && achievementSheet.getLastRow() > 1) {
    var data = achievementSheet.getRange(2, 1, achievementSheet.getLastRow() - 1, 4).getValues();
    data.forEach(function(row) {
      if (row[1] === studentToken) {
        achievementTotal += (Number(row[3]) || 0);
      }
    });
  }

  // 3. Sum adjustments from ledger (if email is known)
  var ledgerTotal = 0;
  if (email) {
    var ledgerSheet = ss.getSheetByName(CONFIG.SHEETS.LEDGER_DB);
    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      var lData = ledgerSheet.getRange(2, 1, ledgerSheet.getLastRow() - 1, 3).getValues();
      lData.forEach(function(row) {
        if (row[1] === email) {
          ledgerTotal += (Number(row[2]) || 0);
        }
      });
    }
  }

  return achievementTotal + ledgerTotal;
}

/**
 * Get balance by token (public API)
 */
function getBalanceByToken(studentToken) {
  try {
    if (!studentToken) {
      return { success: false, error: 'Missing studentToken' };
    }

    var balance = calculateAchievementBalance_(studentToken);
    return { success: true, balance: balance, studentToken: studentToken };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Get all achievements for a token
 */
function getAchievementsByToken(studentToken) {
  try {
    if (!studentToken) {
      return { success: false, error: 'Missing studentToken' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);
    var achievements = [];

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
      data.forEach(function(row) {
        if (row[1] === studentToken) {
          achievements.push({
            earnedAt: row[0],
            achievementId: row[2],
            mp: row[3]
          });
        }
      });
    }

    return { success: true, achievements: achievements, count: achievements.length };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 4: TOKEN/SKEY SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate an anonymized token for a student
 * Format: AdjectiveNoun## (e.g., "SwiftWolf42")
 */
function generateAnonymizedToken_() {
  var adjective = OVERLORD_CONFIG.ADJECTIVES[Math.floor(Math.random() * OVERLORD_CONFIG.ADJECTIVES.length)];
  var noun = OVERLORD_CONFIG.NOUNS[Math.floor(Math.random() * OVERLORD_CONFIG.NOUNS.length)];
  var number = Math.floor(Math.random() * 100);
  var numStr = number < 10 ? '0' + number : String(number);

  return adjective + noun + numStr;
}

/**
 * Get or create a token mapping for a student
 *
 * @param {string} studentId - Google Classroom student ID
 * @param {number} period - Class period
 * @returns {string} The student's token
 */
function getOrCreateTokenMapping_(studentId, period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);

  if (!sheet) {
    sheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
    sheet.hideSheet();
    sheet.appendRow(['StudentId', 'AnonymizedToken', 'Period', 'Email', 'GeneratedAt']);
  }

  // Check if student already has a token
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId)) {
        return data[i][1]; // Return existing token
      }
    }
  }

  // Generate new unique token
  var newToken;
  var attempts = 0;
  var existingTokens = [];

  if (sheet.getLastRow() > 1) {
    existingTokens = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  }

  do {
    newToken = generateAnonymizedToken_();
    attempts++;
  } while (existingTokens.indexOf(newToken) !== -1 && attempts < 100);

  // Get email if possible
  var email = '';
  try {
    email = studentId + '@stu.sandi.net'; // Default format
  } catch (e) {}

  // Save mapping
  sheet.appendRow([studentId, newToken, period || '', email, new Date()]);

  return newToken;
}

/**
 * Get email from token
 */
function getEmailFromToken_(studentToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === studentToken) {
      return data[i][3] || (data[i][0] + '@stu.sandi.net');
    }
  }

  return null;
}

/**
 * Get student ID from token
 */
function getStudentIdFromToken_(studentToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === studentToken) {
      return data[i][0];
    }
  }

  return null;
}

/**
 * Validate a token exists
 */
function validateToken(studentToken) {
  var studentId = getStudentIdFromToken_(studentToken);
  return {
    success: true,
    valid: !!studentId,
    studentToken: studentToken
  };
}

/**
 * Get student info by token (non-PII)
 */
function getStudentInfoByToken(studentToken) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, error: 'Token not found' };
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === studentToken) {
      return {
        success: true,
        studentToken: studentToken,
        period: data[i][2],
        registeredAt: data[i][4]
      };
    }
  }

  return { success: false, error: 'Token not found' };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 5: SKEY DEPLOYMENT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate skey data for a student
 */
function generateSkey_(studentId, period) {
  var token = getOrCreateTokenMapping_(studentId, period);

  return {
    version: '1.0',
    studentToken: token,
    period: period,
    wssUrl: OVERLORD_CONFIG.WSS_URL,
    generatedAt: new Date().toISOString()
  };
}

/**
 * Obfuscate skey data (Base64)
 */
function obfuscateSkey_(skeyData) {
  var jsonStr = JSON.stringify(skeyData);
  return Utilities.base64Encode(jsonStr, Utilities.Charset.UTF_8);
}

/**
 * Deobfuscate skey data
 */
function deobfuscateSkey_(encoded) {
  var decoded = Utilities.base64Decode(encoded, Utilities.Charset.UTF_8);
  var jsonStr = Utilities.newBlob(decoded).getDataAsString();
  return JSON.parse(jsonStr);
}

/**
 * Deploy skeys for an entire period
 * Creates .skey files in Google Drive and shares with students
 */
function deploySkeysForPeriod(period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = { success: 0, failed: 0, errors: [] };

  // 1. Get students for this period from Progress sheet
  var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (!progressSheet || progressSheet.getLastRow() < 3) {
    return { success: false, error: 'No students found' };
  }

  var data = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 3).getValues();
  var students = [];
  var seen = {};

  data.forEach(function(row) {
    var email = row[0];
    var studentPeriod = String(row[2]).replace(/^Period\s*/i, '');

    if (email && !seen[email] && studentPeriod === String(period)) {
      seen[email] = true;
      // Extract student ID from email (assumes format: ID@stu.sandi.net)
      var studentId = email.split('@')[0];
      students.push({ email: email, studentId: studentId, period: period });
    }
  });

  if (students.length === 0) {
    return { success: false, error: 'No students found for period ' + period };
  }

  // 2. Create/get folder structure
  var rootFolder = getOrCreateFolder_('Multimedia Heroes');
  var skeyFolder = getOrCreateSubfolder_(rootFolder, 'MMH_Skeys');
  var periodFolder = getOrCreateSubfolder_(skeyFolder, 'Period_' + period);

  // 3. Generate and deploy skeys
  students.forEach(function(student) {
    try {
      var skeyData = generateSkey_(student.studentId, period);
      var obfuscated = obfuscateSkey_(skeyData);

      // Create file
      var fileName = skeyData.studentToken + '.skey';
      var existingFiles = periodFolder.getFilesByName(fileName);

      if (existingFiles.hasNext()) {
        // Update existing file
        var file = existingFiles.next();
        file.setContent(obfuscated);
      } else {
        // Create new file
        var file = periodFolder.createFile(fileName, obfuscated, 'text/plain');

        // Share with student
        try {
          file.addViewer(student.email);
        } catch (shareErr) {
          console.log('Could not share with ' + student.email + ': ' + shareErr.message);
        }
      }

      results.success++;

    } catch (e) {
      results.failed++;
      results.errors.push(student.email + ': ' + e.message);
    }
  });

  return {
    success: true,
    deployed: results.success,
    failed: results.failed,
    total: students.length,
    folderUrl: periodFolder.getUrl(),
    errors: results.errors
  };
}

/**
 * Get deployment status for all periods
 */
function getSkeyDeploymentStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);

  var status = {};

  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    data.forEach(function(row) {
      var period = row[2] || 'Unknown';
      if (!status[period]) {
        status[period] = { count: 0, tokens: [] };
      }
      status[period].count++;
      status[period].tokens.push(row[1]);
    });
  }

  return { success: true, periods: status };
}

// Folder helpers
function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateSubfolder_(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: INVENTORY SYSTEM (Token-Based)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get inventory by token
 */
function getInventoryByToken(studentToken) {
  try {
    var email = getEmailFromToken_(studentToken);
    if (!email) {
      return { success: false, error: 'Invalid token' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_DB);
    var inventory = {};

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      data.forEach(function(row) {
        if (row[1] === email) {
          var item = row[2];
          // Handle consumed items
          if (item.indexOf('Consumed ') === 0) {
            var originalItem = item.replace('Consumed ', '');
            inventory[originalItem] = (inventory[originalItem] || 0) - 1;
          } else {
            inventory[item] = (inventory[item] || 0) + 1;
          }
        }
      });
    }

    // Filter out zero/negative counts
    var cleanInventory = {};
    Object.keys(inventory).forEach(function(item) {
      if (inventory[item] > 0) {
        cleanInventory[item] = inventory[item];
      }
    });

    return { success: true, inventory: cleanInventory };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Add item to inventory by token
 */
function addInventoryByToken(studentToken, item, quantity) {
  try {
    var email = getEmailFromToken_(studentToken);
    if (!email) {
      return { success: false, error: 'Invalid token' };
    }

    quantity = quantity || 1;
    for (var i = 0; i < quantity; i++) {
      addToInventoryDB_(email, item);
    }

    return { success: true, item: item, quantity: quantity };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Remove item from inventory by token
 */
function removeInventoryByToken(studentToken, item, quantity) {
  try {
    var email = getEmailFromToken_(studentToken);
    if (!email) {
      return { success: false, error: 'Invalid token' };
    }

    // Check current count
    var inventoryResult = getInventoryByToken(studentToken);
    var currentCount = (inventoryResult.inventory && inventoryResult.inventory[item]) || 0;

    if (currentCount < quantity) {
      return { success: false, error: 'Insufficient inventory', current: currentCount, requested: quantity };
    }

    quantity = quantity || 1;
    for (var i = 0; i < quantity; i++) {
      addToInventoryDB_(email, 'Consumed ' + item);
    }

    return { success: true, item: item, quantity: quantity, remaining: currentCount - quantity };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Check if student has item
 */
function hasItemByToken(studentToken, item) {
  var inventoryResult = getInventoryByToken(studentToken);
  var count = (inventoryResult.inventory && inventoryResult.inventory[item]) || 0;
  return { success: true, hasItem: count > 0, count: count };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 7: YOUTUBE SEARCH
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Search YouTube with SafeSearch
 */
function searchYouTube(query, maxResults) {
  try {
    if (!query) {
      return { success: false, error: 'Missing search query' };
    }

    maxResults = maxResults || 5;

    // Get SafeSearch setting
    var safeSearch = getSafeSearchSetting_();

    var searchParams = {
      q: query,
      maxResults: maxResults,
      type: 'video',
      part: 'snippet',
      safeSearch: safeSearch ? 'strict' : 'none',
      videoEmbeddable: true
    };

    var results = YouTube.Search.list('snippet', searchParams);

    if (!results.items || results.items.length === 0) {
      return { success: true, videos: [], message: 'No results found' };
    }

    var videos = results.items.map(function(item) {
      return {
        videoId: item.id.videoId,
        title: item.snippet.title,
        channelTitle: item.snippet.channelTitle,
        thumbnail: item.snippet.thumbnails.medium ? item.snippet.thumbnails.medium.url : item.snippet.thumbnails.default.url,
        description: item.snippet.description.substring(0, 100)
      };
    });

    return { success: true, videos: videos, safeSearch: safeSearch };

  } catch (error) {
    console.error('[searchYouTube] Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get SafeSearch setting from config
 */
function getSafeSearchSetting_() {
  // Could be stored in Properties or a config sheet
  // For now, use the default from OVERLORD_CONFIG
  var props = PropertiesService.getScriptProperties();
  var setting = props.getProperty('SAFE_SEARCH');

  if (setting === null) {
    return OVERLORD_CONFIG.SAFE_SEARCH;
  }

  return setting === 'true';
}

/**
 * Toggle SafeSearch setting (teacher function)
 */
function toggleSafeSearch() {
  var props = PropertiesService.getScriptProperties();
  var current = getSafeSearchSetting_();
  var newValue = !current;

  props.setProperty('SAFE_SEARCH', String(newValue));

  return { success: true, safeSearch: newValue };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 8: SONG REQUEST (Token-Based)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get song request inventory count by token
 */
function getSongRequestInventoryByToken(studentToken) {
  var inventoryResult = getInventoryByToken(studentToken);
  var count = (inventoryResult.inventory && inventoryResult.inventory['Song Request']) || 0;
  return { success: true, count: count };
}

/**
 * Submit song request by token
 */
function submitSongRequestByToken(studentToken, videoId, title, artist) {
  try {
    // 1. Verify inventory
    var inventoryCheck = getSongRequestInventoryByToken(studentToken);
    if (!inventoryCheck.success || inventoryCheck.count < 1) {
      return { success: false, error: 'No Song Request items in inventory' };
    }

    // 2. Get student info
    var email = getEmailFromToken_(studentToken);
    var studentInfo = getStudentInfoByToken(studentToken);
    var period = studentInfo.period || '?';

    // 3. Consume inventory item
    removeInventoryByToken(studentToken, 'Song Request', 1);

    // 4. Add to song requests
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);

    if (!sheet) {
      return { success: false, error: 'Song Requests sheet not found' };
    }

    var link = videoId ? 'https://www.youtube.com/watch?v=' + videoId : '';

    // Columns: Timestamp, Email, Name, Period, Artist, Song, Link, Status
    sheet.appendRow([
      new Date(),
      email || studentToken,
      studentToken, // Use token as display name (anonymous)
      period,
      artist || '',
      title || '',
      link,
      'PENDING'
    ]);

    return {
      success: true,
      message: 'Song request submitted!',
      remainingRequests: inventoryCheck.count - 1
    };

  } catch (error) {
    console.error('[submitSongRequestByToken] Error:', error);
    return { success: false, error: error.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 9: SHEET FORMATTING FOR NEW SHEETS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Format the Achievement Rules sheet
 */
function formatAchievementRulesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENT_RULES);

  if (!sheet) {
    sheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.ACHIEVEMENT_RULES);
  }

  // Apply global dark theme if function exists
  if (typeof applyGlobalDarkTheme_ === 'function') {
    applyGlobalDarkTheme_(sheet);
  }

  // Description row
  sheet.getRange('A1:E1').merge().setValue('🎖️ ACHIEVEMENT RULES — Configure which achievements are repeatable and their cooldowns. Use * as wildcard.');
  if (typeof applyDescriptionStyle_ === 'function') {
    applyDescriptionStyle_(sheet.getRange('A1:E1'));
  }

  // Headers
  var headers = ['Pattern', 'Repeatable', 'CooldownHours', 'MaxPerDay', 'Notes'];
  sheet.getRange('A2:E2').setValues([headers]);
  if (typeof applyHeaderStyle_ === 'function') {
    applyHeaderStyle_(sheet.getRange('A2:E2'));
  }

  // Column widths
  sheet.setColumnWidth(1, 200);  // Pattern
  sheet.setColumnWidth(2, 100);  // Repeatable
  sheet.setColumnWidth(3, 120);  // CooldownHours
  sheet.setColumnWidth(4, 100);  // MaxPerDay
  sheet.setColumnWidth(5, 300);  // Notes

  // Add default rules if empty
  if (sheet.getLastRow() < 3) {
    var defaultRules = [
      ['*-L*-a*-*', false, 0, 0, 'Standard track achievements (one-time)'],
      ['daily-*', true, 24, 1, 'Daily challenges (once per day)'],
      ['practice-*', true, 0, 3, 'Practice achievements (up to 3x/day)'],
      ['bonus-*', true, 0, 0, 'Bonus round achievements (unlimited)']
    ];
    sheet.getRange(3, 1, defaultRules.length, 5).setValues(defaultRules);
  }

  // Data validation for Repeatable column
  var repeatableRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  sheet.getRange('B3:B100').setDataValidation(repeatableRule);
}

/**
 * Setup all Overlord sheets
 */
function setupOverlordSheets() {
  formatAchievementRulesSheet_();

  // Ensure hidden sheets exist
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tokenSheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
  if (!tokenSheet) {
    tokenSheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
    tokenSheet.hideSheet();
    tokenSheet.appendRow(['StudentId', 'AnonymizedToken', 'Period', 'Email', 'GeneratedAt']);
  }

  var achievementSheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);
  if (!achievementSheet) {
    achievementSheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.ACHIEVEMENTS_DB);
    achievementSheet.hideSheet();
    achievementSheet.appendRow(['Timestamp', 'StudentToken', 'AchievementId', 'MP', 'Repeatable']);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Overlord sheets configured!', '✅ Complete', 5);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 10: MENU ADDITIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Add Overlord menu items (call from main onOpen)
 */
function addOverlordMenuItems_(menu) {
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu('🎮 MMH Overlord')
    .addItem('⚙️ Setup Overlord Sheets', 'setupOverlordSheets')
    .addItem('🔑 Deploy Skeys for Period...', 'showDeploySkeyDialog')
    .addItem('📊 View Deployment Status', 'showDeploymentStatus')
    .addSeparator()
    .addItem('🔍 Toggle SafeSearch', 'toggleSafeSearchUI'));

  return menu;
}

/**
 * Show dialog for skey deployment
 */
function showDeploySkeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '🔑 Deploy Skeys',
    'Enter the period number to deploy skeys for:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var period = response.getResponseText().trim();
    if (period) {
      var result = deploySkeysForPeriod(period);
      if (result.success) {
        ui.alert('✅ Deployment Complete',
          'Successfully deployed: ' + result.deployed + '\n' +
          'Failed: ' + result.failed + '\n' +
          'Total students: ' + result.total + '\n\n' +
          'Folder: ' + result.folderUrl,
          ui.ButtonSet.OK);
      } else {
        ui.alert('❌ Deployment Failed', result.error, ui.ButtonSet.OK);
      }
    }
  }
}

/**
 * Show deployment status
 */
function showDeploymentStatus() {
  var status = getSkeyDeploymentStatus();
  var ui = SpreadsheetApp.getUi();

  var message = 'Skey Deployment Status:\n\n';
  Object.keys(status.periods).forEach(function(period) {
    message += 'Period ' + period + ': ' + status.periods[period].count + ' students\n';
  });

  ui.alert('📊 Deployment Status', message, ui.ButtonSet.OK);
}

/**
 * Toggle SafeSearch UI
 */
function toggleSafeSearchUI() {
  var result = toggleSafeSearch();
  var ui = SpreadsheetApp.getUi();

  ui.alert('🔍 SafeSearch',
    'SafeSearch is now ' + (result.safeSearch ? 'ENABLED' : 'DISABLED'),
    ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 11: TEST STUDENTS SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test students are external emails (non-district) that can:
 * - Use the portal/song request system
 * - Have MP/inventory adjusted via dashboard
 * - Don't interact with Google Classroom
 * - Great for demos, showing district people, testing
 */

var TEST_STUDENTS_SHEET = '🧪 Test Students';

/**
 * Create a test student with external email
 * @param {string} name - Display name for the test student
 * @param {string} email - External email (e.g., jasonbermanjazz@gmail.com)
 * @param {string} period - Simulated period (for organization)
 * @param {string} notes - Optional notes about this test student
 */
function createTestStudent(name, email, period, notes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureTestStudentsSheet_();

  // Check if email already exists
  if (sheet.getLastRow() > 2) {
    var existingData = sheet.getRange(3, 2, sheet.getLastRow() - 2, 1).getValues();
    for (var i = 0; i < existingData.length; i++) {
      if (existingData[i][0] === email) {
        return { success: false, error: 'Test student with this email already exists' };
      }
    }
  }

  // Generate unique token
  var token = generateAnonymizedToken_();
  var existingTokens = [];

  // Get existing tokens from both registries
  var tokenSheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
  if (tokenSheet && tokenSheet.getLastRow() > 1) {
    existingTokens = existingTokens.concat(
      tokenSheet.getRange(2, 2, tokenSheet.getLastRow() - 1, 1).getValues().flat()
    );
  }
  if (sheet.getLastRow() > 2) {
    existingTokens = existingTokens.concat(
      sheet.getRange(3, 3, sheet.getLastRow() - 2, 1).getValues().flat()
    );
  }

  // Ensure unique token
  var attempts = 0;
  while (existingTokens.indexOf(token) !== -1 && attempts < 100) {
    token = generateAnonymizedToken_();
    attempts++;
  }

  // Add test student
  // Columns: Name | Email | Token | Period | Notes | CreatedAt | IsTest
  sheet.appendRow([name, email, token, period || 'Test', notes || '', new Date(), true]);

  // Also add to token registry for consistent lookups
  var registrySheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
  if (!registrySheet) {
    registrySheet = ss.insertSheet(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
    registrySheet.hideSheet();
    registrySheet.appendRow(['StudentId', 'AnonymizedToken', 'Period', 'Email', 'GeneratedAt', 'IsTest']);
  }
  registrySheet.appendRow(['TEST_' + email, token, period || 'Test', email, new Date(), true]);

  return {
    success: true,
    name: name,
    email: email,
    token: token,
    period: period,
    message: 'Test student created successfully!'
  };
}

/**
 * Get all test students
 */
function getTestStudents() {
  var sheet = ensureTestStudentsSheet_();
  var students = [];

  if (sheet.getLastRow() > 2) {
    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 7).getValues();
    data.forEach(function(row) {
      students.push({
        name: row[0],
        email: row[1],
        token: row[2],
        period: row[3],
        notes: row[4],
        createdAt: row[5]
      });
    });
  }

  return { success: true, students: students };
}

/**
 * Delete a test student
 */
function deleteTestStudent(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_STUDENTS_SHEET);

  if (!sheet || sheet.getLastRow() < 3) {
    return { success: false, error: 'Test student not found' };
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === email) {
      sheet.deleteRow(i + 3);

      // Also remove from token registry
      var registrySheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
      if (registrySheet && registrySheet.getLastRow() > 1) {
        var regData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 4).getValues();
        for (var j = 0; j < regData.length; j++) {
          if (regData[j][3] === email) {
            registrySheet.deleteRow(j + 2);
            break;
          }
        }
      }

      return { success: true, message: 'Test student deleted' };
    }
  }

  return { success: false, error: 'Test student not found' };
}

/**
 * Generate and download skey for a test student
 */
function generateTestStudentSkey(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_STUDENTS_SHEET);

  if (!sheet || sheet.getLastRow() < 3) {
    return { success: false, error: 'Test student not found' };
  }

  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === email) {
      var skeyData = {
        version: '1.0',
        studentToken: data[i][2],
        period: data[i][3],
        wssUrl: OVERLORD_CONFIG.WSS_URL,
        isTest: true,
        generatedAt: new Date().toISOString()
      };

      var obfuscated = obfuscateSkey_(skeyData);

      return {
        success: true,
        skey: obfuscated,
        token: data[i][2],
        fileName: data[i][2] + '.skey'
      };
    }
  }

  return { success: false, error: 'Test student not found' };
}

/**
 * Deploy skey to Google Drive for a test student
 */
function deployTestStudentSkey(email) {
  var result = generateTestStudentSkey(email);
  if (!result.success) {
    return result;
  }

  var rootFolder = getOrCreateFolder_('Multimedia Heroes');
  var skeyFolder = getOrCreateSubfolder_(rootFolder, 'MMH_Skeys');
  var testFolder = getOrCreateSubfolder_(skeyFolder, 'Test_Students');

  var fileName = result.fileName;
  var existingFiles = testFolder.getFilesByName(fileName);

  var file;
  if (existingFiles.hasNext()) {
    file = existingFiles.next();
    file.setContent(result.skey);
  } else {
    file = testFolder.createFile(fileName, result.skey, 'text/plain');

    // Try to share with the email
    try {
      file.addViewer(email);
    } catch (e) {
      console.log('Could not share with ' + email + ': ' + e.message);
    }
  }

  return {
    success: true,
    fileUrl: file.getUrl(),
    token: result.token,
    message: 'Skey deployed to Google Drive'
  };
}

/**
 * Ensure test students sheet exists
 */
function ensureTestStudentsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_STUDENTS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TEST_STUDENTS_SHEET);

    // Apply dark theme if available
    if (typeof applyGlobalDarkTheme_ === 'function') {
      applyGlobalDarkTheme_(sheet);
    }

    // Description row
    sheet.getRange('A1:G1').merge().setValue('🧪 TEST STUDENTS — External emails for demos & testing. No Classroom integration.');
    if (typeof applyDescriptionStyle_ === 'function') {
      applyDescriptionStyle_(sheet.getRange('A1:G1'));
    }

    // Headers
    var headers = ['Name', 'Email', 'Token', 'Period', 'Notes', 'Created', 'IsTest'];
    sheet.getRange('A2:G2').setValues([headers]);
    if (typeof applyHeaderStyle_ === 'function') {
      applyHeaderStyle_(sheet.getRange('A2:G2'));
    }

    // Column widths
    sheet.setColumnWidth(1, 150);  // Name
    sheet.setColumnWidth(2, 250);  // Email
    sheet.setColumnWidth(3, 120);  // Token
    sheet.setColumnWidth(4, 80);   // Period
    sheet.setColumnWidth(5, 200);  // Notes
    sheet.setColumnWidth(6, 150);  // Created
    sheet.setColumnWidth(7, 60);   // IsTest
  }

  return sheet;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 12: MANUAL ADJUSTMENTS (Teacher Dashboard)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get all students (regular + test) for the dashboard
 */
function getAllStudentsForDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var students = [];

  // Get regular students from token registry
  var tokenSheet = ss.getSheetByName(OVERLORD_CONFIG.SHEETS.TOKEN_REGISTRY);
  if (tokenSheet && tokenSheet.getLastRow() > 1) {
    var data = tokenSheet.getRange(2, 1, tokenSheet.getLastRow() - 1, 6).getValues();
    data.forEach(function(row) {
      var isTest = row[5] === true || String(row[0]).indexOf('TEST_') === 0;
      students.push({
        studentId: row[0],
        token: row[1],
        period: row[2],
        email: row[3],
        createdAt: row[4],
        isTest: isTest,
        balance: 0 // Will be calculated
      });
    });
  }

  // Calculate balances
  students.forEach(function(student) {
    student.balance = calculateAchievementBalance_(student.token);
    var inv = getInventoryByToken(student.token);
    student.inventory = inv.inventory || {};
  });

  return { success: true, students: students };
}

/**
 * Search students by token or email
 */
function searchStudentsOverlord(query) {
  var result = getAllStudentsForDashboard();
  if (!result.success) return result;

  var q = query.toLowerCase();
  var filtered = result.students.filter(function(s) {
    return (s.token && s.token.toLowerCase().indexOf(q) !== -1) ||
           (s.email && s.email.toLowerCase().indexOf(q) !== -1) ||
           (s.studentId && String(s.studentId).toLowerCase().indexOf(q) !== -1);
  });

  return { success: true, students: filtered };
}

/**
 * Manually adjust MP for a student (by token)
 * @param {string} studentToken - The student's token
 * @param {number} amount - Amount to add (positive) or subtract (negative)
 * @param {string} reason - Reason for adjustment
 */
function adjustStudentMP(studentToken, amount, reason) {
  try {
    var numAmount = Number(amount);
    if (isNaN(numAmount) || numAmount === 0) {
      return { success: false, error: 'Invalid amount' };
    }

    // Log as a special "manual" achievement
    var achievementId = 'manual-adjustment-' + Date.now();
    logAchievement_(studentToken, achievementId, numAmount, false);

    // Also log to ledger for audit trail
    var email = getEmailFromToken_(studentToken);
    if (email) {
      addToLedgerDB_(email, 0, 'MP Adjustment', reason + ' (Token: ' + studentToken + ', MP: ' + numAmount + ')');
    }

    var newBalance = calculateAchievementBalance_(studentToken);

    return {
      success: true,
      amount: numAmount,
      newBalance: newBalance,
      message: (numAmount > 0 ? 'Added ' : 'Removed ') + Math.abs(numAmount) + ' MP'
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Manually add inventory item for a student (by token)
 */
function addStudentInventory(studentToken, item, quantity) {
  return addInventoryByToken(studentToken, item, quantity || 1);
}

/**
 * Manually remove inventory item for a student (by token)
 */
function removeStudentInventory(studentToken, item, quantity) {
  return removeInventoryByToken(studentToken, item, quantity || 1);
}

/**
 * Get student details by token (for dashboard)
 */
function getStudentDetailsByToken(studentToken) {
  var balance = calculateAchievementBalance_(studentToken);
  var inventory = getInventoryByToken(studentToken);
  var achievements = getAchievementsByToken(studentToken);
  var info = getStudentInfoByToken(studentToken);

  return {
    success: true,
    token: studentToken,
    period: info.period,
    balance: balance,
    inventory: inventory.inventory || {},
    achievements: achievements.achievements || [],
    isTest: info.studentToken ? (String(getStudentIdFromToken_(studentToken)).indexOf('TEST_') === 0) : false
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 13: WEBAPP ENDPOINTS FOR DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get data for MMH Overlord dashboard tab
 */
function getOverlordDashboardData() {
  var status = getSkeyDeploymentStatus();
  var testStudents = getTestStudents();
  var safeSearch = getSafeSearchSetting_();

  return {
    success: true,
    deploymentStatus: status.periods || {},
    testStudents: testStudents.students || [],
    safeSearch: safeSearch,
    periods: [3, 4, 5, 6, 7] // From CONFIG
  };
}

/**
 * Deploy skeys for a period (called from dashboard)
 */
function deploySkeysPeriodFromDashboard(period) {
  return deploySkeysForPeriod(period);
}

/**
 * Create test student from dashboard
 */
function createTestStudentFromDashboard(name, email, period, notes) {
  return createTestStudent(name, email, period, notes);
}

/**
 * Delete test student from dashboard
 */
function deleteTestStudentFromDashboard(email) {
  return deleteTestStudent(email);
}

/**
 * Deploy test student skey from dashboard
 */
function deployTestStudentSkeyFromDashboard(email) {
  return deployTestStudentSkey(email);
}

/**
 * Adjust MP from dashboard
 */
function adjustMPFromDashboard(studentToken, amount, reason) {
  return adjustStudentMP(studentToken, amount, reason);
}

/**
 * Add inventory from dashboard
 */
function addInventoryFromDashboard(studentToken, item, quantity) {
  return addStudentInventory(studentToken, item, quantity);
}

/**
 * Remove inventory from dashboard
 */
function removeInventoryFromDashboard(studentToken, item, quantity) {
  return removeStudentInventory(studentToken, item, quantity);
}

/**
 * Get student details from dashboard
 */
function getStudentDetailsFromDashboard(studentToken) {
  return getStudentDetailsByToken(studentToken);
}

/**
 * Search students from dashboard
 */
function searchStudentsFromDashboard(query) {
  return searchStudentsOverlord(query);
}
