/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * MULTIMEDIA HEROES — STUDENT REGISTRY MODULE
 *
 * Pulls student data from the district roster spreadsheet and creates a reliable
 * email lookup that bypasses Classroom API permission issues.
 * ═══════════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

function diagnoseNameMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  
  // Get first active classroom
  var classroomData = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 6).getValues();
  var activeClassroom = classroomData.find(function(row) {
    return row[4] === 'TRUE' || row[4] === true;
  });
  
  if (!activeClassroom) {
    console.log('No active classroom found');
    return;
  }
  
  var courseId = activeClassroom[0];
  var courseName = activeClassroom[1];
  
  console.log('Checking classroom: ' + courseName);
  console.log('═══════════════════════════════════════════════════════════');
  
  // Get students directly from Classroom API (not enriched)
  var students = [];
  var pageToken = null;
  do {
    var response = Classroom.Courses.Students.list(courseId, { pageSize: 100, pageToken: pageToken });
    if (response.students) students = students.concat(response.students);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  console.log('Classroom has ' + students.length + ' students');
  console.log('');
  
  // Get registry
  var registry = getStudentRegistry_();
  console.log('Registry has ' + registry.all.length + ' students');
  console.log('');
  
  // Compare
  var matched = 0;
  var unmatched = 0;
  
  console.log('UNMATCHED STUDENTS (Classroom name → no registry match):');
  console.log('─────────────────────────────────────────────────────────');
  
  students.forEach(function(student) {
    var classroomName = student.profile.name ? student.profile.name.fullName : 'Unknown';
    var classroomEmail = student.profile.emailAddress || '(no email from API)';
    
    // Try to find in registry
    var email = lookupStudentEmail_(classroomName, student.userId);
    
    if (email) {
      matched++;
    } else {
      unmatched++;
      console.log('❌ "' + classroomName + '"');
    }
  });
  
  console.log('');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('SUMMARY:');
  console.log('✅ Matched: ' + matched);
  console.log('❌ Unmatched: ' + unmatched);
  console.log('');
  
  if (unmatched > 0) {
    console.log('SAMPLE REGISTRY NAMES (for comparison):');
    registry.all.slice(0, 10).forEach(function(s) {
      console.log('   "' + s.name + '"');
    });
  }
}
var ROSTER_CONFIG = {
  // External roster spreadsheet
  ROSTER_SPREADSHEET_ID: '1pw1hKeHTE6Qsmfyryw2TDkbre--Jrmz5X2_TF-MjoJQ',
  ROSTER_SHEET_NAME: 'Roster',

  // Teacher email (used to filter students)
  TEACHER_EMAIL: 'jberman2@sandi.net',

  // Which periods are Mr. B's
  MY_PERIODS: [3, 4, 5, 6, 7],

  // Column mapping (0-indexed)
  // Roster columns: Student Name, Student ID, Special Status, Restricted, Break Limit, P1, P2, P3, P4, P5, P6, P7
  COLUMNS: {
    STUDENT_NAME: 0,
    STUDENT_ID: 1,
    SPECIAL_STATUS: 2,
    RESTRICTED: 3,
    BREAK_LIMIT: 4,
    PERIOD_1: 5,
    PERIOD_2: 6,
    PERIOD_3: 7,
    PERIOD_4: 8,
    PERIOD_5: 9,
    PERIOD_6: 10,
    PERIOD_7: 11
  },

  // Student email domain
  EMAIL_DOMAIN: '@stu.sandi.net',

  // Local cache sheet name
  CACHE_SHEET_NAME: 'hidden_student_registry'
};

var OVERRIDE_SHEET_NAME = 'hidden_student_overrides';


// ═══════════════════════════════════════════════════════════════════════════════
// MANUAL OVERRIDE SYSTEM
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get all manual overrides
 * Returns object: { classroomUserId: { email, name, note } }
 */
function getStudentOverrides_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERRIDE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var overrides = {};

  data.forEach(function(row) {
    var oderId = row[0];
    if (row[0]) {
      overrides[row[0].toString()] = {
        email: row[1],
        classroomName: row[2],
        rosterName: row[3],
        note: row[4]
      };
    }
  });

  return overrides;
}

/**
 * Create the overrides sheet if it doesn't exist
 */
function ensureOverrideSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERRIDE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(OVERRIDE_SHEET_NAME);
    sheet.appendRow(['Classroom UserID', 'Student Email', 'Classroom Name', 'Roster Name', 'Note']);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#E8E8E8');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 200);
    sheet.hideSheet();
  }

  return sheet;
}

/**
 * Add or update a manual override
 */
function setStudentOverride(classroomUserId, email, classroomName, rosterName, note) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureOverrideSheet_();

  // Check if override already exists
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0].toString() === classroomUserId.toString()) {
        // Update existing row
        sheet.getRange(2 + i, 2, 1, 4).setValues([[email, classroomName, rosterName, note]]);
        return { success: true, action: 'updated' };
      }
    }
  }

  // Add new row
  sheet.appendRow([classroomUserId, email, classroomName, rosterName, note || '']);
  return { success: true, action: 'added' };
}

/**
 * Remove an override
 */
function removeStudentOverride(classroomUserId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OVERRIDE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, error: 'No overrides exist' };
  }

  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0].toString() === classroomUserId.toString()) {
      sheet.deleteRow(2 + i);
      return { success: true };
    }
  }

  return { success: false, error: 'Override not found' };
}

/**
 * Get unmatched students across all active classrooms
 * Used by the Config UI to show what needs manual resolution
 */
function getUnmatchedStudents() {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);

  if (!classroomsSheet || classroomsSheet.getLastRow() < 3) {
    return [];
  }

  var classroomData = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 6).getValues();
  var activeClassrooms = classroomData.filter(function(row) {
    return row[4] === 'TRUE' || row[4] === true;
  });

  var registry = getStudentRegistry_();
  var overrides = getStudentOverrides_();
  var unmatched = [];
  var seen = {};

  activeClassrooms.forEach(function(classroom) {
    var courseId = classroom[0];
    var period = classroom[3];

    try {
      var students = [];
      var pageToken = null;
      do {
        var response = Classroom.Courses.Students.list(courseId, { pageSize: 100, pageToken: pageToken });
        if (response.students) students = students.concat(response.students);
        pageToken = response.nextPageToken;
      } while (pageToken);

      students.forEach(function(student) {
        var userId = student.userId;
        if (seen[userId]) return;
        seen[userId] = true;

        var classroomName = student.profile.name ? student.profile.name.fullName : 'Unknown';
        var hasApiEmail = !!student.profile.emailAddress;

        // Check if already has override
        if (overrides[userId]) {
          return; // Already resolved
        }

        // Check if auto-match works
        var email = lookupStudentEmail_(classroomName, userId);
        if (!email && !hasApiEmail) {
          unmatched.push({
            userId: userId,
            classroomName: classroomName,
            period: period
          });
        }
      });
    } catch (e) {
      console.error('Error checking classroom ' + courseId + ': ' + e.message);
    }
  });

  return unmatched;
}

/**
 * Get all students from registry (for dropdown selection)
 */
function getRegistryStudentsForDropdown() {
  checkAuth_();
  var registry = getStudentRegistry_();

  return registry.all.map(function(s) {
    return {
      email: s.email,
      name: s.name,
      periods: s.periods.join(', ')
    };
  }).sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
}

/**
 * Get current overrides for display
 */
function getCurrentOverrides() {
  checkAuth_();
  var overrides = getStudentOverrides_();

  return Object.keys(overrides).map(function(userId) {
    var o = overrides[userId];
    return {
      userId: userId,
      email: o.email,
      classroomName: o.classroomName,
      rosterName: o.rosterName,
      note: o.note
    };
  });
}

/**
 * API endpoint to save an override from the frontend
 */
function saveStudentOverride(data) {
  checkAuth_();
  return setStudentOverride(data.userId, data.email, data.classroomName, data.rosterName, data.note);
}

/**
 * Save multiple overrides at once
 */
function saveStudentOverridesBatch(overrides) {
  checkAuth_();
  var results = { saved: 0, errors: [] };

  overrides.forEach(function(o) {
    try {
      if (o.email && o.userId) {
        setStudentOverride(o.userId, o.email, o.classroomName, o.rosterName, o.note || '');
        results.saved++;
      }
    } catch (e) {
      results.errors.push(o.classroomName + ': ' + e.message);
    }
  });

  return results;
}

/**
 * API endpoint to delete an override
 */
function deleteStudentOverride(userId) {
  checkAuth_();
  return removeStudentOverride(userId);
}

/**
 * Mark a student as "ignored" (dropped from class, doesn't exist, etc.)
 */
function ignoreUnmatchedStudent(userId, classroomName, reason) {
  checkAuth_();
  return setStudentOverride(userId, 'IGNORED', classroomName, '', reason || 'Dropped/Does not exist');
}


// ═══════════════════════════════════════════════════════════════════════════════
// CORE FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Pulls student data from the district roster and builds a local registry.
 */
function syncStudentRegistry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  console.log('═══════════════════════════════════════════════════════════');
  console.log('SYNCING STUDENT REGISTRY FROM DISTRICT ROSTER');
  console.log('═══════════════════════════════════════════════════════════');

  try {
    // 1. Open the external roster
    var rosterSS = SpreadsheetApp.openById(ROSTER_CONFIG.ROSTER_SPREADSHEET_ID);
    var rosterSheet = rosterSS.getSheetByName(ROSTER_CONFIG.ROSTER_SHEET_NAME);

    if (!rosterSheet) {
      throw new Error('Could not find sheet "' + ROSTER_CONFIG.ROSTER_SHEET_NAME + '" in roster spreadsheet');
    }

    // 2. Get all roster data (skip header row)
    var lastRow = rosterSheet.getLastRow();
    if (lastRow < 2) {
      console.log('Roster is empty.');
      return { total: 0, myStudents: 0 };
    }

    var rosterData = rosterSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    console.log('Loaded ' + rosterData.length + ' students from district roster.');

    // 3. Filter to just MY students (where I'm the teacher for periods 3-7)
    var myStudents = [];
    var teacherEmail = ROSTER_CONFIG.TEACHER_EMAIL.toLowerCase();

    rosterData.forEach(function(row) {
      var studentName = row[ROSTER_CONFIG.COLUMNS.STUDENT_NAME];
      var studentId = row[ROSTER_CONFIG.COLUMNS.STUDENT_ID];

      if (!studentName || !studentId) return;

      // Check each of my periods
      var studentPeriods = [];

      ROSTER_CONFIG.MY_PERIODS.forEach(function(period) {
        var colIndex = ROSTER_CONFIG.COLUMNS['PERIOD_' + period];
        var teacherInPeriod = row[colIndex];

        if (teacherInPeriod && teacherInPeriod.toString().toLowerCase().indexOf(teacherEmail) !== -1) {
          studentPeriods.push(period);
        }
      });

      // If student has me for at least one period, add them
      if (studentPeriods.length > 0) {
        var idStr = studentId.toString().trim();
        var email = idStr + ROSTER_CONFIG.EMAIL_DOMAIN;

        myStudents.push({
          name: studentName.toString().trim(),
          studentId: idStr,
          email: email.toLowerCase(),
          periods: studentPeriods,
          primaryPeriod: studentPeriods[0]
        });
      }
    });

    console.log('Found ' + myStudents.length + ' students in my classes.');

    // 4. Write to local cache sheet
    writeStudentRegistryCache_(ss, myStudents);

    // 5. Update the in-memory cache
    STUDENT_REGISTRY_CACHE = buildRegistryLookup_(myStudents);

    console.log('');
    console.log('✅ Registry sync complete!');
    console.log('   Total students: ' + myStudents.length);

    ss.toast('Synced ' + myStudents.length + ' students from district roster.', '✅ Registry Updated', 5);

    return {
      total: rosterData.length,
      myStudents: myStudents.length
    };

  } catch (error) {
    console.error('❌ Error syncing registry: ' + error.message);
    ss.toast('Error: ' + error.message, '❌ Registry Sync Failed', 10);
    throw error;
  }
}


/**
 * Writes the student registry to a hidden cache sheet
 */
function writeStudentRegistryCache_(ss, students) {
  var sheet = ss.getSheetByName(ROSTER_CONFIG.CACHE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(ROSTER_CONFIG.CACHE_SHEET_NAME);
    sheet.hideSheet();
  }

  sheet.clear();
  sheet.appendRow(['Email', 'Student ID', 'Name', 'Periods', 'Primary Period', 'Last Updated']);

  if (students.length > 0) {
    var rows = students.map(function(s) {
      return [
        s.email,
        s.studentId,
        s.name,
        s.periods.join(', '),
        s.primaryPeriod,
        new Date()
      ];
    });

    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#E8E8E8');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(3, 200);

  console.log('Wrote ' + students.length + ' students to cache sheet.');
}


// ═══════════════════════════════════════════════════════════════════════════════
// LOOKUP FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

var STUDENT_REGISTRY_CACHE = null;

/**
 * Builds lookup dictionaries from student array
 */
function buildRegistryLookup_(students) {
  var byEmail = {};
  var byStudentId = {};
  var byName = {};

  students.forEach(function(s) {
    byEmail[s.email.toLowerCase()] = s;
    byStudentId[s.studentId] = s;

    var normName = s.name.toLowerCase().replace(/[^a-z]/g, '');
    byName[normName] = s;
  });

  return {
    byEmail: byEmail,
    byStudentId: byStudentId,
    byName: byName,
    all: students
  };
}


/**
 * Gets the student registry, loading from cache if needed
 */
function getStudentRegistry_() {
  if (STUDENT_REGISTRY_CACHE) {
    return STUDENT_REGISTRY_CACHE;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROSTER_CONFIG.CACHE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    console.log('Student registry cache is empty. Run syncStudentRegistry() first.');
    return { byEmail: {}, byStudentId: {}, byName: {}, all: [] };
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  var students = data.map(function(row) {
    return {
      email: row[0],
      studentId: row[1].toString(),
      name: row[2],
      periods: row[3].toString().split(', ').map(Number),
      primaryPeriod: row[4]
    };
  });

  STUDENT_REGISTRY_CACHE = buildRegistryLookup_(students);
  return STUDENT_REGISTRY_CACHE;
}


/**
 * Look up a student's email by their name
 */
/**
 * REPLACE the lookupStudentEmail_ function in StudentRegistry.gs with this:
 */
function lookupStudentEmail_(studentName, classroomUserId) {
  // CHECK OVERRIDES FIRST
  if (classroomUserId) {
    var overrides = getStudentOverrides_();
    if (overrides[classroomUserId]) {
      var override = overrides[classroomUserId];
      if (override.email === 'IGNORED') {
        debugLog_("✓ Found override for student ID " + classroomUserId + ": IGNORED");
        return null;
      }
      debugLog_("✓ Found override for student ID " + classroomUserId + ": " + override.email);
      return override.email;
    }
  }

  var registry = getStudentRegistry_();

  if (!studentName) {
    debugLog_("⚠️  Cannot lookup email - no student name provided (ID: " + classroomUserId + ")");
    return null;
  }

  debugLog_("Looking up email for student: '" + studentName + "' (ID: " + classroomUserId + ")");
  
  // Normalize function - strips apostrophes, hyphens, accents, etc.
  var normalize = function(str) {
    return str.toLowerCase()
      .replace(/[''`´']/g, '')      // apostrophes
      .replace(/[\-–—]/g, '')       // hyphens/dashes
      .replace(/[\.]/g, '')         // periods
      .replace(/[^a-z]/g, '');      // anything else non-alpha
  };
  
  var classroomName = studentName.trim();
  var classroomParts = classroomName.split(/\s+/);
  
  var classroomFirst = normalize(classroomParts[0]);
  var classroomLast = normalize(classroomParts[classroomParts.length - 1]);
  
  // Search registry
  for (var i = 0; i < registry.all.length; i++) {
    var student = registry.all[i];
    var rosterName = student.name;
    
    var commaIndex = rosterName.indexOf(',');
    if (commaIndex === -1) continue;
    
    var rosterLast = normalize(rosterName.substring(0, commaIndex).trim());
    var rosterRest = rosterName.substring(commaIndex + 1).trim();
    var rosterFirst = normalize(rosterRest.split(/\s+/)[0]);
    
    // Match: first names match AND last names match
    var firstMatch = classroomFirst === rosterFirst || 
                     classroomFirst.indexOf(rosterFirst) === 0 || 
                     rosterFirst.indexOf(classroomFirst) === 0;
    
    var lastMatch = classroomLast === rosterLast ||
                    classroomLast.indexOf(rosterLast) === 0 ||
                    rosterLast.indexOf(classroomLast) === 0;
    
    if (firstMatch && lastMatch) {
      return student.email;
    }
  }
  
  // Fuzzy match pass
  for (var i = 0; i < registry.all.length; i++) {
    var student = registry.all[i];
    var rosterName = student.name;
    
    var commaIndex = rosterName.indexOf(',');
    if (commaIndex === -1) continue;
    
    var rosterLast = normalize(rosterName.substring(0, commaIndex).trim());
    var rosterRest = rosterName.substring(commaIndex + 1).trim();
    var rosterFirst = normalize(rosterRest.split(/\s+/)[0]);
    
    if (classroomFirst === rosterFirst) {
      if (rosterLast.indexOf(classroomLast) !== -1 || classroomLast.indexOf(rosterLast) !== -1) {
        return student.email;
      }
    }
  }
  
  debugLog_('❌ Could not find email match in registry for student: "' + studentName + '" (ID: ' + classroomUserId + ')');
  debugLog_('   Normalized name parts: first="' + classroomFirst + '", last="' + classroomLast + '"');
  debugLog_('   💡 TIP: Check Config tab to manually link this student or verify name format in registry');
  return null;
}

/**
 * Look up student info by email
 */
function lookupStudentByEmail_(email) {
  var registry = getStudentRegistry_();
  return registry.byEmail[email.toLowerCase()] || null;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ENHANCED getAllStudents_ WRAPPER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Wraps Classroom API and fills in missing emails from the registry
 */
function getAllStudentsEnriched_(courseId) {
  var students = [];
  var pageToken = null;

  do {
    var response = Classroom.Courses.Students.list(courseId, {
      pageSize: 100,
      pageToken: pageToken
    });

    if (response.students) {
      students = students.concat(response.students);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  // Enrich with registry data
  var registry = getStudentRegistry_();
  var enrichedCount = 0;

  students.forEach(function(student) {
    if (!student.profile.emailAddress && student.profile.name) {
      var fullName = student.profile.name.fullName;
      var email = lookupStudentEmail_(fullName, student.userId);

      if (email) {
        student.profile.emailAddress = email;
        enrichedCount++;
      }
    }
  });

  if (enrichedCount > 0) {
    console.log('📧 Enriched ' + enrichedCount + ' students with emails from registry.');
  }

  return students;
}


// ═══════════════════════════════════════════════════════════════════════════════
// DIAGNOSTIC FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test the registry lookup
 */
function testRegistryLookup() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('TESTING STUDENT REGISTRY LOOKUP');
  console.log('═══════════════════════════════════════════════════════════');

  var registry = getStudentRegistry_();

  console.log('Registry loaded with ' + registry.all.length + ' students.');
  console.log('');

  console.log('Sample students:');
  registry.all.slice(0, 5).forEach(function(s, i) {
    console.log((i+1) + '. ' + s.name);
    console.log('   Email: ' + s.email);
    console.log('   ID: ' + s.studentId);
    console.log('   Periods: ' + s.periods.join(', '));
  });
}


/**
 * Test enrichment with a specific classroom
 */
function testEnrichedStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);

  if (!classroomsSheet || classroomsSheet.getLastRow() < 3) {
    console.log('No classrooms configured.');
    return;
  }

  var classroomData = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 6).getValues();
  var activeClassroom = classroomData.find(function(row) {
    return row[4] === 'TRUE' || row[4] === true;
  });

  if (!activeClassroom) {
    console.log('No active classrooms.');
    return;
  }

  var courseId = activeClassroom[0];
  var courseName = activeClassroom[1];

  console.log('═══════════════════════════════════════════════════════════');
  console.log('TESTING ENRICHED STUDENT DATA');
  console.log('Classroom: ' + courseName);
  console.log('═══════════════════════════════════════════════════════════');

  var students = getAllStudentsEnriched_(courseId);

  var withEmail = 0;
  var withoutEmail = 0;

  students.forEach(function(student) {
    var name = student.profile.name ? student.profile.name.fullName : 'Unknown';
    var email = student.profile.emailAddress;

    if (email) {
      withEmail++;
      console.log('✅ ' + name + ' → ' + email);
    } else {
      withoutEmail++;
      console.log('❌ ' + name + ' → (no email found)');
    }
  });

  console.log('');
  console.log('RESULTS:');
  console.log('✅ With email: ' + withEmail);
  console.log('❌ Without email: ' + withoutEmail);
}


// ═══════════════════════════════════════════════════════════════════════════════
// CLEANUP FUNCTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Removes entries with blank emails from the grades database
 */
function cleanBlankEmailsFromGradesDB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("hidden_grades_db");

  console.log('═══════════════════════════════════════════════════════════');
  console.log('CLEANING BLANK EMAILS FROM GRADES DATABASE');
  console.log('═══════════════════════════════════════════════════════════');

  if (!sheet) {
    console.log('❌ hidden_grades_db sheet not found.');
    return;
  }

  if (sheet.getLastRow() < 2) {
    console.log('✅ No data to clean (sheet is empty).');
    return;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  console.log('Found ' + data.length + ' total entries.');

  var cleanedData = [];
  var removedCount = 0;

  data.forEach(function(row) {
    var email = row[0];
    var name = row[2];

    if (email && email.toString().trim() !== '') {
      cleanedData.push(row);
    } else {
      removedCount++;
      console.log('🗑️  Removing: ' + name);
    }
  });

  if (removedCount === 0) {
    console.log('✅ No blank email entries found! Database is clean.');
    return;
  }

  sheet.getRange(2, 1, sheet.getLastRow(), 9).clearContent();

  if (cleanedData.length > 0) {
    sheet.getRange(2, 1, cleanedData.length, 9).setValues(cleanedData);
  }

  console.log('');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🗑️  Removed: ' + removedCount + ' invalid entries');
  console.log('✅ Remaining: ' + cleanedData.length + ' valid entries');

  // Recompile progress
  console.log('Recompiling Progress sheet...');
  compileProgressFromDB_();
  console.log('✅ Progress sheet updated.');

  ss.toast('Removed ' + removedCount + ' invalid entries.', '🧹 Cleanup Complete', 10);
}
