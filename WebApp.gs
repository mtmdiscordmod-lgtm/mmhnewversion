/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * WEB APP BACKEND (Teacher Console)
 * ═══════════════════════════════════════════════════════════════════════════════
 */
function debugSongManager() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classroomsSheet = ss.getSheetByName("🏫 Classrooms");
  var songSheet = ss.getSheetByName("🎵 Song Requests");
  
  // 1. Check Classroom Config for Period 3 ID
  var targetId = "795768825195"; // The ID from your text
  console.log("--- LOOKING FOR CLASSROOM ID: " + targetId + " ---");
  
  var cData = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 6).getValues();
  var foundPeriod = null;
  
  for (var i = 0; i < cData.length; i++) {
    var rawId = cData[i][0];
    var idString = String(rawId).trim();
    
    // Log any near-matches or duplicates
    if (idString.indexOf("795768") !== -1) {
      console.log("Found Row " + (i+3) + ": ID='" + rawId + "' (Type: " + typeof rawId + ") | Period='" + cData[i][3] + "'");
      if (idString == targetId) {
        foundPeriod = cData[i][3]; // Store the period found
        // Don't break yet, let's see if there are duplicates!
      }
    }
  }
  
  if (foundPeriod === null) {
    console.error("❌ CRITICAL: Could not find the Class ID in the sheet!");
    return;
  }
  
  var targetPeriodNorm = String(foundPeriod).trim().replace(/^Period\s*/i, '');
  console.log("✅ MATCH TARGET: Sheet Period is '" + foundPeriod + "' -> Normalized to '" + targetPeriodNorm + "'");

  // 2. Check Song Requests
  console.log("--- CHECKING SONG REQUESTS ---");
  var sData = songSheet.getRange(3, 1, songSheet.getLastRow() - 2, 8).getValues();
  
  var matchCount = 0;
  for (var j = 0; j < sData.length; j++) {
    var row = sData[j];
    var rowPeriod = row[3];
    var rowStatus = row[7];
    var rowNorm = String(rowPeriod).trim().replace(/^Period\s*/i, '');
    
    // Log the Period 3 request specifically
    if (String(rowPeriod) == "3" || String(rowPeriod) == "Period 3") {
      var isMatch = (rowNorm === targetPeriodNorm);
      var isStatus = (rowStatus === 'PENDING' || rowStatus === 'STAGED');
      console.log("Song Row " + (j+3) + ": Period='" + rowPeriod + "' (Norm: " + rowNorm + ") | Status='" + rowStatus + "'");
      console.log("   -> Matches Period? " + isMatch);
      console.log("   -> Matches Status? " + isStatus);
      if (isMatch && isStatus) matchCount++;
    }
  }
  
  console.log("--- RESULT ---");
  console.log("Total visible songs for this class: " + matchCount);
}
var AUTHORIZED_USER = 'jberman2@sandi.net';

/**
 * Serves the HTML interface.
 */
function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail();

  // ROUTING LOGIC:
  // 1. Explicit 'page=request' parameter
  // 2. Student Domain Detection (@stu.sandi.net)
  var isStudent = userEmail.indexOf('@stu.sandi.net') !== -1;
  var isRequestPage = e.parameter.page === 'request';
  var isBonusPage = e.parameter.page === 'bonus';

  if (isBonusPage) {
    return HtmlService.createTemplateFromFile('bonus_client')
      .evaluate()
      .setTitle('MMH Bonus Round')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  } else if (isStudent || isRequestPage) {
    return HtmlService.createTemplateFromFile('studentrequest')
      .evaluate()
      .setTitle('MMH Student Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('MMH Teacher Console')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * HELPER: SECURITY CHECK
 * Throws an error if the user is not authorized.
 */
function checkAuth_() {
  var email = Session.getActiveUser().getEmail();
  if (!email || email.trim().toLowerCase() !== AUTHORIZED_USER.toLowerCase()) {
    throw new Error('⛔ UNAUTHORIZED ACCESS: ' + email);
  }
}

/**
 * 1. DASHBOARD INIT
 * Returns user email, active classrooms, and script URL.
 */
function getDataForDashboard() {
  var email = Session.getActiveUser().getEmail();
  var normalizedEmail = email ? email.trim().toLowerCase() : '';
  var authorizedEmail = AUTHORIZED_USER.toLowerCase();

  var scriptUrl = '';
  try {
    scriptUrl = ScriptApp.getService().getUrl();
  } catch (e) {
    console.log('Script URL not available (not deployed?)');
  }

  if (normalizedEmail !== authorizedEmail) {
    return {
      email: email,
      authorized: false,
      courses: [],
      scriptUrl: scriptUrl
    };
  }

  // Read courses directly from sheet (Read-Only) to prevent overwriting config
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var courseList = [];

  if (sheet && sheet.getLastRow() >= 3) {
    // Col A=ID, B=Name, C=Section, D=Period, E=Active
    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 5).getValues();
    data.forEach(function(row) {
      // Check ID (row[0]) and Active Status (row[4])
      var isActive = (row[4] === true || String(row[4]).toUpperCase() === 'TRUE');
      if (row[0] && isActive) {
        courseList.push({
          id: row[0],
          name: row[1],
          section: row[2] || ''
        });
      }
    });
  }

  return {
    email: email,
    authorized: true,
    courses: courseList,
    scriptUrl: scriptUrl
  };
}

/**
 * HELPER: DEBUG LOGGING
 */
function debugLog_(message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('🐛 Debug Log');
    if (!sheet) {
      sheet = ss.insertSheet('🐛 Debug Log');
      sheet.appendRow(['Timestamp', 'Message']);
    }
    sheet.appendRow([new Date(), message]);
  } catch (e) {
    console.log('Debug Log Error: ' + e.message);
  }
}

/**
 * HELPER: GET MANAGED ASSIGNMENT IDS
 */
function getManagedAssignmentIds_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (typeof CONFIG !== 'undefined') ? CONFIG.SHEETS.ASSIGNMENTS : '📋 Assignments';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  var data = sheet.getRange(3, 4, lastRow - 2, 1).getValues();

  var ids = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) ids.push(String(data[i][0]).trim());
  }
  return ids;
}

/**
 * HELPER: BATCH FETCH PENDING SUBMISSIONS
 */
function fetchPendingSubmissions_(courseId) {
  debugLog_("--- FETCHING SUBMISSIONS for " + courseId + " ---");
  var pendingItems = [];
  var managedIds = getManagedAssignmentIds_();
  debugLog_("Managed IDs Count: " + managedIds.length);
  if (managedIds.length > 0) debugLog_("First Managed ID: '" + managedIds[0] + "'");

  var coursework = [];
  var pageToken = null;
  do {
    var response = Classroom.Courses.CourseWork.list(courseId, {
      pageSize: 100,
      pageToken: pageToken,
      courseWorkStates: ['PUBLISHED']
    });
    if (response.courseWork) coursework = coursework.concat(response.courseWork);
    pageToken = response.nextPageToken;
  } while (pageToken);

  debugLog_("API Coursework Found: " + coursework.length);
  if (coursework.length === 0) return [];
  debugLog_("First API ID: '" + coursework[0].id + "'");

  var students = getCachedStudents_(courseId);
  var studentMap = {};
  students.forEach(function(s) { studentMap[s.userId] = s.profile.name.fullName; });

  var matchedAssignments = 0;
  var totalSubmissionsFound = 0;

  coursework.forEach(function(work) {
    // Robust comparison: Ensure both are trimmed strings
    if (managedIds.indexOf(String(work.id).trim()) === -1) return;
    matchedAssignments++;

    try {
      var subResponse = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, work.id, {
        states: ['TURNED_IN']
      });

      if (subResponse.studentSubmissions) {
        totalSubmissionsFound += subResponse.studentSubmissions.length;
        subResponse.studentSubmissions.forEach(function(sub) {
          var studentName = studentMap[sub.userId] || 'Unknown Student';

          var files = [];
          if (sub.assignmentSubmission && sub.assignmentSubmission.attachments) {
            sub.assignmentSubmission.attachments.forEach(function(att) {
              if (att.driveFile) {
                var fileDate = new Date(0);
                try {
                  fileDate = DriveApp.getFileById(att.driveFile.id).getDateCreated();
                } catch(e) {}

                files.push({
                  id: att.driveFile.id,
                  title: att.driveFile.title || 'Untitled File',
                  url: att.driveFile.alternateLink,
                  created: fileDate.toISOString() // Explicit string conversion
                });
              }
            });
          }

          // Sort using string comparison or Date parsing if needed, but ISO string comparison works for date ordering
          files.sort(function(a, b) { return (b.created > a.created) ? 1 : -1; });

          pendingItems.push({
            assignmentId: work.id,
            assignmentTitle: work.title,
            maxPoints: work.maxPoints || 0,
            submissionId: sub.id,
            userId: sub.userId,
            studentName: studentName,
            files: files,
            currentGrade: sub.assignedGrade || sub.draftGrade || '',
            state: sub.state
          });
        });
      }
    } catch (e) {
      console.log('Error fetching subs for ' + work.id + ': ' + e.message);
      debugLog_("Error fetching subs for " + work.id + ": " + e.message);
    }
  });

  debugLog_("Matched Assignments: " + matchedAssignments);
  debugLog_("Total Submissions Found: " + totalSubmissionsFound);
  debugLog_("Returning " + pendingItems.length + " items.");

  return pendingItems;
}

/**
 * 3. GET GRADING DATA (BATCH)
 */
function getGradingData(courseId) {
  try {
    checkAuth_();
    var pendingList = fetchPendingSubmissions_(courseId);

    if (!pendingList || !Array.isArray(pendingList)) {
      debugLog_("getGradingData Warning: pendingList was " + typeof pendingList);
      return [];
    }

    pendingList.sort(function(a, b) {
      if (a.assignmentTitle !== b.assignmentTitle) {
        return a.assignmentTitle.localeCompare(b.assignmentTitle);
      }
      return a.studentName.localeCompare(b.studentName);
    });

    debugLog_("getGradingData Returning: " + pendingList.length + " items");
    return pendingList;

  } catch (e) {
    debugLog_("getGradingData ERROR: " + e.message);
    throw e;
  }
}

/**
 * 4. GRADE & RETURN & CLEANUP
 * Updates grade, returns submission, and deletes the shortcut/folder from the export to keep it clean.
 */
function gradeAndReturnSubmission(courseId, assignmentId, studentId, grade, submissionId, comments) {
  checkAuth_();

  debugLog_('=== GRADE ATTEMPT ===');
  debugLog_('courseId: ' + courseId);
  debugLog_('assignmentId: ' + assignmentId);
  debugLog_('studentId: ' + studentId);
  debugLog_('grade: ' + grade);
  debugLog_('submissionId: ' + submissionId);
  debugLog_('comments: ' + comments);

  try {
    var numericGrade = Number(grade);
    if (isNaN(numericGrade)) throw new Error("Invalid grade format: " + grade);

    // Robust ID handling
    var cId = String(courseId).trim();
    var aId = String(assignmentId).trim();
    var sId = String(studentId).trim();
    var subId = String(submissionId).trim();

    debugLog_('Trimmed IDs - cId: ' + cId + ', aId: ' + aId + ', subId: ' + subId);

    // --- 1. FETCH DETAILS FOR DOC GEN ---
    var maxPoints = 0;
    var assignmentTitle = "Assignment";
    var studentName = "Student";
    var studentEmail = "";
    var coursePeriod = "Unknown Period";
    var turnedInDate = new Date();

    // A. Assignment Details
    try {
      var courseWork = Classroom.Courses.CourseWork.get(cId, aId);
      maxPoints = courseWork.maxPoints || 0;
      assignmentTitle = courseWork.title;
    } catch(e) { console.log("Error getting coursework details: " + e.message); }

    // B. Student Details
    try {
      var profile = Classroom.Users.get(sId);
      studentName = profile.name.fullName;
      studentEmail = profile.emailAddress;
    } catch(e) {
      debugLog_("❌ Error getting user profile from Classroom API: " + e.message);
      debugLog_("⚠️  Attempting to use student ID as fallback for folder name");
      // Use student ID as fallback to avoid all students being in "Student" folder
      studentName = "Student_" + sId;
    }

    // If name is still default (API failed), try to get it from registry
    if (studentName === "Student" || studentName.indexOf("Student_") === 0) {
      debugLog_("⚠️  Student name not resolved from Classroom API, checking registry...");
      try {
        // Try to get name from registry using student ID
        var registryResult = lookupStudentEmail_(null, sId); // Pass null for name, just use ID
        if (registryResult && registryResult !== 'IGNORED' && registryResult !== '') {
          studentEmail = registryResult;
          // Try to extract name from email (before @)
          var emailParts = registryResult.split('@');
          if (emailParts.length > 0 && emailParts[0]) {
            // Use email prefix as name fallback (better than "Student")
            studentName = emailParts[0].replace(/[0-9]/g, '').replace(/[\._]/g, ' ').trim() || ("Student_" + sId);
            debugLog_("✓ Derived student name from registry email: " + studentName);
          }
        }
      } catch (regErr) {
        debugLog_("Registry lookup for name failed: " + regErr.message);
      }
    }

    // Fallback if API hidden email
    if (!studentEmail || studentEmail === '') {
      debugLog_("⚠️  Email not provided by Classroom API, trying fallbacks...");

      try {
        // Fallback 1: Registry lookup (uses StudentRegistry.gs)
        var registryEmail = lookupStudentEmail_(studentName, sId);

        if (registryEmail && registryEmail !== 'IGNORED' && registryEmail !== '') {
           studentEmail = registryEmail;
           debugLog_("✓ Registry email lookup SUCCESS for " + sId + ": " + studentEmail);
        } else {
           debugLog_("✗ Registry lookup failed or returned IGNORED for " + sId);

           // Fallback 2: Local cache lookup (uses cached grades/progress sheets)
           var cacheEmail = findStudentEmailById_(sId);

           if (cacheEmail && cacheEmail !== '') {
             studentEmail = cacheEmail;
             debugLog_("✓ Local cache lookup SUCCESS for " + sId + ": " + studentEmail);
           } else {
             debugLog_("✗ Local cache lookup also FAILED for " + sId);
             debugLog_("⚠️  WARNING: No email found for student '" + studentName + "' (ID: " + sId + ")");
             debugLog_("⚠️  Doc will be created but student won't be added as commenter");
             // Don't set studentEmail - leave empty, doc will still be created (Fix #1)
           }
        }
      } catch (e) {
        debugLog_("❌ ERROR during email fallback lookup: " + e.message);
        // Continue anyway - doc creation should not depend on email (Fix #1)
      }
    }

    // Log final email resolution status
    if (studentEmail && studentEmail !== 'IGNORED') {
      debugLog_("✓ Final student email resolved: " + studentEmail);
    } else {
      debugLog_("⚠️  Proceeding WITHOUT student email (doc will be created but commenter step will be skipped)");
    }

    // C. Period Details (from cached config or list)
    try {
      var course = Classroom.Courses.get(cId);
      // Try to find period from config cache logic or simple parsing if enabled,
      // but let's rely on the Config Sheet if possible or just use Course Name/Section
      // For simplicity/robustness, we check the sheet
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
      if (configSheet) {
         var data = configSheet.getRange(3, 1, configSheet.getLastRow()-2, 6).getValues();
         for(var i=0; i<data.length; i++) {
           if(String(data[i][0]).trim() === cId) {
             coursePeriod = data[i][3] ? "Period " + data[i][3] : coursePeriod;
             break;
           }
         }
      }
    } catch(e) { console.log("Error resolving period: " + e.message); }

    // D. Fetch Submission Timestamp (Turned In Date)
    try {
      var existingSub = Classroom.Courses.CourseWork.StudentSubmissions.get(cId, aId, subId);
      if (existingSub.updateTime) {
         turnedInDate = existingSub.updateTime;
      }
    } catch (e) { debugLog_("Could not fetch submission time: " + e.message); }

    // E. Calculate Coins
    var newBalance = 0;
    try {
      if (studentEmail) {
        var currentBalance = calculateTimmyCoinBalance_(studentEmail);
        newBalance = currentBalance + numericGrade;
      }
    } catch (e) { console.log("Balance calc error: " + e.message); }


    // --- 2. GENERATE GRADING LOG DOC ---
    // Create doc even without email (only requires student name)
    var newDocId = null;
    if (studentName) {  // Only require student name, not email
      try {
        debugLog_("Creating Grading Log for " + studentName);
        var folder = getOrCreateGradingFolder_(coursePeriod, studentName);
        // Pass new arguments: turnedInDate, newBalance
        var docFile = createGradingDoc_(folder, studentName, assignmentTitle, numericGrade, maxPoints, comments, turnedInDate, newBalance);
        newDocId = docFile.getId();

        // Add Student as Commenter (only if we have valid email)
        if (studentEmail && studentEmail !== 'IGNORED' && studentEmail !== '') {
          try {
            docFile.addCommenter(studentEmail);
            debugLog_("✓ Student added as commenter: " + studentEmail);
          } catch (commentErr) {
            debugLog_("⚠️  Could not add commenter: " + commentErr.message);
            // Continue anyway - doc is still created and will be attached
          }
        } else {
          debugLog_("⚠️  Skipping commenter step (no valid email for student)");
        }

        debugLog_("Created Doc: " + docFile.getUrl());
      } catch (docErr) {
        debugLog_("❌ Error generating doc: " + docErr.message);
        throw new Error("Failed to create grading doc: " + docErr.message);
      }
    } else {
      debugLog_("❌ CRITICAL: Cannot create doc - no student name available");
      throw new Error("Cannot create grading doc without student name");
    }

    // --- 3. ATTACH DOC TO SUBMISSION (The Tricky Part) ---
    if (newDocId) {
      try {
        debugLog_("Attaching Doc to submission...");
        Classroom.Courses.CourseWork.StudentSubmissions.modifyAttachments({
          addAttachments: [{
            driveFile: { id: newDocId }
          }]
        }, cId, aId, subId);
        debugLog_("Attachment successful.");
      } catch (attachErr) {
        debugLog_("Could not attach doc (likely not created by script): " + attachErr.message);
        // Continue anyway - the doc exists in Drive and is shared
      }
    }


    // TEST 1: Can we even READ the submission?
    try {
      var existingSub = Classroom.Courses.CourseWork.StudentSubmissions.get(cId, aId, subId);
      debugLog_('SUCCESS: Can read submission. Current state: ' + existingSub.state);
      debugLog_('Current draftGrade: ' + existingSub.draftGrade);
      debugLog_('Current assignedGrade: ' + existingSub.assignedGrade);
    } catch (readErr) {
      debugLog_('FAILED to read submission: ' + readErr.message);
      return { success: false, error: 'Cannot read submission: ' + readErr.message };
    }

    // TEST 2: Try the patch - MUST set BOTH draftGrade AND assignedGrade
    // CRITICAL: The API does NOT auto-copy draftGrade to assignedGrade like the Classroom UI does!
    // See: https://developers.google.com/workspace/classroom/reference/rest/v1/courses.courseWork.studentSubmissions/return
    try {
      debugLog_('Attempting patch with draftGrade AND assignedGrade: ' + numericGrade);
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        { draftGrade: numericGrade, assignedGrade: numericGrade },
        cId, aId, subId,
        { updateMask: 'draftGrade,assignedGrade' }
      );
      debugLog_('SUCCESS: Patch completed (both draftGrade and assignedGrade set)');
    } catch (patchErr) {
      debugLog_('FAILED patch: ' + patchErr.message);
      return { success: false, error: 'Patch failed: ' + patchErr.message };
    }

    // TEST 3: Try the return
    try {
      debugLog_('Attempting return...');
      Classroom.Courses.CourseWork.StudentSubmissions['return']({}, cId, aId, subId);
      debugLog_('SUCCESS: Return completed');
    } catch (returnErr) {
      debugLog_('FAILED return: ' + returnErr.message);
      return { success: false, error: 'Return failed: ' + returnErr.message };
    }

    // 3. Cleanup ("Inbox Zero") - This still needs Student ID to find the folder by Name
    cleanupGradedShortcuts_(cId, aId, sId);

    // Return detailed success information
    var docUrl = null;
    if (newDocId) {
      try {
        docUrl = DriveApp.getFileById(newDocId).getUrl();
      } catch (urlErr) {
        debugLog_("Could not retrieve doc URL: " + urlErr.message);
      }
    }

    return {
      success: true,
      docId: newDocId,
      docUrl: docUrl,
      docCreated: !!newDocId,
      docAttached: !!newDocId,
      emailResolved: !!(studentEmail && studentEmail !== 'IGNORED'),
      commenterAdded: !!(studentEmail && studentEmail !== 'IGNORED' && newDocId),
      studentName: studentName,
      assignmentTitle: assignmentTitle,
      grade: numericGrade,
      maxPoints: maxPoints
    };

  } catch (e) {
    debugLog_('❌ GRADING ERROR: ' + e.message);
    console.error('Grading Error: ' + e.message);

    // Return detailed error information
    return {
      success: false,
      error: e.message,
      errorStack: e.stack || 'No stack trace available',
      details: {
        courseId: cId,
        assignmentId: aId,
        studentId: sId,
        studentName: studentName || 'Unknown',
        assignmentTitle: assignmentTitle || 'Unknown',
        emailResolved: !!(studentEmail && studentEmail !== 'IGNORED'),
        docCreated: !!newDocId
      }
    };
  }
}

/**
 * HELPER: Get or Create Grading Folder Hierarchy
 * Multimedia Heroes / Grading Logs / Period X / Student Name
 */
function getOrCreateGradingFolder_(period, studentName) {
  try {
    // 1. Root folder
    var rootName = "Multimedia Heroes";
    var rootIter = DriveApp.getFoldersByName(rootName);
    var root;

    if (rootIter.hasNext()) {
      root = rootIter.next();
      debugLog_("✓ Found existing root folder: " + rootName);
    } else {
      root = DriveApp.createFolder(rootName);
      debugLog_("✓ Created new root folder: " + rootName + " (ID: " + root.getId() + ")");
    }

    if (!root) throw new Error("Failed to get/create root folder: " + rootName);

    // 2. Grading Logs subfolder
    var logsName = "Grading Logs";
    var logsIter = root.getFoldersByName(logsName);
    var logsFolder;

    if (logsIter.hasNext()) {
      logsFolder = logsIter.next();
      debugLog_("✓ Found existing Grading Logs folder");
    } else {
      logsFolder = root.createFolder(logsName);
      debugLog_("✓ Created Grading Logs folder (ID: " + logsFolder.getId() + ")");
    }

    if (!logsFolder) throw new Error("Failed to get/create Grading Logs folder");

    // 3. Period subfolder
    var periodName = String(period).indexOf("Period") !== -1 ? String(period) : "Period " + period;
    var pIter = logsFolder.getFoldersByName(periodName);
    var pFolder;

    if (pIter.hasNext()) {
      pFolder = pIter.next();
      debugLog_("✓ Found existing period folder: " + periodName);
    } else {
      pFolder = logsFolder.createFolder(periodName);
      debugLog_("✓ Created period folder: " + periodName + " (ID: " + pFolder.getId() + ")");
    }

    if (!pFolder) throw new Error("Failed to get/create period folder: " + periodName);

    // 4. Student subfolder
    var sIter = pFolder.getFoldersByName(studentName);
    var studentFolder;

    if (sIter.hasNext()) {
      studentFolder = sIter.next();
      debugLog_("✓ Found existing student folder: " + studentName);
    } else {
      studentFolder = pFolder.createFolder(studentName);
      debugLog_("✓ Created student folder: " + studentName + " (ID: " + studentFolder.getId() + ")");
    }

    if (!studentFolder) throw new Error("Failed to get/create student folder for: " + studentName);

    debugLog_("✓ Folder hierarchy complete: " + rootName + "/" + logsName + "/" + periodName + "/" + studentName);
    return studentFolder;

  } catch (e) {
    debugLog_("❌ FOLDER CREATION FAILED: " + e.message);
    debugLog_("   Period: " + period + ", Student: " + studentName);
    throw new Error("Folder creation error: " + e.message);
  }
}

/**
 * HELPER: Generate the Google Doc
 */
function createGradingDoc_(folder, studentName, assignmentTitle, grade, maxPoints, comments, turnedInDate, newBalance) {
  var docName = assignmentTitle + " - Grading Log";
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // Set purple background
  body.setBackgroundColor('#2e1065'); // Deep purple (Violet-950)

  // COLORS
  var COL_CYAN = '#22d3ee';
  var COL_AMBER = '#fbbf24';
  var COL_SLATE = '#94a3b8';
  var COL_DARK = '#020617';
  var COL_PURPLE_TEXT = '#c084fc'; // Purple-400 for better contrast on dark background

  // 0. LOGO
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG.ASSETS && CONFIG.ASSETS.LOGO_URL) {
      var resp = UrlFetchApp.fetch(CONFIG.ASSETS.LOGO_URL);
      var blob = resp.getBlob();
      var imgPara = body.insertParagraph(0, "");
      var img = imgPara.appendInlineImage(blob);
      img.setWidth(100).setHeight(100);
      imgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  } catch(e) { console.log("Logo insertion failed: " + e.message); }

  // 1. HEADER
  var headerPara = body.appendParagraph("MULTIMEDIA HEROES // GRADING LOG");
  headerPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerPara.setAttributes({
    'FOREGROUND_COLOR': COL_CYAN,
    'FONT_FAMILY': 'Consolas',
    'BOLD': true
  });

  var subHeader = body.appendParagraph(assignmentTitle);
  subHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  subHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subHeader.setAttributes({
    'FOREGROUND_COLOR': '#555555',
    'FONT_FAMILY': 'Verdana',
    'BOLD': true
  });

  body.appendHorizontalRule();

  // 2. SCORE SECTION
  var isPerfect = (grade >= maxPoints);
  var scoreText = isPerfect ? "✨ PERFECT SCORE! ✨" : "GRADE REPORT";
  var scoreColor = isPerfect ? COL_AMBER : '#333333';

  var scoreHeader = body.appendParagraph(scoreText);
  scoreHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  scoreHeader.setAttributes({
    'FOREGROUND_COLOR': scoreColor,
    'FONT_SIZE': 18,
    'FONT_FAMILY': 'Consolas',
    'BOLD': true
  });

  // TIMMY COINS SECTION (New)
  body.appendParagraph("");
  var coinText = "💰 EARNINGS: +" + grade + " Coins";
  var coinPara = body.appendParagraph(coinText);
  coinPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  coinPara.setAttributes({
    'FOREGROUND_COLOR': COL_AMBER,
    'FONT_FAMILY': 'Consolas',
    'BOLD': true,
    'FONT_SIZE': 12
  });

  if (newBalance !== undefined && newBalance !== null) {
    var balText = "🏦 NEW BALANCE: " + newBalance + " Coins";
    var balPara = body.appendParagraph(balText);
    balPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    balPara.setAttributes({
      'FOREGROUND_COLOR': COL_AMBER,
      'FONT_FAMILY': 'Consolas',
      'BOLD': true,
      'FONT_SIZE': 12
    });
  }

  if (isPerfect) {
    var congrats = body.appendParagraph("Outstanding work, Agent! You have demonstrated complete mastery of this mission.");
    congrats.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    congrats.setAttributes({ 'ITALIC': true, 'FOREGROUND_COLOR': '#555555' });
  } else {
    var encourage = body.appendParagraph("Mission complete, but there is room for optimization. Check the feedback below to improve your score.");
    encourage.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    encourage.setAttributes({ 'ITALIC': true, 'FOREGROUND_COLOR': '#555555' });
  }

  body.appendParagraph(""); // spacer

  // 3. DETAILS TABLE
  var tDate = turnedInDate ? new Date(turnedInDate).toLocaleDateString() : new Date().toLocaleDateString();
  var cells = [
    ['Date:', new Date().toLocaleDateString()],
    ['Turned In:', tDate],
    ['Student:', studentName],
    ['Score:', grade + " MP / " + maxPoints + " MP"]
  ];

  var table = body.appendTable(cells);
  table.setBorderColor('#dddddd');
  // Style Table
  for(var r=0; r<table.getNumRows(); r++) {
    var row = table.getRow(r);
    row.getCell(0).setAttributes({'BOLD': true, 'FOREGROUND_COLOR': COL_SLATE, 'FONT_FAMILY': 'Consolas'});
    row.getCell(1).setAttributes({'FONT_FAMILY': 'Verdana'});
    if (r === 3 && isPerfect) { // Index 3 is now Score
      row.getCell(1).setAttributes({'FOREGROUND_COLOR': COL_AMBER, 'BOLD': true});
    }
  }

  body.appendParagraph(""); // spacer

  // 4. COMMENTS
  var commentHeader = body.appendParagraph("FEEDBACK / INSTRUCTIONS");
  commentHeader.setAttributes({'FOREGROUND_COLOR': COL_CYAN, 'FONT_FAMILY': 'Consolas', 'BOLD': true});

  var feedbackText = comments && comments.trim() !== "" ? comments : "(No specific comments provided.)";
  var commentPara = body.appendParagraph(feedbackText);
  commentPara.setAttributes({
    'FONT_FAMILY': 'Verdana',
    'FOREGROUND_COLOR': '#000000',
    'FONT_SIZE': 11
  });

  // Footer / Branding
  body.appendParagraph("");
  body.appendHorizontalRule();
  var footer = body.appendParagraph("Multimedia Heroes System // " + new Date().getFullYear());
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.setAttributes({'FOREGROUND_COLOR': '#aaaaaa', 'FONT_SIZE': 8, 'FONT_FAMILY': 'Consolas'});

  doc.saveAndClose();

  // Validate doc was created
  var docId = doc.getId();
  if (!docId) {
    throw new Error("Failed to get doc ID after creation");
  }

  // Get file handle from Drive
  var file = DriveApp.getFileById(docId);
  if (!file) {
    throw new Error("Failed to retrieve doc file from Drive");
  }

  debugLog_("Moving doc " + docId + " to folder: " + folder.getName());

  // Move file to correct folder (removes from root Drive)
  try {
    file.moveTo(folder);
  } catch (moveErr) {
    debugLog_("❌ Failed to move doc to folder: " + moveErr.message);
    throw new Error("Failed to move doc to folder: " + moveErr.message);
  }

  // Validate doc is in correct location
  var parents = file.getParents();
  if (!parents.hasNext()) {
    throw new Error("Doc has no parent folder after move operation!");
  }

  var parentFolder = parents.next();
  debugLog_("✓ Doc successfully moved to folder: " + parentFolder.getName() + " (ID: " + parentFolder.getId() + ")");

  return file;
}

/**
 * HELPER: Deletes the specific student folder in the Export directory to reduce bloat.
 */
function cleanupGradedShortcuts_(courseId, assignmentId, studentId) {
  debugLog_('Starting cleanup for student: ' + studentId);
  try {
    var course = Classroom.Courses.get(courseId);
    var work = Classroom.Courses.CourseWork.get(courseId, assignmentId);
    var studentProfile = Classroom.Users.get(studentId);

    var rootIter = DriveApp.getFoldersByName("MMH Grading Exports");
    if (!rootIter.hasNext()) return;
    var rootFolder = rootIter.next();

    var subFolderName = (course.name + " - " + work.title).replace(/[\/]/g, "-");
    var subIter = rootFolder.getFoldersByName(subFolderName);
    if (!subIter.hasNext()) return;
    var assignFolder = subIter.next();

    var studentName = studentProfile.name.fullName;
    var studentIter = assignFolder.getFoldersByName(studentName);

    if (studentIter.hasNext()) {
      var sFolder = studentIter.next();
      sFolder.setTrashed(true); // Deletes the folder and its shortcuts (does NOT delete original files)
    }
  } catch (e) {
    console.log("Cleanup warning: " + e.message);
  }
}

/**
 * 5. DRIVE ORGANIZER (BATCH + SHORTCUTS)
 */
function organizeDriveFiles(courseId) {
  checkAuth_();
  try {
    var course = Classroom.Courses.get(courseId);
    var className = course.name;
    var pendingItems = fetchPendingSubmissions_(courseId);

    if (pendingItems.length === 0) {
      return { success: true, count: 0, message: "No pending files to organize." };
    }

    var rootName = "MMH Grading Exports";
    var rootIter = DriveApp.getFoldersByName(rootName);
    var rootFolder = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);

    var count = 0;
    var processedAssignments = {};

    pendingItems.forEach(function(item) {
      if (item.files.length > 0) {
        var subFolderName = (className + " - " + item.assignmentTitle).replace(/[\/]/g, "-");

        if (!processedAssignments[subFolderName]) {
           var subIter = rootFolder.getFoldersByName(subFolderName);
           processedAssignments[subFolderName] = subIter.hasNext() ? subIter.next() : rootFolder.createFolder(subFolderName);
        }
        var assignFolder = processedAssignments[subFolderName];

        var studentName = item.studentName;
        var studentFolderIter = assignFolder.getFoldersByName(studentName);
        var studentFolder = studentFolderIter.hasNext() ? studentFolderIter.next() : assignFolder.createFolder(studentName);

        // Deduplication using names
        var existingNames = [];
        var filesIter = studentFolder.getFiles();
        while (filesIter.hasNext()) existingNames.push(filesIter.next().getName());

        item.files.forEach(function(f) {
           try {
             if (existingNames.indexOf(f.title) === -1) {
               studentFolder.createShortcut(f.id);
               existingNames.push(f.title);
               count++;
             }
           } catch (err) { console.log('Error creating shortcut: ' + err.message); }
        });
      }
    });

    return {
      success: true,
      url: rootFolder.getUrl(),
      count: count,
      assignments: Object.keys(processedAssignments).length
    };

  } catch (e) {
    console.error('Drive Org Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 6. LEADERBOARD DATA
 * Fetches and processes Leaderboard stats (Progress Sheet).
 */
function getLeaderboardData() {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Track Map (Code -> Studio)
  var sheetNameTracks = (typeof CONFIG !== 'undefined') ? CONFIG.SHEETS.CONFIG_TRACKS : '⚙️ Tracks';
  var tracksSheet = ss.getSheetByName(sheetNameTracks);
  var studioMap = {}; // { 'SND': 'Sound', 'VIS': 'Visual' }
  if (tracksSheet && tracksSheet.getLastRow() >= 3) {
    var tData = tracksSheet.getRange(3, 1, tracksSheet.getLastRow() - 2, 3).getValues();
    tData.forEach(function(row) {
      if (row[0] && row[2]) studioMap[row[0]] = row[2];
    });
  }

  // 2. Get Progress Data
  var sheetNameProgress = (typeof CONFIG !== 'undefined') ? CONFIG.SHEETS.PROGRESS : '📈 Progress';
  var progressSheet = ss.getSheetByName(sheetNameProgress);
  if (!progressSheet || progressSheet.getLastRow() < 3) return { weekly: [], allTime: [] };

  var pData = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 8).getValues();
  // Col A(0)=Email, B(1)=Name, C(2)=Period, D(3)=TrackCode, G(6)=TotalMP, H(7)=WeekMP

  // 3. Aggregate
  var weeklyMap = {};
  var allTimeMap = {};

  pData.forEach(function(row) {
    var email = row[0];
    var name = row[1];
    var period = row[2];
    var trackCode = row[3];
    var totalMP = Number(row[6]) || 0;
    var weekMP = Number(row[7]) || 0;
    var studio = studioMap[trackCode] || 'General';

    if (!email) return;

    // Aggregate Weekly
    if (!weeklyMap[email]) weeklyMap[email] = { name: name, period: period, mp: 0, studios: {} };
    weeklyMap[email].mp += weekMP;
    if (!weeklyMap[email].studios[studio]) weeklyMap[email].studios[studio] = 0;
    weeklyMap[email].studios[studio] += weekMP;

    // Aggregate All Time
    if (!allTimeMap[email]) allTimeMap[email] = { name: name, period: period, mp: 0, studios: {} };
    allTimeMap[email].mp += totalMP;
    if (!allTimeMap[email].studios[studio]) allTimeMap[email].studios[studio] = 0;
    allTimeMap[email].studios[studio] += totalMP;
  });

  // 4. Sort and Format
  function formatList(map) {
    return Object.values(map)
      .sort(function(a, b) { return b.mp - a.mp; })
      .map(function(item) {
        // Determine dominant studio
        var bestStudio = 'General';
        var maxS = -1;
        Object.keys(item.studios).forEach(function(s) {
          if (item.studios[s] > maxS) { maxS = item.studios[s]; bestStudio = s; }
        });
        return {
          name: item.name,
          period: item.period,
          mp: item.mp,
          studio: bestStudio
        };
      });
  }

  return {
    weekly: formatList(weeklyMap),
    allTime: formatList(allTimeMap)
  };
}

/**
 * 7. AUCTION DATA API
 * Fetches auction items and student roster for fuzzy search.
 */
function getAuctionData() {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. Fetch Items
  var itemSheet = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  var items = [];
  if (itemSheet && itemSheet.getLastRow() >= 3) {
    var itemData = itemSheet.getRange(3, 1, itemSheet.getLastRow() - 2, 4).getValues();
    items = itemData.map(function(row) {
      return {
        name: row[0],
        image: row[1],
        minBid: row[2],
        desc: row[3]
      };
    });
  }

  // B. Fetch Student Roster (from Progress sheet cache)
  var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  var students = [];
  if (progressSheet && progressSheet.getLastRow() >= 3) {
    var pData = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 3).getValues();
    var seen = {};
    pData.forEach(function(row) {
      var email = row[0];
      if (email && !seen[email]) {
        seen[email] = true;
        students.push({
          email: email,
          name: row[1],
          period: row[2]
        });
      }
    });
  }

  // C. Fetch Active Periods
  var classroomSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
  var periods = [];
  if (classroomSheet && classroomSheet.getLastRow() >= 3) {
    // Col 4 = Period (Index 3), Col 5 = Active (Index 4)
    var cData = classroomSheet.getRange(3, 1, classroomSheet.getLastRow() - 2, 6).getValues();
    var uniquePeriods = {};

    cData.forEach(function(row) {
      var period = row[3];
      var isActive = (row[4] === true || row[4] === 'TRUE');

      if (isActive && period !== '' && period !== null) {
        uniquePeriods[period] = true;
      }
    });

    periods = Object.keys(uniquePeriods).sort(function(a, b) {
      // numeric sort attempt
      var numA = parseInt(a.toString().replace(/\D/g, ''));
      var numB = parseInt(b.toString().replace(/\D/g, ''));
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.toString().localeCompare(b.toString());
    });
  }

  return { items: items, students: students, periods: periods };
}

/**
 * 8. CHECK BALANCE API
 */
function checkStudentBalance(email, amount) {
  checkAuth_();

  var balance = calculateTimmyCoinBalance_(email);
  var cost = Number(amount);

  if (isNaN(cost)) return { approved: false, message: "Invalid amount" };

  if (balance >= cost) {
    return { approved: true, balance: balance, message: "Funds Available" };
  } else {
    return { approved: false, balance: balance, message: "Insufficient Funds" };
  }
}

/**
 * 9. FINALIZE AUCTION API
 */
function finalizeAuctionTransaction(email, itemName, amount) {
  checkAuth_();

  var cost = Number(amount);

  // Double check balance server-side to prevent race conditions
  var balance = calculateTimmyCoinBalance_(email);
  if (balance < cost) {
    return { success: false, message: "Insufficient Funds (Balance: " + balance + ")" };
  }

  try {
    addToLedgerDB_(email, -cost, 'Auction', 'Bought: ' + itemName);
    addToInventoryDB_(email, itemName);

    var newBalance = balance - cost;
    return { success: true, newBalance: newBalance };

  } catch (e) {
    console.error("Auction Error: " + e.message);
    return { success: false, message: "System Error" };
  }
}

/**
 * 10. STUDENT PORTAL API
 */

/**
 * Checks student inventory for "Song Request" item
 */
function getStudentProfile() {
  var email = Session.getActiveUser().getEmail();

  // Security: Ensure it's a student or authorized tester
  // (In production, Session.getActiveUser() is reliable for domain users)

  // Check Inventory
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_DB);
  var count = 0;

  if (invSheet && invSheet.getLastRow() >= 2) {
    var data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 3).getValues();
    data.forEach(function(row) {
      if (row[1] === email) {
        if (row[2] === 'Song Request') count++;
        if (row[2] === 'Consumed Song Request') count--;
      }
    });
  }

  return {
    email: email,
    inventory: Math.max(0, count) // Prevent negative
  };
}

/**
 * Handles Song Request Submission
 */
function submitSongRequest(form) {
  var email = Session.getActiveUser().getEmail();

  // 1. Verify Inventory again (Server-side check)
  var profile = getStudentProfile();
  if (profile.inventory < 1) {
    return { success: false, message: "Insufficient Inventory. Please buy a Song Request item." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestSheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);
  if (!requestSheet) return { success: false, message: "System Error: Request DB not found." };

  // 2. Get Student Name (from Progress Cache for speed)
  var name = email;
  var period = "?";
  var progressSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
  if (progressSheet && progressSheet.getLastRow() >= 3) {
    var pData = progressSheet.getRange(3, 1, progressSheet.getLastRow() - 2, 3).getValues();
    for(var i=0; i<pData.length; i++) {
      if(pData[i][0] === email) {
        name = pData[i][1];
        period = pData[i][2];
        break;
      }
    }
  }

  // 3. Deduct Inventory (Log consumption)
  try {
    addToInventoryDB_(email, 'Consumed Song Request');

    // 4. Add Request to Sheet
    // Columns: Timestamp, Email, Name, Period, Artist, Song, Link, Status
    requestSheet.appendRow([
      new Date(),
      email,
      name,
      period,
      form.artist,
      form.song,
      form.link,
      'PENDING'
    ]);

    return { success: true };

  } catch(e) {
    console.error("Song Request Error: " + e.message);
    return { success: false, message: "Transaction failed. Please try again." };
  }
}

/**
 * 11. TEACHER SONG MANAGER API
 */

/**
 * Helper: Normalize Period for Comparison
 * Handles "Period 1", "1", 1 -> "1"
 */
function normalizePeriod_(input) {
  if (input === null || input === undefined) return "";
  var s = String(input).trim();
  // Remove "Period" (case insensitive) and spaces
  s = s.replace(/^Period\s*/i, '');
  return s;
}

/**
 * Fetch pending and staged songs, filtered by period
 */
function getSongManagerData(periodFilter) {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);

  debugLog_("getSongManagerData called with filter: " + periodFilter);

  if (!sheet || sheet.getLastRow() < 3) {
    debugLog_("Song Sheet empty or missing.");
    return { period: null, requests: [] };
  }

  // Columns: [Timestamp, Email, Name, Period, Artist, Song, Link, Status]
  var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 8).getValues();
  debugLog_("Read " + data.length + " rows from Song Sheet.");

  var requests = [];
  data.forEach(function(row, index) {
    // Robust Status Check
    var status = row[7] ? String(row[7]).trim().toUpperCase() : "";

    if (status === 'PENDING' || status === 'STAGED') {
       requests.push({
         rowIndex: index + 3, // 1-based index for sheet updates
         timestamp: String(row[0]), // Force string to prevent serialization crash
         studentName: row[2],
         period: row[3], // Keep original for now
         artist: row[4],
         song: row[5],
         link: row[6],
         status: status
       });
    }
  });

  debugLog_("Found " + requests.length + " requests before period filter.");

  // Resolve Course ID -> Period
  var targetPeriod = null;
  var displayPeriod = null; // What we send to UI

  if (periodFilter) {
    var classroomsSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    if (classroomsSheet && classroomsSheet.getLastRow() >= 3) {
      var cData = classroomsSheet.getRange(3, 1, classroomsSheet.getLastRow() - 2, 4).getValues();
      // Col 1 = ID, Col 4 = Period
      for (var i = 0; i < cData.length; i++) {
        // Loose comparison for Course ID (String vs Number)
        if (String(cData[i][0]).trim() == String(periodFilter).trim()) {
          displayPeriod = cData[i][3];
          targetPeriod = normalizePeriod_(displayPeriod);
          debugLog_("Mapped Course " + periodFilter + " to Period '" + displayPeriod + "' -> Normalized '" + targetPeriod + "'");
          break;
        }
      }
    }
  }

  if (targetPeriod !== null) {
    var originalCount = requests.length;
    requests = requests.filter(function(r) {
      var rNorm = normalizePeriod_(r.period);
      var match = rNorm === targetPeriod;
      if (!match) {
        // Optional: Log misses if debugging heavy
        // debugLog_("Filter mismatch: Row Period '" + r.period + "' (" + rNorm + ") != Target '" + targetPeriod + "'");
      }
      return match;
    });
    debugLog_("Filtered from " + originalCount + " to " + requests.length + " requests for Period " + targetPeriod);
  } else {
    debugLog_("No target period resolved. Returning all " + requests.length + " requests.");
  }

  return {
    period: displayPeriod || targetPeriod, // Return the nice "Period 1" string if available
    requests: requests
  };
}

/**
 * Toggle Status (PENDING <-> STAGED)
 */
function toggleSongStaging(rowIndex, newStatus) {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);

  if (sheet) {
    sheet.getRange(rowIndex, 8).setValue(newStatus); // Col 8 is Status
  }
}

/**
 * Update Song Request Details (Artist, Song, Link)
 */
function updateSongRequest(rowIndex, artist, song, link) {
  checkAuth_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);

  if (sheet) {
    // Cols: 5=Artist, 6=Song, 7=Link
    sheet.getRange(rowIndex, 5).setValue(artist);
    sheet.getRange(rowIndex, 6).setValue(song);
    sheet.getRange(rowIndex, 7).setValue(link);
  }
}

/**
 * Finalize Playlist: Create on YouTube, Add Songs, Update Sheets
 */
function finalizePlaylist(courseId) {
  checkAuth_();

  // 1. Get Staged Songs for this class
  var dataObj = getSongManagerData(courseId);
  var staged = dataObj.requests.filter(function(r) { return r.status === 'STAGED'; });
  var period = dataObj.period;

  if (staged.length === 0) {
    return { success: false, message: "No songs staged." };
  }

  try {
    // 2. Create Playlist
    // Format: [YYYY-MM-DD] - Period [P] - [Day] Power Playlist
    var d = new Date();
    var timezone = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(d, timezone, 'yyyy-MM-dd');
    var dayName = Utilities.formatDate(d, timezone, 'EEEE'); // Full Day Name

    var title = dateStr + " - Period " + period + " - " + dayName + " Power Playlist";

    var playlist = YouTube.Playlists.insert({
      snippet: {
        title: title,
        description: "Generated by Multimedia Heroes for Period " + period,
        tags: ["MMH", "Classroom"]
      },
      status: {
        privacyStatus: "unlisted"
      }
    }, "snippet,status");

    var playlistId = playlist.id;
    var successCount = 0;

    // 3. Add Songs
    staged.forEach(function(req) {
      var videoId = null;

      // A. Try Link First
      if (req.link) {
        var match = req.link.match(/(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^"&?\/\s]{11})/);
        if (match) videoId = match[1];
      }

      // B. Fallback: Search
      if (!videoId) {
        var query = req.artist + " " + req.song + " lyrics"; // 'lyrics' usually helps find safe versions
        var search = YouTube.Search.list("id", {
          q: query,
          maxResults: 1,
          type: "video"
        });
        if (search.items && search.items.length > 0) {
          videoId = search.items[0].id.videoId;
        }
      }

      // C. Add to Playlist
      if (videoId) {
        try {
          YouTube.PlaylistItems.insert({
            snippet: {
              playlistId: playlistId,
              resourceId: {
                kind: "youtube#video",
                videoId: videoId
              }
            }
          }, "snippet");
          successCount++;

          // D. Update Sheet Status -> PLAYED
          var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONG_REQUESTS);
          sheet.getRange(req.rowIndex, 8).setValue('PLAYED');

        } catch (vidErr) {
          console.error("Error adding video " + videoId + ": " + vidErr.message);
        }
      }
    });

    var playlistUrl = "https://www.youtube.com/playlist?list=" + playlistId;
    return {
      success: true,
      message: "Created playlist with " + successCount + " songs.",
      playlistUrl: playlistUrl
    };

  } catch (e) {
    console.error("Playlist Error: " + e.message);
    return { success: false, message: "YouTube Error: " + e.message };
  }
}