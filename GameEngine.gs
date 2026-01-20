/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * GAME ENGINE (Backend)
 * Handles state, logic, and concurrency for the "Bonus Round" roguelike mode.
 * ═══════════════════════════════════════════════════════════════════════════════
 */

var GameService = {

  CACHE_PREFIX: 'MMH_GAME_',
  BUCKET_COUNT: 20,

  /**
   * 1. INITIALIZE GAME
   * Calculates Class MP, sets up initial state in Cache.
   */
  initializeGame: function(courseId, questionSetId) {
    checkAuth_(); // Ensure teacher is calling this

    // A. Calculate Total Class MP
    var totalMP = this.calculateClassMP_(courseId);
    var multiplier = 2.5;
    var initialEnergy = Math.floor(totalMP * multiplier);

    // B. Load Questions
    var questions = this.loadQuestions_(questionSetId); // For now, loads all from sheet
    if (questions.length === 0) throw new Error("No questions found in BONUS_ROUNDS sheet.");

    // C. Create Game ID
    var gameId = Utilities.getUuid();

    // D. Initialize State
    var state = {
      gameId: gameId,
      courseId: courseId,
      active: true,
      currentRound: 0, // Index of questions array
      maxEnergy: initialEnergy,
      currentEnergy: initialEnergy,
      questions: questions, // Store full set in cache (tier 2 storage)
      streak: 0,
      status: 'WAITING' // WAITING, PLAYING, GAME_OVER, VICTORY
    };

    // E. Write to Cache (Expiration: 2 hours)
    var cache = CacheService.getScriptCache();
    cache.put(this.CACHE_PREFIX + gameId, JSON.stringify(state), 7200);

    // F. Clear old buckets for this ID (just in case of collision, unlikely)
    this.clearBuckets_(gameId);

    return {
      success: true,
      gameId: gameId,
      energy: initialEnergy,
      questionCount: questions.length
    };
  },

  /**
   * 2. GENERATE CLASSROOM LINK
   * Posts the "Magic Link" to the stream.
   */
  generateClassroomLink: function(courseId, gameId) {
    checkAuth_();

    var scriptUrl = ScriptApp.getService().getUrl();
    var magicLink = scriptUrl + "?page=bonus&gameId=" + gameId;

    try {
      Classroom.Courses.Announcements.create({
        text: "🚨 BONUS ROUND DETECTED! 🚨\n\nClick the link to join the battle!",
        materials: [{
          link: {
            url: magicLink,
            title: "⚔️ JOIN BONUS ROUND"
          }
        }],
        state: 'PUBLISHED',
        assigneeMode: 'ALL_STUDENTS'
      }, courseId);

      return { success: true, message: "Link posted to Classroom!" };
    } catch (e) {
      console.error("Announcement Error: " + e.message);
      return { success: false, message: "Error posting link: " + e.message };
    }
  },

  /**
   * 3. PROCESS ANSWER SHARD (Student)
   * Writes student answer to a random bucket to avoid lock contention.
   */
  processAnswerShard: function(gameId, studentEmail, answerIndex) {
    // Note: No heavy auth check here for speed, relying on email from client (validated by doGet context ideally,
    // but for high throughput game, we trust the param with basic checks).
    // In a strict environment, we'd validate Session.getActiveUser().getEmail() matches studentEmail.

    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { success: false, error: "Not logged in" };

    // Randomized Bucket Selection (0-19)
    var bucketId = Math.floor(Math.random() * this.BUCKET_COUNT);
    var cacheKey = this.CACHE_PREFIX + gameId + '_BUCKET_' + bucketId;

    var lock = LockService.getScriptLock();
    // Quick lock, fail fast if busy (client will retry or jitter handles it)
    if (lock.tryLock(500)) {
      try {
        var cache = CacheService.getScriptCache();
        var raw = cache.get(cacheKey);
        var bucket = raw ? JSON.parse(raw) : [];

        // Add answer
        bucket.push({
          email: userEmail,
          answer: answerIndex,
          timestamp: new Date().getTime()
        });

        cache.put(cacheKey, JSON.stringify(bucket), 7200);
        return { success: true };
      } catch (e) {
        return { success: false, error: e.message };
      } finally {
        lock.releaseLock();
      }
    } else {
      return { success: false, error: "Busy" };
    }
  },

  /**
   * 4. GAME LOOP TICK (Teacher Driver)
   * Aggregates answers, calculates damage, updates state.
   */
  gameLoopTick: function(gameId) {
    checkAuth_(); // Teacher only

    var cache = CacheService.getScriptCache();
    var stateRaw = cache.get(this.CACHE_PREFIX + gameId);
    if (!stateRaw) return { error: "Game not found" };

    var state = JSON.parse(stateRaw);
    if (state.status !== 'PLAYING') return state; // Just return current state if not active

    var currentQ = state.questions[state.currentRound];
    var correctAnswer = this.parseAnswerJSON_(currentQ.answerJson).correct;

    // A. Aggregate Buckets
    var allAnswers = [];
    var bucketKeys = [];
    for (var i = 0; i < this.BUCKET_COUNT; i++) {
      bucketKeys.push(this.CACHE_PREFIX + gameId + '_BUCKET_' + i);
    }

    var bucketMap = cache.getAll(bucketKeys);

    // Clear buckets immediately after reading to prevent double counting
    // (Ideally we'd use a transaction, but cache.removeAll is good enough for this eventual consistency)
    cache.removeAll(bucketKeys);

    Object.values(bucketMap).forEach(function(json) {
      if (json) {
        var answers = JSON.parse(json);
        allAnswers = allAnswers.concat(answers);
      }
    });

    // B. Deduplicate (Last answer counts? Or first? Let's say First valid answer per question period)
    // Actually, simpler: One answer per student per tick.
    // Since we clear buckets, we just process what we grabbed.

    // C. Calculate Damage
    var wrongCount = 0;
    var correctCount = 0;

    allAnswers.forEach(function(a) {
      if (a.answer == correctAnswer) {
        correctCount++;
      } else {
        wrongCount++;
      }
    });

    // Roguelike Scaling
    // Damage = Base * (1 + (Round * 0.15))
    var baseDamage = currentQ.damage || 10;
    var scaling = 1 + (state.currentRound * 0.15);
    var totalDamage = Math.floor(wrongCount * baseDamage * scaling);

    state.currentEnergy -= totalDamage;
    if (state.currentEnergy < 0) state.currentEnergy = 0;

    // Update Streak
    if (wrongCount === 0 && correctCount > 0) {
      state.streak++;
    } else if (wrongCount > 0) {
      state.streak = 0;
    }

    // Check Game Over
    if (state.currentEnergy <= 0) {
      state.status = 'GAME_OVER';
    }

    // D. Save State
    state.lastTickDamage = totalDamage; // For frontend FX
    state.lastTickCorrect = correctCount;
    state.lastTickWrong = wrongCount;

    cache.put(this.CACHE_PREFIX + gameId, JSON.stringify(state), 7200);

    return state;
  },

  /**
   * 5. GET GAME STATE (Student/Teacher)
   */
  getGameState: function(gameId) {
    // No auth check needed for read-only state (high traffic)
    var cache = CacheService.getScriptCache();
    var stateRaw = cache.get(this.CACHE_PREFIX + gameId);

    if (!stateRaw) return { error: "Game not found" };

    var state = JSON.parse(stateRaw);

    // Mask sensitive data for students
    return {
      active: state.active,
      status: state.status,
      currentRound: state.currentRound,
      totalRounds: state.questions.length,
      currentEnergy: state.currentEnergy,
      maxEnergy: state.maxEnergy,
      question: state.questions[state.currentRound],
      lastDamage: state.lastTickDamage,
      streak: state.streak
    };
  },

  /**
   * 6. TRANSITION TO NEXT ROUND
   */
  nextQuestion: function(gameId) {
    checkAuth_();

    var cache = CacheService.getScriptCache();
    var stateRaw = cache.get(this.CACHE_PREFIX + gameId);
    if (!stateRaw) return { error: "Game not found" };

    var state = JSON.parse(stateRaw);

    if (state.currentRound < state.questions.length - 1) {
      state.currentRound++;
      state.status = 'PLAYING'; // Ensure playing
      // Reset tick stats
      state.lastTickDamage = 0;
      state.lastTickCorrect = 0;
      state.lastTickWrong = 0;

      cache.put(this.CACHE_PREFIX + gameId, JSON.stringify(state), 7200);
      return { success: true, round: state.currentRound };
    } else {
      state.status = 'VICTORY';
      cache.put(this.CACHE_PREFIX + gameId, JSON.stringify(state), 7200);

      // LOG REWARDS HERE (Future)
      this.logVictory_(state.gameId, state.courseId, state.questions.length);

      return { success: true, status: 'VICTORY' };
    }
  },

  startGame: function(gameId) {
    checkAuth_();
    var cache = CacheService.getScriptCache();
    var stateRaw = cache.get(this.CACHE_PREFIX + gameId);
    if (stateRaw) {
      var state = JSON.parse(stateRaw);
      state.status = 'PLAYING';
      cache.put(this.CACHE_PREFIX + gameId, JSON.stringify(state), 7200);
      return { success: true };
    }
    return { success: false };
  },

  endGame: function(gameId) {
    checkAuth_();
    var cache = CacheService.getScriptCache();
    cache.remove(this.CACHE_PREFIX + gameId);
    return { success: true };
  },

  // --- HELPERS ---

  calculateClassMP_: function(courseId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PROGRESS);
    if (!sheet) return 1000; // Default fallback

    // Map courseId to Period string (e.g., "1")
    var period = this.getPeriodByCourseId_(courseId);
    if (!period) return 1000;

    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 7).getValues(); // Col 3 is Period, Col 7 is Total MP
    var total = 0;

    data.forEach(function(row) {
      // Normalize comparison
      var pRow = String(row[2]).replace(/\D/g, ''); // Extract number
      var pTarget = String(period).replace(/\D/g, '');

      if (pRow === pTarget) {
        total += (Number(row[6]) || 0);
      }
    });

    return total > 0 ? total : 1000; // Minimum floor
  },

  getPeriodByCourseId_: function(courseId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG_CLASSROOMS);
    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 4).getValues(); // Col 1 ID, Col 4 Period

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(courseId)) {
        return data[i][3];
      }
    }
    return null;
  },

  loadQuestions_: function(setId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BONUS_ROUNDS);
    if (!sheet) return [];

    // Columns: RoundID, Question, Media, Type, AnswerJSON, Damage, Time
    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 7).getValues();

    var questions = [];
    data.forEach(function(row) {
      if (row[1]) { // If question text exists
        questions.push({
          id: row[0],
          text: row[1],
          media: row[2],
          type: row[3],
          answerJson: row[4], // Keep string for transport, parse when needed
          damage: row[5],
          timeLimit: row[6]
        });
      }
    });

    return questions;
  },

  parseAnswerJSON_: function(jsonStr) {
    try {
      return JSON.parse(jsonStr);
    } catch (e) {
      return { options: [], correct: 0 };
    }
  },

  clearBuckets_: function(gameId) {
    var cache = CacheService.getScriptCache();
    for (var i = 0; i < this.BUCKET_COUNT; i++) {
      cache.remove(this.CACHE_PREFIX + gameId + '_BUCKET_' + i);
    }
  },

  logVictory_: function(gameId, courseId, rounds) {
    // Placeholder for future reward logic
    console.log("Victory logged: " + gameId + " - " + rounds + " rounds.");
  }
};

/**
 * EXPOSED FUNCTIONS FOR CLIENT
 */
function initGame(courseId) { return GameService.initializeGame(courseId); }
function postGameLink(courseId, gameId) { return GameService.generateClassroomLink(courseId, gameId); }
function submitGameAnswer(gameId, answerIdx) { return GameService.processAnswerShard(gameId, '', answerIdx); }
function tickGameLoop(gameId) { return GameService.gameLoopTick(gameId); }
function fetchGameState(gameId) { return GameService.getGameState(gameId); }
function advanceRound(gameId) { return GameService.nextQuestion(gameId); }
function setGameActive(gameId) { return GameService.startGame(gameId); }
function stopGame(gameId) { return GameService.endGame(gameId); }
