function testGradingFlow() {
  console.log("Starting Grading Flow Test...");

  try {
    // 1. Test Folder Creation
    console.log("Testing Folder Creation...");
    var period = "Test Period 99";
    var studentName = "Test Student";
    var folder = getOrCreateGradingFolder_(period, studentName);
    console.log("Folder created/found: " + folder.getUrl());

    // 2. Test Doc Creation (Perfect Score)
    console.log("Testing Doc Creation (Perfect Score)...");
    var doc1 = createGradingDoc_(folder, studentName, "Mission Alpha", 100, 100, "Excellent performance.");
    console.log("Doc 1 created: " + doc1.getUrl());

    // 3. Test Doc Creation (Imperfect Score)
    console.log("Testing Doc Creation (Imperfect Score)...");
    var doc2 = createGradingDoc_(folder, studentName, "Mission Beta", 80, 100, "Missed a few spots.");
    console.log("Doc 2 created: " + doc2.getUrl());

    console.log("Test Complete. Please check 'Multimedia Heroes/Grading Logs/Period Test Period 99/Test Student' in Drive.");

  } catch (e) {
    console.error("Test Failed: " + e.message);
    if (e.stack) console.error(e.stack);
  }
}
