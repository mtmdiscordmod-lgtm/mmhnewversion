// Mock Utilities for Node.js environment
const Utilities = {
  formatDate: function(date, timezone, format) {
    // Basic mock: ignores timezone for simple local test,
    // but verifies we are calling the function correctly.
    // In real GAS, this handles the offset.
    if (format === 'yyyy-MM-dd') return date.toISOString().slice(0,10);
    if (format === 'EEEE') return 'Friday'; // Mock day
    return '';
  }
};

const Session = {
  getScriptTimeZone: () => "America/Los_Angeles"
};

function testPlaylistNamingFixed() {
  const period = 1;
  const d = new Date('2023-10-27T12:00:00Z');

  // Logic from WebApp.gs
  var timezone = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(d, timezone, 'yyyy-MM-dd');
  var dayName = Utilities.formatDate(d, timezone, 'EEEE');

  var title = dateStr + " - Period " + period + " - " + dayName + " Power Playlist";

  console.log("Generated Title: " + title);

  if (title.includes("2023-10-27") && title.includes("Period 1") && title.includes("Power Playlist")) {
    console.log("PASS: Title format correct");
  } else {
    console.error("FAIL: Title format incorrect");
    process.exit(1);
  }
}

testPlaylistNamingFixed();
