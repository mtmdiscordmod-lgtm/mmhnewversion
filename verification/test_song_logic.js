// Mock Objects for testing logic without Google Services
const mockSheet = {
  data: [
    ['2023-01-01', 'stu@sandi.net', 'Student A', 1, 'Rick Astley', 'Never Gonna Give You Up', 'https://youtu.be/dQw4w9WgXcQ', 'PENDING'],
    ['2023-01-01', 'stu2@sandi.net', 'Student B', 1, 'Darude', 'Sandstorm', '', 'STAGED']
  ]
};

// Test 1: Naming Convention
function testPlaylistNaming() {
  const period = 1;
  const d = new Date('2023-10-27T12:00:00'); // Friday
  const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const dateStr = d.toISOString().slice(0,10);
  const dayName = days[d.getDay()];

  const expectedTitle = "2023-10-27 - Period 1 - Friday Power Playlist";
  const actualTitle = dateStr + " - Period " + period + " - " + dayName + " Power Playlist";

  if (expectedTitle !== actualTitle) {
    console.error(`FAIL: Naming Convention.\nExpected: ${expectedTitle}\nActual:   ${actualTitle}`);
    process.exit(1);
  } else {
    console.log("PASS: Naming Convention");
  }
}

// Test 2: Filter Logic (Mocking getSongManagerData)
function testFilterLogic() {
  const requests = [];
  mockSheet.data.forEach((row, index) => {
    const status = row[7];
    if (status === 'PENDING' || status === 'STAGED') {
      requests.push({ period: row[3], status: status });
    }
  });

  const p1Requests = requests.filter(r => r.period === 1);

  if (p1Requests.length !== 2) {
    console.error(`FAIL: Filter Logic. Expected 2 requests, got ${p1Requests.length}`);
    process.exit(1);
  } else {
    console.log("PASS: Filter Logic");
  }
}

testPlaylistNaming();
testFilterLogic();
