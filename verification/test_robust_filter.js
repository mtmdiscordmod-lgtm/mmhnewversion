// Mock normalization function from WebApp.gs
function normalizePeriod_(input) {
  if (input === null || input === undefined) return "";
  var s = String(input).trim();
  s = s.replace(/^Period\s*/i, '');
  return s;
}

// Test Cases
function runTests() {
  const tests = [
    { input: "1", expected: "1", name: "String '1'" },
    { input: 1, expected: "1", name: "Number 1" },
    { input: "Period 1", expected: "1", name: "String 'Period 1'" },
    { input: "period 1", expected: "1", name: "String 'period 1'" },
    { input: "Period  1", expected: "1", name: "String 'Period  1' (spaces)" },
    { input: " 1 ", expected: "1", name: "String ' 1 ' (trim)" }
  ];

  let failures = 0;
  tests.forEach(t => {
    const res = normalizePeriod_(t.input);
    if (res !== t.expected) {
      console.error(`FAIL: ${t.name} -> Expected '${t.expected}', Got '${res}'`);
      failures++;
    } else {
      console.log(`PASS: ${t.name}`);
    }
  });

  if (failures > 0) process.exit(1);
}

runTests();
