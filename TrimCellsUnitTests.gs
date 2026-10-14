/**
 * Runs every trim test, logs PASS/FAIL (or ERROR if a test throws),
 * then logs a summary. Add new test functions to the array below.
 */
function trim_tests() {
  const testFunctions = [
    test_TrimSpaces,
    test_TrimTabs,
    test_TrimNewLines,
    test_TrimNonBreakingSpaces,
    test_TrimZeroWidthCharacters,
    test_TrimKeepsInternalSpacing,
    test_TrimLeavesNonTextAlone,
    test_TrimKeepsEmojiSelector,
    test_TrimIdempotent,
    test_TrimSkipsFormulaLookingText,
    test_TrimAllowsApprovedFormulaCreation,
    test_TrimRejectsUnapprovedFormulaCreation,
  ];

  let passed = 0;
  testFunctions.forEach(fn => {
    let ok = false;
    try {
      ok = fn();
    } catch (e) {
      Logger.log(`${fn.name}: ERROR - ${e.message}`);
    }
    if (ok) passed++;
    Logger.log(`${fn.name}: ${ok ? 'PASS' : 'FAIL'}`);
  });

  Logger.log(`${passed} out of ${testFunctions.length} tests passed`);
}

/* ---------- Helpers ---------- */

/** Turns a grid into a string that distinguishes types (so 42 and "42" differ). */
function describeGrid(grid) {
  return JSON.stringify(grid.map(row => row.map(v =>
    v instanceof Date ? `date:${v.getTime()}` : `${typeof v}:${v}`
  )));
}

/**
 * Builds a test sheet, runs trimCells, and verifies that:
 * 1. The input was stored as written (otherwise the test setup is wrong).
 * 2. The remaining values and their types match `expected`.
 * 3. No formulas were created by trimming.
 * 4. (Optional) The number of cells skipped matches `expectSkipped`.
 *
 * @param {string} sheetName Test sheet name (created if missing).
 * @param {Array<Array>} rows Contents written starting at A1.
 * @param {Array<Array>} expected Expected values after trimming.
 * @param {Object} [options]
 * @param {boolean} [options.text=true] Format cells as plain text first, so
 *     padded strings are stored literally. Use false for mixed-type tests.
 * @param {number} [options.passes=1] How many times to run trimCells.
 * @param {number} [options.expectSkipped] Expected count returned by trimCells.
 * @returns {boolean} True if every check passes.
 */
function runTrimTest(sheetName, rows, expected, options = {}) {
  const { text = true, passes = 1, expectSkipped } = options;
  const sheet = getOrCreateSheet(sheetName);
  const numRows = rows.length;
  const numCols = rows[0].length;

  const range = sheet.getRange(1, 1, numRows, numCols);
  if (text) range.setNumberFormat('@');
  range.setValues(rows);
  SpreadsheetApp.flush();

  const results = [];

  // 1. Setup check: Sheets stored the input as written
  const setupOk = describeGrid(range.getValues()) === describeGrid(rows);
  if (!setupOk) Logger.log('FAIL: setup changed the input before trimming');
  results.push(setupOk);

  let skipped;
  for (let i = 0; i < passes; i++) skipped = trimCells(sheet);
  SpreadsheetApp.flush();

  // 2. Remaining values and types
  const actual = sheet.getRange(1, 1, numRows, numCols).getValues();
  const valuesOk = describeGrid(actual) === describeGrid(expected);
  if (!valuesOk) Logger.log(`FAIL: expected ${describeGrid(expected)} got ${describeGrid(actual)}`);
  results.push(valuesOk);

  // 3. Trimming did not create formulas
  results.push(checkNoFormulas(sheet));

  // 4. Skipped count
  if (expectSkipped !== undefined) {
    const skipOk = skipped === expectSkipped;
    if (!skipOk) Logger.log(`FAIL: expected ${expectSkipped} skipped cell(s), got ${skipped}`);
    results.push(skipOk);
  }

  return results.every(Boolean); // all checks run first, so every failure is logged
}

/* ---------- Tests ---------- */

function test_TrimSpaces() {
  return runTrimTest('TrimSpaces', [
    ['ID', 'Name'],
    ['    Jane', 'Jane   '],               // leading, trailing
    ['     Jane    ', 'Jane'],             // both sides, already-clean neighbor
    ['  Jane Doe  ', 'Mary Ann '],        // inner space kept
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Jane Doe', 'Mary Ann'],
  ]);
}

function test_TrimTabs() {
  return runTrimTest('TrimTabs', [
    ['ID', 'Name'],
    ['\tJane', 'Jane\t'],                  // leading, trailing
    ['\t\tJane\t\t', 'Jane'],              // repeated
    [' \t Jane \t ', 'Jane\tDoe'],         // mixed with spaces; inner tab kept
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane\tDoe'],
  ]);
}

function test_TrimNewLines() {
  return runTrimTest('TrimNewLines', [
    ['ID', 'Name'],
    ['\nJane', 'Jane\n'],                  // leading, trailing
    ['\n\nJane\n\n', 'Jane'],              // repeated
    ['Jane\nDoe', '\n Jane\nDoe \n'],      // inner new line kept
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Jane\nDoe', 'Jane\nDoe'],
  ]);
}

function test_TrimNonBreakingSpaces() {
  return runTrimTest('TrimNonBreakingSpaces', [
    ['ID', 'Name'],
    ['\u00A0Jane', 'Jane\u00A0'],          // no-break space
    ['\u202FJane\u2007', '\u3000Jane\u3000'], // narrow no-break, figure, ideographic
    ['Jane\u00A0Doe', '\u00A0Jane\u00A0Doe\u00A0'], // inner NBSP kept
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Jane\u00A0Doe', 'Jane\u00A0Doe'],
  ]);
}

function test_TrimZeroWidthCharacters() {
  return runTrimTest('TrimZeroWidthCharacters', [
    ['ID', 'Name'],
    ['\u200BJane', 'Jane\u200B'],               // zero-width space
    ['\uFEFFJane\u2060', '\u200C\u200DJane\u200D'], // BOM, word joiner, ZWNJ, ZWJ
    ['\u2800Jane\u3164', 'Jane'],               // Braille blank, Hangul filler
    ['Ja\u200Bne', '\u200BJa\u200Bne\u200B'],   // inner zero-width char kept
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Jane', 'Jane'],
    ['Ja\u200Bne', 'Ja\u200Bne'],
  ]);
}

function test_TrimKeepsInternalSpacing() {
  return runTrimTest('TrimKeepsInternalSpacing', [
    ['ID', 'Name'],
    ['Jane   Doe', '  Jane   Doe  '],     // only the edges change
    ['A  B  C', 'x'],                     // nothing to trim: untouched
  ], [
    ['ID', 'Name'],
    ['Jane   Doe', 'Jane   Doe'],
    ['A  B  C', 'x'],
  ]);
}

function test_TrimLeavesNonTextAlone() {
  return runTrimTest('TrimLeavesNonTextAlone', [
    ['ID', 'Value'],
    ['  Jane  ', 42],                     // text trimmed, number untouched
    [' Pi ', 3.14],
    [' Zero ', 0],                        // 0 and false must not be treated as blank
    [' Yes ', true],
    [' No ', false],
    [' Date ', new Date(2026, 8, 27)],
  ], [
    ['ID', 'Value'],
    ['Jane', 42],
    ['Pi', 3.14],
    ['Zero', 0],
    ['Yes', true],
    ['No', false],
    ['Date', new Date(2026, 8, 27)],
  ], { text: false });
}

function test_TrimKeepsEmojiSelector() {
  return runTrimTest('TrimKeepsEmojiSelector', [
    ['ID', 'Name'],
    [' Jane \u2764\uFE0F ', 'Jane \u{1F44D} '],   // variation selector at the end
  ], [
    ['ID', 'Name'],
    ['Jane \u2764\uFE0F', 'Jane \u{1F44D}'],      // padding gone, emoji intact
  ]);
}

function test_TrimIdempotent() {
  return runTrimTest('TrimIdempotent', [
    ['ID', 'Name'],
    ['Jane', '  Jane  '],
    ['  Doe', 'Doe'],
  ], [
    ['ID', 'Name'],
    ['Jane', 'Jane'],
    ['Doe', 'Doe'],
  ], { passes: 2 });                            // a second run changes nothing
}

function test_TrimSkipsFormulaLookingText() {
  return runTrimTest('TrimSkipsFormulaLookingText', [
    ['ID', 'Name'],
    ['  =SUM(1,2) ', '  =IMPORTXML("x","y") '],  // would become live formulas
    ['  Jane ', 'ok'],
  ], [
    ['ID', 'Name'],
    ['  =SUM(1,2) ', '  =IMPORTXML("x","y") '],  // left as the original text
    ['Jane', 'ok'],                              // normal cell still trimmed
  ], { expectSkipped: 2 });
}

function test_TrimAllowsApprovedFormulaCreation() {
  const sheet = getOrCreateSheet('TrimApprovedFormula');

  sheet.getRange('A1').setNumberFormat('@');
  sheet.getRange('A1').setValue('  =SUM(1,2)  ');

  trimCells(sheet, [], true);

  const value = sheet.getRange('A1').getValue();

  return value === 3;
}

function test_TrimRejectsUnapprovedFormulaCreation() {
  const sheet = getOrCreateSheet('TrimUnapprovedFormula');

  sheet.getRange('A1').setNumberFormat('@');
  sheet.getRange('A1').setValue('  =IMPORTXML("x","y")  ');

  try {
    trimCells(sheet, [], true);
    return false;
  } catch (e) {
    return e.message.includes('unapproved');
  }
}