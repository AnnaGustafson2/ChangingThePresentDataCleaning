/**
 * Runs every empty-cell test, logs PASS/FAIL (or ERROR if a test throws),
 * then logs a summary. Add new test functions to the array below.
 */
function empty_tests() {
  const testFunctions = [
    test_FormulaReturningEmpty,
    test_NonBreakingSpaces,
    test_ZeroWidthCharacters,
    test_LeadingApostrophesLiteral,
    test_LeadingApostropheMarker,
    test_EmptyRowAbove,
    test_EmptyRowMiddle,
    test_EmptyRowBelow,
    test_EmptyColumnLeft,
    test_EmptyColumnMiddle,
    test_EmptyColumnRight,
    test_HeadedEmptyColumnKept,
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

/**
 * Builds a test sheet, runs formula conversion and empty-cell cleaning,
 * and verifies that the resulting values, whitespace, and sheet dimensions
 * match the expected results.
 *
 * @param {string} sheetName Test sheet name (created if missing).
 * @param {Array<Array>} rows Contents written starting at A1.
 * @param {Array<Array>} expected Expected values after cleaning. Blank cells
 *     are ''. Trailing blank rows and columns should be omitted because
 *     they are expected to be deleted.
 * @param {Object} [options] Optional test settings.
 * @param {boolean} [options.text=false] Format cells as plain text before
 *     writing values. Useful for testing literal whitespace and invisible
 *     characters.
 * @returns {boolean} True if all checks pass; false otherwise.
 */
function runEmptyCellTest(sheetName, rows, expected, options = {}) {
  const { text = false } = options;
  const sheet = getOrCreateSheet(sheetName);

  const range = sheet.getRange(1, 1, rows.length, rows[0].length);
  if (text) range.setNumberFormat('@');
  range.setValues(rows);
  SpreadsheetApp.flush();

  convertFormulas(sheet);
  emptyCellClean(sheet);
  SpreadsheetApp.flush();

  const expRows = expected.length;
  const expCols = expected[0].length;
  const results = [];

  // 1. Remaining values
  const actual = sheet.getRange(1, 1, expRows, expCols).getValues();
  const valuesOk = JSON.stringify(actual) === JSON.stringify(expected);
  if (!valuesOk) Logger.log(`FAIL: expected ${JSON.stringify(expected)} got ${JSON.stringify(actual)}`);
  results.push(valuesOk);

  // 2. No whitespace-only cells remain
  const remaining = validateNoWhitespaceCells(sheet);
  if (remaining !== 0) Logger.log(`FAIL: ${remaining} whitespace-only cell(s) remain`);
  results.push(remaining === 0);

  // 3. Sheet was trimmed to the expected dimensions
  const sizeOk = sheet.getMaxRows() === expRows && sheet.getMaxColumns() === expCols;
  if (!sizeOk) {
    Logger.log(`FAIL: expected ${expRows}x${expCols} sheet, got ${sheet.getMaxRows()}x${sheet.getMaxColumns()}`);
  }
  results.push(sizeOk);

  return results.every(Boolean);
}

/* ---------- Tests ---------- */

/**
 * Tests that formulas returning empty strings are cleaned up,
 * while formulas returning actual values are preserved.
 */
function test_FormulaReturningEmpty() {
  return runEmptyCellTest('FormulaReturningEmpty', [
    ['Formula', 'Value'],
    ['Blank (IF)', '=IF(1=2,"NO","")'],
    ['Blank (literal)', '=""'],
    ['Keep', '=CONCAT("a", "b")'],
    ['=IF(1=2,"NO","")', '=""'],
  ], [
    ['Formula', 'Value'],
    ['Blank (IF)', ''],
    ['Blank (literal)', ''],
    ['Keep', 'ab'],
  ]);
}

/**
 * Tests that non-breaking and other Unicode space characters are treated
 * as blank when they are the only content in a cell, while spaces within
 * meaningful content are preserved.
 */
function test_NonBreakingSpaces() {
  return runEmptyCellTest('NonBreakingSpaces', [
    ['ID', 'Name'],
    ['\u00A0', '\u00A0\u00A0'],
    ['\u202F', '\u2007\u3000'],
    ['a\u00A0b', 'keep'],
    ['\u00A0', '\u00A0'],
  ], [
    ['ID', 'Name'],
    ['a\u00A0b', 'keep'],
  ], { text: true });
}

/**
 * Tests that zero-width and other invisible Unicode characters are treated
 * as blank when they are the only content in a cell, while invisible
 * characters within meaningful content are preserved.
 */
function test_ZeroWidthCharacters() {
  return runEmptyCellTest('ZeroWidthCharacters', [
    ['ID', 'Name'],
    ['\u200B', '\u200C\u200D'],
    ['\uFEFF', '\u2060'],
    ['\u2800', '\u3164\uFFA0'],
    ['x\u200B', 'keep'],
    ['\u200B', '\u200B'],
  ], [
    ['ID', 'Name'],
    ['x\u200B', 'keep'],
  ], { text: true });
}

/**
 * Tests that literal apostrophes and apostrophe-only cells are handled
 * correctly, while meaningful text following an apostrophe is preserved.
 */
function test_LeadingApostrophesLiteral() {
  return runEmptyCellTest('LeadingApostrophesLiteral', [
    ['ID', 'Name'],
    ["'", "'   "],
    ["''", '\u2019'],
    ["'\u200B", "'\u00A0"],
    ["'hello", 'keep'],
    ["'", "'"],
  ], [
    ['ID', 'Name'],
    ['hello', 'keep'],
  ], { text: true });
}

/**
 * Tests that Sheets' leading-apostrophe markers are removed when they
 * are used to preserve literal values, while the underlying values remain.
 */
function test_LeadingApostropheMarker() {
  return runEmptyCellTest('LeadingApostropheMarker', [
    ['ID', 'Name'],
    ["'", "'   "],
    ["'02134", "'hello"],
    ["'", "'"],
  ], [
    ['ID', 'Name'],
    ['02134', 'hello'],
  ]);
}

/**
 * Tests that an empty row above the data is removed.
 */
function test_EmptyRowAbove() {
  return runEmptyCellTest('EmptyRowAbove', [
    ['', ''],
    ['ID', 'Name'],
    ['a', 'b'],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ]);
}

/**
 * Tests that an empty row in the middle of the data is removed.
 */
function test_EmptyRowMiddle() {
  return runEmptyCellTest('EmptyRowMiddle', [
    ['ID', 'Name'],
    ['', ''],
    ['', '\u200B'],
    ['a', 'b'],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ], { text: true });
}

/**
 * Tests that an empty row below the data is removed.
 */
function test_EmptyRowBelow() {
  return runEmptyCellTest('EmptyRowBelow', [
    ['ID', 'Name'],
    ['a', 'b'],
    ['', ''],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ]);
}

/**
 * Tests that an empty column to the left of the data is removed.
 */
function test_EmptyColumnLeft() {
  return runEmptyCellTest('EmptyColumnLeft', [
    ['', 'ID', 'Name'],
    ['', 'a', 'b'],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ]);
}

/**
 * Tests that an empty column in the middle of the data is removed.
 */
function test_EmptyColumnMiddle() {
  return runEmptyCellTest('EmptyColumnMiddle', [
    ['ID', '', 'Name'],
    ['a', '', 'b'],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ]);
}

/**
 * Tests that an empty column to the right of the data is removed.
 */
function test_EmptyColumnRight() {
  return runEmptyCellTest('EmptyColumnRight', [
    ['ID', 'Name', ''],
    ['a', 'b', ''],
  ], [
    ['ID', 'Name'],
    ['a', 'b'],
  ]);
}

/**
 * Tests that a column containing a header is preserved even when its
 * data cells are empty.
 */
function test_HeadedEmptyColumnKept() {
  return runEmptyCellTest('HeadedEmptyColumn', [
    ['ID', 'Notes', 'Name'],
    ['a', '', 'b'],
  ], [
    ['ID', 'Notes', 'Name'],
    ['a', '', 'b'],
  ]);
}