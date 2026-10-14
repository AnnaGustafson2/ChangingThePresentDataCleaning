// Shorthand for expected JavaScript types
const S = 'string', N = 'number', B = 'boolean', D = 'object'; // D = Date

/**
 * Runs every test, logs PASS/FAIL (or ERROR if a test throws),
 * then logs a summary. Add new test functions to the arrays below.
 */
function run_tests() {
  const formulaTestFunctions = [
    test_convertFormulas_singleFormulaText,
    test_convertFormulas_singleFormulaNumber,
    test_convertFormulas_singleFormulaBoolean,
    test_convertFormulas_singleFormulaDate,
    test_convertFormulas_singleFormulaBlank,
    test_convertFormulas_multipleFormulas,
    test_convertFormulas_formattingFormulas,
    test_convertFormulas_errorFormulas,
    test_convertFormulas_singleCellFormula,
    test_convertFormulas_dependencies,
    test_convertFormulas_changingData,
    test_convertFormulas_multipleCells,
    test_convertFormulas_approvedFunctions,
    test_everyApprovedFunctionHasATest,
    test_convertFormulas_rangeCriteria,
    test_convertFormulas_textThatLooksLikeNumbers,
  ];

  let formulaPassed = 0;
  formulaTestFunctions.forEach(fn => {
    let ok = false;
    try {
      ok = fn();
    } catch (e) {
      Logger.log(`${fn.name}: ERROR - ${e.message}`);
    }
    if (ok) formulaPassed++;
    Logger.log(`${fn.name}: ${ok ? 'PASS' : 'FAIL'}`);
  });

  const emptyTestFunctions = [
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

  let emptyPassed = 0;
  emptyTestFunctions.forEach(fn => {
    let ok = false;
    try {
      ok = fn();
    } catch (e) {
      Logger.log(`${fn.name}: ERROR - ${e.message}`);
    }
    if (ok) emptyPassed++;
    Logger.log(`${fn.name}: ${ok ? 'PASS' : 'FAIL'}`);
  });

  const trimTestFunctions = [
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

  let trimPassed = 0;
  trimTestFunctions.forEach(fn => {
    let ok = false;
    try {
      ok = fn();
    } catch (e) {
      Logger.log(`${fn.name}: ERROR - ${e.message}`);
    }
    if (ok) trimPassed++;
    Logger.log(`${fn.name}: ${ok ? 'PASS' : 'FAIL'}`);
  });

  Logger.log(`${formulaPassed} out of ${formulaTestFunctions.length} formula tests passed`);
  Logger.log(`${emptyPassed} out of ${emptyTestFunctions.length} empty cell tests passed`);
  Logger.log(`${trimPassed} out of ${trimTestFunctions.length} trim tests passed`);

}

/* ---------- Helpers ---------- */

/** Returns a cleared sheet with the given name, creating it if needed. */
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clear();
  return sheet;
}

/** Trims a sheet down to a single cell (A1). */
function shrinkToOneCell(sheet) {
  if (sheet.getMaxRows() > 1) sheet.deleteRows(2, sheet.getMaxRows() - 1);
  if (sheet.getMaxColumns() > 1) sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
}

/** Passes if the sheet has no formulas left. */
function checkNoFormulas(sheet) {
  const remaining = [];
  sheet.getDataRange().getFormulas().forEach((row, r) =>
    row.forEach((formula, c) => {
      if (formula) remaining.push({ row: r + 1, col: c + 1, formula });
    })
  );
  if (remaining.length) Logger.log(`FAIL: formulas remaining: ${JSON.stringify(remaining)}`);
  return remaining.length === 0;
}

/** Passes if the cell's number format is unchanged. */
function checkFormat(cell, before, after) {
  if (before === after) return true;
  Logger.log(`FAIL: ${cell} format changed from "${before}" to "${after}"`);
  return false;
}

/** Passes if the cell's value has the expected JavaScript type. */
function checkType(sheet, cell, expected) {
  const actual = typeof sheet.getRange(cell).getValue();
  if (actual === expected) return true;
  Logger.log(`FAIL: ${cell} expected type "${expected}" but got "${actual}"`);
  return false;
}

/**
 * Builds a test sheet, runs flag + convert, and verifies that no formulas
 * remain, number formats are unchanged, and values have the expected types.
 *
 * @param {string} sheetName Test sheet name (created if missing).
 * @param {Array<Array>} rows Contents written starting at A1. Strings starting
 *     with "=" become formulas.
 * @param {Array<Array<string|null>>} types Expected type for each cell, laid out
 *     like the sheet. Use null to skip a cell. May extend past `rows` (for spills).
 * @param {Object} [options]
 * @param {Object<string,string>} [options.formats] Number formats to apply, by A1 address.
 * @param {boolean} [options.header=true] Shade the first row.
 * @param {boolean} [options.shrink=false] Shrink the sheet to one cell first.
 * @returns {boolean} True if every check passes.
 */
function runConversionTest(sheetName, rows, types, options = {}) {
  const { formats = {}, header = true, shrink = false } = options;
  const sheet = getOrCreateSheet(sheetName);

  if (shrink) shrinkToOneCell(sheet);
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  if (header) sheet.getRange(1, 1, 1, rows[0].length).setBackground('#d9ead3');
  Object.entries(formats).forEach(([cell, fmt]) => sheet.getRange(cell).setNumberFormat(fmt));

  // Flatten the types grid into { 'A1': 'string', ... }
  const expected = {};
  types.forEach((row, r) =>
    row.forEach((type, c) => {
      if (type) expected[sheet.getRange(r + 1, c + 1).getA1Notation()] = type;
    })
  );
  const cells = Object.keys(expected);

  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  flagUnapprovedFormulas(sheet);
  convertFormulas(sheet);

  const results = [checkNoFormulas(sheet)];
  cells.forEach(cell => {
    const after = sheet.getRange(cell).getNumberFormat();
    const gainedDateFormat = expected[cell] === D && formatsBefore[cell] === '';
    results.push(gainedDateFormat || checkFormat(cell, formatsBefore[cell], after));
    results.push(checkType(sheet, cell, expected[cell]));
  });
  return results.every(Boolean);
}

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

/* ---------- Convert Formula Tests ---------- */

function test_convertFormulas_singleFormulaText() {
  return runConversionTest('SingleFormulaText', [
    ['Formula', 'Value'],
    ['Concatenation', '=CONCAT("John", " Doe")'],
  ], [
    [S, S],
    [S, S],
  ]);
}

function test_convertFormulas_singleFormulaNumber() {
  return runConversionTest('SingleFormulaNumber', [
    ['Formula', 'Value'],
    ['Addition', '=1+1'],
  ], [
    [S, S],
    [S, N],
  ]);
}

function test_convertFormulas_singleFormulaBoolean() {
  return runConversionTest('SingleFormulaBoolean', [
    ['Formula', 'Value'],
    ['And Expression', '=AND(1=0,3=4)'],
  ], [
    [S, S],
    [S, B],
  ]);
}

function test_convertFormulas_singleFormulaDate() {
  return runConversionTest('SingleFormulaDate', [
    ['Formula', 'Value'],
    ['Date', '=DATE(2026, 12, 25)'],
  ], [
    [S, S],
    [S, D],
  ]);
}

function test_convertFormulas_singleFormulaBlank() {
  return runConversionTest('SingleFormulaBlank', [
    ['Formula', 'Value'],
    ['Blank', '=IF(1=2,"NO","")'],
  ], [
    [S, S],
    [S, S],
  ]);
}

function test_convertFormulas_multipleFormulas() {
  return runConversionTest('MultipleFormulas', [
    ['Formula', 'Value'],
    ['Logical: AND', '=AND(1+1=2, 5>20)'],
    ['Logical: FALSE', '=FALSE()'],
    ['Logical: IF', '=IF(1=2,"NO","")'],
    ['Logical: IFERROR', '=IFERROR(1/0, 0)'],
    ['Logical: IFNA', '=IFNA(NA(), 0)']
  ], [
    [S, S],
    [S, B],
    [S, B],
    [S, S],
    [S, N],
    [S, N],
  ]);
}

function test_convertFormulas_formattingFormulas() {
  return runConversionTest('FormattingFormulas', [
    ['Format', 'Value'],
    ['Currency', '=SUM(1000+500)'],
    ['Percentages', '=1/4'],
    ['Dates', '=DATE(2026,9,27)'],
    ['Decimal Places', '=1/3'],
  ], [
    [S, S],
    [S, N],
    [S, N],
    [S, D],
    [S, N],
  ], {
    formats: { B2: '$#,##0.00', B3: '0.00%', B4: 'mm/dd/yyyy', B5: '0.000' },
  });
}

function test_convertFormulas_errorFormulas() {
  return runConversionTest('ErrorFormulas', [
    ['Format', 'Value'],
    ['Illegal Division', '=DIVIDE(1,0)'],
    ['One Row', 1],
    ['Zero Row', 0],
    ['If-Error Illegal Division Wrapper', '=IFERROR(B3/B4, 0)'],
  ], [
    [S, S],
    [S, S], // the error becomes the text "#DIV/0!"
    [S, N],
    [S, N],
    [S, N],
  ]);
}

function test_convertFormulas_singleCellFormula() {
  return runConversionTest('EmptySheet', [
    ['=CONCAT("John"," Smith")'],
  ], [
    [S],
  ], { header: false, shrink: true });
}

function test_convertFormulas_dependencies() {
  return runConversionTest('Dependencies', [
    ['Format', 'Value'],
    ['Number One', 10],
    ['Formula One (Uses Number One)', '=SUM(B2, 50)'],
    ['Formula Two (Uses Number Two)', '=SUM(B5, 50)'],
    ['Number Two', 500],
    ['Formula Three (Uses Formula One and Two)', '=SUM(B3, B4)'],
  ], [
    [S, S],
    [S, N],
    [S, N],
    [S, N],
    [S, N],
    [S, N],
  ]);
}

function test_convertFormulas_changingData() {
  return runConversionTest('ChangingData', [
    ['Format', 'Value'],
    ['Data', '=TODAY()'],
  ], [
    [S, S],
    [S, D],
  ]);
}

function test_convertFormulas_multipleCells() {
  return runConversionTest('MultipleCells', [
    ['Format', 'Value'],
    [1, 1],
    [2, 2],
    [3, 3],
    ['Result', '=ARRAYFORMULA(A2:A4 * B2:B4)'],
  ], [
    [S, S],
    [N, N],
    [N, N],
    [N, N],
    [S, N],
    [null, N], // B6 and B7 are spilled cells
    [null, N],
  ]);
}

// [function, formula, expected type]
const FUNCTION_CASES = [
  // Logical
  ['IFS',        '=IFS(1=2,"a",2=2,"b")', S],
  ['NOT',        '=NOT(1=2)', B],
  ['OR',         '=OR(1=2,2=2)', B],
  ['SWITCH',     '=SWITCH(2,1,"a",2,"b")', S],
  ['TRUE',       '=TRUE()', B],
  ['XOR',        '=XOR(1=1,2=3)', B],
  // Math
  ['SUMIF',      '=SUMIF({1;2;3},">1")', N],
  ['SUMPRODUCT', '=SUMPRODUCT({1;2},{3;4})', N],
  ['PRODUCT',    '=PRODUCT(2,3)', N],
  ['ABS',        '=ABS(-5)', N],
  ['ROUND',      '=ROUND(2.567,1)', N],
  ['ROUNDUP',    '=ROUNDUP(2.1,0)', N],
  ['ROUNDDOWN',  '=ROUNDDOWN(2.9,0)', N],
  ['INT',        '=INT(2.9)', N],
  ['MOD',        '=MOD(7,3)', N],
  ['POWER',      '=POWER(2,3)', N],
  ['SQRT',       '=SQRT(9)', N],
  ['CEILING',    '=CEILING(2.1)', N],
  ['FLOOR',      '=FLOOR(2.9)', N],
  // Statistical
  ['AVERAGE',    '=AVERAGE(1,2,3)', N],
  ['AVERAGEIF',  '=AVERAGEIF({1;2;3},">1")', N],
  ['COUNT',      '=COUNT(1,"a",2)', N],
  ['COUNTA',     '=COUNTA(1,"a",2)', N],
  ['COUNTIF',    '=COUNTIF({1;2;3},">1")', N],
  ['COUNTIFS',   '=COUNTIFS({1;2;3},">1")', N],
  ['MAX',        '=MAX(1,5,3)', N],
  ['MIN',        '=MIN(1,5,3)', N],
  ['MEDIAN',     '=MEDIAN(1,5,3)', N],
  // Text
  ['CONCATENATE','=CONCATENATE("a","b")', S],
  ['LEFT',       '=LEFT("Jane",2)', S],
  ['RIGHT',      '=RIGHT("Jane",2)', S],
  ['MID',        '=MID("Jane",2,2)', S],
  ['LEN',        '=LEN("Jane")', N],
  ['LOWER',      '=LOWER("JANE")', S],
  ['UPPER',      '=UPPER("jane")', S],
  ['PROPER',     '=PROPER("jane doe")', S],
  ['TRIM',       '=TRIM("  a  b ")', S],
  ['SUBSTITUTE', '=SUBSTITUTE("a-b","-","+")', S],
  ['TEXT',       '=TEXT(5,"0.00")', S],
  ['TEXTJOIN',   '=TEXTJOIN(",",TRUE,"a","b")', S],
  ['VALUE',      '=VALUE("5")', N],
  ['REPT',       '=REPT("a",3)', S],
  ['FIND',       '=FIND("a","bab")', N],
  ['SEARCH',     '=SEARCH("A","bab")', N],
  // Date
  ['YEAR',       '=YEAR(DATE(2026,1,2))', N],
  ['MONTH',      '=MONTH(DATE(2026,1,2))', N],
  ['DAY',        '=DAY(DATE(2026,1,2))', N],
  ['DATEDIF',    '=DATEDIF(DATE(2026,1,1),DATE(2026,3,1),"M")', N],
  ['WEEKDAY',    '=WEEKDAY(DATE(2026,1,2))', N],
  // Lookup
  ['VLOOKUP',    '=VLOOKUP(2,{1,"a";2,"b"},2,FALSE)', S],
  ['HLOOKUP',    '=HLOOKUP(2,{1,2;"a","b"},2,FALSE)', S],
  ['XLOOKUP',    '=XLOOKUP(2,{1;2},{"a";"b"})', S],
  ['INDEX',      '=INDEX({1;2;3},2)', N],
  ['MATCH',      '=MATCH(2,{1;2;3},0)', N],
  // Info
  ['ISBLANK',    '=ISBLANK("")', B],
  ['ISNUMBER',   '=ISNUMBER(5)', B],
  ['ISTEXT',     '=ISTEXT("a")', B],
  ['ISERROR',    '=ISERROR(1/0)', B],
  ['ISNA',       '=ISNA(NA())', B],
  // Operator
  ['ADD',        '=ADD(1,2)', N],
  ['MINUS',      '=MINUS(5,2)', N],
  ['MULTIPLY',   '=MULTIPLY(2,3)', N],
  // Date-returning
  ['EDATE',      '=EDATE(DATE(2026,1,31),1)', D],
  ['EOMONTH',    '=EOMONTH(DATE(2026,1,15),0)', D],
];

// Already covered by the earlier tests above
const ALREADY_TESTED = ['AND', 'FALSE', 'IF', 'IFERROR', 'IFNA', 'SUM', 'DATE', 'TODAY',
                        'CONCAT', 'DIVIDE', 'ARRAYFORMULA', 'NA', 'SUMIFS', 'AVERAGEIFS'];

function test_convertFormulas_approvedFunctions() {
  const rows  = [['Function', 'Value'], ...FUNCTION_CASES.map(([name, formula]) => [name + '()', formula])];
  const types = [[S, S], ...FUNCTION_CASES.map(([, , type]) => [S, type])];

  // Sheets infers a date format for these two because their arguments are dates.
  // Force a plain number format so the test checks the value, not the display.
  const formats = {};
  ['DATEDIF', 'WEEKDAY'].forEach(name => {
    const row = FUNCTION_CASES.findIndex(c => c[0] === name) + 2; // +1 for header, +1 for 1-based rows
    formats['B' + row] = '0';
  });

  return runConversionTest('ApprovedFunctions', rows, types, { formats });
}

/** Fails if a function is approved but has no test, so "approved = tested" stays true. */
function test_everyApprovedFunctionHasATest() {
  const covered = new Set([...ALREADY_TESTED, ...FUNCTION_CASES.map(c => c[0])]);
  const missing = APPROVED_FUNCTIONS.filter(name => !covered.has(name));
  if (missing.length) Logger.log(`FAIL: approved but untested: ${missing.join(', ')}`);
  return missing.length === 0;
}

function test_convertFormulas_rangeCriteria() {
  return runConversionTest('RangeCriteria', [
    ['Data', 'Function', 'Value'],
    [1, 'SUMIFS()',     '=SUMIFS(A2:A4,A2:A4,">1")'],
    [2, 'AVERAGEIFS()', '=AVERAGEIFS(A2:A4,A2:A4,">1")'],
    [3, '', ''],
  ], [
    [S, S, S],
    [N, S, N],
    [N, S, N],
    [N, null, null],
  ]);
}

function test_convertFormulas_textThatLooksLikeNumbers() {
  const sheet = getOrCreateSheet('LooksLikeNumbers');
  const range = sheet.getRange('A1:A5');
  range.setValues([["'02134"], ["'1-2"], ["'1E5"], ["'=not a formula"], ["'-5 pending"]]);

  const before = range.getValues();
  if (before[3][0] !== '=not a formula') {
    Logger.log(`FAIL: setup did not produce literal text, got "${before[3][0]}"`);
    return false;
  }

  convertFormulas(sheet);
  const after = sheet.getRange('A1:A5').getValues();

  const ok = JSON.stringify(before) === JSON.stringify(after);
  if (!ok) Logger.log(`FAIL: before ${JSON.stringify(before)} after ${JSON.stringify(after)}`);
  return ok;
}

/* ---------- Empty Cell Tests ---------- */

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

/* ---------- Trim Spaces Tests ---------- */

/**
 * Tests that leading and trailing spaces are removed, while spaces
 * within meaningful content are preserved.
 */
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

/**
 * Tests that leading and trailing tabs (including repeated tabs and tabs
 * mixed with spaces) are removed, while tabs within content are preserved.
 */
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

/**
 * Tests that leading and trailing new lines (including repeated ones and
 * ones mixed with spaces) are removed, while new lines within content
 * are preserved.
 */
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

/**
 * Tests that non-breaking and other Unicode space characters (no-break,
 * narrow no-break, figure, ideographic) are removed at the edges of a cell,
 * while those within meaningful content are preserved.
 */
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

/**
 * Tests that zero-width and other invisible Unicode characters (zero-width
 * space, ZWNJ, ZWJ, word joiner, BOM, Braille blank, Hangul filler) are
 * removed at the edges of a cell, while invisible characters within
 * meaningful content are preserved.
 */
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

/**
 * Tests that only the edges of a cell are trimmed, and that runs of
 * spaces inside the text are left untouched.
 */
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

/**
 * Tests that numbers, booleans, and dates are left unchanged, and that
 * 0 and false are not mistaken for blank values.
 */
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

/**
 * Tests that emoji variation selectors (such as U+FE0F) at the end of
 * a cell are preserved, while the padding around them is removed.
 */
function test_TrimKeepsEmojiSelector() {
  return runTrimTest('TrimKeepsEmojiSelector', [
    ['ID', 'Name'],
    [' Jane \u2764\uFE0F ', 'Jane \u{1F44D} '],   // variation selector at the end
  ], [
    ['ID', 'Name'],
    ['Jane \u2764\uFE0F', 'Jane \u{1F44D}'],      // padding gone, emoji intact
  ]);
}

/**
 * Tests that trimming is idempotent: running trimCells a second time
 * produces no further changes.
 */
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

/**
 * Tests that text which would become a live formula once trimmed is
 * skipped and left exactly as written, that skipped cells are counted,
 * and that normal cells in the same sheet are still trimmed.
 */
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

/**
 * Tests that when formula creation is allowed, trimmed text that uses an
 * approved function becomes a live formula and evaluates to its result.
 */
function test_TrimAllowsApprovedFormulaCreation() {
  const sheet = getOrCreateSheet('TrimApprovedFormula');

  sheet.getRange('A1').setNumberFormat('@');
  sheet.getRange('A1').setValue('  =SUM(1,2)  ');

  trimCells(sheet, [], true);

  const value = sheet.getRange('A1').getValue();

  return value === 3;
}

/**
 * Tests that when formula creation is allowed, trimmed text that uses an
 * unapproved function (such as IMPORTXML) throws an error mentioning
 * "unapproved" instead of creating the formula.
 */
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