// Shorthand for expected JavaScript types
const S = 'string', N = 'number', B = 'boolean', D = 'object'; // D = Date

/**
 * Runs every test, logs PASS/FAIL (or ERROR if a test throws),
 * then logs a summary. Add new test functions to the array below.
 */
function tests() {
  const testFunctions = [
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
    test_convertFormulas_textThatLooksLikeNumbers,
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
    results.push(checkFormat(cell, formatsBefore[cell], sheet.getRange(cell).getNumberFormat()));
    results.push(checkType(sheet, cell, expected[cell]));
  });
  return results.every(Boolean); // all checks run first, so every failure is logged
}

/* ---------- Tests ---------- */

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
    formats: { B2: '$#,##0.00', B3: '0.00%', B4: 'yyyy-mm-dd', B5: '0.000' },
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