/**
 * Runs all test functions, collects their results, and logs
 * each test's name and PASS/FAIL status for user inspection.
 *
 * Add new test functions to the `testFunctions` array above;
 * no other changes are required.
 */
function tests() {
  const testFunctions = [
    test_convertFormulas_singleFormulaText,
    test_convertFormulas_singleFormulaNumber,
    test_convertFormulas_singleFormulaBoolean,
    test_convertFormulas_singleFormulaDate,
    test_convertFormulas_singleFormulaBlank,
    test_convertFormulas_multipleFormulas,
    test_convertFormulas_formatingFormulas,
    test_convertFormulas_errorFormulas,
    test_convertFormulas_singleCellFormula,
    test_convertFormulas_dependencies,
    test_convertFormulas_changingData,
    test_convertFormulas_multipleCells,
    // add new test_ functions here — nothing else to update
  ];

  testFunctions.forEach(fn => {
    Logger.log(`${fn.name}: ${fn() ? 'PASS' : 'FAIL'}`);
  });
}

/**
 * Finds all formulas currently present in a sheet.
 * An empty array is expected when all formulas have been cleaned.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @returns {Array<{row: number, col: number, formula: string}>}
 *     An array containing the row, column, and formula for each
 *     non-empty formula cell.
 */
function findRemainingFormulas(sheet) {
  const range = sheet.getDataRange();
  const formulas = range.getFormulas(); 
  const remaining = [];

  formulas.forEach((row, r) => {
    row.forEach((formula, c) => {
      if (formula !== '') {
        remaining.push({ row: r + 1, col: c + 1, formula: formula });
      }
    });
  });
  return remaining;
}

/**
 * Checks whether any formulas remain in the sheet.
 * Uses findRemainingFormulas() to identify remaining formulas.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check.
 * @returns {boolean} True if no formulas remain; false otherwise.
 */
function run_formula_test(sheet) {
  const remaining = findRemainingFormulas(sheet);

  if (remaining.length === 0) {
    return true;
  } else {
    Logger.log(`FAIL: formulas remaining: ${JSON.stringify(remaining)}`);
    return false;
  }
}

/**
 * Checks whether a cell's format remained unchanged.
 *
 * @param {string} cell The cell address being checked.
 * @param {string} formatBefore The cell's format before cleaning.
 * @param {string} formatAfter The cell's format after cleaning.
 * @returns {boolean} True if the format is unchanged; false otherwise.
 */
function run_format_test(cell, formatBefore, formatAfter) {
  if (formatBefore === formatAfter) {
    return true;
  } else {
    Logger.log(`FAIL: ${cell} format changed from "${formatBefore}" to "${formatAfter}"`);
    return false;
  }
}

/**
 * Checks whether a cell contains the expected JavaScript value type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet containing the cell.
 * @param {string} cell The cell address being checked.
 * @param {string} expectedType The expected JavaScript type.
 * @returns {boolean} True if the actual type matches the expected type; false otherwise.
 */
function run_type_test(sheet, cell, expectedType) {
  const value = sheet.getRange(cell).getValue();
  const actualType = typeof value;
  if (actualType === expectedType) {
    return true;
  } else {
    Logger.log(`FAIL: ${cell} unexpected type from "${expectedType}" to "${actualType}"`);
    return false;
  }
}

/**
 * Runs the standard checks used to verify formula cleaning.
 *
 * Checks that:
 * 1. No formulas remain in the sheet.
 * 2. Cell number formats remain unchanged.
 * 3. Cell values have the expected JavaScript types.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet being tested.
 * @param {string[]} cells The cell addresses to check.
 * @param {Object<string, string>} formatsBefore The original number format for each cell.
 * @param {string[]} expectedTypes The expected JavaScript type for each cell.
 * @returns {boolean[]} An array containing the result of each check.
 */
function run_standard_checks(sheet, cells, formatsBefore, expectedTypes) {
  const results = [];
  results.push(run_formula_test(sheet));
  cells.forEach(cell => {
    const formatAfter = sheet.getRange(cell).getNumberFormat();
    results.push(run_format_test(cell, formatsBefore[cell], formatAfter));
  });
  cells.forEach((cell, i) => {
    results.push(run_type_test(sheet, cell, expectedTypes[i]));
  });
  return results
}
/**
 * Tests convertFormulas() on a sheet containing a single text formula.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleFormulaText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SingleFormulaText');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Concatenation');
  sheet.getRange('B2').setFormula('=CONCAT("John", " Doe")');

  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'string'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a single number formula.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleFormulaNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SingleFormulaNumber');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Addition');
  sheet.getRange('B2').setFormula('=1+1');

  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'number'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a single boolean formula.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleFormulaBoolean() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SingleFormulaBoolean');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('And Expression');
  sheet.getRange('B2').setFormula('=AND(1=0,3=4)');

  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'boolean'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a single date formula.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleFormulaDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SingleFormulaDate');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Date');
  sheet.getRange('B2').setFormula('=DATE(2026, 12, 25)');

  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'object'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a single 
 * blank string returning formula.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleFormulaBlank() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SingleFormulaBlank');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Blank');
  sheet.getRange('B2').setFormula('=IF(1=2,"NO","")');

  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'string'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a multiple formulas
 * with types number, text, boolean, date, and blank string.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_multipleFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MultipleFormulas');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Formula').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Number (SUM)');
  sheet.getRange('B2').setFormula('=SUM(1+1)');
  sheet.getRange('A3').setValue('String (CONCAT) ');
  sheet.getRange('B3').setFormula('=CONCAT("John", " Doe")');
  sheet.getRange('A4').setValue('Boolean (AND)');
  sheet.getRange('B4').setFormula('=AND(1+1=2, 5>20)');
  sheet.getRange('A5').setValue('Blank');
  sheet.getRange('B5').setFormula('=IF(1=2,"NO","")');
  sheet.getRange('A6').setValue('Date');
  sheet.getRange('B6').setFormula('=DATE(2026, 12, 25)');

  const cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'A6', 'B6'];
  const expectedTypes = ['string', 'string', 'string', 'number', 'string', 
    'string', 'string', 'boolean', 'string', 'string', 'string', 'object'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a multiple formulas 
 * formats such as currency, percentages, dates, decimal places.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_formatingFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormattingFormulas');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Format').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Currency');
  sheet.getRange('B2').setFormula('=SUM(1000+500)').setNumberFormat('$#,##0.00');
  sheet.getRange('A3').setValue('Percentages');
  sheet.getRange('B3').setFormula('=1/4').setNumberFormat('0.00%');
  sheet.getRange('A4').setValue('Dates');
  sheet.getRange('B4').setFormula('=DATE(2026,9,27)').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('A5').setValue('Decimal Places');
  sheet.getRange('B5').setFormula('=1/3').setNumberFormat('0.000');

  const cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5'];
  const expectedTypes = ['string', 'string', 'string', 'number', 'string', 'number', 'string', 'object', 'string', 'number'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a sheet containing a formulas with errors.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_errorFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ErrorFormulas');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Format').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Illegal Division');
  sheet.getRange('B2').setFormula('=DIVIDE(1,0)');
  sheet.getRange('A3').setValue('One Row');
  sheet.getRange('B3').setFormula(1);
  sheet.getRange('A4').setValue('Zero Row');
  sheet.getRange('B4').setFormula(0);
  sheet.getRange('A5').setValue('If-Error Illegal Division Wrapper');
  sheet.getRange('B5').setFormula('=IFERROR(B3/B4, 0)');


  const cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5'];
  const expectedTypes = ['string', 'string', 'string', 'string', 'string', 'number', 'string', 'number', 'string', 'number'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on an single-celled sheet.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_singleCellFormula() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmptySheet');
  sheet.clear(); 
  
  var maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.deleteRows(2, maxRows - 1);
  }
  
  var maxColumns = sheet.getMaxColumns();
  if (maxColumns > 1) {
    sheet.deleteColumns(2, maxColumns - 1);
  }
  sheet.getRange('A1').setFormula('=CONCAT("John"," Smith")');

  const cells = ['A1'];
  const expectedTypes = ['string'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a formulas with dependencies on cells
 * above, below, or dependent on other formulas.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_dependencies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dependencies');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Format').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Number One');
  sheet.getRange('B2').setFormula(10);
  sheet.getRange('A3').setValue('Formula One (Uses Number One)');
  sheet.getRange('B3').setFormula('=SUM(B2, 50)');
  sheet.getRange('A4').setValue('Formula Two (Uses Number Two)');
  sheet.getRange('B4').setFormula('=SUM(B5, 50)');
  sheet.getRange('A5').setValue('Number Two');
  sheet.getRange('B5').setFormula(500);
  sheet.getRange('A6').setValue('Formula Three (Uses Formula One and Two)');
  sheet.getRange('B6').setFormula('=SUM(B3, B4)');

  const cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'A6', 'B6'];
  const expectedTypes = ['string', 'string', 'string', 'number', 'string', 'number', 
    'string', 'number', 'string', 'number', 'string', 'number'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a formulas that rely on current data.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_changingData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangingData');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Format').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue('Data');
  sheet.getRange('B2').setFormula('=TODAY()');


  const cells = ['A1', 'B1', 'A2', 'B2'];
  const expectedTypes = ['string', 'string', 'string', 'object'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}

/**
 * Tests convertFormulas() on a formulas that produce more than plain values
 * in multiple cells.
 *
 * Sets up a test sheet with a formula, runs the cleaning method,
 * and verifies that formulas were removed while formats and value
 * types were preserved.
 *
 * @returns {boolean} True if all standard checks pass; false otherwise.
 */
function test_convertFormulas_multipleCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MultipleCells');
  sheet.clear(); 
  sheet.getRange('A1').setValue('Format').setBackground('#d9ead3');
  sheet.getRange('B1').setValue('Value').setBackground('#d9ead3');
  sheet.getRange('A2').setValue(1);
  sheet.getRange('B2').setValue(1);
  sheet.getRange('A3').setValue(2);
  sheet.getRange('B3').setValue(2);
  sheet.getRange('A4').setValue(3);
  sheet.getRange('B4').setValue(3);
  sheet.getRange('A5').setValue('Result');
  sheet.getRange('B5').setFormula('=ARRAYFORMULA(A2:A4 * B2:B4)');


  const cells = ['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'B6', 'B7'];
  const expectedTypes = ['string', 'string', 'number', 'number', 'number', 'number', 'number', 'number', 'string', 
    'number', 'number', 'number'];
  const formatsBefore = {};
  cells.forEach(cell => formatsBefore[cell] = sheet.getRange(cell).getNumberFormat());

  convertFormulas(sheet);

  const results = run_standard_checks(sheet, cells, formatsBefore, expectedTypes);

  return results.every(r => r === true);
}