/**
 * Converts all formulas in a sheet to their current calculated values.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *     The sheet whose formulas should be converted to values.
 */
function convertFormulas(sheet) {
  const range = sheet.getDataRange();
  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}

const APPROVED_FUNCTIONS = ['AND', 'SUM', 'AVERAGE', 'IF', 'CONCAT', 'DATE', 'DIVIDE', 'IFERROR', 'ARRAYFORMULA', 'TODAY'];

/**
 * Finds and logs formulas containing functions that are not approved.
 *
 * Each formula is checked for function names and compared against
 * APPROVED_FUNCTIONS. Any unapproved functions are logged with the
 * sheet name, cell address, and formula.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *     The sheet whose formulas should be checked.
 * @returns {number} The number of cells containing unapproved functions.
 */
function flagUnapprovedFormulas(sheet) {
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();
  let violations = 0;

  formulas.forEach((row, r) => {
    row.forEach((formula, c) => {
      if (!formula) return;

      const bad = findUnapprovedFunctions(formula);
      if (bad.length === 0) return;

      const cell = range.getCell(r + 1, c + 1);
      Logger.log(`[${sheet.getName()}] Unapproved ${bad.join(', ')} in ${cell.getA1Notation()}: ${formula}`);
      violations++;
    });
  });

  return violations;
}

/**
 * Extracts function names from a formula and identifies those
 * that are not included in APPROVED_FUNCTIONS.
 *
 * Text enclosed in double quotes is removed before searching so
 * that text containing parentheses is not mistaken for a function.
 *
 * @param {string} formula The formula to inspect.
 * @returns {string[]} An array of unapproved function names.
 *     Returns an empty array if all detected functions are approved.
 */
function findUnapprovedFunctions(formula) {
  const noStrings = formula.replace(/"(?:[^"]|"")*"/g, '""'); // ignore text in quotes
  const names = [...noStrings.matchAll(/([A-Z][A-Z0-9_.]*)\s*\(/gi)].map(m => m[1].toUpperCase());
  return names.filter(name => !APPROVED_FUNCTIONS.includes(name));
}