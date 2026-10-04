function cleanSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let formulaViolations = 0;
  let whitespaceCells = 0;

  ss.getSheets().forEach(sheet => {
    // Formula Handling
    formulaViolations += flagUnapprovedFormulas(sheet);
    convertFormulas(sheet);

    // Empty Cell Handling
    emptyCellClean(sheet);
    whitespaceCells += validateNoWhitespaceCells(sheet);
  });

  Logger.log(formulaViolations === 0
    ? 'PASS: All formulas use approved functions'
    : `FLAG: Needs manual verification: ${formulaViolations} formula(s) flagged across the spreadsheet`);

  Logger.log(whitespaceCells === 0
    ? 'PASS: No whitespace-only cells remaining'
    : `FLAG: Needs manual verification: ${whitespaceCells} whitespace-only cell(s) remain across the spreadsheet`);
}