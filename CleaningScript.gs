function cleanSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let violations = 0;

  ss.getSheets().forEach(sheet => {
    violations += flagUnapprovedFormulas(sheet);
    convertFormulas(sheet);
  });

  Logger.log(violations === 0
    ? 'All formulas use approved functions'
    : `Needs manual verification: ${violations} formula(s) flagged across the spreadsheet`);
}

function convertFormulas(sheet) {
  const range = sheet.getDataRange();
  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}

const APPROVED_FUNCTIONS = ['AND', 'SUM', 'AVERAGE', 'IF', 'CONCAT', 'DATE', 'DIVIDE', 'IFERROR', 'ARRAYFORMULA', 'TODAY'];

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

function findUnapprovedFunctions(formula) {
  const noStrings = formula.replace(/"(?:[^"]|"")*"/g, '""'); // ignore text in quotes
  const names = [...noStrings.matchAll(/([A-Z][A-Z0-9_.]*)\s*\(/gi)].map(m => m[1].toUpperCase());
  return names.filter(name => !APPROVED_FUNCTIONS.includes(name));
}