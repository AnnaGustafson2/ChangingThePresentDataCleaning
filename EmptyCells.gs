/**
 * Characters considered invisible or whitespace when identifying blank cells.
 *
 * Includes standard whitespace as well as zero-width characters, directional
 * marks, Unicode fillers, variation selectors, and other invisible characters.
 */
const INVISIBLE = '\\s\\u00AD\\u034F\\u061C\\u115F\\u1160\\u17B4\\u17B5\\u180B-\\u180E' +
                  '\\u200B-\\u200F\\u202A-\\u202E\\u2060-\\u206F\\u2800\\u3164\\uFE00-\\uFE0F\\uFEFF\\uFFA0';

/**
 * Regular expression used to determine whether a cell contains no visible
 * content. A cell may contain invisible characters and at most apostrophes.
 */
const BLANK_REGEX = new RegExp(`^[${INVISIBLE}'\\u2018\\u2019]*$`);

/**
 * Cleans empty cells and removes unused rows and columns from a sheet.
 *
 * 1. Clears cells containing only whitespace, invisible characters, or apostrophes.
 * 2. Deletes fully empty rows above and inside the data (Cases 1.1, 1.2).
 * 3. Deletes fully empty columns left of and inside the data (Cases 2.1, 2.2).
 *    Headed columns with no data are kept and logged.
 * 4. Trims unused rows below and columns right of the data (Cases 1.3, 2.3).
 *
 * Row 1 is treated as the header row after empty rows are removed.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to clean.
 */
function emptyCellClean(sheet) {
  trimWhitespaceCells(sheet);
  SpreadsheetApp.flush();
  deleteEmptyRows(sheet);          // 1.1 above, 1.2 middle
  deleteEmptyColumns(sheet);       // 2.1 left, 2.2 middle
  SpreadsheetApp.flush();
  deleteEmptyRowsBelow(sheet);     // 1.3 below
  deleteEmptyColumnsRight(sheet);  // 2.3 right
}

/**
 * Determines whether a cell contains no meaningful content.
 *
 * A cell is considered blank if it is an empty string, null, or a string
 * containing only whitespace, invisible characters, or apostrophes.
 *
 * @param {*} cell The cell value to check.
 * @returns {boolean} True if the cell is considered blank; false otherwise.
 */
function isBlank(cell) {
  return cell === '' || cell === null ||
         (typeof cell === 'string' && BLANK_REGEX.test(cell));
}

/**
 * Converts a 1-based column number to its spreadsheet letter name.
 *
 * @param {number} n The 1-based column number.
 * @returns {string} The corresponding column letter.
 *
 * @example
 * colToLetter(1);  // 'A'
 * colToLetter(27); // 'AA'
 */
function colToLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Clears cells that contain only whitespace or invisible characters.
 *
 * Cells containing other content are not rewritten. Apostrophe-only cells
 * are also cleared because they are treated as having no visible content.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to clean.
 */
function trimWhitespaceCells(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  const toClear = [];
  values.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (isBlank(cell)) {
        toClear.push(colToLetter(startCol + c) + (startRow + r));
      }
    });
  });

  if (toClear.length > 0) {
    sheet.getRangeList(toClear).clearContent();
  }
}

/**
 * Groups sorted 1-based indices into contiguous runs.
 * @param {number[]} indices Sorted ascending.
 * @returns {Array<{start: number, count: number}>}
 */
function groupRuns(indices) {
  const runs = [];
  indices.forEach(i => {
    const last = runs[runs.length - 1];
    if (last && i === last.start + last.count) last.count++;
    else runs.push({ start: i, count: 1 });
  });
  return runs;
}

/**
 * Deletes fully empty rows above and inside the data (Cases 1.1 and 1.2).
 * Runs are deleted bottom-to-top so row numbers don't shift mid-loop.
 * Does nothing if every row is empty, since Sheets requires at least one row.
 */
function deleteEmptyRows(sheet) {
  const values = sheet.getDataRange().getValues();
  const empty = [];
  values.forEach((row, r) => {
    if (row.every(cell => cell === '')) empty.push(r + 1);
  });

  if (empty.length === values.length) return;
  groupRuns(empty).reverse().forEach(run => sheet.deleteRows(run.start, run.count));
}

/**
 * Deletes fully empty columns left of and inside the data (Cases 2.1 and 2.2).
 * Row 1 is treated as the header row (call after deleteEmptyRows).
 *
 * @param {boolean} [keepHeaded=true] Keep columns that have a header but no data.
 *     Their header gives them meaning, so they are logged for review instead.
 */
function deleteEmptyColumns(sheet, keepHeaded = true) {
  const values = sheet.getDataRange().getValues();
  const numCols = values[0].length;
  const remove = [];

  for (let c = 0; c < numCols; c++) {
    const wholeColumnEmpty = values.every(row => row[c] === '');
    const headed = values[0][c] !== '';
    const bodyEmpty = values.slice(1).every(row => row[c] === '');

    if (wholeColumnEmpty) {
      remove.push(c + 1);
    } else if (headed && bodyEmpty) {
      if (keepHeaded) Logger.log(`[${sheet.getName()}] Headed column with no data kept: ${colToLetter(c + 1)}`);
      else remove.push(c + 1);
    }
  }

  if (remove.length === numCols) return;
  groupRuns(remove).reverse().forEach(run => sheet.deleteColumns(run.start, run.count));
}

/**
 * Deletes all unused rows below the sheet's last populated row.
 * Retains at least one row even when the sheet contains no content.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to modify.
 */
function deleteEmptyRowsBelow(sheet) {
  const maxRows = sheet.getMaxRows();
  const keep = Math.max(sheet.getLastRow(), 1);

  if (maxRows > keep) {
    sheet.deleteRows(keep + 1, maxRows - keep);
  }
}

/**
 * Deletes all unused columns to the right of the sheet's last populated column.
 * Retains at least one column even when the sheet contains no content.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to modify.
 */
function deleteEmptyColumnsRight(sheet) {
  const maxColumns = sheet.getMaxColumns();
  const keep = Math.max(sheet.getLastColumn(), 1);

  if (maxColumns > keep) {
    sheet.deleteColumns(keep + 1, maxColumns - keep);
  }
}

/**
 * Determines whether a cell contains any visible content.
 *
 * Apostrophes are removed before checking so that a cell containing only
 * an apostrophe is treated as having no visible content.
 *
 * Non-string values such as numbers, booleans, and Dates are considered
 * meaningful content.
 *
 * @param {*} cell The cell value to check.
 * @returns {boolean} True if the cell contains visible content; false otherwise.
 */
function hasVisibleContent(cell) {
  if (typeof cell !== 'string') return true;
  return /[\p{L}\p{N}\p{P}\p{S}]/u.test(
    cell.replace(/['\u2018\u2019]/g, '')
  );
}

/**
 * Finds cells that contain only whitespace or other non-visible characters.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @returns {string[]} An array of A1-style cell addresses containing
 *     whitespace-only content.
 */
function findWhitespaceCells(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const found = [];

  values.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (cell !== '' && !hasVisibleContent(cell)) {
        found.push(colToLetter(startCol + c) + (startRow + r));
      }
    });
  });

  return found;
}

/**
 * Checks whether a sheet contains any whitespace-only cells.
 *
 * Logs the addresses of any whitespace-only cells found.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to validate.
 * @returns {number} The number of whitespace-only cells found.
 */
function validateNoWhitespaceCells(sheet) {
  const found = findWhitespaceCells(sheet);

  if (found.length > 0) {
    Logger.log(
      `[${sheet.getName()}] Whitespace-only cells found: ${found.join(', ')}`
    );
  }

  return found.length;
}