/** Adds a "Data Cleaning" menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Data Cleaning')
    .addItem('Clean...', 'showCleaningDialog')
    .addToUi();
}

/** Shows the checkbox dialog. */
function showCleaningDialog() {
  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: Arial, sans-serif; font-size: 18px; padding: 8px; }
  label { display: block; margin: 10px 0 0; }
  small { display: block; margin-left: 22px; color: #666; }
  button { margin-top: 14px; padding: 6px 18px; }
  #result { margin-top: 12px; white-space: pre-wrap; max-height: 300px; overflow-y: auto; }
</style>

<b>Select what to clean:</b>
<label><input type="checkbox" id="formulas" checked> Convert formulas to values</label>
<label><input type="checkbox" id="emptyCells" checked> Remove empty rows and columns</label>
<label><input type="checkbox" id="trim" checked> Trim leading/trailing whitespace in text</label>
<div style="margin-left: 20px; margin-bottom: 6px;">
    <label><input type="checkbox" id="allowTrimmedFormulas"> Allow trimming that creates formulas (new formulas will be converted to values)</label>
</div>
<p><small style="margin:0">Applies to every sheet. Changes can only be undone through version history.</small></p>

<button id="run" onclick="run()">Run</button>
<div id="result"></div>

<script>
  function run() {
    const options = {
      formulas: document.getElementById('formulas').checked,
      emptyCells: document.getElementById('emptyCells').checked,
      trim: document.getElementById('trim').checked,
      allowTrimmedFormulas:
        document.getElementById('allowTrimmedFormulas').checked,
    };

    const btn = document.getElementById('run');
    const result = document.getElementById('result');

    btn.disabled = true;
    result.textContent = 'Running...';

    google.script.run
      .withSuccessHandler(function (text) {
        result.textContent = text;
        btn.disabled = false;
      })
      .withFailureHandler(function (err) {
        result.textContent = 'Error: ' + err.message;
        btn.disabled = false;
      })
      .runSelectedCleaning(options);
  }
</script>`)
    .setWidth(600)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Data Cleaning');
}

/** Converts a column letter to its 1-based number ('A' -> 1, 'AA' -> 27). */
function letterToCol(letters) {
  return letters.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
}

/**
 * Parses a comma-delimited list of column letters ("c, F, aa").
 * @param {string} text The text entered in the dialog.
 * @returns {{columns: number[], invalid: string[]}} Sorted, de-duplicated 1-based
 *     column numbers, plus any entries that are not valid column letters.
 */
function parseColumnList(text) {
  const columns = new Set();
  const invalid = [];

  String(text || '').split(',')
    .map(s => s.trim().toUpperCase())
    .filter(Boolean)
    .forEach(entry => {
      if (/^[A-Z]{1,3}$/.test(entry)) columns.add(letterToCol(entry));
      else invalid.push(entry);
    });

  return { columns: [...columns].sort((a, b) => a - b), invalid };
}

/** Called by the dialog. Validates the input, then returns the summary text to display. */
function runSelectedCleaning(options) {
  if (!['formulas', 'emptyCells', 'trim'].some(k => options[k] === true)) {
    return 'Nothing selected.';
  }

  return cleanSheet(options);
}

/** Formats a sheet name for use in a reference, quoting it if needed. */
function sheetRef(sheet) {
  const name = sheet.getName();
  return /^[A-Za-z_]\w*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/**
 * Runs the selected cleaning steps on every sheet, always in a safe order:
 * formulas, empty cells, trimming, emails. Calling it with no argument runs
 * everything except emails (which need columns).
 *
 * @param {{formulas?: boolean, emptyCells?: boolean, trim?: boolean,
 *          emails?: boolean, emailColumns?: number[]}} [options]
 *     emailColumns holds 1-based column numbers.
 * @returns {string} A summary of anything flagged, with cell locations.
 */
function cleanSheet(options = {}) {
  const {
    formulas = true,
    emptyCells = true,
    trim = true,
    allowTrimmedFormulas = false,
  } = options;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const formulaCells = [];
  const whitespaceCells = [];
  const skippedCells = [];

  /* ---------- PASS 1: VALIDATE EXISTING FORMULAS ---------- */

  if (formulas) {
    ss.getSheets().forEach(sheet => {
      const ref = sheetRef(sheet);

      const found = [];

      flagUnapprovedFormulas(sheet, found);

      found.forEach(a1 => {
        formulaCells.push(`${ref}!${a1}`);
      });
    });

    if (formulaCells.length) {
      Logger.log(
        `ABORTED: ${formulaCells.length} formula(s) used unapproved functions:\n` +
        formulaCells.join('\n')
      );

      const shown = formulaCells.slice(0, 25);
      const more = formulaCells.length - shown.length;
      return (
        'Cleaning was cancelled because unapproved formulas were found.\n\n' +
        shown.join('\n') +
        (more > 0
          ? `\n\n...and ${more} more (see Logger).`
          : '')
      );
    }
  }

  /* ---------- PASS 2: CLEAN ---------- */

  ss.getSheets().forEach(sheet => {
    const ref = sheetRef(sheet);

    if (emptyCells) {
      emptyCellClean(sheet);

      findWhitespaceCells(sheet)
        .forEach(a => whitespaceCells.push(`${ref}!${a}`));
    }

    if (trim) {
      const skipped = [];

      trimCells(
        sheet,
        skipped,
        allowTrimmedFormulas
      );

      skipped.forEach(a => {
        skippedCells.push(`${ref}!${a}`);
      });
    }

    if (formulas) {
      convertFormulas(sheet);
    }
  });

  /* ---------- SUMMARY ---------- */

  const summary = [];
  const MAX_SHOWN = 15;

  const pass = message => Logger.log(message);

  const flag = (message, locations) => {
    Logger.log(
      `${message}\n  ${locations.join('\n  ')}`
    );

    const shown = locations.slice(0, MAX_SHOWN);

    summary.push(
      `${message}\n  ${shown.join(', ')}` +
      (locations.length > MAX_SHOWN
        ? `, and ${locations.length - MAX_SHOWN} more (see Logger)`
        : '')
    );
  };

  if (formulas) {
    pass('PASS: All formulas use approved functions');
  }

  if (emptyCells) {
    if (whitespaceCells.length === 0) {
      pass('PASS: No whitespace-only cells remaining');
    } else {
      flag(
        `FLAG: ${whitespaceCells.length} whitespace-only cell(s) remain:`,
        whitespaceCells
      );
    }
  }

  if (trim) {
    if (skippedCells.length === 0) {
      pass('PASS: No trimmed text would have become a formula');
    } else {
      flag(
        `FLAG: ${skippedCells.length} cell(s) skipped because trimming would create a formula:`,
        skippedCells
      );
    }
  }

  return summary.length
    ? 'Finished, but these need review:\n\n' +
        summary.join('\n\n')
    : 'Finished. No issues flagged.';
}

/* ---------- FORMULA CONVERSION ---------- */

/**
 * Converts all formulas in a sheet to their current calculated values.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *     The sheet whose formulas should be converted to values.
 */
function convertFormulas(sheet) {
  const range = sheet.getDataRange();
  const before = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  SpreadsheetApp.flush();

  const after = range.getValues();
  const lostDates = [];
  before.forEach((row, r) => {
    row.forEach((value, c) => {
      if (value instanceof Date && !(after[r][c] instanceof Date)) {
        lostDates.push(colToLetter(startCol + c) + (startRow + r));
      }
    });
  });

  if (lostDates.length > 0) {
    sheet.getRangeList(lostDates).setNumberFormat('m-d-yyyy');
  }
}

const APPROVED_BY_CATEGORY = {
  Logical:     ['AND', 'FALSE', 'IF', 'IFERROR', 'IFNA', 'IFS', 'NOT', 'OR', 'SWITCH', 'TRUE', 'XOR'],
  Math:        ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'PRODUCT', 'ABS', 'ROUND', 'ROUNDUP',
                'ROUNDDOWN', 'INT', 'MOD', 'POWER', 'SQRT', 'CEILING', 'FLOOR'],
  Statistical: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS',
                'MAX', 'MIN', 'MEDIAN'],
  Text:        ['CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER', 'PROPER', 'TRIM',
                'SUBSTITUTE', 'TEXT', 'TEXTJOIN', 'VALUE', 'REPT', 'FIND', 'SEARCH'],
  Date:        ['DATE', 'TODAY', 'YEAR', 'MONTH', 'DAY', 'DATEDIF', 'WEEKDAY', 'EDATE', 'EOMONTH'],
  Lookup:      ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH'],
  Info:        ['NA', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISNA'],
  Operator:    ['CONCAT', 'DIVIDE', 'ADD', 'MINUS', 'MULTIPLY'],
  Array:       ['ARRAYFORMULA'],
};

const APPROVED_FUNCTIONS = Object.values(APPROVED_BY_CATEGORY).flat();

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
function flagUnapprovedFormulas(sheet, found = []) {
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();

  let violations = 0;

  formulas.forEach((row, r) => {
    row.forEach((formula, c) => {
      if (!formula) return;

      const bad = findUnapprovedFunctions(formula);

      if (bad.length === 0) return;

      const a1 =
        colToLetter(range.getColumn() + c) +
        (range.getRow() + r);

      found.push(a1);

      Logger.log(
        `[${sheet.getName()}] Unapproved ${bad.join(', ')} in ${a1}: ${formula}`
      );

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

/* ---------- EMPTY ROWS AND COLUMNS ---------- */

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
  trimWhitespaceCells(sheet); // Could potential not be that often and might slow down a lot
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

/* ---------- TRIM CELLS ---------- */

// Same invisible set, minus variation selectors (\uFE00-\uFE0F), which can legitimately
// trail an emoji (e.g. ❤️ is U+2764 + U+FE0F). Stripping them would change the presentation.
const EDGE_INVISIBLE = INVISIBLE.replace('\\uFE00-\\uFE0F', '');
const EDGE_REGEX = new RegExp(
    `^(?:[${EDGE_INVISIBLE}]+)|(?:[${EDGE_INVISIBLE}]+)$`,
    'g'
);

/**
 * Removes leading and trailing whitespace/invisible characters from text cells
 * ("   Jane  " -> "Jane"). Inner spacing is left alone. Non-text cells and
 * cells that are already trimmed are never rewritten.
 *
 * If trimming would leave text starting with "=", the cell is skipped and
 * logged, because writing it would create a live formula. The log lists any
 * unapproved functions it uses.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to clean.
 * @param {string[]} [skippedCells] Optional list that receives the A1 address
 *     of each skipped cell.
 * @returns {number} The number of cells skipped because they would become formulas.
 */
function trimCells(
  sheet,
  skippedCells = [],
  allowTrimmedFormulas = false
) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  const createdFormulas = [];

  let skipped = 0;
  let changed = false;

  values.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string') return;

      const trimmed = cell.replace(EDGE_REGEX, '');

      if (trimmed === cell) return;

      const address =
        colToLetter(startCol + c) + (startRow + r);

      // Would become a formula
      if (trimmed.startsWith('=')) {

        if (!allowTrimmedFormulas) {
          const bad = findUnapprovedFunctions(trimmed);

          Logger.log(
            `[${sheet.getName()}] Trim skipped at ${address}: result would become a formula` +
            (bad.length
              ? ` (unapproved: ${bad.join(', ')})`
              : '')
          );

          skippedCells.push(address);
          skipped++;
          return;
        }

        // Defer writing until validation passes
        createdFormulas.push({
          row: r,
          col: c,
          address,
          value: trimmed,
        });

        return;
      }

      // Normal trim can be applied immediately in memory
      values[r][c] = trimmed;
      changed = true;
    });
  });

  //
  // Validate deferred formulas before applying them
  //
  if (allowTrimmedFormulas && createdFormulas.length) {
    const violations = [];

    createdFormulas.forEach(({ address, value }) => {
      const bad = findUnapprovedFunctions(value);

      if (bad.length) {
        violations.push(
          `${address} (${bad.join(', ')})`
        );
      }
    });

    if (violations.length) {
      const shown = violations.slice(0, 25);
      const more = violations.length - shown.length;

      throw new Error(
        'Trimming would create formulas that use unapproved functions:\n\n' +
        shown.join('\n') +
        (more > 0
          ? `\n\n...and ${more} more (see Logger).`
          : '')
      );
    }

    createdFormulas.forEach(({ row, col, value }) => {
      values[row][col] = value;
      changed = true;
    });
  }

  // Single write to the sheet
  if (changed) {
    range.setValues(values);
  }

  return skipped;
}