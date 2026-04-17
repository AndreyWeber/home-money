/// <reference path="Types.js" />

/**
 * @fileoverview Main entry points and core transaction-processing logic for
 * the home-money Google Apps Script project.
 *
 * Algorithm overview:
 *  1. Read all transaction records from the 'Raw Data' sheet.
 *  2. For each record whose date is ≤ today:
 *     2.1 Look up the budget row/column by symbol. Skip if not found.
 *     2.2 Build a new cell note and calculate the new cell value.
 *     2.3 For planned payments: update background colour and handle overrun.
 *     2.4 For planned payments where cell value = 0 and note sum > 0: skip
 *         and show a warning dialog.
 *     2.5 Write the new cell value, note, and background colour.
 *  3. Records dated in the future are skipped.
 *  4. Successfully processed records are deleted from 'Raw Data'.
 *
 * Known issues:
 *  - Numbers typed in comments are included in planned-payment calculations,
 *    potentially causing false overrun warnings.
 */

// ── Spreadsheet event handlers ───────────────────────────────────────────────

/**
 * Spreadsheet `onOpen` event handler.
 * Creates the 'Transactions' custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Transactions')
    .addItem('Process Transactions', 'processTransactions')
    //.addItem('Show Daily Expenses', 'showDailyExpenses')
    .addItem('Show Transactions History', 'showTransactionsHistorySidebar')
    .addItem('Show Metadata Viewer', 'showMetadataSidebar')
    .addToUi();
}

// ── Dialog / sidebar launchers ───────────────────────────────────────────────

/**
 * Opens the Daily Expenses modal dialog.
 */
function showDailyExpenses() {
  const html = HtmlService
    .createHtmlOutputFromFile('DailyExpenses')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, DAILY_EXPENSES_DATA_TITLE);
}

/**
 * Opens the Transactions History sidebar for the cell currently selected on
 * the Summary Balance tab.
 */
function showTransactionsHistorySidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('TransactionsHistorySidebar')
    .setTitle('Transactions History');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the Metadata Viewer sidebar.
 */
function showMetadataSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('MetadataViewerSidebar')
    .setTitle('Metadata Viewer');

  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Data providers (called from HTML dialogs via google.script.run) ──────────

/**
 * Calculates and returns daily expenses data for the given date.
 * Called from the DailyExpenses HTML dialog.
 *
 * @param  {string}  dateToShow      - Date string to calculate expenses for
 * @param  {boolean} includePlanned  - Whether to include planned-payment rows
 * @returns {DailyExpensesResult} Actual, expected and overrun amounts
 */
function getDailyExpensesData(dateToShow, includePlanned) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawData = getRawDataTransactionObjects();

  const dailyExpensesSumActual = rawData.reduce((sum, row) => {
    if (!row.dateOfTransaction) return sum;
    if (dateAsUtc(row.dateOfTransaction) !== dateAsUtc(new Date(dateToShow))) return sum;
    if (!includePlanned && row.plannedPayment) return sum;
    return sum + row.value;
  }, 0);

  const summaryBalanceSheet = ss.getSheetByName(Sheets.SUMMARY_BALANCE);
  const rowNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
    (value) => value === FREE_DAILY_CASH_DATA_TITLE,
  );
  const colNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getRange(rowNum + 1, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
    isNumber,
  );
  const dailyExpensesSumExpected = summaryBalanceSheet
    .getRange(rowNum, colNum)
    .getCell(1, 1)
    .getValue();

  return {
    sumActual:   `${dailyExpensesSumActual}${CURRENCY_SUFFIX}`,
    sumExpected: `${dailyExpensesSumExpected.toFixed(2)}${CURRENCY_SUFFIX}`,
    overrun: dailyExpensesSumActual > dailyExpensesSumExpected
      ? `${(dailyExpensesSumActual - dailyExpensesSumExpected).toFixed(2)}${CURRENCY_SUFFIX}`
      : '(нет)',
  };
}

/**
 * Returns the transactions-history JSON for the cell currently selected on
 * the Summary Balance tab, or a default sentinel when another tab is active.
 *
 * @returns {TransactionsHistoryResult} History data for the active cell
 */
const getTransactionsHistoryData = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() === Sheets.SUMMARY_BALANCE) {
    const activeCellA1Notation = activeSheet.getActiveCell().getA1Notation();
    const json = ss
      .getSheetByName(Sheets.TRANSACTIONS_HISTORY)
      .getRange(activeCellA1Notation)
      .getValue();

    return { isSummaryBalanceSheet: true, jsonCellA1Notation: activeCellA1Notation, json };
  }

  return { isSummaryBalanceSheet: false, jsonCellA1Notation: null, jsonStr: null };
};

// ── Cell note and background helpers ────────────────────────────────────────

/**
 * Builds an updated cell note by appending the new transaction data to the
 * existing note. Returns the existing note unchanged when noteToAdd is empty
 * and isPlanned is false.
 *
 * @param  {string}           existingNote - Current note on the cell
 * @param  {Date}             dateToAdd    - Transaction date to append
 * @param  {number}           valueToAdd   - Transaction amount to append
 * @param  {string|null}      noteToAdd    - Transaction comment to append
 * @param  {boolean}          isPlanned    - Whether the transaction is a planned payment
 * @returns {string} Updated cell note
 */
function createCellNote(existingNote, dateToAdd, valueToAdd, noteToAdd, isPlanned) {
  if (!noteToAdd && !isPlanned) return existingNote;

  const prefix = existingNote ? `${existingNote}\u000A` : '';

  if (isPlanned) {
    return noteToAdd
      ? `${prefix}${dateToFormattedString(dateToAdd)}: ${valueToAdd} - ${noteToAdd} - потрачено`
      : `${prefix}${dateToFormattedString(dateToAdd)}: ${valueToAdd} - потрачено`;
  }

  return `${prefix}${dateToFormattedString(dateToAdd)}: ${valueToAdd} - ${noteToAdd}`;
}

/**
 * Sums all numeric tokens found in cellNote.
 * Tokens are split on newlines and spaces; non-numeric tokens are ignored.
 *
 * @param  {string|null} cellNote - Cell note text to parse
 * @returns {number} Sum of all numeric tokens
 */
function calculateCellNoteMoneySum(cellNote) {
  if (!cellNote) return 0;
  return cellNote
    .split(/[\n|\s*]/)
    .reduce((sum, token) => (isNumber(token) ? sum + Number(token) : sum), 0);
}

/**
 * Returns the background colour a planned-payment cell should show based on
 * how much of the budgeted amount has been spent according to the cell note.
 *
 * @param  {string|null} cellNote  - Current cell note
 * @param  {number|null} cellValue - Current cell value (the budget amount)
 * @returns {string} Hex colour code from {@link CELL_BACKGROUND_COLOR}
 */
function calculateCellBackgroung(cellNote, cellValue) {
  if (!cellNote || !cellValue) return CELL_BACKGROUND_COLOR.noFill;

  const noteMoneySum = calculateCellNoteMoneySum(cellNote);
  if (noteMoneySum === 0)       return CELL_BACKGROUND_COLOR.notSpent;
  if (noteMoneySum < cellValue) return CELL_BACKGROUND_COLOR.almostSpent;
  return CELL_BACKGROUND_COLOR.spent;
}

// ── Raw data write helpers ───────────────────────────────────────────────────

/**
 * Deletes the 'Raw Data' row whose timestamp matches the given value.
 *
 * @param  {Date} timeStamp - Timestamp of the row to delete
 */
function deleteRawTransaction(timeStamp) {
  const ds = getRawDataSheet();
  const rowNum = findValueIndex(
    flattenArray(ds.getSheetValues(1, 1, ds.getMaxRows(), 1)),
    (value) => +value === +timeStamp,
  );
  if (rowNum !== 0) ds.deleteRow(rowNum);
}

/**
 * Appends rawDataRow to the transaction-history JSON stored in the
 * corresponding cell on the Transactions History sheet.
 *
 * @param  {RawTransaction}                           rawDataRow                - Processed transaction record
 * @param  {GoogleAppsScript.Spreadsheet.Sheet}       summaryBalanceSheet       - Summary Balance sheet reference
 * @param  {GoogleAppsScript.Spreadsheet.Sheet}       transactionsHistorySheet  - Transactions History sheet reference
 */
function addTransationHistoryRow(rawDataRow, summaryBalanceSheet, transactionsHistorySheet) {
  const rowNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
    (value) => value === rawDataRow.symbol.split(' : ')[0],
  );

  const colNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getRange(rowNum, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
    (value) => value.toString() === EMPTY_STRING,
  ) - 1;

  const cell = transactionsHistorySheet.getRange(rowNum, colNum);
  const cellValue = cell.getValue();

  let existing = null;
  if (cellValue) {
    try {
      existing = JSON.parse(cellValue);
    } catch (e) {
      // Cell contains a value that is not valid JSON (e.g. "undefined", partial write).
      // Treat it as no existing data and start a fresh array.
      console.warn(`addTransationHistoryRow: unparseable cell value — starting fresh. Value: "${cellValue}". Error: ${e.message}`);
    }
  }
  const jsonArr = existing
    ? (Array.isArray(existing) ? existing : [existing])
    : [];
  jsonArr.push(rawDataRow);

  cell.setValue(JSON.stringify(jsonArr));
}

// ── Main processing ──────────────────────────────────────────────────────────

/**
 * Processes all pending raw transactions: updates the Summary Balance sheet,
 * writes transaction history entries, deletes processed raw rows, and saves
 * the latest-transaction metadata.
 */
function processTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summaryBalanceSheet      = ss.getSheetByName(Sheets.SUMMARY_BALANCE);
  const transactionsHistorySheet = ss.getSheetByName(Sheets.TRANSACTIONS_HISTORY);

  const rawTransactions = getRawDataTransactionObjects();
  const processedRawTransactions = [];

  for (const rawTransaction of rawTransactions) {
    if (processTransaction(rawTransaction, summaryBalanceSheet)) {
      addTransationHistoryRow(rawTransaction, summaryBalanceSheet, transactionsHistorySheet);
      deleteRawTransaction(rawTransaction.timestamp);
      processedRawTransactions.push(rawTransaction);
    }
  }

  const latest = getRawDataTransactionObjectWithMaxDate(processedRawTransactions);
  setLatestTransactionDate(latest.dateOfTransaction);
  setLatestTransactionName(latest.comment);
  setLatestTransactionSymbol(latest.symbol);
  setLatestTransactionValue(latest.value);
}

/**
 * Applies a single raw transaction to the Summary Balance sheet.
 *
 * A transaction is skipped (returns `false`) when:
 * - Its date is in the future
 * - Its value is ≤ 0
 * - Its symbol is not found in the sheet
 * - It is a planned payment and the warning condition is triggered
 *   (cell value = 0 but note already contains money values)
 *
 * @param  {RawTransaction}                     rawDataRow           - Transaction record to process
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} summaryBalanceSheet  - Summary Balance sheet reference
 * @returns {boolean} `true` if the transaction was processed, `false` if skipped
 */
function processTransaction(rawDataRow, summaryBalanceSheet) {
  // Skip future-dated transactions
  // ERROR: 31/12/2019 and 2/2/2019 example of wrong comparison <-- do something with it
  if (dateAsUtc(rawDataRow.dateOfTransaction) > dateAsUtc(new Date())) return false;

  if (rawDataRow.value <= 0) return false;

  const rowNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
    (value) => value === rawDataRow.symbol.split(' : ')[0],
  );
  if (rowNum === 0) return false;

  const colNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getRange(rowNum, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
    (value) => value.toString() === EMPTY_STRING,
  ) - 1;

  const cell = summaryBalanceSheet.getRange(rowNum, colNum).getCell(1, 1);
  const cellNote = createCellNote(
    cell.getNote(),
    rawDataRow.dateOfTransaction,
    rawDataRow.value,
    rawDataRow.comment,
    rawDataRow.plannedPayment,
  );
  let cellValue = cell.getValue();

  if (rawDataRow.plannedPayment) {
    const noteMoneySum = calculateCellNoteMoneySum(cellNote);
    if (cellValue === 0 && noteMoneySum > 0) {
      showPlannedPaymentWarning(
        rawDataRow.symbol,
        rawDataRow.value,
        dateTimeToFormattedString(rawDataRow.timestamp),
        noteMoneySum,
      );
      return false;
    }

    cell.setBackground(calculateCellBackgroung(cellNote, cellValue));

    if (noteMoneySum > cellValue) {
      cellValue += noteMoneySum - cellValue;
    }
  } else {
    cellValue += rawDataRow.value;
  }

  cell.setValue(cellValue);
  cell.setNote(cellNote);

  return true;
}

// ── Warning dialog ───────────────────────────────────────────────────────────

/**
 * Shows a warning dialog when a planned-payment transaction would be applied
 * to a cell whose value is zero but whose note already contains money values,
 * indicating a potential double-count.
 *
 * @param  {string} tranSymbol    - Symbol of the transaction triggering the warning
 * @param  {number} tranValue     - Value of the transaction
 * @param  {string} tranTimeStamp - Formatted timestamp of the transaction
 * @param  {number} noteMoneySum  - Sum already recorded in the cell note
 */
function showPlannedPaymentWarning(tranSymbol, tranValue, tranTimeStamp, noteMoneySum) {
  const template = HtmlService.createTemplateFromFile('PlannedPaymentWarning');
  template.data = {
    transactionSymbol:    tranSymbol,
    noteMoneySum,
    transactionValue:     tranValue,
    transactionTimestamp: tranTimeStamp,
  };

  const html = template.evaluate();
  html.setTitle(PLANNED_PAYMENT_WARNING_TITLE);
  html.setHeight(180);
  html.setWidth(500);

  SpreadsheetApp.getActiveSpreadsheet().show(html);
}
