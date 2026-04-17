/**
 * @fileoverview Functions for reading and writing the 'Transactions History'
 * sheet, and for displaying the Transactions History sidebar.
 */

/// <reference path="Types.js" />

// ── Sidebar launcher ─────────────────────────────────────────────────────────

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

// ── Data providers (called from HTML sidebar via google.script.run) ──────────

/**
 * Returns the transactions-history JSON stored in the cell currently selected
 * on the Summary Balance tab, or a sentinel object when another tab is active.
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

// ── Write functions ──────────────────────────────────────────────────────────

/**
 * Appends rawDataRow to the transaction-history JSON stored in the
 * corresponding Transactions History cell, identified by matching the
 * transaction symbol against the Summary Balance row layout.
 *
 * If the cell already contains a JSON value that cannot be parsed (e.g. a
 * partial write or a stray "undefined" string), the existing value is
 * discarded with a warning and a fresh array is started.
 *
 * @param  {RawTransaction}                     rawDataRow               - Processed transaction record to append
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} summaryBalanceSheet      - Summary Balance sheet (used for row/col lookup)
 * @param  {GoogleAppsScript.Spreadsheet.Sheet} transactionsHistorySheet - Transactions History sheet to write into
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

  // TODO: Can be simplified to getRange(rowNum, colNum) only. Should be tested
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
