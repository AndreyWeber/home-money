/**
 * @fileoverview Functions for reading transaction records from the 'Raw Data' sheet.
 */

/// <reference path="Types.js" />

// ── Module-level cache ───────────────────────────────────────────────────────

/** @type {GoogleAppsScript.Spreadsheet.Sheet|null} */
let rawDataSheet = null;

// ── Sheet accessor ───────────────────────────────────────────────────────────

/**
 * Returns the 'Raw Data' sheet object. The result is cached so that
 * SpreadsheetApp is only called once per script execution.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The 'Raw Data' sheet
 */
function getRawDataSheet() {
  if (rawDataSheet === null) {
    rawDataSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(Sheets.RAW_DATA);
  }
  return rawDataSheet;
}

// ── Write functions ──────────────────────────────────────────────────────────

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

// ── Read functions ───────────────────────────────────────────────────────────

/**
 * Returns all transaction records from the 'Raw Data' sheet as an array of
 * objects keyed by normalised column header names.
 *
 * @returns {RawTransaction[]} All raw transaction records
 */
const getRawDataTransactionObjects = () => getRowsData(getRawDataSheet());

/**
 * Returns the element in rawDataArr that has the latest (maximum) transaction
 * date. Entries with missing or invalid dates are ignored.
 *
 * When the array is empty or all entries have invalid dates, returns a sentinel
 * object with an invalid date and empty strings for the text fields.
 *
 * @param  {RawTransaction[]} rawDataArr - Collection of raw transaction records
 * @returns {RawTransaction} Record with the latest transaction date, or a sentinel
 * @throws {Error} When rawDataArr is falsy
 */
const getRawDataTransactionObjectWithMaxDate = (rawDataArr) => rawDataArr
  ? rawDataArr
      .filter(to => to.dateOfTransaction != null && isValidDate(to.dateOfTransaction))
      .reduce(
        (prev, cur) =>
          isValidDate(prev.dateOfTransaction) &&
          dateAsUtc(prev.dateOfTransaction) > dateAsUtc(cur.dateOfTransaction)
            ? prev
            : cur,
        { dateOfTransaction: new Date(NaN), comment: '', symbol: '', value: 0, plannedPayment: false, timestamp: null },
      )
  : _throwErr("'rawData' argument is undefined");
