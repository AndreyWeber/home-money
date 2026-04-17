/**
 * @fileoverview Functions for reading and writing the 'Metadata' sheet.
 */

/// <reference path="Types.js" />

// ── Module-level cache ───────────────────────────────────────────────────────

/** @type {GoogleAppsScript.Spreadsheet.Sheet|null} */
let metadataSheet = null;

// ── Constants ────────────────────────────────────────────────────────────────

/**
 * Key names used in the 'Metadata' sheet.
 * @enum {string}
 */
const MetadataKeys = {
  LATEST_TRANSACTION_DATE:   'LatestTransactionDate',
  LATEST_TRANSACTION_NAME:   'LatestTransactionName',
  LATEST_TRANSACTION_SYMBOL: 'LatestTransactionSymbol',
  LATEST_TRANSACTION_VALUE:  'LatestTransactionValue',
};

// ── Sheet accessor ───────────────────────────────────────────────────────────

/**
 * Returns the 'Metadata' sheet object. The result is cached so that
 * SpreadsheetApp is only called once per script execution.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The 'Metadata' sheet
 */
function getMetadataSheet() {
  if (metadataSheet === null) {
    metadataSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(Sheets.METADATA);
  }
  return metadataSheet;
}

// ── Read functions ───────────────────────────────────────────────────────────

/**
 * Returns all entries from the 'Metadata' sheet as an array of objects.
 *
 * @returns {MetadataEntry[]} All metadata entries
 */
const getMetadataObjects = () => getRowsData(getMetadataSheet());

/**
 * Returns the metadata entry matching the given key, or `undefined` if not found.
 *
 * @param  {string} key - Key to search for
 * @returns {MetadataEntry|undefined} Matching entry, or undefined
 * @throws {Error} When key is falsy
 */
const getMetadataObject = (key) => key
  ? getMetadataObjects().find(el => el.key === key)
  : _throwErr("'key' argument is undefined");

/**
 * Returns all metadata entries serialised as a JSON string.
 *
 * @returns {string} JSON array of all {@link MetadataEntry} objects
 */
const getAllMetataObjectsJson = () => toJsonString(getMetadataObjects);

// ── Write functions ──────────────────────────────────────────────────────────

/**
 * Writes value to the 'Metadata' sheet cell in column B at the row whose
 * column A matches metadataKey.
 *
 * @param  {string}        metadataKey - Key identifying the row to update
 * @param  {string|number} value       - Value to write into column B
 * @throws {Error} When metadataKey does not exist in the sheet
 */
function setMetadataValue(metadataKey, value) {
  const sheet = getMetadataSheet();
  const rowNum = sheet
    .getSheetValues(1, 1, sheet.getMaxRows(), 1)
    .flat()
    .indexOf(metadataKey) + 1;

  if (rowNum === 0) {
    _throwErr(`${Sheets.METADATA} key = '${metadataKey}' doesn't exist`);
  }

  sheet.getRange(`B${rowNum}`).setValue(value);
}

/**
 * Persists the latest transaction date to the 'Metadata' sheet as an ISO string.
 *
 * @param  {Date} date - Latest registered transaction date
 * @throws {Error} When date is falsy
 */
function setLatestTransactionDate(date) {
  if (!date) {
    _throwErr(`'date' argument is undefined or not a date. value: ${date}`);
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_DATE, date.toISOString());
}

/**
 * Persists the latest transaction name to the 'Metadata' sheet.
 *
 * @param  {string} name - Latest registered transaction name
 * @throws {Error} When name is not a string
 */
function setLatestTransactionName(name) {
  if (!isString(name)) {
    _throwErr(`'name' argument is undefined or not a string. value: ${name}`);
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_NAME, name);
}

/**
 * Persists the latest transaction symbol to the 'Metadata' sheet.
 *
 * @param  {string} symbol - Latest registered transaction symbol
 * @throws {Error} When symbol is not a string
 */
function setLatestTransactionSymbol(symbol) {
  if (!isString(symbol)) {
    _throwErr(`'symbol' argument is undefined or not a string. value: ${symbol}`);
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_SYMBOL, symbol);
}

/**
 * Persists the latest transaction value to the 'Metadata' sheet.
 *
 * @param  {number} value - Latest registered transaction value
 * @throws {Error} When value is NaN
 */
function setLatestTransactionValue(value) {
  if (isNaN(value)) {
    _throwErr(`'value' argument is not a number. value: ${value}`);
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_VALUE, value);
}

// ── Custom spreadsheet functions ─────────────────────────────────────────────

/**
 * Returns the latest registered transaction date from the 'Metadata' sheet.
 *
 * @returns {string} ISO date string of the latest transaction date
 * @throws {Error} When the key is not found in the sheet
 * @customfunction
 */
function GET_LATEST_TRAN_DATE() {
  return getMetadataObject(MetadataKeys.LATEST_TRANSACTION_DATE)?.value
    ?? _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_DATE}'`);
}

/**
 * Returns the latest registered transaction name from the 'Metadata' sheet.
 *
 * @returns {string} Name of the latest transaction
 * @throws {Error} When the key is not found in the sheet
 * @customfunction
 */
function GET_LATEST_TRAN_NAME() {
  return getMetadataObject(MetadataKeys.LATEST_TRANSACTION_NAME)?.value
    ?? _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_NAME}'`);
}

/**
 * Returns the latest registered transaction symbol from the 'Metadata' sheet.
 *
 * @returns {string} Symbol of the latest transaction
 * @throws {Error} When the key is not found in the sheet
 * @customfunction
 */
function GET_LATEST_TRAN_SYMBOL() {
  return getMetadataObject(MetadataKeys.LATEST_TRANSACTION_SYMBOL)?.value
    ?? _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_SYMBOL}'`);
}

/**
 * Returns the latest registered transaction value from the 'Metadata' sheet.
 *
 * @returns {number} Value of the latest transaction
 * @throws {Error} When the key is not found in the sheet
 * @customfunction
 */
function GET_LATEST_TRAN_VALUE() {
  return getMetadataObject(MetadataKeys.LATEST_TRANSACTION_VALUE)?.value
    ?? _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_VALUE}'`);
}
