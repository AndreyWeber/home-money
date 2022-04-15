/****************************************************
 * Functions required to work with 'Metadata' sheet *
 ****************************************************/

/********************
 * Quasi-properties *
 ********************/

let metadataSheet = null;

/*************
 * Constants *
 *************/

const MetadataKeys = {
  LATEST_TRANSACTION_DATE: "LatestTransactionDate",
  LATEST_TRANSACTION_NAME: "LatestTransactionName",
  LATEST_TRANSACTION_SYMBOL: "LatestTransactionSymbol"
};

/*************
 * Functions *
 *************/

/**
 * Get 'Metadata' sheet object
 * @returns {Object} 'Metadata' sheet object
 */
function getMetadataSheet() {
  if (metadataSheet === null) {
    metadataSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(Sheets.METADATA);
  }

  return metadataSheet;
}

/**
 * Get collection of all 'Metadata' objects
 * @returns {Array} 'Metadata' objects collection
 */
const getMetadataObjects = () => getRowsData(getMetadataSheet());

/**
 * Get 'Metadata' object by key
 * @param {String} key - 'Metadata' object key to search by
 * @returns {Object} 'Metadata' object
 */
const getMetadataObject = (key) => key
  ? getMetadataObjects.find(el => el.key === key)
  : _throwErr("'key' argument is undefined");

/**
 * Get 'Metadata' object with the latest registered transaction date
 * @returns {Object} 'Metadata' object
 */
const getLatestTransactionDateMetadataObject = () =>
  getMetadataObject(MetadataKeys.LATEST_TRANSACTION_DATE) ||
  _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_DATE}'`);

/**
 * Get 'Metadata' object with the latest registered transaction name
 * @returns {Object} 'Metadata' object
 */
const getLatestTransactionNameMetadataObject = () =>
  getMetadataObject(MetadataKeys.LATEST_TRANSACTION_NAME) ||
  _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_NAME}'`);

/**
 * Get 'Metadata' object with the latest registered transaction symbol
 * @returns {Object} 'Metadata' object
 */
const getLatestTransactionSymbolMetadataObject = () =>
  getMetadataObject(MetadataKeys.LATEST_TRANSACTION_SYMBOL) ||
  _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.LATEST_TRANSACTION_SYMBOL}'`);

/**
 * Get collection of all 'Metadata' objects
 * @returns {Array} collection of all 'Metadata' objects
 */
const getAllMetataObjectsJson = () => toJsonString(getMetadataObjects);

/**
 * Set value property of a 'Metadata' object found by the provided key
 * @param {String} metadataKey - 'Metadata' object key to search by
 * @param {String} value - 'Metadata' object value to update
 */
function setMetadataValue(metadataKey, value) {
  const metadataSheet = getMetadataSheet();

  const rowNum = metadataSheet
    .getSheetValues(1, 1, metadataSheet.getMaxRows(), 1)
    .flat()
    .indexOf(metadataKey) + 1;

  if (rowNum === 0) {
    _throwErr(`${Sheets.METADATA} key = '${metadataKey}' doesn't exist`);
  }

  const cell = metadataSheet.getRange(`B${rowNum}`);
  cell.setValue(value);
}

/**
 * Set value property of Metadata object containing the latest registered transaction date
 * and save it into the appropriate cell of Metadata sheet
 * @param {Date} date - latest registered transacation date
 */
function setLatestTransactionsDate(date) {
  // Will throw an error if 'date' is undefined or has wrong format
  const dateString = dateToFormattedString(date);

  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_DATE, dateString);
}

/**
 * Set value property of Metadata object containing the latest registered transaction name
 * and save it into the appropriate cell of Metadata sheet
 * @param {String} name - latest registered transaction name
 */
function setLatestTransactionName(name) {
  if (!name) {
    _throwErr("'name' argument is undefined");
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_NAME, name);
}

/**
 * Set value property of Metadata object containing the latest registered transaction symbol
 * and save it into the appropriate cell of Metadata sheet
 * @param {String} symbol - latest registered trnsaction symbol
 */
function setLatestTransactionSymbol(symbol) {
  if (!symbol) {
    _throwErr("'symbol' argument is undefined");
  }
  setMetadataValue(MetadataKeys.LATEST_TRANSACTION_SYMBOL, symbol);
}
