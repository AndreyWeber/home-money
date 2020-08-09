/****************************************************
 * Functions required to work with 'Metadata' sheet *
 ****************************************************/

/**
 * Constants
 */

const MetadataKeys = {
  TRANSACTIONS_MAX_DATE: "TransactionsMaxDate"
};

/**
 * Functions
 */

/**
 * Get 'Metadata' sheet reference
 */
const getMetadataSheet = () => SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName(Sheets.METADATA);

/**
 * Get Metadata object by key
 * @param {String} key
 * @returns {Object}
 */
const getMetadataObject = key => key
  ? getRowsData(getMetadataSheet()).find(el => el.key === key)
  : _throwErr("'key' argument is undefined");

/**
 * Get metadata oject containing a latest registered transaction date
 * This is required to check for what date expenses were calculated
 * last time
 * @returns {Object}
 */
const getLatestTransactionDateMetadataObject = () =>
  getMetadataObject(MetadataKeys.TRANSACTIONS_MAX_DATE) ||
  _throwErr(`Can't find ${Sheets.METADATA} key = '${MetadataKeys.TRANSACTIONS_MAX_DATE}'`);

/**
 * Set value property of Metadata object containing a latest registered
 * transaction date
 * @param {Date} date
 */
function setLatestTransactionsDate (date) {
  // Will throw an error if 'date' is undefined or has wrong format
  const dateString = dateToFormattedString(date);

  const metadataSheet = getMetadataSheet();

  // Get rownum of MetadataKeys.TRANSACTIONS_MAX_DATE
  // Tests showed, that this expression performance is good in case
  // of large amount of rows
  const rowNum = metadataSheet
    .getSheetValues(1, 1, metadataSheet.getMaxRows(), 1)
    .flat()
    .indexOf(MetadataKeys.TRANSACTIONS_MAX_DATE) + 1;

  if (rowNum === 0) {
    _throwErr(`${Sheets.METADATA} key = '${MetadataKeys.TRANSACTIONS_MAX_DATE}' doesn't exist`);
  }

  // Required value must be in the 'B' column
  const cell = metadataSheet.getRange(`B${rowNum}`);
  cell.setValue(dateString);
}
