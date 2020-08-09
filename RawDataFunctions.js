/****************************************************
 * Functions required to work with 'Raw Data' sheet *
 ****************************************************/

/**
 * Functions
 */

/**
 * Get 'Raw Data' sheet reference
 * @returns {Object}
 */
const getRawDataSheet = () => SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName(Sheets.RAW_DATA);

/**
 * Get 'Raw Data' objects collection
 * @returns {Object[]}
*/
const getRawDataTransactionObjects = () => getRowsData(getRawDataSheet());

/**
 * Get max date from 'Date of transaction' column of the 'Raw Data' sheet
 * @returns {Date}
 */
const getRawDataTransactionsMaxDate = () => getRawDataTransactionObjects()
  .filter(to => to.dateOfTransaction && isValidDate(to.dateOfTransaction))
  .map(to => to.dateOfTransaction)
  .reduce((prevVal, curVal) =>
    isValidDate(prevVal) && (dateAsUtc(prevVal) > dateAsUtc(curVal))
      ? prevVal
      : curVal,
    new Date(NaN) // Initial value for reducer
  );
