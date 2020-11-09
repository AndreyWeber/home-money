/****************************************************
 * Functions required to work with 'Raw Data' sheet *
 ****************************************************/
/**
 * Quasi-properties
 */

 let rawDataSheet = null;

 /**
 * Functions
 */

/**
 * Get 'Raw Data' sheet reference
 * @returns {Object}
 */
function getRawDataSheet() {
  if (rawDataSheet === null) {
    rawDataSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(Sheets.RAW_DATA);

    return rawDataSheet;
  }

  return rawDataSheet;
}

/**
 * Get 'Raw Data' objects collection
 * @returns {Object[]}
*/
const getRawDataTransactionObjects = () => getRowsData(getRawDataSheet());

/**
 * Get max date from 'Date of transaction' column of the 'Raw Data' sheet
 * @returns {Date} Returns Invalid Date in case of empty 'Date of transaction' column
 */
const getRawDataTransactionsMaxDate = (rawData) => rawData
  ? rawData.filter(to => to.dateOfTransaction && isValidDate(to.dateOfTransaction))
  .map(to => to.dateOfTransaction)
  .reduce((prevVal, curVal) =>
    isValidDate(prevVal) && (dateAsUtc(prevVal) > dateAsUtc(curVal))
      ? prevVal
      : curVal,
    new Date(NaN) // Initial value for reducer
  )
  : _throwErr("'rawData' argument is undefined");
