/****************************************************
 * Functions required to work with 'Raw Data' sheet *
 ****************************************************/

/********************
 * Quasi-properties *
 ********************/

 let rawDataSheet = null;

/*************
 * Functions *
 *************/

/**
 * Get 'RawData' sheet object
 * @returns {Object} 'RawData' sheet object
 */
function getRawDataSheet() {
  if (rawDataSheet === null) {
    rawDataSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(Sheets.RAW_DATA);
  }

  return rawDataSheet;
}

/**
 * Get 'RawData' objects collection
 * @returns {Array} 'RawData' objects collection
*/
const getRawDataTransactionObjects = () => getRowsData(getRawDataSheet());

/**
 * Get a 'RawData' object with max transaction date
 * @param {Array} rawDataArr - collection of 'RawData' objects
 * @returns {Object} 'RawData' object with max date
 */
const getRawDataTransactionObjectWithMaxDate = (rawDataArr) => rawDataArr
  ? rawDataArr.filter(to => to.dateOfTransaction && isValidDate(to.dateOfTransaction))
      .reduce((prevVal, curVal) =>
        isValidDate(prevVal.dateOfTransaction) &&
        (dateAsUtc(prevVal.dateOfTransaction) > dateAsUtc(curVal.dateOfTransaction))
          ? prevVal
          : curVal,
        {
          dateOfTransaction: new Date(NaN),
          comment: "",
          symbol: ""
        } // Initial value for reducer
      )
  : _throwErr("'rawData' argument is undefined");
