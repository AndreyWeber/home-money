// ===================================================================
// ISSUES:
// ===================================================================
// 1. При указании любых цифр в комментариях они учитываются при подсчетах
//    IsPlanned значений.
// 2. Если заранее написать комментарий, содержащий цифры на клетку с
//    IsPlanned значением, то эти цифры будут использованы в посчетах и
//    будет посчитан overrun
// ===================================================================
// Algorithm:
// 1. Get list of transaction records from "Raw Data" sheet.
// 2. For each transaction record which date is less than current system date:
//   2.1. Get row and column number for symbol from transaction. Skip transaction if
//    symbol doesn't exist.
//   2.2. Create cell note and calculate cell value both based on transaction record data and current cell data.
//   2.3. If transaction marked as planned change cell background color and calculate overrun.
//   2.4. If transaction marked as planned and cell contains zero value skip transaction and
//    show warning message.
//   2.5. Set cell new value, note and possibly background color as mentioned in point 2.4.
// 3. All transactions with date greater than current system date will be skipped.
// 4. If transaction record processed successfully delete it and leave it as is if not.
// ===================================================================

/**
 * Spreadsheet open event handler. Creates new menu items related to
 * transactions processing.
 */
function onOpen() {
  // Create transactions menu entries
  SpreadsheetApp.getUi()
    .createMenu('Transactions')
    .addItem('Process Transactions', 'processTransactions')
    //.addItem('Show Daily Expenses', 'showDailyExpenses')
    .addItem('Show Transactions History', 'showTransactionsHistorySidebar')
    .addItem('Show Metadata Viewer', 'showMetadataSidebar')
    .addToUi();
}

// ===================================================================
// Dialogs.
// ===================================================================

/**
 * Shows daily expenses for the chosen date.
 */
function showDailyExpenses() {
  // Create HTML to show
  var html = HtmlService.createHtmlOutputFromFile('DailyExpenses')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  // Show HTML
  SpreadsheetApp.getUi()
    .showModalDialog(html, DAILY_EXPENSES_DATA_TITLE);
}

/**
 * Try to show sidebar with the list of transactions history entries for a
 * cell currently selected on the Summary Balance tab.
 */
function showTransactionsHistorySidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('TransactionsHistorySidebar')
    .setTitle('Transactions History');

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}

function showMetadataSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('MetadataViewerSidebar')
    .setTitle('Metadata Viewer');

  SpreadsheetApp
    .getUi()
    .showSidebar(html);
}

/**
 *
 */
function getDailyExpensesData(dateToShow, includePlanned) {
  // Get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get raw transactions data collection
  var rawDataSheet = ss.getSheetByName(Sheets.RAW_DATA);
  var rawData = getRowsData(rawDataSheet);

  // Calculate actual daily expenses value
  var dailyExpensesSumActual = 0;
  for (var i = 0; i < rawData.length; i++) {
    // Skip row if its date doesn't match specified date
    if (dateAsUtc(rawData[i].dateOfTransaction) != dateAsUtc(new Date(dateToShow))) {
      continue;
    }
    // Skip row if it is not required to include planned payment values in result
    if (!includePlanned && rawData[i].plannedPayment) {
      continue;
    }

    dailyExpensesSumActual += rawData[i].value;
  }

  // Get expected daily expenses value
  var summaryBalanceSheet = ss.getSheetByName(Sheets.SUMMARY_BALANCE);

  var rowNum = findValueIndex(flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
                              function(value) { return (value == FREE_DAILY_CASH_DATA_TITLE) ? true : false; });
  var colNum = findValueIndex(flattenArray(summaryBalanceSheet.getRange(rowNum + 1, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
                              isNumber);
  var dailyExpensesSumExpected = summaryBalanceSheet.getRange(rowNum, colNum).getCell(1, 1).getValue();

  return {
    sumActual: dailyExpensesSumActual + CURRENCY_SUFFIX,
    sumExpected: dailyExpensesSumExpected.toFixed(2) + CURRENCY_SUFFIX,
    overrun: dailyExpensesSumActual > dailyExpensesSumExpected ? (dailyExpensesSumActual - dailyExpensesSumExpected).toFixed(2) + CURRENCY_SUFFIX: "(нет)"
  };
}

// showPlannedPaymentWarning shows warning about collisions between
// cell value with planned payment and transactions planned payment data.
// Arguments:
//   - tranSymbol: transaction symbol
//   - tranValue: transaction value
//   - tranTimeStamp: transaction timestamp
//   - noteMoneySum: sum of money values from cell note
// Returns void.
function showPlannedPaymentWarning(tranSymbol, tranValue, tranTimeStamp, noteMoneySum) {
  // Create template and set data
  var template = HtmlService.createTemplateFromFile('PlannedPaymentWarning');
  template.data = {
    transactionSymbol: tranSymbol,
    noteMoneySum: noteMoneySum,
    transactionValue: tranValue,
    transactionTimestamp: tranTimeStamp
  };

  // Create HTML to show
  var html = template.evaluate();
  html.setTitle(PLANNED_PAYMENT_WARNING_TITLE);
  html.setHeight(180);
  html.setWidth(500);

  // Show HTML
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.show(html);
}

/**
 * Get transactions history stored on "Transactions History" tab
 * Cells on "Summary Balance" are mapped to the same cells in
 * "Transactions History" tab to store history in JSON format
 * @returns {Object}
 */
const getTransactionsHistoryData = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() === Sheets.SUMMARY_BALANCE) {
    const activeCellA1Notation = activeSheet
      .getActiveCell()
      .getA1Notation();
    const jsonStr = ss
      .getSheetByName(Sheets.TRANSACTIONS_HISTORY)
      .getRange(activeCellA1Notation)
      .getValue();

    // Return result
    return {
      isSummaryBalanceSheet: true,
      jsonCellA1Notation: activeCellA1Notation,
      json: jsonStr
    };
  }

  // Return default value
  return {
    isSummaryBalanceSheet: false,
    jsonCellA1Notation: null,
    jsonStr: null
  };
}

// ===================================================================
// Helper methods.
// ===================================================================

// createCellNote creates cell note.
// Arguments:
//   - existingNote: existing cell note should be enhanced
//   - dateToAdd: date should be added to the cell note
//   - valueToAdd: money value should be added to the cell note
//   - noteToAdd: note should be added to the cell note
//   - isPlanned: points if new note belongs to planned payment
// Returns string.
function createCellNote(existingNote, dateToAdd, valueToAdd, noteToAdd, isPlanned) {
  // Return exsiting note if noteToAdd is empty and it not belongs to planned payment
  if (!noteToAdd && !isPlanned) {
    return existingNote;
  }

  // If existing note is not empty add line feed to the end
  if (existingNote) {
    existingNote += "\u000A";
  }

  // Return cell note depending on belongs it to planned payment or not
  if (isPlanned) {
    return existingNote += !noteToAdd
      ? `${dateToFormattedString(dateToAdd)}: ${valueToAdd} - потрачено`
      : `${dateToFormattedString(dateToAdd)}: ${valueToAdd} - ${noteToAdd} - потрачено`;
  }

  // Return enhanced cell note if noteToAdd is not empty
  return existingNote += `${dateToFormattedString(dateToAdd)}: ${valueToAdd} - ${noteToAdd}`;
}

// calculateCellNoteMoneySum calculates sum of money values provided in the cell note.
// Arguments:
//   - cellNote: note of the cell
// Returns number.
function calculateCellNoteMoneySum(cellNote) {
  var result = 0;
  if (!cellNote) {
    return result;
  }

  var noteTokens = cellNote.split(/[\n|\s*]/);
  for (var i = 0; i < noteTokens.length; i++) {
    if (isNumber(noteTokens[i])) {
      result += Number(noteTokens[i]);
    }
  }

  return result;
}

// calculateCellBackgroung calculates cell background color for cells with
// planned payment.
// Arguments:
//   - cellNote: note of the cell
//   - cellValue: value of the cell
// Returns string.
function calculateCellBackgroung(cellNote, cellValue) {
  if (!cellNote || !cellValue) {
    return CELL_BACKGROUND_COLOR.noFill;
  }

  // Calculate sum of money values provided in the cell note
  var noteMoneySum = calculateCellNoteMoneySum(cellNote);

  // Calculate cell background color
  if (noteMoneySum == 0) {
    // Return red color if note doesn't contain money value
    return CELL_BACKGROUND_COLOR.notSpent;
  } else if (noteMoneySum < cellValue) {
    // Return green color if note money value less than cell value
    return CELL_BACKGROUND_COLOR.almostSpent;
  } else if (noteMoneySum >= cellValue) {
    // Return green coloe if note maney value greater or equal cell value
    return CELL_BACKGROUND_COLOR.spent;
  }
}

// ===================================================================
// Main methods.
// ===================================================================

// deleteTransaction deletes transaction record on Raw Data sheet by
// its time stamp value.
// Arguments:
//   - timeStamp: time stamp of data record on Raw Data sheet
// Returns void.
// TODO: Move to 'RawDataFunctions'
function deleteRawTransaction(timeStamp) {
  const ds = getRawDataSheet();

  // Get number of row to delete
  var rowNum = findValueIndex(
    flattenArray(ds.getSheetValues(1, 1, ds.getMaxRows(), 1)),
    (value) => +value === +timeStamp
  );

  if (rowNum !== 0) {
    ds.deleteRow(rowNum);
  }
}

/**
 * Write transactions history data to an appropriate cell on Transactions History sheet
 */
function addTransationHistoryRow(rawDataRow, summaryBalanceSheet, transactionsHistorySheet) {
  // TODO: Probably can be simplified
  const rowNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
    (value) => value === rawDataRow.symbol.split(" : ")[0]
  );

  const colNum = findValueIndex(
    flattenArray(summaryBalanceSheet.getRange(rowNum, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
    (value) => value.toString() === EMPTY_STRING
  ) - 1;

  // TODO: Can be simplified to getRange(rowNum, colNum) only. Should be tested
  const cell = transactionsHistorySheet
    .getRange(rowNum, colNum)
    .getCell(1, 1);

  const cellValue = cell.getValue();

  let jsonObj = cellValue ? JSON.parse(cellValue) : null; // TODO: check if cellValue exists but cannot be parsed at least if it's empty string or "null" or "undefined" string
  if (jsonObj) {
    if (jsonObj instanceof Array) {
      jsonObj.push(rawDataRow);
    } else {
      const arr = [];
      arr.push(jsonObj);
      arr.push(rawDataRow);

      jsonObj = arr;
    }
  } else {
    jsonObj = [];
    jsonObj.push(rawDataRow);
  }

  const jsonStr = JSON.stringify(jsonObj);

  cell.setValue(jsonStr);
}

// processTransactions processes raw transactions data and fill Summary Balance sheet table
// current actual column.
// Returns void.
function processTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summaryBalanceSheet = ss.getSheetByName(Sheets.SUMMARY_BALANCE);
  const transactionsHistorySheet = ss.getSheetByName(Sheets.TRANSACTIONS_HISTORY);

  // Get raw transactions data collection
  const rawTransactions = getRawDataTransactionObjects();

  // Process raw transactions data
  const processedRawTransactions = [];
  for (let i = 0; i < rawTransactions.length; i++) {
    if (processTransaction(rawTransactions[i], summaryBalanceSheet)) {
      // Write transactions history
      addTransationHistoryRow(rawTransactions[i], summaryBalanceSheet, transactionsHistorySheet);
      // Delete raw data row by its timestamp
      deleteRawTransaction(rawTransactions[i].timestamp);
      // Save all processed raw transactions
      processedRawTransactions.push(rawTransactions[i]);
    }
  }

  // Save processed raw transactions max date
  const rawTransactionsMaxDate = getRawDataTransactionsMaxDate(processedRawTransactions);
  setLatestTransactionsDate(rawTransactionsMaxDate);
}

// processTransaction processes particular raw transaction data entry and does
// changes in Summary Balance sheet data. Transation record can be skiped if
// current date is less than transaction date or if transaction value is equal to zero.
// Returns True if transaction record is processed, otherwise - False.
// Arguments:
//   - rawDataRow: array of transaction data objects
//   - summaryBalanceSheet: reference on Summary Balance sheet
// Returns boolean.
function processTransaction(rawDataRow, summaryBalanceSheet) {
  // Skip transaction if transaction date > current date
  // ERROR: 31/12/2019 and 2/2/2019 example of wrong comparison <-- do something with it
  if (dateAsUtc(rawDataRow.dateOfTransaction) > dateAsUtc(new Date())) {
    return false;
  }

  // Skip transaction if transaction value is 0 or less
  if (rawDataRow.value <= 0) {
    return false;
  }

  // Get row number
  var rowNum = findValueIndex(flattenArray(summaryBalanceSheet.getSheetValues(1, 1, summaryBalanceSheet.getMaxRows(), 1)),
                              function(value) { return value === rawDataRow.symbol.split(" : ")[0] ? true : false; });
  // Skip transaction if symbol doesn't exist
  if (rowNum === 0) {
    return false;
  }

  // Get column number
  var colNum = findValueIndex(flattenArray(summaryBalanceSheet.getRange(rowNum, 1, 1, summaryBalanceSheet.getMaxColumns()).getValues()),
                              function(value) { return value.toString() === EMPTY_STRING; }) - 1;

  // Get cell to modify
  var cell = summaryBalanceSheet.getRange(rowNum, colNum).getCell(1, 1);

  // Get cell value and note and set background for cell with planned payment
  var cellNote = createCellNote(cell.getNote(), rawDataRow.dateOfTransaction, rawDataRow.value, rawDataRow.comment, rawDataRow.plannedPayment);
  var cellValue = cell.getValue();
  if (rawDataRow.plannedPayment) {
    // If cell value is eqaul to zero and cell note money sum more than zero
    // than skip transaction and show warning
    var noteMoneySum = calculateCellNoteMoneySum(cellNote);
    if (cellValue === 0 && noteMoneySum > 0) {
      showPlannedPaymentWarning(rawDataRow.symbol, rawDataRow.value, dateTimeToFormattedString(rawDataRow.timestamp), noteMoneySum);
      return false;
    }

    // Calculate cell background color
    var cellBackgr = calculateCellBackgroung(cellNote, cellValue);
    cell.setBackground(cellBackgr);

    // Calculate overrun
    if (noteMoneySum > cellValue) {
      cellValue += noteMoneySum - cellValue
    }
  } else {
    cellValue += rawDataRow.value;
  }

  // Set cell value and note
  cell.setValue(cellValue);
  cell.setNote(cellNote);

  // Transaction processed successfully
  return true;
}
